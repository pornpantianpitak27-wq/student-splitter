import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - Final Stable", layout="wide")

# --- 2. การเชื่อมต่อฐานข้อมูล ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        for col in data.columns:
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "").replace('nan', '')
        return data
    except Exception:
        return pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room'])

df = load_data()
ROOMS_P1 = [f"O1/{i}" for i in range(1, 16)]
ROOMS_P2 = [f"O2/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการรายชื่อนักศึกษา (Stable Version)")

# --- 3. ส่วนกรอกข้อมูล ---
st.subheader("➕ เพิ่มรายชื่อนักศึกษาใหม่")
tab_p1, tab_p2 = st.tabs(["📝 ลงทะเบียน ปี 1 (O1)", "📝 ลงทะเบียน ปี 2 (O2)"])

def student_form(year_label, room_options):
    with st.form(f"form_{year_label}", clear_on_submit=True):
        c1, c2, c3 = st.columns([1, 2, 4])
        with c1: batch = st.text_input("รุ่น", placeholder="23", key=f"b_{year_label}")
        with c2: sid = st.text_input("รหัสนักศึกษา", key=f"s_{year_label}")
        with c3: fullname = st.text_input("ชื่อ-นามสกุล", key=f"f_{year_label}")
        
        c4, c5 = st.columns(2)
        with c4: room = st.selectbox("ห้องเรียน", room_options, key=f"r_{year_label}")
        with c5: st.info(f"ระดับชั้น: {year_label}")

        if st.form_submit_button("💾 บันทึกข้อมูล", use_container_width=True):
            if sid and fullname:
                new_row = pd.DataFrame([{
                    "รุ่น": f"'{batch}",
                    "รหัสนักศึกษา": f"'{sid}",
                    "ชื่อ-นามสกุล": fullname.strip(),
                    "ระดับชั้น": year_label,
                    "Room": room
                }])
                current_df = load_data()
                updated_df = pd.concat([current_df, new_row], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success("✅ บันทึกสำเร็จ!"); st.rerun()
            else:
                st.error("❌ กรุณากรอกรหัสและชื่อ")

with tab_p1: student_form("ปี1", ROOMS_P1)
with tab_p2: student_form("ปี2", ROOMS_P2)

# --- 4. ฟังก์ชันสร้าง Excel (แก้ไขปัญหา IndexError) ---
def create_excel_report(target_year):
    data_to_use = load_data()
    if data_to_use.empty: return None
    
    # กรองข้อมูลตามชั้นปี
    year_data = data_to_use[data_to_use['ระดับชั้น'] == target_year]
    
    # *** จุดสำคัญ: ถ้าไม่มีข้อมูลปีนั้นๆ ให้ return None ทันที เพื่อไม่ให้เกิด IndexError ***
    if year_data.empty:
        return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # ลบ Sheet เริ่มต้น
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    
    # สร้าง Sheet แยกตามห้อง
    rooms_in_year = sorted(year_data['Room'].unique())
    for r_name in rooms_in_year:
        if not r_name: continue
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # หัวไฟล์
        ws.merge_cells('A1:U1')
        ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
        ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A2:U2')
        ws['A2'] = f"ระดับ ปวส. ชั้นปีที่ {'1' if target_year=='ปี1' else '2'} ห้อง {r_name}"
        ws['A2'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A3:K3'); ws['A3'] = "วิชา..........................................................................."
        ws.merge_cells('L3:U3'); ws['L3'] = "ผู้สอน..........................................................................."

        # หัวตาราง (อิงตามไฟล์ต้นฉบับที่คุณต้องการ)
        ws.append(['เลขที่', 'รหัสประจำตัว', 'ชื่อ-สกุล', ''])
        ws.append(['', '', '', 'เดือน'])
        ws.append(['', '', '', 'วันที่'])
        ws.append(['', '', '', 'คาบ'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ'])

        for r in range(4, 8):
            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = Alignment(horizontal='center', vertical='center')

        # ข้อมูลรายชื่อ
        for i, row in enumerate(room_data.itertuples(), 1):
            name_val = getattr(row, 'ชื่อ_นามสกุล', getattr(row, 'ชื่อ', 'ไม่ระบุชื่อ'))
            ws.append([i, row.รหัสนักศึกษา, name_val, ''] + ['' for _ in range(16)] + [f"รุ่น {getattr(row, 'รุ่น', '')}"])
            curr_row = ws.max_row
            for c in range(1, 22):
                ws.cell(row=curr_row, column=c).border = border
                ws.cell(row=curr_row, column=c).alignment = Alignment(horizontal='center')
            ws.cell(row=curr_row, column=3).alignment = Alignment(horizontal='left', indent=1)

        # ท้ายไฟล์
        f_row = ws.max_row + 2
        ws.merge_cells(f'A{f_row}:K{f_row}'); ws[f'A{f_row}'] = "ชื่ออาจารย์ที่ปรึกษา............................................................"
        ws.merge_cells(f'L{f_row}:U{f_row}'); ws[f'L{f_row}'] = "ชื่อหัวหน้าชั้น............................................................"

        # ปรับความกว้างคอลัมน์
        ws.column_dimensions['A'].width = 6; ws.column_dimensions['B'].width = 16; ws.column_dimensions['C'].width = 32; ws.column_dimensions['D'].width = 8
        for c in range(5, 21): ws.column_dimensions[ws.cell(row=7, column=c).column_letter].width = 3.5

    wb.save(output)
    return output.getvalue()

# --- 5. ส่วนแสดงปุ่มดาวน์โหลด ---
st.divider()
st.subheader("🖨️ ออกใบรายชื่อ")
c1, c2 = st.columns(2)

with c1:
    f1 = create_excel_report("ปี1")
    if f1:
        st.download_button("📥 ดาวน์โหลดใบรายชื่อ ปี 1", f1, "ใบรายชื่อ_ปี1.xlsx", use_container_width=True)
    else:
        st.info("💡 ยังไม่มีข้อมูลนักศึกษา ปี 1")

with c2:
    f2 = create_excel_report("ปี2")
    if f2:
        st.download_button("📥 ดาวน์โหลดใบรายชื่อ ปี 2", f2, "ใบรายชื่อ_ปี2.xlsx", use_container_width=True)
    else:
        st.warning("💡 ยังไม่มีข้อมูลนักศึกษา ปี 2 (ปุ่มดาวน์โหลดจะปรากฏเมื่อมีข้อมูล)")

# --- 6. ส่วนค้นหา/แก้ไข ---
with st.expander("🔍 ค้นหาและแก้ไขข้อมูล"):
    search = st.text_input("ค้นหาชื่อหรือรหัส...")
    if not df.empty:
        found = df[df['ชื่อ-นามสกุล'].str.contains(search, na=False) | df['รหัสนักศึกษา'].str.contains(search, na=False)]
        st.data_editor(found, use_container_width=True, key="main_editor")
