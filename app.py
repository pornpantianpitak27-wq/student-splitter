import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - ตามรูปแบบใบรายชื่อ", layout="wide")

# --- 2. การเชื่อมต่อฐานข้อมูล ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        cols_to_fix = ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ', 'นามสกุล', 'ระดับชั้น', 'Room']
        for col in cols_to_fix:
            if col in data.columns:
                data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
                data[col] = data[col].str.replace("'", "").replace('nan', '')
        return data
    except Exception:
        return pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ', 'นามสกุล', 'ระดับชั้น', 'Room'])

df = load_data()
ROOMS_P1 = [f"O1/{i}" for i in range(1, 16)]
ROOMS_P2 = [f"O2/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการรายชื่อและออกใบเช็คชื่อ (16 คาบ)")

# --- 3. ส่วนกรอกข้อมูล (แยกชื่อ-นามสกุลเพื่อความละเอียด) ---
st.subheader("➕ เพิ่มรายชื่อนักศึกษา")
tab_p1, tab_p2 = st.tabs(["📝 ลงทะเบียน ปี 1 (O1)", "📝 ลงทะเบียน ปี 2 (O2)"])

def student_form(year_label, room_options):
    with st.form(f"form_{year_label}", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
        with c1: batch = st.text_input("รุ่น", placeholder="23", key=f"b_{year_label}")
        with c2: sid = st.text_input("รหัสนักศึกษา", key=f"s_{year_label}")
        with c3: fname = st.text_input("ชื่อ", key=f"f_{year_label}")
        with c4: lname = st.text_input("นามสกุล", key=f"l_{year_label}")
        
        c5, c6 = st.columns(2)
        with c5: room = st.selectbox("เลือกห้องเรียน", room_options, key=f"r_{year_label}")
        with c6: st.info(f"ระดับชั้น: {year_label}")

        if st.form_submit_button("💾 บันทึกข้อมูล", use_container_width=True):
            if sid and fname:
                if sid in df['รหัสนักศึกษา'].values:
                    st.warning("⚠️ รหัสนี้มีในระบบแล้ว")
                else:
                    new_row = pd.DataFrame([{"รุ่น": f"'{batch}", "รหัสนักศึกษา": f"'{sid}", "ชื่อ": fname.strip(), "นามสกุล": lname.strip(), "ระดับชั้น": year_label, "Room": room}])
                    updated_df = pd.concat([df, new_row], ignore_index=True)
                    conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                    st.success("✅ บันทึกสำเร็จ!"); st.rerun()
            else:
                st.error("❌ กรุณากรอกข้อมูลให้ครบ")

with tab_p1: student_form("ปี1", ROOMS_P1)
with tab_p2: student_form("ปี2", ROOMS_P2)

# --- 4. ฟังก์ชันสร้าง Excel ตามรูปภาพที่คุณส่งมา ---
def create_excel_report(target_year):
    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    
    year_data = df[df['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # 1. หัวไฟล์ (Header)
        ws.merge_cells('A1:U1')
        ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
        ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')

        ws.merge_cells('A2:U2')
        ws['A2'] = f"ระดับ ปวส. ชั้นปีที่ {'1' if target_year=='ปี1' else '2'} ห้อง {r_name}"
        ws['A2'].alignment = Alignment(horizontal='center')

        ws.merge_cells('A3:K3')
        ws['A3'] = "วิชา..........................................................................."
        ws.merge_cells('L3:U3')
        ws['L3'] = "ผู้สอน..........................................................................."

        # 2. หัวตารางตามรูปภาพ (แถวที่ 4-6)
        # แถวที่ 4: หัวข้อหลัก
        headers = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล', ''] + ['' for _ in range(16)] + ['หมายเหตุ']
        ws.append(headers)
        
        # แถวที่ 5: เดือน/วันที่
        ws.append(['', '', '', 'เดือน'] + ['' for _ in range(16)] + [''])
        ws.append(['', '', '', 'วันที่'] + ['' for _ in range(16)] + [''])
        ws.append(['', '', '', 'คาบ'] + [f'{i+1}' for i in range(16)] + [''])

        # จัดการ Merge และเส้นขอบหัวตาราง (แถว 4 ถึง 7)
        for r in range(4, 8):
            for c in range(1, 22):
                ws.cell(row=r, column=c).border = border
                ws.cell(row=r, column=c).alignment = Alignment(horizontal='center', vertical='center')

        # 3. ข้อมูลนักศึกษา (เริ่มแถวที่ 8)
        for i, row in enumerate(room_data.itertuples(), 1):
            fullname = f"{row.ชื่อ} {row.นามสกุล}"
            ws.append([i, row.รหัสนักศึกษา, fullname, ''] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
            current_row = ws.max_row
            for c in range(1, 22):
                ws.cell(row=current_row, column=c).border = border
                ws.cell(row=current_row, column=c).alignment = Alignment(horizontal='center')
            ws.cell(row=current_row, column=3).alignment = Alignment(horizontal='left', indent=1)

        # 4. ส่วนท้ายไฟล์ (Footer)
        f_row = ws.max_row + 2
        ws.merge_cells(f'A{f_row}:K{f_row}')
        ws[f'A{f_row}'] = "ชื่ออาจารย์ที่ปรึกษา........................................................................"
        ws.merge_cells(f'L{f_row}:U{f_row}')
        ws[f'L{f_row}'] = "ชื่อหัวหน้าชั้น........................................................................"

        # 5. ปรับความกว้างคอลัมน์
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 8 # คอลัมน์ "เดือน/วันที่/คาบ"
        for c in range(5, 21):
            ws.column_dimensions[ws.cell(row=7, column=c).column_letter].width = 3.5

    wb.save(output)
    return output.getvalue()

# --- 5. ปุ่มดาวน์โหลด ---
st.divider()
c1, c2 = st.columns(2)
with c1:
    f1 = create_excel_report("ปี1")
    if f1: st.download_button("📥 ดาวน์โหลดไฟล์ ปี 1 (16 คาบ)", f1, "ใบรายชื่อ_ปี1.xlsx")
with c2:
    f2 = create_excel_report("ปี2")
    if f2: st.download_button("📥 ดาวน์โหลดไฟล์ ปี 2 (16 คาบ)", f2, "ใบรายชื่อ_ปี2.xlsx")

# --- 6. ส่วนแก้ไขข้อมูล ---
with st.expander("🔍 ค้นหา/แก้ไข/ย้ายห้อง"):
    search = st.text_input("ค้นหาชื่อหรือรหัส...")
    if not df.empty:
        found = df[df['ชื่อ'].str.contains(search, na=False) | df['รหัสนักศึกษา'].str.contains(search, na=False)]
        st.data_editor(found, use_container_width=True, key="editor")
