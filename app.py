import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - Complete & Stable", layout="wide")

# --- 2. การเชื่อมต่อฐานข้อมูล ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # ทำความสะอาดข้อมูลเบื้องต้น
        for col in data.columns:
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "").replace('nan', '')
        return data
    except Exception:
        return pd.DataFrame()

df = load_data()

# รายชื่อห้องเรียนตามที่คุณระบุ
ROOMS_P1 = [f"O1/{i}" for i in range(1, 16)]
ROOMS_P2 = [f"O2/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการรายชื่อและออกใบเช็คชื่อ 16 คาบ (สมบูรณ์)")

# --- 3. ส่วนกรอกข้อมูล (แยกช่องชื่อ-นามสกุล เพื่อความละเอียด) ---
st.subheader("➕ เพิ่มรายชื่อนักศึกษาใหม่")
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
                # ตรวจสอบว่ามีคอลัมน์ที่ต้องการไหม ถ้าไม่มีให้สร้างใหม่
                current_df = load_data()
                new_row = pd.DataFrame([{
                    "รุ่น": f"'{batch}", 
                    "รหัสนักศึกษา": f"'{sid}", 
                    "ชื่อ": fname.strip(), 
                    "นามสกุล": lname.strip(), 
                    "ระดับชั้น": year_label, 
                    "Room": room
                }])
                updated_df = pd.concat([current_df, new_row], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success("✅ บันทึกสำเร็จ!"); st.rerun()
            else:
                st.error("❌ กรุณากรอกรหัสและชื่อนักศึกษา")

with tab_p1: student_form("ปี1", ROOMS_P1)
with tab_p2: student_form("ปี2", ROOMS_P2)

# --- 4. ฟังก์ชันสร้าง Excel (จัดรูปแบบตามรูปภาพที่คุณส่งมา) ---
def create_excel_report(target_year):
    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    
    # ดึงข้อมูลล่าสุด
    data_to_use = load_data()
    if data_to_use.empty: return None
    
    year_data = data_to_use[data_to_use['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # ส่วนหัว (Header)
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

        # หัวตาราง (แถวที่ 4-7)
        ws.append(['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล', 'คอลัมน์ว่างสำหรับเดือน/วันที่'])
        ws.append(['', '', '', 'เดือน'])
        ws.append(['', '', '', 'วันที่'])
        ws.append(['', '', '', 'คาบ'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ'])

        # วาดเส้นขอบหัวตาราง
        for r in range(4, 8):
            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # ข้อมูลนักศึกษา (เริ่มแถวที่ 8)
        for i, row_data in enumerate(room_data.itertuples(), 1):
            # ป้องกัน Error ถ้าไม่มีคอลัมน์ชื่อ/นามสกุล
            f_name = getattr(row_data, 'ชื่อ', '')
            l_name = getattr(row_data, 'นามสกุล', '')
            # ถ้าหาไม่เจอจริงๆ ลองหาจากคอลัมน์ "ชื่อ-นามสกุล" แบบเก่า
            fullname = f"{f_name} {l_name}".strip() if f_name else getattr(row_data, 'ชื่อ_นามสกุล', 'ไม่ระบุชื่อ')
            
            ws.append([i, row_data.รหัสนักศึกษา, fullname, ''] + ['' for _ in range(16)] + [f"รุ่น {getattr(row_data, 'รุ่น', '')}"])
            
            curr_row = ws.max_row
            for c in range(1, 22):
                ws.cell(row=curr_row, column=c).border = border
                ws.cell(row=curr_row, column=c).alignment = Alignment(horizontal='center')
            ws.cell(row=curr_row, column=3).alignment = Alignment(horizontal='left', indent=1)

        # ส่วนท้าย (Footer)
        last = ws.max_row + 2
        ws.merge_cells(f'A{last}:K{last}')
        ws[f'A{last}'] = "ชื่ออาจารย์ที่ปรึกษา............................................................"
        ws.merge_cells(f'L{last}:U{last}')
        ws[f'L{last}'] = "ชื่อหัวหน้าชั้น............................................................"

        # ปรับความกว้างคอลัมน์
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 7 
        for c in range(5, 21):
            ws.column_dimensions[ws.cell(row=7, column=c).column_letter].width = 3.5

    wb.save(output)
    return output.getvalue()

# --- 5. ปุ่มดาวน์โหลด ---
st.divider()
col_p1, col_p2 = st.columns(2)
with col_p1:
    f1 = create_excel_report("ปี1")
    if f1: st.download_button("📥 ดาวน์โหลดใบรายชื่อ ปี 1", f1, "ใบรายชื่อ_ปี1.xlsx")
with col_p2:
    f2 = create_excel_report("ปี2")
    if f2: st.download_button("📥 ดาวน์โหลดใบรายชื่อ ปี 2", f2, "ใบรายชื่อ_ปี2.xlsx")

# --- 6. แก้ไขปัญหา Error โดยการตรวจสอบโครงสร้างคอลัมน์ ---
with st.expander("🛠️ ตรวจสอบโครงสร้างฐานข้อมูล (Debug)"):
    st.write("คอลัมน์ปัจจุบันใน Sheets ของคุณคือ:", df.columns.tolist())
    if st.button("🔄 รีเฟรชข้อมูลล่าสุด"):
        st.rerun()
