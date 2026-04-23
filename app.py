import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection
import os

# --- 1. ตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - วิทยาลัยเทคโนโลยีนนทบุรี", layout="wide")

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
        return pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ', 'นามสกุล', 'ระดับชั้น', 'Room'])

df = load_data()
ROOMS_P1 = [f"O1/{i}" for i in range(1, 16)]
ROOMS_P2 = [f"O2/{i}" for i in range(1, 16)]

# --- 3. ส่วน UI ---
st.title("📑 ระบบออกใบรายชื่อสมบูรณ์แบบ (100% Format)")

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

        if st.form_submit_button("💾 บันทึกข้อมูลลงฐานข้อมูล", use_container_width=True):
            if sid and fname:
                new_row = pd.DataFrame([{"รุ่น": f"'{batch}", "รหัสนักศึกษา": f"'{sid}", "ชื่อ": fname.strip(), "นามสกุล": lname.strip(), "ระดับชั้น": year_label, "Room": room}])
                current_df = load_data()
                updated_df = pd.concat([current_df, new_row], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success("✅ บันทึกสำเร็จ!"); st.rerun()

with tab_p1: student_form("ปี1", ROOMS_P1)
with tab_p2: student_form("ปี2", ROOMS_P2)

# --- 4. ฟังก์ชันสร้าง Excel เลย์เอาต์เป๊ะ ---
def create_excel_report(target_year):
    data_to_use = load_data()
    if data_to_use.empty: return None
    year_data = data_to_use[data_to_use['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    bold_font = Font(name='Angsana New', size=14, bold=True)
    normal_font = Font(name='Angsana New', size=13)
    center_align = Alignment(horizontal='center', vertical='center')

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # 4.1 ตารางขวาบน (Row 2-4)
        ws.merge_cells('N2:U2'); ws['N2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['N2'].border = border; ws['N2'].alignment = center_align; ws['N2'].font = bold_font
        ws.merge_cells('N3:O4'); ws['N3'] = "เช็คชื่อนักศึกษา"; ws['N3'].border = border; ws['N3'].alignment = center_align
        ws.merge_cells('P3:R4'); ws['P3'] = "เซ็นสอบกลางภาค"; ws['P3'].border = border; ws['P3'].alignment = center_align
        ws.merge_cells('S3:U4'); ws['S3'] = "เซ็นสอบปลายภาค"; ws['S3'].border = border; ws['S3'].alignment = center_align

        # 4.2 หัวข้อกลาง (Row 5-7)
        ws.merge_cells('A5:U5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:U6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align
        ws.merge_cells('A7:K7'); ws['A7'] = "วิชา..........................................................................."; ws['L7'] = "ผู้สอน..........................................................................."

        # 4.3 หัวตารางหลัก (Row 8-10)
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:K10'); ws['C8'] = "ชื่อ-สกุล"
        ws['L8'] = "เดือน"; ws['L9'] = "วันที่"; ws['L10'] = "คาบ"
        ws.merge_cells('U8:U10'); ws['U8'] = "หมายเหตุ"

        for i in range(1, 9): ws.cell(row=10, column=12+i).value = i # คาบ 1-8
            
        for r in range(8, 11):
            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # 4.4 รายชื่อนักศึกษา (Row 11+)
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            
            # Merge ชื่อ-สกุล (C ถึง K)
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=11)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            
            for c in range(1, 22):
                cell = ws.cell(row=curr, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = normal_font
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)

        # 4.5 ตั้งค่าความกว้างคอลัมน์
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 30
        for c_idx in range(12, 22):
            ws.column_dimensions[get_column_letter(c_idx)].width = 4

        # 4.6 แทรกโลโก้
        if os.path.exists("1523.jpg"):
            img = Image("1523.jpg")
            img.height = 70; img.width = 70 
            ws.add_image(img, 'H1') 

    wb.save(output)
    return output.getvalue()

# --- 5. ปุ่มดาวน์โหลด ---
st.divider()
c1, c2 = st.columns(2)
with c1:
    f1 = create_excel_report("ปี1")
    if f1: st.download_button("📥 โหลดใบรายชื่อ ปี 1", f1, "ใบรายชื่อ_ปี1.xlsx", use_container_width=True)
with c2:
    f2 = create_excel_report("ปี2")
    if f2: st.download_button("📥 โหลดใบรายชื่อ ปี 2", f2, "ใบรายชื่อ_ปี2.xlsx", use_container_width=True)
