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

# รายการชื่อไฟล์โลโก้ที่ระบบจะพยายามค้นหา
LOGO_FILES = ["logo_college.jpg", "1523.jpg", "logo_college.png"]

def get_existing_logo():
    for logo in LOGO_FILES:
        if os.path.exists(logo):
            return logo
    return None

# --- 3. ส่วนหน้าจอหลัก ---
st.title("🏫 ระบบออกใบรายชื่อ (เวอร์ชันแก้ไขโลโก้)")

current_logo = get_existing_logo()
if current_logo:
    st.success(f"✅ ตรวจพบไฟล์โลโก้ '{current_logo}' ระบบพร้อมแทรกรูปใน Excel")
else:
    st.error("❌ ไม่พบไฟล์โลโก้บน GitHub! กรุณาอัปโหลดไฟล์ 'logo_college.jpg' ไว้ที่เดียวกับ app.py")

# [ส่วนการลงทะเบียนคงเดิม...]
tab_p1, tab_p2 = st.tabs(["📝 ปี 1 (O1)", "📝 ปี 2 (O2)"])
# (ใช้ฟอร์มการกรอกข้อมูลจากเวอร์ชันก่อนหน้าได้เลย)

# --- 4. ฟังก์ชันสร้าง Excel พร้อมแทรกโลโก้ ---
def create_excel_report(target_year):
    df_data = load_data()
    if df_data.empty: return None
    year_data = df_data[df_data['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    bold_font = Font(name='Angsana New', size=15, bold=True)
    normal_font = Font(name='Angsana New', size=14)
    center_align = Alignment(horizontal='center', vertical='center')

    logo_path = get_existing_logo()

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # จัดเลย์เอาต์หัวกระดาษตามภาพ Untitled.png
        ws.merge_cells('N2:U2'); ws['N2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['N2'].border = border; ws['N2'].alignment = center_align; ws['N2'].font = bold_font
        ws.merge_cells('N3:O4'); ws['N3'] = "เช็คชื่อนักศึกษา"; ws['N3'].border = border; ws['N3'].alignment = center_align
        ws.merge_cells('P3:R4'); ws['P3'] = "เซ็นสอบกลางภาค"; ws['P3'].border = border; ws['P3'].alignment = center_align
        ws.merge_cells('S3:U4'); ws['S3'] = "เซ็นสอบปลายภาค"; ws['S3'].border = border; ws['S3'].alignment = center_align

        ws.merge_cells('A5:U5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:U6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align

        # หัวตาราง 4 ชั้น ตามภาพ image_df60da.png
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:K10'); ws['C8'] = "ชื่อ-สกุล"
        ws['L8'] = "เดือน"; ws['L9'] = "วันที่"; ws['L10'] = "คาบ"
        ws.merge_cells('U8:U10'); ws['U8'] = "หมายเหตุ"
        for i in range(1, 9): ws.cell(row=10, column=12+i).value = i
            
        for r in range(8, 11):
            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # ข้อมูลนักศึกษา
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=11)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 22):
                ws.cell(row=curr, column=c).border = border; ws.cell(row=curr, column=c).alignment = center_align
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)

        # ตั้งค่าความกว้าง
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 17
        ws.column_dimensions['C'].width = 32
        for c_idx in range(12, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 4

        # *** จุดที่ทำให้โลโก้ปรากฏ ***
        if logo_path:
            try:
                img = Image(logo_path)
                img.height = 75  # ปรับขนาดตามภาพ image_df60da.png
                img.width = 75
                ws.add_image(img, 'H1') # วางตำแหน่งกลางหัวกระดาษ
            except Exception:
                pass

    wb.save(output)
    return output.getvalue()

# --- 5. ปุ่มดาวน์โหลด ---
st.divider()
c1, c2 = st.columns(2)
with c1:
    f1 = create_excel_report("ปี1")
    if f1: st.download_button("📥 โหลดไฟล์ ปี 1", f1, "Report_P1.xlsx", use_container_width=True)
with c2:
    f2 = create_excel_report("ปี2")
    if f2: st.download_button("📥 โหลดไฟล์ ปี 2", f2, "Report_P2.xlsx", use_container_width=True)
