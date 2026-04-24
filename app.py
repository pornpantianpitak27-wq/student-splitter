import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection

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

def get_existing_logo():
    # ตรวจสอบไฟล์โลโก้ที่คุณอัปโหลดไว้ (1523.jpg หรือ logo_college.jpg)
    for name in ["1523.jpg", "logo_college.jpg"]:
        if os.path.exists(name):
            return name
    return None

# --- 3. ส่วนแสดงผลหน้าเว็บ ---
st.title("📑 ระบบออกใบรายชื่อนักศึกษา (ฉบับสมบูรณ์)")

logo_path = get_existing_logo()
if logo_path:
    st.success(f"✅ พบไฟล์โลโก้ '{logo_path}' ระบบจะแทรกลงใน Excel ให้โดยอัตโนมัติ")
else:
    st.error("❌ ไม่พบไฟล์โลโก้บน GitHub! กรุณาตรวจสอบว่ามีไฟล์ '1523.jpg' หรือ 'logo_college.jpg' อยู่ในโฟลเดอร์เดียวกับ app.py หรือไม่")

# --- 4. ฟังก์ชันสร้าง Excel (Layout C-V ตามสั่ง) ---
def create_excel_report(target_year):
    df_all = load_data()
    if df_all.empty: return None
    year_data = df_all[df_all['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active) 
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    bold_font = Font(name='Angsana New', size=15, bold=True)
    normal_font = Font(name='Angsana New', size=14)
    center_align = Alignment(horizontal='center', vertical='center')

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # 4.1 ตารางขวาบน (จบที่คอลัมน์ V)
        ws.merge_cells('O2:V2'); ws['O2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['O2'].border = border; ws['O2'].alignment = center_align; ws['O2'].font = bold_font
        ws.merge_cells('O3:P4'); ws['O3'] = "เช็คชื่อนักศึกษา"; ws['O3'].border = border; ws['O3'].alignment = center_align
        ws.merge_cells('Q3:S4'); ws['Q3'] = "เซ็นสอบกลางภาค"; ws['Q3'].border = border; ws['Q3'].alignment = center_align
        ws.merge_cells('T3:V4'); ws['T3'] = "เซ็นสอบปลายภาค"; ws['T3'].border = border; ws['T3'].alignment = center_align

        # 4.2 หัวข้อหลัก (Merge A ถึง V)
        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align
        
        # 4.3 หัวตาราง: C-D ชื่อ-สกุล, E-U ตารางเช็คชื่อ, V หมายเหตุ
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        ws['E8'] = "เดือน"; ws['E9'] = "วันที่"; ws['E10'] = "คาบ"
        ws.merge_cells('V8:V10'); ws['V8'] = "หมายเหตุ"

        for i in range(1, 17): ws.cell(row=10, column=5+i).value = i # ใส่เลขคาบ 1-16
            
        for r in range(8, 11):
            for c in range(1, 23):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # 4.4 รายชื่อนักศึกษา (จัดตำแหน่ง C-D)
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=4)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 23):
                ws.cell(row=curr, column=c).border = border; ws.cell(row=curr, column=c).alignment = center_align
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)

        # 4.5 ปรับขนาดคอลัมน์
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        for c_idx in range(5, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 3.5
        ws.column_dimensions['V'].width = 12

        # 4.6 *** การแทรกโลโก้ใหม่ (บังคับแสดงผลทุก Sheet) ***
        if logo_path:
            try:
                img = XLImage(logo_path)
                img.width = 85
                img.height = 85
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
    if f1: st.download_button("📥 ดาวน์โหลด ปี 1", f1, "Report_P1.xlsx", use_container_width=True)
with c2:
    f2 = create_excel_report("ปี2")
    if f2: st.download_button("📥 ดาวน์โหลด ปี 2", f2, "Report_P2.xlsx", use_container_width=True)
