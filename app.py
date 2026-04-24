import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection
import os

# --- 1. ตั้งค่าพื้นฐาน ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา", layout="wide")
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
    for name in ["1523.jpg", "logo_college.jpg"]:
        if os.path.exists(name): return name
    return None

# --- 2. ส่วนสร้าง Excel (Layout ใหม่) ---
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
        
        # --- หัวข้อขวาบน (ปรับให้จบที่ V) ---
        ws.merge_cells('O2:V2'); ws['O2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['O2'].border = border; ws['O2'].alignment = center_align; ws['O2'].font = bold_font
        ws.merge_cells('O3:P4'); ws['O3'] = "เช็คชื่อนักศึกษา"; ws['O3'].border = border; ws['O3'].alignment = center_align
        ws.merge_cells('Q3:S4'); ws['Q3'] = "เซ็นสอบกลางภาค"; ws['Q3'].border = border; ws['Q3'].alignment = center_align
        ws.merge_cells('T3:V4'); ws['T3'] = "เซ็นสอบปลายภาค"; ws['T3'].border = border; ws['T3'].alignment = center_align

        # --- หัวข้อหลัก (Merge A ถึง V) ---
        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align
        
        # --- หัวตาราง ( Layout: C-D คือชื่อ, E เริ่มตาราง, V คือหมายเหตุ) ---
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        
        ws['E8'] = "เดือน"; ws['E9'] = "วันที่"; ws['E10'] = "คาบ"
        ws.merge_cells('V8:V10'); ws['V8'] = "หมายเหตุ"

        # ใส่เลขคาบในแถวที่ 10 (เริ่มจาก F เป็นต้นไป หรือตามใจชอบ ในที่นี้เริ่มช่องถัดจากหัวข้อ)
        for i in range(1, 16): # ใส่เลข 1-15 ในช่องคาบ
            ws.cell(row=10, column=5+i).value = i
            
        # ใส่เส้นขอบหัวตาราง (A ถึง V)
        for r in range(8, 11):
            for c in range(1, 23): # 1=A, 22=V
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # --- ข้อมูลนักศึกษา ---
        for i, row in enumerate(room_data.itertuples(), 1):
            curr_row = 10 + i
            ws.cell(row=curr_row, column=1).value = i
            ws.cell(row=curr_row, column=2).value = row.รหัสนักศึกษา
            
            # ชื่อ-นามสกุล อยู่ที่ C และ D (Merge 2 คอลัมน์)
            ws.merge_cells(start_row=curr_row, start_column=3, end_row=curr_row, end_column=4)
            ws.cell(row=curr_row, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            
            # ใส่เส้นขอบทุกช่องในแถว (A ถึง V)
            for c in range(1, 23):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = border; cell.alignment = center_align
            ws.cell(row=curr_row, column=3).alignment = Alignment(horizontal='left', indent=1)

        # --- ปรับความกว้างคอลัมน์ ---
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15 # ชื่อ
        ws.column_dimensions['D'].width = 15 # นามสกุล
        for c_idx in range(5, 22): # คอลัมน์ E ถึง U (ตารางเช็คชื่อ)
            ws.column_dimensions[get_column_letter(c_idx)].width = 3.5
        ws.column_dimensions['V'].width = 12 # หมายเหตุ

        # --- แทรกโลโก้ ---
        logo_path = get_existing_logo()
        if logo_path:
            try:
                img = Image(logo_path)
                img.width, img.height = 80, 80
                ws.add_image(img, 'H1')
            except: pass

    wb.save(output)
    return output.getvalue()

# --- 3. ส่วนแสดงผล Streamlit ---
st.title("🏫 ระบบออกใบรายชื่อ (Update Layout C-V)")

tab1, tab2 = st.tabs(["ลงทะเบียน", "ดาวน์โหลด"])

with tab1:
    st.info("กรุณากรอกข้อมูลนักศึกษาผ่านฟอร์มเดิมของคุณ")

with tab2:
    c1, c2 = st.columns(2)
    with c1:
        f1 = create_excel_report("ปี1")
        if f1: st.download_button("📥 โหลด ปี 1 (C-V Layout)", f1, "Report_P1.xlsx")
    with c2:
        f2 = create_excel_report("ปี2")
        if f2: st.download_button("📥 โหลด ปี 2 (C-V Layout)", f2, "Report_P2.xlsx")
