import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าและการเชื่อมต่อ ---
st.set_page_config(page_title="วิทยาลัยเทคโนโลยีนนทบุรี", layout="wide")
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # ทำความสะอาดข้อมูล: ลบช่องว่างและจัดการค่าว่าง
        for col in data.columns:
            data[col] = data[col].astype(str).str.strip() # ลบช่องว่างหัว-ท้าย
            data[col] = data[col].str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].replace(['nan', 'None', ''], '')
        return data
    except Exception:
        return pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ', 'นามสกุล', 'ระดับชั้น', 'Room'])

def get_logo_image():
    logo_filename = "logo_college.jpg"
    if os.path.exists(logo_filename):
        try: return XLImage(logo_filename)
        except: return None
    return None

# --- 2. ฟังก์ชันใบรายชื่อ (Attendance Report) ---
def create_attendance_report(target_year):
    df_all = load_data()
    # กรองข้อมูลโดยไม่สนช่องว่าง
    year_data = df_all[df_all['ระดับชั้น'].str.contains(target_year, na=False)]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook(); wb.remove(wb.active)
    side = Side(style='thin'); border = Border(left=side, right=side, top=side, bottom=side)
    f_bold = Font(name='Angsana New', size=15, bold=True)
    center = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center', indent=1)

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ใบรายชื่อ-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        ws.print_title_rows = '1:10'

        # ส่วนหัว O-V และหัวเรื่อง
        ws.merge_cells('O2:V2'); ws['O2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['O2'].border = border; ws['O2'].alignment = center; ws['O2'].font = f_bold
        ws.merge_cells('O3:P4'); ws['O3'] = "เช็คชื่อนักศึกษา"; ws['O3'].border = border; ws['O3'].alignment = center
        ws.merge_cells('Q3:S4'); ws['Q3'] = "เซ็นสอบกลางภาค"; ws['Q3'].border = border; ws['Q3'].alignment = center
        ws.merge_cells('T3:V4'); ws['T3'] = "เซ็นสอบปลายภาค"; ws['T3'].border = border; ws['T3'].alignment = center
        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = f_bold; ws['A5'].alignment = center
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = f_bold; ws['A6'].alignment = center

        # หัวตาราง (Merge C-D เฉพาะหัว)
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"; ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        ws['E8']="เดือน"; ws['E9']="วันที่"; ws['E10']="คาบ"; ws.merge_cells('V8:V10'); ws['V8']="หมายเหตุ"
        for i in range(1, 17): ws.cell(row=10, column=5+i).value = i

        for r in range(8, 11):
            for c in range(1, 23):
                cell = ws.cell(row=r, column=c); cell.border = border; cell.alignment = center; cell.font = f_bold

        # ข้อมูลนักศึกษา (แยก C-D)
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.cell(row=curr, column=3).value = row.ชื่อ
            ws.cell(row=curr, column=4).value = row.นามสกุล
            for c in range(1, 23):
                cell = ws.cell(row=curr, column=c); cell.border = border
                cell.alignment = left_align if c in [3, 4] else center
            if i % 25 == 0: ws.row_breaks.append(Break(id=curr))

        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            ws.add_image(img, 'H1')

        ws.column_dimensions['C'].width = 10; ws.column_dimensions['D'].width = 10
        for c_idx in range(5, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 3.5

    wb.save(output); return output.getvalue()

# --- 3. ฟังก์ชันใบกรอกเกรด (Grade Report) ---
def create_grade_report(target_year):
    df_all = load_data()
    year_data = df_all[df_all['ระดับชั้น'].str.contains(target_year, na=False)]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook(); wb.remove(wb.active)
    side = Side(style='thin'); border = Border(left=side, right=side, top=side, bottom=side)
    f_bold = Font(name='Angsana New', size=14, bold=True)
    f_normal = Font(name='Angsana New', size=13)
    rotate_align = Alignment(horizontal='center', vertical='center', textRotation=90)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', indent=1)

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"เกรด-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        ws.print_title_rows = '1:11'

        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            ws.add_image(img, 'I1') 

        # หัวกระดาษตามโครงสร้างที่ส่งมา
        ws.merge_cells('A4:R4'); ws['A4'] = "บัญชีผลการเรียนรายวิชา"; ws['A4'].alignment = center_align; ws['A4'].font = f_bold
        ws.merge_cells('A5:R5'); ws['A5'] = "ภาคเรียนที่  ...............  ปีการศึกษา .........................."; ws['A5'].alignment = center_align; ws['A5'].font = f_normal
        ws.merge_cells('A6:C6'); ws['A6'] = "รหัสวิชา  ………………………….."
        ws.merge_cells('D6:L6'); ws['D6'] = "ชื่อวิชา  ……………………………………………………………………………………"
        ws.merge_cells('M6:R6'); ws['M6'] = "หน่วยกิต ……. หน่วยกิต"
        ws.merge_cells('A7:H7'); ws['A7'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name}"
        ws.merge_cells('I7:R7'); ws['I7'] = "ผู้สอน  ..........................................................................................."

        # หัวตาราง (A-D)
        for col, val in [('A8', 'เลขที่'), ('B8', 'รหัสประจำตัว')]:
            ws.merge_cells(f'{col}:{col[0]}11'); ws[col] = val; ws[col].alignment = center_align
        ws.merge_cells('C8:D11'); ws['C8'] = "ชื่อ - สกุล"; ws['C8'].alignment = center_align

        # ส่วนคะแนน (ทฤษฎี/ปฏิบัติ) และคะแนนเต็ม (แถว 11)
        ws.merge_cells('E8:J8'); ws['E8'] = "ทฤษฎี..........................หน่วยกิต"
        ws.merge_cells('K8:P8'); ws['K8'] = "ปฏิบัติ..................หน่วยกิต"
        pts = {5:'10',6:'10',7:'20',8:'20',9:'40',10:'100',11:'20',12:'10',13:'10',14:'40',15:'20',16:'100'}
        for c, v in pts.items(): 
            ws.cell(row=11, column=c).value = v
            ws.cell(row=11, column=c).alignment = center_align

        # วนลูปข้อมูล (แยก C-D)
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 11 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.cell(row=curr, column=3).value = row.ชื่อ
            ws.cell(row=curr, column=4).value = row.นามสกุล
            for c in range(1, 19):
                cell = ws.cell(row=curr, column=c); cell.border = border
                cell.alignment = left_align if c in [3, 4] else center_align
            if i % 25 == 0: ws.row_breaks.append(Break(id=curr))

        ws.column_dimensions['C'].width = 10; ws.column_dimensions['D'].width = 10

    wb.save(output); return output.getvalue()

# --- 4. หน้าจอหลัก ---
st.title("🏫 ระบบจัดการวิทยาลัย (Complete Version)")

# โหลดข้อมูลมาตรวจสอบก่อนสร้างปุ่ม
df_full = load_data()
has_p1 = not df_full[df_full['ระดับชั้น'].str.contains("ปี1", na=False)].empty
has_p2 = not df_full[df_full['ระดับชั้น'].str.contains("ปี2", na=False)].empty

t1, t2, t3 = st.tabs(["📝 ลงทะเบียน", "🔍 แก้ไขข้อมูล", "📥 ดาวน์โหลดเอกสาร"])

with tab3:
    c1, c2 = st.columns(2)
    with c1:
        st.info("📝 ใบรายชื่อ (Attendance)")
        if has_p1:
            st.download_button("📥 โหลดใบรายชื่อ ปี 1", create_attendance_report("ปี1"), "Attendance_P1.xlsx", use_container_width=True)
        if has_p2:
            st.download_button("📥 โหลดใบรายชื่อ ปี 2", create_attendance_report("ปี2"), "Attendance_P2.xlsx", use_container_width=True)
        else:
            st.warning("⚠️ ไม่พบข้อมูลนักศึกษา ปี 2")

    with c2:
        st.success("📊 ใบกรอกเกรด (Grade Form)")
        if has_p1:
            st.download_button("📥 โหลดใบเกรด ปี 1", create_grade_report("ปี1"), "Grade_P1.xlsx", use_container_width=True)
        if has_p2:
            st.download_button("📥 โหลดใบเกรด ปี 2", create_grade_report("ปี2"), "Grade_P2.xlsx", use_container_width=True)
