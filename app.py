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

# --- 1. ตั้งค่าพื้นฐานและการเชื่อมต่อ ---
st.set_page_config(page_title="วิทยาลัยเทคโนโลยีนนทบุรี", layout="wide")
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

def get_logo_image():
    logo_filename = "logo_college.jpg"
    if os.path.exists(logo_filename):
        try: return XLImage(logo_filename)
        except: return None
    return None

# --- 2. ฟังก์ชันสร้างใบรายชื่อ (Attendance) - ยึดโครงสร้างที่คุณส่งมา ---
def create_attendance_report(target_year):
    df_all = load_data()
    if df_all.empty: return None
    year_data = df_all[df_all['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook(); wb.remove(wb.active)
    side = Side(style='thin'); border = Border(left=side, right=side, top=side, bottom=side)
    f_bold = Font(name='Angsana New', size=15, bold=True)
    center = Alignment(horizontal='center', vertical='center')

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ใบรายชื่อ-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        ws.print_title_rows = '1:10' # ล็อคหัวตารางซ้ำทุกหน้า
        
        # หัวกระดาษ (O-V)
        ws.merge_cells('O2:V2'); ws['O2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['O2'].border = border; ws['O2'].alignment = center; ws['O2'].font = f_bold
        ws.merge_cells('O3:P4'); ws['O3'] = "เช็คชื่อนักศึกษา"; ws['O3'].border = border; ws['O3'].alignment = center
        ws.merge_cells('Q3:S4'); ws['Q3'] = "เซ็นสอบกลางภาค"; ws['Q3'].border = border; ws['Q3'].alignment = center
        ws.merge_cells('T3:V4'); ws['T3'] = "เซ็นสอบปลายภาค"; ws['T3'].border = border; ws['T3'].alignment = center

        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = f_bold; ws['A5'].alignment = center
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = f_bold; ws['A6'].alignment = center
        
        # หัวตาราง
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        ws['E8']="เดือน"; ws['E9']="วันที่"; ws['E10']="คาบ"; ws.merge_cells('V8:V10'); ws['V8']="หมายเหตุ"
        for i in range(1, 17): ws.cell(row=10, column=5+i).value = i
        for r in range(8, 11):
            for c in range(1, 23):
                cell = ws.cell(row=r, column=c); cell.border = border; cell.alignment = center; cell.font = f_bold

        # ข้อมูล (หน้าละ 25 คน)
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=4)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 23):
                cell = ws.cell(row=curr, column=c); cell.border = border; cell.alignment = center
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)
            
            if i % 25 == 0: ws.row_breaks.append(Break(id=curr))

        # โลโก้
        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            ws.add_image(img, 'H1')

        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 10
        for c_idx in range(5, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 3.5
        ws.column_dimensions['V'].width = 10

    wb.save(output); return output.getvalue()

# --- 3. ฟังก์ชันสร้างใบกรอกเกรด (Grade) - แยกคอลัมน์ชื่อ/นามสกุล ---
def create_grade_report(target_year):
    df_all = load_data()
    if df_all.empty: return None
    year_data = df_all[df_all['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook(); wb.remove(wb.active)
    side = Side(style='thin'); border = Border(left=side, right=side, top=side, bottom=side)
    f_bold = Font(name='Angsana New', size=14, bold=True)
    rotate_align = Alignment(horizontal='center', vertical='center', textRotation=90)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', indent=1)

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ใบเกรด-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        ws.print_title_rows = '1:11'

        # โลโก้กึ่งกลาง (คอลัมน์ I)
        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            ws.add_image(img, 'I1')

        # หัวตาราง (แถว 8-11)
        ws.merge_cells('C8:D11'); ws['C8'] = "ชื่อ - สกุล" # หัวผสาน C-D
        ws.merge_cells('A8:A11'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B11'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('E8:J8'); ws['E8'] = "ทฤษฎี..........................หน่วยกิต"
        ws.merge_cells('K8:P8'); ws['K8'] = "ปฏิบัติ..................หน่วยกิต"
        
        # หัวข้อหมุน 90 องศา
        theory = ['เวลา/อุปกรณ์', 'พฤติกรรม', 'งาน/ทดสอบ', 'สอบกลางภาค', 'สอบปลายภาค']
        for idx, val in enumerate(theory):
            cell = ws.cell(row=10, column=5+idx); cell.value = val; cell.alignment = rotate_align
        
        # ข้อมูลรายชื่อ (แยก C-ชื่อ D-นามสกุล)
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 11 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.cell(row=curr, column=3).value = row.ชื่อ      # คอลัมน์ C
            ws.cell(row=curr, column=4).value = row.นามสกุล  # คอลัมน์ D
            for c in range(1, 19):
                cell = ws.cell(row=curr, column=c); cell.border = border
                cell.alignment = left_align if c in [3,4] else center_align
            if i % 25 == 0: ws.row_breaks.append(Break(id=curr))

        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 8
        for c in ['E','F','G','H','I','J','K','L','M','N','O','P']: ws.column_dimensions[c].width = 3

    wb.save(output); return output.getvalue()

# --- 4. ส่วนแสดงผล UI หน้าเว็บ ---
st.title("🏫 ระบบจัดการวิทยาลัย (ใบรายชื่อ & ใบเกรด)")

tab1, tab2, tab3 = st.tabs(["📝 ลงทะเบียน", "🔍 แก้ไขข้อมูล", "📥 ดาวน์โหลด"])

with tab1:
    with st.form("reg_form"):
        c1, c2, c3 = st.columns(3)
        with c1: batch = st.text_input("รุ่น"); sid = st.text_input("รหัสนักศึกษา")
        with c2: fname = st.text_input("ชื่อ"); lname = st.text_input("นามสกุล")
        with c3:
            level = st.selectbox("ระดับชั้น", ["ปี1", "ปี2"])
            room = st.selectbox("ห้อง", [f"{'O1' if level=='ปี1' else 'O2'}/{i}" for i in range(1, 16)])
        if st.form_submit_button("💾 บันทึก"):
            df = load_data()
            new = pd.DataFrame([{"รุ่น": f"'{batch}", "รหัสนักศึกษา": f"'{sid}", "ชื่อ": fname.strip(), "นามสกุล": lname.strip(), "ระดับชั้น": level, "Room": room}])
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=pd.concat([df, new], ignore_index=True))
            st.success("บันทึกแล้ว!"); st.rerun()

with tab2:
    df_edit = load_data()
    if not df_edit.empty:
        edited = st.data_editor(df_edit, num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกการแก้ไข"):
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=edited)
            st.success("อัปเดตแล้ว!"); st.rerun()

with tab3:
    st.subheader("📥 ดาวน์โหลด (หน้าละ 25 คน)")
    c1, c2 = st.columns(2)
    with c1:
        st.info("📝 ใบรายชื่อ (Attendance)")
        st.download_button("📥 โหลดใบรายชื่อ ปี 1", create_attendance_report("ปี1"), "Attendance_P1.xlsx", use_container_width=True)
        st.download_button("📥 โหลดใบรายชื่อ ปี 2", create_attendance_report("ปี2"), "Attendance_P2.xlsx", use_container_width=True)
    with c2:
        st.success("📊 ใบกรอกคะแนน (Grade Form)")
        st.download_button("📥 โหลดใบเกรด ปี 1", create_grade_report("ปี1"), "Grade_P1.xlsx", use_container_width=True)
        st.download_button("📥 โหลดใบเกรด ปี 2", create_grade_report("ปี2"), "Grade_P2.xlsx", use_container_width=True)
