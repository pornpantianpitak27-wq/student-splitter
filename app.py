import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter  # เพิ่มการ Import เพื่อแก้ Error
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
    except:
        return pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ', 'นามสกุล', 'ระดับชั้น', 'Room'])

def get_logo_image():
    if os.path.exists("logo_college.jpg"):
        try: return XLImage("logo_college.jpg")
        except: return None
    return None

# --- 2. ฟังก์ชันสร้างใบเช็คชื่อ (หน้าละ 25 คน) ---
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
    left_align = Alignment(horizontal='left', vertical='center', indent=1)

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"เช็คชื่อ-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        ws.print_title_rows = '1:10' # พิมพ์หัวซ้ำทุกหน้า
        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            ws.add_image(img, 'H1') # โลโก้กึ่งกลาง

        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = f_bold; ws['A5'].alignment = center
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = f_bold; ws['A6'].alignment = center
        
        # หัวตาราง
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        ws.merge_cells('E8:E10'); ws['E8'] = "เดือน/วัน/คาบ"
        for i in range(1, 17): ws.cell(row=10, column=5+i).value = i
        ws.merge_cells('V8:V10'); ws['V8'] = "หมายเหตุ"

        for r in range(8, 11):
            for c in range(1, 23):
                cell = ws.cell(row=r, column=c); cell.border = border; cell.alignment = center; cell.font = f_bold

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

        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 9
        ws.column_dimensions['D'].width = 8
        for c_idx in range(5, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 3.5
        ws.column_dimensions['V'].width = 10

    wb.save(output); return output.getvalue()

# --- 3. ฟังก์ชันสร้างฟอร์มเกรด (หน้าละ 25 คน) ---
def create_grade_report(target_year):
    df_all = load_data()
    if df_all.empty: return None
    year_data = df_all[df_all['ระดับชั้น'] == target_year]
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

        ws.merge_cells('A4:R4'); ws['A4'] = "บัญชีผลการเรียนรายวิชา"; ws['A4'].alignment = center_align; ws['A4'].font = f_bold
        ws.merge_cells('A5:R5'); ws['A5'] = "ภาคเรียนที่  ...............  ปีการศึกษา .........................."; ws['A5'].alignment = center_align; ws['A5'].font = f_normal
        
        ws.merge_cells('A8:A11'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B11'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D11'); ws['C8'] = "ชื่อ - สกุล"; ws['C8'].alignment = center_align

        theory_heads = {'E10':"เวลา/อุปกรณ์",'F10':"พฤติกรรม",'G10':"งาน/ทดสอบ",'H10':"สอบกลางภาค",'I10':"สอบปลายภาค"}
        for cell, val in theory_heads.items(): ws[cell] = val; ws[cell].alignment = rotate_align
        
        prac_heads = {'K10':"คุณภาพของงาน",'L10':"เวลา/อุปกรณ์",'M10':"พฤติกรรม",'N10':"การปฎิบัติงาน",'O10':"สอบทฤษฎีเชิงปฎิบัติ"}
        for cell, val in prac_heads.items(): ws[cell] = val; ws[cell].alignment = rotate_align

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

        ws.column_dimensions['A'].width = 3.86
        ws.column_dimensions['B'].width = 12.29
        ws.column_dimensions['C'].width = 9.0; ws.column_dimensions['D'].width = 8.0 
        for c in ['E','F','G','H','I','J','L','M','N','O']: ws.column_dimensions[c].width = 2.86

    wb.save(output); return output.getvalue()

# --- 4. Streamlit UI ---
st.title("🏫 ระบบพิมพ์เอกสารวิทยาลัย (Complete Version)")

t1, t2, t3 = st.tabs(["📝 ลงทะเบียน", "🔍 แก้ไขข้อมูล", "📥 ดาวน์โหลดเอกสาร"])

with t1:
    with st.form("reg"):
        c1, c2, c3 = st.columns(3)
        with c1: batch = st.text_input("รุ่น"); sid = st.text_input("รหัสนักศึกษา")
        with c2: fname = st.text_input("ชื่อ"); lname = st.text_input("นามสกุล")
        with c3:
            level = st.selectbox("ระดับชั้น", ["ปี1", "ปี2"])
            room = st.selectbox("ห้อง", [f"{'O1' if level=='ปี1' else 'O2'}/{i}" for i in range(1, 16)])
        if st.form_submit_button("💾 บันทึก"):
            if sid and fname:
                df = load_data()
                new = pd.DataFrame([{"รุ่น":f"'{batch}","รหัสนักศึกษา":f"'{sid}","ชื่อ":fname.strip(),"นามสกุล":lname.strip(),"ระดับชั้น":level,"Room":room}])
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=pd.concat([df, new], ignore_index=True))
                st.success("บันทึกสำเร็จ!"); st.rerun()

with t2:
    df_edit = load_data()
    if not df_edit.empty:
        edited = st.data_editor(df_edit, num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกการแก้ไข"):
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=edited)
            st.success("อัปเดตเรียบร้อย!"); st.rerun()

with t3:
    st.subheader("📥 ดาวน์โหลดเอกสาร (หน้าละ 25 คน)")
    c1, c2 = st.columns(2)
    with c1:
        st.write("📝 **ใบเช็คชื่อ**")
        if st.button("Generate Attendance P1"):
            st.download_button("Download P1", create_attendance_report("ปี1"), "Att_P1.xlsx")
        if st.button("Generate Attendance P2"):
            st.download_button("Download P2", create_attendance_report("ปี2"), "Att_P2.xlsx")
    with c2:
        st.write("📊 **ฟอร์มกรอกเกรด**")
        if st.button("Generate Grade P1"):
            st.download_button("Download P1", create_grade_report("ปี1"), "Grade_P1.xlsx")
        if st.button("Generate Grade P2"):
            st.download_button("Download P2", create_grade_report("ปี2"), "Grade_P2.xlsx")
