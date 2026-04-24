import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
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
    except:
        return pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ', 'นามสกุล', 'ระดับชั้น', 'Room'])

def get_logo_image():
    if os.path.exists("logo_college.jpg"):
        try: return XLImage("logo_college.jpg")
        except: return None
    return None

# --- 2. ฟังก์ชันสร้างฟอร์มเกรด (ฉบับสมบูรณ์ - หน้าละ 25 คน) ---
def create_grade_report(target_year):
    df_all = load_data()
    if df_all.empty: return None
    year_data = df_all[df_all['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    # การตั้งค่า Style
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    f_bold = Font(name='Angsana New', size=14, bold=True)
    f_normal = Font(name='Angsana New', size=13)
    rotate_align = Alignment(horizontal='center', vertical='center', textRotation=90)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', indent=1)

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"เกรด-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # --- ตั้งค่าหัวกระดาษ (ซ้ำทุกหน้าอัตโนมัติ) ---
        ws.print_title_rows = '1:11'  # ล็อคแถว 1-11 ให้พิมพ์ทุกหน้า
        
        # ใส่โลโก้กึ่งกลาง (คอลัมน์ I)
        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            ws.add_image(img, 'I1') 

        # หัวเรื่อง
        ws.merge_cells('A4:R4'); ws['A4'] = "บัญชีผลการเรียนรายวิชา"; ws['A4'].alignment = center_align; ws['A4'].font = f_bold
        ws.merge_cells('A5:R5'); ws['A5'] = "ภาคเรียนที่  ...............  ปีการศึกษา .........................."; ws['A5'].alignment = center_align; ws['A5'].font = f_normal
        ws.merge_cells('A6:C6'); ws['A6'] = "รหัสวิชา  ………………………….."
        ws.merge_cells('D6:L6'); ws['D6'] = "ชื่อวิชา  ……………………………………………………………………………………"
        ws.merge_cells('M6:R6'); ws['M6'] = "หน่วยกิต ……. หน่วยกิต"
        ws.merge_cells('A7:H7'); ws['A7'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name}"
        ws.merge_cells('I7:R7'); ws['I7'] = "ผู้สอน  ..........................................................................................."
        for cell in ['A6','D6','M6','A7','I7']: ws[cell].font = f_normal

        # --- หัวตาราง (แถว 8-11) ---
        # ผสานเลขที่ และ รหัส
        ws.merge_cells('A8:A11'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B11'); ws['B8'] = "รหัสประจำตัว"
        # ผสาน ชื่อ-สกุล (เฉพาะส่วนหัว)
        ws.merge_cells('C8:D11'); ws['C8'] = "ชื่อ - สกุล"; ws['C8'].alignment = center_align

        # ทฤษฎี (E-J)
        ws.merge_cells('E8:J8'); ws['E8'] = "ทฤษฎี..........................หน่วยกิต"
        ws.merge_cells('E9:I9'); ws['E9'] = "คะแนนระหว่างภาค"
        theory_heads = {'E10':"เวลา/อุปกรณ์",'F10':"พฤติกรรม",'G10':"งาน/ทดสอบ",'H10':"สอบกลางภาค",'I10':"สอบปลายภาค",'J10':"คะแนนรวม"}
        for cell, val in theory_heads.items(): ws[cell] = val; ws[cell].alignment = rotate_align
        ws.merge_cells('J9:J10')

        # ปฏิบัติ (K-P)
        ws.merge_cells('K8:P8'); ws['K8'] = "ปฏิบัติ..................หน่วยกิต"
        ws.merge_cells('K9:O9'); ws['K9'] = "คะแนนระหว่างภาค"
        prac_heads = {'K10':"คุณภาพของงาน",'L10':"เวลา/อุปกรณ์",'M10':"พฤติกรรม",'N10':"การปฎิบัติงาน",'O10':"สอบทฤษฎีเชิงปฎิบัติ",'P10':"คะแนนรวม"}
        for cell, val in prac_heads.items(): ws[cell] = val; ws[cell].alignment = rotate_align
        ws.merge_cells('P9:P10')

        # ระดับ/หมายเหตุ
        ws.merge_cells('Q8:Q9'); ws['Q8'] = "ระดับ"; ws['Q10'] = "คะแนน"; ws.merge_cells('Q10:Q11')
        ws.merge_cells('R8:R11'); ws['R8'] = "หมายเหตุ"

        # คะแนนเต็ม (แถว 11)
        pts = {5:'10',6:'10',7:'20',8:'20',9:'40',10:'100',11:'20',12:'10',13:'10',14:'40',15:'20',16:'100'}
        for c, v in pts.items():
            cell = ws.cell(row=11, column=c); cell.value = v; cell.alignment = center_align

        # ตีกรอบหัวตาราง
        for r in range(8, 12):
            for c in range(1, 19):
                ws.cell(row=r, column=c).border = border; ws.cell(row=r, column=c).font = f_bold

        # --- ข้อมูลนักศึกษา (แยก C-ชื่อ D-นามสกุล และ Page Break ทุก 25 คน) ---
        for i, row in enumerate(room_data.itertuples(), 1):
            curr_row = 11 + i
            ws.cell(row=curr_row, column=1).value = i
            ws.cell(row=curr_row, column=2).value = row.รหัสนักศึกษา
            ws.cell(row=curr_row, column=3).value = row.ชื่อ      # ชื่ออยู่ C
            ws.cell(row=curr_row, column=4).value = row.นามสกุล  # นามสกุลอยู่ D
            
            for c in range(1, 19):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = border
                cell.alignment = left_align if c in [3, 4] else center_align

            # ใส่ตัวแบ่งหน้า (Page Break) ทุกๆ 25 รายชื่อ
            if i % 25 == 0:
                ws.row_breaks.append(Break(id=curr_row))

        # --- ตั้งค่าความกว้างคอลัมน์ตามสเปก ---
        ws.column_dimensions['A'].width = 3.86
        ws.column_dimensions['B'].width = 12.29
        ws.column_dimensions['C'].width = 9.0  # เฉลี่ย C-D ให้รวมได้ 17
        ws.column_dimensions['D'].width = 8.0 
        for c_idx in ['E','F','G','H','I','J','L','M','N','O']: ws.column_dimensions[c_idx].width = 2.86
        ws.column_dimensions['K'].width = 3.29
        ws.column_dimensions['P'].width = 3.29
        ws.column_dimensions['Q'].width = 5.86
        ws.column_dimensions['R'].width = 8

    wb.save(output)
    return output.getvalue()

# --- 3. ส่วนหน้าจอ Streamlit ---
st.title("🏫 ระบบพิมพ์บัญชีผลการเรียน (Version 25-Rows)")

t1, t2, t3 = st.tabs(["📝 ลงทะเบียน", "🔍 แก้ไขข้อมูล", "📥 ดาวน์โหลด"])

# (Tab 1 & 2 คงเดิมจากเวอร์ชันก่อนหน้า)
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
            st.success("อัปเดตข้อมูลแล้ว!"); st.rerun()

with t3:
    st.subheader("📥 ดาวน์โหลดไฟล์ Excel สำหรับใช้งานจริง")
    st.write("ฟอร์มนี้จัดหน้าละ 25 คน และพิมพ์หัวตารางอัตโนมัติในทุกหน้ากระดาษ")
    col1, col2 = st.columns(2)
    with col1:
        g1 = create_grade_report("ปี1")
        if g1: st.download_button("📊 โหลดฟอร์มเกรด ปี 1", g1, "Grade_P1_Final.xlsx", type="primary")
    with col2:
        g2 = create_grade_report("ปี2")
        if g2: st.download_button("📊 โหลดฟอร์มเกรด ปี 2", g2, "Grade_P2_Final.xlsx", type="primary")
