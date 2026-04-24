import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าพื้นฐาน ---
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

# --- 2. ฟังก์ชันสร้างฟอร์มเกรด (ปรับตำแหน่งโลโก้กึ่งกลาง + ชื่อ-สกุล C-D) ---
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

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"เกรด-{r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # --- 1. ใส่โลโก้ให้กึ่งกลาง (ประมาณคอลัมน์ H/I) ---
        img = get_logo_image()
        if img:
            img.width, img.height = 75, 75
            # วางกึ่งกลางระหว่างคอลัมน์ A-R (ประมาณคอลัมน์ I)
            ws.add_image(img, 'I1') 

        # --- 2. หัวกระดาษ ---
        ws.merge_cells('A4:R4'); ws['A4'] = "บัญชีผลการเรียนรายวิชา"; ws['A4'].alignment = center_align; ws['A4'].font = f_bold
        ws.merge_cells('A5:R5'); ws['A5'] = "ภาคเรียนที่  ...............  ปีการศึกษา .........................."; ws['A5'].alignment = center_align; ws['A5'].font = f_normal
        ws.merge_cells('A6:C6'); ws['A6'] = "รหัสวิชา  ………………………….."
        ws.merge_cells('D6:L6'); ws['D6'] = "ชื่อวิชา  ……………………………………………………………………………………"
        ws.merge_cells('M6:R6'); ws['M6'] = "หน่วยกิต ……. หน่วยกิต"
        ws.merge_cells('A7:H7'); ws['A7'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name}"
        ws.merge_cells('I7:R7'); ws['I7'] = "ผู้สอน  ..........................................................................................."
        for cell in ['A6','D6','M6','A7','I7']: ws[cell].font = f_normal

        # --- 3. โครงสร้างหัวตาราง (แถย 8-11) ---
        # เลขที่ (A), รหัสประจำตัว (B)
        for col, val in [('A8', 'เลขที่'), ('B8', 'รหัสประจำตัว')]:
            ws.merge_cells(f'{col}:{col[0]}11'); ws[col] = val; ws[col].alignment = center_align

        # ชื่อ - สกุล (Merge คอลัมน์ C และ D ตามสั่ง)
        ws.merge_cells('C8:D11')
        ws['C8'] = "ชื่อ - สกุล"
        ws['C8'].alignment = center_align

        # ส่วน ทฤษฎี (E-J)
        ws.merge_cells('E8:J8'); ws['E8'] = "ทฤษฎี..........................หน่วยกิต"
        ws.merge_cells('E9:I9'); ws['E9'] = "คะแนนระหว่างภาค"
        headers_theory = {'E10':"เวลา/อุปกรณ์",'F10':"พฤติกรรม",'G10':"งาน/ทดสอบ",'H10':"สอบกลางภาค",'I10':"สอบปลายภาค",'J10':"คะแนนรวม"}
        for cell, val in headers_theory.items(): ws[cell] = val; ws[cell].alignment = rotate_align
        ws.merge_cells('J9:J10')

        # ส่วน ปฏิบัติ (K-P)
        ws.merge_cells('K8:P8'); ws['K8'] = "ปฏิบัติ..................หน่วยกิต"
        ws.merge_cells('K9:O9'); ws['K9'] = "คะแนนระหว่างภาค"
        headers_practice = {'K10':"คุณภาพของงาน",'L10':"เวลา/อุปกรณ์",'M10':"พฤติกรรม",'N10':"การปฎิบัติงาน",'O10':"สอบทฤษฎีเชิงปฎิบัติ",'P10':"คะแนนรวม"}
        for cell, val in headers_practice.items(): ws[cell] = val; ws[cell].alignment = rotate_align
        ws.merge_cells('P9:P10')

        # ระดับ/หมายเหตุ
        ws.merge_cells('Q8:Q9'); ws['Q8'] = "ระดับ"; ws['Q10'] = "คะแนน"; ws.merge_cells('Q10:Q11')
        ws.merge_cells('R8:R11'); ws['R8'] = "หมายเหตุ"

        # คะแนนเต็ม (แถว 11)
        pts = {5:'10',6:'10',7:'20',8:'20',9:'40',10:'100',11:'20',12:'10',13:'10',14:'40',15:'20',16:'100'}
        for c, v in pts.items(): ws.cell(row=11, column=c).value = v; ws.cell(row=11, column=c).alignment = center_align

        # ตีกรอบหัวตาราง
        for r in range(8, 12):
            for c in range(1, 19):
                cell = ws.cell(row=r, column=c); cell.border = border; cell.font = f_bold

        # --- 4. ความกว้างคอลัมน์ตามสั่ง ---
        ws.column_dimensions['A'].width = 3.86
        ws.column_dimensions['B'].width = 12.29
        ws.column_dimensions['C'].width = 10 # แบ่งความกว้าง C และ D
        ws.column_dimensions['D'].width = 7  # รวม C+D ประมาณ 17
        for c_idx in ['E','F','G','H','I','J','L','M','N','O']: ws.column_dimensions[c_idx].width = 2.86
        ws.column_dimensions['K'].width = 3.29
        ws.column_dimensions['P'].width = 3.29
        ws.column_dimensions['Q'].width = 5.86
        ws.column_dimensions['R'].width = 8

        # --- 5. ใส่รายชื่อนักศึกษา ---
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 11 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            # Merge C และ D สำหรับชื่อนักศึกษาทุกแถว
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=4)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 19):
                ws.cell(row=curr, column=c).border = border
                ws.cell(row=curr, column=c).alignment = center_align if c != 3 else Alignment(horizontal='left', indent=1)

    wb.save(output); return output.getvalue()

# --- ส่วน UI Streamlit ---
st.title("🏫 ระบบจัดการข้อมูลวิทยาลัย (ฉบับปรับปรุงตำแหน่ง)")
# ... (Tab 1 และ Tab 2 เหมือนเดิม) ...

t1, t2, t3 = st.tabs(["📝 ลงทะเบียน", "🔍 แก้ไขข้อมูล", "📥 ดาวน์โหลดเอกสาร"])
# (โค้ด Tab 1, 2 ข้ามไปเพื่อความกระชับ)

with t3:
    st.subheader("📥 ดาวน์โหลดฟอร์มบัญชีผลการเรียน")
    st.info("โลโก้อยู่กึ่งกลางหน้า และคอลัมน์ชื่อ-สกุล ถูกรวม (Merge) ระหว่าง C และ D เรียบร้อยแล้ว")
    col1, col2 = st.columns(2)
    with col1:
        g1 = create_grade_report("ปี1")
        if g1: st.download_button("📊 โหลดฟอร์มเกรด ปี 1", g1, "Grade_P1_Centered.xlsx", type="primary")
    with col2:
        g2 = create_grade_report("ปี2")
        if g2: st.download_button("📊 โหลดฟอร์มเกรด ปี 2", g2, "Grade_P2_Centered.xlsx", type="primary")
