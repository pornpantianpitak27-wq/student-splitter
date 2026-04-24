import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าหน้ากระดาษและเชื่อมต่อ GSheets ---
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

# --- 2. ฟังก์ชันจัดการรูปภาพโลโก้ ---
def get_logo_image():
    logo_filename = "logo_college.jpg"
    if os.path.exists(logo_filename):
        try:
            return XLImage(logo_filename)
        except: return None
    return None

# --- 3. ส่วนแสดงผล UI หน้าเว็บ ---
st.title("🏫 ระบบจัดการข้อมูลนักศึกษาและใบรายชื่อ")

# ตรวจสอบโลโก้ก่อนเริ่ม
if os.path.exists("logo_college.jpg"):
    st.success("✅ ตรวจพบไฟล์โลโก้ในระบบ พร้อมใช้งานสำหรับ Excel")
else:
    st.warning("⚠️ ไม่พบไฟล์ logo_college.jpg ใน GitHub (รูปจะไม่ขึ้นใน Excel)")

# สร้าง Tabs สำหรับแบ่งการทำงาน
tab1, tab2, tab3 = st.tabs(["📝 ลงทะเบียนใหม่", "🔍 ตรวจสอบ/แก้ไขข้อมูล", "📥 ดาวน์โหลดใบรายชื่อ"])

# --- TAB 1: ลงทะเบียนใหม่ ---
with tab1:
    with st.form("reg_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            batch = st.text_input("รุ่น")
            sid = st.text_input("รหัสนักศึกษา")
        with c2:
            fname = st.text_input("ชื่อ")
            lname = st.text_input("นามสกุล")
        with c3:
            level = st.selectbox("ระดับชั้น", ["ปี1", "ปี2"])
            prefix = "O1" if level == "ปี1" else "O2"
            room = st.selectbox("ห้อง", [f"{prefix}/{i}" for i in range(1, 16)])
            
        if st.form_submit_button("💾 บันทึกข้อมูล"):
            if sid and fname:
                df = load_data()
                new_row = pd.DataFrame([{"รุ่น": f"'{batch}", "รหัสนักศึกษา": f"'{sid}", "ชื่อ": fname.strip(), "นามสกุล": lname.strip(), "ระดับชั้น": level, "Room": room}])
                updated_df = pd.concat([df, new_row], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success("บันทึกข้อมูลเรียบร้อยแล้ว!")
                st.rerun()

# --- TAB 2: ตรวจสอบและแก้ไขข้อมูล ---
with tab2:
    df_edit = load_data()
    if not df_edit.empty:
        st.write("### รายชื่อนักศึกษาทั้งหมด")
        edited_df = st.data_editor(df_edit, num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกการเปลี่ยนแปลงทั้งหมด"):
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=edited_df)
            st.success("อัปเดตข้อมูลเรียบร้อย!")
            st.rerun()
    else:
        st.info("ยังไม่มีข้อมูลนักศึกษาในระบบ")

# --- TAB 3: ฟังก์ชันสร้าง Excel และดาวน์โหลด ---
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
    center_align = Alignment(horizontal='center', vertical='center')

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # หัวกระดาษ (O ถึง V)
        ws.merge_cells('O2:V2'); ws['O2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['O2'].border = border; ws['O2'].alignment = center_align; ws['O2'].font = bold_font
        ws.merge_cells('O3:P4'); ws['O3'] = "เช็คชื่อนักศึกษา"; ws['O3'].border = border; ws['O3'].alignment = center_align
        ws.merge_cells('Q3:S4'); ws['Q3'] = "เซ็นสอบกลางภาค"; ws['Q3'].border = border; ws['Q3'].alignment = center_align
        ws.merge_cells('T3:V4'); ws['T3'] = "เซ็นสอบปลายภาค"; ws['T3'].border = border; ws['T3'].alignment = center_align

        # หัวข้อหลัก
        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align
        
        # หัวตาราง Layout C-D-E-V
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        ws['E8'] = "เดือน"; ws['E9'] = "วันที่"; ws['E10'] = "คาบ"
        ws.merge_cells('V8:V10'); ws['V8'] = "หมายเหตุ"

        for i in range(1, 17): ws.cell(row=10, column=5+i).value = i
        for r in range(8, 11):
            for c in range(1, 23):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # ข้อมูลรายชื่อ
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=4)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 23):
                cell = ws.cell(row=curr, column=c)
                cell.border = border; cell.alignment = center_align
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)

        # ปรับความกว้าง
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 14
        for c_idx in range(5, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 3.5
        ws.column_dimensions['V'].width = 12

        # แทรกโลโก้
        img = get_logo_image()
        if img:
            img.width, img.height = 80, 80
            ws.add_image(img, 'H1')

    wb.save(output)
    return output.getvalue()

with tab3:
    c1, c2 = st.columns(2)
    with c1:
        f1 = create_excel_report("ปี1")
        if f1: st.download_button("📥 โหลดใบรายชื่อ ปี 1", f1, "P1_Sheet.xlsx", use_container_width=True)
    with c2:
        f2 = create_excel_report("ปี2")
        if f2: st.download_button("📥 โหลดใบรายชื่อ ปี 2", f2, "P2_Sheet.xlsx", use_container_width=True)
