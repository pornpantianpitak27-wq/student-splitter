import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection
import re

st.set_page_config(page_title="ระบบจัดการใบรายชื่อนักศึกษา", layout="wide")

# --- 1. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    # ดึงข้อมูลจาก Sheets
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    
    # แก้ปัญหาจุดทศนิยม (.0) และค่าว่าง (NaN) ทันทีที่โหลดข้อมูล
    # บังคับให้คอลัมน์เหล่านี้เป็น String (ข้อความ) ทั้งหมด
    cols_to_fix = ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']
    for col in cols_to_fix:
        if col in data.columns:
            # แปลงเป็น string -> ตัด .0 ออก (ถ้ามี) -> เปลี่ยน nan เป็นค่าว่าง
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].replace('nan', '')
    return data

# โหลดข้อมูลมาเก็บไว้ในตัวแปร df
df = load_data()

# --- ตัวเลือกสำหรับเมนูต่างๆ ---
CLASSES = ["ปี1", "ปี2"]
# รายชื่อห้องเรียนตามไฟล์ที่คุณให้มา (O1/1 ถึง O1/11)
ROOMS = [f"O1/{i}" for i in range(1, 12)]

st.title("📑 ระบบจัดการและย้ายห้องนักศึกษาออนไลน์")

# --- 2. ส่วนเพิ่มข้อมูลนักศึกษาใหม่ ---
with st.expander("➕ เพิ่มรายชื่อนักศึกษาใหม่", expanded=False):
    with st.form("add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_batch = st.text_input("รุ่น (เช่น 23)")
            new_id = st.text_input("รหัสนักศึกษา")
        with c2:
            new_name = st.text_input("ชื่อ-นามสกุล")
            new_level = st.selectbox("ระดับชั้น", CLASSES)
        with c3:
            new_room = st.selectbox("ห้องเรียน", ROOMS)
        
        submit_btn = st.form_submit_button("บันทึกข้อมูล")
        
        if submit_btn:
            if new_id and new_name:
                # สร้างข้อมูลใหม่ (ล้างจุดทศนิยมตั้งแต่ตอนบันทึก)
                new_entry = pd.DataFrame([{
                    "รุ่น": str(new_batch).replace('.0', ''),
                    "รหัสนักศึกษา": str(new_id).replace('.0', ''),
                    "ชื่อ-นามสกุล": new_name,
                    "ระดับชั้น": new_level,
                    "Room": new_room
                }])
                updated_df = pd.concat([df, new_entry], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success(f"บันทึกข้อมูลคุณ {new_name} เรียบร้อย!")
                st.rerun()
            else:
                st.warning("กรุณากรอกรหัสและชื่อนักศึกษา")

# --- 3. ส่วนค้นหาและแก้ไข (การย้ายห้อง) ---
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล (ย้ายห้องเรียน)")
search_term = st.text_input("พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    # กรองข้อมูลตามคำค้นหา
    mask = df['ชื่อ-นามสกุล'].str.contains(search_term, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search_term, case=False, na=False)
    filtered_df = df[mask]
    
    # ใช้ data_editor เพื่อให้กดแก้ไขห้องในตารางได้เลย
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "รหัสนักศึกษา": st.column_config.TextColumn() # ป้องกันรหัสเพี้ยนในตาราง
        },
        num_rows="dynamic",
        key="editor"
    )
    
    if st.button("💾 ยืนยันการแก้ไข/ย้ายห้อง"):
        # อัปเดตข้อมูลจากตารางแก้ไขกลับไปที่ DataFrame หลัก
        df.update(edited_df)
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=df)
        st.success("อัปเดตข้อมูลนักศึกษาเรียบร้อยแล้ว!")
        st.rerun()

# --- 4. ส่วนการออกใบรายชื่อ Excel ---
st.divider()
if st.button("🖨️ ออกใบรายชื่อ (แยกตามห้องปัจจุบัน)"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        thin = Side(border_style="thin")
        
        # วนลูปสร้าง Sheet ตามห้องเรียนที่มีข้อมูลจริง
        for r_name in sorted(df['Room'].unique()):
            if not r_name: continue
            
            ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
            room_data = df[df['Room'] == r_name].sort_values('รหัสนักศึกษา')
            
            # วาดหัวฟอร์ม (ปรับให้ตรงตามตัวอย่างที่คุณต้องการ)
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True, size=14)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            # หัวตาราง
            headers = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'คาบ {i+1}' for i in range(16)] + ['หมายเหตุ']
            ws.append(headers)
            
            # ใส่รายชื่อนักศึกษา
            for idx, row in enumerate(room_data.itertuples(), 1):
                ws.append([idx, row.รหัสนักศึกษา, row._3] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
            
            # จัดฟอร์แมตเส้นตาราง
            for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=20):
                for cell in r:
                    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    cell.alignment = Alignment(horizontal='center')
            
            ws.column_dimensions['C'].width = 30 # ปรับความกว้างคอลัมน์ชื่อ

        wb.save(output)
        st.download_button("💾 ดาวน์โหลดใบรายชื่อนักศึกษา.xlsx", output.getvalue(), "Student_Attendance.xlsx")
