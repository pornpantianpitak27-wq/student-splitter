import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="ระบบจัดการใบรายชื่อ (V2.1)", layout="wide")

# --- 1. เชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    # ดึงข้อมูลจาก Google Sheets
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    # วิธีแก้ไขที่ 2: แปลงคอลัมน์สำคัญให้เป็นข้อความ (String) ทันทีที่โหลด
    # เพื่อป้องกัน Error .str.contains กับข้อมูลที่เป็นตัวเลข
    for col in ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']:
        if col in data.columns:
            data[col] = data[col].astype(str).replace('nan', '')
    return data

df = load_data()

# --- ข้อมูลตัวเลือก ---
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการข้อมูลนักศึกษาออนไลน์ (ฉบับสมบูรณ์)")

# --- 2. ส่วนกรอกข้อมูลนักศึกษาใหม่ ---
with st.expander("➕ เพิ่มนักศึกษาใหม่", expanded=False):
    with st.form("input_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            batch = st.text_input("รุ่น")
            std_id = st.text_input("รหัสนักศึกษา")
        with c2:
            name = st.text_input("ชื่อ-นามสกุล")
            level = st.selectbox("ระดับชั้น", CLASSES)
        with c3:
            room = st.selectbox("ห้องเรียน", ROOMS)
        
        submit = st.form_submit_button("💾 บันทึกข้อมูลลง Google Sheets")
        
        if submit:
            if std_id and name:
                new_row = {
                    "รุ่น": str(batch),
                    "รหัสนักศึกษา": str(std_id),
                    "ชื่อ-นามสกุล": str(name),
                    "ระดับชั้น": str(level),
                    "Room": str(room)
                }
                new_data = pd.DataFrame([new_row])
                updated_df = pd.concat([df, new_data], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success(f"บันทึกคุณ {name} เรียบร้อยแล้ว!")
                st.rerun()
            else:
                st.error("⚠️ กรุณากรอกรหัสและชื่อนักศึกษา")

# --- 3. ส่วนค้นหาและแก้ไข (ย้ายห้อง) ---
st.subheader("🔍 ค้นหาและจัดการรายชื่อ")
search = st.text_input("🔎 พิมพ์ชื่อหรือรหัสนักศึกษาเพื่อค้นหา...")

if not df.empty:
    # ค้นหาโดยไม่สนว่าเป็นตัวพิมพ์เล็ก/ใหญ่ และรองรับค่าว่าง
    mask = df['ชื่อ-นามสกุล'].str.contains(search, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search, case=False, na=False)
    filtered_df = df[mask]
    
    st.write(f"พบข้อมูล {len(filtered_df)} รายการ")
    
    # ตารางแก้ไขข้อมูล
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "รหัสนักศึกษา": st.column_config.TextColumn() # บังคับให้เป็น Text
        },
        num_rows="dynamic",
        key="main_editor"
    )
    
    if st.button("✅ ยืนยันการแก้ไขและบันทึก"):
        # นำข้อมูลที่แก้ไปทับใน DataFrame หลัก
        df.update(edited_df)
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=df)
        st.success("อัปเดตข้อมูลในระบบเรียบร้อย!")
        st.rerun()
else:
    st.info("ยังไม่มีข้อมูลในระบบ กรุณาเพิ่มรายชื่อนักศึกษา")

# --- 4. ส่วนส่งออกใบรายชื่อ (Excel) ---
st.divider()
if st.button("🖨️ ออกใบรายชื่อ Excel (แยก Sheet ตามห้อง)"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active) # ลบ Sheet เปล่าที่ติดมาตอนสร้าง
        
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        # วนลูปสร้างทีละห้อง
        for room_name in sorted(df['Room'].unique()):
            if not room_name: continue
            
            ws = wb.create_sheet(title=f"ห้อง {room_name.replace('/', '-')}")
            room_students = df[df['Room'] == room_name].sort_values('รหัสนักศึกษา')
            
            # --- ส่วนหัว (Header) ---
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True, size=14)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            ws.append(['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'คาบ {i+1}' for i in range(16)] + ['หมายเหตุ'])
            
            # --- ใส่รายชื่อ ---
            for i, row in enumerate(room_students.itertuples(), 1):
                ws.append([i, row.รหัสนักศึกษา, row._3] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
            
            # --- จัดรูปแบบเส้นตาราง ---
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=20):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center')
            
            # ปรับความกว้างคอลัมน์ชื่อ
            ws.column_dimensions['C'].width = 30
            ws.column_dimensions['B'].width = 15

        wb.save(output)
        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ใบรายชื่อ .xlsx",
            data=output.getvalue(),
            file_name="ใบรายชื่อนักศึกษา_2568.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
