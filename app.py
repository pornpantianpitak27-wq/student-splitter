import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าหน้ากระดาษ (ต้องอยู่บนสุด) ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา", layout="wide")

# --- 2. เชื่อมต่อข้อมูล ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # ล้างข้อมูลขยะ
        for col in ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']:
            if col in data.columns:
                data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
                data[col] = data[col].str.replace("'", "")
                data[col] = data[col].replace('nan', '')
        return data
    except Exception as e:
        st.error(f"ไม่สามารถโหลดข้อมูลจาก Google Sheets ได้: {e}")
        return pd.DataFrame()

df = load_data()
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการข้อมูลนักศึกษา")

# --- 3. ส่วนกรอกข้อมูลใหม่ (Form) ---
# ตรวจสอบว่ามีส่วนนี้ในโค้ดของคุณหรือไม่
st.subheader("➕ เพิ่มรายชื่อนักศึกษาใหม่")
with st.form("add_student_form", clear_on_submit=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        new_batch = st.text_input("รุ่น")
        new_id = st.text_input("รหัสนักศึกษา")
    with c2:
        new_name = st.text_input("ชื่อ-นามสกุล")
        new_level = st.selectbox("ระดับชั้น", CLASSES)
    with c3:
        new_room = st.selectbox("ห้องเรียน", ROOMS)
    
    submit_btn = st.form_submit_button("💾 บันทึกรายชื่อใหม่")
    
    if submit_btn:
        if new_id and new_name:
            new_row = pd.DataFrame([{
                "รุ่น": f"'{new_batch}",
                "รหัสนักศึกษา": f"'{new_id}",
                "ชื่อ-นามสกุล": new_name,
                "ระดับชั้น": new_level,
                "Room": new_room
            }])
            updated_df = pd.concat([df, new_row], ignore_index=True)
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
            st.success(f"บันทึกข้อมูลคุณ {new_name} เรียบร้อย!")
            st.rerun()
        else:
            st.warning("กรุณากรอกข้อมูลให้ครบถ้วน")

# --- 4. ส่วนค้นหาและย้ายห้อง (Editor) ---
st.divider()
st.subheader("🔍 ค้นหาและจัดการการย้ายห้อง")
search = st.text_input("🔎 พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search, case=False, na=False)
    filtered_df = df[mask].copy()
    filtered_df['สาเหตุการย้าย'] = "" # คอลัมน์ชั่วคราว
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "รหัสนักศึกษา": st.column_config.TextColumn(disabled=True),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES)
        },
        key="data_editor_safe"
    )
    
    if st.button("✅ ยืนยันการแก้ไข"):
        # โค้ดบันทึกข้อมูล (เหมือนเวอร์ชันก่อนหน้า)
        # ...
        st.success("อัปเดตข้อมูลสำเร็จ")
        st.rerun()

# --- 5. ส่วนออกไฟล์ Excel ---
st.divider()
if st.button("🖨️ ออกใบรายชื่อ (Excel)"):
    # โค้ดสร้าง Excel ตามฟอร์มต้นฉบับ
    # ...
    st.info("กำลังสร้างไฟล์...")
