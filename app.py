import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการข้อมูลนักศึกษา", layout="wide")

# --- 1. เตรียมฐานข้อมูลจำลอง (Session State) ---
if 'student_db' not in st.session_state:
    st.session_state.student_db = pd.DataFrame(columns=[
        'รุ่น', 'รหัส4ตัว', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room'
    ])

# ตัวเลือกข้อมูล
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการและแยกรายชื่อนักศึกษา")

# --- 2. ส่วนการกรอกข้อมูลใหม่ (Input Section) ---
with st.expander("➕ เพิ่มข้อมูลนักศึกษาใหม่", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        batch = st.text_input("รุ่น (เช่น 67)")
        student_id = st.text_input("รหัส4ตัว", max_chars=4)
    with col2:
        full_name = st.text_input("ชื่อ-นามสกุล")
        level = st.selectbox("ระดับชั้น", CLASSES)
    with col3:
        room = st.selectbox("ห้องเรียน", ROOMS)
        
    if st.button("บันทึกข้อมูล"):
        if batch and student_id and full_name:
            new_data = {
                'รุ่น': batch,
                'รหัส4ตัว': student_id,
                'ชื่อ-นามสกุล': full_name,
                'ระดับชั้น': level,
                'Room': room
            }
            st.session_state.student_db = pd.concat([st.session_state.student_db, pd.DataFrame([new_data])], ignore_index=True)
            st.success(f"บันทึกข้อมูลคุณ {full_name} เรียบร้อย!")
        else:
            st.warning("กรุณากรอกข้อมูลให้ครบถ้วน")

st.divider()

# --- 3. ส่วนการค้นหาและแก้ไข (Search & Edit) ---
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล")
search_query = st.text_input("ค้นหาด้วยชื่อ หรือ รหัส")

if not st.session_state.student_db.empty:
    # กรองข้อมูลตามคำค้นหา
    mask = st.session_state.student_db['ชื่อ-นามสกุล'].str.contains(search_query) | \
           st.session_state.student_db['รหัส4ตัว'].str.contains(search_query)
    search_results = st.session_state.student_db[mask]

    if not search_results.empty:
        # ใช้ data_editor ของ Streamlit เพื่อให้แก้ไขข้อมูลในตารางได้เลย
        edited_df = st.data_editor(
            search_results, 
            num_rows="dynamic",
            column_config={
                "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
                "Room": st.column_config.SelectboxColumn(options=ROOMS)
            },
            key="data_editor"
        )
        
        # ปุ่มยืนยันการแก้ไข
        if st.button("ยืนยันการแก้ไข/ย้ายห้อง"):
            st.session_state.student_db.update(edited_df)
            st.success("อัปเดตข้อมูลสำเร็จ!")
    else:
        st.info("ไม่พบข้อมูลที่ค้นหา")

st.divider()

# --- 4. ส่วนการส่งออก Excel แยกห้อง (Export) ---
st.subheader("📥 ส่งออกข้อมูล (แยก Sheet ตามห้อง)")

if st.button("เตรียมไฟล์ Excel สำหรับดาวน์โหลด"):
    if not st.session_state.student_db.empty:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # ดึงรายชื่อห้องที่มีข้อมูลจริงๆ
            available_rooms = sorted(st.session_state.student_db['Room'].unique())
            
            for r in available_rooms:
                room_df = st.session_state.student_db[st.session_state.student_db['Room'] == r].copy()
                room_df.insert(0, 'ลำดับ', range(1, len(room_df) + 1))
                
                sheet_name = f"ห้อง {r.replace('/', '-')}"
                room_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="รายชื่อแยกห้อง_อัปเดตล่าสุด.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("ยังไม่มีข้อมูลนักศึกษาในระบบ")

# แสดงตารางรวมทั้งหมดด้านล่าง
st.write("---")
st.write("📋 ข้อมูลทั้งหมดในระบบตอนนี้:")
st.dataframe(st.session_state.student_db)