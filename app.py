import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side

st.set_page_config(page_title="ระบบจัดการใบรายชื่อนักศึกษา", layout="wide")

# --- ตั้งค่าข้อมูลเบื้องต้น ---
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

# --- 1. เตรียมฐานข้อมูล (Session State) ---
if 'student_db' not in st.session_state:
    st.session_state.student_db = pd.DataFrame(columns=[
        'รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room'
    ])

st.title("📋 ระบบจัดการข้อมูลและออกใบรายชื่อ (แยกห้อง)")

# --- 2. ส่วนการกรอกข้อมูล (Input) ---
with st.expander("➕ เพิ่มนักศึกษาใหม่", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        batch = st.text_input("รุ่น")
        student_id = st.text_input("รหัสนักศึกษา")
    with c2:
        name = st.text_input("ชื่อ-นามสกุล")
        level = st.selectbox("ระดับชั้น", CLASSES)
    with c3:
        room = st.selectbox("ห้องเรียน", ROOMS)
    
    if st.button("บันทึกข้อมูล"):
        if student_id and name:
            new_row = {'รุ่น': batch, 'รหัสนักศึกษา': student_id, 'ชื่อ-นามสกุล': name, 'ระดับชั้น': level, 'Room': room}
            st.session_state.student_db = pd.concat([st.session_state.student_db, pd.DataFrame([new_row])], ignore_index=True)
            st.success(f"เพิ่ม {name} เข้าห้อง {room} เรียบร้อย")
        else:
            st.error("กรุณากรอก รหัส และ ชื่อ-นามสกุล")

st.divider()

# --- 3. ค้นหาและแก้ไข (Search & Edit) ---
st.subheader("🔍 ค้นหาและย้ายห้องเรียน")
search = st.text_input("พิมพ์ชื่อเพื่อค้นหา...")
if not st.session_state.student_db.empty:
    filtered_df = st.session_state.student_db[st.session_state.student_db['ชื่อ-นามสกุล'].str.contains(search, na=False)]
    
    # ใช้ Data Editor เพื่อให้กดเปลี่ยนห้องได้เลย
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS)
        },
        num_rows="dynamic"
    )
    
    if st.button("ยืนยันการแก้ไขข้อมูล"):
        st.session_state.student_db.update(edited_df)
        st.success("อัปเดตข้อมูลเรียบร้อยแล้ว!")

st.divider()

# --- 4. ฟังก์ชันส่งออกเป็น "ใบรายชื่อ" (Excel Export) ---
def create_attendance_sheet(df):
    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # ลบ sheet เปล่าที่สร้างมาตอนแรก
    
    rooms = sorted(df['Room'].unique())
    
    for r in rooms:
        ws = wb.create_sheet(title=f"ห้อง {r.replace('/', '-')}")
        room_data = df[df['Room'] == r].copy()
        
        # --- สร้างหัวกระดาษ (Header) ตามตัวอย่าง ---
        ws.merge_cells('A1:U1')
        ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        ws.merge_cells('A2:U2')
        ws['A2'] = f"ระดับ ปวส. ชั้น {room_data['ระดับชั้น'].iloc[0] if not room_data.empty else ''} ห้อง {r}"
        ws['A2'].alignment = Alignment(horizontal='center')
        
        # ส่วนของหัวตาราง
        headers = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-สกุล', 'หมายเหตุ']
        ws.append(headers)
        
        # ใส่ข้อมูลนักศึกษา
        for i, row in enumerate(room_data.itertuples(), 1):
            ws.append([i, row.รหัสนักศึกษา, row._3, f"รุ่น {row.รุ่น}"]) # _3 คือ ชื่อ-นามสกุล
            
        # ตกแต่งเส้นตาราง (Border)
        thin = Side(border_style="thin", color="000000")
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=4):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                
    wb.save(output)
    return output.getvalue()

# --- 5. ปุ่มดาวน์โหลด ---
if st.button("🖨️ ออกใบรายชื่อแยกห้อง (Excel)"):
    if not st.session_state.student_db.empty:
        excel_file = create_attendance_sheet(st.session_state.student_db)
        st.download_button(
            label="💾 กดดาวน์โหลดใบรายชื่อ",
            data=excel_file,
            file_name="ใบรายชื่อนักศึกษา.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("ยังไม่มีข้อมูลให้ส่งออก")

# แสดงข้อมูลดิบด้านล่าง
st.write("📊 ข้อมูลทั้งหมดในระบบ:", st.session_state.student_db)