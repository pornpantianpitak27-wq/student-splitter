import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side

st.set_page_config(page_title="ระบบจัดการใบรายชื่อ", layout="wide")

# --- ตัวเลือกข้อมูล ---
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

# --- เก็บข้อมูลในเครื่อง (Session State) ---
if 'student_db' not in st.session_state:
    st.session_state.student_db = pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room'])

st.title("🏫 ระบบจัดการข้อมูลนักศึกษาและออกใบรายชื่อ")

# --- ส่วนที่ 1: กรอกข้อมูล ---
st.header("1. เพิ่มข้อมูลนักศึกษา")
with st.container(border=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        batch = st.text_input("รุ่น (เช่น 23)")
        std_id = st.text_input("รหัสนักศึกษา (เช่น 6830...)")
    with c2:
        name = st.text_input("ชื่อ-นามสกุล")
        level = st.selectbox("ระดับชั้น", CLASSES)
    with c3:
        room = st.selectbox("ห้องเรียน", ROOMS)
    
    if st.button("➕ บันทึกข้อมูลลงในระบบ"):
        if std_id and name:
            new_row = {'รุ่น': batch, 'รหัสนักศึกษา': std_id, 'ชื่อ-นามสกุล': name, 'ระดับชั้น': level, 'Room': room}
            st.session_state.student_db = pd.concat([st.session_state.student_db, pd.DataFrame([new_row])], ignore_index=True)
            st.success(f"บันทึก {name} สำเร็จ!")
        else:
            st.error("กรุณากรอกรหัสและชื่อนักศึกษา")

# --- ส่วนที่ 2: ค้นหาและแก้ไข ---
st.header("2. ค้นหาและแก้ไข/ย้ายห้อง")
with st.container(border=True):
    search = st.text_input("🔎 พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")
    if not st.session_state.student_db.empty:
        # กรองข้อมูล
        df = st.session_state.student_db
        show_df = df[df['ชื่อ-นามสกุล'].str.contains(search, na=False) | df['รหัสนักศึกษา'].str.contains(search, na=False)]
        
        # ตารางแก้ไข
        edited_df = st.data_editor(
            show_df,
            column_config={
                "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
                "Room": st.column_config.SelectboxColumn(options=ROOMS)
            },
            num_rows="dynamic",
            key="editor"
        )
        
        if st.button("💾 ยืนยันการแก้ไขข้อมูล"):
            st.session_state.student_db.update(edited_df)
            st.rerun()

# --- ส่วนที่ 3: ส่งออก Excel ---
st.header("3. ดาวน์โหลดใบรายชื่อ (แยกตามห้อง)")
if st.button("🖨️ สร้างไฟล์ใบรายชื่อ Excel"):
    if not st.session_state.student_db.empty:
        # สร้างไฟล์ Excel
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        
        db = st.session_state.student_db
        for r in sorted(db['Room'].unique()):
            ws = wb.create_sheet(title=f"ห้อง {r.replace('/', '-')}")
            room_df = db[db['Room'] == r]
            
            # หัวกระดาษ
            ws.merge_cells('A1:D1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            ws.merge_cells('A2:D2')
            ws['A2'] = f"ชั้น {room_df['ระดับชั้น'].iloc[0]} ห้อง {r}"
            ws['A2'].alignment = Alignment(horizontal='center')
            
            # หัวตาราง
            ws.append(['เลขที่', 'รหัสประจำตัว', 'ชื่อ-สกุล', 'หมายเหตุ'])
            for i, row in enumerate(room_df.itertuples(), 1):
                ws.append([i, row.รหัสนักศึกษา, row._3, f"รุ่น {row.รุ่น}"])
            
            # ใส่เส้นตาราง
            thin = Side(border_style="thin")
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=4):
                for cell in row:
                    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

        wb.save(output)
        st.download_button(label="✅ กดดาวน์โหลดที่นี่", data=output.getvalue(), file_name="ใบรายชื่อ.xlsx")
    else:
        st.warning("กรุณากรอกข้อมูลนักศึกษาก่อน")

if st.button("🗑️ ล้างข้อมูลทั้งหมดในระบบ"):
    st.session_state.student_db = pd.DataFrame(columns=['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room'])
    st.rerun()
