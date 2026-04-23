import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="ระบบจัดการใบรายชื่อ", layout="wide")

# --- เชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    return conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)

df = load_data()

# --- ข้อมูลตัวเลือก ---
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการข้อมูลนักศึกษาออนไลน์")

# --- 1. ส่วนกรอกข้อมูล ---
with st.expander("➕ เพิ่มนักศึกษาใหม่", expanded=True):
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
        
        submit = st.form_submit_button("บันทึกข้อมูลลง Google Sheets")
        
        if submit:
            if std_id and name:
                new_data = pd.DataFrame([{"รุ่น": batch, "รหัสนักศึกษา": std_id, "ชื่อ-นามสกุล": name, "ระดับชั้น": level, "Room": room}])
                updated_df = pd.concat([df, new_data], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success("บันทึกสำเร็จ! ข้อมูลถูกส่งไปที่ Google Sheets แล้ว")
                st.rerun()
            else:
                st.error("กรุณากรอกข้อมูลให้ครบ")

# --- 2. ส่วนค้นหาและย้ายห้อง ---
st.subheader("🔍 ค้นหาและย้ายห้องเรียน")
search = st.text_input("พิมพ์ชื่อเพื่อค้นหา...")
if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search, na=False) | df['รหัสนักศึกษา'].str.contains(search, na=False)
    filtered_df = df[mask]
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS)
        },
        num_rows="dynamic",
        key="data_editor"
    )
    
    if st.button("💾 ยืนยันการแก้ไขทั้งหมด"):
        df.update(edited_df)
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=df)
        st.success("อัปเดตข้อมูลใน Google Sheets เรียบร้อย!")
        st.rerun()

# --- 3. ส่วนออกใบรายชื่อ Excel ---
st.divider()
if st.button("🖨️ สร้างไฟล์ใบรายชื่อ Excel (แยกห้อง)"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        thin = Side(border_style="thin")
        
        for r in sorted(df['Room'].unique()):
            ws = wb.create_sheet(title=f"ห้อง {r.replace('/', '-')}")
            room_df = df[df['Room'] == r]
            
            # หัวกระดาษ
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True, size=14)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            # ตารางเช็คชื่อ (วาดโครง)
            ws.append(['เลขที่', 'รหัสประจำตัว', 'ชื่อ-สกุล'] + ['' for _ in range(17)] + ['หมายเหตุ'])
            ws.append(['', '', '', 'วันที่'] + ['' for _ in range(16)] + [''])
            
            for i, row in enumerate(room_df.itertuples(), 1):
                ws.append([i, row.รหัสนักศึกษา, row._3] + ['' for _ in range(17)] + [f"รุ่น {row.รุ่น}"])
            
            # ใส่เส้นตาราง
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=21):
                for cell in row:
                    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

        wb.save(output)
        st.download_button("💾 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "Attendance_List.xlsx")
