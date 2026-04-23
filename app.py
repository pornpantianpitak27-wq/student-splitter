import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าหน้ากระดาษ (ต้องอยู่บรรทัดบนสุดถัดจาก import) ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - แบบฟอร์มมาตรฐาน", layout="wide")

# --- 2. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # ปรับจูนข้อมูล: ลบ .0, ลบเครื่องหมาย ' และจัดการค่าว่าง
        for col in ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']:
            if col in data.columns:
                data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
                data[col] = data[col].str.replace("'", "")
                data[col] = data[col].replace('nan', '')
        return data
    except Exception as e:
        st.error(f"เชื่อมต่อฐานข้อมูลไม่ได้: {e}")
        return pd.DataFrame()

df = load_data()
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("📑 ระบบบริหารจัดการรายชื่อนักศึกษา")

# --- 3. ส่วนเพิ่มข้อมูลใหม่ ---
with st.expander("➕ ลงทะเบียนนักศึกษาใหม่", expanded=False):
    with st.form("add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_batch = st.text_input("รุ่น (เช่น 23)")
            new_id = st.text_input("รหัสนักศึกษา (กรอก 0 นำหน้าได้)")
        with c2:
            new_name = st.text_input("ชื่อ-นามสกุล")
            new_level = st.selectbox("ระดับชั้น", CLASSES)
        with c3:
            new_room = st.selectbox("ห้องเรียน", ROOMS)
        
        if st.form_submit_button("💾 บันทึกข้อมูล"):
            if new_id and new_name:
                new_entry = pd.DataFrame([{
                    "รุ่น": f"'{new_batch}",
                    "รหัสนักศึกษา": f"'{new_id}",
                    "ชื่อ-นามสกุล": new_name,
                    "ระดับชั้น": new_level,
                    "Room": new_room
                }])
                updated_df = pd.concat([df, new_entry], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success("บันทึกสำเร็จ!")
                st.rerun()

# --- 4. ส่วนค้นหาและแก้ไข (รองรับการค้นหาด้วยรหัสแล้วแก้ได้ตลอด) ---
st.divider()
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล/ย้ายห้อง")
q = st.text_input("🔎 ใส่รหัสนักศึกษาหรือชื่อเพื่อค้นหา...")

if not df.empty:
    mask = df['รหัสนักศึกษา'].str.contains(q, case=False, na=False) | \
           df['ชื่อ-นามสกุล'].str.contains(q, case=False, na=False)
    filtered = df[mask].copy()

    if not filtered.empty:
        edited = st.data_editor(
            filtered,
            column_config={
                "รหัสนักศึกษา": st.column_config.TextColumn("รหัส (ห้ามแก้)", disabled=True),
                "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
                "Room": st.column_config.SelectboxColumn(options=ROOMS)
            },
            key="edit_table",
            use_container_width=True
        )

        if st.button("💾 ยืนยันการแก้ไขที่เลือก"):
            main_df = df.copy()
            for _, row in edited.iterrows():
                # อัปเดตทับข้อมูลเดิมโดยอ้างอิงจากรหัส
                main_df.loc[main_df['รหัสนักศึกษา'] == row['รหัสนักศึกษา'], 
                           ['รุ่น', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']] = \
                    [row['รุ่น'], row['ชื่อ-นามสกุล'], row['ระดับชั้น'], row['Room']]
            
            # ล็อคเลข 0 ก่อนส่งคืน Sheets
            for col in ['รุ่น', 'รหัสนักศึกษา']:
                main_df[col] = main_df[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")
            
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=main_df)
            st.success("อัปเดตข้อมูลเรียบร้อย!")
            st.rerun()

# --- 5. ส่วนส่งออก Excel (ตามฟอร์มต้นฉบับ 16 คาบ) ---
st.divider()
if st.button("🖨️ ออกใบรายชื่อแยกตามห้อง (Excel)"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        side = Side(style='thin')
        border = Border(left=side, right=side, top=side, bottom=side)
        
        for r_name in sorted(df['Room'].unique()):
            if not r_name: continue
            ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
            room_data = df[df['Room'] == r_name].sort_values('รหัสนักศึกษา')
            
            # หัวกระดาษ
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].alignment = Alignment(horizontal='center')
            ws['A1'].font = Font(bold=True, size=14)

            ws.merge_cells('A2:U2')
            ws['A2'] = f"ระดับ ปวส. ห้อง {r_name}"
            ws['A2'].alignment = Alignment(horizontal='center')

            ws.merge_cells('A3:K3')
            ws['A3'] = "วิชา..........................................................."
            ws.merge_cells('L3:U3')
            ws['L3'] = "ผู้สอน........................................................."

            # หัวตาราง
            cols = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ']
            ws.append(cols)
            for cell in ws[4]:
                cell.border = border
                cell.alignment = Alignment(horizontal='center')
                cell.font = Font(bold=True)

            # ข้อมูลนักศึกษา
            for i, row in enumerate(room_data.itertuples(), 1):
                ws.append([i, row.รหัสนักศึกษา, row._3] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
                for cell in ws[ws.max_row]:
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
                ws.cell(row=ws.max_row, column=3).alignment = Alignment(horizontal='left')

            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 25
            for c in range(4, 20):
                ws.column_dimensions[ws.cell(row=4, column=c).column_letter].width = 3

        wb.save(output)
        st.download_button("📥 ดาวน์โหลดใบรายชื่อ.xlsx", output.getvalue(), "Student_Attendance.xlsx")
