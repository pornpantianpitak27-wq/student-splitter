import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="ระบบจัดการใบรายชื่อนักศึกษา V3", layout="wide")

# --- 1. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    # ดึงข้อมูลจาก Sheets
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    
    # ล้างข้อมูลขยะ (จุดทศนิยม และ nan)
    cols_to_fix = ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']
    for col in cols_to_fix:
        if col in data.columns:
            # แปลงเป็นข้อความ -> ลบ .0 -> ลบเครื่องหมาย ' ที่เราใส่ไว้ตอนบันทึก (ถ้ามี)
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "")
            data[col] = data[col].replace('nan', '')
    return data

df = load_data()

# --- ตัวเลือกเมนู ---
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("📑 ระบบจัดการรายชื่อนักศึกษา (รองรับเลข 0 นำหน้า)")

# --- 2. ส่วนเพิ่มข้อมูลนักศึกษาใหม่ ---
with st.expander("➕ เพิ่มรายชื่อนักศึกษาใหม่", expanded=False):
    with st.form("add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_batch = st.text_input("รุ่น")
            new_id = st.text_input("รหัสนักศึกษา (กรอกเลข 0 นำหน้าได้)")
        with c2:
            new_name = st.text_input("ชื่อ-นามสกุล")
            new_level = st.selectbox("ระดับชั้น", CLASSES)
        with c3:
            new_room = st.selectbox("ห้องเรียน", ROOMS)
        
        submit_btn = st.form_submit_button("💾 บันทึกข้อมูล")
        
        if submit_btn:
            if new_id and new_name:
                # เคล็ดลับ: ใส่ ' นำหน้าข้อมูลที่เป็นตัวเลข เพื่อไม่ให้ Google Sheets ตัดเลข 0
                new_entry = pd.DataFrame([{
                    "รุ่น": f"'{new_batch}",
                    "รหัสนักศึกษา": f"'{new_id}", 
                    "ชื่อ-นามสกุล": new_name,
                    "ระดับชั้น": new_level,
                    "Room": new_room
                }])
                updated_df = pd.concat([df, new_entry], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                st.success(f"บันทึกรหัส {new_id} เรียบร้อย!")
                st.rerun()
            else:
                st.warning("กรุณากรอกข้อมูลให้ครบถ้วน")

# --- 3. ส่วนค้นหาและแก้ไข ---
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล")
search_term = st.text_input("พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search_term, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search_term, case=False, na=False)
    filtered_df = df[mask]
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "รหัสนักศึกษา": st.column_config.TextColumn()
        },
        num_rows="dynamic",
        key="editor_v3"
    )
    
    if st.button("✅ บันทึกการเปลี่ยนแปลง"):
        # ก่อนบันทึกคืนค่า ต้องใส่ ' กลับเข้าไปเพื่อรักษาเลข 0
        for col in ['รุ่น', 'รหัสนักศึกษา']:
            edited_df[col] = edited_df[col].apply(lambda x: f"'{x}" if not str(x).startswith("'") else x)
        
        df.update(edited_df)
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=df)
        st.success("อัปเดตข้อมูลสำเร็จ!")
        st.rerun()

# --- 4. ส่วนการออกใบรายชื่อ Excel ---
st.divider()
if st.button("🖨️ ออกใบรายชื่อ Excel"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        thin = Side(border_style="thin")
        
        for r_name in sorted(df['Room'].unique()):
            if not r_name: continue
            ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
            room_data = df[df['Room'] == r_name].sort_values('รหัสนักศึกษา')
            
            # หัวตาราง
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True, size=14)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            ws.append(['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'คาบ {i+1}' for i in range(16)] + ['หมายเหตุ'])
            
            for idx, row in enumerate(room_data.itertuples(), 1):
                # ใน Excel บังคับรหัสเป็นข้อความ
                ws.append([idx, f"{row.รหัสนักศึกษา}", row._3] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
            
            for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=20):
                for cell in r:
                    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    cell.alignment = Alignment(horizontal='center')
            
            ws.column_dimensions['C'].width = 30
            ws.column_dimensions['B'].width = 15

        wb.save(output)
        st.download_button("💾 ดาวน์โหลดไฟล์ .xlsx", output.getvalue(), "Attendance_List.xlsx")
