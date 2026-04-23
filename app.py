import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าเบื้องต้น ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - Complete Version", layout="wide")

# --- 2. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # ปรับแต่งข้อมูลให้พร้อมใช้งาน (ลบ .0 และลบเครื่องหมาย ' ออกชั่วคราวเพื่อการค้นหา)
        for col in ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']:
            if col in data.columns:
                data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
                data[col] = data[col].str.replace("'", "")
                data[col] = data[col].replace('nan', '')
        return data
    except Exception as e:
        st.error(f"ไม่สามารถดึงข้อมูลได้: {e}")
        return pd.DataFrame()

df = load_data()
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบบริหารจัดการรายชื่อนักศึกษา (Full Version)")

# --- 3. เมนูระบบกรอกข้อมูล (High-Stability) ---
st.subheader("➕ ลงทะเบียนนักศึกษาใหม่")
with st.form("stable_add_form", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        new_batch = st.text_input("รุ่น", placeholder="เช่น 23")
        new_id = st.text_input("รหัสนักศึกษา", placeholder="รหัส 11 หลัก")
    with col2:
        new_name = st.text_input("ชื่อ-นามสกุล")
        new_level = st.selectbox("ระดับชั้น", CLASSES)
    with col3:
        new_room = st.selectbox("ห้องเรียน", ROOMS)
        st.caption("ตรวจสอบความถูกต้องก่อนบันทึก")

    if st.form_submit_button("💾 บันทึกข้อมูลใหม่", use_container_width=True):
        if not new_id or not new_name:
            st.error("❌ กรุณากรอกรหัสและชื่อ-นามสกุล")
        elif new_id in df['รหัสนักศึกษา'].values:
            st.warning(f"⚠️ รหัส {new_id} มีในระบบแล้ว โปรดใช้ส่วนแก้ไขด้านล่างเพื่อย้ายห้อง")
        else:
            # บันทึกโดยเติม ' เพื่อรักษาเลข 0
            new_row = pd.DataFrame([{
                "รุ่น": f"'{new_batch}",
                "รหัสนักศึกษา": f"'{new_id}",
                "ชื่อ-นามสกุล": new_name.strip(),
                "ระดับชั้น": new_level,
                "Room": new_room
            }])
            updated_df = pd.concat([df, new_row], ignore_index=True)
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
            st.success(f"✅ บันทึกคุณ {new_name} เรียบร้อย!")
            st.rerun()

# --- 4. ส่วนค้นหาและแก้ไข (ป้องกันชื่อซ้ำด้วยระบบ Update) ---
st.divider()
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล (ย้ายห้อง)")
search_q = st.text_input("🔎 พิมพ์รหัสหรือชื่อเพื่อค้นหาเพื่อแก้ไข...")

if not df.empty:
    mask = df['รหัสนักศึกษา'].str.contains(search_q, case=False, na=False) | \
           df['ชื่อ-นามสกุล'].str.contains(search_q, case=False, na=False)
    filtered_df = df[mask].copy()

    if not filtered_df.empty:
        edited_df = st.data_editor(
            filtered_df,
            column_config={
                "รหัสนักศึกษา": st.column_config.TextColumn("รหัส (ห้ามแก้)", disabled=True),
                "Room": st.column_config.SelectboxColumn("ห้องใหม่", options=ROOMS),
                "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES)
            },
            key="edit_editor",
            use_container_width=True
        )

        if st.button("✅ ยืนยันการแก้ไขข้อมูล"):
            main_db = df.copy()
            for _, row in edited_df.iterrows():
                std_id = row['รหัสนักศึกษา']
                # ค้นหาและเขียนทับบรรทัดเดิม (Update) ป้องกันการเกิดชื่อซ้ำ
                main_db.loc[main_db['รหัสนักศึกษา'] == std_id, 
                           ['รุ่น', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']] = \
                    [row['รุ่น'], row['ชื่อ-นามสกุล'], row['ระดับชั้น'], row['Room']]

            # ล็อค Format เลข 0 ก่อนส่งไป Google Sheets
            for col in ['รุ่น', 'รหัสนักศึกษา']:
                main_db[col] = main_db[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")

            conn.update(spreadsheet=st.secrets["gsheet_url"], data=main_db)
            st.success("✅ อัปเดตข้อมูลเรียบร้อย!")
            st.rerun()

# --- 5. ฟังก์ชันสร้าง Excel แยกชั้นปี (16 คาบ) ---
st.divider()
st.subheader("🖨️ ออกใบรายชื่อ (แยกไฟล์ปี 1 / ปี 2)")

def create_excel(target_year):
    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    
    year_data = df[df['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    for r_name in sorted(year_data['Room'].unique()):
        if not r_name: continue
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # หัวฟอร์ม
        ws.merge_cells('A1:U1')
        ws['A1'] = f"บัญชีรายชื่อนักศึกษา {target_year} ภาคเรียนที่ 1 ปีการศึกษา 2568"
        ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A2:U2'); ws['A2'] = f"ระดับ ปวส. ห้อง {r_name}"
        ws['A2'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A3:K3'); ws['A3'] = "วิชา..........................................................."
        ws.merge_cells('L3:U3'); ws['L3'] = "ผู้สอน........................................................."

        # หัวตาราง
        headers = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ']
        ws.append(headers)
        for cell in ws[4]:
            cell.border = border; cell.alignment = Alignment(horizontal='center'); cell.font = Font(bold=True)

        for i, row in enumerate(room_data.itertuples(), 1):
            ws.append([i, row.รหัสนักศึกษา, row._3] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
            for cell in ws[ws.max_row]: cell.border = border; cell.alignment = Alignment(horizontal='center')
            ws.cell(row=ws.max_row, column=3).alignment = Alignment(horizontal='left', indent=1)

        ws.column_dimensions['B'].width = 15; ws.column_dimensions['C'].width = 25
        for c in range(4, 20): ws.column_dimensions[ws.cell(row=4, column=c).column_letter].width = 3.5
            
    wb.save(output)
    return output.getvalue()

c1, c2 = st.columns(2)
with c1:
    file1 = create_excel("ปี1")
    if file1: st.download_button("📥 ดาวน์โหลดไฟล์ ปี 1", file1, "รายชื่อ_ปี1.xlsx", key="p1")
with c2:
    file2 = create_excel("ปี2")
    if file2: st.download_button("📥 ดาวน์โหลดไฟล์ ปี 2", file2, "รายชื่อ_ปี2.xlsx", key="p2")
