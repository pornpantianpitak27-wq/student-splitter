import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา ปวส.", layout="wide")

# --- 2. การเชื่อมต่อฐานข้อมูล Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        for col in ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']:
            if col in data.columns:
                data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
                data[col] = data[col].str.replace("'", "")
                data[col] = data[col].replace('nan', '')
        return data
    except Exception as e:
        st.error(f"ไม่สามารถโหลดข้อมูลได้: {e}")
        return pd.DataFrame()

df = load_data()

# --- ตั้งค่ารายชื่อห้องเรียนแยกตามชั้นปี ---
ROOMS_P1 = [f"O1/{i}" for i in range(1, 16)] # ปี 1 ใช้ O1/1 - O1/15
ROOMS_P2 = [f"O2/{i}" for i in range(1, 16)] # ปี 2 ใช้ O2/1 - O2/15

st.title("📑 ระบบบริหารจัดการรายชื่อนักศึกษา ปวส.")

# --- 3. ส่วนเมนูระบบกรอกข้อมูล (แยก Tab ปี 1 และ ปี 2) ---
st.subheader("➕ ลงทะเบียนนักศึกษาใหม่")

tab_p1, tab_p2 = st.tabs(["📝 กรอกข้อมูล ปี 1", "📝 กรอกข้อมูล ปี 2"])

def student_form(year_label, room_options):
    with st.form(f"form_{year_label}", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            batch = st.text_input(f"รุ่น ({year_label})", placeholder="เช่น 23")
            sid = st.text_input(f"รหัสนักศึกษา ({year_label})")
        with c2:
            name = st.text_input(f"ชื่อ-นามสกุล ({year_label})")
            st.info(f"ระดับชั้นที่กำลังบันทึก: {year_label}")
        with c3:
            # ใช้ลิสต์ห้องตามที่ส่งเข้ามาในฟังก์ชัน
            room = st.selectbox(f"เลือกห้องเรียน ({year_label})", room_options)
            st.caption("ตรวจสอบข้อมูลก่อนบันทึก")
        
        if st.form_submit_button(f"💾 บันทึกข้อมูล {year_label}", use_container_width=True):
            if sid and name:
                if sid in df['รหัสนักศึกษา'].values:
                    st.warning(f"⚠️ รหัส {sid} มีในระบบแล้ว! โปรดใช้ส่วนแก้ไขด้านล่าง")
                else:
                    new_data = pd.DataFrame([{
                        "รุ่น": f"'{batch}",
                        "รหัสนักศึกษา": f"'{sid}",
                        "ชื่อ-นามสกุล": name.strip(),
                        "ระดับชั้น": year_label,
                        "Room": room
                    }])
                    updated_df = pd.concat([df, new_data], ignore_index=True)
                    conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                    st.success(f"✅ บันทึกคุณ {name} ลงห้อง {room} เรียบร้อย!")
                    st.rerun()
            else:
                st.error("❌ กรุณากรอกรหัสและชื่อให้ครบถ้วน")

with tab_p1:
    student_form("ปี1", ROOMS_P1)

with tab_p2:
    student_form("ปี2", ROOMS_P2)

# --- 4. ส่วนค้นหาและแก้ไข (รองรับการย้ายห้องข้ามรูปแบบ) ---
st.divider()
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล / ย้ายห้อง")
search_q = st.text_input("🔎 พิมพ์รหัสหรือชื่อเพื่อค้นหา...")

if not df.empty:
    mask = df['รหัสนักศึกษา'].str.contains(search_q, case=False, na=False) | \
           df['ชื่อ-นามสกุล'].str.contains(search_q, case=False, na=False)
    filtered = df[mask].copy()

    if not filtered.empty:
        # ผสมรายชื่อห้องทั้งหมดเพื่อใช้ใน Editor กรณีมีการย้ายข้ามปี
        ALL_ROOMS = ROOMS_P1 + ROOMS_P2
        
        edited = st.data_editor(
            filtered,
            column_config={
                "รหัสนักศึกษา": st.column_config.TextColumn("รหัส (ห้ามแก้)", disabled=True),
                "ระดับชั้น": st.column_config.SelectboxColumn("ระดับชั้น", options=["ปี1", "ปี2"]),
                "Room": st.column_config.SelectboxColumn("ห้องเรียน", options=ALL_ROOMS)
            },
            key="edit_table_v10",
            use_container_width=True
        )

        if st.button("✅ ยืนยันการอัปเดตข้อมูล"):
            main_db = df.copy()
            for _, row in edited.iterrows():
                main_db.loc[main_db['รหัสนักศึกษา'] == row['รหัสนักศึกษา'], 
                           ['รุ่น', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']] = \
                    [row['รุ่น'], row['ชื่อ-นามสกุล'], row['ระดับชั้น'], row['Room']]
            
            for col in ['รุ่น', 'รหัสนักศึกษา']:
                main_db[col] = main_db[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")
            
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=main_db)
            st.success("✅ อัปเดตข้อมูลสำเร็จ!")
            st.rerun()

# --- 5. ส่วนส่งออก Excel (ฟอร์ม 16 คาบ เหมือนเดิม) ---
st.divider()
st.subheader("🖨️ ออกใบรายชื่อ (แยกไฟล์ตามชั้นปี)")

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
        
        ws.merge_cells('A1:U1')
        ws['A1'] = f"บัญชีรายชื่อนักศึกษา {target_year} ภาคเรียนที่ 1 ปีการศึกษา 2568"
        ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A2:U2'); ws['A2'] = f"ระดับ ปวส. ห้อง {r_name}"
        ws['A2'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A3:K3'); ws['A3'] = "วิชา..........................................................."
        ws.merge_cells('L3:U3'); ws['L3'] = "ผู้สอน........................................................."

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

col_p1, col_p2 = st.columns(2)
with col_p1:
    f1 = create_excel("ปี1")
    if f1: st.download_button("📥 ดาวน์โหลดใบรายชื่อ ปี 1", f1, "รายชื่อ_ปี1.xlsx", key="dl_p1")
with col_p2:
    f2 = create_excel("ปี2")
    if f2: st.download_button("📥 ดาวน์โหลดใบรายชื่อ ปี 2", f2, "รายชื่อ_ปี2.xlsx", key="dl_p2")
