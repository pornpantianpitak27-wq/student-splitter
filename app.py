import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าหน้ากระดาษ (ต้องอยู่บรรทัดแรกๆ) ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา & Log การย้าย", layout="wide")

# --- 2. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    # ดึงข้อมูลจากแผ่นงานหลัก (หน้าแรก)
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    
    # ล้างข้อมูลขยะและรักษาเลข 0 นำหน้า
    cols_to_fix = ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']
    for col in cols_to_fix:
        if col in data.columns:
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "")
            data[col] = data[col].replace('nan', '')
    return data

# โหลดข้อมูล
df = load_data()

# ตัวเลือกพื้นฐาน
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการข้อมูลและประวัติการย้ายห้องนักศึกษา")

# --- 3. ส่วนเพิ่มนักศึกษาใหม่ ---
with st.expander("➕ เพิ่มนักศึกษาใหม่", expanded=False):
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
        
        if st.form_submit_button("💾 บันทึกรายชื่อใหม่"):
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
                st.error("⚠️ กรุณากรอกข้อมูลให้ครบถ้วน")

# --- 4. ส่วนค้นหาและย้ายห้อง (พร้อมบันทึก Log) ---
st.divider()
st.subheader("🔍 ค้นหาและจัดการการย้ายห้อง")
search = st.text_input("🔎 พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search, case=False, na=False)
    filtered_df = df[mask].copy()
    
    # เพิ่มคอลัมน์ "สาเหตุการย้าย" ให้กรอกในตาราง
    filtered_df['สาเหตุการย้าย'] = ""
    
    st.info("💡 เปลี่ยนชื่อห้องในช่อง 'Room' และระบุเหตุผลในช่อง 'สาเหตุการย้าย' แล้วกดปุ่มยืนยันด้านล่าง")
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "รหัสนักศึกษา": st.column_config.TextColumn(disabled=True),
            "สาเหตุการย้าย": st.column_config.TextColumn("สาเหตุการย้าย") # ตัด placeholder ออกเพื่อกัน Error
        },
        num_rows="dynamic",
        key="main_editor_v6"
    )
    
    if st.button("✅ ยืนยันการแก้ไขและบันทึกประวัติ"):
        final_df = df.copy()
        log_entries = []
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        for index, row in edited_df.iterrows():
            std_id = row['รหัสนักศึกษา']
            # หาข้อมูลเดิม
            original_row = df[df['รหัสนักศึกษา'] == std_id]
            if not original_row.empty:
                old_room = original_row['Room'].values[0]
                new_room = row['Room']
                
                # ตรวจสอบว่ามีการเปลี่ยนห้องหรือไม่
                if old_room != new_room:
                    log_entries.append({
                        "วันที่-เวลา": now_str,
                        "รหัสนักศึกษา": f"'{std_id}",
                        "ชื่อ-นามสกุล": row['ชื่อ-นามสกุล'],
                        "ห้องเดิม": old_room,
                        "ห้องใหม่": new_room,
                        "สาเหตุการย้าย": row['สาเหตุการย้าย'] if row['สาเหตุการย้าย'] else "ไม่ระบุสาเหตุ"
                    })
                
                # อัปเดตข้อมูลใน DataFrame หลัก (ใช้ .loc เพื่อทับบรรทัดเดิม ป้องกันชื่อซ้ำ)
                final_df.loc[final_df['รหัสนักศึกษา'] == std_id, ['รุ่น', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']] = \
                    [row['รุ่น'], row['ชื่อ-นามสกุล'], row['ระดับชั้น'], row['Room']]

        # เติมเครื่องหมาย ' เพื่อรักษาเลข 0 ก่อนบันทึก
        for col in ['รุ่น', 'รหัสนักศึกษา']:
            final_df[col] = final_df[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")
            
        # 1. บันทึกหน้าหลัก
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=final_df)
        
        # 2. บันทึกหน้า Log (ถ้ามีการย้ายห้อง)
        if log_entries:
            try:
                log_df = conn.read(spreadsheet=st.secrets["gsheet_url"], worksheet="Log", ttl=0)
                new_log_data = pd.concat([log_df, pd.DataFrame(log_entries)], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], worksheet="Log", data=new_log_data)
                st.success("ย้ายห้องและบันทึกประวัติสำเร็จ!")
            except:
                st.error("⚠️ ไม่พบแผ่นงานชื่อ 'Log' ใน Google Sheets กรุณาสร้างขึ้นใหม่ก่อนใช้งานส่วนนี้")
        else:
            st.success("อัปเดตข้อมูลเรียบร้อย (ไม่มีการย้ายห้อง)")
            
        st.rerun()

# --- 5. ส่วนส่งออก Excel ---
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
            ws = wb.create_sheet(title=f"Room {r_name.replace('/', '-')}")
            room_df = df[df['Room'] == r_name].sort_values('รหัสนักศึกษา')
            
            ws.merge_cells('A1:U1')
            ws['A1'] = f"ใบรายชื่อนักศึกษา ห้อง {r_name} ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True, size=12)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            ws.append(['เลขที่', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ'])
            for i, r in enumerate(room_df.itertuples(), 1):
                ws.append([i, r.รหัสนักศึกษา, r._3] + ['' for _ in range(16)] + [f"รุ่น {r.รุ่น}"])
            
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=20):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions['C'].width = 25

        wb.save(output)
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", output.getvalue(), "Student_List_2025.xlsx")
