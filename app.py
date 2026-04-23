import streamlit as st  # <--- ห้ามลบบรรทัดนี้เด็ดขาด!
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# 1. ต้องเซตค่าหน้ากระดาษเป็นอย่างแรกหลัง Import
st.set_page_config(page_title="ระบบจัดการนักศึกษา + Log การย้าย", layout="wide")

# 2. การเชื่อมต่อข้อมูล
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    # ทำความสะอาดข้อมูล (ป้องกัน .0 และ nan)
    for col in ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']:
        if col in data.columns:
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "")
            data[col] = data[col].replace('nan', '')
    return data

df = load_data()

# 3. ส่วนหัวของเว็บ
st.title("🏫 ระบบจัดการและบันทึกประวัติการย้ายห้อง")

# --- (ส่วนเพิ่มข้อมูลนักศึกษาใหม่ Form เดิมของคุณ) ---
# ... (ก๊อปปี้ส่วนเพิ่มข้อมูลมาวางตรงนี้) ...

# --- 4. ส่วนค้นหาและแก้ไข (ที่คุณต้องการบันทึกประวัติย้ายห้อง) ---
st.subheader("🔍 ค้นหาและจัดการการย้ายห้อง") # <--- จุดที่เคยติด Error จะหายไปถ้าเรียงตามนี้

search_term = st.text_input("พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search_term, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search_term, case=False, na=False)
    filtered_df = df[mask].copy()
    
    # เพิ่มช่องว่างสำหรับกรอกสาเหตุ
    filtered_df['สาเหตุการย้าย'] = ""
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=["ปี1", "ปี2"]),
            "Room": st.column_config.SelectboxColumn(options=[f"O1/{i}" for i in range(1, 16)]),
            "รหัสนักศึกษา": st.column_config.TextColumn(disabled=True),
            "สาเหตุการย้าย": st.column_config.TextColumn(placeholder="ระบุสาเหตุ...")
        },
        num_rows="dynamic",
        key="editor_log_v1"
    )
    
    if st.button("✅ ยืนยันการย้ายห้องและบันทึกประวัติ"):
        final_df = df.copy()
        log_entries = []
        
        for index, row in edited_df.iterrows():
            std_id = row['รหัสนักศึกษา']
            # หาห้องเดิมจากข้อมูลหลัก
            old_room_list = df.loc[df['รหัสนักศึกษา'] == std_id, 'Room'].values
            old_room = old_room_list[0] if len(old_room_list) > 0 else "ไม่ทราบ"
            new_room = row['Room']
            
            if old_room != new_room:
                # เก็บข้อมูลลง Log
                log_entries.append({
                    "วันที่-เวลา": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "รหัสนักศึกษา": f"'{std_id}",
                    "ชื่อ-นามสกุล": row['ชื่อ-นามสกุล'],
                    "ห้องเดิม": old_room,
                    "ห้องใหม่": new_room,
                    "สาเหตุการย้าย": row['สาเหตุการย้าย'] if row['สาเหตุการย้าย'] else "ไม่ระบุ"
                })
                # อัปเดตข้อมูลหน้าหลัก
                final_df.loc[final_df['รหัสนักศึกษา'] == std_id, ['ระดับชั้น', 'Room', 'ชื่อ-นามสกุล', 'รุ่น']] = \
                    [row['ระดับชั้น'], row['Room'], row['ชื่อ-นามสกุล'], row['รุ่น']]

        # บันทึกข้อมูลกลับ
        for col in ['รุ่น', 'รหัสนักศึกษา']:
            final_df[col] = final_df[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")
        
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=final_df)
        
        if log_entries:
            try:
                # บันทึก Log ลง Sheet ชื่อ "Log" (ต้องสร้างเตรียมไว้ใน Google Sheets)
                log_df = conn.read(spreadsheet=st.secrets["gsheet_url"], worksheet="Log", ttl=0)
                new_log_df = pd.concat([log_df, pd.DataFrame(log_entries)], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], worksheet="Log", data=new_log_df)
                st.success("ย้ายห้องและบันทึกประวัติเรียบร้อย!")
            except Exception as e:
                st.error("⚠️ กรุณาสร้าง Sheet ชื่อ 'Log' ใน Google Sheets ของคุณก่อน!")
        
        st.rerun()
