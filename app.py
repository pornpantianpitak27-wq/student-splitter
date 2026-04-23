import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. การตั้งค่าเบื้องต้น ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - Final Version", layout="wide")

# --- 2. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        # อ่านข้อมูลจากหน้าแรก (Main Sheet)
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # Clean ข้อมูล: ลบ .0 และลบเครื่องหมาย ' เพื่อให้นำมาประมวลผลได้ง่าย
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

st.title("🏫 ระบบจัดการรายชื่อและประวัติการย้ายห้อง")

# --- 3. ส่วนเพิ่มนักศึกษาใหม่ ---
with st.expander("➕ เพิ่มรายชื่อนักศึกษาใหม่", expanded=False):
    with st.form("add_student_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_batch = st.text_input("รุ่น (Batch)")
            new_id = st.text_input("รหัสนักศึกษา")
        with c2:
            new_name = st.text_input("ชื่อ-นามสกุล")
            new_level = st.selectbox("ระดับชั้น", CLASSES)
        with c3:
            new_room = st.selectbox("ห้องเรียน (Room)", ROOMS)
        
        if st.form_submit_button("💾 บันทึกข้อมูลใหม่"):
            if new_id and new_name:
                # ตรวจสอบว่ารหัสซ้ำหรือไม่
                if new_id in df['รหัสนักศึกษา'].values:
                    st.error("❌ รหัสนักศึกษานี้มีอยู่ในระบบแล้ว กรุณาใช้ส่วนแก้ไขข้อมูลด้านล่าง")
                else:
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
                st.warning("⚠️ กรุณากรอกข้อมูลให้ครบถ้วน")

# --- 4. ส่วนค้นหาและแก้ไข (Update เท่านั้น ไม่เพิ่มบรรทัดใหม่) ---
st.divider()
st.subheader("🔍 ค้นหาและจัดการการย้ายห้อง (แก้ไขได้ตลอด)")
search_q = st.text_input("🔎 พิมพ์รหัสหรือชื่อเพื่อค้นหา...")

if not df.empty:
    mask = df['รหัสนักศึกษา'].str.contains(search_q, case=False, na=False) | \
           df['ชื่อ-นามสกุล'].str.contains(search_q, case=False, na=False)
    filtered_df = df[mask].copy()

    if not filtered_df.empty:
        # ตารางแก้ไขข้อมูล
        edited_df = st.data_editor(
            filtered_df,
            column_config={
                "รหัสนักศึกษา": st.column_config.TextColumn("รหัส (ห้ามแก้)", disabled=True),
                "Room": st.column_config.SelectboxColumn("ห้องใหม่", options=ROOMS),
                "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES)
            },
            key="main_editor_v7",
            use_container_width=True
        )

        if st.button("✅ ยืนยันการแก้ไขข้อมูล"):
            main_db = df.copy()
            # ใช้ระบบ Update โดยอ้างอิงรหัสนักศึกษา
            for _, row in edited_df.iterrows():
                std_id = row['รหัสนักศึกษา']
                # เขียนทับข้อมูลในแถวเดิมที่มีรหัสนี้
                main_db.loc[main_db['รหัสนักศึกษา'] == std_id, 
                           ['รุ่น', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']] = \
                    [row['รุ่น'], row['ชื่อ-นามสกุล'], row['ระดับชั้น'], row['Room']]

            # ใส่ ' นำหน้าก่อนบันทึกเสมอเพื่อกันเลข 0 หาย
            for col in ['รุ่น', 'รหัสนักศึกษา']:
                main_db[col] = main_db[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")

            conn.update(spreadsheet=st.secrets["gsheet_url"], data=main_db)
            st.success("✅ อัปเดตข้อมูลและย้ายห้องเรียบร้อย (ไม่มีชื่อซ้ำ)")
            st.rerun()

# --- 5. ส่วนส่งออก Excel (ตามฟอร์มต้นฉบับ 16 คาบ) ---
st.divider()
if st.button("🖨️ ออกใบรายชื่อแยกตามห้อง (Excel)"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        
        # ตั้งค่าเส้นขอบ
        side = Side(style='thin')
        border = Border(left=side, right=side, top=side, bottom=side)
        
        for r_name in sorted(df['Room'].unique()):
            if not r_name or r_name == 'nan': continue
            
            ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
            room_data = df[df['Room'] == r_name].sort_values('รหัสนักศึกษา')
            
            # --- ส่วนหัวตามต้นฉบับ ---
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = Font(bold=True, size=14)
            ws['A1'].alignment = Alignment(horizontal='center')

            ws.merge_cells('A2:U2')
            current_lv = room_data['ระดับชั้น'].iloc[0] if not room_data.empty else ""
            ws['A2'] = f"ระดับ ปวส. ชั้น{current_lv} ห้อง {r_name}"
            ws['A2'].alignment = Alignment(horizontal='center')

            ws.merge_cells('A3:K3')
            ws['A3'] = "วิชา..........................................................."
            ws.merge_cells('L3:U3')
            ws['L3'] = "ผู้สอน........................................................."

            # --- หัวตาราง ---
            headers = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ']
            ws.append(headers)
            for cell in ws[4]:
                cell.border = border
                cell.alignment = Alignment(horizontal='center')
                cell.font = Font(bold=True)

            # --- ใส่รายชื่อ ---
            for i, row in enumerate(room_data.itertuples(), 1):
                ws.append([i, row.รหัสนักศึกษา, row._3] + ['' for _ in range(16)] + [f"รุ่น {row.รุ่น}"])
                # ใส่เส้นขอบทุกช่องในบรรทัดนั้น
                for cell in ws[ws.max_row]:
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
                # เฉพาะคอลัมน์ชื่อ ให้ชิดซ้าย
                ws.cell(row=ws.max_row, column=3).alignment = Alignment(horizontal='left', indent=1)

            # ตั้งค่าความกว้างคอลัมน์
            ws.column_dimensions['A'].width = 5
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 25
            for c in range(4, 20):
                ws.column_dimensions[ws.cell(row=4, column=c).column_letter].width = 3.5

        wb.save(output)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (ทุกห้อง)",
            data=output.getvalue(),
            file_name=f"ใบรายชื่อ_2568_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
