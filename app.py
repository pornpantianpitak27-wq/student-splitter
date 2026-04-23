import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา & ออกแบบฟอร์ม", layout="wide")

# --- 2. การเชื่อมต่อ Google Sheets ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    cols_to_fix = ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']
    for col in cols_to_fix:
        if col in data.columns:
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "")
            data[col] = data[col].replace('nan', '')
    return data

df = load_data()
CLASSES = ["ปี1", "ปี2"]
ROOMS = [f"O1/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการข้อมูลนักศึกษาและออกใบรายชื่อ")

# --- (ส่วนเพิ่มข้อมูล และ แก้ไขย้ายห้อง คงไว้ตามเดิม) ---
# ... (ก๊อปปี้ส่วน Form และ Editor จากเวอร์ชันก่อนมาวางตรงนี้ได้เลย) ...

# --- 5. ส่วนส่งออก Excel (ถอดแบบฟอร์มจากต้นฉบับ) ---
st.divider()
st.subheader("🖨️ ออกใบรายชื่อ (รูปแบบเดียวกับไฟล์ต้นฉบับ)")

if st.button("📥 สร้างไฟล์ Excel (แยกห้อง)"):
    if not df.empty:
        output = BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        
        # ตั้งค่าเส้นขอบ
        side = Side(style='thin')
        border = Border(left=side, right=side, top=side, bottom=side)
        
        # ฟอนต์มาตรฐาน
        font_header = Font(name='Sarabun', size=14, bold=True)
        font_sub = Font(name='Sarabun', size=12)
        
        for r_name in sorted(df['Room'].unique()):
            if not r_name: continue
            
            # สร้าง Sheet และตั้งชื่อ
            sheet_title = f"ห้อง {r_name.replace('/', '-')}"
            ws = wb.create_sheet(title=sheet_title)
            room_df = df[df['Room'] == r_name].sort_values('รหัสนักศึกษา')
            
            # --- เริ่มสร้างส่วนหัว (Header) ตามต้นฉบับ ---
            # บรรทัด 1: ชื่อรายงาน
            ws.merge_cells('A1:U1')
            ws['A1'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"
            ws['A1'].font = font_header
            ws['A1'].alignment = Alignment(horizontal='center')
            
            # บรรทัด 2: ระดับชั้น/ห้อง
            ws.merge_cells('A2:U2')
            current_level = room_df['ระดับชั้น'].iloc[0] if not room_df.empty else "ปวส."
            ws['A2'] = f"ระดับ ปวส. ชั้น{current_level}  ห้อง  {r_name}"
            ws['A2'].font = font_sub
            ws['A2'].alignment = Alignment(horizontal='center')
            
            # บรรทัด 3: วิชา / ผู้สอน (แบบช่องว่างให้เขียน)
            ws.merge_cells('A3:K3')
            ws['A3'] = "วิชา........................................................................................"
            ws.merge_cells('L3:U3')
            ws['L3'] = "ผู้สอน......................................................................................"
            ws['A3'].font = font_sub
            ws['L3'].font = font_sub

            # บรรทัด 4: หัวตาราง
            # ลำดับคอลัมน์: เลขที่(A), รหัส(B), ชื่อ(C), คาบ 1-16 (D-S), หมายเหตุ(T)
            headers = ['เลขที่', 'รหัสประจำตัว', 'ชื่อ-นามสกุล'] + [f'{i+1}' for i in range(16)] + ['หมายเหตุ']
            ws.append(headers)
            
            # จัดรูปแบบหัวตาราง
            header_row = ws[4]
            for cell in header_row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.border = border

            # --- ใส่รายชื่อนักศึกษา ---
            for i, row in enumerate(room_df.itertuples(), 1):
                # ตระเตรียมข้อมูลแต่ละแถว
                std_name = row._3 # คอลัมน์ชื่อ-นามสกุล
                std_id = row.รหัสนักศึกษา
                batch = f"รุ่น {row.รุ่น}"
                
                # เขียนข้อมูลลงบรรทัด
                ws.append([i, std_id, std_name] + ['' for _ in range(16)] + [batch])
                
                # ใส่เส้นขอบและจัดกึ่งกลาง
                current_row = ws[ws.max_row]
                for cell in current_row:
                    cell.border = border
                    # เฉพาะชื่อ-สกุล ให้ชิดซ้าย
                    if cell.column == 3:
                        cell.alignment = Alignment(horizontal='left', indent=1)
                    else:
                        cell.alignment = Alignment(horizontal='center')

            # --- ตั้งค่าความกว้างคอลัมน์ ---
            ws.column_dimensions['A'].width = 6   # เลขที่
            ws.column_dimensions['B'].width = 15  # รหัส
            ws.column_dimensions['C'].width = 30  # ชื่อ-สกุล
            ws.column_dimensions['T'].width = 15  # หมายเหตุ
            # คอลัมน์คาบเรียนให้แคบลง
            for col in range(4, 20):
                ws.column_dimensions[ws.cell(row=4, column=col).column_letter].width = 4

        wb.save(output)
        st.download_button(
            label="💾 ดาวน์โหลดใบรายชื่อ (Excel)",
            data=output.getvalue(),
            file_name=f"ใบรายชื่อ_ปี68_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("ระบบจะจัดกลุ่มนักศึกษาตามห้อง (Room) และสร้างตารางเช็คชื่อ 16 คาบให้โดยอัตโนมัติ")
