import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image # เพิ่มการใช้งาน Image เพื่อแทรกโลโก้
from streamlit_gsheets import GSheetsConnection

# --- 1. ตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="ระบบจัดการนักศึกษา - ตามไฟล์จริง + โลโก้", layout="wide")

# --- 2. การเชื่อมต่อฐานข้อมูล ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
        # ล้างข้อมูลให้สะอาด
        for col in data.columns:
            data[col] = data[col].astype(str).str.replace(r'\.0$', '', regex=True)
            data[col] = data[col].str.replace("'", "").replace('nan', '')
        return data
    except Exception:
        return pd.DataFrame()

df = load_data()
ROOMS_P1 = [f"O1/{i}" for i in range(1, 16)]
ROOMS_P2 = [f"O2/{i}" for i in range(1, 16)]

st.title("🏫 ระบบจัดการรายชื่อนักศึกษา ปวส. (Stable Version)")

# --- 3. ส่วนกรอกข้อมูล ---
st.subheader("➕ เพิ่มรายชื่อนักศึกษาใหม่")
tab_p1, tab_p2 = st.tabs(["📝 ปี 1 (O1)", "📝 ปี 2 (O2)"])

def student_form(year_label, room_options):
    with st.form(f"form_{year_label}", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
        with c1: batch = st.text_input("รุ่น", placeholder="23", key=f"b_{year_label}")
        with c2: sid = st.text_input("รหัสนักศึกษา", key=f"s_{year_label}")
        with c3: fname = st.text_input("ชื่อ", key=f"f_{year_label}")
        with c4: lname = st.text_input("นามสกุล", key=f"l_{year_label}")
        
        c5, c6 = st.columns(2)
        with c5: room = st.selectbox("เลือกห้องเรียน", room_options, key=f"r_{year_label}")
        with c6: st.info(f"ระดับชั้นที่เลือก: {year_label}")

        if st.form_submit_button("💾 บันทึกข้อมูล", use_container_width=True):
            if sid and fname:
                if sid in df['รหัสนักศึกษา'].values:
                    st.warning("⚠️ รหัสนักศึกษานี้มีอยู่ในระบบแล้ว")
                else:
                    new_entry = pd.DataFrame([{
                        "รุ่น": f"'{batch}",
                        "รหัสนักศึกษา": f"'{sid}",
                        "ชื่อ": fname.strip(),
                        "นามสกุล": lname.strip(),
                        "ระดับชั้น": year_label,
                        "Room": room
                    }])
                    current_df = load_data()
                    updated_df = pd.concat([current_df, new_entry], ignore_index=True)
                    conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
                    st.success("✅ บันทึกข้อมูลสำเร็จ!"); st.rerun()

with tab_p1: student_form("ปี1", ROOMS_P1)
with tab_p2: student_form("ปี2", ROOMS_P2)

# --- 4. ฟังก์ชันสร้าง Excel (แทรกโลโก้และจัดรูปแบบตามต้นฉบับเป๊ะ) ---
def create_excel_report(target_year):
    data_to_use = load_data()
    if data_to_use.empty: return None
    
    # กรองข้อมูลตามชั้นปี
    year_data = data_to_use[data_to_use['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # ลบ Sheet เริ่มต้น
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    bold_font = Font(bold=True, size=12)
    center_align = Alignment(horizontal='center', vertical='center')
    
    # โหลดไฟล์โลโก้
    # โค้ดส่วนนี้จะพยายามโหลดไฟล์โลโก้ที่คุณอัปโหลดมา หากไฟล์ไม่มี ให้ใส่โลโก้ในโฟลเดอร์เดียวกับโค้ด
    # หรือใส่ URL โดยตรงได้ที่นี่
    try:
        logo_img = Image('image_4.png') 
        # ปรับขนาดโลโก้ (กว้าง/สูง) ให้พอดีกับตำแหน่ง
        # ตัวอย่างเช่นปรับให้สูง 1.5 นิ้ว (144 pixels) หรือตามความเหมาะสมในรูป
        logo_img.height = 120 
        logo_img.width = 120 
    except Exception:
        logo_img = None
        st.warning("⚠️ ไม่พบไฟล์โลโก้ 'image_4.png' ในระบบ โปรแกรมจะสร้างไฟล์โดยไม่มีโลโก้")

    # สร้าง Sheet แยกตามห้อง
    rooms_in_year = sorted(year_data['Room'].unique())
    for r_name in rooms_in_year:
        if not r_name or r_name == 'nan': continue
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # --- หัวไฟล์แถวที่ 1-2 (ตารางขวาบน) ตามรูป Excel เป๊ะ ---
        ws.merge_cells('N1:U1'); ws['N1'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['N1'].border = border; ws['N1'].alignment = center_align; ws['N1'].font = bold_font
        
        headers_sub = ['เช็คชื่อนักศึกษา', 'เซ็นสอบกลางภาค', 'เซ็นสอบปลายภาค']
        cols_sub = [('N2', 'O2'), ('P2', 'R2'), ('S2', 'U2')]
        for i, (start, end) in enumerate(cols_sub):
            ws.merge_cells(f'{start}:{end}')
            ws[start] = headers_sub[i]
            ws[start].border = border; ws[start].alignment = center_align

        # --- ส่วนชื่อหัวข้อหลัก ---
        ws.merge_cells('A3:U3')
        ws['A3'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568" # ปรับปีการศึกษา
        ws['A3'].font = Font(bold=True, size=14); ws['A3'].alignment = center_align

        ws.merge_cells('A4:U4')
        ws['A4'] = f"ระดับ ปวส. ชั้นปีที่ {'1' if target_year=='ปี1' else '2'} ห้อง {r_name} (เรียนวันพฤหัสบดี) ศูนย์บางแค"
        ws['A4'].alignment = center_align

        ws.merge_cells('A5:K5'); ws['A5'] = "วิชา..........................................................................."
        ws.merge_cells('L5:U5'); ws['L5'] = "ผู้สอน..........................................................................."

        # --- หัวตารางตามรูปภาพที่แนบมาเป๊ะๆ (แถวที่ 6-9) ---
        ws.merge_cells('A6:A9'); ws['A6'] = "เลขที่"
        ws.merge_cells('B6:B9'); ws['B6'] = "รหัสประจำตัว"
        ws.merge_cells('C6:K9'); ws['C6'] = "ชื่อ-สกุล"
        
        ws.merge_cells('L6:L6'); ws['L6'] = "เดือน"
        ws.merge_cells('L7:L7'); ws['L7'] = "วันที่"
        ws.merge_cells('L8:L8'); ws['L8'] = "คาบ"
        
        ws.merge_cells('U6:U9'); ws['U6'] = "หมายเหตุ"

        for i in range(1, 9): # ตัวเลขคาบ 1-8
            ws.cell(row=8, column=12+i).value = i
            
        # ตีเส้นหัวตารางทั้งหมด
        for r in range(6, 10):
            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # --- รายชื่อนักศึกษา (เริ่มแถว 10) ---
        for i, row in enumerate(room_data.itertuples(), 1):
            sid_raw = str(getattr(row, 'รหัสนักศึกษา', '')).replace("'", "")
            sid_val = f"'{sid_raw}"
            fname = getattr(row, 'ชื่อ', '')
            lname = getattr(row, 'นามสกุล', '')
            fullname = f"{fname} {lname}".strip()
            
            ws.append([i, sid_val, fullname, ''] + ['' for _ in range(16)] + [f"รุ่น {getattr(row, 'รุ่น', '')}"])
            current_row = ws.max_row
            for c in range(1, 22):
                ws.cell(row=current_row, column=c).border = border; ws.cell(row=current_row, column=c).alignment = center_align
            ws.cell(row=current_row, column=3).alignment = Alignment(horizontal='left', indent=1)

        # --- ท้ายไฟล์ ---
        f_row = ws.max_row + 2
        ws.merge_cells(f'A{f_row}:K{f_row}'); ws[f'A{f_row}'] = "ชื่ออาจารย์ที่ปรึกษา............................................................"
        ws.merge_cells(f'L{f_row}:U{f_row}'); ws[f'L{f_row}'] = "ชื่อหัวหน้าชั้น............................................................"

        # --- ปรับความกว้างคอลัมน์ ---
        ws.column_dimensions['A'].width = 6; ws.column_dimensions['B'].width = 16; ws.column_dimensions['C'].width = 30; ws.column_dimensions['D'].width = 8
        for c in range(5, 21): ws.column_dimensions[ws.cell(row=7, column=c).column_letter].width = 3.5

        # *** แทรกโลโก้ ***
        if logo_img:
            # แทรกที่ตำแหน่งเซลล์ตามรูป Excel
            ws.add_image(logo_img, 'H1') # หรือ H2 ตามตำแหน่งในรูปให้สวยงาม

    wb.save(output)
    return output.getvalue()

# --- 5. ส่วนแสดงปุ่มดาวน์โหลด (ปรับ UI ให้น่าใช้) ---
st.divider()
st.subheader("🖨️ ออกใบรายชื่อ (พร้อมโลโก้)")
col_p1, col_p2 = st.columns(2)

with col_p1:
    st.write("📂 ข้อมูล ปี 1")
    f1 = create_excel_report("ปี1")
    if f1:
        st.download_button("📥 ดาวน์โหลดไฟล์ ปี 1 (xlsx)", f1, "ใบรายชื่อ_ปี1.xlsx", use_container_width=True)
    else:
        st.info("💡 ยังไม่มีข้อมูลนักศึกษา ปี 1")

with col_p2:
    st.write("📂 ข้อมูล ปี 2")
    f2 = create_excel_report("ปี2")
    if f2:
        st.download_button("📥 ดาวน์โหลดไฟล์ ปี 2 (xlsx)", f2, "ใบรายชื่อ_ปี2.xlsx", use_container_width=True)
    else:
        st.warning("💡 ยังไม่มีข้อมูลนักศึกษา ปี 2 (ปุ่มดาวน์โหลดจะปรากฏเมื่อมีข้อมูล)")

# --- 6. ส่วนค้นหา/แก้ไข (คงเดิมเพื่อความต่อเนื่อง) ---
with st.expander("🔍 ค้นหาและจัดการการย้ายห้อง"):
    search = st.text_input("ค้นหาชื่อหรือรหัส...")
    if not df.empty:
        found = df[df['ชื่อ'].str.contains(search, na=False) | df['รหัสนักศึกษา'].str.contains(search, na=False)]
        st.data_editor(found, use_container_width=True, key="main_editor_v12")
