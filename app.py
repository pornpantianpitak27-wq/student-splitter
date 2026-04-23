import streamlit as st
from streamlit_gsheets import GSheetsConnection

# สร้างการเชื่อมต่อกับ Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# ดึงข้อมูลจาก Google Sheets มาแสดง
df = conn.read(spreadsheet=st.secrets["gsheet_url"])

# เวลาจะบันทึกข้อมูลใหม่ ให้ใช้คำสั่งนี้แทน:
if st.button("บันทึกเข้าสู่ระบบ"):
    # (โค้ดเตรียมข้อมูลใหม่...)
    updated_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_df)
    st.success("บันทึกข้อมูลลง Google Sheets สำเร็จ!")
