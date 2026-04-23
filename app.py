import streamlit as st
import pandas as pd
from io import BytesIO

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบแยกรายชื่อนักศึกษา", page_icon="🏫")

st.title("🏫 ระบบแยกรายชื่อนักศึกษา")
st.markdown("### อัปโหลดไฟล์เดียว แยกให้ครบทุกห้องในไฟล์เดียว (แยก Sheet)")

# ส่วนคำแนะนำการใช้งาน
with st.expander("📌 วิธีการเตรียมไฟล์ (คลิกเพื่ออ่าน)"):
    st.write("""
    1. ไฟล์ต้องมีหัวตาราง (คอลัมน์) ที่ชื่อว่า **Room** (ตัวพิมพ์ใหญ่ R)
    2. ข้อมูลห้องเรียนสามารถเป็น 6/1, 6/2 หรือ 1, 2, 3 ก็ได้
    3. รองรับไฟล์นามสกุล .csv และ .xlsx (Excel)
    """)

# 1. ส่วนการอัปโหลดไฟล์
uploaded_file = st.file_uploader("เลือกไฟล์รายชื่อนักศึกษา", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        # อ่านข้อมูลจากไฟล์ที่อัปโหลด
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.success("อ่านข้อมูลสำเร็จ!")
        
        # ตรวจสอบว่ามีคอลัมน์ Room หรือไม่
        if 'Room' in df.columns:
            # ดึงรายชื่อห้องทั้งหมดและเรียงลำดับ
            rooms = sorted(df['Room'].unique())
            
            # เตรียมพื้นที่ในหน่วยความจำเพื่อสร้างไฟล์ Excel
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for room in rooms:
                    # กรองข้อมูลรายห้อง
                    room_df = df[df['Room'] == room].copy()
                    
                    # เพิ่มคอลัมน์ลำดับที่ (นับ 1 ใหม่ทุกห้อง)
                    room_df.insert(0, 'ลำดับ', range(1, len(room_df) + 1))
                    
                    # ตั้งชื่อ Sheet (จัดการตัวอักษรพิเศษที่ Excel ไม่รองรับ)
                    clean_sheet_name = f"ห้อง {str(room).replace('/', '-')}"
                    
                    # เขียนลง Sheet
                    room_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            
            st.info(f"ระบบตรวจพบทั้งหมด {len(rooms)} ห้องเรียน")
            
            # 2. ปุ่มดาวน์โหลด
            st.download_button(
                label="📥 คลิกเพื่อดาวน์โหลดไฟล์แยกห้อง (Excel)",
                data=output.getvalue(),
                file_name="รายชื่อนักศึกษา_แยกห้องเรียบร้อย.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # แสดงตัวอย่างข้อมูลให้ดูด้านล่าง
            st.divider()
            st.write("🔍 ตัวอย่างข้อมูลที่อ่านได้:")
            st.dataframe(df.head(10))
            
        else:
            st.error("❌ ไม่พบคอลัมน์ชื่อ 'Room' ในไฟล์ของคุณ กรุณาตรวจสอบหัวตาราง")
            
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

else:
    st.write("---")
    st.caption("รอการอัปโหลดไฟล์จากคุณ...")