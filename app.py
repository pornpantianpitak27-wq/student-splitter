def load_data():
    # 1. ดึงข้อมูลจาก Sheets
    data = conn.read(spreadsheet=st.secrets["gsheet_url"], ttl=0)
    
    # 2. แก้ปัญหาจุดทศนิยม: แปลงทุกอย่างในคอลัมน์ที่ต้องการให้เป็น "ข้อความ" ก่อน
    # โดยใช้ฟังก์ชัน map เพื่อกำจัด .0 ออกไปให้หมด
    columns_to_fix = ['รุ่น', 'รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']
    
    for col in columns_to_fix:
        if col in data.columns:
            # แปลงเป็น String และถ้ามี .0 ให้ตัดออก
            data[col] = data[col].astype(str).replace(r'\.0$', '', regex=True)
            # ถ้าเป็นค่าว่างของระบบ (nan) ให้เปลี่ยนเป็นช่องว่างจริง
            data[col] = data[col].replace('nan', '')
            
    return data
