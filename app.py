# --- 3. ส่วนค้นหาและแก้ไข (ย้ายห้อง + บันทึกประวัติ) ---
st.subheader("🔍 ค้นหาและจัดการการย้ายห้อง")
search_term = st.text_input("พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search_term, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search_term, case=False, na=False)
    filtered_df = df[mask].copy()
    
    # เพิ่มคอลัมน์ชั่วคราวสำหรับกรอกสาเหตุในหน้าเว็บ (ไม่บันทึกลงหน้าหลัก)
    filtered_df['สาเหตุการย้าย'] = ""
    
    st.info("💡 วิธีการย้าย: เปลี่ยนชื่อห้องในคอลัมน์ 'Room' และพิมพ์สาเหตุในช่อง 'สาเหตุการย้าย' จากนั้นกดบันทึก")
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "รหัสนักศึกษา": st.column_config.TextColumn(disabled=True),
            "สาเหตุการย้าย": st.column_config.TextColumn(placeholder="ระบุสาเหตุที่ย้าย...")
        },
        num_rows="dynamic",
        key="editor_v5"
    )
    
    if st.button("✅ ยืนยันการย้ายห้องและบันทึกประวัติ"):
        from datetime import datetime
        
        final_df = df.copy()
        log_entries = [] # สำหรับเก็บประวัติการย้าย
        
        for index, row in edited_df.iterrows():
            std_id = row['รหัสนักศึกษา']
            old_room = df.loc[df['รหัสนักศึกษา'] == std_id, 'Room'].values[0]
            new_room = row['Room']
            
            # ตรวจสอบว่ามีการเปลี่ยนห้องจริงๆ หรือไม่
            if old_room != new_room:
                # 1. เตรียมข้อมูล Log
                log_entries.append({
                    "วันที่-เวลา": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "รหัสนักศึกษา": f"'{std_id}",
                    "ชื่อ-นามสกุล": row['ชื่อ-นามสกุล'],
                    "ห้องเดิม": old_room,
                    "ห้องใหม่": new_room,
                    "สาเหตุการย้าย": row['สาเหตุการย้าย'] if row['สาเหตุการย้าย'] else "ไม่ระบุ"
                })
                
                # 2. อัปเดตข้อมูลในหน้าหลัก
                final_df.loc[final_df['รหัสนักศึกษา'] == std_id, ['ระดับชั้น', 'Room', 'ชื่อ-นามสกุล', 'รุ่น']] = \
                    [row['ระดับชั้น'], row['Room'], row['ชื่อ-นามสกุล'], row['รุ่น']]

        # บันทึกข้อมูลหน้าหลัก
        for col in ['รุ่น', 'รหัสนักศึกษา']:
            final_df[col] = final_df[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=final_df)
        
        # บันทึกข้อมูลหน้า Log (ถ้ามีการย้ายจริง)
        if log_entries:
            try:
                # ดึงข้อมูล Log เดิมมาเพื่อต่อท้าย
                log_df = conn.read(spreadsheet=st.secrets["gsheet_url"], worksheet="Log", ttl=0)
                new_log_df = pd.concat([log_df, pd.DataFrame(log_entries)], ignore_index=True)
                conn.update(spreadsheet=st.secrets["gsheet_url"], worksheet="Log", data=new_log_df)
                st.success("ย้ายห้องและบันทึกประวัติลงหน้า 'Log' เรียบร้อยแล้ว!")
            except:
                st.error("⚠️ ไม่พบ Sheet ที่ชื่อ 'Log' กรุณาสร้าง Sheet ใหม่ใน Google Sheets และตั้งชื่อว่า Log")
        else:
            st.warning("ไม่มีการเปลี่ยนแปลงห้องเรียน")
            
        st.rerun()
