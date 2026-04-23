# --- 3. ส่วนค้นหาและแก้ไข (ปรับปรุงเพื่อป้องกันชื่อซ้ำ) ---
st.subheader("🔍 ค้นหาและแก้ไขข้อมูล")
search_term = st.text_input("พิมพ์ชื่อหรือรหัสเพื่อค้นหา...")

if not df.empty:
    mask = df['ชื่อ-นามสกุล'].str.contains(search_term, case=False, na=False) | \
           df['รหัสนักศึกษา'].str.contains(search_term, case=False, na=False)
    filtered_df = df[mask].copy()
    
    edited_df = st.data_editor(
        filtered_df,
        column_config={
            "ระดับชั้น": st.column_config.SelectboxColumn(options=CLASSES),
            "Room": st.column_config.SelectboxColumn(options=ROOMS),
            "รหัสนักศึกษา": st.column_config.TextColumn(disabled=True) # ห้ามแก้รหัสในตารางนี้เพื่อป้องกันการจับคู่ผิด
        },
        num_rows="dynamic",
        key="editor_v4"
    )
    
    if st.button("✅ บันทึกการเปลี่ยนแปลง"):
        # สร้างสำเนาข้อมูลหลัก
        final_df = df.copy()
        
        # วนลูปแก้ไขข้อมูลทีละแถวโดยใช้ 'รหัสนักศึกษา' เป็นตัวหาตำแหน่ง
        for index, row in edited_df.iterrows():
            std_id = row['รหัสนักศึกษา']
            # หาตำแหน่งแถวในข้อมูลหลักที่รหัสตรงกัน
            final_df.loc[final_df['รหัสนักศึกษา'] == std_id, ['ระดับชั้น', 'Room', 'ชื่อ-นามสกุล', 'รุ่น']] = \
                [row['ระดับชั้น'], row['Room'], row['ชื่อ-นามสกุล'], row['รุ่น']]

        # เติมเครื่องหมาย ' นำหน้าเพื่อให้ Google Sheets คงสภาพความเป็นข้อความ (ป้องกันเลข 0 หาย)
        for col in ['รุ่น', 'รหัสนักศึกษา']:
            final_df[col] = final_df[col].apply(lambda x: f"'{str(x).replace(\"'\", \"\")}")

        # อัปเดตลง Google Sheets ทั้งหมด (ทับไฟล์เดิม)
        conn.update(spreadsheet=st.secrets["gsheet_url"], data=final_df)
        st.success("อัปเดตการย้ายห้องเรียบร้อยแล้ว! (ข้อมูลเก่าถูกแทนที่แล้ว)")
        st.rerun()
        wb.save(output)
        st.download_button("💾 ดาวน์โหลดไฟล์ .xlsx", output.getvalue(), "Attendance_List.xlsx")
