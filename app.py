# --- 4. ส่วนค้นหาและแก้ไขข้อมูล (เน้นค้นหาด้วยรหัสและแก้ไขได้ตลอด) ---
st.divider()
st.subheader("🔍 ค้นหาด้วยรหัสหรือชื่อ เพื่อแก้ไขข้อมูล")

# ช่องค้นหา
search_query = st.text_input("พิมพ์รหัสนักศึกษา หรือ ชื่อ-นามสกุล เพื่อค้นหา...")

if not df.empty:
    # กรองข้อมูลตามคำค้นหา
    mask = df['รหัสนักศึกษา'].str.contains(search_query, case=False, na=False) | \
           df['ชื่อ-นามสกุล'].str.contains(search_query, case=False, na=False)
    
    filtered_df = df[mask].copy()

    if not filtered_df.empty:
        st.write(f"พบข้อมูล {len(filtered_df)} รายการ")
        
        # ตารางแก้ไขข้อมูล
        # เราเปิดให้แก้ได้เกือบทุกช่อง ยกเว้นรหัสนักศึกษาที่เป็น Key หลัก
        edited_df = st.data_editor(
            filtered_df,
            column_config={
                "รหัสนักศึกษา": st.column_config.TextColumn("รหัสนักศึกษา (ห้ามแก้)", disabled=True),
                "รุ่น": st.column_config.TextColumn("รุ่น"),
                "ชื่อ-นามสกุล": st.column_config.TextColumn("ชื่อ-นามสกุล"),
                "ระดับชั้น": st.column_config.SelectboxColumn("ระดับชั้น", options=CLASSES),
                "Room": st.column_config.SelectboxColumn("ห้องเรียน", options=ROOMS),
            },
            key="editor_unlimited_v1",
            use_container_width=True
        )

        if st.button("💾 บันทึกการเปลี่ยนแปลงทั้งหมด"):
            # สร้างสำเนาข้อมูลเดิมมาเพื่อรอรับการอัปเดต
            updated_main_df = df.copy()
            
            # วนลูปตามรายการที่แสดงในตารางแก้ไข
            for index, row in edited_df.iterrows():
                std_id = row['รหัสนักศึกษา']
                
                # ใช้ .loc ค้นหาแถวที่มีรหัสนักศึกษาตรงกันในฐานข้อมูลหลัก แล้วทับข้อมูลใหม่ลงไป
                # วิธีนี้จะทำให้ข้อมูลเก่าในห้องเดิมถูกเปลี่ยนเป็นข้อมูลใหม่ทันที (ไม่เกิดชื่อซ้ำ)
                updated_main_df.loc[updated_main_df['รหัสนักศึกษา'] == std_id, 
                                   ['รุ่น', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'Room']] = \
                    [row['รุ่น'], row['ชื่อ-นามสกุล'], row['ระดับชั้น'], row['Room']]

            # ก่อนบันทึกลง Google Sheets ต้องเติม ' นำหน้าเลขรหัสเพื่อป้องกันเลข 0 หาย
            for col in ['รุ่น', 'รหัสนักศึกษา']:
                updated_main_df[col] = updated_main_df[col].apply(lambda x: f"'{str(x).replace(chr(39), '')}")

            # ส่งข้อมูลที่แก้ไขแล้วกลับไปยัง Google Sheets
            conn.update(spreadsheet=st.secrets["gsheet_url"], data=updated_main_df)
            
            st.success("✅ แก้ไขข้อมูลเรียบร้อยแล้ว!")
            st.rerun()
    else:
        if search_query:
            st.warning("❌ ไม่พบข้อมูลรหัสหรือชื่อนี้ในระบบ")
