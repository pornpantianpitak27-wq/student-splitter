def create_excel_report(target_year):
    df_all = load_data()
    if df_all.empty: return None
    year_data = df_all[df_all['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active) 
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    bold_font = Font(name='Angsana New', size=15, bold=True)
    normal_font = Font(name='Angsana New', size=14)
    center_align = Alignment(horizontal='center', vertical='center')

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # --- จัด Layout ตามที่คุณต้องการ (C-D ชื่อ, E-U ตาราง, V หมายเหตุ) ---
        ws.merge_cells('O2:V2'); ws['O2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['O2'].border = border; ws['O2'].alignment = center_align; ws['O2'].font = bold_font
        ws.merge_cells('O3:P4'); ws['O3'] = "เช็คชื่อนักศึกษา"; ws['O3'].border = border; ws['O3'].alignment = center_align
        ws.merge_cells('Q3:S4'); ws['Q3'] = "เซ็นสอบกลางภาค"; ws['Q3'].border = border; ws['Q3'].alignment = center_align
        ws.merge_cells('T3:V4'); ws['T3'] = "เซ็นสอบปลายภาค"; ws['T3'].border = border; ws['T3'].alignment = center_align

        ws.merge_cells('A5:V5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:V6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align
        
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:D10'); ws['C8'] = "ชื่อ-สกุล"
        ws['E8'] = "เดือน"; ws['E9'] = "วันที่"; ws['E10'] = "คาบ"
        ws.merge_cells('V8:V10'); ws['V8'] = "หมายเหตุ"

        for i in range(1, 17): ws.cell(row=10, column=5+i).value = i
        for r in range(8, 11):
            for c in range(1, 23):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # --- รายชื่อนักศึกษา ---
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=4)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 23):
                ws.cell(row=curr, column=c).border = border; ws.cell(row=curr, column=c).alignment = center_align
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)

        # ปรับความกว้าง
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        for c_idx in range(5, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 3.5
        ws.column_dimensions['V'].width = 12

        # --- 🚀 วิธีแทรกรูปภาพแบบใหม่ (New Insertion Method) ---
        logo_path = get_existing_logo() # จะดึง 1523.jpg หรือ logo_college.jpg มาใช้
        if logo_path:
            try:
                # สร้าง Image Object ข้างใน Loop ทุกครั้ง
                from openpyxl.drawing.image import Image as XLImage
                img = XLImage(logo_path)
                
                # ปรับขนาดให้เล็กลงเล็กน้อยเพื่อความปลอดภัย (75x75 pixels)
                img.width = 75
                img.height = 75
                
                # ใช้ Anchor แบบระบุ Cell ให้ชัดเจน (วางที่ตำแหน่ง H1)
                ws.add_image(img, 'H1') 
            except Exception as e:
                st.write(f"Error inserting image: {e}") # แสดง Error ถ้าแทรกไม่ได้

    wb.save(output)
    return output.getvalue()
