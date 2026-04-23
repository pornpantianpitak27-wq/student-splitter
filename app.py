def create_excel_report(target_year):
    df_data = load_data()
    if df_data.empty: return None
    year_data = df_data[df_data['ระดับชั้น'] == target_year]
    if year_data.empty: return None

    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    bold_font = Font(name='Angsana New', size=15, bold=True)
    normal_font = Font(name='Angsana New', size=14)
    center_align = Alignment(horizontal='center', vertical='center')

    # เรียกใช้ฟังก์ชันหาไฟล์โลโก้
    logo_path = get_existing_logo()

    for r_name in sorted(year_data['Room'].unique()):
        ws = wb.create_sheet(title=f"ห้อง {r_name.replace('/', '-')}")
        room_data = year_data[year_data['Room'] == r_name].sort_values('รหัสนักศึกษา')
        
        # --- ส่วนหัวตารางและเนื้อหา (เหมือนเดิม) ---
        ws.merge_cells('N2:U2'); ws['N2'] = "บัญชีรายชื่อนี้ใช้สำหรับ"; ws['N2'].border = border; ws['N2'].alignment = center_align; ws['N2'].font = bold_font
        ws.merge_cells('N3:O4'); ws['N3'] = "เช็คชื่อนักศึกษา"; ws['N3'].border = border; ws['N3'].alignment = center_align
        ws.merge_cells('P3:R4'); ws['P3'] = "เซ็นสอบกลางภาค"; ws['P3'].border = border; ws['P3'].alignment = center_align
        ws.merge_cells('S3:U4'); ws['S3'] = "เซ็นสอบปลายภาค"; ws['S3'].border = border; ws['S3'].alignment = center_align
        ws.merge_cells('A5:U5'); ws['A5'] = "บัญชีรายชื่อนักศึกษา ภาคเรียนที่ 1 ปีการศึกษา 2568"; ws['A5'].font = bold_font; ws['A5'].alignment = center_align
        ws.merge_cells('A6:U6'); ws['A6'] = f"ระดับ ปวส. ชั้นปีที่ {target_year[2:]} ห้อง {r_name} ศูนย์บางแค"; ws['A6'].font = bold_font; ws['A6'].alignment = center_align
        ws.merge_cells('A8:A10'); ws['A8'] = "เลขที่"
        ws.merge_cells('B8:B10'); ws['B8'] = "รหัสประจำตัว"
        ws.merge_cells('C8:K10'); ws['C8'] = "ชื่อ-สกุล"
        ws['L8'] = "เดือน"; ws['L9'] = "วันที่"; ws['L10'] = "คาบ"
        ws.merge_cells('U8:U10'); ws['U8'] = "หมายเหตุ"
        for i in range(1, 9): ws.cell(row=10, column=12+i).value = i
        for r in range(8, 11):
            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border; cell.alignment = center_align; cell.font = bold_font

        # ข้อมูลนักศึกษา
        for i, row in enumerate(room_data.itertuples(), 1):
            curr = 10 + i
            ws.cell(row=curr, column=1).value = i
            ws.cell(row=curr, column=2).value = row.รหัสนักศึกษา
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=11)
            ws.cell(row=curr, column=3).value = f"{row.ชื่อ} {row.นามสกุล}"
            for c in range(1, 22):
                ws.cell(row=curr, column=c).border = border; ws.cell(row=curr, column=c).alignment = center_align
            ws.cell(row=curr, column=3).alignment = Alignment(horizontal='left', indent=1)

        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 17
        ws.column_dimensions['C'].width = 32
        for c_idx in range(12, 22): ws.column_dimensions[get_column_letter(c_idx)].width = 4

        # --- 🛠️ ส่วนแก้ไขการแทรกโลโก้ (Fix) ---
        if logo_path:
            try:
                # สร้างวัตถุรูปภาพใหม่ทุกครั้งที่วนลูปสร้าง Sheet
                img = Image(logo_path)
                
                # กำหนดขนาด (ลองปรับให้เล็กลงนิดหน่อยเพื่อให้ชัวร์ว่าไม่ล้นช่อง)
                img.width = 85 
                img.height = 85
                
                # เปลี่ยนตำแหน่งวางจาก H1 เป็นจุดที่ว่าง (เช่น ช่วงคอลัมน์ I แถวที่ 1)
                # เพื่อไม่ให้ไปทับกับข้อความที่ Merge ไว้
                ws.add_image(img, 'I1') 
                
            except Exception as e:
                # ถ้ามี Error ให้ข้ามไปก่อนเพื่อไม่ให้โปรแกรมค้าง
                continue

    wb.save(output)
    return output.getvalue()
