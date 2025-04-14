@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    wb = Workbook()
    ws = wb.active
    ws.title = "Báo giá"

    # Styles
    font_13 = Font(name="Times New Roman", size=13)
    bold_font = Font(name="Times New Roman", size=13, bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Logo
    logo_path = frappe.get_site_path("public", "files", "logo.jpg")
    if os.path.exists(logo_path):
        logo_img = XLImage(logo_path)
        logo_img.width = 180
        logo_img.height = 60
        ws.add_image(logo_img, "A1")

    # Company info
    ws.merge_cells("C1:N1")
    ws["C1"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["C1"].font = bold_font
    ws["C1"].alignment = center_alignment

    ws["A3"] = "Địa chỉ :"
    ws["B3"] = "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức , HCM"
    ws["A4"] = "Hotline :"
    ws["B4"] = "0768.927..526 - 033.566.9526"
    ws["A5"] = "Website :"
    ws["B5"] = "https://thehome.com.vn/"

    for cell in ["A3", "B3", "A4", "B4", "A5", "B5"]:
        ws[cell].font = font_13
        ws[cell].alignment = left_alignment

    # Title
    ws.merge_cells("A6:N6")
    ws["A6"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["A6"].font = bold_font
    ws["A6"].alignment = center_alignment

    # Introduction text
    ws.merge_cells("A8:N8")
    ws["A8"] = "Lời đầu tiên , xin cảm ơn Quý khách hàng đã quan tâm đến sản phẩm nội thất của công ty chúng tôi."
    ws["A8"].font = font_13
    ws["A8"].alignment = left_alignment

    ws.merge_cells("A9:N9")
    ws["A9"] = "Chúng tôi xin gửi đến Quý khách hàng Bảng báo giá như sau :"
    ws["A9"].font = font_13
    ws["A9"].alignment = left_alignment

    # Customer info
    ws["A10"] = "Khách hàng :"
    ws["B10"] = customer.customer_name or ""
    ws["I10"] = "Điện thoại :"
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        ws["J10"] = contact.mobile_no or contact.phone or ""

    ws["A11"] = "Địa chỉ :"
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")
    if address_name:
        address = frappe.get_doc("Address", address_name)
        ws["B11"] = address.address_line1 or ""

    # Table headers
    headers = ["STT", "Tên sản phẩm", "", "", "Kích thước sản phẩm", "", "Mã hàng", "SL", 
              "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"]
    row_num = 13
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col)
        cell.value = header
        cell.font = bold_font
        cell.alignment = center_alignment
        cell.border = thin_border
        cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

    # Merge header cells
    ws.merge_cells(f"B{row_num}:D{row_num}")  # Tên sản phẩm
    ws.merge_cells(f"E{row_num}:F{row_num}")  # Kích thước
    ws.merge_cells(f"I{row_num}:J{row_num}")  # Hình ảnh

    # Items
    for i, item in enumerate(quotation.items, 1):
        row = row_num + i
        
        # Apply borders and alignment to all cells in the row
        for col in range(1, 15):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.font = font_13
            cell.alignment = center_alignment if col in [1, 8, 11, 12, 13, 14] else left_alignment

        # Fill data
        ws.cell(row=row, column=1, value=i)
        ws.merge_cells(f"B{row}:D{row}")
        ws.cell(row=row, column=2, value=item.item_name)
        ws.merge_cells(f"E{row}:F{row}")
        ws.cell(row=row, column=5, value=item.size or "")
        ws.cell(row=row, column=7, value=item.item_code)
        ws.cell(row=row, column=8, value=item.qty)
        ws.merge_cells(f"I{row}:J{row}")
        ws.cell(row=row, column=11, value="Bộ")
        ws.cell(row=row, column=12, value=item.rate)
        ws.cell(row=row, column=13, value=item.discount_percentage)
        ws.cell(row=row, column=14, value=item.amount)

        # Handle image
        if item.image:
            try:
                image_path = ""
                if item.image.startswith("/files/"):
                    image_path = frappe.get_site_path("public", item.image.lstrip("/"))
                elif item.image.startswith("http"):
                    response = requests.get(item.image, timeout=5)
                    if response.status_code == 200:
                        tmp_path = f"/tmp/tmp_item_{i}.png"
                        with open(tmp_path, "wb") as f:
                            f.write(response.content)
                        image_path = tmp_path

                if os.path.exists(image_path):
                    img = XLImage(image_path)
                    img.width = 100
                    img.height = 100
                    ws.add_image(img, f"I{row}")
                    ws.row_dimensions[row].height = 80

    # Totals section
    current_row = row_num + len(quotation.items) + 1
    
    # A. Tổng cộng
    ws.cell(row=current_row, column=1, value="A").font = font_13
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Tổng cộng").font = font_13
    ws.cell(row=current_row, column=14, value=quotation.total).font = font_13

    # B. Phụ phí
    current_row += 1
    ws.cell(row=current_row, column=1, value="B").font = font_13
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Phụ phí").font = font_13
    ws.cell(row=current_row, column=14, value=0).font = font_13

    # C. Đã thanh toán
    current_row += 1
    ws.cell(row=current_row, column=1, value="C").font = font_13
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Đã thanh toán").font = font_13
    ws.cell(row=current_row, column=14, value=0).font = font_13

    # Tổng tiền thanh toán
    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=1, value="Tổng tiền thanh toán (A+B-C)").font = font_13
    ws.cell(row=current_row, column=14, value=quotation.total).font = font_13

    # Signature section
    current_row += 2
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
    ws.cell(row=current_row, column=1, value="Khách hàng").font = bold_font
    ws.cell(row=current_row, column=1).alignment = center_alignment

    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=9)
    ws.cell(row=current_row, column=6, value="Người giao hàng").font = bold_font
    ws.cell(row=current_row, column=6).alignment = center_alignment

    ws.merge_cells(start_row=current_row, start_column=11, end_row=current_row, end_column=14)
    ws.cell(row=current_row, column=11, value="Ngày     Tháng     Năm").font = bold_font
    ws.cell(row=current_row, column=11).alignment = center_alignment

    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
    ws.cell(row=current_row, column=1, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=current_row, column=1).alignment = center_alignment

    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=9)
    ws.cell(row=current_row, column=6, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=current_row, column=6).alignment = center_alignment

    # Note section
    current_row += 2
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=14)
    ws.cell(row=current_row, column=1, value="Lưu ý :     Không đổi trả sản mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất").font = bold_font

    # Set column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 5
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['K'].width = 8
    ws.column_dimensions['L'].width = 12
    ws.column_dimensions['M'].width = 8
    ws.column_dimensions['N'].width = 12

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
