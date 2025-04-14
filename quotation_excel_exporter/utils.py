import frappe
import io
import os
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    wb = Workbook()
    ws = wb.active
    ws.title = "Báo giá"

    font_13 = Font(name="Times New Roman", size=13)
    bold_font = Font(name="Times New Roman", size=13, bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    col_widths = [5, 30, 5, 5, 15, 5, 15, 5, 12, 5, 10, 12, 10, 15]
    for i, width in enumerate(col_widths, 1):
        col_letter = chr(64 + i) if i <= 26 else 'A' + chr(64 + (i - 26))
        ws.column_dimensions[col_letter].width = width

    logo_path = frappe.get_site_path("public", "files", "logo.jpg")
    if os.path.exists(logo_path):
        logo_img = XLImage(logo_path)
        logo_img.width = 130
        logo_img.height = 100
        ws.add_image(logo_img, "A1")
    ws.row_dimensions[1].height = 50
    ws.row_dimensions[2].height = 50
    ws.row_dimensions[3].height = 50

    ws.merge_cells("C1:N1")
    ws["C1"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["C1"].font = Font(name="Times New Roman", size=18, bold=True)
    ws["C1"].alignment = center_alignment

    info = [
        ("A3", "Địa chỉ :", "B3", "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức , HCM"),
        ("A4", "Hotline :", "B4", "0768.927.526 - 033.566.9526"),
        ("A5", "Website :", "B5", "https://thehome.com.vn/")
    ]
    for a_cell, a_text, b_cell, b_text in info:
        ws[a_cell] = a_text
        ws[b_cell] = b_text
        ws[a_cell].font = font_13
        ws[b_cell].font = font_13
        ws[a_cell].alignment = left_alignment
        ws[b_cell].alignment = left_alignment

    ws.merge_cells("C6:N6")
    ws["C6"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["C6"].font = Font(name="Times New Roman", size=16, bold=True)
    ws["C6"].alignment = center_alignment

    ws["A8"] = "Khách hàng :"
    ws["B8"] = customer.customer_name or ""
    ws["A9"] = "Địa chỉ :"
    address = frappe.db.get_value("Dynamic Link", {"link_doctype": "Customer", "link_name": customer.name, "parenttype": "Address"}, "parent")
    if address:
        address_doc = frappe.get_doc("Address", address)
        ws["B9"] = address_doc.address_line1 or ""

    ws["I8"] = "Điện thoại :"
    contact = frappe.db.get_value("Dynamic Link", {"link_doctype": "Customer", "link_name": customer.name, "parenttype": "Contact"}, "parent")
    if contact:
        contact_doc = frappe.get_doc("Contact", contact)
        ws["J8"] = contact_doc.mobile_no or contact_doc.phone or ""

    ws.merge_cells("A10:N10")
    ws["A10"] = "Lời đầu tiên , xin cảm ơn Quý khách hàng đã quan tâm đến sản phẩm nội thất của công ty chúng tôi."
    ws["A10"].font = font_13
    ws["A10"].alignment = left_alignment

    ws.merge_cells("A11:N11")
    ws["A11"] = "Chúng tôi xin gửi đến Quý khách hàng Bảng báo giá như sau :"
    ws["A11"].font = font_13
    ws["A11"].alignment = left_alignment

    headers = ["STT", "Tên sản phẩm", "", "", "Kích thước sản phẩm", "", "Mã hàng", "SL", "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"]
    row_num = 13
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col)
        cell.value = header
        cell.font = bold_font
        cell.alignment = center_alignment
        cell.border = thin_border
        cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

    ws.merge_cells(f"B{row_num}:D{row_num}")
    ws.merge_cells(f"E{row_num}:F{row_num}")
    ws.merge_cells(f"I{row_num}:J{row_num}")

    for i, item in enumerate(quotation.items, 1):
        row = row_num + i
        ws.cell(row=row, column=1, value=i)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.cell(row=row, column=2, value=item.item_name)
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
        ws.cell(row=row, column=5, value=item.size or "")
        ws.cell(row=row, column=7, value=item.item_code)
        ws.cell(row=row, column=8, value=item.qty)
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
        ws.cell(row=row, column=11, value="Bộ")
        ws.cell(row=row, column=12, value=item.rate)
        ws.cell(row=row, column=13, value=item.discount_percentage)
        ws.cell(row=row, column=14, value=item.amount)

        for col in range(1, 15):
            cell = ws.cell(row=row, column=col)
            cell.font = font_13
            cell.border = thin_border
            cell.alignment = center_alignment if col in [1, 8, 11, 12, 13, 14] else left_alignment

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
                else:
                    ws.row_dimensions[row].height = 20
            except Exception as e:
                frappe.log_error(f"Hình ảnh lỗi: {item.image} - {str(e)}")
                ws.row_dimensions[row].height = 20
        else:
            ws.row_dimensions[row].height = 20

    row = row_num + len(quotation.items) + 1
    ws.cell(row=row, column=1, value="A")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
    ws.cell(row=row, column=2, value="Tổng cộng")
    ws.cell(row=row, column=14, value=quotation.total)

    ws.cell(row=row+1, column=1, value="B")
    ws.merge_cells(start_row=row+1, start_column=2, end_row=row+1, end_column=13)
    ws.cell(row=row+1, column=2, value="Phụ phí")
    ws.cell(row=row+1, column=14, value=0)

    ws.cell(row=row+2, column=1, value="C")
    ws.merge_cells(start_row=row+2, start_column=2, end_row=row+2, end_column=13)
    ws.cell(row=row+2, column=2, value="Đã thanh toán")
    ws.cell(row=row+2, column=14, value=0)

    ws.merge_cells(start_row=row+3, start_column=1, end_row=row+3, end_column=13)
    ws.cell(row=row+3, column=1, value="Tổng tiền thanh toán (A+B-C)")
    ws.cell(row=row+3, column=14, value=quotation.total)

    ws.merge_cells(start_row=row+5, start_column=1, end_row=row+5, end_column=4)
    ws.cell(row=row+5, column=1, value="Khách hàng").font = bold_font
    ws.merge_cells(start_row=row+5, start_column=6, end_row=row+5, end_column=9)
    ws.cell(row=row+5, column=6, value="Người giao hàng").font = bold_font
    ws.merge_cells(start_row=row+5, start_column=11, end_row=row+5, end_column=14)
    ws.cell(row=row+5, column=11, value="Ngày     Tháng     Năm").font = bold_font

    ws.merge_cells(start_row=row+6, start_column=1, end_row=row+6, end_column=4)
    ws.cell(row=row+6, column=1, value="(Ký và ghi rõ họ tên)")
    ws.merge_cells(start_row=row+6, start_column=6, end_row=row+6, end_column=9)
    ws.cell(row=row+6, column=6, value="(Ký và ghi rõ họ tên)")

    ws.merge_cells(start_row=row+8, start_column=1, end_row=row+8, end_column=14)
    ws.cell(row=row+8, column=1, value="Lưu ý :     Không đổi trả sản mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất").font = bold_font

    ws.merge_cells(start_row=row+9, start_column=1, end_row=row+9, end_column=14)
    ws.cell(row=row+9, column=1, value="Hình thức thanh toán:").font = font_13

    ws.merge_cells(start_row=row+10, start_column=1, end_row=row+10, end_column=14)
    ws.cell(row=row+10, column=1, value="- Thanh toán 100% giá trị đơn hàng khi nhận được hàng").font = font_13

    ws.merge_cells(start_row=row+11, start_column=1, end_row=row+11, end_column=14)
    ws.cell(row=row+11, column=1, value="- Đặt hàng đặt cọc trước 30% giá trị đơn hàng").font = font_13

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
