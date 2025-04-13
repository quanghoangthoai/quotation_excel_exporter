import frappe
import io
import os
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side

@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    wb = Workbook()
    ws = wb.active
    ws.title = "Báo giá"

    font_13 = Font(name="Times New Roman", size=13)
    bold_font = Font(name="Times New Roman", size=13, bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Logo đứng
    logo_path = frappe.get_site_path("public", "files", "z6473642459612_58e86d169bb72c78b360392b4f81e8bae2152f.jpg")
    if os.path.exists(logo_path):
        logo_img = XLImage(logo_path)
        logo_img.width = 90
        logo_img.height = 130
        ws.add_image(logo_img, "A1")
        ws.row_dimensions[1].height = 70
        ws.row_dimensions[2].height = 70
        ws.row_dimensions[3].height = 70

    # Company name
    ws.merge_cells("C2:N2")
    ws["C2"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["C2"].font = bold_font
    ws["C2"].alignment = center_alignment

    # Company info
    ws["A3"] = "Địa chỉ :"
    ws["B3"] = "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức , HCM"
    ws["A4"] = "Hotline :"
    ws["B4"] = "0768.927..526 - 033.566.9526"
    ws["B5"] = "https://thehome.com.vn/"
    ws["B5"].font = Font(color="0000FF", underline="single", name="Times New Roman")

    # Title
    ws.merge_cells("A7:N7")
    ws["A7"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["A7"].font = bold_font
    ws["A7"].alignment = center_alignment

    # Customer info
    ws["A9"] = "Khách hàng :"
    ws["B9"] = customer.customer_name or ""
    ws["A10"] = "Địa chỉ :"
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")
    address = frappe.get_doc("Address", address_name) if address_name else None
    ws["B10"] = address.address_line1 if address else ""

    ws["J9"] = "Điện thoại :"
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        ws["K9"] = contact.mobile_no or contact.phone or ""

    # Table header
    headers = [
        "STT", "Tên sản phẩm", "", "", "Kích thước sản phẩm", "Mã hàng", "", "SL",
        "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"
    ]
    ws.append([None])
    ws.append(headers)

    for col in range(1, len(headers)+1):
        cell = ws.cell(row=12, column=col)
        cell.font = bold_font
        cell.alignment = center_alignment
        cell.border = thin_border

    ws.merge_cells("B12:D12")
    ws.merge_cells("F12:G12")
    ws.merge_cells("I12:J12")

    # Dòng sản phẩm từ quotation.items
    total = 0
    for i, item in enumerate(quotation.items):
        row = 13 + i
        ws.cell(row=row, column=1, value=i + 1).font = font_13
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.cell(row=row, column=2, value=item.item_name).font = font_13
        ws.cell(row=row, column=5, value=item.size).font = font_13
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=7)
        ws.cell(row=row, column=6, value=item.item_code).font = font_13
        ws.cell(row=row, column=8, value=item.qty).font = font_13
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
        ws.cell(row=row, column=11, value="Bộ").font = font_13
        ws.cell(row=row, column=12, value=item.rate).font = font_13
        ws.cell(row=row, column=13, value=item.discount_percentage).font = font_13
        amount = item.amount or (item.qty * item.rate)
        ws.cell(row=row, column=14, value=amount).font = font_13
        total += amount

    row = 13 + len(quotation.items)
    ws.cell(row=row, column=1, value="A").font = font_13
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
    ws.cell(row=row, column=2, value="Tổng cộng").font = font_13
    ws.cell(row=row, column=14, value=total).font = font_13

    ws.cell(row=row + 1, column=1, value="B").font = font_13
    ws.merge_cells(start_row=row + 1, start_column=2, end_row=row + 1, end_column=13)
    ws.cell(row=row + 1, column=2, value="Phụ phí").font = font_13
    ws.cell(row=row + 1, column=14, value=0).font = font_13

    ws.cell(row=row + 2, column=1, value="C").font = font_13
    ws.merge_cells(start_row=row + 2, start_column=2, end_row=row + 2, end_column=13)
    ws.cell(row=row + 2, column=2, value="Đã thanh toán").font = font_13
    ws.cell(row=row + 2, column=14, value=0).font = font_13

    ws.merge_cells(start_row=row + 3, start_column=1, end_row=row + 3, end_column=13)
    ws.cell(row=row + 3, column=1, value="Tổng tiền thanh toán (A+B-C)").font = font_13
    ws.cell(row=row + 3, column=14, value=total).font = font_13

    # Thêm đoạn cuối mẫu
    footer_row = row + 5
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=3)
    ws.cell(row=footer_row, column=1, value="Khách hàng").font = bold_font

    ws.merge_cells(start_row=footer_row, start_column=5, end_row=footer_row, end_column=7)
    ws.cell(row=footer_row, column=5, value="Người giao hàng").font = bold_font

    ws.merge_cells(start_row=footer_row, start_column=9, end_row=footer_row, end_column=14)
    ws.cell(row=footer_row, column=9, value="Ngày     Tháng     Năm").font = bold_font

    ws.merge_cells(start_row=footer_row + 1, start_column=1, end_row=footer_row + 1, end_column=3)
    ws.cell(row=footer_row + 1, column=1, value="(Ký và ghi rõ họ tên)").font = font_13

    ws.merge_cells(start_row=footer_row + 1, start_column=5, end_row=footer_row + 1, end_column=7)
    ws.cell(row=footer_row + 1, column=5, value="(Ký và ghi rõ họ tên)").font = font_13

    ws.merge_cells(start_row=footer_row + 1, start_column=9, end_row=footer_row + 1, end_column=14)
    ws.cell(row=footer_row + 1, column=9, value="(Ký và ghi rõ họ tên)").font = font_13

    ws.merge_cells(start_row=footer_row + 3, start_column=1, end_row=footer_row + 3, end_column=14)
    ws.cell(row=footer_row + 3, column=1, value="Lưu ý :     Không đổi trả sản mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất").font = bold_font

    ws.merge_cells(start_row=footer_row + 4, start_column=1, end_row=footer_row + 4, end_column=14)
    ws.cell(row=footer_row + 4, column=1, value="Hình  thức thanh toán :").font = bold_font

    ws.merge_cells(start_row=footer_row + 5, start_column=1, end_row=footer_row + 5, end_column=14)
    ws.cell(row=footer_row + 5, column=1, value="- Thanh toán 100% giá trị đơn hàng khi nhận được hàng").font = font_13

    ws.merge_cells(start_row=footer_row + 6, start_column=1, end_row=footer_row + 6, end_column=14)
    ws.cell(row=footer_row + 6, column=1, value="- Đặt hàng đặt cọc trước 30% giá trị đơn hàng").font = font_13

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
