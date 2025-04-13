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

    # Header (Logo, Thông tin công ty, tiêu đề)
    logo_path = frappe.get_site_path("public", "files", "logo.jpg")
    if os.path.exists(logo_path):
        img = XLImage(logo_path)
        img.width, img.height = 120, 50
        ws.add_image(img, "A1")

    ws.merge_cells("C2:N2")
    ws["C2"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["C2"].font = bold_font

    ws["A4"] = "Địa chỉ :"
    ws["B4"] = "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức , HCM"
    ws["A5"] = "Hotline :"
    ws["B5"] = "0768.927..526 - 033.566.9526"
    ws["B6"] = "https://thehome.com.vn/"
    ws["B6"].font = Font(color="0000FF", underline="single", name="Times New Roman")

    ws.merge_cells("A8:N8")
    ws["A8"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["A8"].font = bold_font
    ws["A8"].alignment = center_alignment

    # Khách hàng và liên hệ
    ws["A10"] = "Khách hàng :"
    ws["B10"] = customer.customer_name or ""
    ws["A11"] = "Địa chỉ :"
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")
    address = frappe.get_doc("Address", address_name) if address_name else None
    ws["B11"] = address.address_line1 if address else ""

    ws["J10"] = "Điện thoại :"
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        ws["K10"] = contact.mobile_no or contact.phone or ""

    # Table header
    headers = [
        "STT", "Tên sản phẩm", "", "", "Kích thước sản phẩm", "Mã hàng", "", "SL",
        "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"
    ]
    ws.append([None])
    ws.append(headers)

    for col in range(1, len(headers)+1):
        cell = ws.cell(row=13, column=col)
        cell.font = bold_font
        cell.alignment = center_alignment
        cell.border = thin_border

    ws.merge_cells("B13:D13")
    ws.merge_cells("F13:G13")
    ws.merge_cells("I13:J13")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
