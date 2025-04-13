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
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # Logo
    logo_path = frappe.get_site_path("public", "files", "logo.jpg")
    if os.path.exists(logo_path):
        img = XLImage(logo_path)
        img.width, img.height = 120, 50
        ws.add_image(img, "A1")

    # Header
    ws.merge_cells("C1:N1")
    ws["C1"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["C1"].font = bold_font

    ws["A2"] = "Địa chỉ :"
    ws["B2"] = "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức , HCM"
    ws["A3"] = "Hotline :"
    ws["B3"] = "0768.927..526 - 033.566.9526"
    ws["B4"] = "https://thehome.com.vn/"
    ws["B4"].font = Font(color="0000FF", underline="single", name="Times New Roman")

    ws.merge_cells("A6:N6")
    ws["A6"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["A6"].font = bold_font
    ws["A6"].alignment = Alignment(horizontal="center")

    ws["A8"] = "Khách hàng :"
    ws["B8"] = customer.customer_name or ""

    ws["A9"] = "Địa chỉ :"
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")
    address = frappe.get_doc("Address", address_name) if address_name else None
    ws["B9"] = address.address_line1 if address else ""

    ws["J8"] = "Điện thoại :"
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        ws["K8"] = contact.mobile_no or contact.phone or ""

    headers = [
        "STT", "Tên sản phẩm", "Kích thước sản phẩm", "Mã hàng", "SL",
        "Hình ảnh", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"
    ]
    ws.append([None])
    ws.append(headers)

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=11, column=col)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    start_row = 12
    for i, item in enumerate(quotation.items):
        row = start_row + i
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value=item.item_name)
        ws.cell(row=row, column=3, value=item.size or "")
        ws.cell(row=row, column=4, value=item.item_code)
        ws.cell(row=row, column=5, value=item.qty)
        ws.cell(row=row, column=7, value="Bộ")
        ws.cell(row=row, column=8, value=item.rate or 0)
        ws.cell(row=row, column=9, value=item.discount_percentage or 0)
        ws.cell(row=row, column=10, value=item.amount or (item.qty * item.rate))

        for col in range(1, 11):
            cell = ws.cell(row=row, column=col)
            cell.font = font_13
            cell.alignment = Alignment(vertical="top")
            cell.border = border

        if item.image:
            try:
                image_path = ""
                if item.image.startswith("/files/"):
                    image_path = frappe.get_site_path("public", item.image.lstrip("/"))
                elif item.image.startswith("http"):
                    tmp_path = f"/tmp/tmp_item_{i}.png"
                    with open(tmp_path, "wb") as f:
                        f.write(requests.get(item.image).content)
                    image_path = tmp_path

                if os.path.exists(image_path):
                    img = XLImage(image_path)
                    img.width = 100
                    img.height = 100
                    ws.add_image(img, f"F{row}")
                    ws.row_dimensions[row].height = 80
            except:
                ws.row_dimensions[row].height = 20
        else:
            ws.row_dimensions[row].height = 20

    row = start_row + len(quotation.items)
    ws.cell(row=row, column=1, value="A")
    ws.cell(row=row, column=2, value="Tổng cộng")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    ws.cell(row=row, column=10, value=quotation.total)

    ws.cell(row=row + 1, column=1, value="B")
    ws.cell(row=row + 1, column=2, value="Phụ phí")
    ws.merge_cells(start_row=row + 1, start_column=2, end_row=row + 1, end_column=9)
    ws.cell(row=row + 1, column=10, value=0)

    ws.cell(row=row + 2, column=1, value="C")
    ws.cell(row=row + 2, column=2, value="Đã thanh toán")
    ws.merge_cells(start_row=row + 2, start_column=2, end_row=row + 2, end_column=9)
    ws.cell(row=row + 2, column=10, value=0)

    ws.cell(row=row + 3, column=2, value="Tổng tiền thanh toán (A+B-C)")
    ws.merge_cells(start_row=row + 3, start_column=2, end_row=row + 3, end_column=9)
    ws.cell(row=row + 3, column=10, value=quotation.total)

    ws.cell(row=row + 5, column=1, value="Khách hàng")
    ws.cell(row=row + 5, column=5, value="Người giao hàng")
    ws.cell(row=row + 5, column=8, value="Ngày")
    ws.cell(row=row + 5, column=9, value="Tháng")
    ws.cell(row=row + 5, column=10, value="Năm")

    ws.cell(row=row + 6, column=1, value="(Ký và ghi rõ họ tên)")
    ws.cell(row=row + 6, column=5, value="(Ký và ghi rõ họ tên)")
    ws.cell(row=row + 6, column=8, value="(Ký và ghi rõ họ tên)")

    ws.cell(row=row + 8, column=1, value="Lưu ý :")
    ws.cell(row=row + 8, column=2, value="Không đổi trả sản mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất")
    ws.cell(row=row + 9, column=1, value="Hình  thức thanh toán :")
    ws.cell(row=row + 10, column=1, value="- Thanh toán 100% giá trị đơn hàng khi nhận được hàng")
    ws.cell(row=row + 11, column=1, value="- Đặt hàng đặt cọc trước 30% giá trị đơn hàng")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
