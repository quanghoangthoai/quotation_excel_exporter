import frappe
import io
import os
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font

@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    file_path = frappe.get_site_path("public", "files", "mẫu báo giá.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active

    font = Font(name="Times New Roman", size=13)
    currency_format = '#,##0.00" đ"'

    # Customer name
    cell = ws["B9"]
    cell.value = customer.customer_name or ""
    cell.font = font

    # Phone from Contact
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")

    contact_mobile = ""
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        contact_mobile = contact.mobile_no or contact.phone or ""

    phone_cell = ws["J9"]
    phone_cell.value = contact_mobile
    phone_cell.font = font
    phone_cell.alignment = Alignment(horizontal="left", vertical="center")

    # Address
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")

    address_line1 = ""
    if address_name:
        address = frappe.get_doc("Address", address_name)
        address_line1 = address.address_line1 or ""

    ws["B10"] = address_line1
    ws["B10"].font = font

    # Items
    start_row = 14
    for i, item in enumerate(quotation.items):
        row = start_row + i

        # Tạo ô chủ động bằng ws.cell() để đảm bảo luôn tồn tại
        cell_a = ws.cell(row=row, column=1, value=i + 1)
        cell_a.font = font
        cell_a.alignment = Alignment(horizontal="center", vertical="top")

        # Merge B:D for item_name
        bd_range = f"B{row}:D{row}"
        for m in list(ws.merged_cells.ranges):
            if bd_range == str(m):
                ws.unmerge_cells(bd_range)
        ws.merge_cells(bd_range)
        cell_name = ws.cell(row=row, column=2)
        cell_name.value = item.item_name
        cell_name.font = font
        cell_name.alignment = Alignment(wrap_text=True, vertical="top")

        # Merge E:F for size
        ef_range = f"E{row}:F{row}"
        for m in list(ws.merged_cells.ranges):
            if ef_range == str(m):
                ws.unmerge_cells(ef_range)
        ws.merge_cells(ef_range)
        cell_desc = ws.cell(row=row, column=5)
        cell_desc.value = item.size or ""
        cell_desc.font = font
        cell_desc.alignment = Alignment(wrap_text=True, vertical="top")

        ws.cell(row=row, column=7, value=item.item_code).font = font  # G
        ws.cell(row=row, column=8, value=item.qty).font = font        # H
        ws.cell(row=row, column=12, value=item.rate or 0).font = font  # L
        amt_cell = ws.cell(row=row, column=14, value=item.amount or (item.qty * item.rate))  # N
        amt_cell.font = font
        amt_cell.number_format = currency_format

        # Merge I:J regardless, then insert image if exists
        ij_range = f"I{row}:J{row}"
        for m in list(ws.merged_cells.ranges):
            if ij_range == str(m):
                ws.unmerge_cells(ij_range)
        ws.merge_cells(ij_range)

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
                    ws.add_image(img, f"I{row}")
                    ws.row_dimensions[row].height = 80
            except:
                pass
        else:
            ws.row_dimensions[row].height = 20  # default height

    # Tổng cộng sau danh sách sản phẩm
    total_row = start_row + len(quotation.items) + 1
    for i in range(4):
        r = total_row + i
        cell = ws.cell(row=r, column=14)
        cell.value = quotation.total if i in [0, 3] else 0
        cell.font = font
        cell.number_format = currency_format

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
