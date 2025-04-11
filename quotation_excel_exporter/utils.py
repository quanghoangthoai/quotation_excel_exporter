import frappe
import io
import os
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from copy import copy

@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    # Load template file
    file_path = frappe.get_site_path("public", "files", "mẫu báo giá.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active

    # Unmerge all cells first
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    font = Font(name="Times New Roman", size=13)
    currency_format = '#,##0.00" đ"'

    # Customer name
    ws["B9"] = customer.customer_name or ""
    ws["B9"].font = font

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

    ws["J9"] = contact_mobile
    ws["J9"].font = font
    ws["J9"].alignment = Alignment(horizontal="left", vertical="center")

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

        # Item number
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=1).font = font
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center", vertical="top")

        # Item name (B:D)
        cell_name = ws.cell(row=row, column=2)
        cell_name.value = item.item_name
        cell_name.font = font
        cell_name.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)

        # Size (E:F)
        cell_size = ws.cell(row=row, column=5)
        cell_size.value = item.size or ""
        cell_size.font = font
        cell_size.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)

        # Regular cells
        ws.cell(row=row, column=7, value=item.item_code).font = font  # G
        ws.cell(row=row, column=8, value=item.qty).font = font        # H
        ws.cell(row=row, column=12, value=item.rate or 0).font = font  # L
        
        # Amount
        amt_cell = ws.cell(row=row, column=14)
        amt_cell.value = item.amount or (item.qty * item.rate)
        amt_cell.font = font
        amt_cell.number_format = currency_format

        # Image cells (I:J)
        img_cell = ws.cell(row=row, column=9)
        img_cell.value = ""  # Clear any existing value
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
        
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
                ws.row_dimensions[row].height = 20
        else:
            ws.row_dimensions[row].height = 20

    # Totals
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
