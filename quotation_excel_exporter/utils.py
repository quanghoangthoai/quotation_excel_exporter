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

    # Address from Address doctype
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
        ws[f"A{row}"] = i + 1
        ws[f"A{row}"].font = font

        # Merge B:D for item_name
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=2)
        cell.value = item.item_name
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

        # # Ghi item.size vào E:F (unmerge -> gán -> merge lại)
        # ef_range = f"E{row}:F{row}"
        # for m in list(ws.merged_cells.ranges):
        #     if ef_range == str(m):
        #         ws.unmerge_cells(ef_range)

        # cell_desc = ws.cell(row=row, column=5)
        # cell_desc.value = item.size or ""
        # cell_desc.font = font
        # cell_desc.alignment = Alignment(wrap_text=True, vertical="top")

        # ws.merge_cells(ef_range)

        ws[f"G{row}"] = item.item_code
        ws[f"G{row}"].font = font
        ws[f"H{row}"] = item.qty
        ws[f"H{row}"].font = font
        ws[f"L{row}"] = item.rate or 0
        ws[f"L{row}"].font = font
        ws[f"N{row}"] = item.amount or (item.qty * item.rate)
        ws[f"N{row}"].font = font
        ws[f"N{row}"].number_format = currency_format

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

    # Total
    for r in range(15, 19):
        cell = ws.cell(row=r, column=14)
        cell.value = quotation.total if r in [15, 18] else 0
        cell.font = font
        cell.number_format = currency_format

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
