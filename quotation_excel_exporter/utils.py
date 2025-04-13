import frappe
import io
import os
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from copy import copy

@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    file_path = frappe.get_site_path("public", "files", "mẫu báo giá.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active  # Work directly on active sheet instead of creating new one

    # Styles
    font = Font(name="Times New Roman", size=13)
    currency_format = '#,##0" đ"'
    center_align = Alignment(horizontal='center', vertical='center')
    wrap_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Customer info
    ws["B9"] = customer.customer_name or ""
    ws["B9"].font = font
    ws["B9"].alignment = wrap_align

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
    ws["B10"].alignment = wrap_align

    # Clear existing items rows
    start_row = 14
    max_row = start_row + 20  # Adjust this based on your template
    for row in range(start_row, max_row):
        for col in range(1, 15):  # Columns A to N
            cell = ws.cell(row=row, column=col)
            cell.value = None
            
    # Items section
    for i, item in enumerate(quotation.items):
        row = start_row + i

        # STT
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=1).font = font
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=1).border = thin_border

        # Tên sản phẩm (B:D)
        # First set values and styles for all cells in range
        for col in range(2, 5):  # B to D
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.font = font
            cell.alignment = wrap_align
            if col == 2:
                cell.value = item.item_name
        # Then merge
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)

        # Kích thước (E:F)
        # First set values and styles for all cells in range
        for col in range(5, 7):  # E to F
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.font = font
            cell.alignment = wrap_align
            if col == 5:
                cell.value = item.size or ""
        # Then merge
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)

        # Regular cells
        ws.cell(row=row, column=7, value=item.item_code).font = font
        ws.cell(row=row, column=7).alignment = center_align
        ws.cell(row=row, column=7).border = thin_border

        ws.cell(row=row, column=8, value=item.qty).font = font
        ws.cell(row=row, column=8).alignment = center_align
        ws.cell(row=row, column=8).border = thin_border

        # Image cells (I:J)
        # First set borders for all cells in range
        for col in range(9, 11):  # I to J
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
        # Then merge
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)

        # Rate and Amount
        ws.cell(row=row, column=12, value=item.rate or 0).font = font
        ws.cell(row=row, column=12).alignment = center_align
        ws.cell(row=row, column=12).number_format = currency_format
        ws.cell(row=row, column=12).border = thin_border

        amt_cell = ws.cell(row=row, column=14)
        amt_cell.value = item.amount or (item.qty * item.rate)
        amt_cell.font = font
        amt_cell.alignment = center_align
        amt_cell.number_format = currency_format
        amt_cell.border = thin_border

        # Handle images
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

    # Totals section
    total_row = start_row + len(quotation.items) + 1
    for i in range(4):
        r = total_row + i
        cell = ws.cell(row=r, column=14)
        cell.value = quotation.total if i in [0, 3] else 0
        cell.font = font
        cell.alignment = center_align
        cell.number_format = currency_format
        cell.border = thin_border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
