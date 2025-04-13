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
    ws = wb.active

    # Store template row for items
    template_row = 14
    template_cells = {}
    
    # Save template formatting for the item row
    for col in range(1, 15):  # Columns A to N
        cell = ws.cell(row=template_row, column=col)
        template_cells[col] = {
            'font': copy(cell.font) if cell.font else None,
            'border': copy(cell.border) if cell.border else None,
            'fill': copy(cell.fill) if cell.fill else None,
            'number_format': cell.number_format,
            'alignment': copy(cell.alignment) if cell.alignment else None
        }

    # Fill customer info
    ws["B9"] = customer.customer_name or ""
    
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")

    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        ws["I9"] = contact.mobile_no or contact.phone or ""

    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")

    if address_name:
        address = frappe.get_doc("Address", address_name)
        ws["B10"] = address.address_line1 or ""

    # Insert/delete rows based on number of items
    num_items = len(quotation.items)
    if num_items > 1:
        # Insert required number of rows after template row
        ws.insert_rows(template_row + 1, num_items - 1)
        
        # Copy merged cell ranges for each new row
        for i in range(1, num_items):
            new_row = template_row + i
            # Copy B:D merge
            ws.merge_cells(start_row=new_row, start_column=2, end_row=new_row, end_column=4)
            # Copy E:F merge
            ws.merge_cells(start_row=new_row, start_column=5, end_row=new_row, end_column=6)
            # Copy I:J merge
            ws.merge_cells(start_row=new_row, start_column=9, end_row=new_row, end_column=10)

    # Fill items
    start_row = template_row
    for i, item in enumerate(quotation.items):
        row = start_row + i
        
        # Apply template formatting to all cells in the row
        for col in range(1, 15):
            cell = ws.cell(row=row, column=col)
            if template_cells[col]['font']:
                cell.font = copy(template_cells[col]['font'])
            if template_cells[col]['border']:
                cell.border = copy(template_cells[col]['border'])
            if template_cells[col]['fill']:
                cell.fill = copy(template_cells[col]['fill'])
            if template_cells[col]['alignment']:
                cell.alignment = copy(template_cells[col]['alignment'])
            cell.number_format = template_cells[col]['number_format']

        # Fill data
        ws.cell(row=row, column=1, value=i + 1)  # STT
        ws.cell(row=row, column=2, value=item.item_name)  # Tên sản phẩm
        ws.cell(row=row, column=5, value=item.size or "")  # Kích thước
        ws.cell(row=row, column=7, value=item.item_code)  # Mã hàng
        ws.cell(row=row, column=8, value=item.qty)  # SL
        ws.cell(row=row, column=11, value="Bộ")  # Đơn vị
        ws.cell(row=row, column=12, value=item.rate or 0)  # Đơn giá
        ws.cell(row=row, column=13, value=item.discount_percentage or 0)  # CK
        ws.cell(row=row, column=14, value=item.amount or (item.qty * item.rate))  # Thành tiền

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

    # Update totals section position
    total_row = start_row + num_items
    ws.cell(row=total_row, column=1, value="A")
    ws.cell(row=total_row, column=2, value="Tổng cộng")
    ws.cell(row=total_row, column=14, value=quotation.total)

    ws.cell(row=total_row + 1, column=1, value="B")
    ws.cell(row=total_row + 1, column=2, value="Phụ phí")
    ws.cell(row=total_row + 1, column=14, value=0)

    ws.cell(row=total_row + 2, column=1, value="C")
    ws.cell(row=total_row + 2, column=2, value="Đã thanh toán")
    ws.cell(row=total_row + 2, column=14, value=0)

    ws.cell(row=total_row + 3, column=2, value="Tổng tiền thanh toán (A+B-C)")
    ws.cell(row=total_row + 3, column=14, value=quotation.total)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
