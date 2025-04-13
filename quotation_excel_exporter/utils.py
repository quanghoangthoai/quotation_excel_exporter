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
 # Đường dẫn tới logo
    logo_path = frappe.get_site_path("public", "files", "logo.jpg")
    if os.path.exists(logo_path):
        logo_img = XLImage(logo_path)
        logo_img.width = 180  # Điều chỉnh kích thước ngang
        logo_img.height = 60  # Điều chỉnh kích thước dọc

    # Chèn vào đúng ô A1
    ws.add_image(logo_img, "A1")

    # Tuỳ chọn: chỉnh chiều cao hàng nếu cần để logo không bị cắt
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30
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

    # Store template row formatting
    template_row = 14
    template_styles = {}
    for col in range(1, 15):
        cell = ws.cell(row=template_row, column=col)
        template_styles[col] = {
            'font': copy(cell.font) if cell.font else None,
            'border': copy(cell.border) if cell.border else None,
            'fill': copy(cell.fill) if cell.fill else None,
            'number_format': cell.number_format,
            'alignment': copy(cell.alignment) if cell.alignment else None
        }

    # Get merged ranges in template row
    template_merges = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row == template_row:
            template_merges.append((
                merged_range.min_col,
                merged_range.max_col
            ))

    # Unmerge template row cells first
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row == template_row:
            ws.unmerge_cells(str(merged_range))

    # Insert rows for additional items
    num_items = len(quotation.items)
    if num_items > 1:
        ws.insert_rows(template_row + 1, num_items - 1)

    # Fill items
    for i, item in enumerate(quotation.items):
        row = template_row + i
        
        # First set all cell values and styles
        for col in range(1, 15):
            cell = ws.cell(row=row, column=col)
            # Apply template styles
            if template_styles[col]['font']:
                cell.font = copy(template_styles[col]['font'])
            if template_styles[col]['border']:
                cell.border = copy(template_styles[col]['border'])
            if template_styles[col]['fill']:
                cell.fill = copy(template_styles[col]['fill'])
            if template_styles[col]['alignment']:
                cell.alignment = copy(template_styles[col]['alignment'])
            cell.number_format = template_styles[col]['number_format']

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

        # Then merge cells
        for min_col, max_col in template_merges:
            ws.merge_cells(
                start_row=row,
                start_column=min_col,
                end_row=row,
                end_column=max_col
            )

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
    total_row = template_row + num_items
    
    # Move and update totals section
    ws.cell(row=total_row, column=1, value="A")
    ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=13)
    ws.cell(row=total_row, column=2, value="Tổng cộng")
    ws.cell(row=total_row, column=14, value=quotation.total)

    ws.cell(row=total_row + 1, column=1, value="B")
    ws.merge_cells(start_row=total_row + 1, start_column=2, end_row=total_row + 1, end_column=13)
    ws.cell(row=total_row + 1, column=2, value="Phụ phí")
    ws.cell(row=total_row + 1, column=14, value=0)

    ws.cell(row=total_row + 2, column=1, value="C")
    ws.merge_cells(start_row=total_row + 2, start_column=2, end_row=total_row + 2, end_column=13)
    ws.cell(row=total_row + 2, column=2, value="Đã thanh toán")
    ws.cell(row=total_row + 2, column=14, value=0)
        # Tổng tiền thanh toán (A+B-C)
    cell = ws.cell(row=total_row + 3, column=2)
    cell.value = "Tổng tiền thanh toán (A+B-C)"
    ws.merge_cells(start_row=total_row + 3, start_column=2, end_row=total_row + 3, end_column=13)

ws.cell(row=total_row + 3, column=14, value=quotation.total)


    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
