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
    template_ws = wb.active
    ws = wb.create_sheet("Generated")

    # Copy template exactly including all formatting
    for row in template_ws.iter_rows():
        for cell in row:
            new_cell = ws[cell.coordinate]
            new_cell.value = cell.value
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    # Copy merged cells
    for merged_range in template_ws.merged_cells.ranges:
        ws.merge_cells(str(merged_range))

    # Copy dimensions
    for row_idx, row in enumerate(template_ws.iter_rows(), start=1):
        if template_ws.row_dimensions[row_idx].height:
            ws.row_dimensions[row_idx].height = template_ws.row_dimensions[row_idx].height

    for col_letter, dim in template_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = dim.width

    # Styles
    font = Font(name="Times New Roman", size=13)
    header_font = Font(name="Times New Roman", size=13, bold=True)
    currency_format = '#,##0" đ"'  # Changed format to match template
    center_align = Alignment(horizontal='center', vertical='center')
    wrap_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Borders
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

    # Items section
    start_row = 14
    for i, item in enumerate(quotation.items):
        row = start_row + i

        # STT
        cell_num = ws.cell(row=row, column=1, value=i + 1)
        cell_num.font = font
        cell_num.alignment = center_align
        cell_num.border = thin_border

        # Tên sản phẩm (B:D)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        name_cell = ws.cell(row=row, column=2, value=item.item_name)
        name_cell.font = font
        name_cell.alignment = wrap_align
        name_cell.border = thin_border

        # Kích thước (E:F)
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
        size_cell = ws.cell(row=row, column=5, value=item.size or "")
        size_cell.font = font
        size_cell.alignment = wrap_align
        size_cell.border = thin_border

        # Mã hàng
        code_cell = ws.cell(row=row, column=7, value=item.item_code)
        code_cell.font = font
        code_cell.alignment = center_align
        code_cell.border = thin_border

        # SL
        qty_cell = ws.cell(row=row, column=8, value=item.qty)
        qty_cell.font = font
        qty_cell.alignment = center_align
        qty_cell.border = thin_border

        # Hình ảnh (I:J)
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
        img_cell = ws.cell(row=row, column=9)
        img_cell.border = thin_border

        # Đơn giá
        rate_cell = ws.cell(row=row, column=12, value=item.rate or 0)
        rate_cell.font = font
        rate_cell.alignment = center_align
        rate_cell.number_format = currency_format
        rate_cell.border = thin_border

        # Thành tiền
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

    # Remove template and set title
    wb.remove(template_ws)
    ws.title = template_ws.title

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
