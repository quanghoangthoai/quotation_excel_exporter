import frappe
import io
import os
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, PatternFill
from copy import copy

@frappe.whitelist()
def export_excel_api(quotation_name):
    # Lấy thông tin Quotation và Customer
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    # Load file template Excel
    file_path = frappe.get_site_path("public", "files", "mẫu báo giá final.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active

    # Định nghĩa font
    font_13 = Font(name="Times New Roman", size=13)

    # Điền thông tin khách hàng
    ws["B9"] = customer.customer_name or "N/A"
    ws["B9"].font = font_13

    # Lấy thông tin liên hệ (Contact)
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")

    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        contact_mobile = contact.mobile_no or contact.phone or "N/A"
        ws.cell(row=9, column=10, value=contact_mobile).font = font_13
        ws.cell(row=9, column=10).alignment = Alignment(horizontal="left", vertical="center")

    # Lấy thông tin địa chỉ (Address)
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")

    if address_name:
        address = frappe.get_doc("Address", address_name)
        ws["B10"] = address.address_line1 or "N/A"
        ws["B10"].font = font_13

    # Lưu định dạng của hàng template (row 14)
    template_row = 14
    template_styles = {}
    
    # Lưu chiều cao của hàng template
    template_row_height = ws.row_dimensions[template_row].height if ws.row_dimensions[template_row].height else 20
    
    # Lưu thông tin merge cells của hàng template
    template_merges = []
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row == template_row:
            template_merges.append({
                'min_col': merged_range.min_col,
                'max_col': merged_range.max_col,
                'min_row': merged_range.min_row,
                'max_row': merged_range.max_row
            })

    # Lưu định dạng của hàng template
    for col in range(1, 15):
        cell = ws.cell(row=template_row, column=col)
        template_styles[col] = {
            'font': copy(cell.font) if cell.font else Font(name="Times New Roman", size=13),
            'border': copy(cell.border) if cell.border else Border(),
            'fill': copy(cell.fill) if cell.fill else PatternFill(),
            'number_format': cell.number_format if cell.number_format else 'General',
            'alignment': copy(cell.alignment) if cell.alignment else Alignment(horizontal='left', vertical='center')
        }

    # Chèn thêm hàng nếu có nhiều hơn 1 item
    num_items = len(quotation.items)
    if num_items > 1:
        ws.insert_rows(template_row + 1, num_items - 1)

    # Điền dữ liệu cho từng item
    for i, item in enumerate(quotation.items):
        row = template_row + i
        
        # Áp dụng định dạng template cho toàn bộ hàng
        for col in range(1, 15):
            target_cell = ws.cell(row=row, column=col)
            target_cell.font = copy(template_styles[col]['font'])
            target_cell.border = copy(template_styles[col]['border'])
            target_cell.fill = copy(template_styles[col]['fill'])
            target_cell.alignment = copy(template_styles[col]['alignment'])
            target_cell.number_format = template_styles[col]['number_format']

        # Điền dữ liệu với giá trị mặc định nếu thiếu
        ws.cell(row=row, column=1, value=i + 1)  # STT
        ws.cell(row=row, column=2, value=item.item_name or "N/A")  # Tên sản phẩm
        ws.cell(row=row, column=5, value=item.size or "N/A")  # Kích thước
        ws.cell(row=row, column=7, value=item.item_code or "N/A")  # Mã hàng
        ws.cell(row=row, column=8, value=item.qty or 0)  # SL
        ws.cell(row=row, column=11, value="Bộ")  # Đơn vị
        ws.cell(row=row, column=12, value=item.rate or 0)  # Đơn giá
        ws.cell(row=row, column=13, value=item.discount_percentage or 0)  # CK
        ws.cell(row=row, column=14, value=item.amount or (item.qty * item.rate if item.qty and item.rate else 0))  # Thành tiền

        # Áp dụng merge cells cho hàng này
        for merge in template_merges:
            ws.merge_cells(
                start_row=row,
                start_column=merge['min_col'],
                end_row=row,
                end_column=merge['max_col']
            )

        # Xử lý hình ảnh
        if item.image:
            try:
                image_path = ""
                if item.image.startswith("/files/"):
                    image_path = frappe.get_site_path("public", item.image.lstrip("/"))
                elif item.image.startswith("http"):
                    response = requests.get(item.image, timeout=5)
                    if response.status_code == 200:
                        tmp_path = f"/tmp/tmp_item_{i}.png"
                        with open(tmp_path, "wb") as f:
                            f.write(response.content)
                        image_path = tmp_path

                if os.path.exists(image_path):
                    img = XLImage(image_path)
                    img.width = 100
                    img.height = 100
                    ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
                    ws.add_image(img, f"I{row}")
                    ws.row_dimensions[row].height = 80
                else:
                    ws.row_dimensions[row].height = template_row_height
            except Exception as e:
                frappe.log_error(f"Hình ảnh lỗi: {item.image} - {str(e)}")
                ws.row_dimensions[row].height = template_row_height
        else:
            ws.row_dimensions[row].height = template_row_height

    # Điền thông tin tổng cộng và các mục khác
    total_row = template_row + num_items

    for i in range(4):
        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_row == total_row + i:
                ws.unmerge_cells(str(merged_range))

    ws.cell(row=total_row, column=1, value="A").font = font_13
    ws.cell(row=total_row, column=2, value="Tổng cộng").font = font_13
    ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=13)
    ws.cell(row=total_row, column=14, value=quotation.total or 0).font = font_13

    ws.cell(row=total_row + 1, column=1, value="B").font = font_13
    ws.cell(row=total_row + 1, column=2, value="Phụ phí").font = font_13
    ws.merge_cells(start_row=total_row + 1, start_column=2, end_row=total_row + 1, end_column=13)
    ws.cell(row=total_row + 1, column=14, value=0).font = font_13

    ws.cell(row=total_row + 2, column=1, value="C").font = font_13
    ws.cell(row=total_row + 2, column=2, value="Đã thanh toán").font = font_13
    ws.merge_cells(start_row=total_row + 2, start_column=2, end_row=total_row + 2, end_column=13)
    ws.cell(row=total_row + 2, column=14, value=0).font = font_13

    ws.cell(row=total_row + 3, column=1, value="Tổng tiền thanh toán (A+B-C)").font = font_13
    ws.merge_cells(start_row=total_row + 3, start_column=1, end_row=total_row + 3, end_column=13)
    ws.cell(row=total_row + 3, column=14, value=quotation.total or 0).font = font_13

    # Xuất file Excel
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
