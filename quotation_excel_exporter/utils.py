import frappe
import io
import os
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

@frappe.whitelist()
def export_excel_api(quotation_name):
    # Lấy thông tin Quotation và Customer
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    # Load file template Excel
    file_path = frappe.get_site_path("public", "files", "mẫu báo giá final.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active

    # Định nghĩa font và border
    font_13 = Font(name="Times New Roman", size=13)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')

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
        ws.cell(row=9, column=10).alignment = left_alignment

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

    # Xóa dữ liệu mẫu (nếu có) ở row 14
    for col in range(1, 15):
        ws.cell(row=14, column=col).value = None

    # Chèn thêm hàng nếu có nhiều hơn 1 item
    num_items = len(quotation.items)
    if num_items > 1:
        ws.insert_rows(15, num_items - 1)

    # Điền dữ liệu cho từng item
    for i, item in enumerate(quotation.items):
        row = 14 + i

        # Định dạng và điền dữ liệu cho từng ô
        # STT
        ws.cell(row=row, column=1, value=i + 1).font = font_13
        ws.cell(row=row, column=1).border = thin_border
        ws.cell(row=row, column=1).alignment = center_alignment

        # Tên sản phẩm (merge cột 2-4)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.cell(row=row, column=2, value=item.item_name or "N/A").font = font_13
        ws.cell(row=row, column=2).border = thin_border
        ws.cell(row=row, column=2).alignment = left_alignment

        # Kích thước
        ws.cell(row=row, column=5, value=item.size or "N/A").font = font_13
        ws.cell(row=row, column=5).border = thin_border
        ws.cell(row=row, column=5).alignment = center_alignment

        # Mã hàng (merge cột 6-7)
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=7)
        ws.cell(row=row, column=6, value=item.item_code or "N/A").font = font_13
        ws.cell(row=row, column=6).border = thin_border
        ws.cell(row=row, column=6).alignment = center_alignment

        # Số lượng (SL)
        ws.cell(row=row, column=8, value=item.qty or 0).font = font_13
        ws.cell(row=row, column=8).border = thin_border
        ws.cell(row=row, column=8).alignment = center_alignment

        # Hình ảnh (merge cột 9-10)
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
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
                    ws.add_image(img, f"I{row}")
                    ws.row_dimensions[row].height = 80
                else:
                    ws.row_dimensions[row].height = 30
            except Exception as e:
                frappe.log_error(f"Hình ảnh lỗi: {item.image} - {str(e)}")
                ws.row_dimensions[row].height = 30
        else:
            ws.row_dimensions[row].height = 30

        # Đơn vị
        ws.cell(row=row, column=11, value="Bộ").font = font_13
        ws.cell(row=row, column=11).border = thin_border
        ws.cell(row=row, column=11).alignment = center_alignment

        # Đơn giá
        ws.cell(row=row, column=12, value=item.rate or 0).font = font_13
        ws.cell(row=row, column=12).border = thin_border
        ws.cell(row=row, column=12).alignment = center_alignment
        ws.cell(row=row, column=12).number_format = '#,##0'

        # Chiết khấu (CK)
        ws.cell(row=row, column=13, value=item.discount_percentage or 0).font = font_13
        ws.cell(row=row, column=13).border = thin_border
        ws.cell(row=row, column=13).alignment = center_alignment
        ws.cell(row=row, column=13).number_format = '0.00%'

        # Thành tiền
        ws.cell(row=row, column=14, value=item.amount or (item.qty * item.rate if item.qty and item.rate else 0)).font = font_13
        ws.cell(row=row, column=14).border = thin_border
        ws.cell(row=row, column=14).alignment = center_alignment
        ws.cell(row=row, column=14).number_format = '#,##0'

    # Điền thông tin tổng cộng và các mục khác
    total_row = 14 + num_items

    # Xóa merge cells cũ (nếu có) từ total_row trở đi
    for i in range(4):
        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_row == total_row + i:
                ws.unmerge_cells(str(merged_range))

    # Tổng cộng (A)
    ws.cell(row=total_row, column=1, value="A").font = font_13
    ws.cell(row=total_row, column=1).border = thin_border
    ws.cell(row=total_row, column=2, value="Tổng cộng").font = font_13
    ws.cell(row=total_row, column=2).border = thin_border
    ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=13)
    ws.cell(row=total_row, column=14, value=quotation.total or 0).font = font_13
    ws.cell(row=total_row, column=14).border = thin_border
    ws.cell(row=total_row, column=14).number_format = '#,##0'

    # Phụ phí (B)
    ws.cell(row=total_row + 1, column=1, value="B").font = font_13
    ws.cell(row=total_row + 1, column=1).border = thin_border
    ws.cell(row=total_row + 1, column=2, value="Phụ phí").font = font_13
    ws.cell(row=total_row + 1, column=2).border = thin_border
    ws.merge_cells(start_row=total_row + 1, start_column=2, end_row=total_row + 1, end_column=13)
    ws.cell(row=total_row + 1, column=14, value=0).font = font_13
    ws.cell(row=total_row + 1, column=14).border = thin_border
    ws.cell(row=total_row + 1, column=14).number_format = '#,##0'

    # Đã thanh toán (C)
    ws.cell(row=total_row + 2, column=1, value="C").font = font_13
    ws.cell(row=total_row + 2, column=1).border = thin_border
    ws.cell(row=total_row + 2, column=2, value="Đã thanh toán").font = font_13
    ws.cell(row=total_row + 2, column=2).border = thin_border
    ws.merge_cells(start_row=total_row + 2, start_column=2, end_row=total_row + 2, end_column=13)
    ws.cell(row=total_row + 2, column=14, value=0).font = font_13
    ws.cell(row=total_row + 2, column=14).border = thin_border
    ws.cell(row=total_row + 2, column=14).number_format = '#,##0'

    # Tổng tiền thanh toán (A+B-C)
    ws.cell(row=total_row + 3, column=1, value="Tổng tiền thanh toán (A+B-C)").font = font_13
    ws.cell(row=total_row + 3, column=1).border = thin_border
    ws.merge_cells(start_row=total_row + 3, start_column=1, end_row=total_row + 3, end_column=13)
    ws.cell(row=total_row + 3, column=14, value=quotation.total or 0).font = font_13
    ws.cell(row=total_row + 3, column=14).border = thin_border
    ws.cell(row=total_row + 3, column=14).number_format = '#,##0'

    # Xuất file Excel
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
