import frappe
import io
import os
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    wb = Workbook()
    ws = wb.active
    ws.title = "Báo giá"

    font_13 = Font(name="Times New Roman", size=13)
    font_16 = Font(name="Times New Roman", size=16, bold=True)
    font_18 = Font(name="Times New Roman", size=18, bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Logo
    logo_path = frappe.get_site_path("public", "files", "logo.jpg")
    if os.path.exists(logo_path):
        logo_img = XLImage(logo_path)
        logo_img.width = 180
        logo_img.height = 60
        ws.add_image(logo_img, "A1")
        ws.merge_cells("A1:B3")

    # Công ty
    ws.merge_cells("C1:N1")
    ws["C1"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["C1"].font = font_18
    ws["C1"].alignment = center_alignment

    ws.merge_cells("A5:B5")
    ws["A5"] = "Địa chỉ :"
    ws.merge_cells("A6:B6")
    ws["A6"] = "Hotline :"
    ws.merge_cells("A7:B7")
    ws["A7"] = "Website :"
    ws["C5"] = "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức , HCM"
    ws["C6"] = "0768.927..526 - 033.566.9526"
    ws["C7"] = "https://thehome.com.vn/"

    for cell in ["A5", "A6", "A7", "C5", "C6", "C7"]:
        ws[cell].font = font_13
        ws[cell].alignment = left_alignment

    # Tiêu đề
    ws.merge_cells("A9:N9")
    ws["A9"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["A9"].font = font_16
    ws["A9"].alignment = center_alignment

    # Lời mở đầu
    ws.merge_cells("A11:N11")
    ws["A11"] = "Lời đầu tiên , xin cảm ơn Quý khách hàng đã quan tâm đến sản phẩm nội thất của công ty chúng tôi."
    ws["A11"].font = font_13
    ws["A11"].alignment = left_alignment

    ws.merge_cells("A12:N12")
    ws["A12"] = "Chúng tôi xin gửi đến Quý khách hàng Bảng báo giá như sau :"
    ws["A12"].font = font_13
    ws["A12"].alignment = left_alignment

    # Khách hàng
    ws["A13"] = "Khách hàng :"
    ws["B13"] = customer.customer_name or ""
    ws["I13"] = "Điện thoại :"
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        ws["J13"] = contact.mobile_no or contact.phone or ""

    ws["A14"] = "Địa chỉ :"
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")
    if address_name:
        address = frappe.get_doc("Address", address_name)
        ws["B14"] = address.address_line1 or ""

    # Dòng tiêu đề bảng
    headers = ["STT", "Tên sản phẩm", "", "", "Kích thước", "", "Mã hàng", "SL", "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"]
    row_num = 13
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col)
        cell.value = header
        cell.font = font_13
        cell.alignment = center_alignment
        cell.border = thin_border
        cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

    ws.merge_cells(f"B{row_num}:D{row_num}")
    ws.merge_cells(f"E{row_num}:F{row_num}")
    ws.merge_cells(f"I{row_num}:J{row_num}")

    for i, item in enumerate(quotation.items, 1):
        row = row_num + i
        ws.cell(row=row, column=1, value=i).font = font_13
        ws.merge_cells(f"B{row}:D{row}")
        ws.cell(row=row, column=2, value=item.item_name or "").font = font_13
        ws.merge_cells(f"E{row}:F{row}")
        ws.cell(row=row, column=5, value=item.size or "").font = font_13
        ws.cell(row=row, column=7, value=item.item_code or "").font = font_13
        ws.cell(row=row, column=8, value=item.qty or 0).font = font_13
        ws.merge_cells(f"I{row}:J{row}")
        ws.cell(row=row, column=11, value="Bộ").font = font_13
        ws.cell(row=row, column=12, value=item.rate or 0).font = font_13
        ws.cell(row=row, column=13, value=item.discount_percentage or 0).font = font_13
        ws.cell(row=row, column=14, value=item.amount or (item.qty * item.rate)).font = font_13

        for col in range(1, 15):
            ws.cell(row=row, column=col).border = thin_border
            ws.cell(row=row, column=col).alignment = center_alignment

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
            except Exception:
                pass

    current_row = row_num + len(quotation.items) + 1
    ws.cell(row=current_row, column=1, value="A").font = font_13
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Tổng cộng").font = font_13
    ws.cell(row=current_row, column=14, value=quotation.total).font = font_13

    current_row += 1
    ws.cell(row=current_row, column=1, value="B").font = font_13
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Phụ phí").font = font_13
    ws.cell(row=current_row, column=14, value=0).font = font_13

    current_row += 1
    ws.cell(row=current_row, column=1, value="C").font = font_13
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Đã thanh toán").font = font_13
    ws.cell(row=current_row, column=14, value=0).font = font_13

    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=1, value="Tổng tiền thanh toán (A+B-C)").font = font_13
    ws.cell(row=current_row, column=14, value=quotation.total).font = font_13

    # Footer
    footer_row = current_row + 2
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=4)
    ws.cell(row=footer_row, column=1, value="Khách hàng").font = font_13
    ws.cell(row=footer_row, column=1).alignment = center_alignment

    ws.merge_cells(start_row=footer_row, start_column=6, end_row=footer_row, end_column=9)
    ws.cell(row=footer_row, column=6, value="Người giao hàng").font = font_13
    ws.cell(row=footer_row, column=6).alignment = center_alignment

    ws.merge_cells(start_row=footer_row, start_column=11, end_row=footer_row, end_column=14)
    ws.cell(row=footer_row, column=11, value="Ngày     Tháng     Năm").font = font_13
    ws.cell(row=footer_row, column=11).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=1, end_row=footer_row + 1, end_column=4)
    ws.cell(row=footer_row + 1, column=1, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=1).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=6, end_row=footer_row + 1, end_column=9)
    ws.cell(row=footer_row + 1, column=6, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=6).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=11, end_row=footer_row + 1, end_column=14)
    ws.cell(row=footer_row + 1, column=11, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=11).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 3, start_column=1, end_row=footer_row + 3, end_column=14)
    ws.cell(row=footer_row + 3, column=1, value="Lưu ý: Không đổi trả sản phẩm mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất").font = font_13

    ws.merge_cells(start_row=footer_row + 4, start_column=1, end_row=footer_row + 4, end_column=14)
    ws.cell(row=footer_row + 4, column=1, value="Hình thức thanh toán:").font = font_13

    ws.merge_cells(start_row=footer_row + 5, start_column=1, end_row=footer_row + 5, end_column=14)
    ws.cell(row=footer_row + 5, column=1, value="- Thanh toán 100% giá trị đơn hàng khi nhận được hàng").font = font_13

    ws.merge_cells(start_row=footer_row + 6, start_column=1, end_row=footer_row + 6, end_column=14)
    ws.cell(row=footer_row + 6, column=1, value="- Đặt hàng đặt cọc trước 30% giá trị đơn hàng").font = font_13

    # Set column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 5
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['K'].width = 8
    ws.column_dimensions['L'].width = 12
    ws.column_dimensions['M'].width = 8
    ws.column_dimensions['N'].width = 12

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
