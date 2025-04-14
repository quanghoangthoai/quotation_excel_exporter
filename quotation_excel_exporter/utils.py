import frappe
import io
import os
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from datetime import datetime
import tempfile

@frappe.whitelist()
def export_excel_api(quotation_name):
    # Validate input
    if not quotation_name or not frappe.db.exists("Quotation", quotation_name):
        frappe.throw("Invalid or non-existent Quotation")

    # Fetch Quotation and Customer
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name) if quotation.party_name else None

    # Initialize Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Báo giá"

    # Define Styles
    font_13 = Font(name="Times New Roman", size=13)
    font_13_bold = Font(name="Times New Roman", size=13, bold=True)
    font_16 = Font(name="Times New Roman", size=16, bold=True)
    font_18 = Font(name="Times New Roman", size=18, bold=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Logo (Optional)
    logo_path = frappe.get_site_path("public", "files", "z6473642459612_58e86d169bb72c78b360392b4f81e8bae2152f.jpg")
    if os.path.exists(logo_path):
        try:
            logo_img = XLImage(logo_path)
            logo_img.width = 180
            logo_img.height = 60
            ws.add_image(logo_img, "A1")
            ws.merge_cells("A1:B3")
            ws.row_dimensions[1].height = 60
        except Exception as e:
            frappe.log_error(f"Failed to add logo: {str(e)}")

    # Company Details
    ws.merge_cells("D2:H3")
    ws["D2"] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws["D2"].font = font_18
    ws["D2"].alignment = center_alignment

    ws.merge_cells("A5:B5")
    ws["A5"] = "Địa chỉ:"
    ws.merge_cells("C5:N5")
    ws["C5"] = "Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức, HCM"
    
    ws.merge_cells("A6:B6")
    ws["A6"] = "Hotline:"
    ws.merge_cells("C6:N6")
    ws["C6"] = "0768.927.526 - 033.566.9526"
    
    ws.merge_cells("A7:B7")
    ws["A7"] = "Website:"
    ws.merge_cells("C7:N7")
    ws["C7"] = "https://thehome.com.vn/"

    for cell in ["A5", "A6", "A7", "C5", "C6", "C7"]:
        ws[cell].font = font_13
        ws[cell].alignment = left_alignment

    # Title
    ws.merge_cells("A9:N9")
    ws["A9"] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws["A9"].font = font_16
    ws["A9"].alignment = center_alignment

    # Introduction
    ws.merge_cells("A11:N11")
    ws["A11"] = "Lời đầu tiên, xin cảm ơn Quý khách hàng đã quan tâm đến sản phẩm nội thất của công ty chúng tôi."
    ws["A11"].font = font_13
    ws["A11"].alignment = left_alignment

    ws.merge_cells("A12:N12")
    ws["A12"] = "Chúng tôi xin gửi đến Quý khách hàng Bảng báo giá như sau:"
    ws["A12"].font = font_13
    ws["A12"].alignment = left_alignment

    # Customer Details (Bold "Khách hàng")
    ws["A13"] = "Khách hàng:"
    ws["A13"].font = font_13_bold
    ws.merge_cells("B13:H13")
    ws["B13"] = customer.customer_name if customer else ""
    
    ws["I13"] = "Điện thoại:"
    ws.merge_cells("J13:N13")
    phone = ""
    contact_name = frappe.db.get_value(
        "Dynamic Link",
        {"link_doctype": "Customer", "link_name": quotation.party_name, "parenttype": "Contact"},
        "parent"
    )
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        phone = contact.mobile_no or contact.phone or ""
    ws["J13"] = phone

    ws["A14"] = "Địa chỉ:"
    ws.merge_cells("B14:N14")
    address = ""
    address_name = frappe.db.get_value(
        "Dynamic Link",
        {"link_doctype": "Customer", "link_name": quotation.party_name, "parenttype": "Address"},
        "parent"
    )
    if address_name:
        addr = frappe.get_doc("Address", address_name)
        address = addr.address_line1 or ""
        if addr.address_line2:
            address += ", " + addr.address_line2
        if addr.city:
            address += ", " + addr.city
        if addr.country:
            address += ", " + addr.country
    ws["B14"] = address

    for cell in ["A13", "B13", "I13", "J13", "A14", "B14"]:
        ws[cell].font = font_13 if cell != "A13" else font_13_bold
        ws[cell].alignment = left_alignment

    # Table Headers (Bolded)
    headers = [
        "STT", "Tên sản phẩm", "", "", "Kích thước sản phẩm", "", "Mã hàng", 
        "SL", "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"
    ]
    row_num = 16  # Adjusted to leave space
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col)
        cell.value = header
        cell.font = font_13_bold
        cell.alignment = center_alignment
        cell.border = thin_border

    ws.merge_cells(f"B{row_num}:D{row_num}")
    ws.merge_cells(f"E{row_num}:F{row_num}")
    ws.merge_cells(f"I{row_num}:J{row_num}")

    # Table Data
    temp_files = []
    for i, item in enumerate(quotation.items, 1):
        row = row_num + i
        ws.cell(row=row, column=1, value=i).font = font_13
        ws.merge_cells(f"B{row}:D{row}")
        ws.cell(row=row, column=2, value=item.item_name or "").font = font_13
        ws.merge_cells(f"E{row}:F{row}")
        ws.cell(row=row, column=5, value=frappe.db.get_value("Quotation Item", item.name, "size") or "").font = font_13
        ws.cell(row=row, column=7, value=item.item_code or "").font = font_13
        ws.cell(row=row, column=8, value=item.qty or 0).font = font_13
        ws.merge_cells(f"I{row}:J{row}")
        ws.cell(row=row, column=11, value=item.uom or "Bộ").font = font_13
        ws.cell(row=row, column=12, value=item.rate or 0).font = font_13
        ws.cell(row=row, column=13, value=item.discount_percentage or 0).font = font_13
        ws.cell(row=row, column=14, value=item.amount or 0).font = font_13

        # Image Handling with Centering
        if item.image:
            try:
                image_path = None
                if item.image.startswith("/files/"):
                    image_path = frappe.get_site_path("public", item.image.lstrip("/"))
                elif item.image.startswith("http"):
                    response = requests.get(item.image, timeout=5)
                    if response.status_code == 200:
                        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                        tmp_file.write(response.content)
                        tmp_file.close()
                        image_path = tmp_file.name
                        temp_files.append(image_path)
                if image_path and os.path.exists(image_path):
                    img = XLImage(image_path)
                    img.width = 80
                    img.height = 80
                    ws.add_image(img, f"I{row}")
                    ws.row_dimensions[row].height = 80
                    img.anchor = f"I{row}"
            except Exception as e:
                frappe.log_error(f"Failed to add image for item {item.item_code}: {str(e)}")

        for col in range(1, 15):
            ws.cell(row=row, column=col).border = thin_border
            ws.cell(row=row, column=col).alignment = center_alignment

    # Totals (Formatted as a Table)
    current_row = row_num + len(quotation.items) + 1
    additional_fees = 0
    advance_payment = 0

    ws.cell(row=current_row, column=1, value="A").font = font_13
    ws.cell(row=current_row, column=1).border = thin_border
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Tổng cộng").font = font_13
    for col in range(2, 14):
        ws.cell(row=current_row, column=col).border = thin_border
    ws.cell(row=current_row, column=14, value=quotation.grand_total).font = font_13
    ws.cell(row=current_row, column=14).border = thin_border

    current_row += 1
    ws.cell(row=current_row, column=1, value="B").font = font_13
    ws.cell(row=current_row, column=1).border = thin_border
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Phụ phí").font = font_13
    for col in range(2, 14):
        ws.cell(row=current_row, column=col).border = thin_border
    ws.cell(row=current_row, column=14, value="").font = font_13
    ws.cell(row=current_row, column=14).border = thin_border

    current_row += 1
    ws.cell(row=current_row, column=1, value="C").font = font_13
    ws.cell(row=current_row, column=1).border = thin_border
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=2, value="Đã thanh toán").font = font_13
    for col in range(2, 14):
        ws.cell(row=current_row, column=col).border = thin_border
    ws.cell(row=current_row, column=14, value="0").font = font_13
    ws.cell(row=current_row, column=14).border = thin_border

    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=13)
    ws.cell(row=current_row, column=1, value="Tổng tiền thanh toán (A+B-C)").font = font_13
    for col in range(1, 14):
        ws.cell(row=current_row, column=col).border = thin_border
    ws.cell(row=current_row, column=14, value=quotation.grand_total).font = font_13
    ws.cell(row=current_row, column=14).border = thin_border

    for r in range(current_row - 3, current_row + 1):
        ws.cell(row=r, column=1).alignment = left_alignment if r == current_row else center_alignment
        ws.cell(row=r, column=2).alignment = left_alignment
        ws.cell(row=r, column=14).alignment = center_alignment

    # Footer (Bold Titles)
    footer_row = current_row + 2
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=5)
    ws.cell(row=footer_row, column=1, value="Khách hàng").font = font_13_bold
    ws.cell(row=footer_row, column=1).alignment = center_alignment

    ws.merge_cells(start_row=footer_row, start_column=6, end_row=footer_row, end_column=10)
    ws.cell(row=footer_row, column=6, value="Người giao hàng").font = font_13_bold
    ws.cell(row=footer_row, column=6).alignment = center_alignment

    ws.merge_cells(start_row=footer_row, start_column=11, end_row=footer_row, end_column=14)
    date = quotation.transaction_date or datetime.now()
    ws.cell(row=footer_row, column=11, value=f"Ngày {date.day:02d} Tháng {date.month:02d} Năm {date.year}").font = font_13_bold
    ws.cell(row=footer_row, column=11).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=1, end_row=footer_row + 1, end_column=5)
    ws.cell(row=footer_row + 1, column=1, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=1).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=6, end_row=footer_row + 1, end_column=10)
    ws.cell(row=footer_row + 1, column=6, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=6).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=11, end_row=footer_row + 1, end_column=14)
    ws.cell(row=footer_row + 1, column=11, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=11).alignment = center_alignment

    # Notes (Add 3-Row Gap, Bold Titles and Content)
    notes_row = footer_row + 1 + 3  # 3-row gap
    ws.merge_cells(start_row=notes_row, start_column=1, end_row=notes_row, end_column=14)
    ws.cell(row=notes_row, column=1, value="Lưu ý:").font = font_13_bold
    ws.cell(row=notes_row, column=1).alignment = left_alignment

    ws.merge_cells(start_row=notes_row + 1, start_column=1, end_row=notes_row + 1, end_column=14)
    ws.cell(row=notes_row + 1, column=1, value="Không đổi trả sản phẩm mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất").font = font_13_bold
    ws.cell(row=notes_row + 1, column=1).alignment = left_alignment

    ws.merge_cells(start_row=notes_row + 2, start_column=1, end_row=notes_row + 2, end_column=14)
    ws.cell(row=notes_row + 2, column=1, value="Hình thức thanh toán:").font = font_13_bold
    ws.cell(row=notes_row + 2, column=1).alignment = left_alignment

    ws.merge_cells(start_row=notes_row + 3, start_column=1, end_row=notes_row + 3, end_column=14)
    ws.cell(row=notes_row + 3, column=1, value="- Thanh toán 100% giá trị đơn hàng khi nhận được hàng").font = font_13
    ws.cell(row=notes_row + 3, column=1).alignment = left_alignment

    ws.merge_cells(start_row=notes_row + 4, start_column=1, end_row=notes_row + 4, end_column=14)
    ws.cell(row=notes_row + 4, column=1, value="- Đặt hàng đặt cọc trước 30% giá trị đơn hàng").font = font_13
    ws.cell(row=notes_row + 4, column=1).alignment = left_alignment

    # Set Column Widths
    column_widths = {
        "B": 20, "C": 5, "D": 5, "E": 15, "F": 5, "G": 10,
        "H": 5, "I": 15, "J": 5, "K": 8, "L": 12, "M": 8, "N": 12
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Save and Output
    try:
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
        frappe.local.response.filecontent = output.read()
        frappe.local.response.type = "binary"
    except Exception as e:
        frappe.throw(f"Failed to generate Excel file: {str(e)}")
    finally:
        # Clean up temporary files
        for temp_file in temp_files:
            try:
                os.unlink(temp_file)
            except Exception:
                pass
