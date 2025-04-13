import frappe
import io
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

@frappe.whitelist()
def export_excel_api(quotation_name):
    # Lấy thông tin Quotation và Customer
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    # Tạo workbook mới
    wb = Workbook()
    ws = wb.active
    ws.title = "Báo giá"

    # Định nghĩa font và border
    font_13 = Font(name="Times New Roman", size=13)
    font_13_bold = Font(name="Times New Roman", size=13, bold=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')

    # Thêm logo
    try:
        logo_path = frappe.get_site_path("public", "files", "logo.jpg")  # Đường dẫn đến logo
        if os.path.exists(logo_path):
            img = XLImage(logo_path)
            img.width = 100  # Điều chỉnh kích thước logo
            img.height = 100
            ws.add_image(img, "A1")  # Đặt logo tại ô A1
            ws.row_dimensions[1].height = 80  # Điều chỉnh chiều cao hàng để vừa logo
        else:
            frappe.log_error(f"Logo not found at: {logo_path}")
    except Exception as e:
        frappe.log_error(f"Error adding logo: {str(e)}")

    # Điền thông tin header (dịch xuống dưới để không đè lên logo)
    ws.merge_cells('A3:N3')
    ws['A3'] = "CÔNG TY PHÁT TRIỂN THƯƠNG MẠI THẾ KỶ"
    ws['A3'].font = font_13_bold
    ws['A3'].alignment = center_alignment

    ws.merge_cells('A4:N4')
    ws['A4'] = "Địa chỉ: Số 30 đường 16, KĐT Đông Tăng Long, TP Thủ Đức, HCM"
    ws['A4'].font = font_13
    ws['A4'].alignment = center_alignment

    ws.merge_cells('A5:N5')
    ws['A5'] = "Hotline: 0768.927.526 - 033.566.9526"
    ws['A5'].font = font_13
    ws['A5'].alignment = center_alignment

    ws.merge_cells('A6:N6')
    ws['A6'] = "https://thehome.com.vn/"
    ws['A6'].font = font_13
    ws['A6'].alignment = center_alignment

    ws.merge_cells('A7:N7')
    ws['A7'] = "PHIẾU BÁO GIÁ BÁN HÀNG"
    ws['A7'].font = font_13_bold
    ws['A7'].alignment = center_alignment

    # Điền thông tin khách hàng
    ws['A9'] = "Khách hàng:"
    ws['A9'].font = font_13
    ws['B9'] = customer.customer_name or "N/A"
    ws['B9'].font = font_13

    # Lấy thông tin liên hệ (Contact)
    contact_mobile = "N/A"
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        contact_mobile = contact.mobile_no or contact.phone or "N/A"

    ws['I9'] = "Điện thoại:"
    ws['I9'].font = font_13
    ws['J9'] = contact_mobile
    ws['J9'].font = font_13
    ws['J9'].alignment = left_alignment

    # Lấy thông tin địa chỉ (Address)
    address_line = "N/A"
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")
    if address_name:
        address = frappe.get_doc("Address", address_name)
        address_line = address.address_line1 or "N/A"

    ws['A10'] = "Địa chỉ:"
    ws['A10'].font = font_13
    ws['B10'] = address_line
    ws['B10'].font = font_13

    ws['A11'] = "Lời đầu tiên, xin cảm ơn Quý khách hàng đã quan tâm đến sản phẩm nội thất của công ty chúng tôi."
    ws['A11'].font = font_13
    ws.merge_cells('A11:N11')

    ws['A12'] = "Chúng tôi xin gửi đến Quý khách hàng Bảng báo giá như sau:"
    ws['A12'].font = font_13
    ws.merge_cells('A12:N12')

    # Điền tiêu đề bảng
    headers = ["STT", "Tên sản phẩm", "", "", "Kích thước sản phẩm", "Mã hàng", "", "SL", "Hình ảnh", "", "Đơn vị", "Đơn giá", "CK (%)", "Thành tiền"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=14, column=col, value=header).font = font_13_bold
        ws.cell(row=14, column=col).border = thin_border
        ws.cell(row=14, column=col).alignment = center_alignment

    # Merge các cột tiêu đề
    ws.merge_cells('B14:D14')  # Tên sản phẩm
    ws.merge_cells('F14:G14')  # Mã hàng
    ws.merge_cells('I14:J14')  # Hình ảnh

    # Kiểm tra dữ liệu trong quotation.items
    frappe.log_error(f"Total items: {len(quotation.items)}")
    for i, item in enumerate(quotation.items):
        frappe.log_error(f"Item {i+1}:")
        frappe.log_error(f"  item_name: {item.item_name}")
        frappe.log_error(f"  size: {item.size}")
        frappe.log_error(f"  item_code: {item.item_code}")
        frappe.log_error(f"  qty: {item.qty}")
        frappe.log_error(f"  rate: {item.rate}")
        frappe.log_error(f"  discount_percentage: {item.discount_percentage}")
        frappe.log_error(f"  amount: {item.amount}")

    # Điền dữ liệu cho từng item
    for i, item in enumerate(quotation.items):
        row = 16 + i

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
    total_row = 16 + len(quotation.items)

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

    # Điền footer
    footer_row = total_row + 5
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=3)
    ws.cell(row=footer_row, column=1, value="Khách hàng").font = font_13
    ws.cell(row=footer_row, column=1).alignment = center_alignment

    ws.merge_cells(start_row=footer_row, start_column=5, end_row=footer_row, end_column=7)
    ws.cell(row=footer_row, column=5, value="Người giao hàng").font = font_13
    ws.cell(row=footer_row, column=5).alignment = center_alignment

    ws.merge_cells(start_row=footer_row, start_column=9, end_row=footer_row, end_column=14)
    ws.cell(row=footer_row, column=9, value="Ngày     Tháng     Năm").font = font_13
    ws.cell(row=footer_row, column=9).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=1, end_row=footer_row + 1, end_column=3)
    ws.cell(row=footer_row + 1, column=1, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=1).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=5, end_row=footer_row + 1, end_column=7)
    ws.cell(row=footer_row + 1, column=5, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=5).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 1, start_column=9, end_row=footer_row + 1, end_column=14)
    ws.cell(row=footer_row + 1, column=9, value="(Ký và ghi rõ họ tên)").font = font_13
    ws.cell(row=footer_row + 1, column=9).alignment = center_alignment

    ws.merge_cells(start_row=footer_row + 3, start_column=1, end_row=footer_row + 3, end_column=14)
    ws.cell(row=footer_row + 3, column=1, value="Lưu ý: Không đổi trả sản phẩm mẫu trừ trường hợp sản phẩm bị lỗi từ nhà sản xuất").font = font_13

    ws.merge_cells(start_row=footer_row + 4, start_column=1, end_row=footer_row + 4, end_column=14)
    ws.cell(row=footer_row + 4, column=1, value="Hình thức thanh toán:").font = font_13

    ws.merge_cells(start_row=footer_row + 5, start_column=1, end_row=footer_row + 5, end_column=14)
    ws.cell(row=footer_row + 5, column=1, value="- Thanh toán 100% giá trị đơn hàng khi nhận được hàng").font = font_13

    ws.merge_cells(start_row=footer_row + 6, start_column=1, end_row=footer_row + 6, end_column=14)
    ws.cell(row=footer_row + 6, column=1, value="- Đặt hàng đặt cọc trước 30% giá trị đơn hàng").font = font_13

    # Xuất file Excel
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
