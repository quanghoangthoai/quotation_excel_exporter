import frappe
import io
from openpyxl import load_workbook

@frappe.whitelist()
def export_excel_api(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    # Đường dẫn đến file mẫu
    file_path = frappe.get_site_path("public", "files", "mẫu báo giá.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active

    # Ghi tên khách hàng
    ws["B9"] = customer.customer_name or ""

    # Lấy số điện thoại từ Contact
    contact_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Contact"
    }, "parent")

    contact_mobile = ""
    if contact_name:
        contact = frappe.get_doc("Contact", contact_name)
        contact_mobile = contact.mobile_no or contact.phone or ""

    ws["I9"] = contact_mobile

    # Lấy địa chỉ (address_display)
    address_name = frappe.db.get_value("Dynamic Link", {
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, "parent")

    address_display = ""
    if address_name:
        address = frappe.get_doc("Address", address_name)
        address_display = address.get("address_display") or ""

    ws["B10"] = address_display

    # Ghi sản phẩm vào Excel
    start_row = 14
    for i, item in enumerate(quotation.items):
        row = start_row + i
        ws[f"A{row}"] = i + 1
        ws[f"B{row}"] = item.item_name
        ws[f"E{row}"] = item.description or ""
        ws[f"G{row}"] = item.item_code
        ws[f"H{row}"] = item.qty

    # Tổng cộng
    ws["C17"] = quotation.total or 0
    ws["C18"] = 0
    ws["C19"] = 0
    ws["C20"] = quotation.total or 0

    # Xuất file về browser
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
