from erpnext.selling.doctype.quotation.quotation import Quotation as _Quotation

class Quotation(_Quotation):
    def validate_sales_team(self):
        # Bỏ qua kiểm tra tổng % hoa hồng
        return

