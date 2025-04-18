
# Override Sales Order
from erpnext.selling.doctype.sales_order.sales_order import SalesOrder as _SalesOrder

class SalesOrder(_SalesOrder):
    def validate_sales_team(self):
        # Bỏ kiểm tra tổng % hoa hồng
        return
