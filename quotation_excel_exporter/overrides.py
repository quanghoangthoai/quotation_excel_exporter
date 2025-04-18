# apps/quotation_excel_exporter/quotation_excel_exporter/overrides.py

from erpnext.selling.doctype.sales_order.sales_order import SalesOrder as _SalesOrder

def disable_commission_validation(doc, method):
    # Override method validate_sales_team trên instance để no-op
    doc.validate_sales_team = lambda: None

class SalesOrder(_SalesOrder):
    def validate_sales_team(self):
        # Ghi đè hoàn toàn ở class level, no-op
        return
