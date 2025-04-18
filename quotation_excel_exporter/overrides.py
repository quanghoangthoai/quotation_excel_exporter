# apps/quotation_excel_exporter/quotation_excel_exporter/overrides.py

from erpnext.selling.doctype.sales_order.sales_order import SalesOrder as _SalesOrder

def disable_commission_validation(doc, method=None):
    """
    Called via doc_events before_validate.
    Overrides the validate_sales_team method to skip commission validation.
    """
    doc.validate_sales_team = lambda: None

class SalesOrder(_SalesOrder):
    def validate_sales_team(self):
        # Skip the default 100% commission rate validation
        pass
