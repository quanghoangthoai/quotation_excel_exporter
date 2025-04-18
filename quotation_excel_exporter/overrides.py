# apps/quotation_excel_exporter/quotation_excel_exporter/overrides.py

from erpnext.selling.doctype.sales_order.sales_order import SalesOrder as _SalesOrder

def disable_commission_validation(doc, method=None, *args, **kwargs):
    """
    Called via doc_events before_validate.
    We override the instance method so that validate_sales_team() is a no-op.
    """
    doc.validate_sales_team = lambda: None

class SalesOrder(_SalesOrder):
    def validate_sales_team(self):
        # completely skip the 100% check
        return
