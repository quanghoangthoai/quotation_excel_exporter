from __future__ import unicode_literals

app_name = "quotation_excel_exporter"
app_title = "Quotation Excel Exporter"
app_publisher = "Your Name"
app_description = "Export quotations to Excel using template"
app_email = "you@example.com"
app_license = "MIT"

override_doctype_class = {
    "Sales Order": "quotation_excel_exporter.overrides.SalesOrder",
}

doc_events = {
    "Sales Order": {
        "before_validate": "quotation_excel_exporter.overrides.disable_commission_validation"
    }
}

