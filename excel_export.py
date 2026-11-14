from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


HEADER_FILL = PatternFill("solid", fgColor="17324D")
HEADER_FONT = Font(name="Tahoma", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Tahoma", size=10)
MONEY_FORMAT = '#,##0.00 "ج.م"'


def export_invoices(db, output_path, start_date, end_date):
    sales = db.sales_between(start_date.isoformat(), end_date.isoformat())
    workbook = Workbook()
    invoices_sheet = workbook.active
    invoices_sheet.title = "الفواتير"
    invoices_sheet.sheet_view.rightToLeft = True
    invoices_sheet.append(
        ["رقم الفاتورة", "التاريخ", "العميل", "الهاتف", "الإجمالي", "المدفوع", "المتبقي", "الربح"]
    )

    items_sheet = workbook.create_sheet("الأصناف")
    items_sheet.sheet_view.rightToLeft = True
    items_sheet.append(
        ["رقم الفاتورة", "التاريخ", "العميل", "المنتج", "الكمية", "الوحدة", "سعر الوحدة", "إجمالي الصنف"]
    )

    payments_sheet = workbook.create_sheet("الدفعات")
    payments_sheet.sheet_view.rightToLeft = True
    payments_sheet.append(["رقم الفاتورة", "العميل", "تاريخ التحصيل", "المبلغ", "البيان"])

    for sale in sales:
        invoices_sheet.append(
            [
                sale["invoice_number"],
                sale["created_at"],
                sale["customer_name"],
                sale["customer_phone"],
                sale["total"],
                sale["paid"],
                sale["due"],
                sale["total"] - sale["cost_total"],
            ]
        )
        _sale, items = db.sale(sale["id"])
        for item in items:
            items_sheet.append(
                [
                    sale["invoice_number"],
                    sale["created_at"],
                    sale["customer_name"],
                    item["product_name"],
                    item["quantity"],
                    item["unit"],
                    item["unit_price"],
                    item["quantity"] * item["unit_price"],
                ]
            )
        for payment in db.sale_payments(sale["id"]):
            payments_sheet.append(
                [
                    sale["invoice_number"],
                    sale["customer_name"],
                    payment["payment_date"],
                    payment["amount"],
                    payment["note"],
                ]
            )

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="right")
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = BODY_FONT
                cell.alignment = Alignment(horizontal="right")
        for column_cells in sheet.columns:
            column_letter = column_cells[0].column_letter
            max_length = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_letter].width = min(max(max_length + 3, 14), 36)

    for row in invoices_sheet.iter_rows(min_row=2, min_col=5, max_col=8):
        for cell in row:
            cell.number_format = MONEY_FORMAT
    for row in items_sheet.iter_rows(min_row=2, min_col=7, max_col=8):
        for cell in row:
            cell.number_format = MONEY_FORMAT
    for row in payments_sheet.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].number_format = MONEY_FORMAT

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return {"invoice_count": len(sales), "path": output_path}
