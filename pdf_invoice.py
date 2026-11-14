import os
from pathlib import Path
from xml.sax.saxutils import escape

import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def _register_arabic_font():
    fonts_directory = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    regular_path = next(
        (fonts_directory / name for name in ("tahoma.ttf", "arial.ttf") if (fonts_directory / name).is_file()),
        None,
    )
    bold_path = next(
        (fonts_directory / name for name in ("tahomabd.ttf", "arialbd.ttf") if (fonts_directory / name).is_file()),
        None,
    )
    if not regular_path or not bold_path:
        raise FileNotFoundError("لم يتم العثور على خط عربي مناسب في مجلد خطوط Windows.")
    if "InvoiceArabic" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("InvoiceArabic", str(regular_path)))
        pdfmetrics.registerFont(TTFont("InvoiceArabic-Bold", str(bold_path)))
        pdfmetrics.registerFontFamily(
            "InvoiceArabic",
            normal="InvoiceArabic",
            bold="InvoiceArabic-Bold",
            italic="InvoiceArabic",
            boldItalic="InvoiceArabic-Bold",
        )


def _rtl(value):
    return get_display(arabic_reshaper.reshape(str(value)), base_dir="R")


def _paragraph(value, style):
    return Paragraph(escape(_rtl(value)), style)


def generate_invoice_pdf(db, sale_id, output_path):
    sale, items = db.sale(sale_id)
    if not sale:
        raise ValueError("الفاتورة غير موجودة.")

    _register_arabic_font()
    normal = ParagraphStyle(
        "InvoiceNormal",
        fontName="InvoiceArabic",
        fontSize=10,
        leading=17,
        alignment=TA_RIGHT,
        textColor=colors.HexColor("#222222"),
    )
    bold = ParagraphStyle(
        "InvoiceBold",
        parent=normal,
        fontName="InvoiceArabic-Bold",
    )
    title = ParagraphStyle(
        "InvoiceTitle",
        parent=bold,
        fontSize=19,
        leading=27,
        textColor=colors.HexColor("#17324d"),
        spaceAfter=5,
    )
    header = ParagraphStyle(
        "InvoiceTableHeader",
        parent=bold,
        fontSize=9,
        leading=13,
        textColor=colors.white,
    )

    customer = db.customer(sale["customer_id"]) if sale["customer_id"] else None
    shop_name = db.setting("shop_name", "محل الوميتال")
    elements = [
        _paragraph(shop_name, title),
        _paragraph("فاتورة بيع", bold),
        Spacer(1, 5 * mm),
        _paragraph(f"رقم الفاتورة: {sale['invoice_number']}", normal),
        _paragraph(f"التاريخ: {sale['created_at']}", normal),
        _paragraph(f"العميل: {sale['customer_name']}", normal),
    ]
    if customer and customer["phone"]:
        elements.append(_paragraph(f"الهاتف: {customer['phone']}", normal))
    elements.append(Spacer(1, 7 * mm))

    rows = [[
        _paragraph("المنتج", header),
        _paragraph("الكمية", header),
        _paragraph("سعر الوحدة", header),
        _paragraph("الإجمالي", header),
    ]]
    for item in items:
        rows.append(
            [
                _paragraph(item["product_name"], normal),
                _paragraph(f'{item["quantity"]:g} {item["unit"]}', normal),
                _paragraph(f'{item["unit_price"]:,.2f} ج.م', normal),
                _paragraph(f'{item["quantity"] * item["unit_price"]:,.2f} ج.م', normal),
            ]
        )
    item_table = Table(rows, colWidths=[74 * mm, 30 * mm, 35 * mm, 35 * mm], repeatRows=1)
    item_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17324d")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#b8c2ce")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f7fa")]),
            ]
        )
    )
    elements.extend(
        [
            item_table,
            Spacer(1, 8 * mm),
        ]
    )

    totals = [
        [_paragraph("الإجمالي:", bold), _paragraph(f'{sale["total"]:,.2f} ج.م', bold)],
        [_paragraph("المدفوع:", normal), _paragraph(f'{sale["paid"]:,.2f} ج.م', normal)],
        [_paragraph("المتبقي:", bold), _paragraph(f'{sale["due"]:,.2f} ج.م', bold)],
    ]
    totals_table = Table(totals, colWidths=[45 * mm, 55 * mm], hAlign="RIGHT")
    totals_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#d4dbe3")),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    elements.append(totals_table)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=f"Invoice {sale['invoice_number']}",
    )
    document.build(elements)
    return output_path
