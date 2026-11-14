import os
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch

from database import Database, save_configured_database_path
from excel_export import export_invoices
from openpyxl import load_workbook
from pdf_invoice import generate_invoice_pdf


class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db = Database(Path(self.temp_dir.name) / "test.db")
        self.db.save_product(None, "مفصلة", "اكسسوار", "قطعة", 20, 35, 10)
        self.product_id = self.db.products()[0]["id"]
        self.db.save_customer(None, "عميل تجريبي", "01000000000", "")
        self.customer_id = self.db.customers()[0]["id"]

    def tearDown(self):
        self.db.close()
        self.temp_dir.cleanup()

    def test_sale_updates_stock_and_tracks_profit_and_debt(self):
        sale_id = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 3, "unit_price": 35}],
            40,
        )
        sale, items = self.db.sale(sale_id)
        summary = self.db.summary()

        self.assertEqual(self.db.product(self.product_id)["stock"], 7)
        self.assertEqual(sale["total"], 105)
        self.assertEqual(sale["paid"], 40)
        self.assertEqual(sale["due"], 65)
        self.assertEqual(summary["profit"], 45)
        self.assertEqual(summary["debt"], 65)
        self.assertEqual(items[0]["product_name"], "مفصلة")

        self.db.add_payment(sale_id, 65, "القسط الأخير")
        self.assertEqual(self.db.sale(sale_id)[0]["due"], 0)
        self.assertEqual(self.db.summary()["debt"], 0)

    def test_sale_rejects_quantity_above_stock_without_partial_changes(self):
        with self.assertRaisesRegex(ValueError, "الكمية المتاحة"):
            self.db.create_sale(
                self.customer_id,
                [{"product_id": self.product_id, "quantity": 11, "unit_price": 35}],
                0,
            )
        self.assertEqual(self.db.product(self.product_id)["stock"], 10)
        self.assertEqual(len(self.db.sales()), 0)

    def test_payment_cannot_exceed_remaining_debt(self):
        sale_id = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            10,
        )
        with self.assertRaisesRegex(ValueError, "أكبر من المديونية"):
            self.db.add_payment(sale_id, 26)
        self.assertEqual(self.db.sale(sale_id)[0]["paid"], 10)

    def test_customer_purchase_history_includes_invoice_totals_and_payments(self):
        first_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 2, "unit_price": 35}],
            20,
        )
        self.db.add_payment(first_sale, 15, "دفعة ثانية")
        self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            35,
        )
        self.db.create_sale(
            None,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            35,
        )

        history = self.db.customer_sales(self.customer_id)
        payments = self.db.sale_payments(first_sale)

        self.assertEqual(len(history), 2)
        self.assertEqual(history[0]["total"], 35)
        self.assertEqual(history[0]["due"], 0)
        self.assertEqual(history[1]["total"], 70)
        self.assertEqual(history[1]["paid"], 35)
        self.assertEqual(history[1]["due"], 35)
        self.assertEqual([payment["amount"] for payment in payments], [20, 15])
        self.assertEqual(payments[1]["note"], "دفعة ثانية")

    def test_follow_up_alerts_require_balance_threshold_and_inactivity(self):
        today = date.today()
        high_balance_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            0,
        )
        inactive_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            30,
        )
        recent_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            30,
        )
        recent_high_balance_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            0,
        )
        self.db.connection.execute(
            "UPDATE sales SET created_at=? WHERE id=?",
            ((today - timedelta(days=20)).isoformat(), high_balance_sale),
        )
        self.db.connection.execute(
            "UPDATE payments SET payment_date=? WHERE sale_id=?",
            ((today - timedelta(days=11)).isoformat(), inactive_sale),
        )
        self.db.connection.execute(
            "UPDATE payments SET payment_date=? WHERE sale_id=?",
            ((today - timedelta(days=2)).isoformat(), recent_sale),
        )
        self.db.connection.commit()

        alerts = self.db.follow_up_sales(20, 10, today=today)

        self.assertEqual({row["id"] for row in alerts}, {high_balance_sale})
        self.assertIn("المتبقي يبلغ", next(row for row in alerts if row["id"] == high_balance_sale)["reason"])
        self.assertNotIn(inactive_sale, {row["id"] for row in alerts})
        self.assertNotIn(recent_sale, {row["id"] for row in alerts})
        self.assertNotIn(recent_high_balance_sale, {row["id"] for row in alerts})

    def test_excel_export_includes_period_invoices_items_and_payments(self):
        today = date.today()
        included_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 2, "unit_price": 35}],
            20,
        )
        excluded_sale = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            0,
        )
        self.db.connection.execute(
            "UPDATE sales SET created_at=? WHERE id=?",
            ((today - timedelta(days=40)).isoformat(), excluded_sale),
        )
        self.db.connection.commit()
        self.db.add_payment(included_sale, 15, "دفعة اختبار")
        output = Path(self.temp_dir.name) / "invoices.xlsx"

        result = export_invoices(self.db, output, today - timedelta(days=5), today)

        self.assertEqual(result["invoice_count"], 1)
        workbook = load_workbook(output, data_only=True)
        self.assertEqual(workbook.sheetnames, ["الفواتير", "الأصناف", "الدفعات"])
        self.assertEqual(workbook["الفواتير"].max_row, 2)
        self.assertEqual(workbook["الفواتير"]["G2"].value, 35)
        self.assertEqual(workbook["الأصناف"].max_row, 2)
        self.assertEqual(workbook["الدفعات"].max_row, 3)
        self.assertEqual(workbook["الدفعات"]["E3"].value, "دفعة اختبار")

    def test_database_can_be_moved_and_reopened_from_user_selected_location(self):
        selected_location = Path(self.temp_dir.name) / "custom" / "business.db"
        app_data = Path(self.temp_dir.name) / "local-app-data"
        self.db.set_setting("shop_name", "متجر الاختبار")

        with patch.dict(os.environ, {"LOCALAPPDATA": str(app_data)}):
            self.db.backup_to(selected_location)
            save_configured_database_path(selected_location)
            reopened = Database()
            try:
                self.assertEqual(Path(reopened.path), selected_location.resolve())
                self.assertEqual(reopened.products()[0]["name"], "مفصلة")
                self.assertEqual(reopened.setting("shop_name"), "متجر الاختبار")
            finally:
                reopened.close()

    @unittest.skipUnless(
        any(
            (Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / regular).is_file()
            and (Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / bold).is_file()
            for regular, bold in (("tahoma.ttf", "tahomabd.ttf"), ("arial.ttf", "arialbd.ttf"))
        ),
        "Windows Arabic fonts are not installed.",
    )
    def test_invoice_pdf_is_created_with_pdf_signature(self):
        sale_id = self.db.create_sale(
            self.customer_id,
            [{"product_id": self.product_id, "quantity": 1, "unit_price": 35}],
            10,
        )
        output = Path(self.temp_dir.name) / "invoice.pdf"

        result = generate_invoice_pdf(self.db, sale_id, output)

        self.assertEqual(result, output)
        self.assertTrue(output.read_bytes().startswith(b"%PDF-"))


if __name__ == "__main__":
    unittest.main()
