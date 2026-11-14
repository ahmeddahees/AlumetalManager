import json
import os
import sqlite3
from datetime import date
from pathlib import Path


def default_database_path():
    root = os.environ.get("LOCALAPPDATA") or str(Path.home())
    directory = Path(root) / "AlumetalManager"
    directory.mkdir(parents=True, exist_ok=True)
    return directory / "business.db"


def _database_config_path():
    return default_database_path().parent / "database-location.json"


def configured_database_path():
    config_path = _database_config_path()
    if not config_path.exists():
        return default_database_path()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    path = config.get("database_path") if isinstance(config, dict) else None
    if not path:
        raise ValueError(f"مسار قاعدة البيانات غير صالح في ملف الإعدادات: {config_path}")
    database_path = Path(path)
    if not database_path.is_file():
        raise FileNotFoundError(f"لم يتم العثور على قاعدة البيانات المحددة: {database_path}")
    return database_path


def save_configured_database_path(path):
    config_path = _database_config_path()
    config_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = config_path.with_suffix(".tmp")
    temp_path.write_text(
        json.dumps({"database_path": str(Path(path).resolve())}, ensure_ascii=False),
        encoding="utf-8",
    )
    os.replace(temp_path, config_path)


class Database:
    def __init__(self, path=None):
        path = path or configured_database_path()
        self.path = str(path) if str(path) == ":memory:" else str(Path(path).resolve())
        self.connection = sqlite3.connect(self.path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self._create_tables()

    def close(self):
        self.connection.close()

    def backup_to(self, path):
        destination_path = Path(path).resolve()
        if destination_path.exists():
            raise FileExistsError(f"الملف موجود بالفعل: {destination_path}")
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        destination = sqlite3.connect(str(destination_path))
        try:
            self.connection.backup(destination)
        except sqlite3.Error:
            destination.close()
            if destination_path.exists():
                destination_path.unlink()
            raise
        destination.close()

    def _create_tables(self):
        self.connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT '',
                unit TEXT NOT NULL DEFAULT 'قطعة',
                cost_price REAL NOT NULL DEFAULT 0 CHECK(cost_price >= 0),
                sale_price REAL NOT NULL DEFAULT 0 CHECK(sale_price >= 0),
                stock REAL NOT NULL DEFAULT 0 CHECK(stock >= 0)
            );
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                phone TEXT NOT NULL DEFAULT '',
                address TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY,
                invoice_number TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL,
                customer_id INTEGER REFERENCES customers(id) ON DELETE SET NULL,
                customer_name TEXT NOT NULL DEFAULT '',
                total REAL NOT NULL,
                cost_total REAL NOT NULL
            );
            CREATE TABLE IF NOT EXISTS sale_items (
                id INTEGER PRIMARY KEY,
                sale_id INTEGER NOT NULL REFERENCES sales(id) ON DELETE CASCADE,
                product_id INTEGER REFERENCES products(id) ON DELETE SET NULL,
                product_name TEXT NOT NULL,
                unit TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_price REAL NOT NULL,
                unit_cost REAL NOT NULL
            );
            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY,
                sale_id INTEGER NOT NULL REFERENCES sales(id) ON DELETE CASCADE,
                amount REAL NOT NULL CHECK(amount > 0),
                payment_date TEXT NOT NULL,
                note TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )
        self.connection.commit()

    def setting(self, key, default=""):
        row = self.connection.execute(
            "SELECT value FROM settings WHERE key=?", (key,)
        ).fetchone()
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.connection.execute(
            "INSERT INTO settings(key,value) VALUES(?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value.strip()),
        )
        self.connection.commit()

    def products(self):
        return self.connection.execute(
            "SELECT * FROM products ORDER BY name COLLATE NOCASE"
        ).fetchall()

    def product(self, product_id):
        return self.connection.execute(
            "SELECT * FROM products WHERE id = ?", (product_id,)
        ).fetchone()

    def save_product(self, product_id, name, category, unit, cost, price, stock):
        values = (name.strip(), category.strip(), unit.strip() or "قطعة", cost, price, stock)
        if not values[0]:
            raise ValueError("اسم المنتج مطلوب.")
        if min(cost, price, stock) < 0:
            raise ValueError("الأسعار والكمية لا يمكن أن تكون سالبة.")
        if product_id:
            self.connection.execute(
                """UPDATE products SET name=?, category=?, unit=?, cost_price=?,
                   sale_price=?, stock=? WHERE id=?""",
                (*values, product_id),
            )
        else:
            self.connection.execute(
                """INSERT INTO products(name,category,unit,cost_price,sale_price,stock)
                   VALUES(?,?,?,?,?,?)""",
                values,
            )
        self.connection.commit()

    def delete_product(self, product_id):
        self.connection.execute("DELETE FROM products WHERE id=?", (product_id,))
        self.connection.commit()

    def customers(self):
        return self.connection.execute(
            "SELECT * FROM customers ORDER BY name COLLATE NOCASE"
        ).fetchall()

    def customer(self, customer_id):
        return self.connection.execute(
            "SELECT * FROM customers WHERE id=?", (customer_id,)
        ).fetchone()

    def save_customer(self, customer_id, name, phone, address):
        name = name.strip()
        if not name:
            raise ValueError("اسم العميل مطلوب.")
        if customer_id:
            self.connection.execute(
                "UPDATE customers SET name=?, phone=?, address=? WHERE id=?",
                (name, phone.strip(), address.strip(), customer_id),
            )
        else:
            self.connection.execute(
                "INSERT INTO customers(name,phone,address) VALUES(?,?,?)",
                (name, phone.strip(), address.strip()),
            )
        self.connection.commit()

    def delete_customer(self, customer_id):
        self.connection.execute("DELETE FROM customers WHERE id=?", (customer_id,))
        self.connection.commit()

    def create_sale(self, customer_id, items, paid):
        if not items:
            raise ValueError("أضف منتجًا واحدًا على الأقل للفاتورة.")
        if paid < 0:
            raise ValueError("المبلغ المدفوع لا يمكن أن يكون سالبًا.")

        with self.connection:
            normalized = []
            total = cost_total = 0.0
            for item in items:
                product = self.product(item["product_id"])
                quantity = float(item["quantity"])
                price = float(item["unit_price"])
                if not product:
                    raise ValueError("أحد المنتجات لم يعد موجودًا.")
                if quantity <= 0 or price < 0:
                    raise ValueError("تحقق من الكمية وسعر البيع.")
                if product["stock"] < quantity:
                    raise ValueError(
                        f"الكمية المتاحة من «{product['name']}» هي {product['stock']:g} فقط."
                    )
                normalized.append((product, quantity, price))
                total += quantity * price
                cost_total += quantity * product["cost_price"]
            if paid > total:
                raise ValueError("المبلغ المدفوع أكبر من إجمالي الفاتورة.")

            customer = self.customer(customer_id) if customer_id else None
            today = date.today().isoformat()
            cursor = self.connection.execute(
                """INSERT INTO sales(invoice_number,created_at,customer_id,customer_name,total,cost_total)
                   VALUES('',?,?,?,?,?)""",
                (today, customer_id, customer["name"] if customer else "عميل نقدي", total, cost_total),
            )
            sale_id = cursor.lastrowid
            invoice_number = f"INV-{date.today():%Y%m%d}-{sale_id:05d}"
            self.connection.execute(
                "UPDATE sales SET invoice_number=? WHERE id=?", (invoice_number, sale_id)
            )
            for product, quantity, price in normalized:
                self.connection.execute(
                    """INSERT INTO sale_items(sale_id,product_id,product_name,unit,quantity,unit_price,unit_cost)
                       VALUES(?,?,?,?,?,?,?)""",
                    (
                        sale_id,
                        product["id"],
                        product["name"],
                        product["unit"],
                        quantity,
                        price,
                        product["cost_price"],
                    ),
                )
                self.connection.execute(
                    "UPDATE products SET stock=stock-? WHERE id=?",
                    (quantity, product["id"]),
                )
            if paid:
                self.connection.execute(
                    "INSERT INTO payments(sale_id,amount,payment_date,note) VALUES(?,?,?,?)",
                    (sale_id, paid, today, "دفعة عند البيع"),
                )
        return sale_id

    def sales(self):
        return self.connection.execute(
            """SELECT s.*, COALESCE(SUM(p.amount),0) AS paid,
               s.total-COALESCE(SUM(p.amount),0) AS due
               FROM sales s LEFT JOIN payments p ON p.sale_id=s.id
               GROUP BY s.id ORDER BY s.id DESC"""
        ).fetchall()

    def sales_between(self, start_date, end_date):
        return self.connection.execute(
            """SELECT s.*, COALESCE(SUM(p.amount),0) AS paid,
               s.total-COALESCE(SUM(p.amount),0) AS due,
               COALESCE(c.phone,'') AS customer_phone
               FROM sales s
               LEFT JOIN payments p ON p.sale_id=s.id
               LEFT JOIN customers c ON c.id=s.customer_id
               WHERE s.created_at BETWEEN ? AND ?
               GROUP BY s.id ORDER BY s.created_at,s.id""",
            (start_date, end_date),
        ).fetchall()

    def follow_up_sales(self, threshold, inactive_days, today=None):
        if threshold < 0 or inactive_days < 1:
            raise ValueError("حد التنبيه يجب ألا يكون سالبًا، وعدد الأيام يجب أن يكون واحدًا أو أكثر.")
        today = today or date.today()
        rows = self.connection.execute(
            """SELECT s.*, COALESCE(SUM(p.amount),0) AS paid,
               s.total-COALESCE(SUM(p.amount),0) AS due,
               MAX(p.payment_date) AS last_payment_date
               FROM sales s LEFT JOIN payments p ON p.sale_id=s.id
               GROUP BY s.id HAVING s.total-COALESCE(SUM(p.amount),0)>0.009
               ORDER BY s.created_at,s.id"""
        ).fetchall()
        follow_ups = []
        for row in rows:
            last_activity = row["last_payment_date"] or row["created_at"]
            days_since_payment = (today - date.fromisoformat(last_activity)).days
            balance_exceeded = row["due"] + 0.00001 >= threshold
            payment_inactive = days_since_payment >= inactive_days
            if balance_exceeded and payment_inactive:
                follow_up = dict(row)
                follow_up["last_activity"] = last_activity
                follow_up["days_since_payment"] = days_since_payment
                follow_up["reason"] = (
                    f"المتبقي يبلغ الحد ({threshold:,.2f} ج.م) ولم يحدث تحصيل منذ "
                    f"{days_since_payment} يوم"
                )
                follow_ups.append(follow_up)
        return follow_ups

    def customer_sales(self, customer_id):
        return self.connection.execute(
            """SELECT s.*, COALESCE(SUM(p.amount),0) AS paid,
               s.total-COALESCE(SUM(p.amount),0) AS due
               FROM sales s LEFT JOIN payments p ON p.sale_id=s.id
               WHERE s.customer_id=? GROUP BY s.id ORDER BY s.id DESC""",
            (customer_id,),
        ).fetchall()

    def sale(self, sale_id):
        sale = self.connection.execute(
            """SELECT s.*, COALESCE(SUM(p.amount),0) AS paid,
               s.total-COALESCE(SUM(p.amount),0) AS due
               FROM sales s LEFT JOIN payments p ON p.sale_id=s.id
               WHERE s.id=? GROUP BY s.id""",
            (sale_id,),
        ).fetchone()
        if not sale:
            return None, []
        items = self.connection.execute(
            "SELECT * FROM sale_items WHERE sale_id=? ORDER BY id", (sale_id,)
        ).fetchall()
        return sale, items

    def sale_payments(self, sale_id):
        return self.connection.execute(
            """SELECT amount,payment_date,note FROM payments
               WHERE sale_id=? ORDER BY id""",
            (sale_id,),
        ).fetchall()

    def add_payment(self, sale_id, amount, note=""):
        if amount <= 0:
            raise ValueError("يجب أن يكون مبلغ القسط أكبر من صفر.")
        with self.connection:
            sale = self.connection.execute(
                """SELECT s.total-COALESCE(SUM(p.amount),0) AS due
                   FROM sales s LEFT JOIN payments p ON p.sale_id=s.id
                   WHERE s.id=? GROUP BY s.id""",
                (sale_id,),
            ).fetchone()
            if not sale:
                raise ValueError("الفاتورة غير موجودة.")
            if amount > sale["due"] + 0.00001:
                raise ValueError("مبلغ التحصيل أكبر من المديونية المتبقية.")
            self.connection.execute(
                "INSERT INTO payments(sale_id,amount,payment_date,note) VALUES(?,?,?,?)",
                (sale_id, amount, date.today().isoformat(), note.strip()),
            )

    def summary(self):
        return self.connection.execute(
            """SELECT
               COALESCE(SUM(total),0) AS sales,
               COALESCE(SUM(total-cost_total),0) AS profit,
               COALESCE(SUM(total-(SELECT COALESCE(SUM(amount),0) FROM payments p
                                  WHERE p.sale_id=sales.id)),0) AS debt,
               (SELECT COUNT(*) FROM products) AS product_count,
               (SELECT COALESCE(SUM(stock),0) FROM products) AS stock_count
               FROM sales"""
        ).fetchone()
