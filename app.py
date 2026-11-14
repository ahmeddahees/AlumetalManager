import html
import queue
import sqlite3
import subprocess
import sys
import threading
import tkinter as tk
import unicodedata
import webbrowser
from datetime import date, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from urllib.parse import quote

from database import (
    Database,
    default_database_path,
    save_configured_database_path,
)
from excel_export import export_invoices
from localization import translate
from pdf_invoice import generate_invoice_pdf
from update_manager import (
    NoReleaseFound,
    UpdateError,
    download_release,
    get_latest_release,
    start_windows_update,
)
from version import APP_VERSION


BG = "#f2f5f9"
NAVY = "#142d4e"
BLUE = "#2563a6"
GREEN = "#13825f"
RED = "#bd4b4b"
MUTED = "#6b7b90"
CURRENT_LANGUAGE = "ar"


def money(value):
    currency = "EGP" if CURRENT_LANGUAGE == "en" else "ج.م"
    return f"{value:,.2f} {currency}"


class ManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("إدارة محل الوميتال")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = min(1240, screen_width - 80)
        height = min(800, screen_height - 100)
        self.geometry(f"{width}x{height}+{max((screen_width - width) // 2, 0)}+{max((screen_height - height) // 2, 0)}")
        self.minsize(min(900, width), min(580, height))
        self.configure(bg=BG)
        self.db = Database()
        self._update_results = queue.Queue()
        self._update_check_in_progress = False
        self.language = self.db.setting("language", "ar")
        if self.language not in ("ar", "en"):
            self.language = "ar"
        global CURRENT_LANGUAGE
        CURRENT_LANGUAGE = self.language
        self._install_dialog_localization()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._style()
        self._build()
        self.refresh_all()
        self.after(1500, self.check_for_updates)

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, font=("Tahoma", 10))
        style.configure("Title.TLabel", background=BG, foreground=NAVY, font=("Tahoma", 20, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Tahoma", 10))
        style.configure("HeaderTitle.TLabel", background=NAVY, foreground="white", font=("Tahoma", 19, "bold"))
        style.configure("HeaderSub.TLabel", background=NAVY, foreground="#cbd8e8", font=("Tahoma", 10))
        style.configure("TButton", font=("Tahoma", 10), padding=(12, 8), background="#e7edf4", foreground=NAVY)
        style.map("TButton", background=[("active", "#d9e3ee")])
        style.configure("Primary.TButton", background=BLUE, foreground="white", font=("Tahoma", 10, "bold"))
        style.map("Primary.TButton", background=[("active", "#1d528b")])
        style.configure("Header.TButton", background="#294566", foreground="white", font=("Tahoma", 10, "bold"))
        style.map("Header.TButton", background=[("active", "#36587f")])
        style.configure("Treeview", font=("Tahoma", 10), rowheight=34, background="white", fieldbackground="white", borderwidth=0)
        style.configure("Treeview.Heading", font=("Tahoma", 10, "bold"), background="#e5edf5", foreground=NAVY, padding=(8, 10))
        style.map("Treeview.Heading", background=[("active", "#d9e4f0")])
        style.map("Treeview", background=[("selected", "#d6e8fa")], foreground=[("selected", NAVY)])
        style.configure("TNotebook", background=BG, tabmargins=(4, 8, 4, 0), borderwidth=0)
        style.configure("TNotebook.Tab", font=("Tahoma", 10, "bold"), padding=(16, 11), background="#e4eaf1", foreground=MUTED)
        style.map(
            "TNotebook.Tab",
            background=[("selected", BLUE), ("active", "#d8e3ef")],
            foreground=[("selected", "white"), ("active", NAVY)],
        )
        style.configure("TLabelframe", background=BG, bordercolor="#dce4ed", relief="solid")
        style.configure("TLabelframe.Label", background=BG, foreground=NAVY, font=("Tahoma", 11, "bold"))
        style.configure("TEntry", padding=7, fieldbackground="white")
        style.configure("TCombobox", padding=6, fieldbackground="white")

    def _build(self):
        self.title(self._text("إدارة محل الوميتال"))
        header = tk.Frame(self, bg=NAVY, padx=24, pady=17)
        header.pack(fill="x")
        title_block = tk.Frame(header, bg=NAVY)
        title_block.pack(side="right")
        ttk.Label(title_block, text="إدارة محل الوميتال", style="HeaderTitle.TLabel").pack(anchor="e")
        ttk.Label(title_block, text="المخزون  •  المبيعات  •  العملاء  •  التحصيل", style="HeaderSub.TLabel").pack(anchor="e", pady=(5, 0))
        header_buttons = tk.Frame(header, bg=NAVY)
        header_buttons.pack(side="left", pady=5)
        ttk.Button(header_buttons, text="بيانات المحل", style="Header.TButton", command=self.edit_shop_name).pack(side="left", padx=4)
        ttk.Button(header_buttons, text="مكان قاعدة البيانات", style="Header.TButton", command=self.change_database_location).pack(side="left", padx=4)

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=20, pady=(8, 18))
        self.dashboard_tab = ttk.Frame(self.tabs, padding=22)
        self.products_tab = ttk.Frame(self.tabs, padding=20)
        self.customers_tab = ttk.Frame(self.tabs, padding=20)
        self.invoice_tab = ttk.Frame(self.tabs, padding=20)
        self.sales_tab = ttk.Frame(self.tabs, padding=20)
        self.debts_tab = ttk.Frame(self.tabs, padding=20)
        self.alerts_tab = ttk.Frame(self.tabs, padding=20)
        for tab, label in (
            (self.dashboard_tab, "الرئيسية"),
            (self.products_tab, "المنتجات والمخزون"),
            (self.customers_tab, "العملاء"),
            (self.invoice_tab, "فاتورة جديدة"),
            (self.sales_tab, "الفواتير والمبيعات"),
            (self.debts_tab, "الديون والأقساط"),
            (self.alerts_tab, "تنبيهات المتابعة"),
        ):
            self.tabs.add(tab, text=label)
        self._build_dashboard()
        self._build_products()
        self._build_customers()
        self._build_invoice()
        self._build_sales()
        self._build_debts()
        self._build_alerts()
        self._build_menu()
        self._apply_locale(self)

    def _text(self, value):
        return translate(value, self.language)

    def _install_dialog_localization(self):
        for module, names in (
            (messagebox, ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel")),
            (simpledialog, ("askstring", "askinteger", "askfloat")),
        ):
            for name in names:
                original_name = f"_alumetal_original_{name}"
                if not hasattr(module, original_name):
                    setattr(module, original_name, getattr(module, name))
                original = getattr(module, original_name)

                def localized_call(*args, _original=original, **kwargs):
                    args = list(args)
                    for index in range(min(2, len(args))):
                        args[index] = self._text(args[index])
                    for key in ("title", "message", "prompt"):
                        if key in kwargs:
                            kwargs[key] = self._text(kwargs[key])
                    return _original(*args, **kwargs)

                setattr(module, name, localized_call)

        for name in ("askopenfilename", "asksaveasfilename", "askdirectory"):
            original_name = f"_alumetal_original_{name}"
            if not hasattr(filedialog, original_name):
                setattr(filedialog, original_name, getattr(filedialog, name))
            original = getattr(filedialog, original_name)

            def localized_file_dialog(*args, _original=original, **kwargs):
                if "title" in kwargs:
                    kwargs["title"] = self._text(kwargs["title"])
                if "filetypes" in kwargs:
                    kwargs["filetypes"] = [
                        (self._text(label), pattern)
                        for label, pattern in kwargs["filetypes"]
                    ]
                return _original(*args, **kwargs)

            setattr(filedialog, name, localized_file_dialog)

    def _build_menu(self):
        menu = tk.Menu(self, tearoff=False)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(
            label=self._text("فاتورة جديدة"),
            command=lambda: self.tabs.select(self.invoice_tab),
            accelerator="Ctrl+N",
        )
        file_menu.add_command(label=self._text("تصدير الفواتير إلى Excel"), command=self.export_excel_dialog)
        file_menu.add_separator()
        file_menu.add_command(label=self._text("مكان قاعدة البيانات"), command=self.change_database_location)
        file_menu.add_command(label=self._text("بيانات المحل"), command=self.edit_shop_name)
        file_menu.add_separator()
        file_menu.add_command(
            label=self._text("بحث عن تحديث"),
            command=lambda: self.check_for_updates(manual=True),
        )
        file_menu.add_separator()
        file_menu.add_command(label=self._text("إغلاق البرنامج"), command=self.on_close)

        sections_menu = tk.Menu(menu, tearoff=False)
        sections = (
            (self.dashboard_tab, "الرئيسية"),
            (self.products_tab, "المنتجات والمخزون"),
            (self.customers_tab, "العملاء"),
            (self.invoice_tab, "فاتورة جديدة"),
            (self.sales_tab, "الفواتير والمبيعات"),
            (self.debts_tab, "الديون والأقساط"),
            (self.alerts_tab, "تنبيهات المتابعة"),
        )
        for tab, label in sections:
            sections_menu.add_command(
                label=self._text(label),
                command=lambda selected_tab=tab: self.tabs.select(selected_tab),
            )

        language_menu = tk.Menu(menu, tearoff=False)
        self.language_menu_value = tk.StringVar(value=self.language)
        language_menu.add_radiobutton(
            label="العربية",
            value="ar",
            variable=self.language_menu_value,
            command=lambda: self.change_language("ar"),
        )
        language_menu.add_radiobutton(
            label="English",
            value="en",
            variable=self.language_menu_value,
            command=lambda: self.change_language("en"),
        )

        if self.language == "ar":
            menu_items = [
                (self._text("حول"), None),
                (self._text("اللغة"), language_menu),
                (self._text("الأقسام"), sections_menu),
                (self._text("ملف"), file_menu),
            ]
        else:
            menu_items = [
                (self._text("ملف"), file_menu),
                (self._text("الأقسام"), sections_menu),
                (self._text("اللغة"), language_menu),
                (self._text("حول"), None),
            ]
        for label, submenu in menu_items:
            if submenu is None:
                menu.add_command(label=label, command=self.show_about)
            else:
                menu.add_cascade(label=label, menu=submenu)
        self.menu_bar = menu
        self.file_menu = file_menu
        self.sections_menu = sections_menu
        self.config(menu=menu)
        self.bind_all("<Control-n>", lambda _event: self.tabs.select(self.invoice_tab))

    def _apply_locale(self, root):
        for widget in root.winfo_children():
            if isinstance(widget, tk.Menu):
                continue
            try:
                widget_text = widget.cget("text")
                if widget_text:
                    widget.configure(text=self._text(widget_text))
            except (tk.TclError, TypeError):
                pass

            if isinstance(widget, ttk.Notebook):
                tabs = list(widget.tabs())
                if self.language == "ar":
                    for index, tab in enumerate(reversed(tabs)):
                        widget.insert(index, tab)
                for tab in widget.tabs():
                    label = widget.tab(tab, "text")
                    widget.tab(tab, text=self._text(label))
            if isinstance(widget, ttk.Treeview):
                columns = widget["columns"]
                widget.configure(
                    displaycolumns=tuple(reversed(columns)) if self.language == "ar" else "#all"
                )
                for column in widget["columns"]:
                    heading = widget.heading(column, "text")
                    widget.heading(
                        column,
                        text=self._text(heading),
                        anchor="w" if self.language == "en" else "e",
                    )
                    widget.column(column, anchor="w" if self.language == "en" else "e")
            try:
                if "anchor" in widget.keys():
                    anchor = widget.cget("anchor")
                    if anchor in ("e", "w"):
                        widget.configure(anchor="w" if self.language == "en" else "e")
                if "justify" in widget.keys():
                    justify = widget.cget("justify")
                    if justify in ("left", "right"):
                        widget.configure(justify="left" if self.language == "en" else "right")
            except (tk.TclError, TypeError):
                pass

            try:
                pack_info = widget.pack_info()
                side = pack_info.get("side")
                anchor = pack_info.get("anchor")
                options = {}
                if side in ("left", "right"):
                    options["side"] = (
                        ("right" if side == "left" else "left")
                        if self.language == "en"
                        else side
                    )
                if anchor in ("e", "w"):
                    options["anchor"] = "w" if self.language == "en" else "e"
                if options:
                    widget.pack_configure(**options)
            except tk.TclError:
                pass
            try:
                grid_info = widget.grid_info()
                sticky = grid_info.get("sticky", "")
                if sticky and self.language == "en":
                    widget.grid_configure(
                        sticky=sticky.replace("e", "__west__").replace("w", "e").replace("__west__", "w")
                    )
            except tk.TclError:
                pass

            self._apply_locale(widget)

    def change_language(self, language):
        if language not in ("ar", "en") or language == self.language:
            return
        self.db.set_setting("language", language)
        self.language = language
        global CURRENT_LANGUAGE
        CURRENT_LANGUAGE = language
        self.unbind_all("<Control-n>")
        if hasattr(self, "product_search_popup"):
            self.hide_product_matches()
        self.config(menu="")
        for widget in self.winfo_children():
            widget.destroy()
        self._style()
        self._build()
        self.refresh_all()

    def check_for_updates(self, manual=False):
        if self._update_check_in_progress:
            if manual:
                messagebox.showinfo(
                    "بحث عن تحديث",
                    "يوجد فحص أو تنزيل تحديث جارٍ بالفعل.",
                    parent=self,
                )
            return
        self._start_update_task("check", manual=manual)

    def _start_update_task(self, operation, manual=False, release=None):
        self._update_check_in_progress = True

        def run():
            try:
                if operation == "check":
                    result = get_latest_release(APP_VERSION)
                else:
                    result = download_release(release)
            except NoReleaseFound:
                self._update_results.put((operation, manual, release, None, "no_release"))
            except UpdateError as error:
                self._update_results.put((operation, manual, release, None, str(error)))
            else:
                self._update_results.put((operation, manual, release, result, None))

        worker = threading.Thread(target=run, daemon=True)
        try:
            worker.start()
        except RuntimeError as error:
            self._update_check_in_progress = False
            messagebox.showerror(
                "تعذر التحقق من التحديث",
                str(error),
                parent=self,
            )
            return
        self.after(100, self._poll_update_results)

    def _poll_update_results(self):
        if not self.winfo_exists():
            return
        try:
            operation, manual, release, result, error = self._update_results.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_update_results)
            return

        self._update_check_in_progress = False
        if error == "no_release":
            if manual:
                messagebox.showinfo(
                    "بحث عن تحديث",
                    "لم يُنشر إصدار بعد. سيبدأ فحص التحديثات بعد نشر أول إصدار على GitHub.",
                    parent=self,
                )
            return
        if error:
            messagebox.showwarning(
                "تعذر التحقق من التحديث" if operation == "check" else "تعذر تنزيل التحديث",
                f"تعذر إكمال العملية: {error}",
                parent=self,
            )
            return

        if operation == "check":
            if result is None:
                if manual:
                    messagebox.showinfo(
                        "بحث عن تحديث",
                        "البرنامج لديك بأحدث إصدار.",
                        parent=self,
                    )
                return
            message = (
                f"{self._text('يتوفر تحديث جديد')}: {result.version}\n"
                f"{self._text('هل تريد تنزيله وتثبيته الآن؟')}"
            )
            if messagebox.askyesno("تحديث متاح", message, parent=self):
                if not getattr(sys, "frozen", False):
                    if not webbrowser.open(result.page_url):
                        messagebox.showwarning(
                            "تعذر فتح صفحة التحديث",
                            result.page_url,
                            parent=self,
                        )
                    return
                messagebox.showinfo(
                    "تنزيل التحديث",
                    "سيتم تنزيل التحديث والتحقق من سلامته قبل تثبيته.",
                    parent=self,
                )
                self._start_update_task("download", release=result)
            return

        try:
            start_windows_update(result, sys.executable)
        except UpdateError as error:
            messagebox.showerror(
                "تعذر تثبيت التحديث",
                str(error),
                parent=self,
            )
            return
        messagebox.showinfo(
            "تحديث البرنامج",
            "سيُغلق البرنامج لتثبيت التحديث ثم يُعاد تشغيله.",
            parent=self,
        )
        self.on_close()

    def show_about(self):
        dialog = tk.Toplevel(self)
        dialog.title(self._text("حول"))
        dialog.transient(self)
        dialog.configure(bg=BG)
        ttk.Label(dialog, text=self._text("حول برنامج إدارة محل الوميتال"), style="Title.TLabel").pack(
            padx=24, pady=(22, 14)
        )
        ttk.Label(
            dialog,
            text=self._text("تمت البرمجة والتصميم الأستاذ أحمد دهيس"),
            font=("Tahoma", 11, "bold"),
        ).pack(padx=24, pady=6)
        ttk.Label(dialog, text=self._text("للتواصل")).pack(padx=24, pady=(12, 4))
        for phone in ("01101083878", "01015362440"):
            ttk.Label(dialog, text=phone, font=("Tahoma", 11)).pack(padx=24, pady=3)
        ttk.Button(dialog, text=self._text("إغلاق"), command=dialog.destroy).pack(pady=(15, 20))
        self._center_dialog(dialog)

    def _build_dashboard(self):
        ttk.Label(self.dashboard_tab, text="ملخص النشاط", style="Title.TLabel").pack(anchor="e", pady=(0, 18))
        self.cards = ttk.Frame(self.dashboard_tab)
        self.cards.pack(fill="x")
        self.card_values = {}
        for key, title, color in (
            ("sales", "إجمالي المبيعات", BLUE),
            ("profit", "إجمالي الربح", GREEN),
            ("debt", "الديون المتبقية", RED),
            ("products", "عدد المنتجات", NAVY),
            ("stock", "إجمالي الكميات بالمخزون", NAVY),
        ):
            card = tk.Frame(self.cards, bg="white", highlightbackground="#dce4ed", highlightthickness=1, padx=16, pady=15)
            card.pack(side="right", fill="both", expand=True, padx=5)
            tk.Frame(card, bg=color, height=4).pack(fill="x", pady=(0, 12))
            tk.Label(card, text=title, bg="white", fg=MUTED, font=("Tahoma", 10)).pack(anchor="e")
            value = tk.Label(card, text="—", bg="white", fg=color, font=("Tahoma", 16, "bold"))
            value.pack(anchor="e", pady=(12, 2))
            self.card_values[key] = value
        ttk.Label(
            self.dashboard_tab,
            text="يُحسب الربح من سعر التكلفة المسجل وقت البيع. لإضافة كمية واردة للمخزون، عدّل رصيد المنتج.",
            style="Sub.TLabel",
        ).pack(anchor="e", pady=22)

    def _build_products(self):
        self.products_tree = self._tree(
            self.products_tab,
            (("name", "اسم المنتج", 220), ("category", "التصنيف", 120), ("unit", "الوحدة", 90),
             ("cost", "التكلفة", 120), ("price", "سعر البيع", 120), ("stock", "المتاح", 100)),
        )
        buttons = ttk.Frame(self.products_tab)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="إضافة منتج", style="Primary.TButton", command=self.add_product).pack(side="right", padx=4)
        ttk.Button(buttons, text="تعديل المحدد", command=self.edit_product).pack(side="right", padx=4)
        ttk.Button(buttons, text="حذف المحدد", command=self.delete_product).pack(side="right", padx=4)

    def _build_customers(self):
        self.customers_tree = self._tree(
            self.customers_tab,
            (("name", "اسم العميل", 230), ("phone", "رقم الهاتف", 170), ("address", "العنوان", 300)),
        )
        self.customers_tree.bind("<Double-1>", lambda _event: self.show_customer_account())
        buttons = ttk.Frame(self.customers_tab)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="كشف حساب ومشتريات العميل", style="Primary.TButton", command=self.show_customer_account).pack(side="right", padx=4)
        ttk.Button(buttons, text="إضافة عميل", style="Primary.TButton", command=self.add_customer).pack(side="right", padx=4)
        ttk.Button(buttons, text="تعديل المحدد", command=self.edit_customer).pack(side="right", padx=4)
        ttk.Button(buttons, text="حذف المحدد", command=self.delete_customer).pack(side="right", padx=4)

    def _build_invoice(self):
        form = ttk.LabelFrame(self.invoice_tab, text="بيانات الفاتورة", padding=12)
        form.pack(fill="x", pady=(0, 12))
        ttk.Label(form, text="العميل (اختياري):").grid(row=0, column=3, padx=8, pady=6, sticky="e")
        self.invoice_customer = ttk.Combobox(form, state="readonly", width=28, justify="right")
        self.invoice_customer.grid(row=0, column=2, padx=8, pady=6)
        ttk.Label(form, text="المنتج:").grid(row=0, column=1, padx=8, pady=6, sticky="e")
        self.invoice_product = ttk.Entry(form, width=30, justify="right")
        self.invoice_product.grid(row=0, column=0, padx=8, pady=6)
        self.invoice_product.bind("<KeyRelease>", self.filter_invoice_products)
        self.invoice_product.bind("<FocusIn>", self.show_product_matches)
        self.invoice_product.bind("<Button-1>", self.show_product_matches)
        self.invoice_product.bind("<FocusOut>", self.on_product_search_focus_out)
        self.invoice_product.bind("<Down>", self.select_next_product_match)
        self.invoice_product.bind("<Up>", self.select_previous_product_match)
        self.invoice_product.bind("<Return>", self.choose_product_match)
        self.invoice_product.bind("<Escape>", self.hide_product_matches)
        ttk.Label(form, text="الكمية:").grid(row=1, column=3, padx=8, pady=6, sticky="e")
        self.invoice_quantity = ttk.Entry(form, width=14, justify="right")
        self.invoice_quantity.insert(0, "1")
        self.invoice_quantity.grid(row=1, column=2, padx=8, pady=6, sticky="w")
        ttk.Label(form, text="سعر الوحدة (ج.م):").grid(row=1, column=1, padx=8, pady=6, sticky="e")
        self.invoice_price = ttk.Entry(form, width=14, justify="right")
        self.invoice_price.grid(row=1, column=0, padx=8, pady=6, sticky="w")
        ttk.Button(form, text="إضافة للفاتورة", command=self.add_cart_item).grid(row=2, column=0, padx=8, pady=6, sticky="w")
        self.invoice_search_hint = ttk.Label(
            form, text="اكتب جزءًا من اسم المنتج لعرض المنتجات المطابقة فقط.", style="Sub.TLabel"
        )
        self.invoice_search_hint.grid(row=2, column=1, columnspan=3, padx=8, pady=6, sticky="e")

        self.cart_tree = self._tree(
            self.invoice_tab,
            (("name", "المنتج", 260), ("quantity", "الكمية", 100),
             ("price", "سعر الوحدة", 130), ("total", "الإجمالي", 140)),
            height=5,
        )
        self.cart_items = []
        action_row = ttk.Frame(self.invoice_tab)
        action_row.pack(fill="x", pady=10)
        ttk.Button(action_row, text="حذف الصنف المحدد", command=self.remove_cart_item).pack(side="right", padx=4)
        ttk.Label(action_row, text="المبلغ المدفوع الآن (ج.م):").pack(side="right", padx=(18, 6))
        self.invoice_paid = ttk.Entry(action_row, width=15, justify="right")
        self.invoice_paid.insert(0, "0")
        self.invoice_paid.pack(side="right")
        self.invoice_total_label = ttk.Label(action_row, text="الإجمالي: 0.00 ج.م", style="Title.TLabel")
        self.invoice_total_label.pack(side="left", padx=8)
        ttk.Button(action_row, text="حفظ الفاتورة", style="Primary.TButton", command=self.save_invoice).pack(side="left", padx=10)

    def _build_sales(self):
        self.sales_tree = self._tree(
            self.sales_tab,
            (("invoice", "رقم الفاتورة", 185), ("date", "التاريخ", 120), ("customer", "العميل", 190),
             ("total", "الإجمالي", 130), ("paid", "المدفوع", 130), ("due", "المتبقي", 130)),
        )
        self.sales_tree.bind("<Double-1>", lambda _event: self.show_selected_sale_details(self.sales_tree))
        buttons = ttk.Frame(self.sales_tab)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="طباعة الفاتورة المحددة", style="Primary.TButton", command=self.print_selected_invoice).pack(side="right", padx=4)
        ttk.Button(buttons, text="تفاصيل الفاتورة", command=lambda: self.show_selected_sale_details(self.sales_tree)).pack(side="right", padx=4)
        ttk.Button(buttons, text="تسجيل دفعة على الفاتورة", command=lambda: self.record_payment(self.sales_tree)).pack(side="right", padx=4)
        ttk.Button(buttons, text="تصدير Excel", command=self.export_excel_dialog).pack(side="left", padx=4)
        ttk.Button(buttons, text="واتساب", command=lambda: self.send_selected_invoice_whatsapp(self.sales_tree)).pack(side="left", padx=4)

    def _build_debts(self):
        self.debts_tree = self._tree(
            self.debts_tab,
            (("invoice", "رقم الفاتورة", 190), ("date", "التاريخ", 120), ("customer", "العميل", 210),
             ("total", "قيمة الفاتورة", 135), ("paid", "المسدد", 135), ("due", "المتبقي", 135)),
        )
        self.debts_tree.bind("<Double-1>", lambda _event: self.show_selected_sale_details(self.debts_tree))
        buttons = ttk.Frame(self.debts_tab)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(
            buttons, text="تسجيل تحصيل / قسط", style="Primary.TButton",
            command=lambda: self.record_payment(self.debts_tree),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="طباعة الفاتورة المحددة",
            command=lambda: self.print_selected_invoice(self.debts_tree),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="تفاصيل الفاتورة",
            command=lambda: self.show_selected_sale_details(self.debts_tree),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="واتساب",
            command=lambda: self.send_selected_invoice_whatsapp(self.debts_tree),
        ).pack(side="left", padx=4)

    def _build_alerts(self):
        header = ttk.Frame(self.alerts_tab)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="فواتير تحتاج إلى متابعة", style="Title.TLabel").pack(side="right")
        ttk.Button(header, text="إعدادات التنبيه", command=self.edit_alert_settings).pack(side="left", padx=4)
        ttk.Button(header, text="تحديث", command=self.refresh_alerts).pack(side="left", padx=4)
        self.alerts_summary = ttk.Label(self.alerts_tab, text="", style="Sub.TLabel")
        self.alerts_summary.pack(anchor="e", pady=(0, 10))
        self.alerts_tree = self._tree(
            self.alerts_tab,
            (("invoice", "الفاتورة", 165), ("date", "التاريخ", 110), ("customer", "العميل", 170),
             ("due", "المتبقي", 125), ("last_activity", "آخر تحصيل / بيع", 145),
             ("days", "أيام دون تحصيل", 110), ("reason", "سبب التنبيه", 230)),
        )
        self.alerts_tree.bind(
            "<Double-1>", lambda _event: self.show_selected_sale_details(self.alerts_tree)
        )
        buttons = ttk.Frame(self.alerts_tab)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(
            buttons, text="تسجيل تحصيل", style="Primary.TButton",
            command=lambda: self.record_payment(self.alerts_tree),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="إرسال واتساب",
            command=lambda: self.send_selected_invoice_whatsapp(self.alerts_tree),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="طباعة الفاتورة",
            command=lambda: self.print_selected_invoice(self.alerts_tree),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="تفاصيل الفاتورة",
            command=lambda: self.show_selected_sale_details(self.alerts_tree),
        ).pack(side="right", padx=4)

    def _tree(self, parent, columns, height=10):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        names = [col[0] for col in columns]
        tree = ttk.Treeview(frame, columns=names, show="headings", height=height, selectmode="browse")
        tree.tag_configure("odd", background="#f7f9fc")
        tree.tag_configure("even", background="white")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        for key, title, width in columns:
            tree.heading(key, text=title, anchor="e")
            tree.column(key, width=width, minwidth=65, anchor="e", stretch=True)
        tree.pack(side="right", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")
        return tree

    def refresh_all(self):
        self.refresh_dashboard()
        self.refresh_products()
        self.refresh_customers()
        self.refresh_sales()
        self.refresh_debts()
        self.refresh_alerts()
        self.refresh_invoice_choices()

    def refresh_dashboard(self):
        data = self.db.summary()
        self.card_values["sales"].config(text=money(data["sales"]))
        self.card_values["profit"].config(text=money(data["profit"]))
        self.card_values["debt"].config(text=money(data["debt"]))
        self.card_values["products"].config(text=str(data["product_count"]))
        self.card_values["stock"].config(text=f'{data["stock_count"]:g}')

    def refresh_products(self):
        self.products_tree.delete(*self.products_tree.get_children())
        for index, row in enumerate(self.db.products()):
            self.products_tree.insert(
                "", "end", iid=str(row["id"]),
                values=(row["name"], row["category"], row["unit"], money(row["cost_price"]),
                        money(row["sale_price"]), f'{row["stock"]:g}'),
                tags=("even" if index % 2 == 0 else "odd",),
            )

    def refresh_customers(self):
        self.customers_tree.delete(*self.customers_tree.get_children())
        for index, row in enumerate(self.db.customers()):
            self.customers_tree.insert(
                "", "end", iid=str(row["id"]),
                values=(row["name"], row["phone"], row["address"]),
                tags=("even" if index % 2 == 0 else "odd",),
            )

    def refresh_sales(self):
        self.sales_tree.delete(*self.sales_tree.get_children())
        for index, row in enumerate(self.db.sales()):
            self.sales_tree.insert(
                "", "end", iid=str(row["id"]),
                values=(row["invoice_number"], row["created_at"], row["customer_name"],
                        money(row["total"]), money(row["paid"]), money(row["due"])),
                tags=("even" if index % 2 == 0 else "odd",),
            )

    def refresh_debts(self):
        self.debts_tree.delete(*self.debts_tree.get_children())
        for index, row in enumerate(self.db.sales()):
            if row["due"] > 0.009:
                self.debts_tree.insert(
                    "", "end", iid=str(row["id"]),
                    values=(row["invoice_number"], row["created_at"], row["customer_name"],
                            money(row["total"]), money(row["paid"]), money(row["due"])),
                    tags=("even" if index % 2 == 0 else "odd",),
                )

    def refresh_alerts(self):
        if not hasattr(self, "alerts_tree"):
            return
        threshold = float(self.db.setting("alert_threshold", "5000"))
        inactive_days = int(self.db.setting("alert_days", "10"))
        rows = self.db.follow_up_sales(threshold, inactive_days)
        self.alerts_tree.delete(*self.alerts_tree.get_children())
        for index, sale in enumerate(rows):
            self.alerts_tree.insert(
                "", "end", iid=str(sale["id"]),
                values=(
                    sale["invoice_number"], sale["created_at"], sale["customer_name"],
                    money(sale["due"]), sale["last_activity"], sale["days_since_payment"],
                    sale["reason"],
                ),
                tags=("even" if index % 2 == 0 else "odd",),
            )
        if self.language == "en":
            summary = (
                f"Open invoices with a balance of {money(threshold)} or more and no payment "
                f"for at least {inactive_days} days: {len(rows)}"
            )
        else:
            summary = (
                f"فواتير متبقيها {money(threshold)} أو أكثر ولم يحدث عليها تحصيل منذ "
                f"{inactive_days} أيام أو أكثر: {len(rows)}"
            )
        self.alerts_summary.config(text=summary)

    def refresh_invoice_choices(self):
        self.product_map = {}
        products = self.db.products()
        for row in products:
            available = self._text(" — متاح ")
            label = f'{row["name"]}{available}{row["stock"]:g} {row["unit"]}'
            self.product_map[label] = row
        self.product_labels = list(self.product_map)
        walk_in = self._text("عميل نقدي")
        self.customer_map = {walk_in: None}
        for row in self.db.customers():
            self.customer_map[row["name"]] = row["id"]
        self.invoice_customer["values"] = list(self.customer_map)
        if self.invoice_customer.get() not in self.customer_map:
            self.invoice_customer.set(walk_in)
        self.invoice_product.delete(0, "end")
        self.hide_product_matches()

    def filter_invoice_products(self, event=None):
        if event and event.keysym in {
            "Up", "Down", "Left", "Right", "Return", "Escape", "Tab",
            "Home", "End", "Prior", "Next",
        }:
            return
        query = self.invoice_product.get().strip().casefold()
        matches = [
            label for label in self.product_labels
            if not query or query in self.product_map[label]["name"].casefold()
        ]
        self._update_product_match_list(matches)
        if query and matches:
            self.invoice_search_hint.config(text=f'{self._text("نتائج البحث: ")}{len(matches)} {self._text("منتج")}')
        elif not query and matches:
            self.invoice_search_hint.config(text=f'{self._text("كل المنتجات: ")}{len(matches)} {self._text("منتج")}')
        elif query:
            self.invoice_search_hint.config(text=self._text("لا توجد منتجات مطابقة لاسم البحث."))
        else:
            self.invoice_search_hint.config(text=self._text("لا توجد منتجات مضافة بعد."))

    def show_product_matches(self, _event=None):
        self.filter_invoice_products()

    def _update_product_match_list(self, matches):
        if not matches:
            self.hide_product_matches()
            return
        if not hasattr(self, "product_search_popup") or not self.product_search_popup.winfo_exists():
            self.product_search_popup = tk.Toplevel(self)
            self.product_search_popup.withdraw()
            self.product_search_popup.overrideredirect(True)
            self.product_search_popup.transient(self)
            popup_frame = ttk.Frame(self.product_search_popup, padding=2)
            popup_frame.pack(fill="both", expand=True)
            self.product_match_list = tk.Listbox(
                popup_frame,
                height=1,
                activestyle="none",
                font=("Tahoma", 10),
                bg="white",
                fg=NAVY,
                selectbackground="#d6e8fa",
                selectforeground=NAVY,
                highlightthickness=1,
                highlightbackground="#dce4ed",
                relief="flat",
                exportselection=False,
                takefocus=False,
            )
            scrollbar = ttk.Scrollbar(popup_frame, orient="vertical", command=self.product_match_list.yview)
            self.product_match_list.configure(yscrollcommand=scrollbar.set)
            self.product_match_list.pack(side="right", fill="both", expand=True)
            scrollbar.pack(side="left", fill="y")
            self.product_match_list.bind("<Button-1>", self.choose_clicked_product_match)

        self.product_match_list.delete(0, "end")
        for label in matches:
            self.product_match_list.insert("end", label)
        self.product_match_list.configure(height=min(len(matches), 8))
        self.product_match_list.selection_clear(0, "end")
        self.product_match_list.activate(0)
        self.product_match_list.selection_set(0)

        self.update_idletasks()
        width = self.invoice_product.winfo_width()
        row_height = self.product_match_list.winfo_reqheight() / min(len(matches), 8)
        height = min(len(matches), 8) * row_height + 6
        x = self.invoice_product.winfo_rootx()
        y = self.invoice_product.winfo_rooty() + self.invoice_product.winfo_height()
        screen_bottom = self.winfo_screenheight()
        if y + height > screen_bottom:
            y = max(self.invoice_product.winfo_rooty() - height, 0)
        self.product_search_popup.geometry(f"{width}x{int(height)}+{x}+{y}")
        self.product_search_popup.deiconify()
        self.product_search_popup.lift()
        self.invoice_product.focus_set()

    def hide_product_matches(self, _event=None):
        if hasattr(self, "product_search_popup") and self.product_search_popup.winfo_exists():
            self.product_search_popup.withdraw()

    def on_product_search_focus_out(self, _event=None):
        self.after_idle(self._hide_product_matches_if_focus_left)

    def _hide_product_matches_if_focus_left(self):
        focus = self.focus_get()
        if focus not in (self.invoice_product, getattr(self, "product_match_list", None)):
            self.hide_product_matches()

    def choose_clicked_product_match(self, event):
        index = self.product_match_list.nearest(event.y)
        self.product_match_list.selection_clear(0, "end")
        self.product_match_list.selection_set(index)
        self.product_match_list.activate(index)
        self.choose_product_match()
        return "break"

    def select_next_product_match(self, _event=None):
        self._move_product_match_selection(1)
        return "break"

    def select_previous_product_match(self, _event=None):
        self._move_product_match_selection(-1)
        return "break"

    def _move_product_match_selection(self, step):
        if not self.product_match_list.size():
            return
        selection = self.product_match_list.curselection()
        current = selection[0] if selection else 0
        index = max(0, min(current + step, self.product_match_list.size() - 1))
        self.product_match_list.selection_clear(0, "end")
        self.product_match_list.selection_set(index)
        self.product_match_list.activate(index)
        self.product_match_list.see(index)

    def choose_product_match(self, _event=None):
        if not hasattr(self, "product_match_list") or not self.product_match_list.winfo_exists():
            return "break"
        selection = self.product_match_list.curselection()
        if not selection:
            return "break"
        self.invoice_product.delete(0, "end")
        self.invoice_product.insert(0, self.product_match_list.get(selection[0]))
        self.hide_product_matches()
        self.on_product_selected()
        self.invoice_product.focus_set()
        self.invoice_product.icursor("end")
        return "break"

    def _selected_id(self, tree, label):
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("تنبيه", f"حدد {label} أولًا.", parent=self)
            return None
        return int(selection[0])

    def _product_dialog(self, title, record=None):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=BG)
        fields = {}
        defaults = {
            "name": record["name"] if record else "",
            "category": record["category"] if record else "",
            "unit": record["unit"] if record else "قطعة",
            "cost": str(record["cost_price"] if record else "0"),
            "price": str(record["sale_price"] if record else "0"),
            "stock": str(record["stock"] if record else "0"),
        }
        for row, (key, label) in enumerate((
            ("name", "اسم المنتج"), ("category", "التصنيف"), ("unit", "الوحدة"),
            ("cost", "سعر التكلفة"), ("price", "سعر البيع"), ("stock", "الكمية بالمخزون"),
        )):
            ttk.Label(dialog, text=label).grid(row=row, column=1, padx=12, pady=7, sticky="e")
            entry = ttk.Entry(dialog, width=30, justify="right")
            entry.insert(0, defaults[key])
            entry.grid(row=row, column=0, padx=12, pady=7)
            fields[key] = entry

        def save():
            try:
                self.db.save_product(
                    record["id"] if record else None,
                    fields["name"].get(), fields["category"].get(), fields["unit"].get(),
                    float(fields["cost"].get()), float(fields["price"].get()), float(fields["stock"].get()),
                )
            except (ValueError, sqlite3.Error) as error:
                messagebox.showerror("تعذر حفظ المنتج", str(error), parent=dialog)
                return
            dialog.destroy()
            self.refresh_all()

        ttk.Button(dialog, text="حفظ", style="Primary.TButton", command=save).grid(row=6, column=0, padx=12, pady=12, sticky="w")
        dialog.bind("<Return>", lambda _event: save())
        fields["name"].focus_set()
        self._center_dialog(dialog)

    def add_product(self):
        self._product_dialog("إضافة منتج")

    def edit_product(self):
        product_id = self._selected_id(self.products_tree, "منتجًا")
        if product_id is not None:
            self._product_dialog("تعديل المنتج", self.db.product(product_id))

    def delete_product(self):
        product_id = self._selected_id(self.products_tree, "منتجًا")
        if product_id is None:
            return
        if messagebox.askyesno("تأكيد الحذف", "هل تريد حذف المنتج؟ سيظل اسمه محفوظًا في الفواتير السابقة.", parent=self):
            self.db.delete_product(product_id)
            self.refresh_all()

    def _customer_dialog(self, title, record=None):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=BG)
        fields = {}
        for row, (key, label) in enumerate((("name", "اسم العميل"), ("phone", "رقم الهاتف"), ("address", "العنوان"))):
            ttk.Label(dialog, text=label).grid(row=row, column=1, padx=12, pady=8, sticky="e")
            entry = ttk.Entry(dialog, width=32, justify="right")
            entry.insert(0, record[key] if record else "")
            entry.grid(row=row, column=0, padx=12, pady=8)
            fields[key] = entry

        def save():
            try:
                self.db.save_customer(
                    record["id"] if record else None,
                    fields["name"].get(), fields["phone"].get(), fields["address"].get(),
                )
            except ValueError as error:
                messagebox.showerror("تعذر حفظ العميل", str(error), parent=dialog)
                return
            dialog.destroy()
            self.refresh_all()

        ttk.Button(dialog, text="حفظ", style="Primary.TButton", command=save).grid(row=3, column=0, padx=12, pady=12, sticky="w")
        dialog.bind("<Return>", lambda _event: save())
        fields["name"].focus_set()
        self._center_dialog(dialog)

    def add_customer(self):
        self._customer_dialog("إضافة عميل")

    def edit_customer(self):
        customer_id = self._selected_id(self.customers_tree, "عميلًا")
        if customer_id is not None:
            self._customer_dialog("تعديل العميل", self.db.customer(customer_id))

    def delete_customer(self):
        customer_id = self._selected_id(self.customers_tree, "عميلًا")
        if customer_id is None:
            return
        if messagebox.askyesno("تأكيد الحذف", "هل تريد حذف العميل؟ ستبقى الفواتير السابقة محفوظة.", parent=self):
            self.db.delete_customer(customer_id)
            self.refresh_all()

    def show_customer_account(self):
        customer_id = self._selected_id(self.customers_tree, "عميلًا")
        if customer_id is None:
            return
        customer = self.db.customer(customer_id)
        if not customer:
            messagebox.showerror("خطأ", "العميل غير موجود.", parent=self)
            return
        sales = self.db.customer_sales(customer_id)

        dialog = tk.Toplevel(self)
        dialog.title(f'كشف حساب العميل - {customer["name"]}')
        dialog.transient(self)
        dialog.geometry("940x600")
        dialog.minsize(720, 470)
        dialog.configure(bg=BG)
        header = ttk.Frame(dialog, padding=(18, 14))
        header.pack(fill="x")
        ttk.Label(header, text=customer["name"], style="Title.TLabel").pack(anchor="e")
        details = f'الهاتف: {customer["phone"] or "—"}    |    العنوان: {customer["address"] or "—"}'
        ttk.Label(header, text=details, style="Sub.TLabel").pack(anchor="e", pady=(5, 0))

        total_sales = sum(row["total"] for row in sales)
        total_paid = sum(row["paid"] for row in sales)
        total_due = sum(row["due"] for row in sales)
        summary = ttk.Frame(dialog, padding=(14, 0, 14, 12))
        summary.pack(fill="x")
        ttk.Label(summary, text=f"عدد الفواتير: {len(sales)}").pack(side="right", padx=12)
        ttk.Label(summary, text=f"إجمالي المشتريات: {money(total_sales)}").pack(side="right", padx=12)
        ttk.Label(summary, text=f"إجمالي المدفوع: {money(total_paid)}").pack(side="right", padx=12)
        ttk.Label(summary, text=f"المتبقي: {money(total_due)}", foreground=RED).pack(side="right", padx=12)

        tree = self._tree(
            dialog,
            (("invoice", "رقم الفاتورة", 185), ("date", "تاريخ البيع", 125),
             ("total", "الإجمالي", 150), ("paid", "المدفوع", 150), ("due", "المتبقي", 150)),
            height=9,
        )
        for index, sale in enumerate(sales):
            tree.insert(
                "", "end", iid=str(sale["id"]),
                values=(sale["invoice_number"], sale["created_at"], money(sale["total"]),
                        money(sale["paid"]), money(sale["due"])),
                tags=("even" if index % 2 == 0 else "odd",),
            )
        tree.bind("<Double-1>", lambda _event: self.show_selected_sale_details(tree, dialog))
        buttons = ttk.Frame(dialog, padding=12)
        buttons.pack(fill="x")
        ttk.Button(
            buttons, text="طباعة كشف مشتريات العميل",
            command=lambda: self.print_customer_statement(customer_id),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="طباعة الفاتورة المحددة", style="Primary.TButton",
            command=lambda: self.print_selected_invoice(tree, dialog),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="إرسال الفاتورة واتساب",
            command=lambda: self.send_selected_invoice_whatsapp(tree, dialog),
        ).pack(side="right", padx=4)
        ttk.Button(
            buttons, text="تفاصيل الفاتورة المحددة",
            command=lambda: self.show_selected_sale_details(tree, dialog),
        ).pack(side="right", padx=4)
        ttk.Button(buttons, text="إغلاق", command=dialog.destroy).pack(side="left", padx=4)
        self._center_dialog(dialog)

    def show_selected_sale_details(self, tree, parent=None):
        sale_id = self._selected_id(tree, "فاتورة")
        if sale_id is not None:
            self.show_sale_details(sale_id, parent)

    def show_sale_details(self, sale_id, parent=None):
        sale, items = self.db.sale(sale_id)
        if not sale:
            messagebox.showerror("خطأ", "الفاتورة غير موجودة.", parent=parent or self)
            return
        dialog = tk.Toplevel(self)
        dialog.title(f'تفاصيل الفاتورة {sale["invoice_number"]}')
        dialog.transient(parent or self)
        dialog.geometry("900x650")
        dialog.minsize(700, 500)
        dialog.configure(bg=BG)

        heading = ttk.Frame(dialog, padding=(18, 14))
        heading.pack(fill="x")
        ttk.Label(heading, text="تفاصيل الفاتورة", style="Title.TLabel").pack(anchor="e")
        ttk.Label(
            heading,
            text=f'رقم الفاتورة: {sale["invoice_number"]}    |    التاريخ: {sale["created_at"]}    |    العميل: {sale["customer_name"]}',
            style="Sub.TLabel",
        ).pack(anchor="e", pady=(5, 0))

        ttk.Label(dialog, text="الأصناف", font=("Tahoma", 11, "bold")).pack(anchor="e", padx=18, pady=(4, 6))
        items_tree = self._tree(
            dialog,
            (("name", "المنتج", 260), ("quantity", "الكمية", 120),
             ("unit_price", "سعر الوحدة", 160), ("total", "إجمالي الصنف", 170)),
            height=6,
        )
        for index, item in enumerate(items):
            items_tree.insert(
                "", "end", values=(
                    item["product_name"], f'{item["quantity"]:g} {item["unit"]}',
                    money(item["unit_price"]), money(item["quantity"] * item["unit_price"]),
                ), tags=("even" if index % 2 == 0 else "odd",),
            )

        payments = self.db.sale_payments(sale_id)
        ttk.Label(dialog, text="سجل الدفعات والأقساط", font=("Tahoma", 11, "bold")).pack(anchor="e", padx=18, pady=(12, 6))
        payment_tree = self._tree(
            dialog,
            (("date", "التاريخ", 150), ("amount", "المبلغ", 180), ("note", "البيان", 360)),
            height=4,
        )
        for index, payment in enumerate(payments):
            payment_tree.insert(
                "", "end", values=(payment["payment_date"], money(payment["amount"]), payment["note"] or "—"),
                tags=("even" if index % 2 == 0 else "odd",),
            )

        totals = ttk.Frame(dialog, padding=(18, 8))
        totals.pack(fill="x")
        ttk.Label(totals, text=f'الإجمالي: {money(sale["total"])}', font=("Tahoma", 10, "bold")).pack(side="right", padx=10)
        ttk.Label(totals, text=f'المدفوع: {money(sale["paid"])}').pack(side="right", padx=10)
        ttk.Label(totals, text=f'المتبقي: {money(sale["due"])}', foreground=RED).pack(side="right", padx=10)

        buttons = ttk.Frame(dialog, padding=12)
        buttons.pack(fill="x")
        ttk.Button(
            buttons, text="طباعة الفاتورة", style="Primary.TButton",
            command=lambda: self.print_invoice(sale_id),
        ).pack(side="right", padx=4)
        if sale["customer_id"]:
            ttk.Button(
                buttons, text="إرسال واتساب",
                command=lambda: self.send_invoice_whatsapp(sale_id, dialog),
            ).pack(side="right", padx=4)
        if sale["due"] > 0.009:
            ttk.Button(
                buttons, text="تسجيل تحصيل / قسط",
                command=lambda: self.record_payment_for_sale(sale_id, dialog),
            ).pack(side="right", padx=4)
        ttk.Button(buttons, text="إغلاق", command=dialog.destroy).pack(side="left", padx=4)
        self._center_dialog(dialog)

    def record_payment_for_sale(self, sale_id, parent=None):
        self.record_payment(sale_id=sale_id, parent=parent)

    def on_product_selected(self, _event=None):
        product = self.product_map.get(self.invoice_product.get())
        if product:
            self.invoice_price.delete(0, "end")
            self.invoice_price.insert(0, f'{product["sale_price"]:g}')

    def add_cart_item(self):
        product = self.product_map.get(self.invoice_product.get())
        if not product:
            messagebox.showwarning("تنبيه", "اختر منتجًا.", parent=self)
            return
        try:
            quantity = float(self.invoice_quantity.get())
            price = float(self.invoice_price.get())
            if quantity <= 0 or price < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("بيانات غير صحيحة", "أدخل كمية موجبة وسعرًا صحيحًا.", parent=self)
            return
        existing = next((item for item in self.cart_items if item["product_id"] == product["id"]), None)
        if existing:
            existing["quantity"] += quantity
            existing["unit_price"] = price
        else:
            self.cart_items.append(
                {"product_id": product["id"], "name": product["name"], "unit": product["unit"],
                 "quantity": quantity, "unit_price": price}
            )
        self.refresh_cart()

    def refresh_cart(self):
        self.cart_tree.delete(*self.cart_tree.get_children())
        for index, item in enumerate(self.cart_items):
            total = item["quantity"] * item["unit_price"]
            self.cart_tree.insert(
                "", "end", iid=str(index),
                values=(item["name"], f'{item["quantity"]:g} {item["unit"]}',
                        money(item["unit_price"]), money(total)),
                tags=("even" if index % 2 == 0 else "odd",),
            )
        total = sum(item["quantity"] * item["unit_price"] for item in self.cart_items)
        self.invoice_total_label.config(text=f'{self._text("الإجمالي: ")}{money(total)}')

    def remove_cart_item(self):
        selection = self.cart_tree.selection()
        if selection:
            self.cart_items.pop(int(selection[0]))
            self.refresh_cart()

    def save_invoice(self):
        try:
            paid = float(self.invoice_paid.get())
            customer_id = self.customer_map.get(self.invoice_customer.get())
            sale_id = self.db.create_sale(customer_id, self.cart_items, paid)
        except (ValueError, TypeError) as error:
            messagebox.showerror("تعذر حفظ الفاتورة", str(error), parent=self)
            return
        sale, _items = self.db.sale(sale_id)
        self.cart_items.clear()
        self.invoice_paid.delete(0, "end")
        self.invoice_paid.insert(0, "0")
        self.refresh_cart()
        self.refresh_all()
        if messagebox.askyesno(
            "تم حفظ الفاتورة",
            f'تم حفظ الفاتورة {sale["invoice_number"]} بنجاح.\nهل تريد فتحها للطباعة؟',
            parent=self,
        ):
            self.print_invoice(sale_id)

    def print_selected_invoice(self, tree=None, parent=None):
        sale_id = self._selected_id(tree or self.sales_tree, "فاتورة")
        if sale_id is not None:
            self.print_invoice(sale_id)

    def send_selected_invoice_whatsapp(self, tree, parent=None):
        sale_id = self._selected_id(tree, "فاتورة")
        if sale_id is not None:
            self.send_invoice_whatsapp(sale_id, parent)

    def send_invoice_whatsapp(self, sale_id, parent=None):
        sale, _items = self.db.sale(sale_id)
        if not sale:
            messagebox.showerror("خطأ", "الفاتورة غير موجودة.", parent=parent or self)
            return
        customer = self.db.customer(sale["customer_id"]) if sale["customer_id"] else None
        if not customer or not customer["phone"].strip():
            messagebox.showerror(
                "رقم الهاتف غير متوفر",
                "أضف رقم هاتف العميل من صفحة العملاء أولًا.",
                parent=parent or self,
            )
            return
        phone_digits = "".join(
            str(unicodedata.digit(character))
            for character in customer["phone"]
            if character.isdigit()
        )
        if phone_digits.startswith("00"):
            phone_digits = phone_digits[2:]
        elif phone_digits.startswith("0"):
            phone_digits = "20" + phone_digits[1:]
        if len(phone_digits) < 10:
            messagebox.showerror(
                "رقم الهاتف غير صحيح",
                "تحقق من رقم هاتف العميل؛ يجب أن يتضمن كود الدولة.",
                parent=parent or self,
            )
            return

        invoice_folder = Path(self.db.path).resolve().parent / "Invoices"
        pdf_path = invoice_folder / f'{sale["invoice_number"]}.pdf'
        try:
            generate_invoice_pdf(self.db, sale_id, pdf_path)
        except (OSError, ValueError) as error:
            messagebox.showerror("تعذر إنشاء ملف الفاتورة PDF", str(error), parent=parent or self)
            return

        message = (
            f"فاتورة PDF من {self.db.setting('shop_name', 'محل الوميتال')}\n"
            f"رقم الفاتورة: {sale['invoice_number']}\n"
            f"المتبقي: {money(sale['due'])}\n"
            "أرفق ملف الفاتورة PDF لإكمال الإرسال."
        )
        url = f"https://wa.me/{phone_digits}?text={quote(message)}"
        browser_opened = webbrowser.open(url)
        try:
            subprocess.Popen(["explorer.exe", "/select,", str(pdf_path)])
        except OSError as error:
            messagebox.showerror(
                "تم إنشاء ملف PDF",
                f"تم حفظ الفاتورة هنا:\n{pdf_path}\nتعذر فتح مجلد الملف: {error}",
                parent=parent or self,
            )
            return
        if not browser_opened:
            messagebox.showerror(
                "تعذر فتح واتساب",
                f"تم إنشاء ملف PDF هنا:\n{pdf_path}\nلكن لم يتمكن البرنامج من فتح المتصفح.",
                parent=parent or self,
            )
            return
        messagebox.showinfo(
            "الفاتورة جاهزة للإرسال",
            "تم إنشاء ملف PDF وفتح محادثة العميل.\n"
            "أرفق ملف الفاتورة المحدد من مجلد Invoices داخل واتساب، ثم اضغط إرسال.",
            parent=parent or self,
        )

    def export_excel_dialog(self):
        dialog = tk.Toplevel(self)
        dialog.title("تصدير الفواتير إلى Excel")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=BG)
        today = date.today()
        defaults = {
            "start": today.replace(day=1).isoformat(),
            "end": today.isoformat(),
        }
        fields = {}
        for row, (key, label) in enumerate(
            (("start", "من تاريخ (YYYY-MM-DD)"), ("end", "إلى تاريخ (YYYY-MM-DD)"))
        ):
            ttk.Label(dialog, text=label).grid(row=row, column=1, padx=12, pady=8, sticky="e")
            entry = ttk.Entry(dialog, width=22, justify="right")
            entry.insert(0, defaults[key])
            entry.grid(row=row, column=0, padx=12, pady=8)
            fields[key] = entry

        def export():
            try:
                start_date = date.fromisoformat(fields["start"].get().strip())
                end_date = date.fromisoformat(fields["end"].get().strip())
                if end_date < start_date:
                    raise ValueError("تاريخ النهاية يجب أن يكون بعد تاريخ البداية.")
            except ValueError as error:
                messagebox.showerror("تاريخ غير صحيح", str(error), parent=dialog)
                return

            output_path = filedialog.asksaveasfilename(
                parent=dialog,
                title="حفظ ملف Excel",
                initialfile=f"الفواتير_{start_date.isoformat()}_{end_date.isoformat()}.xlsx",
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx")],
            )
            if not output_path:
                return
            try:
                result = export_invoices(self.db, output_path, start_date, end_date)
            except (OSError, ValueError) as error:
                messagebox.showerror("تعذر تصدير الفواتير", str(error), parent=dialog)
                return
            dialog.destroy()
            messagebox.showinfo(
                "تم التصدير",
                f'تم تصدير {result["invoice_count"]} فاتورة إلى:\n{result["path"]}',
                parent=self,
            )

        buttons = ttk.Frame(dialog, padding=10)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Button(buttons, text="تصدير", style="Primary.TButton", command=export).pack(side="right", padx=4)
        ttk.Button(buttons, text="إلغاء", command=dialog.destroy).pack(side="left", padx=4)
        dialog.bind("<Return>", lambda _event: export())
        self._center_dialog(dialog)

    def edit_alert_settings(self):
        dialog = tk.Toplevel(self)
        dialog.title("إعدادات تنبيهات المتابعة")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=BG)
        threshold = ttk.Entry(dialog, width=22, justify="right")
        threshold.insert(0, self.db.setting("alert_threshold", "5000"))
        ttk.Label(dialog, text="حد المبلغ المتبقي (ج.م):").grid(row=0, column=1, padx=12, pady=8, sticky="e")
        threshold.grid(row=0, column=0, padx=12, pady=8)
        inactive_days = ttk.Entry(dialog, width=22, justify="right")
        inactive_days.insert(0, self.db.setting("alert_days", "10"))
        ttk.Label(dialog, text="أيام دون تحصيل:").grid(row=1, column=1, padx=12, pady=8, sticky="e")
        inactive_days.grid(row=1, column=0, padx=12, pady=8)

        def save():
            try:
                threshold_value = float(threshold.get())
                days_value = int(inactive_days.get())
                if threshold_value < 0 or days_value < 1:
                    raise ValueError("أدخل حدًا غير سالب وعدد أيام أكبر من صفر.")
            except ValueError:
                messagebox.showerror(
                    "إعدادات غير صحيحة",
                    "أدخل مبلغًا صحيحًا غير سالب وعدد أيام صحيحًا أكبر من صفر.",
                    parent=dialog,
                )
                return
            self.db.set_setting("alert_threshold", str(threshold_value))
            self.db.set_setting("alert_days", str(days_value))
            dialog.destroy()
            self.refresh_alerts()

        ttk.Button(dialog, text="حفظ", style="Primary.TButton", command=save).grid(
            row=2, column=0, padx=12, pady=12, sticky="w"
        )
        dialog.bind("<Return>", lambda _event: save())
        self._center_dialog(dialog)

    def print_invoice(self, sale_id):
        sale, items = self.db.sale(sale_id)
        if not sale:
            messagebox.showerror("خطأ", "الفاتورة غير موجودة.", parent=self)
            return
        rows = "".join(
            "<tr><td>{}</td><td>{:g} {}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape(item["product_name"]), item["quantity"], html.escape(item["unit"]),
                money(item["unit_price"]), money(item["quantity"] * item["unit_price"]),
            )
            for item in items
        )
        shop = html.escape(self.db.setting("shop_name", "محل الوميتال"))
        document = f"""<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8">
<title></title>
<style>
@page{{size:A4;margin:0}}
*{{box-sizing:border-box}}
html,body{{margin:0;padding:0}}
body{{font-family:Tahoma,Arial,sans-serif;color:#222;background:#fff}}
.invoice{{padding:14mm;max-width:210mm;margin:0 auto}}
h1{{font-size:24px;margin:0 0 6px}}h2{{font-size:19px;margin:16px 0 8px}}
.meta{{color:#444;line-height:1.9;margin:8px 0 22px}}
table{{width:100%;border-collapse:collapse}}
th,td{{border:1px solid #bbb;padding:10px;text-align:right}}
th{{background:#f0f3f6}}
.totals{{width:360px;max-width:100%;margin:20px 0 0 auto}}
.totals div{{padding:8px;border-bottom:1px solid #ddd}}
@media print{{.invoice{{width:210mm;min-height:297mm;margin:0;padding:14mm}}}}
</style><main class="invoice"><h1>{shop}</h1><h2>فاتورة بيع</h2>
<div class="meta">رقم الفاتورة: {html.escape(sale['invoice_number'])} &nbsp; | &nbsp; التاريخ: {html.escape(sale['created_at'])}
<br>العميل: {html.escape(sale['customer_name'])}</div>
<table><thead><tr><th>المنتج</th><th>الكمية</th><th>سعر الوحدة</th><th>الإجمالي</th></tr></thead>
<tbody>{rows}</tbody></table>
<div class="totals"><div>الإجمالي: {money(sale['total'])}</div>
<div>المدفوع: {money(sale['paid'])}</div><div><b>المتبقي: {money(sale['due'])}</b></div></div>
<script>window.onload=()=>setTimeout(()=>window.print(),300)</script></main></html>"""
        output = default_database_path().parent / f"{sale['invoice_number']}.html"
        output.write_text(document, encoding="utf-8")
        webbrowser.open(output.resolve().as_uri())

    def print_customer_statement(self, customer_id):
        customer = self.db.customer(customer_id)
        if not customer:
            messagebox.showerror("خطأ", "العميل غير موجود.", parent=self)
            return

        sales = self.db.customer_sales(customer_id)
        total_sales = sum(sale["total"] for sale in sales)
        total_paid = sum(sale["paid"] for sale in sales)
        total_due = sum(sale["due"] for sale in sales)
        transactions = []
        for sale in reversed(sales):
            _sale, items = self.db.sale(sale["id"])
            item_rows = "".join(
                "<tr><td>{}</td><td>{:g} {}</td><td>{}</td><td>{}</td></tr>".format(
                    html.escape(item["product_name"]),
                    item["quantity"],
                    html.escape(item["unit"]),
                    money(item["unit_price"]),
                    money(item["quantity"] * item["unit_price"]),
                )
                for item in items
            )
            payments = self.db.sale_payments(sale["id"])
            payment_rows = "".join(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                    html.escape(payment["payment_date"]),
                    money(payment["amount"]),
                    html.escape(payment["note"] or "—"),
                )
                for payment in payments
            )
            if not payment_rows:
                payment_rows = '<tr><td colspan="3" class="empty">لا توجد دفعات مسجلة</td></tr>'
            transactions.append(
                f"""<section class="sale">
<h3>فاتورة {html.escape(sale['invoice_number'])}
<span>التاريخ: {html.escape(sale['created_at'])}</span></h3>
<table><thead><tr><th>المنتج</th><th>الكمية</th><th>سعر الوحدة</th><th>الإجمالي</th></tr></thead>
<tbody>{item_rows}</tbody></table>
<table class="payments"><thead><tr><th>تاريخ التحصيل</th><th>المبلغ</th><th>البيان</th></tr></thead>
<tbody>{payment_rows}</tbody></table>
<div class="sale-totals">إجمالي الفاتورة: {money(sale['total'])}
&nbsp; | &nbsp; المدفوع: {money(sale['paid'])}
&nbsp; | &nbsp; المتبقي: <b>{money(sale['due'])}</b></div>
</section>"""
            )

        transaction_html = "".join(transactions) or '<p class="empty">لا توجد مشتريات مسجلة لهذا العميل.</p>'
        shop = html.escape(self.db.setting("shop_name", "محل الوميتال"))
        customer_name = html.escape(customer["name"])
        phone = html.escape(customer["phone"] or "—")
        address = html.escape(customer["address"] or "—")
        document = f"""<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8">
<title></title><style>
@page{{size:A4;margin:0}}
*{{box-sizing:border-box}}
html,body{{margin:0;padding:0}}
body{{font-family:Tahoma,Arial,sans-serif;color:#222;background:#fff}}
.document{{padding:14mm;max-width:210mm;margin:0 auto}}
h1{{font-size:23px;margin:0 0 5px}}h2{{font-size:19px;margin:12px 0}}
.meta{{line-height:1.9;color:#444;margin:8px 0 16px}}
.summary{{display:flex;justify-content:space-between;gap:8px;margin:12px 0 20px}}
.summary div{{border:1px solid #bbb;padding:10px;flex:1;text-align:center}}
.sale{{margin:18px 0 22px;break-inside:avoid-page}}
.sale h3{{font-size:14px;margin:0 0 8px;padding:8px;background:#edf1f5}}
.sale h3 span{{float:left;font-size:12px;font-weight:normal}}
table{{width:100%;border-collapse:collapse;margin:6px 0}}
th,td{{border:1px solid #bbb;padding:7px;text-align:right;font-size:12px}}
th{{background:#f0f3f6}}
.payments{{width:80%;margin-right:auto;margin-top:8px}}
.sale-totals{{text-align:left;padding:7px 0;border-bottom:1px solid #ddd;font-size:12px}}
.empty{{color:#666;text-align:center;padding:10px}}
@media print{{.document{{width:210mm;margin:0;padding:14mm}}}}
</style><main class="document"><h1>{shop}</h1><h2>كشف مشتريات وحساب عميل</h2>
<div class="meta">العميل: {customer_name}<br>الهاتف: {phone}<br>العنوان: {address}</div>
<div class="summary"><div>عدد الفواتير<br><b>{len(sales)}</b></div>
<div>إجمالي المشتريات<br><b>{money(total_sales)}</b></div>
<div>إجمالي المدفوع<br><b>{money(total_paid)}</b></div>
<div>إجمالي المبلغ المتبقي<br><b>{money(total_due)}</b></div></div>
{transaction_html}</main><script>window.onload=()=>setTimeout(()=>window.print(),300)</script></html>"""
        output = default_database_path().parent / f"كشف-حساب-عميل-{customer_id}.html"
        output.write_text(document, encoding="utf-8")
        webbrowser.open(output.resolve().as_uri())

    def record_payment(self, tree=None, sale_id=None, parent=None):
        tree = tree or self.debts_tree
        if sale_id is None:
            sale_id = self._selected_id(tree, "فاتورة")
            if sale_id is None:
                return
        sale, _items = self.db.sale(sale_id)
        if not sale:
            messagebox.showerror("خطأ", "الفاتورة غير موجودة.", parent=parent or self)
            return
        if sale["due"] <= 0.009:
            messagebox.showinfo("الفاتورة مسددة", "لا يوجد مبلغ متبقٍ على هذه الفاتورة.", parent=parent or self)
            return
        amount = simpledialog.askfloat(
            "تسجيل تحصيل", f'المتبقي على الفاتورة {money(sale["due"])}\nأدخل مبلغ التحصيل:',
            minvalue=0.01, parent=parent or self,
        )
        if amount is None:
            return
        note = simpledialog.askstring("بيان القسط", "ملاحظة (اختياري):", parent=parent or self) or ""
        try:
            self.db.add_payment(sale_id, amount, note)
        except ValueError as error:
            messagebox.showerror("تعذر تسجيل التحصيل", str(error), parent=parent or self)
            return
        self.refresh_all()
        if parent and parent.winfo_exists():
            parent.destroy()
        self.show_sale_details(sale_id)
        messagebox.showinfo("تم", "تم تسجيل التحصيل بنجاح.", parent=self)

    def edit_shop_name(self):
        name = simpledialog.askstring(
            "بيانات المحل", "اسم المحل الذي سيظهر على الفاتورة:",
            initialvalue=self.db.setting("shop_name", "محل الوميتال"), parent=self,
        )
        if name and name.strip():
            self.db.set_setting("shop_name", name)

    def change_database_location(self):
        current_path = Path(self.db.path).resolve()
        selected_path = filedialog.asksaveasfilename(
            parent=self,
            title="اختيار مكان قاعدة البيانات",
            initialdir=str(current_path.parent),
            initialfile=current_path.name,
            defaultextension=".db",
            filetypes=[("قاعدة بيانات البرنامج", "*.db"), ("كل الملفات", "*.*")],
        )
        if not selected_path:
            return
        target_path = Path(selected_path).resolve()
        if target_path == current_path:
            messagebox.showinfo("مكان قاعدة البيانات", "قاعدة البيانات محفوظة بالفعل في هذا المكان.", parent=self)
            return

        target_exists = target_path.exists()
        if target_exists and not messagebox.askyesno(
            "قاعدة بيانات موجودة",
            "الملف المحدد يحتوي بالفعل على قاعدة بيانات.\n"
            "هل تريد التبديل إليها؟ لن يتم نسخ بيانات البرنامج الحالية إليها.",
            parent=self,
        ):
            return

        candidate = None
        try:
            if not target_exists:
                self.db.backup_to(target_path)
            candidate = Database(target_path)
            save_configured_database_path(target_path)
        except (OSError, sqlite3.Error, ValueError) as error:
            if candidate:
                candidate.close()
            messagebox.showerror("تعذر تغيير مكان قاعدة البيانات", str(error), parent=self)
            return

        self.db.close()
        self.db = candidate
        self.refresh_all()
        confirmation = "تم التبديل إلى قاعدة البيانات الموجودة في:" if target_exists else "تم نسخ بياناتك واستخدام قاعدة البيانات الجديدة في:"
        messagebox.showinfo(
            "تم تغيير مكان قاعدة البيانات",
            f"{confirmation}\n{target_path}",
            parent=self,
        )

    def _center_dialog(self, dialog):
        self._apply_locale(dialog)
        dialog.title(self._text(dialog.title()))
        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def on_close(self):
        global CURRENT_LANGUAGE
        CURRENT_LANGUAGE = self.language
        for module, names in (
            (messagebox, ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel")),
            (simpledialog, ("askstring", "askinteger", "askfloat")),
        ):
            for name in names:
                original_name = f"_alumetal_original_{name}"
                if hasattr(module, original_name):
                    setattr(module, name, getattr(module, original_name))
        for name in ("askopenfilename", "asksaveasfilename", "askdirectory"):
            original_name = f"_alumetal_original_{name}"
            if hasattr(filedialog, original_name):
                setattr(filedialog, name, getattr(filedialog, original_name))
        self.db.close()
        self.destroy()
if __name__ == "__main__":
    ManagerApp().mainloop()
