import pdfplumber
import pandas as pd
import arabic_reshaper
from bidi.algorithm import get_display
import os
import re
import unicodedata
import threading
import queue
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from decimal import Decimal
from payroll_core import parse_text_pages, standard, export_amount, EARNING, DEDUCTION, INFORMATIONAL
from payroll_excel import create_excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

"""
النسخة الذكية: مسح تلقائي للبنود + قوائم منسدلة + تصدير إكسيل احترافي
المتطلبات:
pip install pdfplumber pandas arabic-reshaper python-bidi openpyxl
"""

def to_hindi_nums(text):
    if not text: return ""
    hindi_map = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")
    return str(text).translate(hindi_map)

def to_english_nums(text):
    if not text: return ""
    eng_map = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    return str(text).translate(eng_map)

def parse_amount_to_float(val):
    """
    تحويل أي قيمة مالية (أرقام هندية/عربية، فواصل آلاف، فواصل عشرية) 
    إلى رقم عشري حقيقي (float) لتتمكن برامج مثل Excel من إجراء الحسابات عليه
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    # تحويل الأرقام المشرقية ٠١٢٣٤٥٦٧٨٩ إلى أرقام إنجليزية
    hindi_to_eng = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    s = s.translate(hindi_to_eng)
    # إزالة فواصل الآلاف
    s = s.replace(",", "").replace("،", "").replace(" ", "")
    # تحويل فاصلة الكسر العشرية العربية ٫ إلى نقطة .
    s = s.replace("٫", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0

def decode_oracle_text(text):
    """
    فك الترتيب البصري المعكوس لنصوص أوراكل وتطبيعها إلى نص عربي قياسي سليم
    مع الحفاظ على الأرقام والنسب المئوية والكلمات الإنجليزية دون تشويه
    """
    if not text: return ""
    lines = str(text).split("\n")
    cleaned_lines = []
    for line in lines:
        line = line.strip()
        if not line: continue
        # 1. عكس السطر لإعادة ترتيب الكلمات والحروف للترتيب المنطقي
        rev = line[::-1]
        # 2. إعادة عكس الأرقام (إنجليزية أو هندية) مع الفواصل العشرية
        rev = re.sub(r'[\d٠-٩]+(?:[.,٫][\d٠-٩]+)?', lambda m: m.group(0)[::-1], rev)
        # 3. تصحيح علامة النسبة المئوية % إذا تغير موضعها
        rev = re.sub(r'%([\d٠-٩]+(?:[.,٫][\d٠-٩]+)?)', r'\1%', rev)
        # 4. إعادة عكس الكلمات اللاتينية/الإنجليزية إن وجدت
        rev = re.sub(r'[A-Za-z]{2,}', lambda m: m.group(0)[::-1], rev)
        # 5. تطبيع أشكال الحروف التشكيلية (Presentation Forms) لحروف عربية قياسية
        norm = unicodedata.normalize('NFKC', rev)
        cleaned_lines.append(norm.strip())
    # دمج الأسطر بمسافة واحدة وتنظيف المسافات الزائدة
    res = " ".join(cleaned_lines)
    return re.sub(r'\s+', ' ', res).strip()

def normalize(text):
    if not text: return ""
    text = str(text).strip()
    hindi_to_eng = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    text = text.translate(hindi_to_eng)
    text = "".join(text.split())
    mapping = {"أ": "ا", "إ": "ا", "آ": "ا", "ة": "ه", "ى": "ي"}
    for k, v in mapping.items():
        text = text.replace(k, v)
    return text.lower()

def is_val(text):
    """فحص ما إذا كانت الخلية تحتوي على قيمة مالية صالحة وتجنب الأكواد الحسابية"""
    if not text: return False
    # أي خلية تحتوي على حروف (عربية أو لاتينية) هي اسم بند وليست مبلغاً
    if re.search(r'[A-Za-z\u0621-\u064A\uFB50-\uFDFF\uFE70-\uFEFC]', str(text)):
        return False
    s_text = str(text).replace(",", "").replace("،", "").strip()
    hindi_to_eng = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    s_text = s_text.translate(hindi_to_eng)
    # استبعاد الأكواد المالية الحكومية ذات الـ 8 أرقام (مثل 21110105 أو 41110201)
    if len(s_text) == 8 and s_text.isdigit():
        return False
    digits = [c for c in s_text if c.isdigit()]
    if not digits:
        return False
    return ("." in s_text or "٫" in s_text) or len(digits) <= 7

def is_ignored_item(clean_text):
    """استبعاد العناوين والإجماليات والمجاميع الفرعية وصافي الراتب من قائمة البنود"""
    # استبعاد البنود التجميعية التي تبدأ بـ x لأنها مجاميع فرعية لبنود موجودة بالفعل في الجدول وتتسبب في مضاعفة المبالغ
    if clean_text.strip().lower().startswith("x"):
        return True
    norm = normalize(clean_text)
    ignore_keywords = [
        "اجمالي", "إجمالي", "اجمالى", "إجمالى", "صافي", "صافى", 
        "الاسم", "القيمة", "الكود", "بيان", "الاستقطاعات", "الاستحقاقات",
        "مخاطبة", "صرف", "ملاحظات", "فقط", "تعتمد"
    ]
    return any(kw in norm for kw in ignore_keywords)

UNKNOWN = "غير معروف"

def extract_emp_info(text):
    """استخراج اسم وكود الموظف من نص خام (ترويسة الصفحة)"""
    text_std = unicodedata.normalize('NFKC', text or "")
    text_norm = text_std.replace("\n", " ").translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))

    emp_code = UNKNOWN
    m = re.search(r"\b\d{5,10}\s*-\s*(\d{3,6})\b", text_norm)
    if m:
        emp_code = m.group(1)
    else:
        m = re.search(r"\b(\d{3,6})\s*-\s*\d{5,10}\b", text_norm)
        if m:
            emp_code = m.group(1)

    emp_name = UNKNOWN
    for line in text_std.split("\n")[:30]:
        line_reversed = line[::-1]
        if "اسم" in line_reversed or "موظف" in line_reversed:
            for keyword in ["اسم الموظف:", "اسم الموظف", "الموظف:", "الموظف", "اسم:", "اسم"]:
                if keyword in line_reversed:
                    raw_name = line_reversed.split(keyword)[-1]
                    for stop_word in ["رقم", "قومى", "قومي", "نوع", "درج", "ادار", "فرع", "بنك", "تأمين"]:
                        if stop_word in raw_name:
                            raw_name = raw_name.split(stop_word)[0]
                    cleaned = "".join([c for c in raw_name if not c.isdigit() and c not in [":", "-", "_", "*"]]).strip()
                    if len(cleaned) >= 3:
                        emp_name = cleaned
                    break
            if emp_name != UNKNOWN:
                break
    return emp_name, emp_code

def parse_page(page):
    """Read page text only; grouping and item identification require document context."""
    text = page.extract_text() or ""
    # PDFium / pdfplumber may expose presentation-order Arabic lines reversed.
    # Only decode when explicit Oracle labels/sections become more recognizable.
    def semantic_score(value):
        normalized = standard(value)
        return sum(normalized.count(token) for token in (
            "اسم الموظف", "رقم الصرف", "الاستحقاقات", "الاستقطاعات",
            "الصافي", "الصافى", "حكومة", "اجمالي"))
    visual="\n".join(decode_oracle_text(line) for line in text.splitlines())
    if semantic_score(visual)>semantic_score(text):
        text=visual
    # Table extraction is a fallback when the page has no readable Unicode text.
    # We do not infer financial values from numeric cells purely by their position.
    if not text.strip():
        rows = page.extract_tables() or []
        text = "\n".join(" ".join(str(c or "") for c in row) for table in rows for row in table)
    return {"raw_text": text, "page": page.page_number}


def parse_page_range(path, start, end):
    """تُنفذ داخل عملية منفصلة: تحليل مجموعة صفحات"""
    out = []
    with pdfplumber.open(path) as pdf:
        for idx in range(start, end):
            page = pdf.pages[idx]
            out.append(parse_page(page))
            page.flush_cache()
    return out

class SmartPayrollApp:
    def __init__(self, root):
        self.root = root
        self.root.title("مستخرج رواتب أوراكل الذكي - متعدد الأنماط")
        self.root.geometry("1180x760")
        
        self.pdf_path = ""
        self.items_db = {"الاستقطاعات": [], "الاستحقاقات": []}
        self.pages_data = []
        self.current_export_info = None
        self.parsed_employees = []
        self.msg_queue = queue.Queue()
        
        self.setup_ui()

    def setup_ui(self):
        # 1. الجزء العلوي: اختيار الملف والمسح
        top = tk.Frame(self.root, pady=10)
        top.pack(fill=tk.X)
        tk.Button(
            top, 
            text="1. اختر ملف PDF للبدء بالمسح", 
            command=self.select_file, 
            bg="#2196F3", 
            fg="white", 
            font=("Segoe UI", 10, "bold"), 
            padx=15, 
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT, padx=10)

        self.lbl_file = tk.Label(top, text="بانتظار اختيار الملف...", fg="gray", font=("Segoe UI", 10))
        self.lbl_file.pack(side=tk.RIGHT, padx=10)
        self.progress = ttk.Progressbar(top, mode="determinate", length=250)
        self.progress.pack(side=tk.LEFT, padx=10)

        # 2. تخصيص نمط الاستخراج
        self.sel_frame = tk.LabelFrame(self.root, text="2. تخصيص نمط الاستخراج", padx=12, pady=10, font=("Segoe UI", 10, "bold"))
        self.sel_frame.pack(fill=tk.X, padx=10, pady=5)

        # سطر الإعدادات (الصف الأول داخل الإطار)
        row1 = tk.Frame(self.sel_frame)
        row1.pack(fill=tk.X, pady=4)

        tk.Label(row1, text="نمط الاستخراج:", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)
        self.mode_options = [
            "بند مالي واحد محدد",
            "كافة بنود الاستحقاقات (المستحق)",
            "كافة بنود الاستقطاعات (المستقطع)",
            "الكشف الشامل (المستحق والمستقطع والصافي)"
        ]
        self.combo_mode = ttk.Combobox(row1, values=self.mode_options, state="readonly", width=34, font=("Segoe UI", 9))
        self.combo_mode.set(self.mode_options[0])
        self.combo_mode.pack(side=tk.RIGHT, padx=8)
        self.combo_mode.bind("<<ComboboxSelected>>", self.on_mode_change)

        tk.Label(row1, text="النوع:", font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=(15, 5))
        self.combo_cat = ttk.Combobox(row1, values=["الاستقطاعات", "الاستحقاقات"], state="readonly", width=14, font=("Segoe UI", 9))
        self.combo_cat.set("الاستقطاعات")
        self.combo_cat.pack(side=tk.RIGHT, padx=8)
        self.combo_cat.bind("<<ComboboxSelected>>", self.update_item_list)

        tk.Label(row1, text="البند المالي:", font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=(15, 5))
        self.combo_item = ttk.Combobox(row1, state="disabled", width=35, font=("Segoe UI", 9))
        self.combo_item.pack(side=tk.RIGHT, padx=8)

        # سطر الإجراء والشرح (الصف الثاني داخل الإطار)
        row2 = tk.Frame(self.sel_frame)
        row2.pack(fill=tk.X, pady=(6, 2))

        self.btn_extract = tk.Button(
            row2,
            text="⚡ تنفيذ الاستخراج وعرض البيانات",
            command=self.process_selected,
            bg="#FF9800",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=18,
            pady=3,
            cursor="hand2"
        )
        self.btn_extract.pack(side=tk.LEFT, padx=10)

        self.lbl_mode_desc = tk.Label(row2, text="يستخرج جدولاً للموظفين مع قيمة البند المحدد فقط وحساب الإجمالي.", fg="#455A64", font=("Segoe UI", 9))
        self.lbl_mode_desc.pack(side=tk.RIGHT, padx=10)

        # 3. جدول عرض البيانات في الواجهة مع شريطي تمرير عمودي وأفقي
        tree_container = tk.Frame(self.root)
        tree_container.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        self.tree = ttk.Treeview(tree_container, show='headings')
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        # الإعداد المبدئي لأعمدة الجدول
        self.cols = ("م", "كود الموظف", "اسم الموظف", "المبلغ")
        self.tree["columns"] = self.cols
        for c in self.cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor=tk.CENTER, width=150)

        # 4. التصدير
        self.btn_xlsx = tk.Button(
            self.root,
            text="3. تصدير التقرير المالي إلى Excel",
            command=self.export,
            state=tk.DISABLED,
            bg="#2E7D32",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            pady=10,
            cursor="hand2"
        )
        self.btn_xlsx.pack(fill=tk.X, padx=10, pady=10)

    def on_mode_change(self, event=None):
        mode = self.combo_mode.get()
        if mode == "بند مالي واحد محدد":
            self.combo_cat.config(state="readonly")
            if self.pages_data and self.items_db.get(self.combo_cat.get()):
                self.combo_item.config(state="readonly")
            else:
                self.combo_item.config(state="disabled")
            self.lbl_mode_desc.config(text="يستخرج جدولاً للموظفين مع قيمة البند المحدد وحساب الإجمالي.")
        elif mode == "كافة بنود الاستحقاقات (المستحق)":
            self.combo_cat.config(state="disabled")
            self.combo_item.config(state="disabled")
            ent_count = len(self.items_db.get("الاستحقاقات", []))
            self.lbl_mode_desc.config(text=f"كشف تفصيلي بكافة بنود الاستحقاقات ({ent_count} بند) وإجمالي المستحق لكل موظف.")
        elif mode == "كافة بنود الاستقطاعات (المستقطع)":
            self.combo_cat.config(state="disabled")
            self.combo_item.config(state="disabled")
            ded_count = len(self.items_db.get("الاستقطاعات", []))
            self.lbl_mode_desc.config(text=f"كشف تفصيلي بكافة بنود الاستقطاعات ({ded_count} بند) وإجمالي المستقطع لكل موظف.")
        elif mode == "الكشف الشامل (المستحق والمستقطع والصافي)":
            self.combo_cat.config(state="disabled")
            self.combo_item.config(state="disabled")
            ent_count = len(self.items_db.get("الاستحقاقات", []))
            ded_count = len(self.items_db.get("الاستقطاعات", []))
            self.lbl_mode_desc.config(text=f"الكشف الشامل: {ent_count} استحقاق + {ded_count} استقطاع + إجمالي المستحق وإجمالي المستقطع وصافي الراتب.")

    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.pdf_path = p
            self.pages_data = []
            self.current_export_info = None
            self.lbl_file.config(text=f"جاري مسح: {os.path.basename(p)}...", fg="blue")
            self.combo_item.config(state="disabled")
            self.btn_xlsx.config(state=tk.DISABLED)
            self.progress["value"] = 0
            self.msg_queue = queue.Queue()
            threading.Thread(target=self._parse_worker, args=(p, self.msg_queue), daemon=True).start()
            self.root.after(100, self._poll_queue)

    def _parse_worker(self, path, q):
        """تحليل كل صفحات الملف مرة واحدة فقط بالتوازي على عدة أنوية"""
        try:
            with pdfplumber.open(path) as pdf:
                total = len(pdf.pages)
            chunk = 10
            ranges = [(s, min(s + chunk, total)) for s in range(0, total, chunk)]
            workers = max(1, min((os.cpu_count() or 2) - 1, len(ranges), 12))
            results = {}
            done = 0
            with ProcessPoolExecutor(max_workers=workers) as ex:
                futs = {ex.submit(parse_page_range, path, s, e): s for s, e in ranges}
                for f in as_completed(futs):
                    s = futs[f]
                    results[s] = f.result()
                    done += len(results[s])
                    q.put(("progress", done, total))
            pages = []
            for s in sorted(results):
                pages.extend(results[s])
            q.put(("done", pages))
        except Exception as e:
            q.put(("error", str(e)))

    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg[0] == "progress":
                    _, done, total = msg
                    self.progress["maximum"] = total
                    self.progress["value"] = done
                    self.lbl_file.config(text=f"جاري مسح الملف: {done} من {total} صفحة...", fg="blue")
                elif msg[0] == "done":
                    self.pages_data = msg[1]
                    self._on_parse_done()
                    return
                elif msg[0] == "error":
                    self.lbl_file.config(text="فشل المسح", fg="red")
                    messagebox.showerror("خطأ في المسح", f"حدث خطأ أثناء قراءة الملف:\n{msg[1]}")
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_parse_done(self):
        """بناء قوائم البنود من البيانات المخزنة في الذاكرة"""
        items_set = {"الاستقطاعات": set(), "الاستحقاقات": set()}
        self.parsed_employees = self.get_unified_employees(self.pages_data)
        # Export is available immediately after document analysis; preview modes
        # are optional and must never determine which employees are exported.
        self.btn_xlsx.config(state=tk.NORMAL if self.parsed_employees else tk.DISABLED)
        for emp in self.parsed_employees:
            for (cat, account, label) in emp["items"]:
                if cat in items_set:
                    items_set[cat].add(f"{account} | {label}")
        for cat in items_set:
            self.items_db[cat] = sorted(items_set[cat])

        total_pages = len(self.pages_data)
        total_items = len(self.items_db["الاستقطاعات"]) + len(self.items_db["الاستحقاقات"])
        self.lbl_file.config(text=f"تم المسح: {os.path.basename(self.pdf_path)} ({total_pages} صفحة - {total_items} بند)", fg="green")
        self.combo_cat.set("الاستقطاعات")
        self.update_item_list()
        self.on_mode_change()
        if total_items == 0:
            messagebox.showwarning("تنبيه", "لم يتم العثور على أي بنود. تأكد أن الملف يحتوي على جداول نصوص وليس صورا.")
        else:
            messagebox.showinfo("تم المسح بنجاح", f"تم العثور على {total_items} بند مالي مختلف ({len(self.items_db['الاستحقاقات'])} استحقاق و {len(self.items_db['الاستقطاعات'])} استقطاع).\nيمكنك الآن اختيار نمط الاستخراج المطلوب وعرضه أو تصديره فوراً.")

    def update_item_list(self, event=None):
        cat = self.combo_cat.get()
        if cat in self.items_db and self.items_db[cat]:
            items = self.items_db[cat]
            self.combo_item.config(state="readonly", values=items)
            self.combo_item.set(items[0])
        else:
            self.combo_item.config(values=[])
            self.combo_item.set("لا توجد بنود")
            self.combo_item.config(state="disabled")

    def get_unified_employees(self, pages_data):
        """Build employees from the complete ordered text stream, not page-local guesses."""
        records = parse_text_pages([p.get("raw_text", "") for p in pages_data])
        for emp in records:
            # Preserve the legacy display shape while retaining structured identity.
            emp["display_items"] = {(kind, label): value for (kind, code, label), value in emp["items"].items()}
        return records

    def fix_ar(self, t):
        if not t: return ""
        try: return get_display(arabic_reshaper.reshape(str(t).strip()))
        except: return str(t)

    def process_selected(self):
        """استخراج فوري من البيانات المخزنة وتحديث الجدول وفق النمط المختار"""
        if not self.pages_data:
            messagebox.showwarning("تنبيه", "يرجى اختيار ومسح ملف PDF أولاً.")
            return

        mode = self.combo_mode.get()
        # مسح محتويات الجدول القديمة
        for i in self.tree.get_children():
            self.tree.delete(i)

        unified_pages = getattr(self, "parsed_employees", None) or self.get_unified_employees(self.pages_data)
        self.current_employees = unified_pages
        def lookup(pg, cat, item):
            matches = [(k, v) for k, v in pg["items"].items()
                       if k[0] == cat and f"{k[1]} | {k[2]}" == item]
            if not matches: return None if pg["status"] != "Verified" else Decimal(0)
            return sum((v for _, v in matches), Decimal(0))

        if mode == "بند مالي واحد محدد":
            target_item = self.combo_item.get()
            category = self.combo_cat.get()
            if not target_item or target_item == "لا توجد بنود":
                messagebox.showwarning("تنبيه", "يرجى اختيار البند المراد استخراجه.")
                return

            key = (category, target_item)
            cols = ("م", "كود الموظف", "اسم الموظف", "المبلغ")
            self.tree["columns"] = cols
            self.tree.heading("م", text="م")
            self.tree.column("م", anchor=tk.CENTER, width=50)
            self.tree.heading("كود الموظف", text="كود الموظف")
            self.tree.column("كود الموظف", anchor=tk.CENTER, width=120)
            self.tree.heading("اسم الموظف", text="اسم الموظف")
            self.tree.column("اسم الموظف", anchor=tk.E, width=300)
            self.tree.heading("المبلغ", text=f"المبلغ ({target_item})")
            self.tree.column("المبلغ", anchor=tk.E, width=160)

            results = []
            for pg in unified_pages:
                amount_val = lookup(pg, category, target_item)
                if amount_val is not None:
                    amount_num = float(amount_val)
                    idx = len(results) + 1
                    self.tree.insert("", tk.END, values=(
                        idx,
                        to_hindi_nums(pg["code"]),
                        self.fix_ar(pg["name"]),
                        to_hindi_nums(f"{amount_num:,.2f}")
                    ))
                    results.append({
                        "كود الموظف": to_english_nums(pg["code"]),
                        "اسم الموظف": pg["name"],
                        "المبلغ": amount_num
                    })

            if not results:
                self.current_export_info = None
                self.btn_xlsx.config(state=tk.DISABLED)
                messagebox.showinfo("نتيجة", "لم يتم العثور على أي موظف مسجل لديه هذا البند.")
                return

            self.current_export_info = {
                "mode": "single",
                "target_item": target_item,
                "category": category,
                "data": results
            }
            total_sum = sum(r["المبلغ"] for r in results)
            self.lbl_file.config(text=f"تم استخراج {len(results)} سجل لـ [{target_item}] | الإجمالي: {total_sum:,.2f} ج.م", fg="green")
            self.btn_xlsx.config(state=tk.NORMAL)

        elif mode in ("كافة بنود الاستحقاقات (المستحق)", "كافة بنود الاستقطاعات (المستقطع)"):
            is_ent = (mode == "كافة بنود الاستحقاقات (المستحق)")
            cat_name = "الاستحقاقات" if is_ent else "الاستقطاعات"
            items_list = sorted(self.items_db.get(cat_name, []))
            if not items_list:
                messagebox.showwarning("تنبيه", f"لم يتم العثور على أي بنود في {cat_name}.")
                return

            tot_col_name = f"إجمالي {cat_name}"
            cols = ["م", "كود الموظف", "اسم الموظف"] + items_list + [tot_col_name]
            self.tree["columns"] = cols
            self.tree.heading("م", text="م")
            self.tree.column("م", anchor=tk.CENTER, width=50)
            self.tree.heading("كود الموظف", text="كود الموظف")
            self.tree.column("كود الموظف", anchor=tk.CENTER, width=100)
            self.tree.heading("اسم الموظف", text="اسم الموظف")
            self.tree.column("اسم الموظف", anchor=tk.E, width=220)

            for itm in items_list:
                self.tree.heading(itm, text=itm)
                self.tree.column(itm, anchor=tk.E, width=130)

            self.tree.heading(tot_col_name, text=f"★ {tot_col_name}")
            self.tree.column(tot_col_name, anchor=tk.E, width=140)

            for idx, pg in enumerate(unified_pages, start=1):
                row_vals = [idx, to_hindi_nums(pg["code"]), self.fix_ar(pg["name"])]
                row_sum = 0.0
                for itm in items_list:
                    amt = lookup(pg, cat_name, itm)
                    if amt is not None: row_sum += float(amt)
                    row_vals.append(to_hindi_nums(f"{amt:,.2f}") if amt is not None else "غير متاح")
                row_vals.append(to_hindi_nums(f"{row_sum:,.2f}"))
                self.tree.insert("", tk.END, values=row_vals)

            self.current_export_info = {
                "mode": "category",
                "category": cat_name,
                "items": items_list,
                "data": unified_pages
            }
            self.lbl_file.config(text=f"تم استخراج {len(unified_pages)} موظف بـ {len(items_list)} بند {cat_name}", fg="green")
            self.btn_xlsx.config(state=tk.NORMAL)

        elif mode == "الكشف الشامل (المستحق والمستقطع والصافي)":
            ent_items = sorted(self.items_db.get("الاستحقاقات", []))
            ded_items = sorted(self.items_db.get("الاستقطاعات", []))
            if not ent_items and not ded_items:
                messagebox.showwarning("تنبيه", "لم يتم العثور على أي بنود في الملف.")
                return

            cols = ["م", "كود الموظف", "اسم الموظف"] + ent_items + ["إجمالي المستحق"] + ded_items + ["إجمالي المستقطع", "صافي الراتب"]
            self.tree["columns"] = cols
            self.tree.heading("م", text="م")
            self.tree.column("م", anchor=tk.CENTER, width=50)
            self.tree.heading("كود الموظف", text="كود الموظف")
            self.tree.column("كود الموظف", anchor=tk.CENTER, width=100)
            self.tree.heading("اسم الموظف", text="اسم الموظف")
            self.tree.column("اسم الموظف", anchor=tk.E, width=220)

            for itm in ent_items:
                self.tree.heading(itm, text=f"[+] {itm}")
                self.tree.column(itm, anchor=tk.E, width=120)
            self.tree.heading("إجمالي المستحق", text="★ إجمالي المستحق")
            self.tree.column("إجمالي المستحق", anchor=tk.E, width=130)

            for itm in ded_items:
                self.tree.heading(itm, text=f"[-] {itm}")
                self.tree.column(itm, anchor=tk.E, width=120)
            self.tree.heading("إجمالي المستقطع", text="★ إجمالي المستقطع")
            self.tree.column("إجمالي المستقطع", anchor=tk.E, width=130)

            self.tree.heading("صافي الراتب", text="💎 صافي الراتب")
            self.tree.column("صافي الراتب", anchor=tk.E, width=140)

            for idx, pg in enumerate(unified_pages, start=1):
                tot_ent = pg.get("printed_gross")
                tot_ded = pg.get("printed_deductions")
                net = pg.get("printed_net")
                # Official Oracle totals are authoritative. Never fabricate a net amount.

                row_vals = [idx, to_hindi_nums(pg["code"]), self.fix_ar(pg["name"])]
                for itm in ent_items:
                    amt = lookup(pg, "الاستحقاقات", itm)
                    row_vals.append(to_hindi_nums(f"{amt:,.2f}") if amt is not None else "غير متاح")
                row_vals.append(to_hindi_nums(f"{tot_ent:,.2f}") if tot_ent is not None else "غير متاح")

                for itm in ded_items:
                    amt = lookup(pg, "الاستقطاعات", itm)
                    row_vals.append(to_hindi_nums(f"{amt:,.2f}") if amt is not None else "غير متاح")
                row_vals.append(to_hindi_nums(f"{tot_ded:,.2f}") if tot_ded is not None else "غير متاح")

                row_vals.append(to_hindi_nums(f"{net:,.2f}") if net is not None else "غير متاح")
                self.tree.insert("", tk.END, values=row_vals)

            self.current_export_info = {
                "mode": "comprehensive",
                "ent_items": ent_items,
                "ded_items": ded_items,
                "data": unified_pages
            }
            self.lbl_file.config(text=f"تم استخراج الكشف الشامل لـ {len(unified_pages)} موظف ({len(ent_items)} استحقاق، {len(ded_items)} استقطاع)", fg="green")
            self.btn_xlsx.config(state=tk.NORMAL)

    def export(self):
        """Interactive employee/item scope, without overwriting prior source records."""
        if not getattr(self,"parsed_employees",None):
            messagebox.showwarning("تنبيه","يجب مسح ملف يحتوي على موظفين أولًا")
            return
        employees=self.parsed_employees
        popup=tk.Toplevel(self.root)
        popup.title("اختيار الموظفين والبنود للتصدير")
        popup.geometry("990x720")
        popup.transient(self.root)
        popup.grab_set()
        area=tk.Frame(popup,padx=14,pady=12)
        area.pack(fill=tk.BOTH,expand=True)
        tk.Label(area,text="اختر موظفًا أو أكثر، وبندًا أو أكثر، أو الإجماليات الرسمية",font=("Segoe UI",12,"bold")).pack(anchor="e")
        tk.Label(area,text="الموظفون (Ctrl/Shift للاختيار المتعدد، أو زر تحديد الكل)",font=("Segoe UI",10,"bold")).pack(anchor="e")
        emp_frame=tk.Frame(area)
        emp_frame.pack(fill=tk.BOTH,expand=True)
        emp_scroll=tk.Scrollbar(emp_frame)
        emp_scroll.pack(side=tk.LEFT,fill=tk.Y)
        emp_list=tk.Listbox(emp_frame,selectmode=tk.EXTENDED,yscrollcommand=emp_scroll.set,
                            font=("Segoe UI",10),exportselection=False)
        emp_list.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        emp_scroll.config(command=emp_list.yview)
        for e in employees:
            emp_list.insert(tk.END,f"{e.get('name',UNKNOWN)}  |  {e.get('disbursement_number') or e.get('national_id') or 'غير معروف'}")
        def select_all_employee():
            emp_list.selection_set(0,tk.END)
        tk.Button(area,text="تحديد كل الموظفين",command=select_all_employee).pack(anchor="e",pady=5)
        select_all_employee()
        flags=tk.Frame(area);flags.pack(fill=tk.X,pady=5)
        gross=tk.BooleanVar(value=False)
        deduction=tk.BooleanVar(value=False)
        net=tk.BooleanVar(value=True)
        for caption,var in (("إجمالي الاستحقاقات",gross),("إجمالي الاستقطاعات",deduction),("الصافي الرسمي",net)):
            tk.Checkbutton(flags,text=caption,variable=var).pack(side=tk.RIGHT,padx=9)
        tk.Label(area,text="بنود فردية (يمكن اختيار عدة بنود؛ Ctrl/Shift)",font=("Segoe UI",10,"bold")).pack(anchor="e")
        available=sorted({k for e in employees for k in e["items"]},key=lambda x:(x[0],x[1],x[2]))
        it_frame=tk.Frame(area);it_frame.pack(fill=tk.BOTH,expand=True)
        it_scroll=tk.Scrollbar(it_frame);it_scroll.pack(side=tk.LEFT,fill=tk.Y)
        it_list=tk.Listbox(it_frame,selectmode=tk.EXTENDED,yscrollcommand=it_scroll.set,
                           font=("Segoe UI",10),exportselection=False)
        it_list.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        it_scroll.config(command=it_list.yview)
        for kind,code,label in available:it_list.insert(tk.END,f"{kind} | {code} | {label}")
        tk.Button(area,text="تحديد جميع البنود",command=lambda:it_list.selection_set(0,tk.END)).pack(anchor="e",pady=4)
        def save_selection():
            indexes=list(emp_list.curselection())
            item_indexes=list(it_list.curselection())
            if not indexes:
                messagebox.showwarning("تنبيه","حدد موظفًا واحدًا على الأقل",parent=popup)
                return
            if not (gross.get() or deduction.get() or net.get() or item_indexes):
                messagebox.showwarning("تنبيه","حدد إجماليًا أو بندًا واحدًا على الأقل",parent=popup)
                return
            destination=filedialog.asksaveasfilename(parent=popup,defaultextension=".xlsx",
                          filetypes=[("Excel",".xlsx")],initialfile="كشف_مرتبات_مخصص.xlsx")
            if not destination:return
            try:
                result=create_excel(destination,employees,self.pdf_path,
                    selected_indexes=indexes,
                    selected_items=[available[i] for i in item_indexes],
                    include_gross=gross.get(),include_deductions=deduction.get(),include_net=net.get())
            except Exception as exc:
                messagebox.showerror("خطأ تصدير",str(exc),parent=popup)
                return
            popup.destroy()
            messagebox.showinfo("نجاح",f"تم تصدير {result['employees']} موظف و {result['details']} صف بند، مع {result['warnings']} تحذير مراجعة.")
        tk.Button(area,text="تصدير Excel",command=save_selection,
                  font=("Segoe UI",11,"bold"),bg="#2E7D32",fg="white").pack(fill=tk.X,pady=8)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = SmartPayrollApp(root)
    root.mainloop()
