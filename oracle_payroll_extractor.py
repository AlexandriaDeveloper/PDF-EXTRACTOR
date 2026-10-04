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
    """استبعاد العناوين والإجماليات وصافي الراتب من قائمة البنود"""
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
    """تحليل صفحة واحدة: الاسم، الكود، وكل البنود {(القسم, البند): المبلغ}"""
    tables = page.extract_tables()

    # الترويسة (الاسم والكود) موجودة غالباً في أول خلية بالجدول - تغنينا عن extract_text المكلفة
    header = ""
    if tables and tables[0] and tables[0][0]:
        header = "\n".join(str(c) for c in tables[0][0] if c)
    emp_name, emp_code = extract_emp_info(header)
    if emp_name == UNKNOWN or emp_code == UNKNOWN:
        n2, c2 = extract_emp_info(page.extract_text() or "")
        if emp_name == UNKNOWN: emp_name = n2
        if emp_code == UNKNOWN: emp_code = c2

    items = {}
    for t in tables:
        t_dec = decode_oracle_text(" ".join(str(c) for r in t for c in r if c))
        is_deduc_table = ("الاستقطاعات" in t_dec) and ("الاستحقاقات" not in t_dec)
        is_entit_table = ("الاستحقاقات" in t_dec)
        for row in t:
            if any(c and str(c).count("\n") > 3 for c in row):
                continue
            val_flags = [is_val(c) for c in row]
            if not any(val_flags):
                continue
            for i, cell in enumerate(row):
                if not cell or val_flags[i]:
                    continue
                raw_s = str(cell).strip()
                if len(raw_s) < 3:
                    continue
                clean_s = decode_oracle_text(raw_s)
                if not re.search(r'[\u0621-\u064A]', clean_s) or len(clean_s) < 3 or is_ignored_item(clean_s):
                    continue
                if is_deduc_table:
                    cat = "الاستقطاعات"
                elif is_entit_table:
                    cat = "الاستقطاعات" if (len(row) > 6 and i < len(row) // 2) else "الاستحقاقات"
                else:
                    cat = "الاستقطاعات" if (len(row) <= 4 or i < len(row) // 2) else "الاستحقاقات"
                amount = next((str(c).strip() for j, c in enumerate(row) if j != i and val_flags[j]), "")
                if amount and (cat, clean_s) not in items:
                    items[(cat, clean_s)] = parse_amount_to_float(amount)
                break  # بند واحد لكل صف
    return {"name": emp_name, "code": emp_code, "items": items}

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
        for pg in self.pages_data:
            for (cat, item) in pg["items"]:
                if cat in items_set:
                    items_set[cat].add(item)
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
        """توحيد بيانات الموظف إذا كان يمتلك أكثر من صفحة في ملف الرواتب"""
        emps = {}
        for pg in pages_data:
            code = str(pg.get("code", "")).strip()
            name = str(pg.get("name", "")).strip()
            key = (code, name) if (code or name) else id(pg)
            if key not in emps:
                emps[key] = {
                    "code": code,
                    "name": name,
                    "items": dict(pg.get("items", {}))
                }
            else:
                emps[key]["items"].update(pg.get("items", {}))
        return list(emps.values())

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

        unified_pages = self.get_unified_employees(self.pages_data)

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
                amount_val = pg["items"].get(key)
                if amount_val is not None:
                    amount_num = parse_amount_to_float(amount_val)
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
                    amt = pg["items"].get((cat_name, itm), 0.0)
                    row_sum += amt
                    row_vals.append(to_hindi_nums(f"{amt:,.2f}") if amt > 0 else "-")
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
                tot_ent = sum(pg["items"].get(("الاستحقاقات", itm), 0.0) for itm in ent_items)
                tot_ded = sum(pg["items"].get(("الاستقطاعات", itm), 0.0) for itm in ded_items)
                net = tot_ent - tot_ded

                row_vals = [idx, to_hindi_nums(pg["code"]), self.fix_ar(pg["name"])]
                for itm in ent_items:
                    amt = pg["items"].get(("الاستحقاقات", itm), 0.0)
                    row_vals.append(to_hindi_nums(f"{amt:,.2f}") if amt > 0 else "-")
                row_vals.append(to_hindi_nums(f"{tot_ent:,.2f}"))

                for itm in ded_items:
                    amt = pg["items"].get(("الاستقطاعات", itm), 0.0)
                    row_vals.append(to_hindi_nums(f"{amt:,.2f}") if amt > 0 else "-")
                row_vals.append(to_hindi_nums(f"{tot_ded:,.2f}"))

                row_vals.append(to_hindi_nums(f"{net:,.2f}"))
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
        """تصدير تقرير إكسيل احترافي بتنسيق مالي كامل ومحدد حسب النمط"""
        if not self.current_export_info:
            messagebox.showwarning("تنبيه", "لا توجد بيانات مستخرجة لتصديرها.")
            return

        info = self.current_export_info
        mode = info["mode"]
        pdf_name = os.path.basename(self.pdf_path) if self.pdf_path else "ملف غير محدد"

        if mode == "single":
            safe_name = re.sub(r'[\\/*?:"<>|]', '_', info["target_item"]).strip()
            default_file = f"كشف_{safe_name}.xlsx"
            dialog_title = f"حفظ كشف {info['target_item']}"
        elif mode == "category":
            safe_name = "الاستحقاقات" if info["category"] == "الاستحقاقات" else "الاستقطاعات"
            default_file = f"كشف_{safe_name}_كافة_الموظفين.xlsx"
            dialog_title = f"حفظ كشف {info['category']}"
        elif mode == "comprehensive":
            default_file = "كشف_المرتبات_الشامل.xlsx"
            dialog_title = "حفظ كشف المرتبات الشامل"
        else:
            default_file = "تقرير_الرواتب.xlsx"
            dialog_title = "حفظ التقرير"

        f = filedialog.asksaveasfilename(
            title=dialog_title,
            initialfile=default_file,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not f:
            return

        try:
            if mode == "single":
                self._save_single_item_excel(f, info["data"], info["target_item"], info["category"], pdf_name)
            elif mode == "category":
                self._save_category_excel(f, info["data"], info["category"], info["items"], pdf_name)
            elif mode == "comprehensive":
                self._save_comprehensive_excel(f, info["data"], info["ent_items"], info["ded_items"], pdf_name)
            
            messagebox.showinfo("نجاح التصدير", f"تم تصدير ملف الإكسيل وتنسيقه باحترافية عالية:\n{os.path.basename(f)}")
        except Exception as e:
            messagebox.showerror("خطأ في التصدير", f"حدث خطأ أثناء تصدير الملف:\n{str(e)}")

    def _save_single_item_excel(self, filepath, rows_data, target_item, category, pdf_name):
        """تصدير تقرير إكسيل لبند واحد بتنسيق مالي أنيق واتجاه صفحة من اليمين لليسار"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "كشف البند"
        ws.views.sheetView[0].rightToLeft = True
        ws.views.sheetView[0].showGridLines = True

        c_navy_header = "1B365D"
        c_navy_table  = "203764"
        c_info_bar    = "E9EEF4"
        c_zebra_odd   = "F2F6FB"
        c_total_fill  = "D9E1F2"
        c_border_gray = "D9D9D9"

        side_thin = Side(style='thin', color=c_border_gray)
        border_data = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)

        # 1. العنوان الرئيسي المدمج
        ws.merge_cells('A1:D1')
        title_cell = ws['A1']
        title_cell.value = f"كشف استخراج: {target_item} ({category})"
        title_cell.font = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
        title_cell.fill = PatternFill(start_color=c_navy_header, end_color=c_navy_header, fill_type='solid')
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 36

        # 2. شريط المعلومات والملخص
        total_count = len(rows_data)
        total_sum = sum(r["المبلغ"] for r in rows_data)

        ws.merge_cells('A2:D2')
        info_cell = ws['A2']
        info_cell.value = f"الملف المصدر: {pdf_name}   |   إجمالي عدد الموظفين: {total_count:,}   |   إجمالي المبلغ: {total_sum:,.2f} ج.م"
        info_cell.font = Font(name='Calibri', size=10, bold=True, color='1B365D')
        info_cell.fill = PatternFill(start_color=c_info_bar, end_color=c_info_bar, fill_type='solid')
        info_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 22
        ws.row_dimensions[3].height = 10

        # 3. ترويسة الجدول
        headers = ["م", "كود الموظف", "اسم الموظف", "المبلغ (ج.م)"]
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h_text)
            cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color=c_navy_table, end_color=c_navy_table, fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=side_thin, right=side_thin, top=side_thin,
                bottom=Side(style='medium', color='142340')
            )
        ws.row_dimensions[4].height = 26

        # 4. صفوف البيانات
        for idx, row in enumerate(rows_data):
            row_num = 5 + idx
            fill_color = c_zebra_odd if idx % 2 == 1 else "FFFFFF"
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

            # م
            c_serial = ws.cell(row=row_num, column=1, value=idx + 1)
            c_serial.alignment = Alignment(horizontal='center', vertical='center')
            c_serial.fill = row_fill; c_serial.border = border_data

            # كود الموظف نصي للحفاظ على الأصفار في البداية
            c_code = ws.cell(row=row_num, column=2, value=str(row.get("كود الموظف", "")))
            c_code.alignment = Alignment(horizontal='center', vertical='center')
            c_code.fill = row_fill; c_code.border = border_data
            c_code.number_format = '@'

            # اسم الموظف
            c_name = ws.cell(row=row_num, column=3, value=str(row.get("اسم الموظف", "")))
            c_name.alignment = Alignment(horizontal='right', vertical='center')
            c_name.fill = row_fill; c_name.border = border_data

            # المبلغ
            amt_val = parse_amount_to_float(row.get("المبلغ", 0.0))
            c_amt = ws.cell(row=row_num, column=4, value=amt_val)
            c_amt.alignment = Alignment(horizontal='right', vertical='center')
            c_amt.fill = row_fill; c_amt.border = border_data
            c_amt.number_format = '#,##0.00;(#,##0.00);"-"'

            ws.row_dimensions[row_num].height = 22

        # 5. صف الإجمالي الكلي المحاسبي
        tot_row = 5 + total_count
        tot_fill = PatternFill(start_color=c_total_fill, end_color=c_total_fill, fill_type='solid')
        tot_border = Border(
            top=Side(style='thin', color='203764'),
            bottom=Side(style='double', color='203764'),
            left=side_thin, right=side_thin
        )

        c1 = ws.cell(row=tot_row, column=1, value="")
        c1.fill = tot_fill; c1.border = tot_border

        c2 = ws.cell(row=tot_row, column=2, value=f"العدد: {total_count}")
        c2.font = Font(name='Calibri', size=10, bold=True, color='203764')
        c2.alignment = Alignment(horizontal='center', vertical='center')
        c2.fill = tot_fill; c2.border = tot_border

        c3 = ws.cell(row=tot_row, column=3, value="الإجمالي الكلي")
        c3.font = Font(name='Calibri', size=11, bold=True, color='203764')
        c3.alignment = Alignment(horizontal='right', vertical='center')
        c3.fill = tot_fill; c3.border = tot_border

        c4 = ws.cell(row=tot_row, column=4, value=f"=SUM(D5:D{tot_row - 1})")
        c4.font = Font(name='Calibri', size=11, bold=True, color='203764')
        c4.alignment = Alignment(horizontal='right', vertical='center')
        c4.fill = tot_fill; c4.border = tot_border
        c4.number_format = '#,##0.00;(#,##0.00);"-"'
        ws.row_dimensions[tot_row].height = 28

        # 6. ضبط عرض الأعمدة وتجميد الألواح
        max_name_len = max([len(str(r.get("اسم الموظف", ""))) for r in rows_data] + [10]) if rows_data else 10
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = max(max_name_len + 6, 32)
        ws.column_dimensions['D'].width = 18

        ws.freeze_panes = 'A5'
        wb.save(filepath)

    def _save_category_excel(self, filepath, pages_data, category, items_list, pdf_name):
        """تصدير تقرير إكسيل لكافة بنود الاستحقاقات أو الاستقطاعات بالتفصيل والإجماليات"""
        is_ent = (category == "الاستحقاقات")
        title_text = "كشف تفصيلي لكافة بنود الاستحقاقات (المستحق)" if is_ent else "كشف تفصيلي لكافة بنود الاستقطاعات (المستقطع)"
        sheet_title = "الاستحقاقات" if is_ent else "الاستقطاعات"

        # ثيم الألوان
        theme_banner = "1E5631" if is_ent else "781D29"
        theme_header = "2D6A4F" if is_ent else "8A1C14"
        theme_total_col = "1B4332" if is_ent else "5C1D24"
        theme_info_bg = "E8F5E9" if is_ent else "FFE4E6"
        theme_info_fg = "1E5631" if is_ent else "781D29"
        theme_zebra = "F4FBF7" if is_ent else "FFF8F8"
        theme_tot_row_bg = "C7E9C0" if is_ent else "F8D7DA"
        theme_tot_border = "1E5631" if is_ent else "781D29"

        c_border_gray = "D9D9D9"
        side_thin = Side(style='thin', color=c_border_gray)
        border_data = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title
        ws.views.sheetView[0].rightToLeft = True
        ws.views.sheetView[0].showGridLines = True

        m_items = len(items_list)
        tot_col_idx = 4 + m_items
        last_col_letter = get_column_letter(tot_col_idx)

        # 1. العنوان المدمج
        ws.merge_cells(f'A1:{last_col_letter}1')
        t = ws['A1']
        t.value = title_text
        t.font = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
        t.fill = PatternFill(start_color=theme_banner, end_color=theme_banner, fill_type='solid')
        t.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 36

        # 2. شريط الملخص
        total_count = len(pages_data)
        ws.merge_cells(f'A2:{last_col_letter}2')
        info = ws['A2']
        info.value = f"الملف المصدر: {pdf_name}   |   إجمالي عدد الموظفين: {total_count:,}   |   عدد البنود المشمولة: {m_items}"
        info.font = Font(name='Calibri', size=10, bold=True, color=theme_info_fg)
        info.fill = PatternFill(start_color=theme_info_bg, end_color=theme_info_bg, fill_type='solid')
        info.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 22
        ws.row_dimensions[3].height = 10

        # 3. ترويسة الجدول
        headers = ["م", "كود الموظف", "اسم الموظف"] + items_list + [f"إجمالي {category}"]
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h_text)
            cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
            bg = theme_total_col if col_idx == tot_col_idx else theme_header
            cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(left=side_thin, right=side_thin, top=side_thin, bottom=Side(style='medium', color='142340'))
        ws.row_dimensions[4].height = 32

        # 4. صفوف البيانات
        for idx, pg in enumerate(pages_data):
            row_num = 5 + idx
            fill_color = theme_zebra if idx % 2 == 1 else "FFFFFF"
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

            # م
            c1 = ws.cell(row=row_num, column=1, value=idx + 1)
            c1.alignment = Alignment(horizontal='center', vertical='center')
            c1.fill = row_fill; c1.border = border_data

            # كود الموظف نصي
            c2 = ws.cell(row=row_num, column=2, value=str(pg["code"]))
            c2.alignment = Alignment(horizontal='center', vertical='center')
            c2.fill = row_fill; c2.border = border_data
            c2.number_format = '@'

            # اسم الموظف
            c3 = ws.cell(row=row_num, column=3, value=str(pg["name"]))
            c3.alignment = Alignment(horizontal='right', vertical='center')
            c3.fill = row_fill; c3.border = border_data

            # قيم البنود
            for i_idx, itm in enumerate(items_list):
                c_idx = 4 + i_idx
                amt = pg["items"].get((category, itm), 0.0)
                c = ws.cell(row=row_num, column=c_idx, value=amt)
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.fill = row_fill; c.border = border_data
                c.number_format = '#,##0.00;(#,##0.00);"-"'

            # إجمالي الصف بمعادلة SUM
            first_item_letter = get_column_letter(4)
            last_item_letter = get_column_letter(3 + m_items)
            c_tot = ws.cell(row=row_num, column=tot_col_idx, value=f"=SUM({first_item_letter}{row_num}:{last_item_letter}{row_num})")
            c_tot.font = Font(name='Calibri', size=11, bold=True, color='000000')
            c_tot.alignment = Alignment(horizontal='right', vertical='center')
            c_tot.fill = row_fill; c_tot.border = border_data
            c_tot.number_format = '#,##0.00;(#,##0.00);"-"'

            ws.row_dimensions[row_num].height = 21

        # 5. سطر الإجمالي الكلي المحاسبي
        tot_row = 5 + total_count
        tot_fill = PatternFill(start_color=theme_tot_row_bg, end_color=theme_tot_row_bg, fill_type='solid')
        tot_border = Border(top=Side(style='thin', color=theme_tot_border), bottom=Side(style='double', color=theme_tot_border), left=side_thin, right=side_thin)

        ws.cell(row=tot_row, column=1, value="").fill = tot_fill
        ws.cell(row=tot_row, column=1).border = tot_border
        c2 = ws.cell(row=tot_row, column=2, value=f"العدد: {total_count}")
        c2.font = Font(name='Calibri', size=10, bold=True, color=theme_tot_border)
        c2.alignment = Alignment(horizontal='center', vertical='center')
        c2.fill = tot_fill; c2.border = tot_border

        c3 = ws.cell(row=tot_row, column=3, value="الإجمالي الكلي")
        c3.font = Font(name='Calibri', size=11, bold=True, color=theme_tot_border)
        c3.alignment = Alignment(horizontal='right', vertical='center')
        c3.fill = tot_fill; c3.border = tot_border

        for c_idx in range(4, tot_col_idx + 1):
            col_let = get_column_letter(c_idx)
            c = ws.cell(row=tot_row, column=c_idx, value=f"=SUM({col_let}5:{col_let}{tot_row - 1})")
            c.font = Font(name='Calibri', size=11, bold=True, color=theme_tot_border)
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.fill = tot_fill; c.border = tot_border
            c.number_format = '#,##0.00;(#,##0.00);"-"'

        ws.row_dimensions[tot_row].height = 28

        # 6. ضبط عروض الأعمدة وتجميد الألواح
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 34
        for c_idx in range(4, tot_col_idx + 1):
            ws.column_dimensions[get_column_letter(c_idx)].width = 17

        ws.freeze_panes = 'D5'
        wb.save(filepath)

    def _save_comprehensive_excel(self, filepath, pages_data, ent_items, ded_items, pdf_name):
        """تصدير الكشف الشامل: المستحق والمستقطع والصافي مع سوبر هيدر وتنسيق محاسبي متكامل"""
        m = len(ent_items)
        k = len(ded_items)
        
        tot_ent_col = 4 + m
        first_ded_col = 5 + m
        last_ded_item_col = 4 + m + k
        tot_ded_col = 5 + m + k
        net_col = 6 + m + k
        last_col_letter = get_column_letter(net_col)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "الكشف الشامل"
        ws.views.sheetView[0].rightToLeft = True
        ws.views.sheetView[0].showGridLines = True

        # لوحة الألوان الرسمية
        c_emp_banner  = "1B365D"   # كحلي لبيانات الموظف
        c_ent_banner  = "1E5631"   # أخضر داكن للاستحقاقات
        c_ded_banner  = "781D29"   # نبيتي للاستقطاعات
        c_net_banner  = "0F4C81"   # أزرق ملكي للصافي
        
        c_ent_sub     = "2D6A4F"
        c_ded_sub     = "8A1C14"
        c_ent_tot_sub = "143D28"
        c_ded_tot_sub = "4A1218"
        c_net_sub     = "092C48"

        c_border_gray = "D9D9D9"
        side_thin = Side(style='thin', color=c_border_gray)
        border_data = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)

        # 1. العنوان الرئيسي
        ws.merge_cells(f'A1:{last_col_letter}1')
        t = ws['A1']
        t.value = "كشف المرتبات والأجور الشامل (الاستحقاقات والاستقطاعات وصافي الراتب)"
        t.font = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
        t.fill = PatternFill(start_color="142340", end_color="142340", fill_type='solid')
        t.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 36

        # 2. شريط الملخص
        total_count = len(pages_data)
        ws.merge_cells(f'A2:{last_col_letter}2')
        info = ws['A2']
        info.value = f"الملف المصدر: {pdf_name}   |   إجمالي عدد الموظفين: {total_count:,}   |   عدد بنود الاستحقاق: {m}   |   عدد بنود الاستقطاع: {k}"
        info.font = Font(name='Calibri', size=10, bold=True, color='142340')
        info.fill = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type='solid')
        info.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 22
        ws.row_dimensions[3].height = 10

        # 3. السوبر هيدر (Super-Headers) في الصف 4
        # بيانات الموظف (A4:C4)
        ws.merge_cells('A4:C4')
        c_emp = ws['A4']
        c_emp.value = "بيانات الموظف الأساسية"
        c_emp.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        c_emp.fill = PatternFill(start_color=c_emp_banner, end_color=c_emp_banner, fill_type='solid')
        c_emp.alignment = Alignment(horizontal='center', vertical='center')

        # الاستحقاقات
        ent_start_let = get_column_letter(4)
        ent_end_let = get_column_letter(tot_ent_col)
        ws.merge_cells(f'{ent_start_let}4:{ent_end_let}4')
        c_ent = ws[f'{ent_start_let}4']
        c_ent.value = "بنود الاستحقاقات (المستحق)"
        c_ent.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        c_ent.fill = PatternFill(start_color=c_ent_banner, end_color=c_ent_banner, fill_type='solid')
        c_ent.alignment = Alignment(horizontal='center', vertical='center')

        # الاستقطاعات
        ded_start_let = get_column_letter(first_ded_col)
        ded_end_let = get_column_letter(tot_ded_col)
        ws.merge_cells(f'{ded_start_let}4:{ded_end_let}4')
        c_ded = ws[f'{ded_start_let}4']
        c_ded.value = "بنود الاستقطاعات (المستقطع)"
        c_ded.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        c_ded.fill = PatternFill(start_color=c_ded_banner, end_color=c_ded_banner, fill_type='solid')
        c_ded.alignment = Alignment(horizontal='center', vertical='center')

        # الصافي
        net_let = get_column_letter(net_col)
        c_net = ws[f'{net_let}4']
        c_net.value = "صافي الراتب"
        c_net.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        c_net.fill = PatternFill(start_color=c_net_banner, end_color=c_net_banner, fill_type='solid')
        c_net.alignment = Alignment(horizontal='center', vertical='center')

        ws.row_dimensions[4].height = 24

        # 4. أسماء الأعمدة الفردية (Sub-Headers) في الصف 5
        headers_5 = [
            ("م", "203764"),
            ("كود الموظف", "203764"),
            ("اسم الموظف", "203764")
        ]
        for itm in ent_items:
            headers_5.append((itm, c_ent_sub))
        headers_5.append(("إجمالي المستحق", c_ent_tot_sub))
        for itm in ded_items:
            headers_5.append((itm, c_ded_sub))
        headers_5.append(("إجمالي المستقطع", c_ded_tot_sub))
        headers_5.append(("صافي المبلغ المستحق", c_net_sub))

        for c_idx, (h_title, h_color) in enumerate(headers_5, start=1):
            cell = ws.cell(row=5, column=c_idx, value=h_title)
            cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color=h_color, end_color=h_color, fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(left=side_thin, right=side_thin, top=side_thin, bottom=Side(style='medium', color='142340'))
        ws.row_dimensions[5].height = 32

        # 5. صفوف البيانات
        for idx, pg in enumerate(pages_data):
            row_num = 6 + idx
            fill_color = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

            # م
            c1 = ws.cell(row=row_num, column=1, value=idx + 1)
            c1.alignment = Alignment(horizontal='center', vertical='center')
            c1.fill = row_fill; c1.border = border_data

            # كود الموظف نصي
            c2 = ws.cell(row=row_num, column=2, value=str(pg["code"]))
            c2.alignment = Alignment(horizontal='center', vertical='center')
            c2.fill = row_fill; c2.border = border_data
            c2.number_format = '@'

            # اسم الموظف
            c3 = ws.cell(row=row_num, column=3, value=str(pg["name"]))
            c3.alignment = Alignment(horizontal='right', vertical='center')
            c3.fill = row_fill; c3.border = border_data

            # مبالغ الاستحقاقات
            for i_idx, itm in enumerate(ent_items):
                c_idx = 4 + i_idx
                amt = pg["items"].get(("الاستحقاقات", itm), 0.0)
                c = ws.cell(row=row_num, column=c_idx, value=amt)
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.fill = row_fill; c.border = border_data
                c.number_format = '#,##0.00;(#,##0.00);"-"'

            # معادلة إجمالي المستحق
            ent_f_let = get_column_letter(4)
            ent_l_let = get_column_letter(3 + m)
            c_tot_ent = ws.cell(row=row_num, column=tot_ent_col, value=f"=SUM({ent_f_let}{row_num}:{ent_l_let}{row_num})")
            c_tot_ent.font = Font(name='Calibri', size=10, bold=True, color='1E5631')
            c_tot_ent.alignment = Alignment(horizontal='right', vertical='center')
            c_tot_ent.fill = PatternFill(start_color="E8F5E9" if idx % 2 == 1 else "F1F8F2", end_color="E8F5E9", fill_type='solid')
            c_tot_ent.border = border_data
            c_tot_ent.number_format = '#,##0.00;(#,##0.00);"-"'

            # مبالغ الاستقطاعات
            for j_idx, itm in enumerate(ded_items):
                c_idx = first_ded_col + j_idx
                amt = pg["items"].get(("الاستقطاعات", itm), 0.0)
                c = ws.cell(row=row_num, column=c_idx, value=amt)
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.fill = row_fill; c.border = border_data
                c.number_format = '#,##0.00;(#,##0.00);"-"'

            # معادلة إجمالي الاستقطاع
            ded_f_let = get_column_letter(first_ded_col)
            ded_l_let = get_column_letter(last_ded_item_col)
            c_tot_ded = ws.cell(row=row_num, column=tot_ded_col, value=f"=SUM({ded_f_let}{row_num}:{ded_l_let}{row_num})")
            c_tot_ded.font = Font(name='Calibri', size=10, bold=True, color='781D29')
            c_tot_ded.alignment = Alignment(horizontal='right', vertical='center')
            c_tot_ded.fill = PatternFill(start_color="FFE4E6" if idx % 2 == 1 else "FFF1F2", end_color="FFE4E6", fill_type='solid')
            c_tot_ded.border = border_data
            c_tot_ded.number_format = '#,##0.00;(#,##0.00);"-"'

            # معادلة الصافي: إجمالي المستحق - إجمالي المستقطع
            tot_ent_let = get_column_letter(tot_ent_col)
            tot_ded_let = get_column_letter(tot_ded_col)
            c_net_val = ws.cell(row=row_num, column=net_col, value=f"={tot_ent_let}{row_num}-{tot_ded_let}{row_num}")
            c_net_val.font = Font(name='Calibri', size=11, bold=True, color='0F4C81')
            c_net_val.alignment = Alignment(horizontal='right', vertical='center')
            c_net_val.fill = PatternFill(start_color="D9E1F2" if idx % 2 == 1 else "EBF1F5", end_color="D9E1F2", fill_type='solid')
            c_net_val.border = border_data
            c_net_val.number_format = '#,##0.00;(#,##0.00);"-"'

            ws.row_dimensions[row_num].height = 21

        # 6. صف الإجمالي الكلي المحاسبي في الأسفل
        tot_row = 6 + total_count
        tot_border = Border(top=Side(style='thin', color='142340'), bottom=Side(style='double', color='142340'), left=side_thin, right=side_thin)

        # م
        ws.cell(row=tot_row, column=1, value="").border = tot_border
        ws.cell(row=tot_row, column=1).fill = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type='solid')

        # كود
        c2 = ws.cell(row=tot_row, column=2, value=f"العدد: {total_count}")
        c2.font = Font(name='Calibri', size=10, bold=True, color='142340')
        c2.alignment = Alignment(horizontal='center', vertical='center')
        c2.fill = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type='solid')
        c2.border = tot_border

        # اسم
        c3 = ws.cell(row=tot_row, column=3, value="الإجمالي الكلي")
        c3.font = Font(name='Calibri', size=11, bold=True, color='142340')
        c3.alignment = Alignment(horizontal='right', vertical='center')
        c3.fill = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type='solid')
        c3.border = tot_border

        # مجاميع الأعمدة
        for c_idx in range(4, net_col + 1):
            col_let = get_column_letter(c_idx)
            if c_idx < tot_ent_col:
                bg_col = "C7E9C0"
                fg_col = "1E5631"
            elif c_idx == tot_ent_col:
                bg_col = "A1D99B"
                fg_col = "143D28"
            elif c_idx < tot_ded_col:
                bg_col = "F8D7DA"
                fg_col = "781D29"
            elif c_idx == tot_ded_col:
                bg_col = "F1AEB5"
                fg_col = "4A1218"
            else: # Net
                bg_col = "B8CCE4"
                fg_col = "092C48"

            c = ws.cell(row=tot_row, column=c_idx, value=f"=SUM({col_let}6:{col_let}{tot_row - 1})")
            c.font = Font(name='Calibri', size=11, bold=True, color=fg_col)
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type='solid')
            c.border = tot_border
            c.number_format = '#,##0.00;(#,##0.00);"-"'

        ws.row_dimensions[tot_row].height = 28

        # 7. ضبط عروض الأعمدة وتجميد بيانات الموظف
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 34
        for c_idx in range(4, net_col + 1):
            ws.column_dimensions[get_column_letter(c_idx)].width = 17

        ws.freeze_panes = 'D6'
        wb.save(filepath)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = SmartPayrollApp(root)
    root.mainloop()
