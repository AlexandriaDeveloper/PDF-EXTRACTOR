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
        self.root.title("مستخرج رواتب أوراكل - نظام الاختيار الذكي")
        self.root.geometry("1100x750")
        
        self.pdf_path = ""
        self.items_db = {"الاستقطاعات": [], "الاستحقاقات": []}
        self.extracted_df = None
        self.pages_data = []
        self.msg_queue = queue.Queue()
        
        self.setup_ui()

    def setup_ui(self):
        # الجزء العلوي: اختيار الملف
        top = tk.Frame(self.root, pady=10)
        top.pack(fill=tk.X)
        tk.Button(top, text="1. اختر ملف PDF للبدء بالمسح", command=self.select_file, bg="#2196F3", fg="white", padx=15).pack(side=tk.RIGHT, padx=10)
        self.lbl_file = tk.Label(top, text="بانتظار اختيار الملف...", fg="gray")
        self.lbl_file.pack(side=tk.RIGHT, padx=10)
        self.progress = ttk.Progressbar(top, mode="determinate", length=250)
        self.progress.pack(side=tk.LEFT, padx=10)

        # جزء الاختيار (يتم تفعيله بعد المسح)
        self.sel_frame = tk.LabelFrame(self.root, text="2. تخصيص الاستخراج", padx=10, pady=10)
        self.sel_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # قائمة النوع
        tk.Label(self.sel_frame, text="اختر النوع:").pack(side=tk.RIGHT, padx=5)
        self.combo_cat = ttk.Combobox(self.sel_frame, values=["الاستقطاعات", "الاستحقاقات"], state="readonly", width=15)
        self.combo_cat.pack(side=tk.RIGHT, padx=10)
        self.combo_cat.bind("<<ComboboxSelected>>", self.update_item_list)

        # قائمة البند
        tk.Label(self.sel_frame, text="اختر البند:").pack(side=tk.RIGHT, padx=5)
        self.combo_item = ttk.Combobox(self.sel_frame, state="disabled", width=45)
        self.combo_item.pack(side=tk.RIGHT, padx=10)
        
        tk.Button(self.sel_frame, text="عرض النتائج لهذا البند", command=self.process_selected, bg="#FF9800", fg="white", padx=15).pack(side=tk.LEFT, padx=10)

        # الجدول
        grid = tk.Frame(self.root)
        grid.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
        self.cols = ("اسم الموظف", "كود الموظف", "قيمة المبلغ")
        self.tree = ttk.Treeview(grid, columns=self.cols, show='headings')
        for c in self.cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor=tk.CENTER, width=200)
        
        vsb = ttk.Scrollbar(grid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # التصدير
        self.btn_xlsx = tk.Button(self.root, text="3. تصدير الجدول إلى Excel", command=self.export, state=tk.DISABLED, bg="#4CAF50", fg="white", pady=10)
        self.btn_xlsx.pack(fill=tk.X, padx=10, pady=10)

    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.pdf_path = p
            self.pages_data = []
            self.lbl_file.config(text=f"جاري مسح: {os.path.basename(p)}...", fg="blue")
            self.combo_item.config(state="disabled")
            self.progress["value"] = 0
            self.msg_queue = queue.Queue()
            # التحليل في خيط خلفي حتى لا تتجمد الواجهة
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
                    self.lbl_file.config(text=f"جاري تحليل الملف: {done} من {total} صفحة...", fg="blue")
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
                items_set[cat].add(item)
        for cat in items_set:
            self.items_db[cat] = sorted(items_set[cat])

        self.lbl_file.config(text=f"تم المسح: {os.path.basename(self.pdf_path)} ({len(self.pages_data)} صفحة)", fg="green")
        self.combo_cat.set("الاستقطاعات")
        self.update_item_list()
        total_items = len(self.items_db["الاستقطاعات"]) + len(self.items_db["الاستحقاقات"])
        if total_items == 0:
            messagebox.showwarning("تنبيه", "لم يتم العثور على أي بنود. تأكد أن الملف يحتوي على جداول نصوص وليس صورا.")
        else:
            messagebox.showinfo("تم المسح", f"تم العثور على {total_items} بند مالي مختلف.\nالاستخراج الآن فوري لأي بند.")

    def process_selected(self):
        """استخراج فوري من البيانات المخزنة دون إعادة قراءة الملف"""
        target_item = self.combo_item.get()
        category = self.combo_cat.get()
        if not target_item or target_item == "لا توجد بنود" or not self.pages_data:
            return
        key = (category, target_item)
        for i in self.tree.get_children():
            self.tree.delete(i)
        results = []
        for pg in self.pages_data:
            amount_val = pg["items"].get(key)
            if amount_val is not None:
                amount_num = parse_amount_to_float(amount_val)
                # عرض المبلغ في الجدول منسقاً
                self.tree.insert("", tk.END, values=(
                    self.fix_ar(pg["name"]),
                    to_hindi_nums(pg["code"]),
                    to_hindi_nums(f"{amount_num:,.2f}")
                ))
                # حفظ المبلغ كـ float حقيقي للعمليات الحسابية والتصدير
                results.append({
                    "كود الموظف": to_english_nums(pg["code"]),
                    "اسم الموظف": pg["name"],
                    "المبلغ": amount_num
                })

        self.lbl_file.config(text=f"تم استخراج {len(results)} سجل لـ: {target_item}", fg="green")
        if results:
            self.extracted_df = pd.DataFrame(results)
            self.btn_xlsx.config(state=tk.NORMAL)
        else:
            self.extracted_df = None
            self.btn_xlsx.config(state=tk.DISABLED)
            messagebox.showinfo("نتيجة", "لم يتم العثور على أي موظف لديه هذا البند.")

    def update_item_list(self, event=None):
        cat = self.combo_cat.get()
        if cat in self.items_db:
            items = self.items_db[cat]
            self.combo_item.config(state="readonly", values=items)
            if items:
                self.combo_item.set(items[0])
            else:
                self.combo_item.set("لا توجد بنود")

    def fix_ar(self, t):
        if not t: return ""
        try: return get_display(arabic_reshaper.reshape(str(t).strip()))
        except: return str(t)

    def export(self):
        """تصدير تقرير إكسيل احترافي بتنسيق مالي كامل"""
        if self.extracted_df is None or self.extracted_df.empty:
            messagebox.showwarning("تنبيه", "لا توجد بيانات لتصديرها.")
            return

        target_item = self.combo_item.get()
        category = self.combo_cat.get()
        safe_item_name = re.sub(r'[\\/*?:"<>|]', '_', target_item).strip()
        default_file = f"{safe_item_name}.xlsx"

        f = filedialog.asksaveasfilename(
            initialfile=default_file,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if f:
            try:
                self._save_styled_excel(f, target_item, category)
                messagebox.showinfo("نجاح التصدير", f"تم تصدير ملف الإكسيل وتنسيقه باحترافية:\n{os.path.basename(f)}")
            except Exception as e:
                # تصدير احتياطي في حال حدوث أي خطأ في التنسيق
                self.extracted_df.to_excel(f, index=False)
                messagebox.showinfo("تم التصدير", f"تم التصدير بالطريقة القياسية:\n{os.path.basename(f)}\n(ملاحظة: {str(e)})")

    def _save_styled_excel(self, filepath, target_item, category):
        """تصدير تقرير إكسيل بتنسيق مالي أنيق واتجاه صفحة من اليمين لليسار مع معادلة الإجمالي"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "كشف المبالغ"

        # 1. ضبط اتجاه الصفحة من اليمين لليسار وإظهار خطوط الشبكة
        ws.views.sheetView[0].rightToLeft = True
        ws.views.sheetView[0].showGridLines = True

        # ألوان كلاسيكية رسمية
        c_navy_header = "1B365D"   # كحلي ملكي للشريط العلوي
        c_navy_table  = "203764"   # كحلي رسمي لترويسة الجدول
        c_info_bar    = "E9EEF4"   # خلفية هادئة لشريط الملخص
        c_zebra_odd   = "F2F6FB"   # تبادل الأسطر لراحة العين
        c_total_fill  = "D9E1F2"   # خلفية سطر الإجمالي الكلي
        c_border_gray = "D9D9D9"   # لون حدود الخلايا

        side_thin = Side(style='thin', color=c_border_gray)
        border_data = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)

        # 2. العنوان الرئيسي المدمج (Banner)
        ws.merge_cells('A1:D1')
        title_cell = ws['A1']
        title_cell.value = f"كشف استخراج: {target_item} ({category})"
        title_cell.font = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
        title_cell.fill = PatternFill(start_color=c_navy_header, end_color=c_navy_header, fill_type='solid')
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 36

        # 3. شريط المعلومات والملخص
        total_sum = float(self.extracted_df["المبلغ"].sum()) if "المبلغ" in self.extracted_df else 0.0
        total_count = len(self.extracted_df)
        pdf_name = os.path.basename(self.pdf_path) if self.pdf_path else "ملف غير محدد"

        ws.merge_cells('A2:D2')
        info_cell = ws['A2']
        info_cell.value = f"الملف المصدر: {pdf_name}   |   إجمالي عدد الموظفين: {total_count:,}   |   إجمالي المبلغ: {total_sum:,.2f} ج.م"
        info_cell.font = Font(name='Calibri', size=10, bold=True, color='1B365D')
        info_cell.fill = PatternFill(start_color=c_info_bar, end_color=c_info_bar, fill_type='solid')
        info_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 22

        # سطر فاصل فارغ
        ws.row_dimensions[3].height = 10

        # 4. ترويسة الجدول
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

        # 5. صفوف البيانات
        for idx, row in self.extracted_df.iterrows():
            row_num = 5 + idx
            fill_color = c_zebra_odd if idx % 2 == 1 else "FFFFFF"
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

            # م (مسلسل)
            c_serial = ws.cell(row=row_num, column=1, value=idx + 1)
            c_serial.alignment = Alignment(horizontal='center', vertical='center')
            c_serial.fill = row_fill
            c_serial.border = border_data

            # كود الموظف
            c_code = ws.cell(row=row_num, column=2, value=str(row.get("كود الموظف", "")))
            c_code.alignment = Alignment(horizontal='center', vertical='center')
            c_code.fill = row_fill
            c_code.border = border_data
            c_code.number_format = '@'

            # اسم الموظف
            c_name = ws.cell(row=row_num, column=3, value=str(row.get("اسم الموظف", "")))
            c_name.alignment = Alignment(horizontal='right', vertical='center')
            c_name.fill = row_fill
            c_name.border = border_data

            # المبلغ (رقم عشري حقيقي مخصص للعمليات الحسابية والجمع التلقائي في إكسيل)
            amt_val = parse_amount_to_float(row.get("المبلغ", 0.0))
            c_amt = ws.cell(row=row_num, column=4, value=amt_val)
            c_amt.alignment = Alignment(horizontal='right', vertical='center')
            c_amt.fill = row_fill
            c_amt.border = border_data
            c_amt.number_format = '#,##0.00'

            ws.row_dimensions[row_num].height = 22

        # 6. صف الإجمالي الكلي المحاسبي
        tot_row = 5 + total_count
        tot_fill = PatternFill(start_color=c_total_fill, end_color=c_total_fill, fill_type='solid')
        tot_border = Border(
            top=Side(style='thin', color='203764'),
            bottom=Side(style='double', color='203764'),  # التسطير المزدوج المحاسبي المعتمد
            left=side_thin, right=side_thin
        )

        c1 = ws.cell(row=tot_row, column=1, value="")
        c1.fill = tot_fill
        c1.border = tot_border

        c2 = ws.cell(row=tot_row, column=2, value=f"العدد: {total_count}")
        c2.font = Font(name='Calibri', size=10, bold=True, color='203764')
        c2.alignment = Alignment(horizontal='center', vertical='center')
        c2.fill = tot_fill
        c2.border = tot_border

        c3 = ws.cell(row=tot_row, column=3, value="الإجمالي الكلي")
        c3.font = Font(name='Calibri', size=11, bold=True, color='203764')
        c3.alignment = Alignment(horizontal='right', vertical='center')
        c3.fill = tot_fill
        c3.border = tot_border

        # معادلة إكسيل حقيقية لجمع عمود المبالغ
        c4 = ws.cell(row=tot_row, column=4, value=f"=SUM(D5:D{tot_row - 1})")
        c4.font = Font(name='Calibri', size=11, bold=True, color='203764')
        c4.alignment = Alignment(horizontal='right', vertical='center')
        c4.fill = tot_fill
        c4.border = tot_border
        c4.number_format = '#,##0.00'
        ws.row_dimensions[tot_row].height = 28

        # 7. ضبط عرض الأعمدة وتجميد الألواح
        max_name_len = max([len(str(x)) for x in self.extracted_df["اسم الموظف"]] + [10]) if not self.extracted_df.empty else 10
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = max(max_name_len + 6, 32)
        ws.column_dimensions['D'].width = 18

        # تجميد الصفوف الأربعة الأولى بحيث تظل الترويسة ثابتة عند التمرير
        ws.freeze_panes = 'A5'

        wb.save(filepath)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = SmartPayrollApp(root)
    root.mainloop()
