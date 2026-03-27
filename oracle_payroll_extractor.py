import pdfplumber
import pandas as pd
import arabic_reshaper
from bidi.algorithm import get_display
import os
import re
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

"""
النسخة الذكية: مسح تلقائي للبنود + قوائم منسدلة
المتطلبات:
pip install pdfplumber pandas arabic-reshaper python-bidi openpyxl
"""

def to_hindi_nums(text):
    if not text: return ""
    hindi_map = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")
    return str(text).translate(hindi_map)

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
    if not text: return False
    s_text = str(text).replace(",", "").strip()
    hindi_to_eng = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    s_text = s_text.translate(hindi_to_eng)
    return any(c.isdigit() for c in s_text) and ("." in s_text or len(s_text) < 10)

class SmartPayrollApp:
    def __init__(self, root):
        self.root = root
        self.root.title("مستخرج رواتب أوراكل - نظام الاختيار الذكي")
        self.root.geometry("1100x750")
        
        self.pdf_path = ""
        self.items_db = {"الاستقطاعات": [], "الاستحقاقات": []}
        self.extracted_df = None
        
        self.setup_ui()

    def setup_ui(self):
        # الجزء العلوي: اختيار الملف
        top = tk.Frame(self.root, pady=10)
        top.pack(fill=tk.X)
        tk.Button(top, text="1. اختر ملف PDF للبدء بالمسح", command=self.select_file, bg="#2196F3", fg="white", padx=15).pack(side=tk.RIGHT, padx=10)
        self.lbl_file = tk.Label(top, text="بانتظار اختيار الملف...", fg="gray")
        self.lbl_file.pack(side=tk.RIGHT, padx=10)

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
        self.combo_item = ttk.Combobox(self.sel_frame, state="disabled", width=40)
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
            self.lbl_file.config(text=f"جاري مسح: {os.path.basename(p)}...", fg="blue")
            self.root.update()
            self.pre_scan_pdf()

    def pre_scan_pdf(self):
        """مسح أولي شامل لجمع كافة أسماء البنود من كل الصفحات"""
        try:
            items_set = {"الاستقطاعات": set(), "الاستحقاقات": set()}
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for t in tables:
                        # حالة القسم الحالية في هذا الجدول
                        current_cat = None 
                        
                        for row in t:
                            row_txt = " ".join([str(c) for c in row if c])
                            row_norm = normalize(row_txt)
                            
                            # تحديد القسم (نظام أوراكل قد يعكس الكلمات)
                            if "استقطاع" in row_norm or "تاعاطقتس" in row_norm:
                                # قد يكون الجدول مزدوجاً، وقتها نعتمد على مكان الكلمة في الصف
                                pass # سيتم التصنيف لاحقاً حسب العمود
                            
                            # سنبحث عن كلمات مفتاحية في كل خلية
                            deduc_idx, entit_idx = -1, -1
                            for i, cell in enumerate(row):
                                if not cell: continue
                                c_norm = normalize(str(cell))
                                if "استقطاع" in c_norm or "تاعاطقتس" in c_norm: deduc_idx = i
                                if "استحقاق" in c_norm or "تاقاقحتس" in c_norm: entit_idx = i
                            
                            if deduc_idx != -1 or entit_idx != -1:
                                # صف ترويسة - لا نستخرج منه بنود
                                continue

                            # استخراج البنود من الصفوف العادية
                            # إذا كان الصف يحتوي على نص + رقم، فهو غالباً بند مالي
                            has_num = any(is_val(c) for c in row)
                            if has_num:
                                for i, cell in enumerate(row):
                                    if not cell or is_val(cell): continue
                                    cell_s = str(cell).strip()
                                    if len(cell_s) < 3: continue
                                    
                                    # استبعاد الكلمات العامة
                                    if any(w in cell_s for w in ["الاسم", "الكود", "جمال", "صافي", "رقم", "بيان"]): continue
                                    
                                    # تحديد القسم بناء على موقع الخلية (في الجداول المزدوجة الاستقطاعات عادة يسار/أول والاستحقاقات يمين/أخر)
                                    # لكن أوراكل غالباً ما تضع الاستقطاعات من عمود 0-3 والاستحقاقات من 4-7
                                    if len(row) > 6:
                                        if i < len(row)//2: items_set["الاستقطاعات"].add(cell_s)
                                        else: items_set["الاستحقاقات"].add(cell_s)
                                    else:
                                        # لو جدول واحد، نعتمد على آخر قسم تم رصده في الصفحة
                                        # سنضيفه للقسمين مؤقتاً لو لم يتحدد ونترك المستخدم يختار
                                        items_set["الاستقطاعات"].add(cell_s)
                                        items_set["الاستحقاقات"].add(cell_s)

            # تنقية القوائم النهائية
            for cat in items_set:
                final_list = []
                seen_norm = set()
                for item in items_set[cat]:
                    norm = normalize(item)
                    if norm not in seen_norm and len(norm) > 2:
                        final_list.append(item)
                        seen_norm.add(norm)
                self.items_db[cat] = sorted(final_list)

            self.lbl_file.config(text=f"تم المسح: {os.path.basename(self.pdf_path)}", fg="green")
            self.combo_cat.set("الاستقطاعات")
            self.update_item_list()
            
            total_items = len(self.items_db["الاستقطاعات"]) + len(self.items_db["الاستحقاقات"])
            if total_items == 0:
                messagebox.showwarning("تنبيه", "لم يتم العثور على أي بنود. تأكد أن الملف يحتوي على جداول نصوص وليس صورا.")
            else:
                messagebox.showinfo("تم المسح", f"تم العثور على {total_items} بند مالي مختلف.")
            
        except Exception as e:
            messagebox.showerror("خطأ في المسح", f"حدث خطأ أثناء قراءة الملف:\n{str(e)}")

    def update_item_list(self, event=None):
        cat = self.combo_cat.get()
        if cat in self.items_db:
            self.combo_item.config(state="readonly", values=self.items_db[cat])
            if self.items_db[cat]: self.combo_item.set(self.items_db[cat][0])
            else: self.combo_item.set("لا توجد بنود")

    def fix_ar(self, t):
        if not t: return ""
        try: return get_display(arabic_reshaper.reshape(str(t).strip()))
        except: return str(t)

    def process_selected(self):
        target_item = self.combo_item.get()
        category = self.combo_cat.get()
        if not target_item or target_item == "لا توجد بنود": return
        
        target_norm = normalize(target_item)
        
        try:
            for i in self.tree.get_children(): self.tree.delete(i)
            results = []
            
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    words = page.extract_words()
                    emp_name = "غير معروف"
                    emp_code = "غير معروف"

                    # 1. الكود: استخدام النص الفعلي لعدم دمج رقم الفرع مع الوقت
                    # نبحث عن 30200105 يليها مسافات اختيارية ثم شرطة ثم الكود
                    full_text = page.extract_text() or ""
                    
                    # تحويل Arabic Presentation Forms (ﻢﺳﺍ) للحروف العربية القياسية (اسم)
                    full_text_std = unicodedata.normalize('NFKC', full_text)
                    
                    full_text_norm = full_text_std.replace("\n", " ").translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
                    
                    m_code = re.search(r"30200105\s*-\s*(\d{3,6})", full_text_norm)
                    if m_code: 
                        emp_code = m_code.group(1)
                    else:
                        m_code_rev = re.search(r"(\d{3,6})\s*-\s*30200105", full_text_norm)
                        if m_code_rev: 
                            emp_code = m_code_rev.group(1)
                            
                    # 2. الاسم: البحث في النص بعد تحويله + عكسه (لأن PDF أوراكل يخزنه بترتيب بصري مقلوب)
                    for line in full_text_std.split("\n")[:30]:
                        # عكس السطر لتحويله من الترتيب البصري للترتيب المنطقي
                        line_reversed = line[::-1]
                        
                        if "اسم" in line_reversed or "موظف" in line_reversed:
                            for keyword in ["اسم الموظف:", "اسم الموظف", "الموظف:", "الموظف", "اسم:", "اسم"]:
                                if keyword in line_reversed:
                                    raw_name = line_reversed.split(keyword)[-1]
                                    # نقطع النص عند اصطدامه بأي بيانات أخرى
                                    for stop_word in ["رقم", "قومى", "نوع", "درج", "ادار", "فرع", "بنك", "تأمين"]:
                                        if stop_word in raw_name:
                                            raw_name = raw_name.split(stop_word)[0]
                                    # تنظيف ما تبقى ليكون هو الاسم الصافي
                                    cleaned_name = "".join([c for c in raw_name if not c.isdigit() and c not in [":", "-", "_", "*"]]).strip()
                                    if len(cleaned_name) >= 3:
                                        emp_name = cleaned_name
                                    break
                            if emp_name != "غير معروف":
                                break
                            
                    tables = page.extract_tables()
                    for t in tables:
                        # تحديد جهة البحث
                        start_idx = -1
                        for r in t:
                            for i, c in enumerate(r):
                                if c and category in str(c):
                                    start_idx = i
                                    break
                            if start_idx != -1: break

                        for row in t:
                            # البحث في الصف
                            area = row[start_idx:] if start_idx != -1 else row
                            row_txt = normalize(" ".join([str(c) for c in area if c]))
                            
                            if (target_norm in row_txt) or (target_norm[::-1] in row_txt):
                                amount = ""
                                for cell in area:
                                    if is_val(cell):
                                        amount = str(cell).strip()
                                        break
                                
                                # تجاهل الصفوف بدون قيمة مالية
                                if amount:
                                    self.tree.insert("", tk.END, values=(
                                        self.fix_ar(emp_name),
                                        to_hindi_nums(emp_code),
                                        to_hindi_nums(amount)
                                    ))
                                    results.append({"اسم الموظف": emp_name, "كود الموظف": emp_code, "المبلغ": amount})
                                    break

            if results:
                self.extracted_df = pd.DataFrame(results)
                self.btn_xlsx.config(state=tk.NORMAL)
                messagebox.showinfo("نجاح", f"تم العثور على {len(results)} سجلات.")
            else:
                messagebox.showinfo("نتيجة", "لم يتم العثور على أي موظف لديه هذا البند.")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء المعالجة:\n{str(e)}")

    def export(self):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if f:
            self.extracted_df.to_excel(f, index=False)
            messagebox.showinfo("نجاح", "تم التصدير.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SmartPayrollApp(root)
    root.mainloop()
