"""Oracle PDF Extractor: verified payroll reading and flexible Arabic Excel export.

Tkinter is deliberately kept separate from the parser and XLSX formatter.
Start with: python oracle_payroll_extractor.py
"""
from __future__ import annotations

from decimal import Decimal
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from payroll_core import ImportReport, PayrollEmployee, import_pdf
from payroll_export import ExportSelection, available_items, make_workbook, save_excel

APP_BG="#f4f7fb"
NAVY="#173a56"
BLUE="#186b91"
INFO="#536b7c"


class SmartPayrollApp:
    def __init__(self, root: tk.Tk):
        self.root=root
        self.root.title("مستخرج مرتبات Oracle | التحقق والتصدير")
        self.root.geometry("1330x840")
        self.root.minsize(1020,680)
        self.root.configure(background=APP_BG)
        self.report: ImportReport | None=None
        self.pdf_path=""
        self.msg_queue: queue.Queue=queue.Queue()
        self.running=False
        self.employee_visible: list[int]=[]
        self.item_visible: list[str]=[]
        self.chosen_employee_indexes: set[int]=set()
        self.chosen_item_keys: set[str]=set()
        self.item_choices={}
        self.scope=tk.StringVar(value="all")
        self.gross=tk.BooleanVar(value=False)
        self.deductions=tk.BooleanVar(value=False)
        self.net=tk.BooleanVar(value=True)
        self.employee_search=tk.StringVar()
        self.item_search=tk.StringVar()
        self.status=tk.StringVar(value="اختر كشف صرفية موظف PDF لبدء استخراج البنود.")
        self.summary=tk.StringVar(value="لا توجد بيانات حتى الآن")
        self.build_ui()

    def build_ui(self):
        style=ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Treeview",rowheight=28,font=("Segoe UI",10))
        style.configure("Treeview.Heading",font=("Segoe UI",10,"bold"),foreground=NAVY)
        style.configure("Accent.TButton",font=("Segoe UI",10,"bold"),padding=8)
        style.configure("TLabel",background=APP_BG,font=("Segoe UI",10),foreground=NAVY)

        header=tk.Frame(self.root,bg=NAVY,padx=22,pady=17)
        header.pack(fill=tk.X)
        tk.Label(header,text="ORACLE PAYROLL EXTRACTOR",bg=NAVY,fg="white",
                 font=("Segoe UI",17,"bold")).pack(side=tk.RIGHT)
        tk.Label(header,text="استخراج بيانات الموظفين • تحقق مالي • Excel قابل للمراجعة",
                 bg=NAVY,fg="#d3e6f1",font=("Segoe UI",11)).pack(side=tk.LEFT)

        controls=tk.Frame(self.root,bg=APP_BG,padx=16,pady=10)
        controls.pack(fill=tk.X)
        self.open_btn=tk.Button(controls,text="١  اختيار ملف PDF",command=self.select_file,
                    bg=BLUE,fg="white",relief=tk.FLAT,padx=18,pady=10,font=("Segoe UI",11,"bold"))
        self.open_btn.pack(side=tk.RIGHT,padx=(8,0))
        self.progress=ttk.Progressbar(controls,length=210,mode="determinate")
        self.progress.pack(side=tk.RIGHT,padx=8)
        tk.Label(controls,textvariable=self.status,bg=APP_BG,fg=INFO,
                 font=("Segoe UI",10),anchor="e").pack(side=tk.RIGHT,fill=tk.X,expand=True)

        info=tk.Frame(self.root,bg="#e8f0f5",padx=15,pady=8)
        info.pack(fill=tk.X,padx=15)
        tk.Label(info,textvariable=self.summary,bg="#e8f0f5",fg=NAVY,
                 font=("Segoe UI",10,"bold"),anchor="e").pack(side=tk.RIGHT,fill=tk.X,expand=True)

        chooser=tk.Frame(self.root,bg=APP_BG,padx=13,pady=10)
        chooser.pack(fill=tk.BOTH,expand=False)
        chooser.columnconfigure(0,weight=1)
        chooser.columnconfigure(1,weight=1)

        people=tk.LabelFrame(chooser,text="٢  نطاق الموظفين",bg=APP_BG,fg=NAVY,
                             font=("Segoe UI",11,"bold"),padx=12,pady=10)
        people.grid(row=0,column=1,sticky="nsew",padx=6)
        tk.Radiobutton(people,text="كل الموظفين",variable=self.scope,value="all",
                       command=self._scope_changed,bg=APP_BG).pack(side=tk.TOP,anchor="e")
        tk.Radiobutton(people,text="موظف أو مجموعة موظفين مختارة",variable=self.scope,
                       value="selected",command=self._scope_changed,bg=APP_BG).pack(anchor="e")
        ep=tk.Frame(people,bg=APP_BG)
        ep.pack(fill=tk.X,pady=(5,4))
        e_search=tk.Entry(ep,textvariable=self.employee_search,justify=tk.RIGHT,
                          font=("Segoe UI",10))
        e_search.pack(side=tk.RIGHT,fill=tk.X,expand=True)
        tk.Label(ep,text="بحث:",bg=APP_BG).pack(side=tk.RIGHT,padx=4)
        self.employee_search.trace_add("write",lambda *_:self._render_people())
        er=tk.Frame(people,bg=APP_BG)
        er.pack(fill=tk.BOTH,expand=True)
        self.people_list=tk.Listbox(er,selectmode=tk.MULTIPLE,height=8,exportselection=False,
                         activestyle="none",font=("Segoe UI",10),justify=tk.RIGHT)
        self.people_list.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        sp=ttk.Scrollbar(er,orient=tk.VERTICAL,command=self.people_list.yview)
        sp.pack(side=tk.RIGHT,fill=tk.Y)
        self.people_list.configure(yscrollcommand=sp.set)
        self.people_list.bind("<<ListboxSelect>>",self._change_people)
        pb=tk.Frame(people,bg=APP_BG)
        pb.pack(fill=tk.X,pady=3)
        ttk.Button(pb,text="تحديد نتائج البحث",command=lambda:self._select_visible_people(True)).pack(side=tk.RIGHT)
        ttk.Button(pb,text="إلغاء تحديد النتائج",command=lambda:self._select_visible_people(False)).pack(side=tk.RIGHT,padx=6)

        details=tk.LabelFrame(chooser,text="٣  تحديد المطلوب من الاستحقاقات والاستقطاعات",
                              bg=APP_BG,fg=NAVY,font=("Segoe UI",11,"bold"),padx=12,pady=10)
        details.grid(row=0,column=0,sticky="nsew",padx=6)
        cb=tk.Frame(details,bg=APP_BG)
        cb.pack(fill=tk.X)
        for title,var in (("إجمالي الاستحقاقات الرسمي",self.gross),
                          ("إجمالي الاستقطاعات الرسمي",self.deductions),
                          ("صافي المرتب الرسمي",self.net)):
            tk.Checkbutton(cb,text=title,variable=var,bg=APP_BG,
                  command=self.preview,anchor="e").pack(side=tk.RIGHT,padx=8)
        q=tk.Frame(details,bg=APP_BG)
        q.pack(fill=tk.X,pady=7)
        tk.Entry(q,textvariable=self.item_search,justify=tk.RIGHT,
                 font=("Segoe UI",10)).pack(side=tk.RIGHT,fill=tk.X,expand=True)
        tk.Label(q,text="البنود:",bg=APP_BG).pack(side=tk.RIGHT,padx=4)
        self.item_search.trace_add("write",lambda *_:self._render_items())
        ir=tk.Frame(details,bg=APP_BG)
        ir.pack(fill=tk.BOTH,expand=True)
        self.items_list=tk.Listbox(ir,selectmode=tk.MULTIPLE,height=9,exportselection=False,
                         activestyle="none",font=("Segoe UI",10),justify=tk.RIGHT)
        self.items_list.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        si=ttk.Scrollbar(ir,orient=tk.VERTICAL,command=self.items_list.yview)
        si.pack(side=tk.RIGHT,fill=tk.Y)
        self.items_list.configure(yscrollcommand=si.set)
        self.items_list.bind("<<ListboxSelect>>",self._change_items)
        ib=tk.Frame(details,bg=APP_BG)
        ib.pack(fill=tk.X,pady=3)
        ttk.Button(ib,text="تحديد البنود الظاهرة",command=lambda:self._select_visible_items(True)).pack(side=tk.RIGHT)
        ttk.Button(ib,text="مسح التحديد",command=lambda:self._select_visible_items(False)).pack(side=tk.RIGHT,padx=6)

        actions=tk.Frame(self.root,bg=APP_BG,padx=20,pady=7)
        actions.pack(fill=tk.X)
        self.export_btn=tk.Button(actions,text="٤  تصدير Excel مخصص",command=self.export,
                     state=tk.DISABLED,bg="#20744a",fg="white",relief=tk.FLAT,
                     padx=25,pady=10,font=("Segoe UI",11,"bold"))
        self.export_btn.pack(side=tk.RIGHT,padx=7)
        ttk.Button(actions,text="معاينة الاختيار",command=self.preview).pack(side=tk.RIGHT,padx=5)
        tk.Label(actions,text="المبالغ غير المثبتة تظل فارغة، والمعلوماتي ليس أجرًا نقديًا.",
                 bg=APP_BG,fg=INFO).pack(side=tk.RIGHT,padx=12)

        preview_frame=tk.LabelFrame(self.root,text="معاينة التقرير قبل التصدير",bg=APP_BG,
                                    font=("Segoe UI",11,"bold"),fg=NAVY,padx=10,pady=8)
        preview_frame.pack(fill=tk.BOTH,expand=True,padx=19,pady=(4,14))
        self.tree=ttk.Treeview(preview_frame,columns=[],show="headings",height=7)
        self.tree.grid(row=0,column=0,sticky="nsew")
        scrollbar=ttk.Scrollbar(preview_frame,orient=tk.VERTICAL,command=self.tree.yview)
        scrollbar.grid(row=0,column=1,sticky="ns")
        horizontal=ttk.Scrollbar(preview_frame,orient=tk.HORIZONTAL,command=self.tree.xview)
        horizontal.grid(row=1,column=0,sticky="ew")
        self.tree.configure(yscrollcommand=scrollbar.set,xscrollcommand=horizontal.set)
        preview_frame.rowconfigure(0,weight=1)
        preview_frame.columnconfigure(0,weight=1)
        self._set_preview_columns(["اسم الموظف","الحالة"])

    def _set_preview_columns(self,headers:list[str]):
        cols=[f"col{i}" for i in range(len(headers))]
        self.tree["columns"]=cols
        for col,title in zip(cols,headers):
            self.tree.heading(col,text=title)
            self.tree.column(col,anchor=tk.E,width=max(135,min(330,8*len(title))),stretch=False)
        for row in self.tree.get_children():
            self.tree.delete(row)

    def select_file(self):
        filename=filedialog.askopenfilename(title="اختر صرفية موظف PDF",filetypes=[("PDF","*.pdf")])
        if not filename:
            return
        if self.running:
            return
        self.pdf_path=filename
        self.report=None
        self._reset_selection()
        self.running=True
        self.progress["value"]=0
        self.export_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.DISABLED)
        self.status.set("جارٍ تحليل صفحات الملف والتحقق من البنود ...")
        def worker():
            try:
                result=import_pdf(filename,lambda done,total:
                    self.msg_queue.put(("progress",done,total)))
                self.msg_queue.put(("done",result))
            except Exception as exc:
                self.msg_queue.put(("error",f"{type(exc).__name__}: {exc}"))
        threading.Thread(target=worker,daemon=True).start()
        self.root.after(80,self._poll)

    def _poll(self):
        try:
            while True:
                task=self.msg_queue.get_nowait()
                if task[0]=="progress":
                    _,done,total=task
                    self.progress.configure(maximum=max(1,total),value=done)
                    self.status.set(f"جارٍ فحص الصفحة {done} من {total}")
                elif task[0]=="done":
                    self.running=False
                    self.open_btn.config(state=tk.NORMAL)
                    self.report=task[1]
                    self.item_choices=available_items(self.report)
                    self._render_people()
                    self._render_items()
                    self.export_btn.config(state=tk.NORMAL if self.report.employees else tk.DISABLED)
                    verified=sum(e.verified for e in self.report.employees)
                    self.summary.set(f"المصدر: {os.path.basename(self.pdf_path)} | الموظفون: {len(self.report.employees)}"
                                     f" | بلا تحذيرات: {verified} | يحتاج مراجعة: {len(self.report.employees)-verified}"
                                     f" | البنود المختلفة: {len(self.item_choices)}")
                    self.status.set("اكتمل التحليل. راجع التحذيرات قبل الاعتماد المالي.")
                    self.preview()
                    if not self.report.verified:
                        messagebox.showwarning("تحتاج البيانات مراجعة",
                            "اكتمل استخراج البيانات، لكن توجد مبالغ أو هويات غير مثبتة.\n"
                            "لن تتحول المبالغ الناقصة إلى صفر في Excel. انظر صفحة «معلومات ومراجعة» بعد التصدير.")
                    return
                elif task[0]=="error":
                    self.running=False
                    self.open_btn.config(state=tk.NORMAL)
                    self.status.set("تعذر قراءة PDF")
                    messagebox.showerror("خطأ في قراءة الملف",task[1])
                    return
        except queue.Empty:
            pass
        self.root.after(80,self._poll)

    def _reset_selection(self):
        self.chosen_employee_indexes.clear()
        self.chosen_item_keys.clear()
        self.item_choices={}
        self.employee_visible=[]
        self.item_visible=[]
        self.employee_search.set("")
        self.item_search.set("")
        self.scope.set("all")

    def _scope_changed(self):
        self.preview()

    def _render_people(self):
        search=self.employee_search.get().strip()
        self.employee_visible=[]
        self.people_list.delete(0,tk.END)
        for i,e in enumerate(self.report.employees if self.report else []):
            haystack=" ".join((e.name,e.national_id or "",e.disbursement_number or "",
                               e.institutional_entity_code or ""))
            if search and search not in haystack:
                continue
            n=len(self.employee_visible)
            self.employee_visible.append(i)
            self.people_list.insert(tk.END,f"{e.name} — {e.disbursement_number or e.national_id or 'هوية غير مثبتة'}")
            if i in self.chosen_employee_indexes:
                self.people_list.selection_set(n)

    def _render_items(self):
        query=self.item_search.get().strip()
        self.item_visible=[]
        self.items_list.delete(0,tk.END)
        for token,entry in self.item_choices.items():
            label=f"{ {'Earning':'استحقاق','Deduction':'استقطاع','Informational':'معلوماتي'}.get(entry.kind,entry.kind) } | {entry.account_code or 'بلا حساب'} | {entry.label}"
            if query and query not in label:
                continue
            n=len(self.item_visible)
            self.item_visible.append(token)
            self.items_list.insert(tk.END,label)
            if token in self.chosen_item_keys:
                self.items_list.selection_set(n)

    def _change_people(self,event=None):
        active=set(self.people_list.curselection())
        for list_index,original_index in enumerate(self.employee_visible):
            if list_index in active:
                self.chosen_employee_indexes.add(original_index)
            else:
                self.chosen_employee_indexes.discard(original_index)
        if self.scope.get()=="selected":
            self.preview()

    def _change_items(self,event=None):
        active=set(self.items_list.curselection())
        for list_index,token in enumerate(self.item_visible):
            if list_index in active:
                self.chosen_item_keys.add(token)
            else:
                self.chosen_item_keys.discard(token)
        self.preview()

    def _select_visible_people(self,enabled:bool):
        self.scope.set("selected")
        for i in self.employee_visible:
            if enabled:self.chosen_employee_indexes.add(i)
            else:self.chosen_employee_indexes.discard(i)
        self._render_people()
        self.preview()

    def _select_visible_items(self,enabled:bool):
        for key in self.item_visible:
            if enabled:self.chosen_item_keys.add(key)
            else:self.chosen_item_keys.discard(key)
        self._render_items()
        self.preview()

    def _selection(self)->ExportSelection:
        return ExportSelection(
            employee_indexes=None if self.scope.get()=="all" else tuple(sorted(self.chosen_employee_indexes)),
            item_keys=tuple(sorted(self.chosen_item_keys)),
            gross=self.gross.get(),deductions=self.deductions.get(),net=self.net.get())

    def preview(self):
        if not self.report:
            return
        try:
            workbook=make_workbook(self.report,self._selection())
        except ValueError as e:
            self.status.set(str(e))
            self._set_preview_columns(["اسم الموظف","الحالة"])
            return
        sheet=workbook["ملخص الموظفين"]
        headers=[c.value for c in sheet[1]]
        self._set_preview_columns(headers)
        for row in sheet.iter_rows(min_row=2,max_row=min(sheet.max_row-1,201),values_only=True):
            values=[f"{v:,.2f}" if isinstance(v,Decimal) else v or "" for v in row]
            self.tree.insert("",tk.END,values=values)
        total=self.report.employees if self.scope.get()=="all" else self.chosen_employee_indexes
        self.status.set(f"المعاينة: {len(total)} موظف و {len(self.chosen_item_keys)} بند محدد"
                        " | أي قيمة ناقصة في الكشف تحتاج مراجعة")

    def export(self):
        if not self.report:
            messagebox.showwarning("تنبيه","يرجى تحليل PDF أولاً")
            return
        try:
            selection=self._selection()
            # Validate before opening a save dialog.
            make_workbook(self.report,selection)
        except ValueError as e:
            messagebox.showwarning("اختيارات غير مكتملة",str(e))
            return
        target=filedialog.asksaveasfilename(
            title="حفظ تقرير مرتبات Oracle",defaultextension=".xlsx",
            initialfile="كشف_مرتبات_Oracle_مخصص.xlsx",filetypes=[("ملف Excel","*.xlsx")])
        if not target:
            return
        try:
            save_excel(target,self.report,selection)
            messagebox.showinfo("تم الحفظ",
                "تم تصدير ملف Excel بثلاث صفحات: الملخص والتفاصيل والمراجعة.\n"
                "الملف يحتوي بيانات شخصية ومالية؛ احتفظ به في مكان آمن.")
        except (PermissionError,ValueError,OSError) as e:
            messagebox.showerror("تعذر تصدير Excel",str(e))


if __name__=="__main__":
    root=tk.Tk()
    SmartPayrollApp(root)
    root.mainloop()
