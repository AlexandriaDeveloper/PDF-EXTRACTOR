"""Financial regression tests. All identities and values here are fictional."""
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from payroll_core import import_page_texts, import_pdf, key, number, _parse_employee, _page_stream
from payroll_export import ExportSelection, available_items, make_workbook, save_excel

NID_A="00000000000001"
NID_B="00000000000002"
PAY_A="30200105-00001"
PAY_B="30200105-00002"


def header(name="موظف تجريبي محمد علي", national=NID_A, pay=PAY_A):
    return (f"نظام أوراكل للرواتب\nاسم الموظف: {name} رقم الصرف: {pay} "
            f"الرقم القومي: {national}\nنوع الصرف: دفع الكتروني "
            "البنك: بنك مصر الفرع: مركز البطاقات\n"
            "الرقم التاميني: 00123456 الدرجة: مدير عام الادارة: كلية الطب\n")


def initial(name="موظف تجريبي محمد علي", national=NID_A, pay=PAY_A):
    return header(name,national,pay)+(
        "الاستحقاقات\n"
        "21110442 بدل الجامعة 100.00\n"
        "اجمالي الاستحقاقات 100.00\n"
        "21120202 حكومة 1 اصابة عمل 5.00\n"
        "21120202 حكومة 0.25% اصابة عمل 5.00\n"
        "21120201 حكومة 3 تامين صحي 5.00\n"
        "الاستقطاعات\n"
        "41110201 مؤمن 9 شيخوخة 40.00\n"
        "اجمالي الاستقطاعات 40.00\n"
    )


def continuation():
    return ("نظام أوراكل للرواتب\n"
            "21120101 حكومة 12 شيخوخة 5.00\n"
            "الصافي: 80.00\n"
            "هذا البيان للمراجعة فقط\n")


def valid_report():
    # Government-share entries intentionally have repeated account codes,
    # but different descriptions and valid two-decimal amounts.
    return import_page_texts("fixture.pdf", [initial(), continuation()])


class ParserTests(unittest.TestCase):
    def test_money_unicode_and_fail_closed(self):
        self.assertEqual(Decimal("15449.03"),number("١٥،٤٤٩٫٠٣"))
        self.assertEqual(Decimal("15449.03"),number("۱۵۴۴۹٫۰۳"))
        self.assertIsNone(number("wrong"))
        self.assertIsNone(number("29001131804917 ج"))
        self.assertEqual("الجامعه",key("اَلْجَامِعَة"))

    def test_identity_metadata_cannot_be_payroll_item(self):
        report=valid_report()
        self.assertEqual(1,len(report.employees))
        e=report.employees[0]
        self.assertEqual("30200105",e.institutional_entity_code)
        self.assertEqual(PAY_A,e.disbursement_number)
        self.assertEqual(NID_A,e.national_id)
        self.assertTrue(e.verified,e.warnings)
        self.assertEqual(Decimal("80.00"),e.net)
        self.assertEqual(2,e.end_page)
        self.assertEqual(1,len([x for x in e.entries if x.kind=="Earning"]))
        self.assertEqual(1,len([x for x in e.entries if x.kind=="Deduction"]))
        info=[x for x in e.entries if x.kind=="Informational"]
        self.assertEqual(4,len(info))
        self.assertNotIn("30200105",[x.account_code for x in e.entries])
        self.assertNotIn(NID_A," ".join(x.raw_text for x in e.entries))
        self.assertEqual(Decimal("20.00"),sum((x.amount for x in info),Decimal(0)))

    def test_repeated_header_is_continuation_and_new_id_starts_new_employee(self):
        p1=initial()
        p2=header() + continuation()
        p3=initial("موظف تجريبي احمد علي",NID_B,PAY_B)+"الصافي: 80.00"
        report=import_page_texts("synthetic.pdf",[p1,p2,p3])
        self.assertEqual(2,len(report.employees))
        self.assertEqual([NID_A,NID_B],[x.national_id for x in report.employees])
        self.assertEqual((1,2),(report.employees[0].start_page,report.employees[0].end_page))

    def test_unproven_totals_are_not_verified(self):
        text=initial()+"الصافي: 88.00"
        report=import_page_texts("wrong.pdf",[text,continuation()])
        self.assertFalse(report.verified)
        self.assertTrue(any("الصافي" in w for w in report.employees[0].warnings))

    def test_different_employee_never_silently_merged(self):
        report=import_page_texts("two.pdf",[
            initial()+"الصافي: 80.00",
            initial("موظف تجريبي احمد علي",NID_B,PAY_B)+"الصافي: 80.00",
        ])
        self.assertEqual(2,len(report.employees))
        self.assertNotEqual(report.employees[0].stable_id,report.employees[1].stable_id)

    def test_orphan_page_warns_without_associating_random_employee(self):
        r=import_page_texts("orphan.pdf",[
            "21110442 بدل الجامعة 90.00\nالصافي: 90.00",
            initial()+"الصافي: 80.00",
        ])
        self.assertEqual(1,len(r.employees))
        self.assertEqual(2,r.employees[0].start_page)
        self.assertTrue(r.warnings)

    def test_trailing_account_can_be_extracted(self):
        txt=header() + (
            "الاستحقاقات\n21110442 بدل الجامعة 100.00\nاجمالي الاستحقاقات 100.00\n"
            "الاستقطاعات\nعمولة تحصيل 0.50 11320129\n"
            "41110201 مؤمن 9 شيخوخة 39.50\nاجمالي الاستقطاعات 40.00\n"
            "الصافي: 60.00"
        )
        e=import_page_texts("fee.pdf",[txt]).employees[0]
        self.assertIn("11320129",[x.account_code for x in e.entries])
        self.assertEqual(Decimal("40"),sum((x.amount for x in e.entries if x.kind=="Deduction"),Decimal(0)))

    def test_missing_fields_cannot_become_verified(self):
        r=import_page_texts("empty.pdf",["نظام أوراكل للرواتب\nالاستحقاقات\n21110442 بدل 100.00"])
        self.assertFalse(r.verified)
        self.assertEqual(0,len(r.employees))


class PdfSourceTests(unittest.TestCase):
    def test_uses_structured_table_only_when_more_semantically_complete(self):
        class Page:
            def extract_text(self):
                return "noise\nabcde"
            def extract_tables(self):
                return [[
                    ["اسم الموظف: موظف تجريبي محمد علي رقم الصرف: 30200105-00001 الرقم القومي: 00000000000001"],
                    ["الاستحقاقات"],
                    ["21110442","بدل الجامعة","100.00"],
                ]]
        stream=_page_stream(Page())
        self.assertIn("21110442",stream)
        self.assertIn("الاستحقاقات",stream)

    def test_pdf_smoke_reads_real_synthetic_pdf_binary(self):
        # Construct a tiny non-sensitive PDF without third-party PDF generators.
        contents=b"BT /F1 12 Tf 72 700 Td (PAYROLL SMOKE SAMPLE) Tj ET"
        objects=[
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Length "+str(len(contents)).encode()+b" >>\nstream\n"+contents+b"\nendstream",
        ]
        data=bytearray(b"%PDF-1.4\n")
        positions=[0]
        for i,obj in enumerate(objects,1):
            positions.append(len(data))
            data.extend(f"{i} 0 obj\n".encode()+obj+b"\nendobj\n")
        xref=len(data)
        data.extend(f"xref\n0 {len(objects)+1}\n".encode())
        data.extend(b"0000000000 65535 f \n")
        for position in positions[1:]:
            data.extend(f"{position:010d} 00000 n \n".encode())
        data.extend(b"trailer\n<< /Size 6 /Root 1 0 R >>\n")
        data.extend(f"startxref\n{xref}\n%%EOF\n".encode())
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/"fixture.pdf"
            p.write_bytes(data)
            report=import_pdf(p)
            self.assertEqual(1,report.pages)
            self.assertEqual([],report.employees)
            self.assertFalse(report.verified)


    def test_table_cells_restore_sections_totals_and_identity_without_duplicates(self):
        from payroll_tables import extract_table_semantics
        def visual(value):
            # Simulated pdfplumber presentation-order Arabic.
            text=value[::-1]
            import re
            return re.sub(r"\d+(?:[.,٫،]\d+)*",lambda m:m.group()[::-1],text)
        class FakePage:
            def extract_tables(self):
                return [
                    [
                        [visual("اسم الموظف: موظف تجريبي محمد علي رقم الصرف: 00001-30200105 الرقم القومي: 00000000000001"),None,None,None,None],
                        [None,visual("الاستحقاقات"),None,None,None],
                        [None,visual("القيمة"),None,None,None],
                        [None,"100.00",None,visual("بدل الجامعة"),"21110442"],
                        [None,"100.00",None,visual("اجمالي الاستحقاقات"),None],
                        [None,"5.00",None,visual("حكومة 1 إصابة عمل"),"21120202"],
                    ],
                    [
                        [visual("الاستقطاعات"),None,None,None],
                        [visual("القيمة"),None,None,None],
                        [None,"40.00",visual("مؤمن 9 شيخوخة"),"41110201"],
                        ["40.00",None,visual("اجمالي الاستقطاعات"),None],
                    ]
                ]
        text=extract_table_semantics(FakePage())
        self.assertIsNotNone(text)
        self.assertIn("الاستحقاقات",text)
        self.assertIn("الاستقطاعات",text)
        self.assertEqual(1,text.count("21110442"))
        self.assertIn("اجمالي الاستحقاقات 100.00",text)
        self.assertIn("اجمالي الاستقطاعات 40.00",text)

    def test_unclassified_continuation_is_promoted_only_with_official_gross_match(self):
        first=(header()+"الاستحقاقات\n21110442 بدل الجامعة 100.00\n"
               "اجمالي الاستحقاقات 115.00\nالاستقطاعات\n"
               "41110201 مؤمن 9 شيخوخة 40.00\nاجمالي الاستقطاعات 40.00")
        second="21110443 حافز تكميلي 15.00 [قسم غير مؤكد]\nالصافي: 75.00"
        report=import_page_texts("supplement.pdf",[first,second])
        employee=report.employees[0]
        self.assertEqual(Decimal("115"),sum((e.amount for e in employee.entries
                          if e.kind=="Earning"),Decimal(0)))
        self.assertTrue(any("صفحات تكميلية" in w for w in employee.warnings))
        self.assertFalse(report.verified)

class ExcelTests(unittest.TestCase):
    def test_selective_export_three_sheets_numeric_and_text_codes(self):
        report=valid_report()
        e=report.employees[0]
        choices=available_items(report)
        earning=next(k for k,x in choices.items() if x.kind=="Earning")
        w=make_workbook(report,ExportSelection((0,),(earning,),True,True,True))
        self.assertEqual(["ملخص الموظفين","تفاصيل البنود","معلومات ومراجعة"],w.sheetnames)
        ws=w.active
        self.assertTrue(ws.sheet_view.rightToLeft)
        self.assertEqual(NID_A,ws.cell(2,2).value)
        self.assertEqual(PAY_A,ws.cell(2,3).value)
        self.assertEqual("30200105",ws.cell(2,4).value)
        self.assertEqual(Decimal("100.00"),ws.cell(2,6).value)
        self.assertEqual(Decimal("40.00"),ws.cell(2,7).value)
        self.assertEqual(Decimal("80.00"),ws.cell(2,8).value)
        self.assertEqual(Decimal("100.00"),ws.cell(2,9).value)
        self.assertGreater(w["تفاصيل البنود"].max_row,1)
        with tempfile.TemporaryDirectory() as d:
            p=save_excel(Path(d)/"out.xlsx",report,ExportSelection((0,),(earning,),True,False,True))
            reopened=load_workbook(p,data_only=False)
            self.assertEqual(100,reopened.active.cell(2,6).value)
            self.assertEqual(NID_A,reopened.active.cell(2,2).value)

    def test_invalid_selection_rejected(self):
        report=valid_report()
        with self.assertRaises(ValueError):
            make_workbook(report,ExportSelection((),(),False,False,True))
        with self.assertRaises(ValueError):
            make_workbook(report,ExportSelection(None,("unknown",),False,False,False))
        with self.assertRaises(ValueError):
            make_workbook(report,ExportSelection(None,(),False,False,False))

    def test_review_absences_are_blank_not_zero(self):
        report=valid_report()
        e=report.employees[0]
        k=next(k for k,x in available_items(report).items() if x.kind=="Earning")
        from payroll_core import PayrollEmployee
        report.employees.append(PayrollEmployee(name="موظف وهمي",entries=list(e.entries)))
        e.entries=[v for v in e.entries if v.choice_key!=k]
        e.warnings.append("مصدر غير مؤكد")
        ws=make_workbook(report,ExportSelection(None,(k,),False,False,False)).active
        self.assertIsNone(ws.cell(2,6).value)

    def test_excel_untrusted_text_never_becomes_formula(self):
        report=valid_report()
        report.employees[0].name="=HYPERLINK(\"https://example.invalid\",\"x\")"
        workbook=make_workbook(report,ExportSelection(net=True))
        cell=workbook.active.cell(2,1)
        self.assertEqual("s",cell.data_type)
        self.assertEqual(report.employees[0].name,cell.value)

    def test_distinguish_informational_and_deductions(self):
        r=valid_report()
        keys=available_items(r)
        self.assertEqual(6,len(keys))
        self.assertEqual(4,sum(v.kind=="Informational" for v in keys.values()))
        # Distinct informational descriptions share an account code.
        shares=[v for v in r.employees[0].entries if v.kind=="Informational"]
        self.assertEqual(4,len(shares))
        self.assertEqual(4,len({v.choice_key for v in shares}))


if __name__=="__main__":
    unittest.main()
