"""Financial regression tests. All identities and values here are fictional."""
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from payroll_core import import_page_texts, key, number, _parse_employee
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
        self.assertEqual(4,len(keys))
        self.assertEqual(3,sum(v.kind=="Informational" for v in keys.values()))
        # Distinct informational descriptions share an account code.
        shares=[v for v in r.employees[0].entries if v.kind=="Informational"]
        self.assertEqual(4,len(shares))
        self.assertEqual(3,len({v.choice_key for v in shares}))


if __name__=="__main__":
    unittest.main()
