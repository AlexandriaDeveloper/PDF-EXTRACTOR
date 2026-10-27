"""Regression cases for the real Oracle *shape*: mirrored Arabic and two panes.

Uses fabricated employee details only. No real private salary data is stored.
"""
import unittest
from decimal import Decimal
from payroll_table_adapter import extract_table_candidates, reverse_visual
from payroll_recovery import recover_records
from payroll_core import EARNING, DEDUCTION, INFORMATIONAL, UNKNOWN, identifier

def visual(text):
    # Reversed glyph order with number runs restored to represent pdfplumber's
    # visual-order Arabic extraction.
    return reverse_visual(text)

class FakeOraclePage:
    page_number=1
    def __init__(self):
        self.heading=("اسم الموظف: موظف اختباري محمد علي "
                      "رقم الصرف: 30200105-00001 الرقم القومي: 00000000000001")
        self.lines=[
            "نظام أوراكل للرواتب",
            self.heading,
            "الاستقطاعات الاستحقاقات",
            "الاسم القيمة الاسم القيمة",
            "21110442",
            "بدل الجامعة",
            "100.00",
            "41110201",
            "خصم اختباري",
            "40.00",
        ]
        self._tables=[
            [[visual(self.heading)]],
            [[visual("الاستقطاعات"), "", "", visual("الاستحقاقات"), "", ""],
             [visual("41110201"),visual("خصم اختباري"),"40.00",
              visual("21110442"),visual("بدل الجامعة"),"100.00"]]
        ]
    def extract_text(self):
        return "\n".join(visual(s) for s in self.lines)
    def extract_tables(self):
        return self._tables

class RecoveryTests(unittest.TestCase):
    def test_page_reader_collects_table_evidence_from_visual_arabic(self):
        from oracle_payroll_extractor import parse_page
        page=FakeOraclePage()
        result=parse_page(page)
        self.assertEqual(len(result["table_items"]),2)
        self.assertIn("اسم الموظف",result["raw_text"])
        self.assertIn("30200105-00001",result["raw_text"])
        records,diag=recover_records([result])
        self.assertTrue(diag["recovery_used"])
        self.assertEqual(len(records),1)
        self.assertEqual(len(records[0]["entries"]),2)

    def test_two_pane_tables_produce_two_separate_candidates(self):
        page=FakeOraclePage()
        rows=extract_table_candidates(page)
        self.assertEqual(len(rows),2)
        keys={(r["kind"],r["account_code"],r["amount"]) for r in rows}
        self.assertIn((EARNING,"21110442",Decimal("100.00")),keys)
        self.assertIn((DEDUCTION,"41110201",Decimal("40.00")),keys)
        self.assertTrue(all(r["account_code"]!="30200105" for r in rows))
    def test_two_pane_items_without_section_heading_in_same_table(self):
        page=FakeOraclePage()
        page._tables=[
            [[visual(page.heading)]],
            [[visual("41110201"),visual("خصم اختباري"),"40.00",
              visual("21110442"),visual("بدل الجامعة"),"100.00"]]
        ]
        items=extract_table_candidates(page)
        self.assertEqual(len(items),2)
        self.assertEqual({x["account_code"] for x in items},{"21110442","41110201"})
        self.assertEqual({x["kind"] for x in items},{EARNING,DEDUCTION})

    def test_semantic_zero_is_recovered_as_explicit_review(self):
        page=FakeOraclePage()
        # text deliberately contains split code/label/amount lines. The strict
        # financial parser cannot link them on one line.
        page_dict={
            "raw_text":"\n".join(page.lines),
            "table_items":extract_table_candidates(page),
            "legacy_name":UNKNOWN,"legacy_code":UNKNOWN,
            "text_chars":len(page.extract_text())
        }
        employees,diagnostics=recover_records([page_dict])
        self.assertEqual(len(employees),1)
        self.assertEqual(len(employees[0]["items"]),2)
        self.assertEqual(employees[0]["national_id"],"00000000000001")
        self.assertEqual(employees[0]["institutional_code"],"30200105")
        self.assertEqual(employees[0]["status"],"ReviewRequired")
        self.assertTrue(diagnostics["recovery_used"])
        self.assertEqual(diagnostics["table_candidates"],2)
    def test_separate_employee_headers_are_not_merged(self):
        original=FakeOraclePage()
        first={
            "raw_text":"\n".join(original.lines),
            "table_items":extract_table_candidates(original),
            "legacy_name":UNKNOWN,"legacy_code":UNKNOWN,
            "text_chars":250
        }
        second_text="\n".join(original.lines).replace(
            "30200105-00001","30200105-00002").replace(
            "00000000000001","00000000000002").replace(
            "موظف اختباري محمد علي","موظف اختباري ثاني")
        second={**first,"raw_text":second_text}
        records,d=recover_records([first,second])
        self.assertEqual(len(records),2)
        self.assertEqual(records[0]["pages"],[1])
        self.assertEqual(records[1]["pages"],[2])
        self.assertNotEqual(records[0]["national_id"],records[1]["national_id"])
    def test_reversed_payment_number_is_restored_as_institutional_prefix(self):
        info=identifier(
            "اسم الموظف: موظف اختباري "
            "رقم الصرف: 00001-30200105 "
            "الرقم القومي: 00000000000001"
        )
        self.assertEqual(info["disbursement_number"],"30200105-00001")
        self.assertEqual(info["institutional_code"],"30200105")
        self.assertEqual(info["code"],"00001")
        self.assertEqual(info["national_id"],"00000000000001")

    def test_empty_pdf_keeps_zero_and_never_claims_success(self):
        result,diagnostics=recover_records([
            {"raw_text":"Document without payslip data","table_items":[],
             "legacy_name":UNKNOWN,"legacy_code":UNKNOWN,"text_chars":25}
        ])
        self.assertEqual(result,[])
        self.assertFalse(diagnostics["recovery_used"])
        self.assertEqual(diagnostics["semantic_items"],0)

if __name__=="__main__":unittest.main()
