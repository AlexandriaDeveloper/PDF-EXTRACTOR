"""Oracle split-table regression tests (fabricated identities and amounts only)."""
import unittest
from decimal import Decimal
from payroll_oracle_tables import (
    logical,parse_header,parse_row,parse_totals,parse_document,parse_money
)

def mirror(s):
    # Same reverser is involutive for Arabic text and numeric identifiers.
    return logical(s)

def first_page(*, disbursement="30200105-00001",national="00000000000001"):
    header=("اسم الموظف: موظف اختباري رقم الصرف: "+disbursement+
            " الرقم القومي: "+national)
    # Oracle's actual 6-column credit table, 4-column deduction table:
    credits=[
       ["",mirror("الاستحقاقات"),None,None,None,None],
       ["", "100.00", "", mirror("بدل الجامعة"),"21110442",""],
       ["", "50.00", "", mirror("حافز اختبار"),"21110328",""],
       ["", "150.00", "", mirror("اجمالي الاستحقاقات"),"",""],
       ["", "", "5.00",mirror("حكومة 1 إصابة عمل"),"21120202",""],
       ["", "7.00", "",mirror("x إجمالي فرعي"),"",""]
    ]
    deductions=[
       ["",mirror("الاستقطاعات"),None,None],
       ["", "20.00",mirror("مؤمن 9 شيخوخة"),"41110201"],
       ["22.00", "",mirror("عمولة تحصيل"),"11320129"],
       ["42.00", "",mirror("اجمالي الاستقطاعات"),""],
    ]
    return {"oracle_original_text":mirror("نظام أوراكل للرواتب")+"\n"+mirror(header),
            "oracle_tables":[credits,deductions]}

def second_page():
    return {"oracle_original_text":mirror("نظام أوراكل للرواتب"),
        "oracle_tables":[[
            ["","","","7.00","",mirror("حكومة 12 شيخوخة"),"21120101",""],
            ["","120.00",mirror("الصافي:"),"","","","",""],
        ]]}

class RealLayoutTests(unittest.TestCase):
    def test_full_two_page_record_reconciles_with_one_hundred_percent_evidence(self):
        result,diag=parse_document([first_page(),second_page()])
        self.assertEqual(len(result),1)
        employee=result[0]
        self.assertEqual(employee["national_id"],"00000000000001")
        self.assertEqual(employee["institutional_code"],"30200105")
        self.assertEqual(employee["disbursement_number"],"30200105-00001")
        self.assertEqual(employee["status"],"Verified")
        self.assertEqual(len(employee["entries"]),5)
        self.assertEqual(employee["printed_gross"],Decimal("150.00"))
        self.assertEqual(employee["printed_deductions"],Decimal("42.00"))
        self.assertEqual(employee["printed_net"],Decimal("120.00"))
        self.assertEqual(diag["tabular_verified"],1)
        self.assertEqual(diag["tabular_items"],5)
        self.assertIn(("الاستقطاعات","11320129","عمولة تحصيل"),employee["items"])
        self.assertNotIn("30200105",[e["account_code"] for e in employee["entries"]])
    def test_bad_printed_net_fails_closed(self):
        page2=second_page()
        page2["oracle_tables"][0][1][1]="119.00"
        result,diag=parse_document([first_page(),page2])
        self.assertEqual(result[0]["status"],"ReviewRequired")
        self.assertEqual(diag["tabular_review"],1)
        self.assertTrue(any("الصافي" in w for w in result[0]["warnings"]))
    def test_missing_deduction_is_review_not_zero(self):
        page=first_page()
        page["oracle_tables"][1][2]=["","","",""]
        result,diag=parse_document([page,second_page()])
        self.assertEqual(result[0]["status"],"ReviewRequired")
    def test_multiple_employees_maintain_identity_boundaries(self):
        result,diag=parse_document([first_page(),second_page(),
            first_page(disbursement="30200105-00002",national="00000000000002"),
            second_page()])
        self.assertEqual(len(result),2)
        self.assertEqual(diag["tabular_verified"],2)
        self.assertEqual([x["pages"] for x in result],[[1,2],[3,4]])
        self.assertNotEqual(result[0]["national_id"],result[1]["national_id"])
    def test_unknown_layout_not_reported_verified(self):
        result,diag=parse_document([
            {"oracle_original_text":"Hello","oracle_tables":[[["some random text","3.50"]]]}
        ])
        self.assertEqual(result,[])
        self.assertFalse(diag["tabular_recognized"])
    def test_arabic_decimal_and_identifiers(self):
        self.assertEqual(parse_money("١٢،٦٠١٫٩٩"),Decimal("12601.99"))
        self.assertIsNone(parse_money("٣٠٢٠٠١٠٥"))
        self.assertIsNone(parse_money("٠٠٠٠٠٠٠٠٠٠٠٠٠١"))

if __name__=="__main__":unittest.main()
