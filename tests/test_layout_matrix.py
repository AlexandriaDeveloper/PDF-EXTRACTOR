"""Layout matrix: same payroll meaning across independent PDF presentations.

Every case uses wholly synthetic identities and amounts, not actual payslips.
"""
import copy
import unittest
from decimal import Decimal
from payroll_oracle_tables import parse_document,extract_page
from test_verified_oracle_tables import first_page,second_page,mirror

def reformat(*,swapped=False,wide=False,reverse_cells=False,plain_arabic=False,
             codes=False,extra_page=False,govt_deduction=False):
    first=copy.deepcopy(first_page())
    second=copy.deepcopy(second_page())
    if codes:
        # Ledger identifiers are data; no special knowledge of 211/411/113 codes.
        for table in first['oracle_tables']+second['oracle_tables']:
            for row in table:
                for i,value in enumerate(row):
                    if value=="21110442":row[i]="98765432"
                    if value=="21110328":row[i]="87654321"
                    if value=="41110201":row[i]="76543210"
                    if value=="11320129":row[i]="65432109"
    if govt_deduction:
        # A government-labelled amount *within the deductions section*
        # is a deduction, not an additional informational credit.
        ded=first['oracle_tables'][1]
        for row in ded:
            if "11320129" in row:
                row[2]=mirror("حكومة تأمين الموظف")
    if swapped:
        first['oracle_tables'].reverse()
    if wide:
        for page in (first,second):
            for table in page['oracle_tables']:
                for row in table:
                    row.insert(0,None)
                    row.append(None)
    if reverse_cells:
        for page in (first,second):
            for table in page['oracle_tables']:
                for row in table:
                    row.reverse()
    if plain_arabic:
        for page in (first,second):
            # The mirrored test fixtures use the same reversal function that
            # provides the logical form on its second application.
            page['oracle_original_text']="\n".join(
                mirror(s) for s in page['oracle_original_text'].splitlines()
            )
            for table in page['oracle_tables']:
                for row in table:
                    for i,value in enumerate(row):
                        if isinstance(value,str) and any("\u0621"<=c<="\u06ff" for c in value):
                            row[i]=mirror(value)
    result=[first,second]
    if extra_page:
        result.insert(1,{"oracle_original_text":"تقرير الرواتب صفحة إضافية",
                         "oracle_tables":[[["معلومات","بدون بنود"]]]})
    return result

class LayoutMatrixTests(unittest.TestCase):
    def verified(self,pages):
        employees,diags=parse_document(pages)
        self.assertEqual(len(employees),1)
        one=employees[0]
        self.assertEqual(one['status'],'Verified',one['warnings'])
        self.assertEqual(one['printed_net'],Decimal('120'))
        self.assertEqual(sum(e['amount'] for e in one['entries']
                             if e['kind']=='الاستحقاقات'),Decimal('150'))
        self.assertEqual(sum(e['amount'] for e in one['entries']
                             if e['kind']=='الاستقطاعات'),Decimal('42'))
        self.assertEqual(diags['tabular_verified'],1)
        return one
    def test_tables_swapped_does_not_change_financial_category(self):
        self.verified(reformat(swapped=True))
    def test_extra_empty_columns_do_not_change_financial_category(self):
        self.verified(reformat(wide=True))
    def test_account_codes_are_not_a_fixed_taxonomy(self):
        emp=self.verified(reformat(codes=True))
        self.assertIn(("الاستحقاقات","98765432","بدل الجامعة"),emp['items'])
        self.assertIn(("الاستقطاعات","65432109","عمولة تحصيل"),emp['items'])
    def test_extra_page_does_not_break_record_boundary(self):
        emp=self.verified(reformat(extra_page=True))
        self.assertEqual(emp['pages'],[1,2,3])
    def test_reversed_table_column_order_is_supported(self):
        self.verified(reformat(reverse_cells=True))
    def test_natural_unicode_arabic_is_supported(self):
        self.verified(reformat(plain_arabic=True))
    def test_government_share_in_deduction_section_is_not_informational(self):
        emp=self.verified(reformat(govt_deduction=True))
        row=next(e for e in emp['entries'] if e['account_code']=="11320129")
        self.assertEqual(row['kind'],'الاستقطاعات')
    def test_separate_header_lines_do_not_require_one_specific_print_layout(self):
        pages=reformat()
        pages[0]["oracle_original_text"]=(
            mirror("نظام الرواتب")+"\n"
            +mirror("اسم الموظف: موظف اختباري")+"\n"
            +mirror("رقم الصرف: 30200105-00001")+"\n"
            +mirror("الرقم القومي: 00000000000001")
        )
        employee=self.verified(pages)
        self.assertEqual(employee["disbursement_number"],"30200105-00001")

    def test_mixed_earnings_and_deductions_in_single_physical_table(self):
        header=("اسم الموظف: موظف اختباري رقم الصرف: 30200105-00001 "
                "الرقم القومي: 00000000000001")
        table=[
            ["",mirror("الاستقطاعات"),None,None,
             "",mirror("الاستحقاقات"),None,None],
            ["20.00",mirror("خصم تجريبي"),"76543210","",
             "100.00",mirror("بدل الجامعة"),"98765432",""]
        ]
        page=extract_page(mirror(header),[table],1)
        self.assertEqual(len(page["entries"]),2)
        items={(x["account_code"],x["kind"],x["amount"]) for x in page["entries"]}
        self.assertIn(("76543210","الاستقطاعات",Decimal('20')),items)
        self.assertIn(("98765432","الاستحقاقات",Decimal('100')),items)
    def test_incomplete_second_employee_header_blocks_cross_employee_merging(self):
        first,continuation=reformat()
        broken=copy.deepcopy(first)
        broken["oracle_original_text"]=(
            mirror("اسم الموظف: موظف ثاني رقم الصرف: 30200105-00002")
            +"\n"+mirror("الرقم القومي: غير متاح")
        )
        result,diagnostics=parse_document([first,continuation,broken])
        self.assertEqual(len(result),1)
        self.assertEqual(result[0]["pages"],[1,2])
        self.assertEqual(result[0]["status"],"ReviewRequired")
        self.assertEqual(diagnostics["unrecognized_employee_headers"],1)

    def test_bad_total_remains_review(self):
        pages=reformat(codes=True,swapped=True)
        pages[1]['oracle_tables'][0][1][1]="119.00"
        employees,diags=parse_document(pages)
        self.assertEqual(employees[0]['status'],'ReviewRequired')
        self.assertEqual(diags['tabular_review'],1)

if __name__=='__main__':
    unittest.main()
