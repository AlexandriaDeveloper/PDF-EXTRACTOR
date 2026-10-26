"""Private verification against LOCAL Oracle files, without exposing identities.

This verifier prints only record counts and available aggregate official net.
"""
from __future__ import annotations
import argparse
from decimal import Decimal
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from payroll_core import import_pdf


def main():
    parser=argparse.ArgumentParser(description="Verify a local Oracle payslip PDF without leaking employee identity")
    parser.add_argument("pdf",type=Path)
    parser.add_argument("--expect-employees",type=int)
    parser.add_argument("--expect-net",type=Decimal)
    parser.add_argument("--strict",action="store_true",help="reject all review warnings")
    args=parser.parse_args()
    if not args.pdf.is_file():
        parser.error("PDF file not found")

    report=import_pdf(args.pdf)
    complete=sum(e.net is not None for e in report.employees)
    warning_count=len(report.warnings)+sum(len(e.warnings) for e in report.employees)
    total=sum((e.net for e in report.employees if e.net is not None),Decimal("0"))
    verified=sum(e.verified for e in report.employees)
    print(f"Pages: {report.pages}; employees: {len(report.employees)}; "
          f"with official net: {complete}; verified employees: {verified}")
    print(f"Available official net: {total:.2f}; warnings: {warning_count}")
    ok=True
    if args.expect_employees is not None and args.expect_employees!=len(report.employees):
        print("FAIL: employee-count reference mismatch")
        ok=False
    if args.expect_net is not None and abs(args.expect_net-total)>Decimal("0.02"):
        print("FAIL: official net reference mismatch")
        ok=False
    if args.strict and warning_count:
        print("FAIL: review warnings exist")
        ok=False
    if args.expect_employees is None or args.expect_net is None:
        print("NOTE: independent controls were not both supplied; this is NOT proof of accounting accuracy")
    return 0 if ok else 1


if __name__=="__main__":
    raise SystemExit(main())
