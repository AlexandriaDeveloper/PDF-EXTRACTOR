# PDF-EXTRACTOR — Oracle payroll payslips

Windows desktop utility built with Python, Tkinter and pdfplumber. Reads Oracle salary statements and exports employee/item-specific Excel workbooks.

## What changed

- The payroll parser is isolated in `payroll_core.py`; the GUI is `oracle_payroll_extractor.py`; XLSX creation is in `payroll_excel.py`.
- Employee identities are distinct from financial entries. The **institutional code** (e.g., `30200105`) is not a line-item account. Keep the full disbursement identifier and the national ID as text.
- Financial rows are accepted only in a recognized entitlement/deduction section, or for an explicitly labeled government contribution. Accounts, labels and amounts are retained with source page references.
- Employer contributions are **Informational** and must not be included in employee cash earnings.
- Official Oracle gross, deductions and net are read independently and checked against extracted items and two explicit possible employer-share equations. Any ambiguity, duplicate item or missing printed total creates a **ReviewRequired** record.
- Monetary calculations use `Decimal` and missing values remain empty; they are never coerced to invented zero.
- The Excel dialog lets you select one, several or all employees and one or several line items, gross/deductions/net totals, or any mixture. The resulting **native XLSX** has summary, item-detail, and review sheets in RTL, numeric amount cells and text identifiers.
- Employee names and PDF-derived text are escaped when written to XLSX to avoid spreadsheet formula injection.

## Developer setup

Python 3.12:

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -p "test_*.py" -v
python oracle_payroll_extractor.py
```

The CI suite runs tests on Windows and Linux and builds a Windows PyInstaller executable. The current EXE is downloaded from the latest passing GitHub Actions `OraclePayrollExtractor-Windows` artifact. Build files and executables are **not tracked in Git**.

## Operational accuracy caveat

The included tests cover synthetic text, metadata boundaries, category separation, employer contributions, duplicate item warnings, official net checks, selected employee exports and readable XLSX values. This is *not proof* that pdfplumber text extraction exactly reproduces the user's original Oracle PDFs. On real local files, compare employee counts, identity fields, all ledger codes, printed totals and reconciliation warnings before relying on exports. Visual-order Arabic or image-only scans without a trustworthy text layer may require a separate extraction adapter and explicit manual review.

**Do not upload real employee salary slips, national IDs, unredacted debug output or generated workbooks to the public repository.**

## Suggested smoke test on private real data

1. Keep a private copy of the original PDF on the workstation.
2. Select the PDF in the GUI and inspect employee and item count.
3. Export all employees, earnings, deductions, official net and selected item columns.
4. Check a few payroll lines against the source PDF (including multi-page records).
5. Verify printed gross/deductions/net and confirm **no employee has ReviewRequired** before treating output as validated. If warnings exist, investigate and retain them; never manually zero-fill missing amounts.

There is no automatic payment authorization in this utility.
