"""Semantic, fail-closed Arabic Oracle payslip extraction.

No GUI imports, Excel dependency, page counts, PDF coordinates or employee IDs
are hard-coded. Actual source-page text is never written to diagnostic logs.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from concurrent.futures import ProcessPoolExecutor, as_completed
import os
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata
from typing import Callable

import pdfplumber

_TRANSLATION = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_DIACRITICS = re.compile(r"[\u064b-\u065f\u0670\u0640\u200e\u200f\u061c]")
_NUM = re.compile(r"(?<![\d])[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?(?![\d])")
_DECIMAL_NUM = re.compile(r"(?<!\d)[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{1,2}(?!\d)")
_ACCOUNT = re.compile(r"(?<!\d)\d{8}(?!\d)")
_NATIONAL_ID = re.compile(r"(?<!\d)\d{14}(?!\d)")
_DISBURSEMENT = re.compile(r"(?<!\d)\d{7,9}\s*-\s*\d{4,6}(?!\d)")
_REVERSED_DISBURSEMENT = re.compile(r"(?<!\d)(\d{4,6})\s*-\s*(\d{8})(?!\d)")
_SECTION_EARNING = ("الاستحقاقات",)
_SECTION_DEDUCTION = ("الاستقطاعات",)
_HEADINGS = (
    "اسم الموظف", "الرقم القومي", "رقم الصرف", "نوع الصرف",
    "البنك", "الفرع", "الرقم التاميني", "الدرجه", "الاداره",
    "كود الجهه", "جهه العمل", "العنوان", "الوظيفه",
)
_KIND_AR = {"Earning": "الاستحقاقات", "Deduction": "الاستقطاعات",
            "Informational": "معلوماتي"}


def normalize(text: object) -> str:
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text)).translate(_TRANSLATION)
    s = _DIACRITICS.sub("", s)
    return s.replace("٫", ".").replace("٬", ",").replace("،", ",")


def key(text: object) -> str:
    s = normalize(text)
    for x, y in {"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه"}.items():
        s = s.replace(x, y)
    return re.sub(r"\s+", " ", s).strip().lower()


def number(value: object) -> Decimal | None:
    s = normalize(value).strip().replace(",", "").replace(" ", "")
    if not re.fullmatch(r"[-+]?\d+(?:\.\d{1,2})?", s):
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def _reverse_visual_line(line: str) -> str:
    """Fallback for visually reversed PDFs; preserve numeric order and Latin words."""
    rev = line[::-1]
    rev = re.sub(r"[\d]+(?:[.,][\d]+)?", lambda m: m.group()[::-1], rev)
    rev = re.sub(r"%(\d+(?:\.\d+)?)", r"\1%", rev)
    return re.sub(r"[A-Za-z]{2,}", lambda m: m.group()[::-1], rev)


def _orientation_score(text: str) -> int:
    v = key(text)
    return sum(3 for s in (
        "اسم الموظف", "رقم الصرف", "الرقم القومي", "الاستحقاقات",
        "الاستقطاعات", "اجمالي", "الصافي", "البنك", "حكومه",
    ) if s in v) + len(re.findall(r"(?<!\d)\d{8}(?!\d)", v))


def logical_lines(source: str) -> list[str]:
    """Keep explicit line boundaries; don't reverse already-logical Unicode."""
    result: list[str] = []
    for raw in normalize(source).splitlines():
        direct = re.sub(r"\s+", " ", raw).strip()
        if not direct:
            continue
        reversed_line = re.sub(r"\s+", " ", normalize(_reverse_visual_line(direct))).strip()
        result.append(reversed_line if _orientation_score(reversed_line) > _orientation_score(direct) else direct)
    return result


def _has(text: str, phrase: str) -> bool:
    return key(phrase) in key(text)


def _first_labeled(text: str, label: str, next_labels: tuple[str, ...] = ()) -> str | None:
    """Allow labels to be embedded on one line without consuming the next field."""
    t = key(text)
    match = re.search(re.escape(key(label)) + r"\s*:\s*", t)
    if not match:
        return None
    rest = t[match.end():]
    end = len(rest)
    for nxt in next_labels:
        m = re.search(re.escape(key(nxt)) + r"\s*:", rest)
        if m:
            end = min(end, m.start())
    value = rest[:end].strip(" :")
    return value if value else None


def _metadata_line(line: str) -> bool:
    k = key(line)
    if any(re.search(r"(?:^|\s)" + re.escape(field) + r"\s*:", k)
           for field in _HEADINGS):
        return True
    return ("نظام اوراكل" in k or "كشف رواتب" in k or "المرتب عن شهر" in k
            or "هذا البيان للمراجعه" in k or "br_" in k)


def _section_heading(line: str) -> str | None:
    k = key(line)
    if "اجمالي" in k or "الصافي" in k or "القيمه" in k:
        return None
    if k == "الاستحقاقات" or re.fullmatch(r"[-=: ]*الاستحقاقات[-=: ]*", k):
        return "Earning"
    if k == "الاستقطاعات" or re.fullmatch(r"[-=: ]*الاستقطاعات[-=: ]*", k):
        return "Deduction"
    return None


def _summary_kind(line: str) -> str | None:
    k = key(line)
    if "اجمالي الاستحقاقات" in k:
        return "gross"
    if "اجمالي الاستقطاعات" in k:
        return "deductions"
    if re.search(r"(?:^|\s)الصافي\s*:", k) or re.search(r"(?:^|\s)الصافي\s+المستحق", k):
        return "net"
    return None


def _amount_on_line(line: str) -> Decimal | None:
    # Prefer explicit two-decimal currency. Never interpret an ID or account as a total.
    matches = _DECIMAL_NUM.findall(normalize(line))
    if not matches:
        # A report may print an integer total: allow it only if it isn't an identifier.
        matches = [m.group() for m in _NUM.finditer(normalize(line))
                   if 1 <= len(re.sub(r"\D", "", m.group())) <= 7]
    return number(matches[-1]) if matches else None


@dataclass(frozen=True)
class PayrollEntry:
    kind: str
    account_code: str | None
    label: str
    amount: Decimal
    page: int
    raw_text: str
    source_section: str | None = None

    @property
    def choice_key(self) -> str:
        return f"{self.kind}|{self.account_code or ''}|{key(self.label)}"


@dataclass
class PayrollEmployee:
    name: str = ""
    national_id: str | None = None
    disbursement_number: str | None = None
    institutional_entity_code: str | None = None
    bank: str | None = None
    start_page: int = 0
    end_page: int = 0
    entries: list[PayrollEntry] = field(default_factory=list)
    gross: Decimal | None = None
    deductions: Decimal | None = None
    net: Decimal | None = None
    warnings: list[str] = field(default_factory=list)

    @property
    def verified(self) -> bool:
        return not self.warnings

    @property
    def stable_id(self) -> str:
        return self.national_id or self.disbursement_number or f"unresolved:{self.start_page}"


@dataclass
class ImportReport:
    filename: str
    employees: list[PayrollEmployee]
    pages: int
    warnings: list[str] = field(default_factory=list)

    @property
    def verified(self) -> bool:
        return not self.warnings and all(e.verified for e in self.employees)


def _employee_header(lines: list[str]) -> dict[str, str | None]:
    joined = " ".join(lines[:34])
    labeled_pay = _first_labeled(joined, "رقم الصرف",
        ("الرقم القومي", "الرقم القومى", "نوع الصرف", "البنك", "الفرع"))
    pay_match = _DISBURSEMENT.search(labeled_pay or "")
    if pay_match:
        pay = re.sub(r"\s+", "", pay_match.group())
    else:
        # pdfplumber may reverse the *order* of both numeric groups while
        # keeping each group internally readable. 54865-30200105 must be
        # canonicalized to 30200105-54865, never guessed from another field.
        reversed_pay = _REVERSED_DISBURSEMENT.search(labeled_pay or "")
        pay = (reversed_pay.group(2) + "-" + reversed_pay.group(1)) if reversed_pay else None
    labeled_id = _first_labeled(joined, "الرقم القومي",
        ("نوع الصرف", "البنك", "الفرع", "الرقم التاميني", "الدرجه", "الاداره"))
    nid_match = _NATIONAL_ID.search(labeled_id or "")
    # Support Arabic spelling "قومى" via key().
    name = _first_labeled(joined, "اسم الموظف",
        ("رقم الصرف", "الرقم القومي", "نوع الصرف", "البنك"))
    if name:
        name = re.sub(r"\s+", " ", name).strip()
    bank = _first_labeled(joined, "البنك",
        ("الفرع", "الرقم التاميني", "الدرجه", "الاداره"))
    return {
        "name": name, "national_id": nid_match.group() if nid_match else None,
        "disbursement_number": pay,
        "institutional_entity_code": pay.split("-")[0] if pay and len(pay.split("-")[0]) == 8 else None,
        "bank": bank,
    }


def _header_present(lines: list[str]) -> bool:
    return any("اسم الموظف" in key(l) for l in lines[:34])


def _entry_from_text(line: str, state: str, page: int) -> PayrollEntry | None:
    if _metadata_line(line) or _summary_kind(line):
        return None
    k = key(line)
    if ("اجمالي" in k or "الصافي" in k or k.startswith("x ")
        or k.startswith("× ") or k.startswith("x") and not _ACCOUNT.search(line)
        or "القيمه" in k and "الاسم" in k):
        return None

    account_matches = list(_ACCOUNT.finditer(line))
    if len(account_matches) != 1:
        return None
    code = account_matches[0].group()
    # Reject a known metadata code. More importantly, the section gate below
    # prevents identity fields from ever entering this parser.
    stripped = (line[:account_matches[0].start()] + " " +
                line[account_matches[0].end():]).strip()
    amounts = list(_DECIMAL_NUM.finditer(stripped))
    if not amounts:
        # Integer salary values are permitted, but never a percentage
        # and only if an anchored account code and label also exist.
        amounts = [m for m in _NUM.finditer(stripped)
                   if 1 <= len(re.sub(r"\D", "", m.group())) <= 7]
    if not amounts:
        return None
    valid = [m for m in amounts if not stripped[m.end():].lstrip().startswith("%")]
    if not valid:
        return None
    chosen = valid[-1]
    amount = number(chosen.group())
    if amount is None:
        return None
    label = (stripped[:chosen.start()] + " " + stripped[chosen.end():]).strip()
    label = re.sub(r"\s+", " ", label)
    if not re.search(r"[\u0621-\u064a]", label) or len(label) < 3:
        return None
    # Government contributions are informative even when Oracle prints them
    # near the deduction section: they are not employee cash earnings.
    kind = "Informational" if "حكومه" in key(label) else state
    return PayrollEntry(kind, code, label, amount, page, line, state)


def _parse_employee(pages: list[tuple[int, str]]) -> PayrollEmployee:
    source_lines = [logical_lines(page) for _, page in pages]
    header = next((_employee_header(lines) for lines in source_lines
                   if _header_present(lines)), {})
    employee = PayrollEmployee(
        name=header.get("name") or "",
        national_id=header.get("national_id"),
        disbursement_number=header.get("disbursement_number"),
        institutional_entity_code=header.get("institutional_entity_code"),
        bank=header.get("bank"), start_page=pages[0][0],
        end_page=pages[-1][0],
    )
    if not employee.name or re.search(r"\d", employee.name):
        employee.warnings.append("اسم الموظف غير واضح")
    if not employee.national_id:
        employee.warnings.append("الرقم القومي غير مثبت")
    if not employee.disbursement_number:
        employee.warnings.append("رقم الصرف الكامل غير مثبت")

    state: str | None = None
    started = False
    pending: str | None = None
    pending_page = 0
    official: dict[str, Decimal] = {}
    for (page_num, _), lines in zip(pages, source_lines):
        for line in lines:
            summary = _summary_kind(line)
            if summary:
                if pending:
                    item = _entry_from_text(pending, state or "Informational", pending_page)
                    if item:
                        employee.entries.append(item)
                    else:
                        employee.warnings.append(f"بند مالي غير مثبت بالصفحة {pending_page}")
                    pending = None
                value = _amount_on_line(line)
                if value is None:
                    employee.warnings.append(f"إجمالي رسمي غير مقروء بالصفحة {page_num}: {summary}")
                elif summary in official and official[summary] != value:
                    employee.warnings.append(f"إجمالي رسمي متعارض: {summary}")
                else:
                    official[summary] = value
                # Employer contribution detail can appear after gross/deductions.
                state = "Informational"
                continue

            heading = _section_heading(line)
            if heading:
                if pending:
                    item = _entry_from_text(pending, state or "Informational", pending_page)
                    if item:
                        employee.entries.append(item)
                    else:
                        employee.warnings.append(f"بند مالي غير مثبت بالصفحة {pending_page}")
                    pending = None
                started = True
                state = heading
                continue
            if not started:
                continue
            if _metadata_line(line):
                if pending:
                    item = _entry_from_text(pending, state or "Informational", pending_page)
                    if item:
                        employee.entries.append(item)
                    else:
                        employee.warnings.append(f"بند مالي غير مثبت بالصفحة {pending_page}")
                pending = None
                continue
            if re.match(r"^\s*(?:x|×)\s*",line,flags=re.IGNORECASE):
                pending = None
                continue

            # Candidate entries must be anchored by an account or an explicitly
            # identifiable employer contribution; do not mine metadata numbers.
            accounts = list(_ACCOUNT.finditer(line))
            if len(accounts) > 1:
                if pending:
                    item = _entry_from_text(pending, state or "Informational", pending_page)
                    if item:
                        employee.entries.append(item)
                    pending = None
                # Multiple code-first rows flattened onto one text line:
                # preserve each candidate and flag rows that can't be proved.
                fragments = re.split(r"(?=(?<!\d)\d{8}(?!\d))", line)
                for frag in fragments:
                    if not _ACCOUNT.search(frag):
                        continue
                    item = _entry_from_text(frag, state or "Informational", page_num)
                    if item and (state != "Informational" or item.kind == "Informational"):
                        employee.entries.append(item)
                    else:
                        employee.warnings.append(f"سطر متعدد البنود يحتاج مراجعة بالصفحة {page_num}")
                continue
            if accounts:
                if pending:
                    item = _entry_from_text(pending, state or "Informational", pending_page)
                    if item and (state != "Informational" or item.kind == "Informational"):
                        employee.entries.append(item)
                    elif not item:
                        employee.warnings.append(f"بند مالي غير مثبت بالصفحة {pending_page}")
                pending = line
                pending_page = page_num
                item = _entry_from_text(pending, state or "Informational", pending_page)
                if item:
                    if state != "Informational" or item.kind == "Informational":
                        employee.entries.append(item)
                    pending = None
                continue
            if pending:
                pending = pending + " " + line
                item = _entry_from_text(pending, state or "Informational", pending_page)
                if item:
                    if state != "Informational" or item.kind == "Informational":
                        employee.entries.append(item)
                    pending = None
            elif "حكومه" in key(line):
                employee.warnings.append(f"حصة جهة عمل بلا كود مثبت بالصفحة {page_num}")
    if pending:
        item = _entry_from_text(pending, state or "Informational", pending_page)
        if item and (state != "Informational" or item.kind == "Informational"):
            employee.entries.append(item)
        else:
            employee.warnings.append(f"بند مالي غير مثبت بالصفحة {pending_page}")

    employee.gross = official.get("gross")
    employee.deductions = official.get("deductions")
    employee.net = official.get("net")
    for name, value in (("الاستحقاقات", employee.gross),
                        ("الاستقطاعات", employee.deductions),
                        ("الصافي", employee.net)):
        if value is None:
            employee.warnings.append(f"الإجمالي الرسمي غير موجود: {name}")

    sums = {kind: sum((e.amount for e in employee.entries if e.kind == kind), Decimal("0"))
            for kind in ("Earning", "Deduction", "Informational")}
    informational_deductions = sum(
        (e.amount for e in employee.entries
         if e.kind == "Informational" and e.source_section == "Deduction"),
        Decimal("0"))
    tolerance = Decimal("0.02")
    if employee.gross is not None and abs(sums["Earning"] - employee.gross) > tolerance:
        employee.warnings.append("بنود الاستحقاقات لا تساوي إجمالي Oracle الرسمي")
    if employee.deductions is not None:
        acceptable = (sums["Deduction"], sums["Deduction"] + informational_deductions,
                      sums["Deduction"] + sums["Informational"])
        if min(abs(employee.deductions - value) for value in acceptable) > tolerance:
            employee.warnings.append("بنود الاستقطاعات/حصص الحكومة لا تساوي الإجمالي الرسمي")
    if not any(e.kind == "Earning" for e in employee.entries):
        employee.warnings.append("لا توجد بنود استحقاق مؤكدة")
    if not any(e.kind == "Deduction" for e in employee.entries):
        employee.warnings.append("لا توجد بنود استقطاع مؤكدة")
    if employee.gross is not None and employee.deductions is not None and employee.net is not None:
        # The real Oracle payslips can print employer contributions inside
        # official deductions while cash net remains earnings minus employee
        # cash deductions. Historical layouts may use another convention:
        # never force a single equation without a matching source control.
        candidate_cash = employee.gross - sums["Deduction"]
        candidate_legacy = employee.gross + sums["Informational"] - employee.deductions
        if min(abs(employee.net - candidate_cash),
               abs(employee.net - candidate_legacy)) > tolerance:
            employee.warnings.append("الصافي لا يتطابق مع المعادلات المالية المدعومة")
    return employee


def import_page_texts(filename: str, page_texts: list[str]) -> ImportReport:
    """Parse provided page text; suitable for deterministic, privacy-safe tests."""
    groups: list[list[tuple[int, str]]] = []
    warnings: list[str] = []
    for i, raw in enumerate(page_texts, start=1):
        lines = logical_lines(raw)
        if _header_present(lines):
            hdr = _employee_header(lines)
            # Repeated employee header on a continuation page should not split
            # the record; a different verified ID must split it.
            if not groups:
                groups.append([])
            elif groups[-1]:
                old_headers = [_employee_header(logical_lines(v)) for _, v in groups[-1]
                               if _header_present(logical_lines(v))]
                old = old_headers[0] if old_headers else {}
                same_nid = bool(hdr.get("national_id") and hdr.get("national_id") == old.get("national_id"))
                same_pay = bool(hdr.get("disbursement_number") and hdr.get("disbursement_number") == old.get("disbursement_number"))
                if not (same_nid or same_pay):
                    groups.append([])
        elif not groups:
            warnings.append(f"صفحة {i} بدون رأس موظف؛ لم تُنسب إلى موظف آخر")
            continue
        groups[-1].append((i, raw))

    employees = [_parse_employee(group) for group in groups if group]
    ids: set[str] = set()
    for e in employees:
        if e.stable_id in ids:
            e.warnings.append("هوية موظف مكررة في الملف")
            warnings.append("توجد هوية موظف مكررة")
        ids.add(e.stable_id)
    if not employees:
        warnings.append("لم يتم إثبات أي موظف في الملف")
    return ImportReport(filename, employees, len(page_texts), warnings)


def _candidate_quality(text: str) -> int:
    """Score semantic completeness, not horizontal PDF coordinates."""
    lines = logical_lines(text)
    score = 0
    for line in lines:
        if _header_present([line]):
            score += 5
        if _section_heading(line):
            score += 5
        if _summary_kind(line) and _amount_on_line(line) is not None:
            score += 4
        if _ACCOUNT.search(line) and _DECIMAL_NUM.search(line):
            score += 3
        if "رقم الصرف" in key(line) and _DISBURSEMENT.search(line):
            score += 3
    return score


def _page_stream(page: object) -> str:
    """Select between native text order and table order without double-counting.

    Parsing *both* representations as separate detail records would duplicate
    money. Instead select a single, semantically stronger representation.
    """
    text = page.extract_text() or ""
    # Avoid expensive table extraction when a text page already has plausible
    # Oracle sections or anchored financial lines.
    text_score = _candidate_quality(text)
    if text_score >= 14:
        return text
    try:
        tables = page.extract_tables() or []
    except (AttributeError, ValueError):
        tables = []
    table_lines: list[str] = []
    for table in tables:
        for row in table:
            if not row:
                continue
            cells = [str(c).replace("\n", " ").strip() for c in row if c is not None and str(c).strip()]
            if cells:
                table_lines.append(" ".join(cells))
    table_text = "\n".join(table_lines)
    return table_text if _candidate_quality(table_text) > text_score else text


def _pdf_chunk(path: str, start: int, end: int) -> list[tuple[int, str]]:
    """Independent PDF handle per process: never share pdfminer page objects."""
    from payroll_tables import extract_table_semantics

    pages: list[tuple[int, str]] = []
    with pdfplumber.open(path) as document:
        for index in range(start, end):
            page = document.pages[index]
            try:
                semantic = extract_table_semantics(page)
                pages.append((index, semantic if semantic is not None else _page_stream(page)))
            finally:
                page.flush_cache()
    return pages


def import_pdf(path: str | Path, progress: Callable[[int, int], None] | None = None) -> ImportReport:
    pdf_path = Path(path)
    with pdfplumber.open(str(pdf_path)) as document:
        total = len(document.pages)
    if total == 0:
        return ImportReport(pdf_path.name, [], 0, ["PDF فارغ"])

    chunk_size = 12
    chunks = [(start, min(total, start + chunk_size))
              for start in range(0, total, chunk_size)]
    workers = min(8, max(1, (os.cpu_count() or 2) - 1), len(chunks))
    results: dict[int, list[tuple[int, str]]] = {}
    completed = 0
    # Single-page files incur no process-launch overhead; large private Oracle
    # reports are parsed in parallel and restored to original page order.
    if workers == 1:
        for start, end in chunks:
            results[start] = _pdf_chunk(str(pdf_path), start, end)
            completed += end - start
            if progress:
                progress(completed, total)
    else:
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(_pdf_chunk, str(pdf_path), start, end): start
                       for start, end in chunks}
            for future in as_completed(futures):
                start = futures[future]
                records = future.result()
                results[start] = records
                completed += len(records)
                if progress:
                    progress(completed, total)
    ordered = [text for start in sorted(results) for _, text in results[start]]
    if len(ordered) != total:
        raise RuntimeError("PDF extraction returned an incomplete page set")
    return import_page_texts(pdf_path.name, ordered)
