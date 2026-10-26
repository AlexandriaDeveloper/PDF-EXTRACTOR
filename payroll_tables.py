"""Read Oracle's real table-cell structure without hardcoded PDF coordinates.

The tables are a *source representation*, not the payroll domain model. Each
row is accepted only when an account, description, and one amount are proven.
The result is a semantic text stream consumed by payroll_core.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re
import unicodedata

_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


def normalize(value: object) -> str:
    return (unicodedata.normalize("NFKC", str(value or "")).translate(_DIGITS)
            .replace("٫", ".").replace("٬", ",").replace("،", ","))


def reverse_visual(value: object) -> str:
    original = unicodedata.normalize("NFKC", str(value or "")).translate(_DIGITS)
    output: list[str] = []
    for line in original.splitlines():
        flipped = line[::-1]
        flipped = re.sub(r"\d+(?:[.,٫،]\d+)*",
                         lambda match: match.group()[::-1], flipped)
        output.append(flipped)
    return normalize(" ".join(output)).strip()


def compact(value: object) -> str:
    s = normalize(value)
    for a, b in {"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه"}.items():
        s = s.replace(a, b)
    return re.sub(r"\s+", "", s).replace("اال", "الا")


def account_cell(value: object) -> str | None:
    value = normalize(value).strip()
    return value if re.fullmatch(r"\d{8}", value) else None


def money_cell(value: object) -> Decimal | None:
    text = normalize(value).strip().replace(" ", "")
    if not re.fullmatch(r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?", text):
        return None
    try:
        return Decimal(text.replace(",", ""))
    except InvalidOperation:
        return None


def extract_table_semantics(page: object) -> str | None:
    """Select separately anchored salary items from each table.

    Long concatenated cells are *not* parsed as items (they duplicate the
    detailed rows elsewhere). Missing/ambiguous amounts remain review markers.
    """
    tables = page.extract_tables() or []
    if not tables:
        return None

    lines: list[str] = []
    detail_count = 0
    for table in tables:
        if not table:
            continue
        # Header is independently parsed as identity metadata by payroll_core.
        if table[0] and table[0][0]:
            candidate = reverse_visual(table[0][0])
            if "اسمالموظف" in compact(candidate):
                lines.append(candidate)

        section = ""
        for row in table[:3]:
            for cell in row:
                if not cell or len(str(cell)) > 60:
                    continue
                label = compact(reverse_visual(cell))
                if "الاستحقاقات" in label:
                    section = "Earning"
                if "الاستقطاعات" in label:
                    section = "Deduction"
        if section:
            lines.append("الاستحقاقات" if section == "Earning" else "الاستقطاعات")

        for row in table:
            if not row:
                continue
            accounts = [(i, account_cell(v)) for i, v in enumerate(row) if account_cell(v)]
            values = [(i, money_cell(v)) for i, v in enumerate(row) if money_cell(v) is not None]
            words = [(i, reverse_visual(v), compact(reverse_visual(v)))
                     for i, v in enumerate(row)
                     if v and re.search(r"[\u0621-\u064a]", normalize(v))]
            # Only concise summary cells paired with a single numeric total
            # may declare official figures. Ignore giant merged header cells.
            for _, label, sign in words:
                if len(label) > 90 or len(values) != 1:
                    continue
                if "جمالي" in sign and "الاستحقاقات" in sign:
                    lines.append("اجمالي الاستحقاقات " + str(values[0][1]))
                elif "جمالي" in sign and "الاستقطاعات" in sign:
                    lines.append("اجمالي الاستقطاعات " + str(values[0][1]))
                elif "الصافي" in sign:
                    lines.append("الصافي: " + str(values[0][1]))

            if len(accounts) != 1:
                continue
            account_position, account = accounts[0]
            labels = [(i, text, sign) for i, text, sign in words
                      if i != account_position and len(sign) > 3
                      and "جمالي" not in sign and "الصافي" not in sign]
            if not labels:
                continue
            _, description, signature = max(labels, key=lambda item: len(item[2]))
            if signature.startswith(("x", "×")):
                continue
            amounts = [v for i, v in values if i != account_position]
            if len(amounts) != 1:
                # The semantic importer marks unproven account details for
                # review; no financial amount is ever invented.
                lines.append(account + " " + description + " [مبلغ غير مؤكد]")
                continue
            if not section and "حكوم" not in signature:
                lines.append(account + " " + description + " [قسم غير مؤكد]")
                continue
            lines.append(account + " " + description + " " + str(amounts[0]))
            detail_count += 1
    return "\n".join(lines) if detail_count else None
