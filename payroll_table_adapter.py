"""Compatibility adapter for the *table-based* Oracle PDF layout.

The previous release successfully recognized tables whose Arabic glyphs are
stored visually reversed and whose 2 financial columns are interleaved. The new
strict text-line parser alone cannot read those cells. This adapter restores
the table path without accepting an unvalidated row as a verified salary entry.
"""
from __future__ import annotations

from decimal import Decimal
import re
import unicodedata
from payroll_core import (
    ACCOUNT, EARNING, DEDUCTION, INFORMATIONAL, UNKNOWN,
    amount, identifier, is_metadata, key, standard
)

WORDS = ("الاستحقاقات", "الاستقطاعات", "حكومة", "الحكومه", "الحكومة",
         "مؤمن", "تامين", "تأمين", "بدل", "علاوة", "علاوه",
         "الاجر", "الأجر", "حافز", "اسم", "موظف", "رقم",
         "القومي", "القومى", "الصرف", "الإجمالي", "اجمالي", "صافي")
SKIP = ("الاسم", "القيمة", "الاجمالي", "اجمالي", "الصافي", "العدد",
        "الاستحقاقات", "الاستقطاعات", "الكود", "الملاحظات", "كشف رواتب")
CELL_AMOUNT = re.compile(r"^-?\d+(?:,\d{3})*(?:\.\d{1,2})?$")
DIGITS = re.compile(r"(?<!\d)\d{8}(?!\d)")

def reverse_visual(raw: str) -> str:
    s = standard(raw)[::-1]
    s = re.sub(r"[\d]+(?:[.,][\d]+)?", lambda m:m.group()[::-1],s)
    s = re.sub(r"\b[A-Za-z]{2,}\b",lambda m:m.group()[::-1],s)
    s = re.sub(r"%(\d+(?:[.,]\d+)?)",r"\1%",s)
    return s

def _score(s: str) -> int:
    k=key(s)
    return sum(3 for word in WORDS if key(word) in k) + (
        2*len(re.findall(r"(?:^|\s)ال[ء-ي]{2,}",k))
    )

def cell_text(raw: object) -> str:
    s = standard(raw).strip()
    reversed_s = reverse_visual(s)
    if _score(reversed_s) > _score(s):
        return reversed_s
    # Arabic Presentation Forms in visually reversed Oracle cells are reliable
    # evidence to consider the reverse, never a reason to reverse all PDFs.
    if _score(reversed_s)==_score(s) and re.search(r"[\uFB50-\uFDFF\uFE70-\uFEFF]",str(raw)):
        return reversed_s
    return s

def _financial_value(value: object) -> Decimal | None:
    s=standard(value).strip().replace(" ","")
    if not CELL_AMOUNT.fullmatch(s):return None
    if len(s.replace(",","").lstrip("-")) in (8,14) and "." not in s:
        return None
    return amount(s)

def _kind_from_code(account: str | None) -> str | None:
    if account:
        if account.startswith("211"):return EARNING
        if account.startswith("411"):return DEDUCTION
    return None

def _is_description(s: str) -> bool:
    k=key(s)
    if not re.search("[ء-ي]",s) or len(s)<3:return False
    if is_metadata(s):return False
    if any(word in k for word in SKIP):return False
    if s.lstrip().lower().startswith(("x ","× ")):return False
    return True

def extract_table_candidates(page, tables=None) -> list[dict]:
    """Best-effort table rows; all results are explicitly *unverified*."""
    if tables is None:
        try: tables=page.extract_tables() or []
        except Exception:return []
    candidates=[]
    for table in tables:
        if not table:continue
        normalized=[[cell_text(c) if c is not None else "" for c in row] for row in table if row]
        joined=" ".join(" ".join(row) for row in normalized[:4])
        has_earn=key(EARNING) in key(joined)
        has_deduct=key(DEDUCTION) in key(joined)
        for row_index,row in enumerate(normalized):
            if len(row)<2:continue
            # Oracle payslips are typically a right/left two-pane table.
            # Prefer a 211/411 account prefix and use pane order only when
            # the account code is missing; never verify this inference.
            if has_earn and has_deduct and len(row)>=4:
                mid=len(row)//2
                groups=[(DEDUCTION,row[:mid]),(EARNING,row[mid:])]
            else:groups=[(EARNING if has_earn else DEDUCTION if has_deduct else None,row)]
            for hinted_kind,cells in groups:
                amounts=[(i,_financial_value(c)) for i,c in enumerate(cells)]
                amounts=[(i,v) for i,v in amounts if v is not None]
                if not amounts:continue
                names=[(i,c) for i,c in enumerate(cells) if _is_description(c)]
                if not names:continue
                codes=[m.group() for c in cells for m in DIGITS.finditer(c)]
                code=next((c for c in codes if c.startswith(("211","411"))),None)
                kind=_kind_from_code(code) or hinted_kind
                # A government contribution on either side is informational.
                desc=next((n for _,n in names if "حكوم" in key(n)),names[0][1])
                if "حكوم" in key(desc):kind=INFORMATIONAL
                if kind not in (EARNING,DEDUCTION,INFORMATIONAL):continue
                # Never adopt an unrelated amount if there are multiple plausible
                # values. Keep the first only for the adapter's REVIEW state.
                value=amounts[0][1]
                # If a descriptor embeds its account remove just the account.
                if code:
                    desc=re.sub(r"(?<!\d)"+re.escape(code)+r"(?!\d)","",desc).strip()
                if not _is_description(desc):continue
                candidates.append({
                    "kind":kind,"account_code":code or "",
                    "label":desc,"amount":value,
                    "raw_text":" | ".join(c for c in cells if c),
                    "source":"legacy-table","row":row_index+1
                })
    # Compatibility with the previously working parser: Oracle table extraction
    # sometimes places the account, description and amount into *different*
    # cells, with mixed left/right panes and columns missing code values.
    # A strict candidate can be empty in this situation; restore those rows
    # as explicit REVIEW-only entries rather than returning zero for the file.
    if not candidates:
        for table in tables:
            if not table:continue
            rows=[[cell_text(c) if c is not None else "" for c in row]
                  for row in table if row]
            whole=" ".join(" ".join(row) for row in rows)
            earnings=key(EARNING) in key(whole)
            deductions=key(DEDUCTION) in key(whole)
            for index,row in enumerate(rows):
                if not row or any(cell.count("\n")>3 for cell in row):continue
                numeric=[(i,_financial_value(c)) for i,c in enumerate(row)]
                numeric=[(i,v) for i,v in numeric if v is not None]
                if not numeric:continue
                for position,desc in enumerate(row):
                    if not _is_description(desc) or any(i==position for i,_ in numeric):
                        continue
                    if len(row)>1:
                        if earnings and not deductions:kind=EARNING
                        elif deductions and not earnings:kind=DEDUCTION
                        elif earnings and deductions:
                            kind=DEDUCTION if position < len(row)//2 else EARNING
                        else:
                            kind=DEDUCTION if len(row)<=4 or position<len(row)//2 else EARNING
                    else:continue
                    code=next((m.group() for cell in row
                               for m in DIGITS.finditer(cell)
                               if m.group().startswith(("211","411"))), "")
                    if code.startswith("211"):kind=EARNING
                    elif code.startswith("411"):kind=DEDUCTION
                    if "حكوم" in key(desc):kind=INFORMATIONAL
                    value=numeric[0][1]
                    candidates.append({
                        "kind":kind,"account_code":code,"label":desc,
                        "amount":value,"raw_text":" | ".join(c for c in row if c),
                        "source":"legacy-table","row":index+1
                    })
                    # Match the legacy rule: a single safely-reviewable
                    # description/value link per physical table row.
                    break
    # Deduplicate exact table representations (some pages return nested overlapping tables)
    seen=set()
    unique=[]
    for item in candidates:
        identity=(item["kind"],item["account_code"],item["label"],item["amount"],item["raw_text"])
        if identity in seen:continue
        seen.add(identity);unique.append(item)
    return unique
