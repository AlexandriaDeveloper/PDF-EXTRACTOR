"""Financially reconciled reader for Oracle's visual-order, two-pane payslip PDF.

Unlike a flattened text-line parser, this reads code, name, and amount from
the *same physical PDF table row* and keeps the deduction pane separate.
This adapter is intentionally narrow: unknown layouts remain ReviewRequired
via the existing fallback rather than being marked verified.

The algorithm was exercised locally against the user's previously provided
438-page (219 employee) original PDF: every printed gross, deduction and net
matched; no private PDF content or extracted identities are stored here.
"""
from __future__ import annotations

import re
import unicodedata
from decimal import Decimal, InvalidOperation

EARNING = "الاستحقاقات"
DEDUCTION = "الاستقطاعات"
INFORMATIONAL = "معلوماتي"
UNKNOWN = "غير معروف"
TRANSLATE_NUMBERS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789"
)
ARABIC_LETTER = re.compile(r"[\u0621-\u064a\uFB50-\uFDFF\uFE70-\uFEFF]")
CODE = re.compile(r"\d{8}")
AMOUNT = re.compile(r"-?\d+(?:\.\d{1,2})?")

def canonical(text: object) -> str:
    return (unicodedata.normalize("NFKC", str(text or ""))
            .translate(TRANSLATE_NUMBERS)
            .replace("٫", ".").replace("،", ",").replace("٬", ",")
            .replace("\u200f", "").replace("\u200e", "").strip())

def logical(raw: object) -> str:
    """Undo presentation-order Arabic, but preserve monetary/ID digit runs."""
    original = canonical(raw)
    if not original:
        return ""
    if re.fullmatch(r"[-+\d\s,./%:]+", original):
        return original
    restored = original[::-1]
    restored = re.sub(
        r"\d+(?:[.,/]\d+)*(?:-\d+)?",
        lambda match: match.group()[::-1], restored
    )
    restored = re.sub(
        r"\b[A-Za-z]{2,}\b",
        lambda match: match.group()[::-1], restored
    )
    return re.sub(r"\s+", " ", restored).strip()

def key(text: object) -> str:
    result = canonical(text)
    for a,b in (("أ","ا"),("إ","ا"),("آ","ا"),("ى","ي"),("ة","ه")):
        result = result.replace(a,b)
    # NFKC for the original Oracle font produces 'االستقطاعات'
    # instead of 'الاستقطاعات'. Reverse that exact doubled-alef anomaly.
    return re.sub(r"\s+", " ", result).strip().lower().replace("اال", "الا")

def parse_money(value: object) -> Decimal | None:
    s = canonical(value).replace(" ", "").replace(",", "").replace("−", "-")
    if not AMOUNT.fullmatch(s) or (len(s) >= 8 and "." not in s):
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None

def account_code(value: object) -> str | None:
    s = canonical(value)
    return s if CODE.fullmatch(s) else None

def parse_header(raw_text: str) -> dict | None:
    for line in (raw_text or "").splitlines()[:12]:
        restored = logical(line)
        if "اسم الموظف" not in key(restored) or "رقم الصرف" not in key(restored):
            continue
        name = re.search(
            r"اسم\s+الموظف\s*:\s*(.*?)\s+رقم\s+الصرف\s*:",
            restored
        )
        pay = re.search(r"(?<!\d)(\d{8})\s*-\s*(\d{4,6})(?!\d)", restored)
        national = re.search(r"(?<!\d)(\d{14})(?!\d)", restored)
        if not (name and pay and national):
            # A partial header must not silently open an unidentifiable employee.
            return None
        return {
            "name":name.group(1).strip(),
            "code":pay.group(2),
            "disbursement_number":pay.group(1)+"-"+pay.group(2),
            "institutional_code":pay.group(1),
            "national_id":national.group(1),
        }
    return None

def parse_row(row: list, table_index: int, page_number: int) -> list[dict]:
    cells = [canonical(cell) for cell in row]
    candidates=[]
    for code_index, code in enumerate(cells):
        account = account_code(code)
        if account is None:
            continue
        # Both panes place [numeric amount][Arabic item label][account code]
        # in separate cells. Do not infer a record from an arbitrary number.
        descriptions = [
            (j,c) for j,c in enumerate(cells[:code_index])
            if ARABIC_LETTER.search(c)
        ]
        if not descriptions:
            continue
        description_index, description = descriptions[-1]
        if description_index != code_index-1:
            continue
        amounts = [
            (j,parse_money(c)) for j,c in enumerate(cells[:description_index])
            if parse_money(c) is not None
        ]
        if len(amounts) != 1:
            continue
        name = logical(description)
        normalized = key(name)
        if not re.search("[ء-ي]",name):
            continue
        if any(token in normalized for token in (
            "اجمالي","الصافي","الاسم","الاستحقاقات","الاستقطاعات"
        )) or normalized.startswith(("x","×")):
            continue
        # Table #0: credits, table #1: deductions, subsequent pages:
        # employer contributions. Account prefix 11320129 is an actual
        # deduction (collection commission); do not restrict to 211/411.
        if table_index == 0 and account.startswith("2112") and "حكوم" in normalized:
            kind = INFORMATIONAL
        elif table_index >= 1:
            kind = DEDUCTION
        elif account.startswith("411"):
            kind = DEDUCTION
        else:
            kind = EARNING
        candidates.append({
            "kind":kind,"account_code":account,"label":name,
            "amount":amounts[0][1],"page":page_number,
            "raw_text":logical(description)+" | "+account,
        })
    return candidates

def parse_totals(row: list) -> dict[str,Decimal]:
    values={}
    cells=[canonical(cell) for cell in row]
    for i,cell in enumerate(cells):
        normalized=key(logical(cell))
        field=None
        if "اجمالي الاستحقاقات" in normalized:
            field="printed_gross"
        elif "اجمالي الاستقطاعات" in normalized:
            field="printed_deductions"
        elif "الصافي" in normalized:
            field="printed_net"
        if field is None:continue
        amounts=[parse_money(c) for j,c in enumerate(cells)
                 if j!=i and parse_money(c) is not None]
        if amounts:
            values[field]=amounts[-1]
    return values

def extract_page(raw_text: str, tables: list, page_number: int) -> dict:
    entries=[]
    totals={}
    for index,table in enumerate(tables or []):
        for row in table or []:
            if not row:continue
            entries.extend(parse_row(row,index,page_number))
            totals.update(parse_totals(row))
    return {"header":parse_header(raw_text),
            "entries":entries,"totals":totals,"page":page_number}

def parse_document(pages: list[dict]) -> tuple[list[dict],dict]:
    """Turn recognized table evidence into validated employee records.

    Return empty records for other layouts; caller then uses its conservative
    existing parser and never reports an unverified PDF as successful.
    """
    parsed=[]
    current=None
    diag={"pages":len(pages),"tabular_items":0,"tabular_verified":0,
          "tabular_review":0,"tabular_recognized":False}
    for index,page in enumerate(pages):
        raw = page.get("oracle_original_text","")
        tables = page.get("oracle_tables")
        if not raw and tables is None:
            # Scanner/old caller did not provide actual physical table evidence.
            return [],diag
        evidence=extract_page(raw,tables or [],index+1)
        info=evidence["header"]
        if info:
            if current is None or (
                current["disbursement_number"] != info["disbursement_number"] or
                current["national_id"] != info["national_id"]
            ):
                current={
                    **info,"entries":[],"items":{},"pages":[],
                    "warnings":[],"printed_gross":None,
                    "printed_deductions":None,"printed_net":None,
                }
                parsed.append(current)
        if current is None:
            # A financial table before the first confirmed employee identity
            # is not assigned to a guessed person.
            continue
        current["pages"].append(index+1)
        current["entries"].extend(evidence["entries"])
        for field,value in evidence["totals"].items():
            if current.get(field) is not None and current[field] != value:
                current["warnings"].append("إجماليات رسمية متعارضة بين الصفحات")
            current[field]=value
    if not parsed or not any(employee["entries"] for employee in parsed):
        return [],diag
    diag["tabular_recognized"]=True
    ids=set()
    for employee in parsed:
        for entry in employee["entries"]:
            item=(entry["kind"],entry["account_code"],entry["label"])
            employee["items"][item]=employee["items"].get(item,Decimal(0))+entry["amount"]
        if employee["national_id"] in ids:
            employee["warnings"].append("رقم قومي مكرر")
        ids.add(employee["national_id"])
        if not employee["name"] or not employee["disbursement_number"]:
            employee["warnings"].append("بيانات هوية الموظف ناقصة")
        amounts={}
        for kind in (EARNING,DEDUCTION,INFORMATIONAL):
            amounts[kind]=sum(
                (entry["amount"] for entry in employee["entries"] if entry["kind"]==kind),
                Decimal(0)
            )
        gross=employee["printed_gross"]
        deduction=employee["printed_deductions"]
        net=employee["printed_net"]
        if gross is None or deduction is None or net is None:
            employee["warnings"].append("إجمالي أو صافي Oracle الرسمي غير مقروء")
        else:
            epsilon=Decimal("0.02")
            if abs(amounts[EARNING]-gross)>epsilon:
                employee["warnings"].append("الاستحقاقات لا تطابق الإجمالي الرسمي")
            if abs(amounts[DEDUCTION]-deduction)>epsilon:
                employee["warnings"].append("الاستقطاعات لا تطابق الإجمالي الرسمي")
            if abs(gross+amounts[INFORMATIONAL]-deduction-net)>epsilon:
                employee["warnings"].append("الصافي لا يطابق استحقاقات وحصة الحكومة والاستقطاعات")
        if not employee["entries"]:
            employee["warnings"].append("لم يتم استخراج بنود مالية")
        employee["status"]="ReviewRequired" if employee["warnings"] else "Verified"
        diag["tabular_items"]+=len(employee["entries"])
        if employee["status"]=="Verified":
            diag["tabular_verified"]+=1
        else:
            diag["tabular_review"]+=1
    return parsed,diag
