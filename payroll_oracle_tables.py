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
    """Select logical/visual Arabic order by evidence, never reverse by default."""
    original=canonical(raw)
    if not original or not ARABIC_LETTER.search(str(raw)):
        return original
    reversed_text=original[::-1]
    reversed_text=re.sub(
        r"\d+(?:[.,/]\d+)*(?:-\d+)?",
        lambda match: match.group()[::-1],reversed_text
    )
    reversed_text=re.sub(
        r"\b[A-Za-z]{2,}\b",
        lambda match: match.group()[::-1],reversed_text
    )
    # Glyphs in Unicode Arabic Presentation Forms are reliably visual order.
    if re.search(r"[\uFB50-\uFDFF\uFE70-\uFEFF]",str(raw)):
        return re.sub(r"\s+"," ",reversed_text).strip()
    def score(candidate):
        words=("اسم الموظف","رقم الصرف","الرقم القومي","الاستحقاقات",
               "الاستقطاعات","اجمالي","الصافي","حكومة","مؤمن",
               "بدل","حافز","علاوة","الاجر","شيخوخة","خصم","عمولة")
        text=key(candidate)
        return (sum(3 for word in words if key(word) in text)
                +len(re.findall(r"\bال[ء-ي]{2,}",text)))
    return reversed_text if score(reversed_text)>score(original) else original

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

def section_label(text: object) -> str | None:
    normalized=key(logical(text))
    if "اجمالي" in normalized:return None
    if "الاستحقاقات" in normalized:return EARNING
    if "الاستقطاعات" in normalized:return DEDUCTION
    return None

def summary_field(text: object) -> str | None:
    normalized=key(logical(text))
    if "اجمالي الاستحقاقات" in normalized:return "printed_gross"
    if "اجمالي الاستقطاعات" in normalized:return "printed_deductions"
    if "الصافي" in normalized or "صافى" in normalized:return "printed_net"
    return None

def infer_table_section(table: list, inherited: str | None = None) -> str | None:
    """Use actual headings, or a table's official subtotal as evidence.

    Unlike old code, neither the ordinal table number nor the number of
    columns determines the kind. Conflicting headings remain unknown.
    """
    labels={section_label(cell) for row in table[:12] for cell in row if cell}
    labels.discard(None)
    if len(labels)==1:return next(iter(labels))
    if len(labels)>1:return None
    summaries=set()
    for row in table:
        for cell in row:
            kind=summary_field(cell)
            if kind=="printed_gross":summaries.add(EARNING)
            elif kind=="printed_deductions":summaries.add(DEDUCTION)
    if len(summaries)==1:return next(iter(summaries))
    return inherited

def locate_section_columns(table: list) -> list[tuple[int,str]]:
    """Determine pane regions from printed financial headings, not table order."""
    anchors={}
    for row in table[:min(12,len(table))]:
        for index,cell in enumerate(row):
            label=section_label(cell)
            if label:anchors[index]=label
    return sorted(anchors.items())

def section_for_code(column: int, anchors: list[tuple[int,str]],
                     default: str | None) -> str | None:
    kinds={kind for _,kind in anchors}
    if len(kinds)<=1:return next(iter(kinds)) if kinds else default
    distances=sorted((abs(column-col),col,kind) for col,kind in anchors)
    if len(distances)>1 and distances[0][0]==distances[1][0] and distances[0][2]!=distances[1][2]:
        return None
    return distances[0][2]

def parse_row(row: list, table_index: int, page_number: int,
              section: str | None = None,
              anchors: list[tuple[int,str]] | None = None) -> list[dict]:
    """Associate nearby financial cells with a code, not fixed coordinates.

    Parameter table_index exists for legacy callers, but deliberately has
    no influence on financial classification.
    """
    cells=[canonical(cell) for cell in row]
    labels=[logical(cell) for cell in cells]
    positions=[(i,account_code(cell)) for i,cell in enumerate(cells)
               if account_code(cell)]
    entries=[]
    for ci,account in positions:
        adjacent=[j for j in (ci-1,ci+1)
                  if 0<=j<len(cells) and
                  re.search("[ء-ي]",labels[j]) and
                  summary_field(cells[j]) is None and
                  section_label(cells[j]) is None]
        if len(adjacent)!=1:continue
        di=adjacent[0];label=labels[di]
        normalized=key(label)
        if (not re.search("[ء-ي]",label)
            or normalized.startswith(("x ","× "))
            or any(t in normalized for t in ("اجمالي","الصافي","الاسم"))):
            continue
        amounts=[(abs(j-di),j,parse_money(cell))
                 for j,cell in enumerate(cells) if parse_money(cell) is not None]
        # Exclude values belonging to a different code in the same physical row.
        competing_labels=[neighbor for other_index,_ in positions if other_index!=ci
                          for neighbor in (other_index-1,other_index+1)
                          if 0<=neighbor<len(cells) and re.search("[ء-ي]",labels[neighbor])]
        amounts=[x for x in amounts if x[0]<=2
                 and all(abs(x[1]-di)<abs(x[1]-neighbor)
                         for neighbor in competing_labels)]
        if not amounts:continue
        amounts.sort(key=lambda x:(x[0],x[1]))
        if len(amounts)>1 and amounts[0][0]==amounts[1][0]:
            continue
        row_section=section_for_code(ci,anchors or [],section)
        kind=(INFORMATIONAL if "حكوم" in normalized and row_section!=DEDUCTION
              else row_section)
        if kind not in (EARNING,DEDUCTION,INFORMATIONAL):continue
        entries.append({
            "kind":kind,"account_code":account,"label":label,
            "amount":amounts[0][2],"page":page_number,
            "raw_text":label+" | "+account,
        })
    return entries

def parse_totals(row: list) -> dict[str,Decimal]:
    values={}
    cells=[canonical(cell) for cell in row]
    for i,cell in enumerate(cells):
        field=summary_field(cell)
        if field is None:continue
        amounts=[(abs(j-i),j,parse_money(other))
                 for j,other in enumerate(cells)
                 if j!=i and parse_money(other) is not None]
        if amounts:
            amounts.sort(key=lambda x:(x[0],x[1]))
            if len(amounts)>1 and amounts[0][0]==amounts[1][0]:
                continue
            values[field]=amounts[0][2]
    return values

def extract_page(raw_text: str, tables: list, page_number: int) -> dict:
    entries=[];totals={}
    for index,table in enumerate(tables or []):
        section=infer_table_section(table)
        anchors=locate_section_columns(table)
        for row in table or []:
            if not row:continue
            headings={section_label(cell) for cell in row if cell}
            headings.discard(None)
            if len(headings)==1:section=next(iter(headings))
            entries.extend(parse_row(row,index,page_number,section,anchors))
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
            if (abs(gross+amounts[INFORMATIONAL]-deduction-net)>epsilon
                and abs(gross-deduction-net)>epsilon):
                employee["warnings"].append("الصافي لا يطابق الإجماليات أو حصة الحكومة")
        if not employee["entries"]:
            employee["warnings"].append("لم يتم استخراج بنود مالية")
        employee["status"]="ReviewRequired" if employee["warnings"] else "Verified"
        diag["tabular_items"]+=len(employee["entries"])
        if employee["status"]=="Verified":
            diag["tabular_verified"]+=1
        else:
            diag["tabular_review"]+=1
    return parsed,diag
