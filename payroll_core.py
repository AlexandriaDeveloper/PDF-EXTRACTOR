"""Conservative Oracle payslip parsing, identity boundaries and financial checks.

Records retain evidence and ambiguity instead of converting unknown values to zero.
PDF text extraction differs between files; no amount is trusted only because a
neighbouring cell happens to be numeric.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
import re
import unicodedata

EARNING = "الاستحقاقات"
DEDUCTION = "الاستقطاعات"
INFORMATIONAL = "معلوماتي"
UNKNOWN = "غير معروف"
SECTION = {
    "الاستحقاقات": EARNING, "الاستقطاعات": DEDUCTION
}
ACCOUNT = re.compile(r"(?<!\d)(\d{8})(?!\d)")
NATIONAL_ID = re.compile(r"(?<!\d)(\d{14})(?!\d)")
PAYMENT = re.compile(r"(?<!\d)(\d{8})\s*-\s*(\d{4,6})(?!\d)")
MONEY = re.compile(r"(?<![\d-])-?\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?(?!\d)|(?<![\d-])-?\d+(?:\.\d{1,2})?(?!\d)")
DECIMAL_AMOUNT = re.compile(r"^-?\d+(?:\.\d{1,2})?$")
LABEL = re.compile(r"^(?:اسم\s*الموظف|رقم\s*الصرف|الرقم\s*القومي|الرقم\s*القومى|البنك|الفرع|نوع\s*الصرف|الرقم\s*التاميني|الرقم\s*التأميني|الدرجه|الدرجة|الاداره|الإدارة)\s*:")
HEADER_WORDS = ("اسم الموظف", "الرقم القومي", "الرقم القومى", "نوع الصرف", "الرقم التاميني", "البنك:", "الفرع:", "الدرجة:", "الادارة:")
SUMMARY = ("اجمالي", "الاجمالي", "الصافي", "الصافى")
SUBTOTAL = re.compile(r"^\s*[x×]\s*", re.IGNORECASE)

def standard(text: object) -> str:
    s = unicodedata.normalize("NFKC", str(text or ""))
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789"))
    return s.replace("٫", ".").replace("٬", ",").replace("،", ",").replace("\u200f", "").replace("\u200e", "")

def key(text: object) -> str:
    s = standard(text)
    for a,b in (("أ","ا"),("إ","ا"),("آ","ا"),("ة","ه"),("ى","ي")):
        s=s.replace(a,b)
    return re.sub(r"\s+", " ", s).strip().lower()

def amount(text: object) -> Decimal | None:
    s=standard(text).replace(",","").strip()
    if not DECIMAL_AMOUNT.fullmatch(s):
        return None
    if "." not in s and len(s.lstrip("-")) in (8, 14):
        # Identifier-shaped integers must not be treated as cash amounts.
        return None
    try: return Decimal(s)
    except InvalidOperation: return None

def monetary_tokens(s: str) -> list[tuple[str,int,int]]:
    s=standard(s)
    return [(m.group(),m.start(),m.end()) for m in MONEY.finditer(s)
            if not (len(m.group()) == 8 and m.group().isdigit())
            and not (len(m.group()) == 14 and m.group().isdigit())]

def maybe_section(line: str) -> str | None:
    k=key(line)
    if k in (key(EARNING), key(DEDUCTION)):
        return EARNING if k==key(EARNING) else DEDUCTION
    return None

def is_metadata(line: str) -> bool:
    k=key(line)
    return bool(LABEL.match(k)) or any(key(token) in k for token in HEADER_WORDS)

def identifier(text: str) -> dict:
    s=standard(text)
    pay=PAYMENT.search(s)
    national=NATIONAL_ID.search(s)
    name=None
    m=re.search(r"اسم\s*الموظف\s*:\s*(.*?)(?=رقم\s*الصرف\s*:|الرقم\s*القوم[يى]\s*:|$)",s,re.S)
    if m:
        value=re.sub(r"\s+"," ",m.group(1)).strip()
        if value and not re.search(r"\d",value) and 3<len(value)<130:
            name=value
    return {
        "name":name or UNKNOWN,
        "code":pay.group(2) if pay else UNKNOWN,  # legacy display only
        "disbursement_number": (pay.group(1)+"-"+pay.group(2)) if pay else None,
        "institutional_code":pay.group(1) if pay else None,
        "national_id":national.group() if national else None,
    }

def normalize_line_item(raw: str, section: str | None) -> tuple[tuple[str,str,str],Decimal] | None:
    """Accept one item only with an account, financial context and credible amount."""
    s=standard(raw).strip()
    if not s or is_metadata(s) or SUBTOTAL.match(s):
        return None
    if any(word in key(s) for word in SUMMARY):
        return None
    code=ACCOUNT.search(s)
    if not code:
        return None
    remainder=(s[:code.start()]+" "+s[code.end():]).strip()
    # A national ID and an institutional identifier are not monetary values.
    if NATIONAL_ID.search(remainder) or PAYMENT.search(s):
        return None
    tokens=monetary_tokens(remainder)
    if not tokens: return None
    # If a token has a decimal separator use the last such token, otherwise
    # conservatively reject ambiguous multi-number rows, excluding percent tags.
    plausible=[]
    for token, start, end in tokens:
        if remainder[end:].lstrip().startswith("%"): continue
        if len(token)==4 and token.isdigit() and token.startswith(("19","20")): continue
        plausible.append((token,start,end))
    if not plausible:return None
    decimals=[t for t in plausible if "." in t[0] or "," in t[0]]
    selected=decimals[-1] if decimals else plausible[-1]
    val=amount(selected[0])
    if val is None:return None
    description=(remainder[:selected[1]]+" "+remainder[selected[2]:]).strip()
    description=re.sub(r"\s+"," ",description)
    if not re.search("[ء-ي]",description) or len(description)<3:return None
    if is_metadata(description):return None
    if section is None and "حكوم" not in key(description):return None
    kind=INFORMATIONAL if "حكوم" in key(description) else section
    if kind is None:return None
    return (kind,code.group(1),description),val

def parse_text_pages(pages: list[str]) -> list[dict]:
    """One record per explicit employee opening; following pages continue it."""
    employees=[]
    current=None
    section=None
    def start(page_index:int, info:dict):
        nonlocal current,section
        current={**info,"items":{},"entries":[],"warnings":[],"pages":[page_index+1],
                 "printed_gross":None,"printed_deductions":None,"printed_net":None}
        employees.append(current)
        section=None
    for page_index,page in enumerate(pages):
        lines=standard(page).splitlines()
        joined="\n".join(lines)
        info=identifier(joined)
        # A new employee is established by an explicit name label, not just a
        # national ID or a recurring account number on a continuation page.
        new_name=info["name"] != UNKNOWN and "اسم الموظف" in key(joined)
        if new_name:
            if current is None or (info["national_id"] and info["national_id"]!=current.get("national_id")) or (
                info["disbursement_number"] and info["disbursement_number"]!=current.get("disbursement_number")):
                start(page_index,info)
        if current is None:
            continue
        if page_index+1 not in current["pages"]:current["pages"].append(page_index+1)
        for line in lines:
            line=line.strip()
            k=key(line)
            if not line:continue
            if is_metadata(line):continue
            for label,field in (("اجمالي الاستحقاقات","printed_gross"),
                                ("اجمالي الاستقطاعات","printed_deductions"),
                                ("الصافي:","printed_net"),("الصافى:","printed_net")):
                if key(label) in k:
                    tokens=monetary_tokens(line)
                    if tokens:
                        value=amount(tokens[-1][0])
                        if value is not None:current[field]=value
                    section=None
                    break
            else:
                heading=maybe_section(line)
                if heading:
                    section=heading
                    continue
                if not section and "حكوم" not in k:continue
                parsed=normalize_line_item(line,section)
                if parsed is None:continue
                item,value=parsed
                if item in current["items"]:
                    # Repeated lines may denote either legitimate repeated payments
                    # or duplicated PDF text. Flag instead of overwrite/double count.
                    current["warnings"].append("بند مكرر يحتاج مراجعة: "+item[1]+" "+item[2])
                current["entries"].append({"kind":item[0],"account_code":item[1],
                                           "label":item[2],"amount":value,"page":page_index+1,"raw_text":line})
                current["items"][item]=current["items"].get(item,Decimal(0))+value
    identities={}
    for employee in employees:
        if employee.get("national_id") is None:employee["warnings"].append("الرقم القومي غير مثبت")
        if employee.get("disbursement_number") is None:employee["warnings"].append("رقم الصرف غير مثبت")
        if employee.get("name")==UNKNOWN:employee["warnings"].append("اسم الموظف غير مثبت")
        identity=employee.get("national_id") or employee.get("disbursement_number")
        if identity:
            if identity in identities:
                employee["warnings"].append("تكرار معرف الموظف في السجلات")
                identities[identity]["warnings"].append("تكرار معرف الموظف في السجلات")
            identities[identity]=employee
        earning=sum((v for (k,_,_),v in employee["items"].items() if k==EARNING),Decimal(0))
        deduction=sum((v for (k,_,_),v in employee["items"].items() if k==DEDUCTION),Decimal(0))
        if employee["printed_gross"] is None:employee["warnings"].append("إجمالي الاستحقاقات الرسمي غير متاح")
        elif earning != employee["printed_gross"]:employee["warnings"].append("اختلاف إجمالي الاستحقاقات عن البنود")
        if employee["printed_deductions"] is None:employee["warnings"].append("إجمالي الاستقطاعات الرسمي غير متاح")
        elif deduction != employee["printed_deductions"]:employee["warnings"].append("اختلاف إجمالي الاستقطاعات عن البنود")
        if employee["printed_net"] is None:
            employee["warnings"].append("صافي Oracle الرسمي غير متاح")
        elif employee["printed_gross"] is not None and employee["printed_deductions"] is not None:
            gross=employee["printed_gross"];ded=employee["printed_deductions"]
            net=employee["printed_net"]
            employer=sum((v for (k,_,_),v in employee["items"].items()
                          if k==INFORMATIONAL),Decimal(0))
            # Oracle layouts differ in whether employer shares enter printed
            # deduction totals. Both known equations are explicit checks;
            # unexplained differences must remain reviewable.
            if abs(gross-ded-net)>Decimal("0.02") and abs(gross+employer-ded-net)>Decimal("0.02"):
                employee["warnings"].append("الصافي الرسمي لا يتطابق مع الإجماليات وحصص جهة العمل")
        employee["status"]="ReviewRequired" if employee["warnings"] else "Verified"
    return employees

def sum_available(items: dict, kind: str) -> Decimal:
    return sum((v for (k,_,_),v in items.items() if k==kind),Decimal(0))

def export_amount(employee: dict, item_key: tuple[str,str,str], verified: bool) -> Decimal | None:
    if item_key in employee["items"]:return employee["items"][item_key]
    return Decimal(0) if verified and employee.get("status")=="Verified" else None
