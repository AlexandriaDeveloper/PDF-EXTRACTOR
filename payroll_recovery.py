"""Recover table-based payslips when flattened PDF text loses row boundaries.

A conservative text parser is primary. The older table route is activated only
if semantic extraction found zero ledger lines while PDF tables yielded
identifiable line candidates. Any recovered data is marked ReviewRequired.
"""
from __future__ import annotations
from decimal import Decimal
from payroll_core import UNKNOWN, identifier, key, parse_text_pages

def recover_records(pages: list[dict]) -> tuple[list[dict],dict]:
    strict=parse_text_pages([p.get("raw_text","") for p in pages])
    semantic_items=sum(len(r["entries"]) for r in strict)
    table_items=sum(len(p.get("table_items",[])) for p in pages)
    diagnostics={
        "pages":len(pages),"semantic_items":semantic_items,"table_candidates":table_items,
        "text_pages":sum(bool(p.get("text_chars",0)) for p in pages),
        "recovery_used":False
    }
    if not table_items:
        return strict,diagnostics
    recovered=[]
    current=None

    def has(value):
        return value is not None and value!=UNKNOWN and str(value).strip()!=""

    def different(a,b):
        if a is None:return bool(b)
        for field in ("national_id","disbursement_number"):
            if has(a.get(field)) and has(b.get(field)) and a[field]!=b[field]:
                return True
        for field in ("name","code"):
            if has(a.get(field)) and has(b.get(field)) and key(a[field])!=key(b[field]):
                return True
        return False

    for page_index,page in enumerate(pages,1):
        info=identifier(page.get("raw_text",""))
        old_name=page.get("legacy_name",UNKNOWN)
        old_code=page.get("legacy_code",UNKNOWN)
        if not has(info.get("name")) and has(old_name):
            info["name"]=old_name
        if not has(info.get("code")) and has(old_code):
            info["code"]=old_code
        identity_hint=any(has(info.get(f)) for f in (
            "national_id","disbursement_number","name","code"))
        if (identity_hint or page.get("table_items")) and (
            current is None or (identity_hint and different(current,info))):
            current={**info,"items":{},"entries":[],"warnings":[
                "تم استخدام قراءة جدولية احتياطية؛ البيانات بحاجة للمطابقة مع أصل Oracle"
            ],"pages":[],"printed_gross":None,"printed_deductions":None,
                     "printed_net":None,"status":"ReviewRequired"}
            recovered.append(current)
        if current is None:continue
        if page_index not in current["pages"]:current["pages"].append(page_index)
        for field in ("name","code","national_id","disbursement_number","institutional_code"):
            if not has(current.get(field)) and has(info.get(field)):
                current[field]=info[field]
        for item in page.get("table_items",[]):
            k=(item["kind"],item["account_code"],item["label"])
            if k in current["items"]:
                current["warnings"].append("احتمال تكرار بند مستخرج من الجدول؛ راجع صفحة "+str(page_index))
            current["items"][k]=current["items"].get(k,Decimal(0))+item["amount"]
            current["entries"].append({
                "kind":item["kind"],"account_code":item["account_code"],
                "label":item["label"],"amount":item["amount"],
                "page":page_index,"raw_text":item["raw_text"]
            })
            if not item["account_code"]:
                current["warnings"].append("كود الحساب غير واضح في صفحة "+str(page_index))
    # Reuse totals that the strict text parser proved independently, matched
    # only by an explicit common identity or verified page overlap.
    for employee in recovered:
        for parsed in strict:
            same_identifier=(
                employee.get("national_id") is not None and
                employee.get("national_id")==parsed.get("national_id")
            ) or (
                employee.get("disbursement_number") is not None and
                employee.get("disbursement_number")==parsed.get("disbursement_number")
            )
            if not same_identifier:continue
            for field in ("printed_gross","printed_deductions","printed_net"):
                if parsed.get(field) is not None:
                    employee[field]=parsed[field]
            break
        if not employee["entries"]:
            employee["warnings"].append("لا توجد بنود مالية مؤكدة")
        if not has(employee.get("national_id")):
            employee["warnings"].append("الرقم القومي غير مثبت")
        if not has(employee.get("disbursement_number")):
            employee["warnings"].append("رقم الصرف الكامل غير مثبت")
        # Do not declare recovered items Verified or infer missing as zero.
        employee["status"]="ReviewRequired"
    if not semantic_items:
        diagnostics["recovery_used"]=bool(recovered)
        return recovered,diagnostics

    # Some employees can be parsed semantically while others have visually
    # interleaved accounts. Apply table recovery *only* to employees for whom
    # the primary parser found no entries, and do not duplicate the good ones.
    def identity_match(a,b):
        for field in ("national_id","disbursement_number"):
            if has(a.get(field)) and has(b.get(field)) and a[field]==b[field]:
                return True
        return False

    result=[]
    consumed=set()
    for original in strict:
        matches=[(i,fallback) for i,fallback in enumerate(recovered)
                 if i not in consumed and (
                     identity_match(original,fallback) or
                     (not has(original.get("national_id")) and
                      not has(original.get("disbursement_number")) and
                      set(original.get("pages",[])) & set(fallback.get("pages",[])))
                 )]
        if original.get("entries") or not matches:
            result.append(original)
            # Block duplicate recovered rows for an already parsed employee.
            if original.get("entries"):
                consumed.update(i for i,_ in matches)
            continue
        i,candidate=matches[0]
        consumed.add(i)
        if not candidate.get("entries"):
            result.append(original)
            continue
        for field in ("printed_gross","printed_deductions","printed_net"):
            if original.get(field) is not None:candidate[field]=original[field]
        diagnostics["recovery_used"]=True
        result.append(candidate)
    for i,candidate in enumerate(recovered):
        if i in consumed or not candidate.get("entries"):
            continue
        if any(identity_match(existing,candidate) for existing in result):
            continue
        diagnostics["recovery_used"]=True
        result.append(candidate)
    result.sort(key=lambda e:min(e.get("pages",[10**9])))
    return result,diagnostics
