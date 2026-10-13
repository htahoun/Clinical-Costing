"""
Split a month's Cost of Goods Sold into DOH cost types from the clinic's stock
movement, and give the implant stock pools.

Sources
  * One-month report (sheet "Movement <Mon> YYYY"), e.g. input/source/Financial_Report_Jan_2025_1.xlsx
  * The full-year workbook input/source/Inventory_movement_2025.xlsx (twelve sheets
    "Movement Jan 2025" … "Movement Dec 2025"), three systems (Finance Director, 13 Oct 2026):
      Jan–Apr  value = "Transaction Value"
      May–Nov  value = "Supplier Price"      (= quantity x supplier net price)
      Dec      value = "Buying Price"        (dispatch system: "Received" rows are usage,
                                              "Returned" rows are returns, no transaction type)

Consumption = "02- Usage" (negative) net of "03- Internal Return" (positive); "04-/05- Exchange"
cancel. External laboratory invoices are not stock (paid from cash in hand) and are added from the
Consumption sheet. Rulings 13 Oct 2026: April is a copy of March and is accepted as the ledger has it;
in August the external-lab usage line has no supplier price and is valued at its transaction value,
and the remaining difference to the P&L (431.82) goes to MS; external-lab lines inside the movement are
separate purchases from the cash-in-hand line, both Prostheses.

Category names and sub-category numbers differ between the systems, so lines are matched by the
category number and by keywords, not by the full name.

  python3 scripts/cogs_from_movement.py input/source/Inventory_movement_2025.xlsx 8 5254
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl

D2 = Decimal("0.01")
MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
BASIS = {m: "Transaction Value" for m in range(1, 5)} | {m: "Supplier Price" for m in range(5, 12)} | {12: "Buying Price"}
CONSUMING = {"02- Usage", "03- Internal Return"}


def dec(v) -> Decimal:
    if v in (None, "", "--", "NA") or not isinstance(v, (int, float, Decimal, str)):
        return Decimal(0)
    try:
        return Decimal(str(v).replace(",", ""))          # exact; rounded once per cost type at the end
    except Exception:  # noqa: BLE001
        return Decimal(0)


def cost_type_for(category: str, sub: str) -> str:
    """DOH cost type from the stock category number and sub-category keywords (all three systems)."""
    cat, s = (category or "").strip(), (sub or "").strip().lower()
    num = cat[:2]
    if any(k in (cat + " " + s).lower() for k in ("imaging", "x-ray", "radiograph")):
        return "Imaging"
    if "nitrous" in s or "oxygen" in s:
        return "Pharmacy"                                   # sedation gases, wherever filed
    if num == "02":                                         # drugs, needles & triage
        if "ana" in s and "sthe" in s:                      # local anaesthetics / anesthetics
            return "Pharmacy"
        if "medic" in s:                                    # medicaments / medicines
            return "Pharmacy"
        return "MS"                                         # needles, syringes, vital-signs / emergency supplies
    if num == "05":                                         # implants and accessories
        if "screwdriver" in s or "drilling" in s:
            return "MS"                                     # reusable surgical kits
        return "Prostheses"
    if num == "06":                                         # laboratory: in-house prosthesis materials + external lab lines
        return "Prostheses"
    if num in ("01", "03", "04"):                           # devices / infection control / dental materials
        return "MS"
    if "facility" in cat.lower() or "management of information" in cat.lower():
        return "OHF"                                        # pantry, stationery, computers, MOI
    raise KeyError(f"no mapping for category {category!r} / {sub!r}")


def pool_for(category: str, sub: str) -> str | None:
    """Implant stock pool (fixtures / abutments / bone_graft) or None."""
    if not (category or "").startswith("05"):
        return None
    s = (sub or "").lower()
    if "fixture" in s or "healing" in s or "cover screw" in s:
        return "fixtures"
    if "abutment" in s or "impression" in s:
        return "abutments"
    if "bone" in s or "graft" in s:
        return "bone_graft"
    return None


def _sheet(wb, month: int | None):
    if month is None:
        return next(s for s in wb.worksheets if s.title.lower().startswith("movement"))
    want = f"movement {MONTH_ABBR[month - 1].lower()}"
    try:
        return next(s for s in wb.worksheets if s.title.lower().startswith(want))
    except StopIteration:
        return next(s for s in wb.worksheets if s.title.lower().startswith("movement"))


def _month_of(ws) -> int:
    m = re.search(r"movement\s+([A-Za-z]{3})", ws.title, re.I)
    return MONTH_ABBR.index(m.group(1).title()) + 1 if m else 1


def iter_consumption(report_path: str | Path, month: int | None = None):
    """Yield (category, sub_category, consumed AED) per consuming line, on the month's value basis."""
    wb = openpyxl.load_workbook(report_path, data_only=True, read_only=True)
    ws = _sheet(wb, month)
    m = month or _month_of(ws)
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h).strip() if h is not None else "" for h in next(rows)]
    hi: dict[str, int] = {}
    for i, h in enumerate(hdr):
        hi.setdefault(h, i)
    if m == 12 and "Transaction Type" not in hi:            # December dispatch system
        for r in rows:
            if not r or r[hi["Status"]] not in ("Received", "Returned") or r[hi["Item Category"]] in (None, ""):
                continue
            yield str(r[hi["Item Category"]]), str(r[hi["Item Sub Category"]]), dec(r[hi["Buying Price"]])
        return
    basis = BASIS[m] if BASIS[m] in hi else "Transaction Value"
    for r in rows:
        if not r or r[hi["Transaction Type"]] not in CONSUMING:
            continue
        cat, sub = str(r[hi["Category"]]).strip(), str(r[hi["Sub-Category"]]).strip()
        v = dec(r[hi[basis]])
        if v == 0 and "external lab" in sub.lower():         # August: external-lab line without a supplier price
            v = dec(r[hi["Transaction Value"]])
        yield cat, sub, -v                                   # usage is negative in the sheet; returns positive


def cogs_split(report_path: str | Path, external_lab: Decimal = Decimal(0), month: int | None = None) -> tuple[dict[str, Decimal], list[str]]:
    """Return ({cost type: AED consumed}, log). Consumption is positive."""
    out: dict[str, Decimal] = defaultdict(Decimal)
    detail: dict[tuple, Decimal] = defaultdict(Decimal)
    tag = f"M{month:02d}" if month else "movement"
    for cat, sub, v in iter_consumption(report_path, month):
        ct = cost_type_for(cat, sub)
        out[ct] += v
        detail[(ct, cat, sub)] += v
    out = {ct: v.quantize(D2, rounding=ROUND_HALF_UP) for ct, v in out.items()}
    log = [f"{tag}: {ct:<11}{cat[:46]:<48}{sub[:46]:<48}{float(v):>12,.2f}" for (ct, cat, sub), v in sorted(detail.items())]
    if external_lab:
        out["Prostheses"] = out.get("Prostheses", Decimal(0)) + external_lab
        log.append(f"{tag}: Prostheses  External LAB invoices (Consumption sheet, cash in hand, not stock){'':<24}{float(external_lab):>12,.2f}")
    return dict(out), log


def implant_pools(report_path: str | Path, month: int | None = None) -> dict[str, Decimal]:
    """{pool: AED consumed} for the implant sub-categories (usage net of returns)."""
    pools: dict[str, Decimal] = defaultdict(Decimal)
    for cat, sub, v in iter_consumption(report_path, month):
        p = pool_for(cat, sub)
        if p:
            pools[p] += v
    return {k: v.quantize(D2, rounding=ROUND_HALF_UP) for k, v in pools.items()}


if __name__ == "__main__":
    p = sys.argv[1]
    mo = int(sys.argv[2]) if len(sys.argv) > 2 else None
    ext = Decimal(sys.argv[3]) if len(sys.argv) > 3 else Decimal(0)
    split, log = cogs_split(p, ext, mo)
    print("\n".join(log))
    print(f"\n{'Cost type':<12}{'AED':>14}")
    for ct, v in sorted(split.items(), key=lambda kv: -kv[1]):
        print(f"{ct:<12}{float(v):>14,.2f}")
    print(f"{'TOTAL':<12}{float(sum(split.values())):>14,.2f}")
    print("pools", {k: float(v) for k, v in implant_pools(p, mo).items()})
