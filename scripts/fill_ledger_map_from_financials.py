#!/usr/bin/env python3
"""
Fill the ledger-by-cost-type workbook from Arjun's FY2025 financial workbook
(input/source/Financial_2025_Arjun_Updated_2.xlsx) — evidence, not estimates.

  python3 scripts/fill_ledger_map_from_financials.py \
      --financials input/source/Financial_2025_Arjun_Updated_2.xlsx \
      --template input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx \
      --out input/Ledger_by_Cost_Type_FY2025.xlsx

Sources inside the financial workbook:
  * "Monthly"      — the accrual P&L by expense head and month; its Total row IS the locked ledger.
  * "Salary"       — payroll by employee and month with a Dir/Ind role label.
  * "Consumption " — cost of goods sold by month split into clinic consumables, CSSD, internal lab, external lab.
  * "TB"           — 2025 depreciation charge by asset class (accumulated depreciation movement).
Every head is mapped to one DOH cost type (Standard v2.1 Appendix 1) with the
cost-centre class (Stage 2) recorded in the Mapping sheet. Each month reconciles
to the ledger to the fils.
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from costsub_schema import COST_TYPES, OVERHEAD_COST_TYPES  # noqa: E402
from gl_new import GL_2025  # noqa: E402

D2 = Decimal("0.01")


def dec(v) -> Decimal:
    if v in (None, "", "--") or not isinstance(v, (int, float, Decimal, str)):
        return Decimal(0)
    try:
        return Decimal(str(v).replace(",", "")).quantize(D2, rounding=ROUND_HALF_UP)
    except Exception:  # noqa: BLE001 — header text in a numeric column
        return Decimal(0)


# Expense heads of the Monthly sheet -> (cost type, cost-centre class, basis). Heads split from other
# sheets (Cost of Goods Sold, Salaries & Wages, Depreciation Expense) are handled explicitly below.
HEAD_MAP = {
    "Third Party & Professional Services": ("OHC", "Overhead — shared services",
        "MySmyle Company EMR subscription, management consultancy and IT services (25,000/month) + Calibrate/Platform Technologies: services charged by a sister company = corporate/shared-services overhead"),
    "Legal & Professional Expenses": ("OHF", "Overhead — administration",
        "DOH facility/clinician licences, immigration and visas, trade licence, civil defence, ADHICS/Tasneef, typing/PRO fees"),
    "Repair & Maintenance Expenses": ("OHF", "Overhead — facility", "general premises maintenance supplies and works"),
    "Insurance": ("OHF", "Overhead — administration", "prepaid insurance amortised (staff medical insurance, Daman)"),
    "Staff Qualification And Education": ("OHF", "Overhead — HR / IT",
        "HR administration: visas, work permits, staff benefits, final settlements, training, licences, information-management (IT) charges"),
    "Office Supplies": ("OHF", "Overhead — administration", "stationery, printer ink, newspapers"),
    "Pantry": ("OHF", "Overhead — back office", "pantry, water, cleaning products (staff and patient refreshments)"),
    "Utilities Expenses": ("OHF", "Overhead — facility", "ADDC electricity/water, Etisalat telephone and internet"),
    "Travelling & Transportation": ("OHF", "Overhead — administration", "taxis, parking, petrol, car wash"),
    "Other Expenses": ("OHF", "Overhead — administration", "courier, flags/badges, miscellaneous"),
    "Marketing": ("OHF", "Overhead — administration", "Meta/WhatsApp API, Adobe, promotions"),
    "Dues & Subscriptions": ("OHF", "Overhead — IT / administration", "software subscriptions, memberships"),
    "Equipment Maintenanace": ("GS", "Direct — dental surgeries", "maintenance of dental chairs and clinical equipment"),
    "External Lab": ("Prostheses", "Direct — prostheses", "outsourced dental laboratory invoices (expense head) — prostheses bought from external labs (ruling 13 Oct 2026)"),
    "Clinic Rent": ("OHF", "Overhead — facility", "lease of the clinic premises (unit 1302, 85,000/year), allocated to the Dental bucket by floor space"),
    "Fines": ("OHF", "Overhead — administration", "regulatory and traffic fines (kept: Standard 3.1.1 takes all GL expenses except balance-sheet items)"),
    "Claim Rejection": ("OHF", "Overhead — patient administration", "written-off rejected claims"),
    "Finance Charges": ("OHF", "Overhead — administration", "bank and card charges"),
}
SALARY_MAP = {   # Dir/Ind label in the Salary sheet -> cost type
    "Direct Dr": ("SWDoc", "Direct — dental surgeries", "dentists, oral surgeon, orthodontist (Salary sheet 'Direct Dr')"),
    "Nurse/licensed": ("SWNurs", "Direct — dental surgeries", "licensed dental nurses"),
    "Nurse/No License": ("SWNurs", "Direct — dental surgeries", "dental assistants (unlicensed nurses)"),
    "Cssd": ("SWAllied", "Direct — CSSD (cost centre 07)", "sterilisation technicians"),
    "Indirect/LAB": ("SWAllied", "Direct — prostheses (cost centre 09, in-house CAD/CAM)", "dental technicians making prostheses; routed to the Prosthesis bucket on prosthesis claims (ruling 13 Oct 2026)"),
    "Indirect/RAP": ("OHF", "Overhead — reception / patient administration", "reception and patient-administration staff"),
    "Indirect /INS": ("OHF", "Overhead — patient administration", "insurance claims staff"),
    "Indirect /Management": ("OHF", "Overhead — management", "clinic management"),
    "Indirect /Purchase": ("OHF", "Overhead — administration", "purchasing"),
    "Indirect/Accounts": ("OHF", "Overhead — administration", "accounts"),
    "Indirect it": ("OHF", "Overhead — IT", "IT"),
    "Indirect/ Hr": ("OHF", "Overhead — HR", "HR"),
    "External": ("OHF", "Overhead — administration", "external accountant, drivers, cleaners and other outsourced support staff"),
}
CONSUMPTION_MAP = [   # row-label prefix in the Consumption sheet -> cost type
    ("Consumables -( Dr", "MS", "Direct — dental surgeries", "clinical consumables issued to the dental surgeries"),
    ("Consumables -  ( CSSD", "MS", "Direct — CSSD", "sterilisation and vital-signs consumables"),
    ("Internal LAB", "Prostheses", "Direct — prostheses", "in-house CAD/CAM prosthesis materials (ruling 13 Oct 2026: no laboratory; prostheses)"),
    ("External LAB", "Prostheses", "Direct — prostheses", "prostheses bought from external laboratories, booked in cost of goods sold"),
]
# Staff on-costs (ruling 13 Oct 2026, Option 3 "by the nature of each cost", employees only). Each head's
# on-cost part is split every month across the four staff groups by that month's key:
#   salary    - benefits, bonus, end-of-service, training (they are computed from pay)
#   headcount - visas / work permits / labour cards and staff medical insurance (a fee per person)
#   licensed  - DOH clinician licences, Dataflow, CME: dentists and licensed nurses only, by headcount
# The rest of the head stays as mapped (facility licences, accreditation, trade licence, civil defence, tenancy -> OHF).
ONCOST_PARTS = {
    "Staff Qualification And Education": [("benefits, bonus, end-of-service, training", Decimal(1), "salary")],
    "Insurance": [("staff medical insurance", Decimal(1), "headcount")],
    "Legal & Professional Expenses": [
        ("clinician licences (DOH, Dataflow, CME)", Decimal("34349.57") / Decimal("253569.45475"), "licensed"),
        ("visas, work permits, labour cards", Decimal("145165.32291666666") / Decimal("253569.45475"), "headcount"),
    ],
}
ONCOST_GROUPS = {   # cost type -> Salary sheet roles; the nine external consultants carry no on-costs
    "SWDoc": ["Direct Dr"],
    "SWNurs": ["Nurse/licensed", "Nurse/No License"],
    "SWAllied": ["Cssd", "Indirect/LAB"],
    "OHF": ["Indirect/RAP", "Indirect /INS", "Indirect /Management", "Indirect /Purchase", "Indirect/Accounts", "Indirect it", "Indirect/ Hr"],
}
LICENSED_ROLES = {"SWDoc": ["Direct Dr"], "SWNurs": ["Nurse/licensed"]}
ONCOST_CLASS = {"SWDoc": "Direct — dental surgeries (dentists' on-costs)", "SWNurs": "Direct — dental surgeries (nurses' on-costs)",
                "SWAllied": "Direct — CSSD / prostheses (technicians' on-costs)", "OHF": "Overhead — back-office staff on-costs"}

DEP_DIRECT = {"Acc. Dep - Machinery & Equipment"}   # dental chairs, imaging units, autoclaves
DEP_OVERHEAD = {"Acc. Dep - Furniture & Fittings", "Acc. Dep - Office Equipment", "Acc. Dep - Computer Accessories", "Acc. Dep - Software"}

QUESTIONS = [
    "Q1. Third Party & Professional Services (370,348.18) is mapped to OHC because it is MySmyle Company's EMR/management/IT recharge plus Calibrate and Platform Technologies. If you prefer OHF (no 'corporate office'), say so: one-line change, overhead ratio unchanged.",
    "Q2. Staff on-costs, employees only (ruling 13 Oct 2026, Option 3 by nature of the cost): benefits / bonus / end-of-service / training (889,431.99) by the month's salary share of dentists / nurses / CSSD + technicians / back office; visas / labour (145,165.32) and staff medical insurance (50,409.29) by the month's headcount; clinician licences (34,349.57) to dentists and licensed nurses by headcount. The rest of Legal & Professional (facility licences, accreditation, trade licence, civil defence, tenancy) stays OHF.",
    "Q3. Laboratory cost type is 0 by ruling (13 Oct 2026): the clinic has no laboratory; in-house CAD/CAM materials, external lab invoices and implants are Prostheses, and the dental technicians' pay (SWAllied) is routed to the Prosthesis bucket. Pharmacy and Imaging consumables are split from stock only for the months with a stock-movement sheet; the other months keep them inside MS until the sheets arrive.",
    "Q4. Fines (17,013.03), Claim Rejection (3,402.20) and Finance Charges (63,005.85) are kept in OHF because the locked ledger includes them. The Standard takes all GL expenses except balance-sheet items; confirm you want them in.",
    "Q5. Cost-centre classes follow the cash-ledger coding (clinics 01-06, CSSD 07, OPG 08, LAB 09 = direct; reception 10-11, back office 12, server room 13 = overhead).",
]


def read_monthly(ws) -> dict[str, list[Decimal]]:
    out = {}
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[1], str) and r[1].strip():
            out[r[1].strip()] = [dec(v) for v in r[2:14]]
    return out


def read_salary(ws, headcount: bool = False) -> dict[str, list[Decimal]]:
    """{role: 12 monthly salary totals}; with headcount=True, {role: 12 monthly counts of persons paid}."""
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if r[0] == "Name")
    mon_col = {12 - k: 2 + k for k in range(12)}          # Dec2025 .. Jan 2025 in columns 2..13
    out: dict[str, list[Decimal]] = defaultdict(lambda: [Decimal(0)] * 12)
    for r in rows[hdr_i + 1:]:
        if not r[0] or r[1] in (None, ""):
            continue
        cat = str(r[1]).strip()
        for m, ci in mon_col.items():
            v = dec(r[ci])
            out[cat][m - 1] += (Decimal(1) if v else Decimal(0)) if headcount else v
    return out


def read_consumption(ws) -> dict[str, list[Decimal]]:
    out = {}
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[0], str):
            for prefix, *_ in CONSUMPTION_MAP:
                if r[0].startswith(prefix):
                    out[prefix] = [-dec(v) for v in r[1:13]]
    return out


def read_depreciation(ws) -> tuple[Decimal, Decimal]:
    direct = overhead = Decimal(0)
    for r in ws.iter_rows(values_only=True):
        if r and r[0] in DEP_DIRECT:
            direct += dec(r[7])
        elif r and r[0] in DEP_OVERHEAD:
            overhead += dec(r[7])
    return direct, overhead


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--financials", default="input/source/Financial_2025_Arjun_Updated_2.xlsx")
    ap.add_argument("--template", default="input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx")
    ap.add_argument("--out", default="input/Ledger_by_Cost_Type_FY2025.xlsx")
    ap.add_argument("--movement", action="append", default=[], metavar="M=PATH",
                    help="month number = monthly Financial Report with a 'Movement' sheet; its stock usage replaces "
                         "the Consumption-sheet split of Cost of Goods Sold for that month (e.g. 1=input/source/Financial_Report_Jan_2025_1.xlsx)")
    args = ap.parse_args(argv)
    movement = {}
    for spec in args.movement:
        m, path = spec.split("=", 1)
        movement[int(m)] = path

    fin = openpyxl.load_workbook(args.financials, data_only=True, read_only=True)
    monthly = read_monthly(fin["Monthly"])
    salary = read_salary(fin["Salary"])
    cons = read_consumption(fin["Consumption "])
    dep_direct, dep_oh = read_depreciation(fin["TB"])

    # --- build the mapping lines: (source, head, cost type, class, basis, 12 months) ---
    lines: list[tuple] = []
    from cogs_from_movement import cogs_split
    mv_split: dict[int, dict] = {}
    mv_log: list[str] = []
    for m, path in movement.items():
        ext = cons["External LAB"][m - 1]            # external lab invoices are not stock
        split, log = cogs_split(path, ext)
        mv_split[m] = split
        mv_log += log
        tot = sum(split.values())
        diff = monthly["Cost of Goods Sold"][m - 1] - tot
        if abs(diff) > Decimal("0.10"):
            raise SystemExit(f"movement file for month {m}: usage {tot} != Cost of Goods Sold {monthly['Cost of Goods Sold'][m - 1]}")
        if diff:
            split["MS"] = split.get("MS", Decimal(0)) + diff
            log.append(f"month {m}: {diff:+} fils rounding between stock lines and the P&L absorbed on MS")
    for prefix, ct, cls, basis in CONSUMPTION_MAP:
        vals = [Decimal(0) if (m + 1) in mv_split else cons[prefix][m] for m in range(12)]
        lines.append(("Consumption sheet", f"Cost of Goods Sold: {prefix.strip(' -(')}", ct, cls, basis, vals))
    mv_types = sorted({ct for sp in mv_split.values() for ct in sp})
    for ct in mv_types:
        vals = [mv_split.get(m + 1, {}).get(ct, Decimal(0)) for m in range(12)]
        cls = {"Prostheses": "Direct — prostheses (implants + in-house lab + external lab)", "Pharmacy": "Direct — dental surgeries",
               "MS": "Direct — dental surgeries / CSSD", "OHF": "Overhead — back office"}.get(ct, "Direct")
        basis = {"Prostheses": "implant fixtures (net of internal returns), healing/cover screws, abutments, implant impression tools, bone-graft materials; in-house CAD/CAM materials (blocks, blanks, 3D inks, stone, milling tools); external lab invoices",
                 "Pharmacy": "local anaesthetic cartridges, nitrous oxide and oxygen",
                 "MS": "restorative, endodontic, burs, impression, bleaching, surgical, infection-control supplies, needles/syringes, hand instruments",
                 "OHF": "pantry and stationery issued from stock"}.get(ct, "")
        lines.append(("Stock movement", f"Cost of Goods Sold ({', '.join(f'M{m:02d}' for m in sorted(mv_split))}): {ct}", ct, cls, basis, vals))
    for cat, vals in salary.items():
        ct, cls, basis = SALARY_MAP[cat]
        lines.append(("Salary sheet", f"Salaries & Wages: {cat}", ct, cls, basis, vals))
    dep_m = monthly["Depreciation Expense"]
    tot_dep = dep_direct + dep_oh
    lines.append(("TB depreciation", "Depreciation: Machinery & Equipment", "Depreciation", "Direct — dental surgeries / imaging",
                  "dental chairs, imaging units, autoclaves (2025 charge 126,679.42)", [(v * dep_direct / tot_dep).quantize(D2) for v in dep_m]))
    lines.append(("TB depreciation", "Depreciation: furniture, office equipment, computers, software", "OHDepreciation", "Overhead — facility / IT",
                  "fit-out, office equipment, computers, software (2025 charge 30,297.55)", [(v * dep_oh / tot_dep).quantize(D2) for v in dep_m]))
    headcount = read_salary(fin["Salary"], headcount=True)
    keys = {
        "salary": {g: [sum(salary.get(r, [Decimal(0)] * 12)[m] for r in roles) for m in range(12)] for g, roles in ONCOST_GROUPS.items()},
        "headcount": {g: [sum(headcount.get(r, [Decimal(0)] * 12)[m] for r in roles) for m in range(12)] for g, roles in ONCOST_GROUPS.items()},
        "licensed": {g: [sum(headcount.get(r, [Decimal(0)] * 12)[m] for r in LICENSED_ROLES.get(g, [])) for m in range(12)] for g in ONCOST_GROUPS},
    }
    for head, (ct, cls, basis) in HEAD_MAP.items():
        if head not in monthly:
            raise SystemExit(f"head {head!r} not found in the Monthly sheet")
        parts = ONCOST_PARTS.get(head)
        if not parts:
            lines.append(("Monthly sheet", head, ct, cls, basis, monthly[head]))
            continue
        rest = list(monthly[head])
        for label, share, key in parts:
            part = [(v * share).quantize(D2, rounding=ROUND_HALF_UP) for v in monthly[head]]
            rest = [r_ - p_ for r_, p_ in zip(rest, part)]
            split = {g: [Decimal(0)] * 12 for g in ONCOST_GROUPS}
            for m in range(12):
                w = {g: keys[key][g][m] for g in ONCOST_GROUPS}
                base = sum(w.values())
                alloc = {g: (part[m] * w[g] / base).quantize(D2, rounding=ROUND_HALF_UP) for g in ONCOST_GROUPS}
                big = max(ONCOST_GROUPS, key=lambda g: w[g])
                alloc[big] += part[m] - sum(alloc.values())          # rounding residual on the largest group
                for g in ONCOST_GROUPS:
                    split[g][m] = alloc[g]
            for g in ONCOST_GROUPS:
                if any(split[g]):
                    lines.append(("Monthly sheet", f"{head}: {label} → {g}", g, ONCOST_CLASS[g],
                                  f"staff on-cost; monthly {key} key ({', '.join(LICENSED_ROLES.get(g, []) if key == 'licensed' else ONCOST_GROUPS[g])}); ruling 13 Oct 2026, Option 3", split[g]))
        if any(rest):
            lines.append(("Monthly sheet", f"{head}: non-on-cost part", ct, cls, basis + " (facility licences, accreditation, trade licence, civil defence, tenancy)", rest))

    # --- check the split heads reproduce the Monthly sheet, and every month the ledger ---
    checks = []
    for head, src_lines in (("Cost of Goods Sold", [l for l in lines if l[0] in ("Consumption sheet", "Stock movement")]),
                            ("Salaries & Wages", [l for l in lines if l[0] == "Salary sheet"]),
                            ("Depreciation Expense", [l for l in lines if l[0] == "TB depreciation"])):
        for m in range(12):
            s = sum(l[5][m] for l in src_lines)
            if s != monthly[head][m]:
                checks.append(f"{head} month {m+1}: split {s} != Monthly {monthly[head][m]}")
    totals = [sum(l[5][m] for l in lines) for m in range(12)]
    for m in range(12):
        gl = GL_2025[m + 1]
        if totals[m] != gl:
            checks.append(f"month {m+1}: mapped total {totals[m]} != ledger {gl} ({totals[m]-gl:+})")
    if checks:
        print("RECONCILIATION DIFFERENCES:\n  " + "\n  ".join(checks))

    # --- aggregate to cost types ---
    ct_month: dict[str, list[Decimal]] = {t: [Decimal(0)] * 12 for t in COST_TYPES}
    for _, _, ct, _, _, vals in lines:
        for m in range(12):
            ct_month[ct][m] += vals[m]
    # the Monthly sheet is 5.70 short of the P&L in total (a blank sub-category); tie each month to the locked ledger on OHF
    for m in range(12):
        diff = GL_2025[m + 1] - sum(ct_month[t][m] for t in COST_TYPES)
        if diff:
            ct_month["OHF"][m] += diff
            checks.append(f"month {m+1}: {diff:+} added to OHF to tie to the locked ledger")

    # --- write the workbook ---
    wb = openpyxl.load_workbook(args.template)
    ws = wb["Cost types by month"]
    DESCRIPTIONS = {   # ruling 13 Oct 2026 — overrides the template wording
        "Laboratory": "Pathology laboratory supplies — none: the clinic has no laboratory (ruling 13 Oct 2026); dental lab work is Prostheses",
        "Prostheses": "Implants, abutments, bone graft; crowns, bridges, veneers, onlays, dentures and every other lab-fabricated prosthesis (in-house CAD/CAM materials and external lab invoices)",
        "SWAllied": "Salaries and wages — allied health incl. technicians (CSSD technicians; dental technicians, routed to the Prosthesis bucket)",
    }
    for r in ws.iter_rows(min_row=2):
        if r[0].value in COST_TYPES:
            if r[0].value in DESCRIPTIONS:
                r[1].value = DESCRIPTIONS[r[0].value]
            for m in range(12):
                r[2 + m].value = float(ct_month[r[0].value][m])   # columns C..N
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="D9E1F2")
    mp = wb.create_sheet("Mapping", 1)
    mp.append(["Source", "Ledger head / split", "DOH cost type", "Cost-centre class (Stage 2)", "Basis",
               "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "FY2025"])
    for c in mp[1]:
        c.font = bold
        c.fill = head_fill
    for src, head, ct, cls, basis, vals in sorted(lines, key=lambda l: (COST_TYPES.index(l[2]), -sum(l[5]))):
        mp.append([src, head, ct, cls, basis, *[float(v) for v in vals], float(sum(vals))])
    mp.append([])
    mp.append(["", "TOTAL mapped", "", "", "", *[float(t) for t in totals], float(sum(totals))])
    mp.append(["", "Ledger (locked)", "", "", "", *[float(GL_2025[m]) for m in range(1, 13)], float(sum(GL_2025.values()))])
    for c in mp[mp.max_row - 1] + mp[mp.max_row]:
        c.font = bold
    for col, w in zip("ABCDE", (16, 48, 14, 36, 90)):
        mp.column_dimensions[col].width = w
    for r in mp.iter_rows(min_row=2):
        for c in r[5:]:
            c.number_format = "#,##0.00"

    q = wb.create_sheet("Questions", 2)
    q.append(["Mapping built from Arjun's FY2025 financial workbook (Monthly P&L, Salary, Consumption, TB). Points to confirm:"])
    q["A1"].font = bold
    for line in QUESTIONS:
        q.append([line])
    q.append([])
    oh = [sum(ct_month[t][m] for t in OVERHEAD_COST_TYPES) for m in range(12)]
    q.append([f"Overhead ratio by month: " + ", ".join(f"{m+1:02d} {float(oh[m] / GL_2025[m+1] * 100):.1f}%" for m in range(12))
              + f"; FY2025 {float(sum(oh) / sum(GL_2025.values()) * 100):.1f}% (DOH band 20% to under 50%)"])
    if checks:
        q.append([])
        q.append(["Reconciliation notes:"])
        for c in checks:
            q.append([c])
    q.column_dimensions["A"].width = 160
    if mv_log:
        cd = wb.create_sheet("Consumables detail", 3)
        cd.append(["Stock movement -> cost type (usage net of internal returns; exchanges cancel; external lab from the P&L)"])
        cd["A1"].font = bold
        for line in mv_log:
            cd.append([line])
        cd.column_dimensions["A"].width = 150
    bs = wb.create_sheet("Bucket shares")
    bs.append(["Rule", *[f"M{m:02d}" for m in range(1, 13)], "Basis"])
    for c in bs[1]:
        c.font = bold
        c.fill = head_fill
    lab = salary.get("Indirect/LAB", [Decimal(0)] * 12)
    cssd = salary.get("Cssd", [Decimal(0)] * 12)
    bs.append(["SWAllied lab share", *[float((lab[m] / (lab[m] + cssd[m])) if (lab[m] + cssd[m]) else 0) for m in range(12)],
               "dental technicians / (dental technicians + CSSD technicians), Salary sheet; the engine pre-assigns this share of SWAllied to the Prosthesis bucket on the prosthesis-coded claims (ruling 13 Oct 2026)"])
    bs.append(["SWDoc to Anaesthesia", *[0.02] * 12, "ruling 10 Oct 2026"])
    bs.append(["SWNurs to Anaesthesia", *[0.01] * 12, "ruling 10 Oct 2026"])
    bs.append(["SWNurs to Imaging", *[0.04] * 12, "ruling 10 Oct 2026"])
    bs.append(["Depreciation to Imaging", *[0.04] * 12, "ruling 10 Oct 2026 (equipment depreciation; OPG / X-ray units)"])
    for r in bs.iter_rows(min_row=2):
        for c in r[1:13]:
            c.number_format = "0.0%"
    bs.column_dimensions["A"].width = 26
    bs.column_dimensions["N"].width = 100
    if "Assumptions" in wb.sheetnames:
        del wb["Assumptions"]
    wb.save(args.out)

    print(f"wrote {args.out}")
    print(f"{'Cost type':<16}" + "".join(f"{m:>12}" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")) + f"{'FY2025':>14}{'share':>8}")
    grand = sum(GL_2025.values())
    for t in COST_TYPES:
        v = ct_month[t]
        if any(v):
            print(f"{t:<16}" + "".join(f"{float(x):>12,.0f}" for x in v) + f"{float(sum(v)):>14,.2f}{float(sum(v) / grand * 100):>7.1f}%")
    print(f"{'TOTAL':<16}" + "".join(f"{float(x):>12,.0f}" for x in totals) + f"{float(sum(totals)):>14,.2f}")
    print(f"{'Ledger':<16}" + "".join(f"{float(GL_2025[m]):>12,.0f}" for m in range(1, 13)) + f"{float(grand):>14,.2f}")
    print(f"overhead ratio FY2025: {float(sum(oh) / grand * 100):.1f}%  (monthly min {min(float(oh[m]/GL_2025[m+1]*100) for m in range(12)):.1f}%, max {max(float(oh[m]/GL_2025[m+1]*100) for m in range(12)):.1f}%)")
    for c in checks:
        print("note:", c)
    return 0


if __name__ == "__main__":
    sys.exit(main())
