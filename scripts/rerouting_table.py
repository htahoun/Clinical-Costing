#!/usr/bin/env python3
"""
Rerouting dictionary: Vision Dental Clinic's internal ledger categories -> DOH cost
types (Standard v2.1 Appendix 1), cost-centre class (Stage 2) and DOH bucket
(scripts/bucket_rules.py).

  python3 scripts/rerouting_table.py --out output/Rerouting_Internal_to_DOH.xlsx

Sources: the Monthly P&L heads (locked ledger), the Salary sheet role labels,
the Consumption sheet, the stock-movement categories (January), the TB
depreciation classes and the cash-ledger sub-categories (indicative amounts
for heads whose accrual lines are not itemised).
Confidence: CONFIRMED = ruling or itemised source; INFERRED = composition read
from the cash ledger; NEEDS GL DETAIL = the accrual head mixes items that map
differently and only Arjun's account-level lines can split it.
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from bucket_rules import RULES  # noqa: E402

FIN = "input/source/Financial_2025_Arjun_Updated_2.xlsx"

# (internal head, internal sub-category / line, plain meaning, DOH cost type, cost-centre class, Standard basis, confidence)
ROWS = [
    # ---- Cost of Goods Sold (stock issued to the clinic; from the monthly stock movement) ----
    ("Cost of Goods Sold", "02- Drugs / 01- Local Anaesthetics", "local anaesthetic cartridges", "Pharmacy", "Direct", "App.1 Pharmacy: drugs and services", "CONFIRMED (ruling)"),
    ("Cost of Goods Sold", "04- Dental Materials / 02- Nitrous & Oxygen Sedation", "sedation gases", "Pharmacy", "Direct", "App.1 Pharmacy: drugs and services", "CONFIRMED (ruling)"),
    ("Cost of Goods Sold", "02- Drugs / 03- Vital Signs, 04- Needles & Syringes", "needles, syringes, emergency supplies", "MS", "Direct", "App.1 MS: all other medical and surgical supplies", "CONFIRMED"),
    ("Cost of Goods Sold", "05- Implants (fixtures, screws, abutments, impression tools, bone graft)", "implants and implant accessories", "Prostheses", "Direct", "App.1 Prostheses: implantable devices and human tissue", "CONFIRMED (ruling)"),
    ("Cost of Goods Sold", "06- Laboratory (blocks, blanks, 3D inks, stone, stains, milling tools)", "in-house CAD/CAM prosthesis materials", "Prostheses", "Direct", "Ruling 13 Oct 2026: the clinic has no laboratory; prostheses made in-house (App.2 Prosthesis bucket: prothetics, implants)", "CONFIRMED (ruling)"),
    ("Cost of Goods Sold", "External LAB (cost of goods sold line)", "prostheses bought from external laboratories", "Prostheses", "Direct", "Ruling 13 Oct 2026: prostheses at invoiced cost (App.2 Prosthesis bucket)", "CONFIRMED (ruling)"),
    ("Cost of Goods Sold", "04- Dental Materials (restorative, endodontic, burs, impression, bleaching, scaling, surgical, orthodontic)", "dental consumables", "MS", "Direct", "App.1 MS", "CONFIRMED"),
    ("Cost of Goods Sold", "03- Prevention & Control of Infection (PPE, disinfectants, wraps, monitoring)", "infection-control supplies", "MS", "Direct", "App.1 MS", "CONFIRMED"),
    ("Cost of Goods Sold", "01- Medical & LAB Devices / hand instruments expensed on issue", "small instruments and devices", "MS", "Direct", "App.1 MS", "CONFIRMED"),
    ("Cost of Goods Sold", "08- Facility Management (pantry, stationery from stock)", "pantry and stationery issued from stock", "OHF", "Overhead", "App.1 OHF: expenses of overhead cost centres", "CONFIRMED"),
    # ---- Salaries & Wages (payroll sheet role labels) ----
    ("Salaries & Wages", "Direct Dr", "dentists, oral surgeon, orthodontist", "SWDoc", "Direct", "App.1 SWDoc: salaries and wages medical", "CONFIRMED"),
    ("Salaries & Wages", "Nurse/licensed, Nurse/No License", "dental nurses and assistants", "SWNurs", "Direct", "App.1 SWNurs", "CONFIRMED"),
    ("Salaries & Wages", "Cssd", "sterilisation technicians", "SWAllied", "Direct", "App.1 SWAllied incl. technicians (ruling: CSSD is direct)", "CONFIRMED (ruling)"),
    ("Salaries & Wages", "Indirect/LAB", "dental technicians (in-house prostheses)", "SWAllied", "Direct", "App.1 SWAllied incl. technicians; bucket = Prosthesis on prosthesis-coded claims (ruling 13 Oct 2026)", "CONFIRMED (ruling)"),
    ("Salaries & Wages", "Indirect/RAP", "registrations and appointments (reception, patient administration)", "OHF", "Overhead", "App.3 Patient Administration is an overhead cost centre", "CONFIRMED (your definition)"),
    ("Salaries & Wages", "Indirect /INS", "insurance department management (claims)", "OHF", "Overhead", "App.3 Patient Administration / Finance overhead cost centres", "CONFIRMED (your definition)"),
    ("Salaries & Wages", "Indirect /Management", "clinic management", "OHF", "Overhead", "App.3 Executive Services (Management)", "CONFIRMED"),
    ("Salaries & Wages", "Indirect /Purchase, Indirect/Accounts, Indirect it, Indirect/ Hr", "purchasing, accounts, IT, HR", "OHF", "Overhead", "App.3 Finance / Computing / Human Resources", "CONFIRMED"),
    ("Salaries & Wages", "External", "external staff consultants: accountants, quality managers, insurance auditors", "OHF", "Overhead", "Definition 2.24 overhead: finance, quality and administrative functions; App.3 Finance", "CONFIRMED (your definition)"),
    # ---- Third Party & Professional Services ----
    ("Third Party & Professional Services", "08- Management of Information: MySmyle EMR subscription, management consultancy and IT services", "sister-company recharge for EMR, management and IT", "OHC", "Overhead", "App.1 OHC: overheads from corporate office / shared services", "INFERRED (cash ledger)"),
    ("Third Party & Professional Services", "07- 3rd Party Consultancy: Calibrate claims service, Platform Technologies costing submission", "outsourced claims and costing services", "OHF", "Overhead", "App.3 Patient Administration / Finance", "INFERRED (cash ledger)"),
    # ---- Legal & Professional Expenses ----
    ("Legal & Professional Expenses", "04- DOH: facility licence renewal", "facility licence", "OHF", "Overhead", "facility-level regulatory cost", "NEEDS GL DETAIL"),
    ("Legal & Professional Expenses", "04- DOH: clinician licence issuance / renewal, Dataflow, CME", "dentists' and nurses' professional licences", "SWDoc / SWNurs (on-cost)", "Direct", "App.1 Cost of Labour: salaries, wages and on-costs follow the staff category", "NEEDS GL DETAIL"),
    ("Legal & Professional Expenses", "02- Immigration and Labor: visas, work permits, labour cards", "staff visa and labour on-costs", "on-cost by staff category (SWDoc / SWNurs / SWAllied / OHF)", "Mixed", "App.1 Cost of Labour: on-costs", "NEEDS GL DETAIL"),
    ("Legal & Professional Expenses", "01- Trade licence, 03- Civil Defense, 15- Tenancy contract fees", "trade licence, fire safety, lease registration", "OHF", "Overhead", "App.1 OHF", "INFERRED (cash ledger)"),
    ("Legal & Professional Expenses", "05- Jawda / Tasneef / ADHICS / Malaffi", "accreditation and information-security audits", "OHF", "Overhead", "App.1 OHF; App.3 Executive Services", "INFERRED (cash ledger)"),
    ("Legal & Professional Expenses", "10- Transportation (PRO runs)", "government-service transport", "OHF", "Overhead", "App.1 OHF", "INFERRED (cash ledger)"),
    # ---- Staff Qualification And Education ----
    ("Staff Qualification And Education", "04- Staff Benefits and Bonus, 07- Staff Final Settlement, 05- Training", "benefits, bonuses, end-of-service, training", "on-cost by staff category (SWDoc / SWNurs / SWAllied / OHF)", "Mixed", "App.1 Cost of Labour: on-costs follow the staff category", "NEEDS GL DETAIL"),
    ("Staff Qualification And Education", "02- Immigration and Labor (if booked here rather than Legal)", "visas and labour on-costs", "on-cost by staff category", "Mixed", "App.1 Cost of Labour: on-costs", "NEEDS GL DETAIL"),
    ("Staff Qualification And Education", "08- Management of Information (if booked here): software, servers, IT security", "clinic IT", "OHF", "Overhead", "App.3 Computing", "NEEDS GL DETAIL"),
    ("Staff Qualification And Education", "MySmyle staff salaries paid through VDC (cash ledger 'Mysmyle Indirect Staff Salary')", "sister-company payroll routed through the clinic", "OHC if VDC bears it; excluded if recoverable from MySmyle", "Overhead", "App.1 OHC: shared services; Standard 3.1.2 offsets and recoveries", "NEEDS GL DETAIL"),
    # ---- Smaller heads ----
    ("Repair & Maintenance Expenses", "01- General Maintenance", "premises maintenance supplies and works", "OHF", "Overhead", "App.3 Maintenance, Engineering & Repairs", "CONFIRMED"),
    ("Insurance", "Daman staff health insurance (prepaid, amortised)", "staff medical insurance", "on-cost by staff category (SWDoc / SWNurs / SWAllied / OHF)", "Mixed", "App.1 Cost of Labour: on-costs", "INFERRED (Insurance Exp sheet)"),
    ("Insurance", "Malpractice / property insurance (if any)", "professional indemnity, property", "SWDoc on-cost / OHF", "Mixed", "App.3 Medical Indemnity Insurance; Insurance - Property", "NEEDS GL DETAIL"),
    ("Office Supplies", "02- Office Supplies and Stationery", "stationery, printer ink, newspapers", "OHF", "Overhead", "App.1 OHF", "CONFIRMED"),
    ("Pantry", "05- Pantry: staff refreshments, cleaning products, water", "pantry and cleaning consumables", "OHF", "Overhead", "App.1 OHF (patient refreshments would be Hotel; immaterial)", "CONFIRMED"),
    ("Utilities Expenses", "03- Utilities and Communications: ADDC, Etisalat, Doocti telephony", "electricity, water, telephone, internet", "OHF", "Overhead", "App.3 Utilities / Telecommunications", "CONFIRMED"),
    ("Travelling & Transportation", "04- Transportation and Vehicle", "taxis, parking, fuel", "OHF", "Overhead", "App.1 OHF", "CONFIRMED"),
    ("Other Expenses", "07- Miscellaneous: courier, flags, sundries", "sundry administration", "OHF", "Overhead", "App.1 OHF", "CONFIRMED"),
    ("Marketing", "09- Marketing and Advertisement: Meta, WhatsApp API, Adobe, promotions", "marketing", "OHF", "Overhead", "App.3 Public Relations & Marketing", "CONFIRMED"),
    ("Dues & Subscriptions", "software subscriptions (Dok32, Microsoft 365, hosting), memberships", "clinic software and memberships", "OHF", "Overhead", "App.3 Computing", "INFERRED (cash ledger)"),
    ("Equipment Maintenanace", "04- Equipments Maintenance", "dental chairs and clinical equipment service", "GS", "Direct", "App.1 GS: other goods and services in direct cost centres", "CONFIRMED"),
    ("External Lab", "05/06- External Lab invoices (expense head)", "prostheses bought from external laboratories", "Prostheses", "Direct", "Ruling 13 Oct 2026: prostheses at invoiced cost (App.2 Prosthesis bucket)", "CONFIRMED (ruling)"),
    ("Clinic Rent", "06- Rent: unit 1302, 85,000 / year", "clinic premises lease", "OHF", "Overhead", "facility overhead allocated to Dental by floor space (App.3 Floor)", "CONFIRMED"),
    ("Fines", "regulatory, traffic and insurance fines", "fines", "OHF", "Overhead", "Standard 3.1.1: all GL expenses except balance-sheet items (kept; confirm)", "CONFIRMED (kept)"),
    ("Depreciation Expense", "Acc. Dep - Machinery & Equipment", "dental chairs, OPG / X-ray, autoclaves, lab equipment", "Depreciation", "Direct", "App.1 Depreciation in direct cost centres (4 % to Imaging bucket by ruling)", "CONFIRMED"),
    ("Depreciation Expense", "Acc. Dep - Furniture & Fittings, Office Equipment, Computers, Software", "fit-out, office equipment, IT", "OHDepreciation", "Overhead", "App.1 OH Depreciation", "CONFIRMED"),
    ("Claim Rejection", "written-off rejected claims", "revenue write-off booked as expense", "OHF", "Overhead", "kept per locked ledger; arguably a revenue adjustment (confirm)", "CONFIRMED (kept)"),
    ("Finance Charges", "11- Bank Charges: card and bank fees", "bank and card charges", "OHF", "Overhead", "App.1 OHF (financing costs are not excluded by the Standard; confirm)", "CONFIRMED (kept)"),
]

MISSING = [
    "M1. Staff Qualification And Education (889,431.99): the account-level lines. Which part is staff on-costs (benefits, end-of-service, training, visas) by staff category, which part is IT (Management of Information), and whether MySmyle staff salaries paid through VDC sit inside it.",
    "M2. Legal & Professional Expenses (253,569.45): split between facility licence / trade licence / civil defence / accreditation (overhead) and clinician licences and visas (labour on-costs by staff category).",
    "M3. Insurance (50,409.29): staff medical (on-cost by staff category) versus malpractice or property.",
    "M4. Dues & Subscriptions (80,526.17): which subscriptions serve the clinic rooms (direct GS) versus administration/IT (OHF).",
    "M5. Third Party & Professional Services (370,348.18): confirm it is the MySmyle EMR/management/IT recharge (OHC) plus Calibrate / Platform Technologies (OHF).",
    "M6. For on-costs: the allocation key you want (headcount per category, or salary per category) — I propose salary per category.",
    "M7. The eleven monthly stock-movement sheets (February to December) to split Cost of Goods Sold as January was split.",
]


# Account-level accrual amounts returned by the Finance Director on 13 Oct 2026
# (input/source/Rerouting_Internal_to_DOH_2_FinanceDirector_2026-10-13.xlsx); every split ties to its head.
LINE_AMOUNTS = {
    ("Third Party & Professional Services", "08- Management of Information"): 351150.23,
    ("Third Party & Professional Services", "07- 3rd Party Consultancy"): 19197.95,
    ("Legal & Professional Expenses", "04- DOH: facility"): 34009.68,
    ("Legal & Professional Expenses", "04- DOH: clinician"): 34349.57,
    ("Legal & Professional Expenses", "02- Immigration"): 145165.32291666666,
    ("Legal & Professional Expenses", "01- Trade licence"): 9133.03183333333,
    ("Legal & Professional Expenses", "05- Jawda"): 30911.85,
    ("Legal & Professional Expenses", "10- Transportation"): 0.0,
    ("Staff Qualification And Education", "04- Staff Benefits"): 889431.99,
    ("Staff Qualification And Education", "02- Immigration"): 0.0,
    ("Staff Qualification And Education", "08- Management"): 0.0,
    ("Staff Qualification And Education", "MySmyle staff"): 0.0,
    ("Insurance", "Malpractice"): 0.0,
}
# January 2025 cost of goods sold by stock line (Financial_Report_Jan_2025_1.xlsx, usage net of internal returns;
# external lab from the Consumption sheet). Feb-Dec await the eleven monthly movement sheets.
COGS_JAN = {
    "02- Drugs / 01- Local": 682.00, "04- Dental Materials / 02-": 570.00, "02- Drugs / 03-": 2338.48,
    "05- Implants": 52796.00, "06- Laboratory": 13538.94, "External LAB": 1870.00,
    "04- Dental Materials (restorative": 17571.77, "03- Prevention": 6169.15, "01- Medical": 5583.33, "08- Facility": 1007.80,
}


def head_totals():
    wb = openpyxl.load_workbook(FIN, data_only=True, read_only=True)
    ws = wb["Monthly"]
    out = {}
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[1], str) and r[1].strip():
            vals = [float(v) if isinstance(v, (int, float)) else 0.0 for v in r[2:14]]
            out[r[1].strip()] = sum(vals)
    sal = defaultdict(float)
    ws = wb["Salary"]
    rows = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r[0] == "Name")
    for r in rows[hi + 1:]:
        if r[0] and r[1]:
            try:
                sal[str(r[1]).strip()] += float(r[16] or 0)
            except (TypeError, ValueError):
                pass
    cash = defaultdict(float)
    ws = wb["Cash_in_hand_2025-01-01_to_2025"]
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h) for h in next(rows)]
    hx = {h: i for i, h in enumerate(hdr)}
    for r in rows:
        if r and r[hx["Transaction Type"]] == "01- Payment":
            try:
                cash[(str(r[hx["Category"]])[:3], str(r[hx["Sub-Category"]])[:3])] += float(r[hx["Remarks"]])
            except (TypeError, ValueError):
                pass
    cons = {}
    ws = wb["Consumption "]
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[0], str):
            if r[0].startswith("Internal LAB"):
                cons["Internal LAB"] = -float(r[13] or 0)
            elif r[0].startswith("External LAB"):
                cons["External LAB"] = -float(r[13] or 0)
    return out, sal, cash, cons


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="output/Rerouting_Internal_to_DOH.xlsx")
    args = ap.parse_args(argv)
    heads, sal, cash, cons = head_totals()
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="D9E1F2")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rerouting table"
    ws.append(["Internal ledger head (our language)", "FY2025 AED (accrual): the line where Arjun split the head, else the head total",
               "Internal sub-category / line", "What it is",
               "DOH cost type (Appendix 1)", "Cost-centre class (Stage 2)", "DOH bucket (bucket rules)", "Standard basis", "Confidence",
               "Line detail: cash-ledger indication, or Jan 2025 from the stock movement (Feb-Dec awaited)"])
    for c in ws[1]:
        c.font = bold
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="top")
    sal_map = {"Direct Dr": ["Direct Dr"], "Nurse/licensed, Nurse/No License": ["Nurse/licensed", "Nurse/No License"], "Cssd": ["Cssd"],
               "Indirect/LAB": ["Indirect/LAB"], "Indirect/RAP": ["Indirect/RAP"], "Indirect /INS": ["Indirect /INS"],
               "Indirect /Management": ["Indirect /Management"], "External": ["External"],
               "Indirect /Purchase, Indirect/Accounts, Indirect it, Indirect/ Hr": ["Indirect /Purchase", "Indirect/Accounts", "Indirect it", "Indirect/ Hr"]}
    cash_hint = {"02- Drugs / 01- Local": None}
    for head, sub, meaning, ct, cls, basis, conf in ROWS:
        bucket = ""
        key = ct.split(" ")[0].split("/")[0]
        if key in RULES:
            bucket = ", ".join(f"{b} {float(s)*100:.0f}%" if not isinstance(s, str) else f"{b} (lab share)" for b, s, _ in RULES[key])
        elif "on-cost" in ct:
            bucket = "follows the staff category"
        ind = ""
        if head == "Salaries & Wages" and sub in sal_map:
            ind = sum(sal.get(k, 0.0) for k in sal_map[sub])
        elif head == "Cost of Goods Sold":
            jan = next((v for k, v in COGS_JAN.items() if sub.startswith(k)), None)
            ind = f"Jan 2025 from stock movement: {jan:,.2f}; Feb-Dec: awaiting the eleven monthly movement sheets" if jan is not None else ""
        elif head == "Depreciation Expense":
            ind = 126679.42 if "Machinery" in sub else 30297.55
        else:
            # cash sub-categories named in the line, matched by their numeric prefix within the right cash category
            cat_for_head = {"Third Party & Professional Services": "04-", "Legal & Professional Expenses": "04-", "Staff Qualification And Education": "03-",
                            "Repair & Maintenance Expenses": "02-", "Office Supplies": "02-", "Pantry": "02-", "Utilities Expenses": "02-",
                            "Travelling & Transportation": "02-", "Other Expenses": "02-", "Marketing": "04-", "Equipment Maintenanace": "01-",
                            "External Lab": "01-", "Clinic Rent": "02-", "Finance Charges": "04-"}
            c3 = cat_for_head.get(head)
            import re
            tags = re.findall(r"(\d{2}-)", sub)
            if c3 and tags:
                tot = sum(v for (cc, s3), v in cash.items() if cc == c3 and s3 in tags)
                if head == "Staff Qualification And Education" and "08-" in tags:
                    tot = sum(v for (cc, s3), v in cash.items() if cc == "04-" and s3 == "08-")
                if head in ("Legal & Professional Expenses", "Staff Qualification And Education") and "02-" in tags:
                    tot = sum(v for (cc, s3), v in cash.items() if cc == "04-" and s3 == "02-")
                ind = tot if tot else ""
        if sub == "Indirect/LAB":
            bucket = "Prosthesis 100% (dental technicians' pool, pre-assigned on the prosthesis-coded claims)"
        line_amt = next((v for (h, pfx), v in LINE_AMOUNTS.items() if h == head and sub.startswith(pfx)), None)
        if head == "Salaries & Wages" and sub in sal_map:
            line_amt = ind
        ws.append([head, line_amt if line_amt is not None else heads.get(head, ""), sub, meaning, ct, cls, bucket, basis, conf, ind if ind != "" else ""])
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for c in (r[1], r[9]):
            c.number_format = "#,##0.00"
        conf = r[8].value or ""
        if conf.startswith("NEEDS"):
            r[8].fill = PatternFill("solid", fgColor="F8CBAD")
        elif conf.startswith("INFERRED"):
            r[8].fill = PatternFill("solid", fgColor="FFF2CC")
        else:
            r[8].fill = PatternFill("solid", fgColor="E2EFDA")
    for col, w in zip("ABCDEFGHIJ", (30, 16, 52, 40, 30, 16, 30, 50, 22, 18)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B2"

    m = wb.create_sheet("Missing from Arjun")
    m.append(["What is still needed to finish the rerouting at account level"])
    m["A1"].font = bold
    for line in MISSING:
        m.append([line])
    m.column_dimensions["A"].width = 160
    for r in m.iter_rows(min_row=2):
        r[0].alignment = Alignment(wrap_text=True)

    k = wb.create_sheet("Legend")
    for line in ["CONFIRMED: a ruling of the Finance Director or an itemised source (payroll roles, stock movement, depreciation classes).",
                 "INFERRED: composition read from the cash ledger sub-categories; amounts indicative (cash, not accrual).",
                 "NEEDS GL DETAIL: the accrual head mixes items that the Standard places differently; only the account-level lines can split it.",
                 "On-cost: Appendix 1 Cost of Labour = salaries, wages and on-costs, so visas, licences, benefits, end-of-service and staff insurance follow the staff category they belong to.",
                 "DOH bucket column: the fixed bucket shares of scripts/bucket_rules.py (10 Oct 2026 rulings); on-costs take the bucket of their staff category."]:
        k.append([line])
    k.column_dimensions["A"].width = 160
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    print(f"wrote {args.out}: {len(ROWS)} rerouting lines, {len(MISSING)} open items")
    return 0


if __name__ == "__main__":
    sys.exit(main())
