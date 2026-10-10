#!/usr/bin/env python3
"""
PROPOSED ledger-by-cost-type map (built without the expense lines — an estimate
for Arjun to confirm or replace, 10 Oct 2026).

  python3 scripts/propose_ledger_map.py --template input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx \
      --out input/Ledger_by_Cost_Type_FY2025_PROPOSED.xlsx

Logic (see the Assumptions sheet it writes):
  * Start from the clinic's own cost structure as implied by the costed files
    (Cherie's allocation table scaled to the ledger).
  * Reclassify by the Standard: front-office/admin/management pay, premises
    (rent, service charges, utilities, maintenance, security, insurance),
    marketing, IT and professional fees are OVERHEAD cost-centre expenses
    (OHF); their depreciation is OHDepreciation; group recharges are OHC.
    Because the rejected files carried only 4.2 % overheads, these items were
    smeared across the direct columns by the scale factor; they are pulled
    back out by scaling the direct block down.
  * Overheads are treated as FIXED monthly amounts (rent and admin salaries do
    not move with activity); the direct block takes the rest of each month's
    ledger in the direct mix.
  * Each month ties to the locked ledger to the fils (SWDoc absorbs residuals).
"""
from __future__ import annotations

import argparse
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from costsub_schema import COST_TYPES, OVERHEAD_COST_TYPES  # noqa: E402
from gl_new import GL_2025, GL_2025_TOTAL  # noqa: E402

# Proposed FY2025 mix, % of total expenditure. Sums to 100.
MIX = {
    "SWDoc": Decimal("25.5"),        # dentists / oral surgeon / orthodontist pay (implied 32.0 before pulling overheads out)
    "SWNurs": Decimal("9.0"),        # dental nurses (implied 11.4)
    "SWAllied": Decimal("3.2"),      # hygienists, radiographer, lab technician (implied 4.4)
    "SWNonClin": Decimal("1.0"),     # clinic clerks / chairside non-clinical staff recorded in the dental cost centre
    "Laboratory": Decimal("7.3"),    # outsourced dental laboratory invoices = the costed Laboratory bucket (469k)
    "Imaging": Decimal("2.6"),       # imaging sensors, films, outsourced CBCT = the costed Imaging bucket (169k)
    "Pharmacy": Decimal("3.3"),      # local anaesthetics, drugs (Anaesthesia bucket 264k less the dentist's time)
    "Prostheses": Decimal("3.4"),    # crowns, dentures, implants = the costed Prosthesis bucket (217k)
    "MS": Decimal("17.2"),           # dental and surgical consumables (implied 18.1 + 5.3 before pulling overheads out)
    "Hotel": Decimal("0.8"),         # cleaning, linen, sterilisation wraps
    "GS": Decimal("1.0"),            # other goods and services in the clinic rooms (clinician licences, waste, small equipment)
    "Depreciation": Decimal("1.7"),  # dental chairs, imaging units (implied 2.0, part moved to OHDepreciation)
    "OHF": Decimal("21.5"),          # premises rent/service charges/utilities/maintenance/security/insurance (~12 %),
                                     # reception/admin/management pay (~6 %), marketing/IT/insurance/professional fees (~3.5 %)
    "OHDepreciation": Decimal("1.3"),  # fit-out, furniture, IT depreciation
    "OHC": Decimal("1.2"),           # group / shared-service recharges (implied 1.1; 0 if there is no corporate office)
    "NABMs": Decimal("0"),
}
assert sum(MIX.values()) == 100, sum(MIX.values())

ASSUMPTIONS = [
    "This sheet is a PROPOSAL built without the FY2025 expense lines. Arjun must confirm or replace every figure from the general ledger.",
    "",
    "A1. Overheads (OHF + OHDepreciation + OHC) = 24.0 % of FY2025 expenditure = AED {oh:,.2f}, treated as a FIXED amount per month ({ohm:,.2f}).",
    "    Rationale: rent and service charges for the premises (~11-12 %), reception/admin/management salaries (~6 %), utilities, maintenance,",
    "    security, insurance, marketing, IT and professional fees (~4-5 %). The rejected files carried 4.2 %, which DOH flagged as unusually low.",
    "    Overheads are fixed per month, so busy months show a lower ratio; 24 % keeps every month at or above the DOH floor of 20 %.",
    "A2. Front office, administration and management staff are an OVERHEAD cost centre (Standard App. 1: SWNonClin is only for non-clinical staff",
    "    recorded in DIRECT cost centres). Their pay is moved from SWNonClin to OHF; 1.0 % stays as SWNonClin for clinic clerks/chairside non-clinical staff.",
    "A3. Premises costs are an OVERHEAD (facility) cost centre allocated to the Dental bucket by floor space, so they are OHF, not GS.",
    "A3b. Laboratory, Prostheses and Imaging cost types are set equal to the clinic's own costed buckets (outsourced lab 469k, prostheses 217k, imaging 169k).",
    "A4. Depreciation: dental chairs and imaging units are DIRECT (1.5 %); fit-out, furniture and IT are OVERHEAD (OHDepreciation 1.2 %).",
    "A5. OHC 1.2 % is kept close to the implied 1.1 % by the costed files (group / shared-service recharges). Set to 0 if Vision Dental has no corporate office.",
    "A6. The direct block (76.0 %) keeps the proportions of the clinic's own allocation table, scaled down uniformly: the scale factor in the",
    "    original build spread the overhead money across the direct columns, so pulling it out proportionally reverses that exactly.",
    "A7. Same direct mix every month (no monthly expense detail available). Overheads fixed; direct types take the rest of each month's ledger.",
    "A8. Every month reconciles to the locked ledger to the fils (SWDoc absorbs rounding). FY2025 total {gl:,.2f}.",
    "A9. NABMs = 0 (no DOH non-activity mandate funding for a private clinic). Accrual basis assumed as in the locked ledger.",
    "",
    "What this proposal CANNOT establish without the ledger lines: the actual rent and service charge; actual admin/management payroll;",
    "whether a corporate office exists (OHC); month-to-month movements (bonuses, annual licence fees, year-end depreciation);",
    "third-party expenses not in the GL (Standard 3.1.2). These change the overhead ratio and the cost-type split, not the claim totals.",
]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", default="input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx")
    ap.add_argument("--out", default="input/Ledger_by_Cost_Type_FY2025_PROPOSED.xlsx")
    args = ap.parse_args(argv)

    oh_share = sum(MIX[t] for t in OVERHEAD_COST_TYPES)
    oh_year = (GL_2025_TOTAL * oh_share / 100).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    oh_month = (oh_year / 12).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    direct_types = [t for t in COST_TYPES if t not in OVERHEAD_COST_TYPES]
    direct_share = sum(MIX[t] for t in direct_types)

    figures: dict[int, dict[str, Decimal]] = {}
    for m in range(1, 13):
        gl = GL_2025[m]
        row: dict[str, Decimal] = {}
        for t in OVERHEAD_COST_TYPES:
            row[t] = (oh_month * MIX[t] / oh_share).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        rest = gl - sum(row.values())
        for t in direct_types:
            row[t] = (rest * MIX[t] / direct_share).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        row["SWDoc"] += gl - sum(row.values())          # tie to the fils
        assert sum(row.values()) == gl
        figures[m] = row

    wb = openpyxl.load_workbook(args.template)
    ws = wb["Cost types by month"]
    for r in ws.iter_rows(min_row=2):
        t = r[0].value
        if t in COST_TYPES:
            for m in range(1, 13):
                r[1 + m].value = float(figures[m][t])
    ws.title = "Cost types by month"
    a = wb.create_sheet("Assumptions", 1)
    a.append(["PROPOSED mapping — assumptions for Arjun to confirm or replace"])
    a["A1"].font = Font(bold=True)
    for line in ASSUMPTIONS:
        a.append([line.format(oh=oh_year, ohm=oh_month, gl=GL_2025_TOTAL)])
    a.append([])
    a.append(["Proposed mix, % of total", "Share %", "FY2025 AED", "Basis"])
    for c in a[a.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
    basis = {
        "SWDoc": "dentists' pay; implied 32.0 % before overheads were pulled out",
        "SWNurs": "dental nurses; implied 11.4 %", "SWAllied": "hygienists, radiographer, lab technician; implied 4.4 %",
        "SWNonClin": "clinic clerks / chairside non-clinical staff in the dental cost centre",
        "Laboratory": "outsourced dental laboratory invoices; implied 7.3 %", "Imaging": "imaging supplies; implied 2.6 %",
        "Pharmacy": "anaesthetics and drugs; implied 4.1 %", "Prostheses": "crowns, dentures, implants; implied 3.4 %",
        "MS": "dental and surgical consumables; implied 18.1 % + 5.3 %", "Hotel": "cleaning, linen, sterilisation wraps",
        "GS": "other goods and services in the clinic rooms", "Depreciation": "dental chairs, imaging units; implied 2.0 %",
        "OHF": "premises, admin/management pay, marketing, IT, insurance, professional fees",
        "OHDepreciation": "fit-out, furniture, IT depreciation", "OHC": "group recharges; close to implied (1.1 %)", "NABMs": "none",
    }
    for t in COST_TYPES:
        a.append([t, float(MIX[t]), float((GL_2025_TOTAL * MIX[t] / 100).quantize(Decimal("0.01"))), basis[t]])
    a.append(["Overhead ratio", float(oh_share), float(oh_year), "DOH band: 20 % to under 50 %"])
    a.column_dimensions["A"].width = 26
    a.column_dimensions["B"].width = 10
    a.column_dimensions["C"].width = 16
    a.column_dimensions["D"].width = 90
    wb.save(args.out)
    print(f"wrote {args.out}: overheads {oh_share}% = {oh_year:,.2f} ({oh_month:,.2f}/month); direct {direct_share}%")
    for m in range(1, 13):
        print(f"  {m:02d} ledger {GL_2025[m]:>12,.2f}  OHF {figures[m]['OHF']:>10,.2f}  SWDoc {figures[m]['SWDoc']:>10,.2f}  MS {figures[m]['MS']:>10,.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
