#!/usr/bin/env python3
"""
Month model report: one workbook the Finance Director can check.

  python3 scripts/month_report.py --month 1 --sent input/source/mysmyle_zip/CostSubmission_MF4159_2025_M01_PROD.xml \
      --corrected output/xml_JAN_MODEL/CostSubmission_MF4159_2025_M01_PROD.xml \
      --ledger input/Ledger_by_Cost_Type_FY2025.xlsx \
      --activity input/source/Clinical_2025_one_sheet_Claude_Updated_Cherie.xlsx \
      --log output/xml_JAN_MODEL/rebalance_log.txt --out output/January_Model_Report.xlsx

Sheets: Summary, Cost types (ledger vs corrected file), Buckets (sent vs corrected),
Bucket profiles, Implant claims, Claim examples, Log.
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from costsub_reader import read_submission  # noqa: E402
from costsub_schema import BUCKETS, COST_TYPES, OVERHEAD_COST_TYPES  # noqa: E402
from gl_new import GL_2025  # noqa: E402
from implant_codes import IMPLANT_CODES, IMPLANT_POOLS  # noqa: E402
from rebalance_two_view import MONTH_SHEETS  # noqa: E402

BOLD = Font(bold=True)
HEAD = PatternFill("solid", fgColor="D9E1F2")


def header(ws, row):
    ws.append(row)
    for c in ws[ws.max_row]:
        c.font = BOLD
        c.fill = HEAD


def money(sub):
    out = {}
    for cid, c in sub["claims"].items():
        m = defaultdict(int)
        for e in c["encounters"]:
            for k, v in e["money"].items():
                m[k] += v
        out[cid] = m
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--sent", required=True)
    ap.add_argument("--corrected", required=True)
    ap.add_argument("--ledger", required=True)
    ap.add_argument("--activity", required=True)
    ap.add_argument("--log", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    A, B = money(read_submission(args.sent)), money(read_submission(args.corrected))
    gl = int(GL_2025[args.month] * 1000)
    lw = openpyxl.load_workbook(args.ledger, data_only=True)["Cost types by month"]
    ledger = {r[0]: int(round((r[1 + args.month] or 0) * 1000)) for r in lw.iter_rows(min_row=2, values_only=True) if r[0] in COST_TYPES}
    aw = openpyxl.load_workbook(args.activity, data_only=True, read_only=True)[MONTH_SHEETS[args.month]]
    rows = aw.iter_rows(values_only=True)
    hdr = [str(h) for h in next(rows)]
    hi = {h: i for i, h in enumerate(hdr)}
    codes: dict[str, list[str]] = defaultdict(list)
    empty = 0
    for r in rows:
        if r is None or r[hi["ID"]] in (None, ""):
            empty += 1
            if empty > 50:
                break
            continue
        empty = 0
        codes[str(r[hi["ID"]]).strip()].append(str(r[hi["ProcedureCode"]]).strip().removesuffix(".0"))
    log_lines = Path(args.log).read_text().splitlines()

    def view(m, prefix):
        return sum(v for k, v in m.items() if k.startswith(prefix))

    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "Summary"
    tb = sum(view(m, "CostBucket") for m in B.values())
    tt = sum(view(m, "CostTypeDetails") for m in B.values())
    oh = sum(m.get(f"CostTypeDetails{t}", 0) for m in B.values() for t in OVERHEAD_COST_TYPES)
    ab = sum(view(m, "CostBucket") for m in A.values())
    at = sum(view(m, "CostTypeDetails") for m in A.values())
    match_b = sum(1 for m in B.values() if view(m, "CostBucket") == view(m, "CostTypeDetails"))
    match_a = sum(1 for m in A.values() if view(m, "CostBucket") == view(m, "CostTypeDetails"))
    impl = [c for c in B if set(codes.get(c, [])) & IMPLANT_CODES]
    rows_ = [
        ["Month", MONTH_SHEETS[args.month] + " 2025"],
        ["Ledger (locked)", gl / 1000],
        ["Claims in file", len(B)],
        ["", ""],
        ["", "Sent file", "Corrected file"],
        ["Cost-bucket view total", ab / 1000, tb / 1000],
        ["Cost-type view total", at / 1000, tt / 1000],
        ["Claims where buckets = cost types", f"{match_a} of {len(A)}", f"{match_b} of {len(B)}"],
        ["Ties to the ledger (cost-type view)", "no" if at != gl else "yes", "yes" if tt == gl else "no"],
        ["Overhead cost ratio", f"{sum(m.get(f'CostTypeDetails{t}', 0) for m in A.values() for t in OVERHEAD_COST_TYPES) / at * 100:.1f}%", f"{oh / tt * 100:.1f}%"],
        ["Claims with implant codes", "", len(impl)],
        ["Claims carrying the Prostheses cost type", sum(1 for m in A.values() if m.get("CostTypeDetailsProstheses")), sum(1 for m in B.values() if m.get("CostTypeDetailsProstheses"))],
        ["Prostheses cost type total", sum(m.get("CostTypeDetailsProstheses", 0) for m in A.values()) / 1000, sum(m.get("CostTypeDetailsProstheses", 0) for m in B.values()) / 1000],
        ["Claim set and identity fields", "", "unchanged"],
    ]
    for r in rows_:
        s.append(r)
    s["A1"].font = BOLD
    for c in s[5]:
        c.font = BOLD
    s.column_dimensions["A"].width = 42
    s.column_dimensions["B"].width = 18
    s.column_dimensions["C"].width = 18
    for r in s.iter_rows(min_row=2):
        for c in r[1:]:
            if isinstance(c.value, float):
                c.number_format = "#,##0.000"

    ct = wb.create_sheet("Cost types")
    header(ct, ["Cost type", "Ledger (Arjun)", "Corrected file", "Difference", "Sent file", "Share of month"])
    for t in COST_TYPES:
        lv = ledger.get(t, 0)
        bv = sum(m.get(f"CostTypeDetails{t}", 0) for m in B.values())
        av = sum(m.get(f"CostTypeDetails{t}", 0) for m in A.values())
        if lv or bv or av:
            ct.append([t, lv / 1000, bv / 1000, (bv - lv) / 1000, av / 1000, bv / tt])
    ct.append(["TOTAL", sum(ledger.values()) / 1000, tt / 1000, (tt - sum(ledger.values())) / 1000, at / 1000, 1])
    for c in ct[ct.max_row]:
        c.font = BOLD
    for r in ct.iter_rows(min_row=2):
        for c in r[1:5]:
            c.number_format = "#,##0.000"
        r[5].number_format = "0.0%"
    ct.column_dimensions["A"].width = 18

    bk = wb.create_sheet("Buckets")
    header(bk, ["Bucket", "Sent: direct", "Sent: overheads", "Corrected: direct", "Corrected: overheads", "Corrected: total"])
    for b in BUCKETS:
        ad = sum(m.get(f"CostBucketDirect{b}", 0) for m in A.values())
        ao = sum(m.get(f"CostBucketOverheads{b}", 0) for m in A.values())
        bd = sum(m.get(f"CostBucketDirect{b}", 0) for m in B.values())
        bo = sum(m.get(f"CostBucketOverheads{b}", 0) for m in B.values())
        if ad or ao or bd or bo:
            bk.append([b, ad / 1000, ao / 1000, bd / 1000, bo / 1000, (bd + bo) / 1000])
    bk.append(["TOTAL", ab / 1000 - 0, 0, sum(m.get(f"CostBucketDirect{b}", 0) for m in B.values() for b in BUCKETS) / 1000,
               sum(m.get(f"CostBucketOverheads{b}", 0) for m in B.values() for b in BUCKETS) / 1000, tb / 1000])
    for c in bk[bk.max_row]:
        c.font = BOLD
    for r in bk.iter_rows(min_row=2):
        for c in r[1:]:
            c.number_format = "#,##0.000"
    bk.column_dimensions["A"].width = 14

    pf = wb.create_sheet("Bucket profiles")
    header(pf, ["Bucket", "Cost-type profile fitted for the month (share of the bucket)"])
    for ln in log_lines:
        if "profile" in ln:
            b = ln.split("bucket ")[1].split(" profile")[0]
            pf.append([b, ln.split("profile ")[1]])
    pf.column_dimensions["A"].width = 14
    pf.column_dimensions["B"].width = 150

    ic = wb.create_sheet("Implant claims")
    header(ic, ["Claim", "Implant codes billed", "Fixture units", "Abutment units", "Bone-graft units", "Claim cost",
                "Prostheses (corrected)", "Prosthesis bucket (corrected)", "Dental bucket sent", "Dental bucket corrected", "Note"])
    caps = {ln.split("claim ")[1].split(" ")[0]: ln.split(": ", 2)[-1] for ln in log_lines if "capped" in ln}
    for cid in sorted(impl, key=lambda c: -B[c].get("CostTypeDetailsProstheses", 0)):
        cs = [c for c in codes[cid] if c in IMPLANT_CODES]
        u = {p: sum(1 for c in cs if c in IMPLANT_POOLS[p]) for p in IMPLANT_POOLS}
        ic.append([cid, ", ".join(sorted(cs)), u["fixtures"], u["abutments"], u["bone_graft"],
                   view(B[cid], "CostTypeDetails") / 1000, B[cid].get("CostTypeDetailsProstheses", 0) / 1000,
                   (B[cid].get("CostBucketDirectProsthesis", 0)) / 1000,
                   (A[cid].get("CostBucketDirectDental", 0) + A[cid].get("CostBucketOverheadsDental", 0) + view(A[cid], "CostTypeDetails")) / 1000,
                   (B[cid].get("CostBucketDirectDental", 0) + B[cid].get("CostBucketOverheadsDental", 0)) / 1000,
                   caps.get(cid, "")])
    for r in ic.iter_rows(min_row=2):
        for c in r[5:10]:
            c.number_format = "#,##0.000"
    ic.column_dimensions["A"].width = 22
    ic.column_dimensions["B"].width = 40
    ic.column_dimensions["K"].width = 70

    ex = wb.create_sheet("Claim examples")
    fix = [c for c in impl if "79933" in codes.get(c, [])]
    picks = [(max(fix, key=lambda c: view(B[c], "CostTypeDetails")), "Implant installation claim")] if fix else []
    picks.append(("1022MF4159-a9c9e9", "Caries claim, no implant codes"))
    for cid, label in picks:
        if cid not in B:
            continue
        ex.append([label, cid, "codes: " + ", ".join(codes.get(cid, []))])
        ex[ex.max_row][0].font = BOLD
        header(ex, ["Cell", "Sent", "Corrected"])
        for grp in ("CostBucketDirect", "CostBucketOverheads", "CostTypeDetails"):
            sa = sb = 0
            for k in sorted(set(A[cid]) | set(B[cid])):
                if k.startswith(grp) and (A[cid].get(k) or B[cid].get(k)):
                    ex.append([k, A[cid].get(k, 0) / 1000, B[cid].get(k, 0) / 1000])
                    sa += A[cid].get(k, 0)
                    sb += B[cid].get(k, 0)
            ex.append([grp + " TOTAL", sa / 1000, sb / 1000])
            ex[ex.max_row][0].font = BOLD
        ex.append([])
    for r in ex.iter_rows(min_row=1):
        for c in r[1:3]:
            if isinstance(c.value, float):
                c.number_format = "#,##0.000"
    ex.column_dimensions["A"].width = 36
    ex.column_dimensions["B"].width = 14
    ex.column_dimensions["C"].width = 14

    lg = wb.create_sheet("Log")
    for ln in log_lines:
        lg.append([ln])
    lg.column_dimensions["A"].width = 160
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    print(f"wrote {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
