#!/usr/bin/env python3
"""
Pre-flight validator: every hard portal/schema rule from docs/CLINICAL_COSTING.md §5,
checked on any Cost.Submission XML (ours or a platform's) before it goes to Shafafiya.

  python3 scripts/validate_xml.py output/xml_final/*.xml [--xsd path] [--year 2025]
      [--expect-disposition PRODUCTION] [--strict-totals]

Checks: XSD (if present); ReceiverID HAAD; Sender/Provider/Facility = MF4159;
DispositionFlag; RecordCount = claims; date format dd/mm/yyyy HH:MM; all dates in
one month and in the file's month; money text exactly 3 dp; no negatives; no
zero-total claims; one <Claim> per ID; principal diagnosis only; Procedure empty
unless TheatreAttendanceFlag=1; CostBucketOverheads all zero; TotalCost = sum of
cells; bucket-vs-type cross-foot report; file ≤ 5 MB; month total vs ledger.
Exit 1 on any FAIL.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from costsub_reader import (DATE_RE, MONEY_RE, claim_total, encounter_total, file_total,  # noqa: E402
                            read_submission, xsd_validate)
from costsub_schema import (COST_TYPE_ELEMENTS, DIRECT_ELEMENTS, DISPOSITIONS, FACILITY_ID,  # noqa: E402
                            MAX_FILE_BYTES, OVERHEAD_ELEMENTS, RECEIVER_ID)
from gl_new import GL  # noqa: E402


def validate(path: str, xsd: str | None, year: int, expect_disp: str | None, strict_totals: bool) -> list[tuple[str, str]]:
    """Return list of (level, message); level in FAIL / WARN / INFO."""
    out: list[tuple[str, str]] = []
    F = lambda m: out.append(("FAIL", m))
    W = lambda m: out.append(("WARN", m))
    I = lambda m: out.append(("INFO", m))
    sub = read_submission(path)
    for e in sub["errors"]:
        F(e)
    if sub["root"] != "Cost.Submission":
        F(f"root element is <{sub['root']}>, expected <Cost.Submission>")
    if sub["size"] > MAX_FILE_BYTES:
        F(f"file is {sub['size']:,} bytes > 5 MB; split the month")
    xsd_res, msgs = xsd_validate(path, xsd)
    (F if xsd_res == "FAIL" else W if xsd_res == "NO XSD" else I)(f"XSD: {xsd_res}" + ("; " + msgs[0] if msgs else ""))
    for m in msgs[1:6]:
        I(f"XSD: {m}")

    h = sub["header"]
    if h.get("SenderID") != FACILITY_ID:
        F(f"SenderID {h.get('SenderID')!r} != {FACILITY_ID}")
    if h.get("ReceiverID") != RECEIVER_ID:
        F(f"ReceiverID {h.get('ReceiverID')!r} != {RECEIVER_ID}")
    if h.get("DispositionFlag") not in DISPOSITIONS:
        F(f"DispositionFlag {h.get('DispositionFlag')!r} not in {DISPOSITIONS}")
    elif expect_disp and h.get("DispositionFlag") != expect_disp:
        F(f"DispositionFlag is {h.get('DispositionFlag')}, expected {expect_disp}")
    if not DATE_RE.match(h.get("TransactionDate", "")):
        F(f"TransactionDate {h.get('TransactionDate')!r} not dd/mm/yyyy HH:MM")
    if str(len(sub["claims"])) != h.get("RecordCount"):
        F(f"RecordCount {h.get('RecordCount')} != {len(sub['claims'])} claims in file")

    months = Counter()
    enc_ids = Counter()
    zero, neg = [], []
    bucket_eq_type = type_eq_total = bucket_eq_total = 0
    for cid, c in sub["claims"].items():
        if c.get("ProviderID") != FACILITY_ID:
            F(f"{cid}: ProviderID {c.get('ProviderID')!r}")
        if not c.get("PatientID"):
            F(f"{cid}: blank PatientID")
        if claim_total(c) == 0:
            zero.append(cid)
        for e in c["encounters"]:
            f = e["fields"]
            if f.get("FacilityID") != FACILITY_ID:
                F(f"{cid}: FacilityID {f.get('FacilityID')!r}")
            if not f.get("ID"):
                F(f"{cid}: blank Encounter ID")
            enc_ids[f.get("ID")] += 1
            if not f.get("ActivityClinicianID"):
                F(f"{cid}: blank ActivityClinicianID")
            for k in ("StartDateTime", "EndDateTime"):
                v = f.get(k, "")
                if not DATE_RE.match(v):
                    F(f"{cid}: {k} {v!r} not dd/mm/yyyy HH:MM")
                else:
                    d = datetime.strptime(v, "%d/%m/%Y %H:%M")
                    months[(d.year, d.month)] += 1
                    if d.year != year:
                        F(f"{cid}: {k} {v} outside FY{year}")
            if f.get("StartDateTime", "") and f.get("EndDateTime", "") and \
               DATE_RE.match(f["StartDateTime"]) and DATE_RE.match(f["EndDateTime"]) and \
               datetime.strptime(f["EndDateTime"], "%d/%m/%Y %H:%M") < datetime.strptime(f["StartDateTime"], "%d/%m/%Y %H:%M"):
                F(f"{cid}: EndDateTime before StartDateTime")
            types = [t for t, _ in e["diagnoses"]]
            if not e["diagnoses"] or types[0] != "Principal":
                F(f"{cid}: first diagnosis is not Principal ({types})")
            if any(t != "Principal" for t in types):
                F(f"{cid}: non-principal diagnosis present (locked: principal only)")
            if any(not code for _, code in e["diagnoses"]):
                F(f"{cid}: blank diagnosis code")
            flag = f.get("TheatreAttendanceFlag", "0")
            procs = [p for p in e["procedures"] if p[0]]
            if flag != "1" and procs:
                F(f"{cid}: Procedure given but TheatreAttendanceFlag={flag}")
            if flag == "1" and not procs:
                W(f"{cid}: TheatreAttendanceFlag=1 but no procedure code")
            for k, txt in e["money_text"].items():
                if not MONEY_RE.match(txt):
                    F(f"{cid}: {k} = {txt!r} is not 3 dp")
            if any(v < 0 for v in e["money"].values()):
                neg.append(cid)
            if any(e["money"].get(k, 0) for k in OVERHEAD_ELEMENTS):
                F(f"{cid}: CostBucketOverheads not all zero (overheads live in CostTypeDetails)")
            tot = e["money"].get("TotalCost")
            if tot is not None and tot != encounter_total(e):
                F(f"{cid}: TotalCost {tot/1000:.3f} != sum of cells {encounter_total(e)/1000:.3f}")
            b = sum(e["money"].get(k, 0) for k in DIRECT_ELEMENTS + OVERHEAD_ELEMENTS)
            t = sum(e["money"].get(k, 0) for k in COST_TYPE_ELEMENTS)
            bucket_eq_type += (b == t)
            type_eq_total += (t == (tot if tot is not None else encounter_total(e)))
            bucket_eq_total += (b == (tot if tot is not None else encounter_total(e)))
    for cid in zero:
        F(f"{cid}: zero-total claim (DOH blocks these; exclude)")
    for cid in neg:
        F(f"{cid}: negative cost cell")
    for eid, n in enc_ids.items():
        if n > 1:
            W(f"Encounter ID {eid} appears {n} times in this file")
    if len(months) > 1:
        W(f"file spans months {sorted(months)} — expected one month per file")
    n = len(sub["claims"])
    I(f"cross-foot: buckets==types in {bucket_eq_type}/{n}; types==total in {type_eq_total}/{n}; "
      f"buckets==total in {bucket_eq_total}/{n}  (our method: buckets + types = total)")
    if strict_totals and (bucket_eq_type != n or type_eq_total != n):
        F("strict cross-foot requested: buckets, cost types and TotalCost do not each sum to the total")
    total = file_total(sub)
    if len(months) == 1:
        (y, m), _ = next(iter(months.items()))
        gl = GL.get(y, {}).get(m)
        if gl is not None:
            diff = total - int(gl * 1000)
            (I if diff == 0 else F)(f"month total {total/1000:,.3f} vs ledger {gl:,.3f}: {'ties to the fils' if diff == 0 else f'{diff/1000:+.3f}'}")
    I(f"{n} claims, {sub['size']:,} bytes, total {total/1000:,.3f}, disposition {h.get('DispositionFlag')}")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--xsd", default=str(Path(__file__).resolve().parents[1] / "docs/shafafiya/CostSubmission_repaired.xsd"))
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--expect-disposition", choices=DISPOSITIONS)
    ap.add_argument("--strict-totals", action="store_true",
                    help="also FAIL unless buckets, types and TotalCost each equal the total (DOH cross-foot hypothesis)")
    args = ap.parse_args(argv)
    rc = 0
    for f in args.files:
        res = validate(f, args.xsd, args.year, args.expect_disposition, args.strict_totals)
        fails = [m for lvl, m in res if lvl == "FAIL"]
        warns = [m for lvl, m in res if lvl == "WARN"]
        print(f"\n== {Path(f).name}: {'FAIL' if fails else 'PASS'} ({len(fails)} fail, {len(warns)} warn)")
        shown = Counter()
        for lvl, m in res:
            key = (lvl, re.sub(r"^\S+: ", "", m)[:40])
            shown[key] += 1
            if shown[key] <= 8:
                print(f"  {lvl:<5}{m}")
        for (lvl, k), c in shown.items():
            if c > 8:
                print(f"  {lvl:<5}... {c - 8} more like '{k}'")
        rc |= bool(fails)
    return rc


if __name__ == "__main__":
    sys.exit(main())
