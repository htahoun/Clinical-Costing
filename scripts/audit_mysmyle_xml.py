#!/usr/bin/env python3
"""
The standard two-file Cost.Submission audit (docs/CLINICAL_COSTING.md §7).

  python3 scripts/audit_mysmyle_xml.py <my_xml> <their_xml> [--xsd path] [--tol 0.0005]

File sizes; XSD both; header fields; claim-set equality both ways; every
identity/meta field per claim; per-cell money diffs > tolerance with counts and
nets per field; zero-total and negative claims; largest cell diffs. Verdict in
plain words, money balance first. Same standard every run.
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from costsub_reader import claim_total, file_total, read_submission, xsd_validate  # noqa: E402

META_FIELDS = ["FacilityID", "ID", "Type", "StartDateTime", "EndDateTime", "StartType", "EndType",
               "ActivityClinicianID", "ServiceCode", "Specialty", "SubSpecialty", "DRGCode",
               "OPAttendanceType", "TheatreAttendanceFlag", "TriageCategory", "CriticalCareHours",
               "VentilationHours"]


def flat_money(claim: dict) -> dict:
    m: dict = defaultdict(int)
    for e in claim["encounters"]:
        for k, v in e["money"].items():
            m[k] += v
    return m


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mine")
    ap.add_argument("theirs")
    ap.add_argument("--xsd", default=str(Path(__file__).resolve().parents[1] / "docs/shafafiya/CostSubmission_repaired.xsd"))
    ap.add_argument("--tol", type=float, default=0.0005, help="AED tolerance per cell")
    args = ap.parse_args(argv)
    tol = int(round(args.tol * 1000))

    A, B = read_submission(args.mine), read_submission(args.theirs)
    la, lb = Path(args.mine).name, Path(args.theirs).name
    print(f"A = {la}  ({A['size']:,} bytes, {len(A['claims'])} claims)")
    print(f"B = {lb}  ({B['size']:,} bytes, {len(B['claims'])} claims)")
    for lab, s, p in (("A", A, args.mine), ("B", B, args.theirs)):
        r, msgs = xsd_validate(p, args.xsd)
        print(f"XSD {lab}: {r}" + (f" — {msgs[0]}" if msgs else ""))
    print("\nHeader")
    for k in sorted(set(A["header"]) | set(B["header"])):
        a, b = A["header"].get(k), B["header"].get(k)
        print(f"  {k:<18}{str(a):<28}{str(b):<28}{'' if a == b else '<-- differs'}")

    sa, sb = set(A["claims"]), set(B["claims"])
    only_a, only_b, both = sorted(sa - sb), sorted(sb - sa), sorted(sa & sb)
    print(f"\nClaim sets: common {len(both)}; only in A {len(only_a)}; only in B {len(only_b)}")
    for lab, lst, src in (("only A", only_a, A), ("only B", only_b, B)):
        for cid in lst[:15]:
            print(f"  {lab}: {cid}  total {claim_total(src['claims'][cid]) / 1000:,.3f}")
        if len(lst) > 15:
            print(f"  {lab}: ... {len(lst) - 15} more")

    meta_diffs = Counter()
    meta_examples = {}
    for cid in both:
        ca, cb = A["claims"][cid], B["claims"][cid]
        for k in ("ProviderID", "PatientID"):
            if ca.get(k) != cb.get(k):
                meta_diffs[k] += 1
                meta_examples.setdefault(k, (cid, ca.get(k), cb.get(k)))
        ea = ca["encounters"][0] if ca["encounters"] else {"fields": {}, "diagnoses": [], "procedures": []}
        eb = cb["encounters"][0] if cb["encounters"] else {"fields": {}, "diagnoses": [], "procedures": []}
        if len(ca["encounters"]) != len(cb["encounters"]):
            meta_diffs["encounter count"] += 1
        for k in META_FIELDS:
            if ea["fields"].get(k, "") != eb["fields"].get(k, ""):
                meta_diffs[k] += 1
                meta_examples.setdefault(k, (cid, ea["fields"].get(k, ""), eb["fields"].get(k, "")))
        if ea["diagnoses"] != eb["diagnoses"]:
            meta_diffs["Diagnosis"] += 1
            meta_examples.setdefault("Diagnosis", (cid, ea["diagnoses"], eb["diagnoses"]))
        if ea["procedures"] != eb["procedures"]:
            meta_diffs["Procedure"] += 1
            meta_examples.setdefault("Procedure", (cid, ea["procedures"], eb["procedures"]))
    print(f"\nIdentity/meta differences on common claims: {sum(meta_diffs.values())}")
    for k, n in meta_diffs.most_common():
        cid, a, b = meta_examples.get(k, ("", "", ""))
        print(f"  {k:<22}{n:>6}   e.g. {cid}: {a!r} vs {b!r}")

    cell_counts, cell_nets = Counter(), Counter()
    largest = []
    tot_a = tot_b = 0
    for cid in both:
        ma, mb = flat_money(A["claims"][cid]), flat_money(B["claims"][cid])
        for k in sorted(set(ma) | set(mb)):
            d = ma.get(k, 0) - mb.get(k, 0)
            if abs(d) > tol:
                cell_counts[k] += 1
                cell_nets[k] += d
                largest.append((abs(d), cid, k, ma.get(k, 0), mb.get(k, 0)))
        tot_a += claim_total(A["claims"][cid])
        tot_b += claim_total(B["claims"][cid])
    print(f"\nMoney on common claims: A {tot_a / 1000:,.3f}  B {tot_b / 1000:,.3f}  net {(tot_a - tot_b) / 1000:+,.3f}")
    print(f"Whole file totals:      A {file_total(A) / 1000:,.3f}  B {file_total(B) / 1000:,.3f}  "
          f"net {(file_total(A) - file_total(B)) / 1000:+,.3f}")
    print(f"Cells differing by > {args.tol} AED: {sum(cell_counts.values())}")
    for k, n in cell_counts.most_common():
        print(f"  {k:<34}{n:>7}   net {cell_nets[k] / 1000:+,.3f}")
    largest.sort(reverse=True)
    if largest:
        print("Largest cell differences:")
        for d, cid, k, a, b in largest[:10]:
            print(f"  {cid:<24}{k:<34}A {a / 1000:>12,.3f}  B {b / 1000:>12,.3f}  diff {(a - b) / 1000:+,.3f}")

    for lab, s in (("A", A), ("B", B)):
        z = [c for c, v in s["claims"].items() if claim_total(v) == 0]
        n = [c for c, v in s["claims"].items() if any(x < 0 for e in v["encounters"] for x in e["money"].values())]
        print(f"{lab}: zero-total claims {len(z)}; negative claims {len(n)}" + (f"  e.g. {z[:3]}" if z else ""))

    print("\nVERDICT")
    net = file_total(A) - file_total(B)
    print(f"  Money: files {'balance to the fils' if net == 0 else f'differ by {net / 1000:+,.3f} AED'} overall; "
          f"{'claim sets identical' if not only_a and not only_b else f'{len(only_a)} claims only in A, {len(only_b)} only in B'}; "
          f"{sum(meta_diffs.values())} identity differences; {sum(cell_counts.values())} cells differ beyond tolerance.")
    mirror = [(k, cell_nets[k]) for k in cell_counts]
    pos = sum(v for _, v in mirror if v > 0)
    negs = -sum(v for _, v in mirror if v < 0)
    if mirror and pos == negs:
        print("  The cell differences mirror each other exactly (money moved between buckets, none created or lost).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
