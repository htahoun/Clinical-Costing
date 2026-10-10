#!/usr/bin/env python3
"""
Ledger-calibrated two-view engine (DOH SY02 corrective action, rulings of 10 Oct 2026).

  python3 scripts/rebalance_two_view.py --indir input/source/mysmyle_zip \
      --ledger input/Ledger_by_Cost_Type_FY2025.xlsx --outdir output/xml_resubmission [--zip]

Two anchors, one fit, per month:
  * Bucket anchor (WHERE): each claim's bucket amounts as costed — the seven
    bucket columns as sent, plus the ten cost-type columns placed in the Dental
    bucket (the clinic's one direct cost centre). Claim totals and the monthly
    tie to the ledger are unchanged.
  * Cost-type anchor (WHAT): the ledger mapped to the sixteen DOH cost types
    (Standard v2.1 Appendix 1), by month, from input/Ledger_by_Cost_Type_FY2025.xlsx.
  * Fit: a bucket x cost-type matrix is balanced to both margins (iterative
    proportional fitting, RAS) from a nature seed (SEED). Each claim's cost-type
    cells = its bucket amounts x the fitted bucket profiles, so on every claim
    cost types = cost buckets, and over the month cost types = ledger by type.
  * Overheads: each bucket's OHF/OHDepreciation/OHC share is written to
    CostBucketOverheads, the rest to CostBucketDirect (Standard: allocated
    overhead sits in the bucket's overhead component).

Exact integer-fils arithmetic; per claim both views are forced equal; per month
every cost-type column is forced to the ledger figure to the fils. Everything
else in the files is preserved verbatim; TransactionDate is refreshed.
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from datetime import datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from lxml import etree  # noqa: E402

from build2_lib import to_fils  # noqa: E402
from costsub_reader import fils  # noqa: E402
from costsub_schema import (BUCKETS, COST_TYPES, DATE_FMT, OVERHEAD_COST_TYPES,  # noqa: E402
                            TZ_OFFSET_HOURS, fmt_money_fils)
from gl_new import GL  # noqa: E402

HOME_BUCKET = "Dental"
DIRECT_TYPES = [t for t in COST_TYPES if t not in OVERHEAD_COST_TYPES]

# Nature seed: which cost types each bucket may contain, with prior weights.
# Zero (absent) means "never". The RAS fit rescales these to both margins.
SEED = {
    # Dental is the clinic's one direct cost centre: it may hold every cost type (small prior weights for
    # the types that belong mainly to another bucket), so the fit always has a flexible absorber.
    "Dental": {"SWDoc": 4, "SWNurs": 2, "SWAllied": 1, "SWNonClin": 1, "MS": 2, "Hotel": 0.2, "GS": 0.5,
               "Depreciation": 1, "Pharmacy": 0.2, "Laboratory": 0.1, "Imaging": 0.1, "Prostheses": 0.1,
               "OHF": 1, "OHDepreciation": 0.5, "OHC": 0.5},
    # Ruling 10 Oct 2026: overhead cost types sit in the Dental bucket only (CostBucketOverheads/Dental).
    # Pass-through purchases: mainly their own cost type, with a small weight on supplies/technician time
    # so the fit can absorb a ledger figure that differs from the costed bucket.
    "Laboratory": {"Laboratory": 1, "MS": 0.05, "SWAllied": 0.05},
    "Prosthesis": {"Prostheses": 1, "MS": 0.05},
    "Imaging": {"Imaging": 2, "SWAllied": 0.5, "Depreciation": 0.5, "MS": 0.2},
    "Anaesthesia": {"Pharmacy": 2, "SWDoc": 0.5, "MS": 0.3},
    "OP": {"SWNonClin": 1, "GS": 0.2, "MS": 0.1},
    "Physician": {"SWDoc": 1, "MS": 0.05},
    "Pharmacy": {"Pharmacy": 1, "MS": 0.05},
    "Allied": {"SWAllied": 1, "MS": 0.05},
    "Other": {"GS": 1, "MS": 0.5},
}
TOL = Decimal("0.0000001")


def read_ledger(path: Path) -> dict[int, dict[str, int]]:
    """{month: {cost type: fils}} from the template; validates each month to the locked ledger."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Cost types by month"]
    out: dict[int, dict[str, int]] = {m: {} for m in range(1, 13)}
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = row[0]
        if t not in COST_TYPES:
            continue
        for m in range(1, 13):
            out[m][t] = to_fils(row[1 + m]) if row[1 + m] not in (None, "") else 0
    problems = []
    for m in range(1, 13):
        if len(out[m]) != len(COST_TYPES):
            problems.append(f"month {m}: missing cost types {sorted(set(COST_TYPES) - set(out[m]))}")
        gl = int(GL[2025][m] * 1000)
        s = sum(out[m].values())
        if s != gl:
            problems.append(f"month {m}: cost types sum {s/1000:,.3f} != ledger {gl/1000:,.3f} ({(s-gl)/1000:+.3f})")
        if any(v < 0 for t, v in out[m].items() if t != "NABMs"):
            problems.append(f"month {m}: negative cost type")
    if problems:
        raise SystemExit("ledger workbook not usable:\n  " + "\n  ".join(problems))
    return out


def ras(rows: dict[str, int], cols: dict[str, int], seed: dict, max_iter=20000) -> dict[str, dict[str, Decimal]]:
    """Balance a seed matrix to row totals (buckets) and column totals (cost types). Decimal."""
    R = {b: Decimal(v) for b, v in rows.items() if v}
    C = {t: Decimal(v) for t, v in cols.items() if v}
    M = {b: {t: Decimal(str(seed.get(b, {}).get(t, 0))) for t in C} for b in R}
    for b in R:
        if sum(M[b].values()) == 0:
            raise SystemExit(f"bucket {b} has no admissible cost type in SEED with a non-zero ledger total")
    for t in C:
        if sum(M[b][t] for b in R) == 0:
            raise SystemExit(f"cost type {t} ({C[t]/1000:,.3f} in the ledger) has no admissible bucket in SEED")
    for _ in range(max_iter):
        for b in R:
            s = sum(M[b].values())
            f = R[b] / s
            for t in C:
                M[b][t] *= f
        for t in C:
            s = sum(M[b][t] for b in R)
            f = C[t] / s
            for b in R:
                M[b][t] *= f
        if all(abs(sum(M[b].values()) - R[b]) <= R[b] * TOL for b in R) and \
           all(abs(sum(M[b][t] for b in R) - C[t]) <= C[t] * TOL for t in C):
            break
    else:
        worst = max(max(abs(sum(M[b].values()) - R[b]) / R[b] for b in R),
                    max(abs(sum(M[b][t] for b in R) - C[t]) / C[t] for t in C))
        if worst > Decimal("0.000001"):
            raise SystemExit(f"RAS did not converge (worst relative error {worst}); the seed is incompatible with the two margins")
    return M


def round_to_total(exact: dict[str, Decimal], total: int) -> dict[str, int]:
    """Round a vector to fils so that it sums exactly to `total` (largest remainders)."""
    base = {k: int(v.to_integral_value(rounding="ROUND_FLOOR")) for k, v in exact.items()}
    rem = {k: exact[k] - base[k] for k in exact}
    diff = total - sum(base.values())
    for k in sorted(rem, key=lambda k: rem[k], reverse=True)[:max(diff, 0)]:
        base[k] += 1
    if diff < 0:  # cannot happen with floor, kept for safety
        for k in sorted(rem, key=lambda k: rem[k])[:-diff]:
            base[k] -= 1
    return base


def process_month(encs: list[dict], ledger_m: dict[str, int], log: list, fname: str):
    """encs: [{'cid','d','o','t'}] raw cells. Returns new (d, o, t) per encounter."""
    # bucket anchor per claim: sent bucket columns + cost-type columns into Dental
    bucket_amt = []
    rows: dict[str, int] = {b: 0 for b in BUCKETS}
    for e in encs:
        ba = {b: e["d"][b] + e["o"][b] for b in BUCKETS}
        ba[HOME_BUCKET] += sum(e["t"].values())
        bucket_amt.append(ba)
        for b in BUCKETS:
            rows[b] += ba[b]
    # tie the bucket anchor to the ledger to the fils (sent files drift by a few fils per month)
    resid = sum(ledger_m.values()) - sum(rows.values())
    if resid:
        i = max(range(len(bucket_amt)), key=lambda k: sum(bucket_amt[k].values()))
        bucket_amt[i][HOME_BUCKET] += resid
        rows[HOME_BUCKET] += resid
        log.append(f"{fname}: bucket anchor {resid/1000:+.3f} AED vs ledger absorbed by claim {encs[i]['cid']} ({HOME_BUCKET})")
    assert sum(rows.values()) == sum(ledger_m.values()), f"{fname}: bucket anchor != ledger"
    M = ras(rows, ledger_m, SEED)
    profiles = {b: {t: M[b][t] / Decimal(rows[b]) for t in M[b]} for b in M}   # share of bucket b in type t
    results = []
    col_sum = {t: 0 for t in COST_TYPES}
    for e, ba in zip(encs, bucket_amt):
        total = sum(ba.values())
        exact_t = {t: Decimal(0) for t in COST_TYPES}
        exact_bt = {}   # (bucket, type) -> Decimal, for the overhead split per bucket
        for b, v in ba.items():
            if v and b in profiles:
                for t, share in profiles[b].items():
                    amt = Decimal(v) * share
                    exact_t[t] += amt
                    exact_bt[(b, t)] = amt
        nt = round_to_total(exact_t, total)
        # overhead component per bucket = its overhead-type share (rounded, capped by the bucket amount)
        no = {b: 0 for b in BUCKETS}
        nd = {b: 0 for b in BUCKETS}
        for b, v in ba.items():
            if not v:
                continue
            oh = sum(exact_bt.get((b, t), Decimal(0)) for t in OVERHEAD_COST_TYPES)
            ohf = min(int(oh.quantize(Decimal(1), rounding=ROUND_HALF_UP)), v)
            no[b], nd[b] = ohf, v - ohf
        # keep the overhead bucket group equal to the overhead types on the claim
        oh_types = sum(nt[t] for t in OVERHEAD_COST_TYPES)
        delta = oh_types - sum(no.values())
        if delta:
            b = max((b for b in BUCKETS if ba[b]), key=lambda b: ba[b])
            no[b] += delta
            nd[b] -= delta
        assert sum(nd.values()) + sum(no.values()) == sum(nt.values()) == total
        for t in COST_TYPES:
            col_sum[t] += nt[t]
        results.append((nd, no, nt))
    # force every cost-type column to the ledger to the fils (residual vector sums to zero)
    resid = {t: ledger_m[t] - col_sum[t] for t in COST_TYPES}
    if any(resid.values()):
        need = [t for t in COST_TYPES if resid[t] < 0]
        cand = [i for i, (_, _, nt) in enumerate(results) if all(nt[t] + resid[t] >= 0 for t in need)]
        i = max(cand, key=lambda i: sum(results[i][2].values()))
        nd, no, nt = results[i]
        for t in COST_TYPES:
            nt[t] += resid[t]
        oh_delta = sum(resid[t] for t in OVERHEAD_COST_TYPES)
        if oh_delta:
            b = max((b for b in BUCKETS if no[b] + nd[b]), key=lambda b: no[b] + nd[b])
            no[b] += oh_delta
            nd[b] -= oh_delta
        assert sum(nd.values()) + sum(no.values()) == sum(nt.values())
        log.append(f"{fname}: fils residuals {{{', '.join(f'{t}:{v:+d}' for t, v in resid.items() if v)}}} absorbed by claim {encs[i]['cid']}")
    return results, profiles


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--indir", required=True)
    ap.add_argument("--ledger", required=True, help="input/Ledger_by_Cost_Type_FY2025.xlsx (Arjun)")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--zip", action="store_true")
    args = ap.parse_args(argv)

    ledger = read_ledger(Path(args.ledger))
    files = sorted(Path(args.indir).glob("*.xml"))
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone(timedelta(hours=TZ_OFFSET_HOURS))).strftime(DATE_FMT)
    log: list[str] = []
    print(f"{'File':<42}{'Claims':>7}{'Buckets':>16}{'Cost types':>16}{'Ledger':>16}{'Match':>7}{'OH%':>7}")
    rows_out = []
    for f in files:
        m = re.search(r"_M(\d\d)_", f.name)
        if not m:
            continue
        month = int(m.group(1))
        tree = etree.parse(str(f))
        root = tree.getroot()
        encs, nodes = [], []
        for claim in root.findall("Claim"):
            cid = claim.findtext("ID")
            for enc in claim.findall("Encounter"):
                d = {b: fils(enc.find("CostBucketDirect").findtext(b)) for b in BUCKETS}
                o = {b: fils(enc.find("CostBucketOverheads").findtext(b)) for b in BUCKETS}
                t = {c: fils(enc.find("CostTypeDetails").findtext(c)) for c in COST_TYPES}
                encs.append({"cid": cid, "d": d, "o": o, "t": t})
                nodes.append(enc)
        results, profiles = process_month(encs, ledger[month], log, f.name)
        for b, prof in profiles.items():
            log.append(f"{f.name}: bucket {b} profile " + ", ".join(f"{t} {float(s)*100:.1f}%" for t, s in prof.items() if s))
        B = T = OH = 0
        for enc, (nd, no, nt) in zip(nodes, results):
            for name, vals in (("CostBucketDirect", nd), ("CostBucketOverheads", no), ("CostTypeDetails", nt)):
                g = enc.find(name)
                for child in g:
                    child.text = fmt_money_fils(vals[child.tag])
            B += sum(nd.values()) + sum(no.values())
            T += sum(nt.values())
            OH += sum(nt[t] for t in OVERHEAD_COST_TYPES)
        td = root.find("Header/TransactionDate")
        if td is not None:
            td.text = now
        out = outdir / f.name
        tree.write(str(out), encoding="UTF-8", pretty_print=False)
        gl = int(GL[args.year][month] * 1000)
        ok = "OK" if B == T == gl else "FAIL"
        print(f"{f.name:<42}{len(encs):>7}{B/1000:>16,.3f}{T/1000:>16,.3f}{gl/1000:>16,.3f}{ok:>7}{OH/T*100:>6.1f}%")
        rows_out.append((f.name, ok))
    (outdir / "rebalance_log.txt").write_text("\n".join(log) + "\n")
    print(f"\n{len(rows_out)} files written to {outdir}; log in rebalance_log.txt")
    if args.zip:
        zp = outdir.parent / f"{outdir.name}.zip"
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, _ in rows_out:
                z.write(outdir / name, name)
        print(f"zip: {zp}")
    return 1 if any(ok != "OK" for _, ok in rows_out) else 0


if __name__ == "__main__":
    sys.exit(main())
