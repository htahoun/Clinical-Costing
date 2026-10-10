#!/usr/bin/env python3
"""
Cost Bucket / Cost Type Match correction (DOH SY02 rejection, 9 Oct 2026).

  python3 scripts/fix_crossfoot.py --indir input/source/mysmyle_zip --outdir output/xml_resubmission \
      [--overheads separated|direct] [--no-tie] [--zip]

DOH reads cost buckets (WHERE the money was spent) and cost types (WHAT it was
spent on) as two complete views of the same claim cost: each must add up to
the full amount on its own. The sent files split each dirham once across the
two views. This script re-expresses every claim so both views carry the full
amount, moving cells only — no money is created or lost:

  cost-type view  = existing cost types
                  + each bucket column shown by its nature (NATURE_MAP)
  bucket view     = existing bucket columns
                  + each cost-type column placed in the clinic's one department
                    (Dental): direct types -> CostBucketDirect/Dental,
                    overhead types -> CostBucketOverheads/Dental (--overheads
                    separated, default) or CostBucketDirect/Dental (direct).

Per month the cost-type view is then tied to the ledger to the fils (largest
claim absorbs the residual in both views) unless --no-tie. Everything else in
the files (identity fields, order, header RecordCount) is preserved verbatim;
TransactionDate is refreshed. Prints the verification table and writes a
change log.
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lxml import etree  # noqa: E402

from costsub_reader import fils  # noqa: E402
from costsub_schema import (BUCKETS, COST_TYPES, DATE_FMT, OVERHEAD_COST_TYPES,  # noqa: E402
                            TZ_OFFSET_HOURS, fmt_money_fils)
from gl_new import GL  # noqa: E402

# Finance Director's ruling, 10 Oct 2026: what each bucket column is made of.
NATURE_MAP = {
    "Dental": "MS",            # dental supplies
    "Laboratory": "Laboratory",
    "Prosthesis": "Prostheses",
    "Imaging": "Imaging",
    "Anaesthesia": "Pharmacy",  # anaesthetic drugs
    "OP": "SWNonClin",
    "Physician": "SWDoc",
}
FALLBACK_NATURE = "MS"         # any other (currently zero) bucket column
HOME_BUCKET = "Dental"         # the clinic's single department


def _group(enc, name):
    g = enc.find(name)
    if g is None:
        raise SystemExit(f"encounter without <{name}>")
    return g


def read_cells(enc):
    d = {b: fils(_group(enc, "CostBucketDirect").findtext(b)) for b in BUCKETS}
    o = {b: fils(_group(enc, "CostBucketOverheads").findtext(b)) for b in BUCKETS}
    t = {c: fils(_group(enc, "CostTypeDetails").findtext(c)) for c in COST_TYPES}
    return d, o, t


def write_cells(enc, d, o, t):
    for name, vals in (("CostBucketDirect", d), ("CostBucketOverheads", o), ("CostTypeDetails", t)):
        g = _group(enc, name)
        for child in g:
            child.text = fmt_money_fils(vals[child.tag])


def reexpress(d, o, t, overheads: str, log: list, cid: str):
    """Return (new_direct, new_overheads, new_types). Money moves, never changes."""
    total = sum(d.values()) + sum(o.values()) + sum(t.values())
    nd, no, nt = dict(d), dict(o), dict(t)
    # 1. bucket columns -> cost-type view by nature
    for b, v in d.items():
        if v:
            nature = NATURE_MAP.get(b)
            if nature is None:
                nature = FALLBACK_NATURE
                log.append(f"{cid}: bucket {b} {v/1000:.3f} has no nature mapping; shown as {nature}")
            nt[nature] += v
    for b, v in o.items():          # pre-existing overhead-bucket cells (none in SY02 files)
        if v:
            nt["OHF"] += v
            log.append(f"{cid}: CostBucketOverheads {b} {v/1000:.3f} shown as OHF")
    # 2. cost-type columns -> bucket view, in the clinic's one department
    for c, v in t.items():
        if v:
            if c in OVERHEAD_COST_TYPES and overheads == "separated":
                no[HOME_BUCKET] += v
            else:
                nd[HOME_BUCKET] += v
    assert sum(nd.values()) + sum(no.values()) == total == sum(nt.values()), cid
    return nd, no, nt


def tie_month(encs: list, gl: int, log: list, fname: str):
    """Make the cost-type view (and so the bucket view) equal the ledger to the fils."""
    view = sum(sum(t.values()) for _, _, _, t in encs)
    diff = gl - view
    if not diff:
        return 0
    i = max(range(len(encs)), key=lambda k: sum(encs[k][3].values()))
    cid, nd, no, nt = encs[i]
    kt = max(nt, key=nt.get)
    nt[kt] += diff
    if nd[HOME_BUCKET] >= max(no.values(), default=0):
        nd[HOME_BUCKET] += diff
    else:
        kb = max(no, key=no.get)
        no[kb] += diff
    log.append(f"{fname}: residual {diff/1000:+.3f} AED vs ledger absorbed by claim {cid} ({kt} and {HOME_BUCKET} bucket)")
    return diff


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--indir", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--overheads", choices=["separated", "direct"], default="separated")
    ap.add_argument("--no-tie", action="store_true", help="do not force the month to the ledger figure")
    ap.add_argument("--zip", action="store_true")
    args = ap.parse_args(argv)

    files = sorted(Path(args.indir).glob("*.xml"))
    if not files:
        raise SystemExit(f"no XML files in {args.indir}")
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone(timedelta(hours=TZ_OFFSET_HOURS))).strftime(DATE_FMT)
    log: list[str] = []
    rows = []
    print(f"{'File':<42}{'Claims':>7}{'Before: buckets':>16}{'types':>14}{'After: each view':>18}{'Ledger':>15}{'Tie':>8}")
    for f in files:
        m = re.search(r"_M(\d\d)_", f.name)
        month = int(m.group(1)) if m else None
        tree = etree.parse(str(f))
        root = tree.getroot()
        b0 = t0 = 0
        encs, nodes = [], []
        for claim in root.findall("Claim"):
            cid = claim.findtext("ID")
            for enc in claim.findall("Encounter"):
                d, o, t = read_cells(enc)
                b0 += sum(d.values()) + sum(o.values())
                t0 += sum(t.values())
                nd, no, nt = reexpress(d, o, t, args.overheads, log, cid)
                encs.append((cid, nd, no, nt))
                nodes.append(enc)
        gl = int(GL[args.year][month] * 1000) if month else None
        tie = "n/a"
        if gl is not None and not args.no_tie:
            tie_month(encs, gl, log, f.name)
            tie = "OK" if sum(sum(t.values()) for _, _, _, t in encs) == gl else "FAIL"
        for enc, (_, nd, no, nt) in zip(nodes, encs):
            write_cells(enc, nd, no, nt)
            assert sum(nd.values()) + sum(no.values()) == sum(nt.values())
        td = root.find("Header/TransactionDate")
        if td is not None:
            td.text = now
        out = outdir / f.name
        tree.write(str(out), encoding="UTF-8", pretty_print=False)
        after = sum(sum(t.values()) for _, _, _, t in encs)
        print(f"{f.name:<42}{len(encs):>7}{b0/1000:>16,.3f}{t0/1000:>14,.3f}{after/1000:>18,.3f}"
              f"{(gl/1000 if gl else 0):>15,.3f}{tie:>8}")
        rows.append((f.name, len(encs), b0, t0, after, gl, tie, out.stat().st_size))
    (outdir / "crossfoot_fix_log.txt").write_text("\n".join(log) + "\n")
    print(f"\n{len(files)} files written to {outdir}; overheads {args.overheads}; {len(log)} log lines in crossfoot_fix_log.txt")
    if args.zip:
        zp = outdir.parent / f"{outdir.name}.zip"
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            for r in rows:
                z.write(outdir / r[0], r[0])
        print(f"zip: {zp}")
    return 1 if any(r[6] == "FAIL" for r in rows) else 0


if __name__ == "__main__":
    sys.exit(main())
