#!/usr/bin/env python3
"""
Generation-2 Shafafiya Cost.Submission XML generator + inline validation.

  python3 scripts/make_xml_final.py --workbook output/Clinical_Costing_2025_MF4159_FINAL.xlsx \
      --outdir output/xml_final --disposition PTE_SUBMIT --supplies-bucket Dental \
      [--xsd docs/shafafiya/CostSubmission_repaired.xsd] [--zip]

One <Claim> per claim ID per month, costs summed across the claim's lines in
integer fils. Conventions: docs/CLINICAL_COSTING.md §4–§5. Zero-total claims
are excluded and logged; negative claims abort the month. Prints the
verification table (month, claims, size, total, ledger tie, XSD result).
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from collections import OrderedDict
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from lxml import etree  # noqa: E402

from build2_lib import to_fils  # noqa: E402
from costsub_reader import xsd_validate  # noqa: E402
from costsub_schema import (DATE_FMT, DIAGNOSIS_ORDER, DISPOSITIONS, ENCOUNTER_ORDER, FACILITY_ID,  # noqa: E402
                            HEADER_ORDER, MAX_FILE_BYTES, MONEY_ELEMENTS, MONEY_GROUPS, PRINCIPAL,
                            PROCEDURE_ORDER, RECEIVER_ID, ROOT, TZ_OFFSET_HOURS, bucket_map, fmt_money_fils)
from gl_new import gl_fils  # noqa: E402

GST = timezone(timedelta(hours=TZ_OFFSET_HOURS))


def read_costed(workbook: str, sheet: str = "Sheet1"):
    wb = openpyxl.load_workbook(workbook, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    headers = [str(h) for h in next(rows)]
    hi = {h: i for i, h in enumerate(headers)}
    money_cols = [h for h in headers if h.startswith(("CostBucketDirect", "CostTypeDetails"))]
    lines = []
    for r in rows:
        if r is None or r[hi["Claim ID"]] in (None, ""):
            continue
        g = lambda k: r[hi[k]] if k in hi else None
        lines.append({
            "month": int(g("Month")), "claim_id": str(g("Claim ID")).strip(), "serial": str(g("Serial") or "").strip(),
            "patient_id": str(g("FileNo.") or "").strip(), "clinician": str(g("Doctor No") or "").strip(),
            "specialty": str(g("Speciality") or "").strip(), "start": str(g("StartDateTime")).strip(),
            "end": str(g("EndDateTime") or g("StartDateTime")).strip(), "code": str(g("Code") or "").strip(),
            "diagnosis": str(g("Diagnosis") or "").strip(), "theatre": str(g("TheatreAttendanceFlag") or "0").strip(),
            "procedure": str(g("Procedure") or "").strip(), "duration": str(g("Duration") or "0").strip(),
            "enc_type": str(g("Type") or "1").strip(),
            "money": {c: to_fils(g(c)) for c in money_cols},
        })
    return lines, money_cols


def build_claims(lines: list[dict], bmap: dict, log: list[str]) -> OrderedDict:
    claims: OrderedDict[str, dict] = OrderedDict()
    for ln in lines:
        c = claims.get(ln["claim_id"])
        if c is None:
            c = claims[ln["claim_id"]] = {
                "ID": ln["claim_id"], "PatientID": ln["patient_id"], "EncounterID": ln["serial"],
                "Clinician": ln["clinician"], "Specialty": ln["specialty"], "Start": ln["start"],
                "End": ln["end"], "Type": ln["enc_type"], "Diagnosis": ln["diagnosis"],
                "Theatre": "0", "Procedure": "", "Duration": "0", "money": {e: 0 for e in MONEY_ELEMENTS},
                "lines": 0,
            }
        if ln["theatre"] == "1":
            c["Theatre"] = "1"
            if ln["procedure"] and not c["Procedure"]:
                c["Procedure"], c["Duration"] = ln["procedure"], ln["duration"] or "0"
        for col, f in ln["money"].items():
            el = bmap.get(col, col)
            if el not in c["money"]:
                raise SystemExit(f"workbook column {col} maps to {el}, which is not a Gen2 element")
            c["money"][el] += f
        c["lines"] += 1
    for cid, c in list(claims.items()):
        total = sum(c["money"].values())
        c["total"] = total
        if total == 0:
            log.append(f"claim {cid}: zero total — EXCLUDED (DOH blocks zero-total encounters)")
            del claims[cid]
        elif total < 0 or any(v < 0 for v in c["money"].values()):
            raise SystemExit(f"claim {cid}: negative cost cell — DOH blocks negatives; fix the source")
    return claims


def claim_element(c: dict) -> etree._Element:
    cl = etree.Element("Claim")
    etree.SubElement(cl, "ID").text = c["ID"]
    etree.SubElement(cl, "ProviderID").text = FACILITY_ID
    enc = etree.SubElement(cl, "Encounter")
    simple = {
        "FacilityID": FACILITY_ID, "ID": c["EncounterID"], "ActivityClinicianID": c["Clinician"],
        "PatientID": c["PatientID"], "Type": c["Type"], "Start": c["Start"], "End": c["End"],
        "StartType": "1", "Specialty": c["Specialty"], "SubSpecialty": "", "DRGCode": "",
        "ServiceCode": "", "EndType": "1", "EDTriageLevel": "0", "OPAttendanceType": "",
        "TheatreAttendanceFlag": c["Theatre"], "CriticalCareTime": "0", "VentilationTime": "0",
    }
    for name in ENCOUNTER_ORDER:
        if name == "Diagnosis":
            d = etree.SubElement(enc, "Diagnosis")
            vals = {"Type": PRINCIPAL, "Code": c["Diagnosis"]}
            for k in DIAGNOSIS_ORDER:
                etree.SubElement(d, k).text = vals[k]
        elif name == "Procedure":
            p = etree.SubElement(enc, "Procedure")
            vals = {"Type": "", "Code": c["Procedure"] if c["Theatre"] == "1" else "",
                    "Duration": c["Duration"] if c["Theatre"] == "1" else "0"}
            for k in PROCEDURE_ORDER:
                etree.SubElement(p, k).text = vals[k]
        elif name in MONEY_GROUPS:
            g = etree.SubElement(enc, name)
            for child in MONEY_GROUPS[name]:
                etree.SubElement(g, child).text = fmt_money_fils(c["money"].get(name + child, 0))
        else:
            etree.SubElement(enc, name).text = simple[name]
    return cl


def write_month(claims: OrderedDict, disposition: str, out: Path, now: datetime) -> None:
    root = etree.Element(ROOT)
    hdr = etree.SubElement(root, "Header")
    vals = {"SenderID": FACILITY_ID, "ReceiverID": RECEIVER_ID, "TransactionDate": now.strftime(DATE_FMT),
            "RecordCount": str(len(claims)), "DispositionFlag": disposition}
    for k in HEADER_ORDER:
        etree.SubElement(hdr, k).text = vals[k]
    for c in claims.values():
        root.append(claim_element(c))
    out.parent.mkdir(parents=True, exist_ok=True)
    etree.ElementTree(root).write(str(out), xml_declaration=True, encoding="UTF-8", pretty_print=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--disposition", default="PTE_SUBMIT", choices=DISPOSITIONS)
    ap.add_argument("--supplies-bucket", default="Dental", help="Dental (accepted by DOH in SY02) or another Gen2 bucket")
    ap.add_argument("--xsd", default=str(Path(__file__).resolve().parents[1] / "docs/shafafiya/CostSubmission_repaired.xsd"))
    ap.add_argument("--months", nargs="*", type=int, help="subset of months to generate")
    ap.add_argument("--zip", action="store_true", help="also write a zip of the twelve files")
    args = ap.parse_args(argv)

    lines, _ = read_costed(args.workbook)
    bmap = bucket_map(args.supplies_bucket)
    outdir = Path(args.outdir)
    now = datetime.now(GST)
    log: list[str] = []
    months = sorted({ln["month"] for ln in lines})
    if args.months:
        months = [m for m in months if m in args.months]
    print(f"{'File':<44}{'Claims':>7}{'Size':>10}{'Total AED':>16}{'Ledger':>16}{'Tie':>6}  XSD")
    results, written = [], []
    for m in months:
        mlines = [ln for ln in lines if ln["month"] == m]
        claims = build_claims(mlines, bmap, log)
        out = outdir / f"CostSubmission_{FACILITY_ID}_{args.year}{m:02d}.xml"
        write_month(claims, args.disposition, out, now)
        size = out.stat().st_size
        total = sum(c["total"] for c in claims.values())
        gl = gl_fils(args.year, m)
        tie = "OK" if total == gl else f"{(total - gl) / 1000:+.3f}"
        xsd, msgs = xsd_validate(out, args.xsd)
        if size > MAX_FILE_BYTES:
            log.append(f"{out.name}: {size} bytes exceeds 5 MB — split the month")
        print(f"{out.name:<44}{len(claims):>7}{size:>10,}{total / 1000:>16,.3f}{gl / 1000:>16,.3f}{tie:>6}  {xsd}")
        for msg in msgs[:5]:
            print(f"    {msg}")
        results.append((out.name, len(claims), size, total, gl, tie, xsd))
        written.append(out)
    logp = outdir / "generation_log.txt"
    logp.write_text("\n".join(log) + "\n")
    print(f"\n{len(written)} file(s) in {outdir}; {len(log)} log entries in {logp.name}; "
          f"disposition {args.disposition}; Supplies -> {args.supplies_bucket}")
    if args.zip:
        zp = outdir.parent / f"CostSubmission_{FACILITY_ID}_{args.year}_XML_{args.disposition}.zip"
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            for w in written:
                z.write(w, w.name)
        print(f"zip: {zp}")
    bad = [r for r in results if r[5] != "OK" or r[6] == "FAIL"]
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
