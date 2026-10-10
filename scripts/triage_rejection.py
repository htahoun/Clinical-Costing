#!/usr/bin/env python3
"""
Rejection triage: digest whatever DOH/Shafafiya returned (response XML, Excel,
CSV, or pasted text/email) into one table — message, count, claims affected,
suspected cause, and the fix in our pipeline.

  python3 scripts/triage_rejection.py <rejection files...> [--xml output/xml_final] \
      [--out output/rejection_triage.xlsx]

Each line of every file is scanned for claim IDs (nnnnMF4159-xxxxxx), encounter
serials, element names and error codes/messages. Messages are matched against the
cause catalogue (CATALOGUE) that maps known Shafafiya/ADCCDC rejections to the
pipeline switch or data fix that clears them. Unknown messages are listed for the
Finance Director and the clinical costing engineer to work through together.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

CLAIM_RE = re.compile(r"\b\d{1,6}MF4159-[0-9a-f]{6}\b", re.I)
ELEMENT_RE = re.compile(r"\b(Cost\.Submission|Header|Claim|Encounter|SenderID|ReceiverID|TransactionDate|"
                        r"RecordCount|DispositionFlag|ProviderID|PatientID|FacilityID|ActivityClinicianID|"
                        r"ServiceCode|Specialty|SubSpecialty|DRGCode|OPAttendanceType|TheatreAttendanceFlag|"
                        r"Diagnosis|Procedure|CostBucket\w+|CostTypeDetails\w+|TotalCost|StartDateTime|EndDateTime)\b")
CODE_RE = re.compile(r"\b(?:error|err|code|rule)\s*[:#]?\s*([A-Z]{1,5}[-_]?\d{2,5})\b", re.I)

# (pattern, cause, fix). Order matters: first match wins.
CATALOGUE = [
    (r"disposition|PTE_SUBMIT|test file|production", "File carried a test DispositionFlag (PTE_SUBMIT) or was sent to the wrong environment.",
     "Regenerate with --disposition PRODUCTION and submit to the production endpoint (scripts/make_xml_final.py, scripts/submit_shafafiya.py)."),
    (r"schema|xsd|not valid|invalid xml|element .* (is not expected|was not expected|missing)|unexpected element|well.?formed",
     "XML does not match the current CostSubmission XSD (element name, order, or a missing/extra element).",
     "Place the official XSD in docs/shafafiya/, run scripts/validate_xml.py, align scripts/costsub_schema.py (ENCOUNTER_ORDER / names) and regenerate."),
    (r"duplicate|already (been )?submitted|already exists|resubmission", "Claim or encounter IDs already exist on the portal (earlier upload accepted, or the same claim in two monthly files).",
     "Confirm with DOH which upload is live. Resubmit only on DOH request; if two files carry one claim, keep it in its StartDateTime month only."),
    (r"total.*(not|does not|doesn't).*(equal|match|sum)|sum of|cross|bucket.*type|cost type.*bucket", "DOH cross-foot: TotalCost must equal the sum of cost buckets AND the sum of cost types.",
     "Our locked split sums buckets + types = total. If DOH enforces buckets = types = total, the 17-column split must be re-expressed as two 100% views (needs the Finance Director's ruling; see docs/REJECTION_PLAYBOOK.md H1)."),
    (r"zero|0\.000|no cost|empty cost", "A claim/encounter with zero total cost reached the file.",
     "scripts/make_xml_final.py excludes zero-total claims; check generation_log.txt and the source lines."),
    (r"negative", "A negative cost cell.", "Fix at source (activity line or allocation split); negatives are blocked."),
    (r"clinician|ActivityClinicianID|licen[cs]e.*(doctor|clinician|practitioner)", "ActivityClinicianID is not a DOH-licensed clinician at MF4159 (wrong format, inactive licence, or staff number instead of DOH licence).",
     "Map Doctor No to the DOH clinician licence (e.g. GDxxxxx) in the activity master; regenerate."),
    (r"facility|provider|sender|licen[cs]e.*(facility|provider)|MF4159", "Facility/Provider/Sender ID rejected (licence mismatch or the sending account is not linked to MF4159).",
     "Confirm the Shafafiya account is linked to MF4159; all three IDs must be MF4159."),
    (r"receiver", "ReceiverID is not HAAD.", "ReceiverID must be HAAD (only enumerated value)."),
    (r"date|period|financial year|outside|future", "A date is outside the reporting period, badly formatted, or in the future.",
     "Dates must be dd/mm/yyyy HH:MM, all inside FY2025; scripts/validate_xml.py flags offenders."),
    (r"patient|PatientID|member|emirates", "PatientID rejected (blank, or not the identifier DOH expects).",
     "We send FileNo.; if DOH wants the Emirates ID / member ID, swap the column in the activity master."),
    (r"diagnos|icd", "Diagnosis code invalid or not an accepted ICD-10 code / missing principal diagnosis.",
     "Check the code list; principal only; regenerate."),
    (r"procedure|cpt|activity code|service code", "Procedure/activity/service code rejected (or ServiceCode required but empty).",
     "ServiceCode is sent empty by design; if DOH wants the CPT code there, fill it from the first line's code (one-line change in make_xml_final.py)."),
    (r"specialt", "Specialty value not in the DOH list.", "Map Speciality text to the DOH specialty code list; regenerate."),
    (r"size|too large|5 ?mb", "File larger than 5 MB.", "Split the month into two files (--months, or split by claim range)."),
    (r"record ?count", "RecordCount does not equal the number of claims.", "Regenerate; the generator writes RecordCount = claims."),
    (r"overhead", "CostBucketOverheads expected non-zero or the reverse.", "Overheads are reported in CostTypeDetails by design; ask DOH which view they validate."),
    (r"encounter.*id|ID.*encounter", "Encounter ID rejected (duplicate across claims or wrong format).", "Encounter.ID = first line's activity serial; check duplicates with scripts/validate_xml.py."),
    (r"authent|password|login|credential|unauthori", "Portal credentials/authorisation problem, not the file.", "Re-check the Shafafiya username (PTE uses the `test` suffix) and password."),
]


def lines_from(path: Path) -> list[str]:
    suf = path.suffix.lower()
    if suf in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        out = []
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                if r and any(v not in (None, "") for v in r):
                    out.append(" | ".join("" if v is None else str(v) for v in r))
        return out
    if suf == ".csv":
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            return [" | ".join(r) for r in csv.reader(f) if any(r)]
    if suf == ".xml":
        from lxml import etree
        try:
            tree = etree.parse(str(path))
            out = []
            for el in tree.iter():
                if el.text and el.text.strip() and len(el) == 0:
                    tag = el.tag.split("}", 1)[-1]
                    out.append(f"{tag}: {el.text.strip()}")
            return out
        except Exception as ex:  # noqa: BLE001
            print(f"  {path.name}: not parseable as XML ({ex}); reading as text")
    return path.read_text(encoding="utf-8", errors="replace").splitlines()


def classify(msg: str):
    for pat, cause, fix in CATALOGUE:
        if re.search(pat, msg, re.I):
            return cause, fix
    return "UNKNOWN — discuss", "Work through with the Finance Director; add to CATALOGUE once understood."


def normalise(msg: str) -> str:
    m = CLAIM_RE.sub("<claim>", msg)
    m = re.sub(r"\d+\.\d{3}", "<amt>", m)
    m = re.sub(r"\d{2}/\d{2}/\d{4}( \d{2}:\d{2})?", "<date>", m)
    m = re.sub(r"\s+", " ", m).strip()
    return m[:160]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--xml", help="folder of our submitted XMLs, to say which month each claim sits in")
    ap.add_argument("--out", default="output/rejection_triage.xlsx")
    args = ap.parse_args(argv)

    claim_month = {}
    if args.xml:
        from costsub_reader import read_submission
        for p in sorted(Path(args.xml).glob("*.xml")):
            for cid in read_submission(p)["claims"]:
                claim_month[cid.lower()] = p.name
    groups: dict[str, dict] = defaultdict(lambda: {"count": 0, "claims": set(), "elements": Counter(),
                                                   "codes": Counter(), "examples": [], "files": set()})
    for f in args.files:
        p = Path(f)
        if not p.exists():
            print(f"missing: {p}")
            continue
        raw = lines_from(p)
        print(f"{p.name}: {len(raw)} lines")
        for ln in raw:
            if not ln.strip():
                continue
            key = normalise(ln)
            g = groups[key]
            g["count"] += 1
            g["claims"].update(c.lower() for c in CLAIM_RE.findall(ln))
            g["elements"].update(ELEMENT_RE.findall(ln))
            g["codes"].update(c.upper() for c in CODE_RE.findall(ln))
            g["files"].add(p.name)
            if len(g["examples"]) < 3:
                g["examples"].append(ln.strip()[:300])

    rows = []
    for key, g in sorted(groups.items(), key=lambda kv: -kv[1]["count"]):
        cause, fix = classify(key + " " + " ".join(g["elements"]))
        months = Counter(claim_month.get(c, "?") for c in g["claims"])
        rows.append([key, g["count"], len(g["claims"]), ", ".join(sorted(g["elements"]))[:80],
                     ", ".join(sorted(g["codes"])), ", ".join(f"{m}:{n}" for m, n in months.most_common()),
                     cause, fix, " || ".join(g["examples"]), ", ".join(sorted(g["files"]))])

    print(f"\n{'Count':>6} {'Claims':>6}  Message (normalised)                                   -> suspected cause")
    for r in rows[:40]:
        print(f"{r[1]:>6} {r[2]:>6}  {r[0][:55]:<55} -> {r[6][:70]}")
    all_claims = set().union(*(g["claims"] for g in groups.values())) if groups else set()
    print(f"\n{len(rows)} distinct messages; {len(all_claims)} distinct claim IDs mentioned; "
          f"{sum(1 for r in rows if r[6].startswith('UNKNOWN'))} unknown message types")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Triage"
    ws.append(["Message (normalised)", "Occurrences", "Claims", "Elements named", "Error codes", "Months (from our XML)",
               "Suspected cause", "Fix in pipeline", "Examples", "Source files"])
    for r in rows:
        ws.append(r)
    cs = wb.create_sheet("Claims")
    cs.append(["Claim ID", "Our file", "Messages"])
    per_claim = defaultdict(list)
    for key, g in groups.items():
        for c in g["claims"]:
            per_claim[c].append(key)
    for c in sorted(per_claim):
        cs.append([c, claim_month.get(c, "?"), " || ".join(per_claim[c])[:500]])
    for sh in wb.worksheets:
        for cell in sh[1]:
            cell.font = Font(bold=True)
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    print(f"wrote {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
