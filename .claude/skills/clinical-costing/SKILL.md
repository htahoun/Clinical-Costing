---
name: clinical-costing
description: Build, audit, submit and — when DOH rejects — triage and resubmit the Abu Dhabi DOH Clinical Cost Data Collection (ADCCDC) for Vision Dental Clinic L.L.C (licence MF4159) — scale-only costing of claim-level activity to the monthly expenditure ledger, Generation-2 Shafafiya Cost.Submission XML creation with XSD and portal-rule validation, rejection triage, and the standard claims/XML audit. Use whenever the user mentions clinical costing, ADCCDC, DOH cost submission, Shafafiya, Cost.Submission XML, a rejection, cost buckets or cost types, the expenditure ledger, the CPT allocation table, monthly factors, or asks to audit claim exports or platform XMLs (MySmyle, Calibrate, ClaimScanner).
---

# Clinical Costing — Vision Dental Clinic (MF4159)

Produces the annual Abu Dhabi Clinical Cost Data Collection submission to the
Department of Health via Shafafiya, audits anything claiming to be it, and
works DOH rejections to resubmission. The master reference with every locked
decision is `docs/CLINICAL_COSTING.md` — read it before changing anything.
Rejections follow `docs/REJECTION_PLAYBOOK.md`. Engines live in `scripts/`.

## Working etiquette (the user's standing rules)

1. **Digest → confirm → wait for "go".** Never build from a freshly uploaded
   file. Read it, state the understanding (figures, counts, deltas; for a
   rejection the triage table), ask up to 4 on-screen MCQs only when
   genuinely ambiguous, and wait for explicit permission before producing
   output.
2. The user is the Finance Director, often on a phone; the coding team is
   non-technical. Explain in plain language, deliver downloadable files, and
   keep spreadsheets — not XML — as the human-facing format.
3. A decision the user confirms is **locked**: record it, apply it everywhere,
   never re-litigate it (e.g. removed claims stay removed).
4. Always reconcile and show the proof: a month that does not tie to the
   ledger to the fils is a defect, not a rounding note.
5. The repository is private: commit inputs, outputs, submitted XMLs and
   rejection files alongside code and docs so comparisons are reproducible.

## The pipeline (seven steps)

1. **Activity**: one row = one unit of one CPT code inside one claim (per-month
   tabs or one sheet; month comes from `StartDateTime`). Accepted headings:
   `docs/activity-template-columns.md`. The reader strips trailing `B` from
   serials and fills blank `Speciality` with the clinician's modal speciality.
2. **Allocation table** (`load_alloc` in `scripts/build2_lib.py`): per CPT
   code a true Cost/Unit plus a 17-column % split that sums to 100. The locked
   fixes (re-keys, dedupes, rescales, orphan map) are applied and logged.
3. **Scale-only costing (locked — no "golden percentage" nudging):**
   ```bash
   python3 scripts/build4_final.py --activity input/activity.xlsx --alloc input/allocation.xlsx \
       --out output/Clinical_Costing_2025_MF4159_FINAL.xlsx
   ```
   modelled month cost = Σ Cost/Unit; factor = ledger ÷ modelled; cell =
   Cost/Unit × factor × column %; 3 dp per cell; column sums forced to their
   rounded targets; month total forced to `round(GL, 3)` exactly (integer
   fils throughout). Ledger: `scripts/gl_new.py`.
4. **Workbook**: Sheet1 (lines + `Source` provenance), Reconciliation
   (12 months, zero differences), Monthly Factors, Merge Log.
5. **XML**: twelve monthly Generation-2 `Cost.Submission` files:
   ```bash
   python3 scripts/make_xml_final.py --workbook output/Clinical_Costing_2025_MF4159_FINAL.xlsx \
       --outdir output/xml_final --disposition PRODUCTION --supplies-bucket Other --zip
   python3 scripts/validate_xml.py output/xml_final/*.xml --expect-disposition PRODUCTION
   ```
   One `<Claim>` per claim ID, costs summed across the claim's lines; zero-total
   claims excluded and logged; the verification table (claims, size, total,
   ledger tie, XSD) is printed. Every schema-dependent name is in
   `scripts/costsub_schema.py`; the XSD goes in `docs/shafafiya/`.
6. **Audit** (`python3 scripts/audit_mysmyle_xml.py <mine> <theirs>`): XSD both
   files, headers, claim-set equality, all identity/meta fields, per-cell money
   diffs > 0.0005, zero-total and negative checks, largest diffs, verdict.
7. **Rejection** (`python3 scripts/triage_rejection.py input/rejection/* --xml output/xml_final`):
   one table of distinct messages → count → claims → month → suspected cause →
   fix; classify (packaging / schema / reference data / money); fix once in the
   engine; regenerate, validate, audit one month against the sent file;
   resubmit (`scripts/submit_shafafiya.py upload … --env prod --send`, credentials
   from `SHAFAFIYA_USER` / `SHAFAFIYA_PASS`); record the ruling.

## Locked XML conventions (Generation 2)

- `Claim.ID` = claim reference (e.g. `1747MF4159-5a058f`); `Encounter.ID` =
  the claim's **first line's** activity serial; `ActivityClinicianID` =
  `Doctor No`; `PatientID` = `FileNo.`; Facility/Provider/Sender = `MF4159`.
- `ReceiverID` = `HAAD` (only enumerated value). Dates `dd/mm/yyyy HH:MM`
  (GMT+4). Amounts AED, 3 dp. Files ≤ 5 MB → one file per month.
- Type/StartType/EndType = 1; SubSpecialty, DRGCode, ServiceCode,
  OPAttendanceType empty; **Principal diagnosis only**; Procedure empty
  (Duration 0) unless `TheatreAttendanceFlag = 1`; triage/critical-care/
  ventilation 0; `CostBucketOverheads` carries each bucket's overhead share
  (= OHF + OHDepreciation + OHC).
- **Supplies is reported in the `Dental` bucket** (ruling 10 Oct 2026).
- **Two complete views (DOH Cost Bucket / Cost Type Match):** on every claim
  Σ CostBucketDirect + Σ CostBucketOverheads = Σ CostTypeDetails = claim cost;
  overhead cost types sit in `CostBucketOverheads/Dental`. The generator
  re-expresses by nature; `scripts/rebalance_two_view.py` calibrates the
  cost-type view to the ledger by cost type.
- Zero-total claims are blocked by DOH: exclude and log them.
- `DispositionFlag`: `PTE_SUBMIT` for the test environment, `PRODUCTION`
  for live. Transport: SOAP `UploadTransaction`, base64;
  PTE `https://shafafiyapte.doh.gov.ae/v3/webservices.asmx` (username +
  `test`), production `https://shafafiya.doh.gov.ae/v3/webservices.asmx`.
  Resubmission only on DOH request (a rejection is that request).

## Definition of done

Twelve XMLs, each: `validate_xml.py` PASS, XSD PASS, month total = ledger to
the fils, claim count reconciled to the workbook, < 5 MB, correct
DispositionFlag. Deliver workbook + zip + the verification table (per month:
claims, size, total, ledger tie, XSD result). Commit and push everything, data included.
After a rejection, additionally: triage table, the ruling recorded in the
master reference, and one month audited against the previously sent file.
