---
name: clinical-costing
description: Build, audit and submit the Abu Dhabi DOH Clinical Cost Data Collection (ADCCDC) for Vision Dental Clinic L.L.C (licence MF4159) — scale-only costing of claim-level activity to the monthly expenditure ledger, Generation-2 Shafafiya Cost.Submission XML creation with XSD validation, and the standard claims/XML audit. Use whenever the user mentions clinical costing, ADCCDC, DOH cost submission, Shafafiya, Cost.Submission XML, cost buckets or cost types, the expenditure ledger, the CPT allocation table, monthly factors, or asks to audit claim exports or platform XMLs (MySmyle, Calibrate, ClaimScanner).
---

# Clinical Costing — Vision Dental Clinic (MF4159)

Produces the annual Abu Dhabi Clinical Cost Data Collection submission to the
Department of Health via Shafafiya, and audits anything claiming to be it.
The master reference with every locked decision is `docs/CLINICAL_COSTING.md`
— read it before changing anything. Reusable engines live in `scripts/`.

## Working etiquette (the user's standing rules)

1. **Digest → confirm → wait for "go".** Never build from a freshly uploaded
   file. Read it, state the understanding (figures, counts, deltas), ask up to
   4 on-screen MCQs only when genuinely ambiguous, and wait for explicit
   permission before producing output.
2. The user is the Finance Director, often on a phone; the coding team is
   non-technical. Explain in plain language, deliver downloadable files, and
   keep spreadsheets — not XML — as the human-facing format.
3. A decision the user confirms is **locked**: record it, apply it everywhere,
   never re-litigate it (e.g. removed claims stay removed).
4. Always reconcile and show the proof: a month that does not tie to the
   ledger to the fils is a defect, not a rounding note.

## The pipeline (six steps)

1. **Activity**: one row = one unit of one CPT code inside one claim, in the
   73-column workbook template (per-month tabs or one sheet; month comes from
   `StartDateTime`). Clean serials (strip trailing `B`), fill blank
   `Speciality` with the clinician's modal speciality.
2. **Allocation table** (`load_alloc` in `scripts/build2_lib.py`): per CPT
   code a true Cost/Unit plus a 17-column % split that sums to 100. Apply the
   locked fixes (re-keys, dedupes, rescales, orphan map) listed in
   `docs/CLINICAL_COSTING.md`.
3. **Scale-only costing (locked — no "golden percentage" nudging):**
   - modelled month cost = Σ Cost/Unit over the month's rows
   - factor = ledger month expenditure ÷ modelled cost
   - every cell = Cost/Unit × factor × column %
   - round every cell to 3 dp, force each column's rounded sum to the rounded
     column target, then force the month total to `round(GL, 3)` **exactly**
     (adjust the largest cell). Engine: `scripts/build4_final.py`.
4. **Workbook**: Sheet1 (lines + `Source` provenance column), Reconciliation
   (12 months, zero differences), Monthly Factors, Merge Log.
5. **XML**: twelve monthly Generation-2 `Cost.Submission` files via
   `scripts/make_xml_final.py` — one `<Claim>` per claim ID, costs summed
   across the claim's lines, conventions below. Validate every file against
   `docs/shafafiya/CostSubmission_repaired.xsd` before delivering.
6. **Audit** (`scripts/audit_mysmyle_xml.py <mine> <theirs>`): XSD both files,
   headers, claim-set equality, all identity/meta fields, per-cell money
   diffs > 0.0005, zero-total and negative checks, largest diffs. Same
   standard every time, so results are comparable across months.

## Locked XML conventions (Generation 2)

- `Claim.ID` = claim reference (e.g. `1747MF4159-5a058f`); `Encounter.ID` =
  the claim's **first line's** activity serial; `ActivityClinicianID` =
  `Doctor No`; `PatientID` = `FileNo.`; Facility/Provider/Sender = `MF4159`.
- `ReceiverID` = `HAAD` (only enumerated value). Dates `dd/mm/yyyy HH:MM`
  (GMT+4). Amounts AED, 3 dp. Files ≤ 5 MB → one file per month.
- Type/StartType/EndType = 1; SubSpecialty, DRGCode, ServiceCode,
  OPAttendanceType empty; **Principal diagnosis only**; Procedure empty
  (Duration 0) unless `TheatreAttendanceFlag = 1`; triage/critical-care/
  ventilation 0; `CostBucketOverheads` all 0.000; overheads live in
  `CostTypeDetails`.
- **Supplies is parked in the `Other` bucket** (Gen2 removed Supplies).
  Open DOH question whether dental clinics should use `Dental` instead —
  MySmyle uses `Dental`; the amounts mirror-match. One-line change in
  `BUCKET_MAP` if DOH rules `Dental`.
- Zero-total claims are blocked by DOH: exclude and log them.
- `DispositionFlag`: `PTE_SUBMIT` for the test environment, `PRODUCTION`
  for live. Transport: SOAP `UploadTransaction`, base64;
  PTE `https://shafafiyapte.doh.gov.ae/v3/webservices.asmx` (username +
  `test`), production `https://shafafiya.doh.gov.ae/v3/webservices.asmx`.

## Definition of done

Twelve XMLs, each: XSD PASS, month total = ledger to the fils, claim count
reconciled to the workbook, < 5 MB. Deliver workbook + zip + a verification
table (per month: claims, size, total, XSD result). Commit and push.
