# Clinical Costing Master Reference — Vision Dental Clinic L.L.C (MF4159)

The single source of truth for the Abu Dhabi DOH Clinical Cost Data Collection
(ADCCDC) work: methodology, locked decisions, conventions, data fixes, audit
standard and history. Maintained from the FY2025 submission project
(completed 2026). Owner: the Finance Director (Dr. Hussein Tahoun).

> **Status (10 Oct 2026).** The FY2025 XML set was submitted and **rejected by
> DOH**. The rejection files and DOH comments are being worked through with the
> procedure in §11 and `docs/REJECTION_PLAYBOOK.md`. This repository was rebuilt
> from this reference on 10 Oct 2026: every engine in §8 is implemented and
> proven end to end on synthetic data (`tests/run_pipeline.sh`), but the
> FY2025 inputs, the FINAL workbook, the twelve submitted XMLs and the official
> XSD are **not** in the repository (patient-level data stays out of GitHub;
> see §12 for what to upload).

---

## 1. Facility and submission profile

| Item | Value |
|---|---|
| Facility | Vision Dental Clinic L.L.C, Abu Dhabi |
| DOH licence / all XML IDs | **MF4159** (Facility, Provider, Sender) |
| Submission | FY2025 ADCCDC, monthly Cost.Submission XML via Shafafiya |
| Standard | DOH Clinical Costing Standard **DOH/ST/HPS/CC/V1/2024** |
| DOH contact | dohclinicalcosting@doh.gov.ae (Annexure report + policy questions) |
| Data suppliers | Cherie (activity + CPT allocation table), Arjun/accounting (monthly expenditure ledger) |

The DOH standard: six stages (define outputs → cost ledger → assign to cost
centres → allocate overheads → cost patient products → reconcile & submit);
14 cost types (Appendix 1); 15 cost buckets (Appendix 2); overhead allocation
statistics (Appendix 3); patient products (Appendix 4). Clause 3.6.2.1.1
requires reconciliation to audited financials; the Annexure reconciliation
workbook is emailed to DOH.

## 2. The locked costing method (scale-only)

One activity row = one unit of one CPT code inside one claim. Month comes
from `StartDateTime`.

- The allocation table gives each CPT code a true **Cost/Unit** and a
  **17-column % split** summing to 100:
  `Anaesthesia, Imaging, Physician, OP, OtherLab, Prosthesis, Supplies,
  SWNurs, SWDoc, SWAllied, SWNonClin, MS, Hotel, GS, Depreciation, OHF, OHC`.
- Per month: modelled cost = Σ Cost/Unit; **factor = ledger expenditure ÷
  modelled cost**; every cell = Cost/Unit × factor × column %.
- Rounding: 3 dp per cell; each column's rounded sum is forced to the rounded
  column target; the month total is forced to `round(GL, 3)` **exactly**
  (largest cell absorbs the residual). The engine works in integer fils so
  the tie is exact by construction, never by float luck.
- **No golden-percentage nudging** — the user ruled the earlier "nudge to a
  golden mix" idea wrong. Scale only; the per-code split decides the mix.
- Factors have ranged roughly 0.7–1.7 by month; that variance is accepted.

Excel ↔ engine column map (TMAP): `Anaesthesia→CostBucketDirectAnaesthesia`,
`Imaging→CostBucketDirectImaging`, `Physician→CostBucketDirectPhysician`,
`OP→CostBucketDirectOP`, `OtherLab→CostBucketDirectLaboratory`,
`Prosthesis→CostBucketDirectProsthesis`, `Supplies→CostBucketDirectSupplies`,
and `SWNurs…OHC→CostTypeDetails*`. All other cost columns are written 0.0;
`total cost` = row sum. TMAP lives in `scripts/costsub_schema.py`.

**Consequence to keep in mind (see §11, hypothesis H1):** because the 17 columns
together sum to 100 %, in our files *direct buckets + cost types = TotalCost*.
The DOH model treats buckets and cost types as two views of the same money,
each summing to the total on its own. MySmyle's files carry the same shape as
ours, so this was accepted as the platform convention; it is nonetheless the
first thing to test against the rejection message.

## 3. Allocation-table fixes (locked)

Applied by `load_alloc()` in `scripts/build2_lib.py` and written to the
workbook's Merge Log:

- Duplicate code 14611 @ 378.00 re-keyed as **14612** (user-confirmed typo;
  14611 = 330.75).
- Identical duplicate rows for 12701 and 29301 de-duplicated.
- Five rows rescaled to exactly 100%: 01901, 02504, 66301, 72111, 72119.
- Orphan map (codes missing from the table priced from siblings):
  `12702→12701, 29302→29301, 42321→42311, 27113→27111`.
- Code **99555** (+E additional materials) priced 0; blank code `00000`
  carried at zero cost and logged.
- Earlier v1-era finding, fixed at source by Cherie's updated table: rows with
  zero overheads (and one missing Depreciation) are an error — every code must
  carry its overhead share. The loader warns if it sees one again.
- Any *other* duplicate or non-100 % row is a WARNING in the Merge Log, kept
  as first-row-wins / rescaled, and must be raised with the Finance Director.

## 4. Data hygiene rules (locked)

- Strip trailing `B` from encounter serials.
- Blank `Speciality` → clinician's modal speciality learned from the full
  source files (notably GD11525 → Oral Surgeon; GD5633, GD22503, GD45264,
  GD36537 → General Dentist). The known values are pinned in
  `KNOWN_MODAL_SPECIALTY`; anything else is learned from the file.
- `Encounter.ID` in XML = the claim's **first line's** activity serial;
  `ActivityClinicianID` = `Doctor No`; `PatientID` = `FileNo.`;
  `Claim.ID` = the claim reference (e.g. `1747MF4159-5a058f`).
- Lines whose `StartDateTime` falls outside the financial year are excluded
  and logged, never silently costed into a neighbouring month.

## 5. Shafafiya Cost.Submission — two schema generations

**Generation 1** (FY2024 window): 15 buckets **including Supplies**; no
ActivityClinicianID/ServiceCode; 15 cost types. This matches the 73-column
Excel template.

**Generation 2** (current XSD, used for FY2025):
- Buckets (16): Allied, Anaesthesia, **Dental**, ED, **Home**, ICU, Imaging,
  Laboratory, Physician, OP, OR, Other, Pharmacy, Prosthesis, SPS, Ward —
  **no Supplies**.
- Encounter adds required `ActivityClinicianID` and `ServiceCode` (we send it
  empty).
- Cost types (16, in order): SWNurs, SWDoc, SWAllied, SWNonClin, Laboratory,
  Imaging, Pharmacy, Prostheses, MS, Hotel, GS, Depreciation, OHF,
  **OHDepreciation**, OHC, NABMs.
- **Supplies is parked in `Other`** (clearly flagged). OPEN QUESTION with DOH:
  should dental clinics use `Dental`? MySmyle's platform uses `Dental`; the
  amounts mirror-match ours in `Other`. If DOH says Dental: regenerate with
  `--supplies-bucket Dental` (`scripts/make_xml_final.py`); nothing else changes.

**Hard portal/schema rules:** ReceiverID `HAAD` only; dates `dd/mm/yyyy HH:MM`
GMT+4; AED 3 dp; ≤ 5 MB per file (split by month); one `<Claim>` per claim ID;
zero-total encounters blocked (exclude + log); negatives blocked; **Principal
diagnosis only**; Procedure only when `TheatreAttendanceFlag=1` (else empty,
Duration 0); CostBucketOverheads all 0.000 (overheads live in
CostTypeDetails); DispositionFlag `PTE_SUBMIT`/`PTE_VALIDATE_ONLY` (test) vs
`PRODUCTION` (live). Transport: SOAP `UploadTransaction`, base64; PTE
`https://shafafiyapte.doh.gov.ae/v3/webservices.asmx` (username + `test`
suffix), production `https://shafafiya.doh.gov.ae/v3/webservices.asmx`.
Resubmission only on DOH request (a rejection *is* that request).
Official XSD + technical PDFs: `docs/shafafiya/` (the served XSD is truncated;
validate against `CostSubmission_repaired.xsd`). `scripts/validate_xml.py`
checks every rule in this paragraph on any file, ours or a platform's.

**Where the schema lives in code.** Every element name, the element order,
the bucket and cost-type lists and the enumerations are in one file,
`scripts/costsub_schema.py`. The element order there was rebuilt from this
reference on 10 Oct 2026 and **must be confirmed against the official XSD**
once it is back in `docs/shafafiya/`; aligning it is a one-file change.

## 6. The FY2025 build history (what governs)

1. **v1** → superseded (provisional activity data).
2. **v2** → Cherie's corrected activity (one sheet), updated allocation table,
   expenditure file totalling 6,728,853.39.
3. **Three-source forensic audit** (MySmyle vs Calibrate vs ClaimScanner, no
   external truth): **ClaimScanner = payment truth** (internal money equations
   perfect); **MySmyle = line-amount accuracy** (0 diffs vs adjudication) but
   missing 25 real claims; **Calibrate = completeness** but 15 claims
   understated by 5,444.60. Uncollected revenue 256,737.43, of which 107
   never-resubmitted rejections worth 29,736.14.
4. **v3 "best of three"**: Calibrate base + those 15 claims replaced with
   MySmyle lines + 23 MySmyle-only records, `Source` provenance column,
   SelfPay included, full costing rebuild.
5. **The Finance Director's manual edit of v3 is the accepted master**
   (24,074 lines / 9,738 claims): claim `1296MF4159-aa1544` deleted (his
   final decision — do not raise again), zero-value claims `1680MF4159-6d63f3`
   and `727MF4159-82438e` deleted, one 99555 line removed from
   `2336MF4159-4d736e`.
6. **Updated ledger (Arjun)**: FY2025 total **6,453,401.98** —
   Jan 511,302.30 · Feb 532,341.65 · Mar 606,868.87 · Apr 518,123.67 ·
   May 488,066.69 · Jun 616,494.09 · Jul 480,166.82 · Aug 526,990.10 ·
   Sep 509,542.05 · Oct 546,725.98 · Nov 493,746.03 · Dec 623,033.73
   (`scripts/gl_new.py`).
7. **FINAL build**: `output/Clinical_Costing_2025_MF4159_FINAL.xlsx` +
   `output/xml_final/` (twelve XMLs, all XSD PASS, every month ties to the
   fils, no zero-cost claims remained) + zip.
8. **Submission to DOH → REJECTED** (2026). Rejection files and comments to be
   digested per §11. The FINAL build stays the accepted master until the
   rejection shows a data defect; a schema or packaging defect is fixed in the
   generator and the same master is regenerated.

## 7. The standard audit (run it the same way every time)

`python3 scripts/audit_mysmyle_xml.py <my_xml> <their_xml>`: file sizes; XSD
both; header fields; claim-set equality both ways; every identity/meta field
per claim (IDs, clinician, patient, dates, types, specialty, diagnosis,
procedure, flags); per-cell money diffs > 0.0005 with counts and nets per
field; zero-total and negative claims; largest cell diffs. Verdict in plain
words, money balance first.

**Converged result (Jan + Dec final audits):** identical claim sets, zero
metadata differences, totals within < 0.1 AED/month. Only two residual
differences: (a) Supplies in `Dental` (MySmyle) vs `Other` (ours) — perfect
mirror, no money created or lost; (b) our files tie the ledger to the fils,
MySmyle's are a few fils short (per-cell rounding without reconciliation).

## 8. Repository map

| Path | Purpose |
|---|---|
| `scripts/costsub_schema.py` | **Single schema map**: Gen2 element names/order, buckets, cost types, TMAP, Supplies switch |
| `scripts/build2_lib.py` | Constants, allocation loader with locked fixes, orphan map, activity reader with hygiene rules, exact-fils rounding |
| `scripts/gl_new.py` | The governing monthly expenditure figures |
| `scripts/build4_final.py` | FINAL engine: accepted lines × ledger, scale-only; writes Sheet1 / Reconciliation / Monthly Factors / Merge Log |
| `scripts/make_xml_final.py` | Gen2 XML generator + inline XSD validation + verification table; `--disposition`, `--supplies-bucket`, `--zip` |
| `scripts/costsub_reader.py` | Tolerant reader for any Cost.Submission XML (flat or nested money elements) |
| `scripts/validate_xml.py` | Pre-flight: every portal rule in §5 + ledger tie + cross-foot report (`--strict-totals`) |
| `scripts/audit_mysmyle_xml.py` | Two-file XML audit (the standard, §7) |
| `scripts/triage_rejection.py` | **Rejection triage**: digests DOH response files into message → count → claims → cause → fix |
| `scripts/submit_shafafiya.py` | SOAP `UploadTransaction` client, dry-run by default, refuses PTE/PRODUCTION mismatches |
| `tests/make_synthetic.py`, `tests/run_pipeline.sh` | Synthetic data and the end-to-end proof |
| `docs/REJECTION_PLAYBOOK.md` | The rejection procedure and hypothesis list |
| `docs/activity-template-columns.md` | Accepted column headings for the activity master and allocation table |
| `docs/shafafiya/` | Official XSDs (+ repaired copy), technical PDFs, validated sample XML — **to be re-uploaded** |
| `docs/cost-types-plain-english.md`, `docs/The_Fifteen_Boxes.docx` | Layman's cost-type guides (previous project; not yet in this repo) |
| `output/Clinical_Costing_2025_MF4159_FINAL.xlsx` | Final costed workbook (git-ignored) |
| `output/xml_final/`, `output/CostSubmission_MF4159_2025_XML_*.zip` | The twelve validated XMLs (git-ignored) |
| `output/Claims_Audit_2025_Three_Sources.xlsx` | Forensic audit evidence workbook (git-ignored) |

`scripts/build3_best_of_three.py` (the historical v3 merge) is not reproduced;
its result is the accepted master of §6 item 5.

## 9. Open items

1. **DOH rejection of the FY2025 set — in progress** (§11).
2. **DOH ruling pending:** Supplies under `Dental` vs `Other` (email sent to
   dohclinicalcosting@doh.gov.ae). If "Dental": `--supplies-bucket Dental`,
   regenerate 12 files.
3. Live files must carry `PRODUCTION`; `scripts/submit_shafafiya.py` refuses a
   test-flagged file on the live endpoint and vice versa.
4. DOH Annexure reconciliation report can be filled from the FINAL build on
   request.
5. Platform/ops follow-ups from the forensic audit: restore the 25 missing
   claims in MySmyle, fix Calibrate's 15 understated claims, payer-mapping
   reconciliation, work the 107 never-resubmitted rejections (29,736.14).
6. Official XSD and technical PDFs to be placed back in `docs/shafafiya/`;
   then confirm `scripts/costsub_schema.py` against them.

## 10. Working etiquette with the Finance Director

Digest every upload fully and report back (figures, counts, deltas) before
building; ask at most 4 on-screen MCQs and only when genuinely ambiguous;
wait for the explicit "go". Plain language (he is often on a phone; the coder
team is non-technical). Locked decisions are final — flag new contrary
evidence once, then accept the ruling. Every delivery: files + verification
table + commit/push. Patient-level data never goes to GitHub (`input/`,
`output/`, spreadsheets and zips are git-ignored).

## 11. Rejection handling (standard procedure)

Full procedure: `docs/REJECTION_PLAYBOOK.md`. In short:

1. **Collect**: the DOH/Shafafiya response files (portal response XML, email,
   Excel/CSV error list, screenshots transcribed), the exact files that were
   sent (the zip), which environment and account sent them, and the date.
2. **Digest** with `scripts/triage_rejection.py <files> --xml <sent xmls>`:
   one table of distinct messages, counts, claims affected, month, suspected
   cause, fix. Report it to the Finance Director in plain words before
   touching anything.
3. **Classify** each message as *packaging* (disposition, environment,
   credentials, file size), *schema* (XSD element/order/format), *reference
   data* (clinician licence, specialty list, diagnosis codes, patient ID
   type), or *money* (cross-foot, zero/negative, ledger tie). Only a *money*
   or *reference-data* defect touches the accepted master; the rest is fixed
   in the generator and regenerated from the same master.
4. **Fix once, in the engine**, so the same cause cannot recur; re-run
   `scripts/validate_xml.py` on all twelve files; audit one month against the
   previously sent file with `scripts/audit_mysmyle_xml.py` to show exactly
   what changed and that money was neither created nor lost.
5. **Resubmit** the full set (or the months DOH named) with `PRODUCTION`;
   record the ruling here and in `docs/REJECTION_PLAYBOOK.md`.

**Hypotheses to test first against the FY2025 rejection** (ordered by
likelihood; see the playbook for the test and the fix of each):
H1 cross-foot (buckets ≠ types ≠ total); H2 `PTE_SUBMIT` sent to production;
H3 Supplies in `Other` vs `Dental`; H4 element order/names vs the XSD;
H5 `ServiceCode` empty but required; H6 `ActivityClinicianID` not a DOH
licence; H7 Specialty text not in the DOH code list; H8 duplicate claim or
encounter IDs across the twelve files.

## 12. What this repository still needs from the previous project

Upload (they are git-ignored, so they live in `input/` and `output/` locally):
the accepted activity master (24,074 lines / 9,738 claims), Cherie's
allocation table, the FINAL workbook, the twelve submitted XMLs (zip), the
official XSD + technical PDFs for `docs/shafafiya/`, and the rejection files.
With those in place the whole chain in §8 runs unchanged.
