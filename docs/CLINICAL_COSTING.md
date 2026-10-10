# Clinical Costing Master Reference — Vision Dental Clinic L.L.C (MF4159)

The single source of truth for the Abu Dhabi DOH Clinical Cost Data Collection
(ADCCDC) work: methodology, locked decisions, conventions, data fixes, audit
standard and history. Maintained from the FY2025 submission project
(completed 2026). Owner: the Finance Director (Dr. Hussein Tahoun).

> **Status (10 Oct 2026).** The FY2025 (SY02) set was submitted on 29 Aug 2026
> and **rejected by DOH on 9 Oct 2026** on one matrix: *Cost Bucket / Cost Type
> Match: No* (score −5). One warning: *Overhead Cost Ratio below 20%*. Root
> cause, rulings and corrective action: §11. The twelve files DOH evaluated are
> in `input/source/mysmyle_zip/`, the letter in `input/rejection/`, the DOH
> Standard v2.1 (Jan 2026) and Quick Start Guide in `docs/shafafiya/`.
> Waiting on: Arjun's ledger mapped to the sixteen cost types
> (`input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx`).

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

**Why this was rejected (SY02, §11).** The 17 columns sum to 100 % *together*,
so each dirham was placed once, in either a bucket or a cost type. The DOH
Standard treats cost buckets (aggregated final direct cost centres: *where*)
and cost types (aggregated ledger accounts: *what*) as two complete views of
the same money; each must equal the full claim cost. From 10 Oct 2026 the
17-column split still decides the **bucket mix** and the claim total, and the
cost-type view is produced separately (§11 corrective action).

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
- **Supplies is reported in `Dental`** (ruling 10 Oct 2026: that is what DOH
  evaluated without objection in the SY02 files). `--supplies-bucket` remains
  available if DOH ever rules otherwise.

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
`scripts/costsub_schema.py`, **confirmed on 10 Oct 2026 against the twelve
files Shafafiya accepted**: money groups are nested
(`<CostBucketDirect><Dental>…`), the encounter uses `Start`/`End`, carries
`PatientID`, `EDTriageLevel`, `CriticalCareTime`, `VentilationTime`, and
`Procedure` has a `Type`; there is no `TotalCost` element. Overheads: the
Standard places allocated overhead in the bucket's overhead component, so
`CostBucketOverheads` carries the OHF / OHDepreciation / OHC share of each
bucket and equals the overhead cost types (ruling 10 Oct 2026; replaces the
earlier "all 0.000" convention).

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
8. **Submission to DOH (MySmyle production files, 29 Aug 2026) → REJECTED
   9 Oct 2026** on Cost Bucket / Cost Type Match; warning on overhead ratio.
   The claim set (9,738 claims, every month = ledger within fils) is
   untouched; only the presentation of each claim's cost changes (§11).
9. **SY02 resubmission (in preparation)**: the twelve sent files re-expressed
   by `scripts/rebalance_two_view.py` once Arjun's ledger-by-cost-type is in.

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
| `scripts/rebalance_two_view.py` | **SY02 corrective engine**: ledger-calibrated two-margin fit; writes the resubmission XMLs |
| `scripts/fix_crossfoot.py` | Nature re-expression (fallback without the ledger map; also used by the generator) |
| `scripts/make_ledger_template.py` | Builds Arjun's ledger-by-cost-type template (+ "Current implied" reference sheet) |
| `scripts/fill_ledger_map_from_financials.py` | Fills the template from Arjun's financial workbook (Mapping, Questions, Consumables detail sheets) — the governing map |
| `scripts/cogs_from_movement.py` | Splits a month's consumables into cost types from the stock-movement sheet of the monthly Financial Report |
| `input/source/Financial_Report_Jan_2025_1.xlsx` | January 2025 monthly financial report (stock movement by item category) |
| `input/Ledger_by_Cost_Type_FY2025.xlsx` | The filled, evidence-based ledger-by-cost-type map (pending confirmation of its Questions sheet) |
| `input/source/Financial_2025_Arjun_Updated_2.xlsx` | Arjun's FY2025 financial workbook (TB, P&L, Monthly, Salary, Consumption, cash ledger by cost centre) |
| `scripts/propose_ledger_map.py` | Earlier common-sense proposal (superseded) |
| `output/xml_resubmission_PROPOSED/` | Twelve two-view files from the proposed map (proposal, not yet the resubmission) |
| `input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx` | The template to be filled and returned as `input/Ledger_by_Cost_Type_FY2025.xlsx` |
| `input/source/mysmyle_zip/` | The twelve files DOH evaluated (MySmyle, PRODUCTION, 29 Aug 2026) |
| `input/rejection/` | DOH SY02 feedback letter (PDF + text) |
| `docs/shafafiya/Abu_Dhabi_Clinical_Costing_Standard_v2_1_2026-01-27.pdf` | The governing Standard (and `.txt`); Quick Start Guide alongside |
| `scripts/submit_shafafiya.py` | SOAP `UploadTransaction` client, dry-run by default, refuses PTE/PRODUCTION mismatches |
| `tests/make_synthetic.py`, `tests/run_pipeline.sh` | Synthetic data and the end-to-end proof |
| `docs/REJECTION_PLAYBOOK.md` | The rejection procedure and hypothesis list |
| `docs/activity-template-columns.md` | Accepted column headings for the activity master and allocation table |
| `docs/shafafiya/` | Standard v2.1, Quick Start Guide (in); official XSD — **to be re-uploaded** |
| `docs/cost-types-plain-english.md`, `docs/The_Fifteen_Boxes.docx` | Layman's cost-type guides (previous project; not yet in this repo) |
| `output/Clinical_Costing_2025_MF4159_FINAL.xlsx` | Final costed workbook |
| `output/xml_final/`, `output/CostSubmission_MF4159_2025_XML_*.zip` | The twelve validated XMLs |
| `output/Claims_Audit_2025_Three_Sources.xlsx` | Forensic audit evidence workbook |
| `input/rejection/`, `input/source/` | DOH rejection files as received; source uploads (activity master, allocation table, sent XMLs) |

`scripts/build3_best_of_three.py` (the historical v3 merge) is not reproduced;
its result is the accepted master of §6 item 5.

## 9. Open items

1. **SY02 resubmission** — waiting on Arjun's ledger mapped to the sixteen cost
   types by month (`input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx`); then run
   the engine, validate, audit, update the reconciliation report, resubmit
   with `PRODUCTION` (§11).
2. **Overhead ratio 4.2 %** (DOH warning, accepted for SY02). The Finance
   Director ruled on 10 Oct 2026 not to re-engineer the allocation now. Arjun's
   ledger mapping will show the true ratio; if it lands in 20–50 % the
   warning clears itself in the resubmission, otherwise it is an SY03 item.
3. The reconciliation report submitted with SY02 must be re-issued with the
   two-view totals (not yet uploaded to this repository).
4. DOH Annexure reconciliation report can be filled from the FINAL build on
   request.
5. Platform/ops follow-ups from the forensic audit: restore the 25 missing
   claims in MySmyle, fix Calibrate's 15 understated claims, payer-mapping
   reconciliation, work the 107 never-resubmitted rejections (29,736.14).
6. Official XSD to be placed in `docs/shafafiya/` (validation currently reports
   `NO XSD`; the structure itself is confirmed by the accepted files).
7. 18 encounter serials are shared by two claims each in the sent files
   (hygiene only; DOH did not object).
8. MySmyle: ask whether their generator can adopt the two-view layout for
   SY03, or keep generating from this repository.

## 10. Working etiquette with the Finance Director

Digest every upload fully and report back (figures, counts, deltas) before
building; ask at most 4 on-screen MCQs and only when genuinely ambiguous;
wait for the explicit "go". Plain language (he is often on a phone; the coder
team is non-technical). Locked decisions are final — flag new contrary
evidence once, then accept the ruling. Every delivery: files + verification
table + commit/push. The repository is private (ruling 10 Oct 2026): inputs,
outputs, submitted XMLs and rejection files are committed so comparisons are
reproducible from GitHub.

## 11. Rejection handling — SY02 case file and standard procedure

### 11.1 SY02 rejection (letter dated 9 Oct 2026)

| Matrix | Criteria | Score | Tolerance |
|---|---|---|---|
| Cost Bucket / Cost Type Match | No | −5 | **Reject** |
| Overhead Cost Ratio | Below 20 % | 0 | Warning |
| All nine other matrices | compliant | 10 / 5 | Acceptable |

**Root cause (verified against Standard v2.1).** Cost types are aggregations of
ledger expense accounts (*what*, Appendix 1); cost buckets are aggregations of
final direct cost centres (*where*, Appendix 2). Every dirham carries both
labels, so on every claim Σ buckets = Σ cost types = claim cost. Our
17-column allocation split placed each dirham once, in a bucket *or* a cost
type: over FY2025 the bucket view held 2,290,172.29 (35.5 %) and the cost-type
view 4,163,229.87 (64.5 %); the two together equal the ledger 6,453,401.98.
Only 11 of 9,738 claims matched, by coincidence. A presentation defect, not a
costing defect: claim set, totals, ledger tie and identity fields are all
sound.

**Warning root cause.** OHF + OHC in the files = 272,572.57 = 4.2 % of total;
DOH expects 20 % to under 50 %. The allocation table gives small overhead
percentages and books non-clinical staff and depreciation as direct types.
The true ratio is whatever Arjun's ledger says once cost centres are
classified direct/overhead (Stage 2).

**Rulings (Finance Director, 10 Oct 2026).**
1. The 17 columns are one list (7 buckets + 10 cost types) — confirmed.
2. Nature of the bucket columns in the cost-type view: Supplies→MS,
   Laboratory→Laboratory, Prosthesis→Prostheses, Imaging→Imaging,
   Anaesthesia→Pharmacy, OP→SWNonClin, Physician→SWDoc.
3. Home of the cost-type columns in the bucket view: `Dental` (the clinic's one
   direct cost centre); overhead types in `CostBucketOverheads/Dental`.
4. Method: **ledger-calibrated two-margin fit** — Arjun maps the FY2025 ledger
   to the sixteen cost types by month (cost-type anchor); the costed claims
   give the bucket anchor; a bucket × cost-type matrix is balanced to both
   (RAS) from a nature seed; each claim's cost types = its bucket amounts ×
   the fitted bucket profiles. Engine: `scripts/rebalance_two_view.py`.
   Fallback without the ledger map: `scripts/fix_crossfoot.py` (nature
   re-expression only).
5. Overhead ratio is not re-engineered for SY02.

**Corrective action sequence.**
1. Arjun fills `input/Ledger_by_Cost_Type_FY2025_TEMPLATE.xlsx` (each month
   must reconcile to the locked ledger; returned as
   `input/Ledger_by_Cost_Type_FY2025.xlsx`).
2. `python3 scripts/rebalance_two_view.py --indir input/source/mysmyle_zip
   --ledger input/Ledger_by_Cost_Type_FY2025.xlsx --outdir
   output/xml_resubmission --zip`.
3. `python3 scripts/validate_xml.py output/xml_resubmission/*.xml
   --expect-disposition PRODUCTION` — every claim must show buckets = types,
   every month tie to the ledger.
4. `python3 scripts/audit_mysmyle_xml.py <new month> <sent month>` — identical
   claim sets, zero identity differences, money moved between cells only.
5. Re-issue the reconciliation report with the two-view totals; CFO sign-off.
6. Upload the twelve files (PRODUCTION) via Shafafiya; record the outcome in
   `docs/REJECTION_PLAYBOOK.md` §5.

**Evidence-based ledger map (10 Oct 2026, from Arjun's FY2025 financial
workbook `input/source/Financial_2025_Arjun_Updated_2.xlsx`).** Built by
`scripts/fill_ledger_map_from_financials.py` → `input/Ledger_by_Cost_Type_FY2025.xlsx`
(sheets: Cost types by month, Mapping, Questions). Sources: the Monthly
P&L (its Total row is the locked ledger), the Salary sheet by role, the
Consumption sheet, the TB depreciation movement. Result: SWDoc 28.2 %,
MS 12.2 %, OHC 5.7 % (MySmyle EMR/management/IT recharge), OHF 39.0 %,
overhead ratio 45.2 % for the year (monthly 31–58 %). Five points await the
Finance Director's confirmation (Questions sheet), chiefly whether
"Staff Qualification And Education" (889,431.99) is HR/IT overhead or staff
on-costs. January demo run: 831/831 claims match, both views 511,302.300.
Supersedes the proposal below once confirmed.
**Consumables from stock movement (10 Oct 2026).** The annual ledger books all
consumables as one line. The monthly Financial Report's "Movement" sheet splits
it by item category; `scripts/cogs_from_movement.py` maps usage (net of
internal returns, exchanges cancel, external lab from the P&L) to cost types:
implants, screws, abutments, bone graft and implant accessories → Prostheses;
local anaesthetics and sedation gases → Pharmacy; lab blocks/blanks/inks/stone →
Laboratory; the rest → MS; pantry/stationery → OHF. Rulings: CSSD stays direct
(SWAllied / MS); implants incl. bone graft → Prostheses; anaesthetics →
Pharmacy. January verified to the fils (102,127.49): Prostheses 52,796.00,
MS 31,662.77, Laboratory 15,408.95, Pharmacy 1,252.00, OHF 1,007.80. Pass
`--movement <m>=<file>` to `fill_ledger_map_from_financials.py` for each
month whose report is uploaded; other months keep the Consumption-sheet split.
Known limit: the fit spreads a cost type across every claim of a bucket, so
implant cost lands on all Dental-bucket claims; attributing it to implant
procedures only needs the activity master (CPT codes are not in the XML).

**Proposed ledger map (10 Oct 2026, superseded by the evidence-based map above).** Built by
`scripts/propose_ledger_map.py` without the expense lines:
`input/Ledger_by_Cost_Type_FY2025_PROPOSED.xlsx` (Assumptions sheet inside).
Overheads 24.0 % of expenditure (OHF 21.5, OHDepreciation 1.3, OHC 1.2) as a
fixed monthly amount; direct block 76.0 % in the clinic's own proportions;
Laboratory, Prostheses and Imaging cost types equal to the costed buckets.
Run through the engine → `output/xml_resubmission_PROPOSED/` (+ zip):
9,738/9,738 claims match, all twelve months tie in both views to the fils,
overhead bucket = overhead types, monthly overhead ratio 20.7–26.9 %. These
files are a **proposal**: they become the resubmission only after Arjun
confirms or replaces the map from the ledger and the engine is re-run.

### 11.2 Standard procedure

Full procedure: `docs/REJECTION_PLAYBOOK.md`. In short: collect → digest
(`scripts/triage_rejection.py`) → classify (packaging / schema / reference
data / money) → fix once in the engine → regenerate all twelve → validate →
audit one month against the sent file → resubmit with `PRODUCTION` → record
the ruling here.

## 12. What this repository still needs from the previous project

Still to upload: the accepted activity master (24,074 lines / 9,738 claims),
Cherie's allocation table, the FINAL workbook, the official XSD, and the
reconciliation report submitted with SY02. The sent XMLs, the DOH letter, the
Standard v2.1 and the Quick Start Guide are in. With Arjun's ledger map the
SY02 resubmission runs from the sent files alone.
