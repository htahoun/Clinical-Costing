# Clinical Costing Master Reference — Vision Dental Clinic L.L.C (MF4159)

The single source of truth for the Abu Dhabi DOH Clinical Cost Data Collection
(ADCCDC) work: methodology, locked decisions, conventions, data fixes, audit
standard and history. Maintained from the FY2025 submission project
(completed 2026). Owner: the Finance Director (Dr. Hussein Tahoun).

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
  (largest cell absorbs the residual).
- **No golden-percentage nudging** — the user ruled the earlier "nudge to a
  golden mix" idea wrong. Scale only; the per-code split decides the mix.
- Factors have ranged roughly 0.7–1.7 by month; that variance is accepted.

Excel ↔ engine column map (TMAP): `Anaesthesia→CostBucketDirectAnaesthesia`,
`Imaging→CostBucketDirectImaging`, `Physician→CostBucketDirectPhysician`,
`OP→CostBucketDirectOP`, `OtherLab→CostBucketDirectLaboratory`,
`Prosthesis→CostBucketDirectProsthesis`, `Supplies→CostBucketDirectSupplies`,
and `SWNurs…OHC→CostTypeDetails*`. All other cost columns are written 0.0;
`total cost` = row sum.

## 3. Allocation-table fixes (locked)

Applied by `load_alloc()` in `scripts/build2_lib.py`:

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
  carry its overhead share.

## 4. Data hygiene rules (locked)

- Strip trailing `B` from encounter serials.
- Blank `Speciality` → clinician's modal speciality learned from the full
  source files (notably GD11525 → Oral Surgeon; GD5633, GD22503, GD45264,
  GD36537 → General Dentist).
- `Encounter.ID` in XML = the claim's **first line's** activity serial;
  `ActivityClinicianID` = `Doctor No`; `PatientID` = `FileNo.`;
  `Claim.ID` = the claim reference (e.g. `1747MF4159-5a058f`).

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
  amounts mirror-match ours in `Other`. If DOH says Dental: change one line
  (`'CostBucketDirectSupplies':'Other'` → `'Dental'`) in the generator and
  regenerate.

**Hard portal/schema rules:** ReceiverID `HAAD` only; dates `dd/mm/yyyy HH:MM`
GMT+4; AED 3 dp; ≤ 5 MB per file (split by month); one `<Claim>` per claim ID;
zero-total encounters blocked (exclude + log); negatives blocked; **Principal
diagnosis only**; Procedure only when `TheatreAttendanceFlag=1` (else empty,
Duration 0); CostBucketOverheads all 0.000 (overheads live in
CostTypeDetails); DispositionFlag `PTE_SUBMIT`/`PTE_VALIDATE_ONLY` (test) vs
`PRODUCTION` (live). Transport: SOAP `UploadTransaction`, base64; PTE
`https://shafafiyapte.doh.gov.ae/v3/webservices.asmx` (username + `test`
suffix), production `https://shafafiya.doh.gov.ae/v3/webservices.asmx`.
Resubmission only on DOH request. Official XSD + technical PDFs:
`docs/shafafiya/` (the served XSD is truncated; validate against
`CostSubmission_repaired.xsd`).

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

## 7. The standard audit (run it the same way every time)

`python3 scripts/audit_mysmyle_xml.py <my_xml> <their_xml>` (run from
`docs/`): file sizes; XSD both; header fields; claim-set equality both ways;
every identity/meta field per claim (IDs, clinician, patient, dates, types,
specialty, diagnosis, procedure, flags); per-cell money diffs > 0.0005 with
counts and nets per field; zero-total and negative claims; largest cell
diffs. Verdict in plain words, money balance first.

**Converged result (Jan + Dec final audits):** identical claim sets, zero
metadata differences, totals within < 0.1 AED/month. Only two residual
differences: (a) Supplies in `Dental` (MySmyle) vs `Other` (ours) — perfect
mirror, no money created or lost; (b) our files tie the ledger to the fils,
MySmyle's are a few fils short (per-cell rounding without reconciliation).

## 8. Repository map

| Path | Purpose |
|---|---|
| `scripts/build2_lib.py` | Constants, GL loader, allocation loader with locked fixes, orphan map |
| `scripts/build3_best_of_three.py` | The v3 merge + costing build (historical) |
| `scripts/build4_final.py` | FINAL engine: frozen accepted lines × updated ledger, scale-only |
| `scripts/gl_new.py` | The governing monthly expenditure figures |
| `scripts/make_xml_final.py` | Gen2 XML generator + inline XSD validation |
| `scripts/audit_mysmyle_xml.py` | Two-file XML audit (the standard) |
| `docs/shafafiya/` | Official XSDs (+ repaired copy), technical PDFs, validated sample XML, README |
| `docs/cost-types-plain-english.md`, `docs/The_Fifteen_Boxes.docx` | Layman's cost-type guides |
| `output/Clinical_Costing_2025_MF4159_FINAL.xlsx` | Final costed workbook (Sheet1 + Reconciliation + Monthly Factors + Merge Log) |
| `output/xml_final/`, `output/CostSubmission_MF4159_2025_XML_FINAL.zip` | The twelve validated XMLs |
| `output/Claims_Audit_2025_Three_Sources.xlsx` | Forensic audit evidence workbook |

## 9. Open items

1. **DOH ruling pending:** Supplies under `Dental` vs `Other` (email sent to
   dohclinicalcosting@doh.gov.ae). If "Dental": one-line change, regenerate 12
   files.
2. Our XML set carries `PTE_SUBMIT`; regenerate with `PRODUCTION` for live
   (MySmyle already produced live files).
3. DOH Annexure reconciliation report can be filled from the FINAL build on
   request.
4. Platform/ops follow-ups from the forensic audit: restore the 25 missing
   claims in MySmyle, fix Calibrate's 15 understated claims, payer-mapping
   reconciliation, work the 107 never-resubmitted rejections (29,736.14).

## 10. Working etiquette with the Finance Director

Digest every upload fully and report back (figures, counts, deltas) before
building; ask at most 4 on-screen MCQs and only when genuinely ambiguous;
wait for the explicit "go". Plain language (he is often on a phone; the coder
team is non-technical). Locked decisions are final — flag new contrary
evidence once, then accept the ruling. Every delivery: files + verification
table + commit/push.
