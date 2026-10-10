# Rejection Playbook — DOH / Shafafiya Cost.Submission

How a rejected ADCCDC file is handled, start to finish, so that every rejection
is worked the same way and the fix lands once, in the engine.

## 0. What to collect before any analysis

| Item | Why |
|---|---|
| The response DOH/Shafafiya returned (portal response XML, email text, Excel/CSV error list, screenshots transcribed to text) | The triage tool reads all of these |
| The exact files that were sent (the zip) | To say which month each rejected claim sits in and to diff against the fix |
| Environment and account used (PTE vs production; username) | Hypothesis H2 |
| Date/time of upload and of the rejection | Period and sequencing questions |
| Any DOH comments or phone notes | Often the only statement of the real rule |

## 1. Digest

```bash
python3 scripts/triage_rejection.py input/rejection/* --xml output/xml_final --out output/rejection_triage.xlsx
```

Produces one table: normalised message → occurrences → claims affected →
elements named → error codes → months → suspected cause → fix in pipeline, plus a
per-claim sheet. Read it back to the Finance Director in plain words first.
Unknown message types are listed as such; once understood, add them to
`CATALOGUE` in `scripts/triage_rejection.py` so the next run classifies them.

## 2. Classify

| Class | Examples | Touches the accepted master? | Fix lives in |
|---|---|---|---|
| Packaging | test disposition on live, wrong endpoint, credentials, > 5 MB, RecordCount | No | `make_xml_final.py --disposition`, `submit_shafafiya.py`, split month |
| Schema | element missing/unexpected/out of order, date or money format | No | `scripts/costsub_schema.py` (names/order), regenerate |
| Reference data | clinician licence, specialty list, diagnosis code, patient-ID type | Columns only | activity master mapping columns, regenerate |
| Money | cross-foot rule, zero/negative, ledger tie, Supplies bucket | Method or bucket switch | Finance Director ruling first; then engine |

## 3. Hypotheses for the FY2025 rejection (test in this order)

| # | Hypothesis | Test | Fix if confirmed |
|---|---|---|---|
| H1 | DOH validates TotalCost = Σ direct buckets = Σ cost types. Our locked 17-column split gives buckets + types = total, so each view is roughly half the total. | `scripts/validate_xml.py --strict-totals` on a sent file reports "buckets==types in 0/n". Message would mention total/sum/cost type/bucket. | Finance Director ruling. Mechanically: express the split as two 100 % views (direct 7 columns rescaled to 100 % for buckets; all 17 columns' *cost-type nature* mapped to the 16 types at 100 %) so both views equal the total. The scale-only method and ledger tie are unchanged; only the presentation doubles. |
| H2 | Files sent to production carried `PTE_SUBMIT`. | Header of the sent files. | `--disposition PRODUCTION`, resend. |
| H3 | Supplies must sit in `Dental` (or must not sit in `Other`). | Message names `Other`/`Dental`; DOH email reply. | `--supplies-bucket Dental`. |
| H4 | Element names/order differ from the current XSD. | Drop the official XSD in `docs/shafafiya/`, run `validate_xml.py`; message says "element … not expected". | Align `ENCOUNTER_ORDER`/names in `costsub_schema.py`. |
| H5 | `ServiceCode` required non-empty (e.g. CPT of the encounter). | Message names ServiceCode. | Fill from the claim's first line's code (one line in `make_xml_final.py`). |
| H6 | `ActivityClinicianID` must be the DOH clinician licence, not the staff number. | Message names clinician/licence; compare Doctor No format to GDxxxxx. | Add a licence column to the activity master; map. |
| H7 | `Specialty` must be a DOH code, not free text. | Message names Specialty; check the Shafafiya code list in the technical PDFs. | Map text → code in the activity reader. |
| H8 | A claim or encounter serial appears in more than one monthly file, or was already live from MySmyle's upload. | Message says duplicate/already submitted; `validate_xml.py` warns on repeated encounter IDs; compare the twelve claim sets. | Keep each claim in its StartDateTime month; agree with DOH which upload is live before resending. |

## 4. Fix, prove, resubmit

1. Fix in the engine (never by hand-editing XML).
2. Regenerate all twelve files with `PRODUCTION`; `validate_xml.py` all twelve
   with `--expect-disposition PRODUCTION`.
3. `audit_mysmyle_xml.py <new month> <sent month>` for at least one month:
   identical claim sets, money balance to the fils, only the intended cells
   moved.
4. Verification table in the delivery (month, claims, size, total, ledger tie,
   XSD result).
5. Resubmit the set DOH asked for (`submit_shafafiya.py upload … --env prod
   --send`, credentials from the environment) and save the response files to
   `output/soap/`.
6. Record the ruling in `docs/CLINICAL_COSTING.md` §6/§9/§11 and add the
   message pattern to the triage catalogue.

## 5. Log of rejections

| Date | Set | Messages (summary) | Root cause | Fix | Resubmitted |
|---|---|---|---|---|---|
| 2026 (to be filled) | FY2025, 12 files | awaiting the rejection files | | | |
