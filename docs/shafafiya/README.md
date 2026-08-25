# Shafafiya Cost Submission — official sources

Downloaded 25 Aug 2026 from the DOH Shafafiya dictionary
(https://www.doh.gov.ae/en/Shafafiya/dictionary).

| File | What it is |
|---|---|
| `Cost-Data-Collection-Technical.pdf` | DOH Cost Data Collection Technical Document v1.12 (July 2025): schema sample, field definitions, portal validations, process flow |
| `CostSubmission.xsd` | Official "Cost Submission (NEW)" XSD, exactly as served by DOH (note: the served file is truncated — missing the final `</xs:schema>` tag) |
| `CostSubmission_repaired.xsd` | Same file with the missing closing tag appended, so validators can load it |
| `CommonTypes.xsd` | Official Common Types XSD (date format, DispositionFlag enum, ID types) |
| `Cost_Submission_SAMPLE_MF4159.xml` | Our template: one real MF4159 dental encounter, element order per the XSD, **validates PASS** against the repaired XSD |

## Key facts enforced by the schema / portal

- Dates `dd/mm/yyyy HH:MM`, GMT+4 (`DateTimeForm` pattern)
- `ReceiverID` must be `HAAD` (the only enumerated value)
- `DispositionFlag`: `PTE_SUBMIT` / `PTE_VALIDATE_ONLY` for the Public Test
  Environment, `PRODUCTION` for live
- Amounts in AED, max 3 decimals; file ≤ 5 MB (split by month/quarter if larger)
- One `<Claim>` per Claim.ID — never merge claims; multiple encounters in one
  claim roll up under one `<Encounter>`
- Zero encounter total is blocked unless a `<Resubmission>` element is present;
  negative totals are blocked; duplicate cost types/buckets per encounter blocked
- Resubmission only on DOH request, and only for Claim.IDs already submitted

## Two generations of the schema — IMPORTANT

**Generation 1** (used for the FY2024 submission, window Sept 2025). Matches the
Technical Document v1.12 sample AND the production guide's Appendix A (Aug 2025):
15 cost buckets **including `Supplies`**; no ActivityClinicianID, no ServiceCode,
15 cost types (no OHDepreciation). This is the layout of our Excel template.

**Generation 2** (the current "Cost Submission (NEW)" XSD on the dictionary —
presumably the format for the FY2025 window):
- Buckets: **`Supplies` is REMOVED**; `Dental` and `Home` are added (16 elements)
- `ActivityClinicianID` (required) and `ServiceCode` added to Encounter
- Cost types: `OHDepreciation` added (16 types)

Open question for DOH (dohclinicalcosting@doh.gov.ae): under Generation 2, where do
dental clinics report Supplies cost, and should dentistry use the new `Dental` bucket
instead of Physician/OP? Our sample parks Supplies in `Other`, clearly flagged.

## How the file is actually sent (production guide v1.0)

- SOAP web service, same channel as e-claims. Operation: **UploadTransaction**,
  file sent base64-encoded.
- Production: `https://shafafiya.doh.gov.ae/v3/webservices.asmx` (WSDL: `?wsdl`)
- PTE: `https://shafafiyapte.doh.gov.ae/v3/webservices.asmx`, username = Shafafiya
  username + `test` suffix, same password
- New user registration: `https://bpmweb.doh.gov.ae/usermanagement/MainPage.html`,
  then authorisation letter + email to shafafiyasupport@doh.gov.ae and IT@doh.gov.ae
- Known issues DOH warns about: report **Principal diagnosis only** (one code);
  report Procedure **only when TheatreAttendanceFlag = 1**
- `Connecting-to-Shafafiya-Webservices-Guide.pdf` in this folder is the official guide
  (includes Appendix A/B full XML samples)
