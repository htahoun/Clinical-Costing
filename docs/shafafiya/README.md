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

## Fields in the XSD that the Technical Document sample does NOT show

- `ActivityClinicianID` — required, after Encounter ID (our dentist licence, e.g. GD22503)
- `ServiceCode` — after DRGCode (may be empty)
- Cost buckets `Dental` and `Home` — the XSD has **17** buckets per block, not 15
- Cost type `OHDepreciation` — the XSD has **16** cost types, not 15

All four must be present in the XML for it to validate, even when zero/empty.
