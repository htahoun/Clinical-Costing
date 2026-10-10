# docs/shafafiya — official DOH Shafafiya Cost.Submission artefacts

Place here (they are referenced by the scripts):

| File | Used by |
|---|---|
| `CostSubmission.xsd` | the served XSD (known to be truncated) — keep for reference |
| `CostSubmission_repaired.xsd` | **the one the scripts validate against** (`make_xml_final.py`, `validate_xml.py`, `audit_mysmyle_xml.py` default to this path) |
| technical PDFs (Shafafiya web services guide, ADCCDC data dictionary / code lists) | specialty codes, clinician ID rules, SOAP operation names |
| a validated sample XML | regression reference for `audit_mysmyle_xml.py` |
| the WSDL of `v3/webservices.asmx` | confirm the SOAP namespace/action in `submit_shafafiya.py` |

Status (10 Oct 2026): these files were part of the previous project and are
not yet in this repository. The DOH hosts (`shafafiya.doh.gov.ae`,
`shafafiyapte.doh.gov.ae`, `www.shafafiya.org`) are not reachable from the
cloud sandbox, so they must be uploaded by the Finance Director. Until the XSD
is here every verification table shows `NO XSD` and the element order in
`scripts/costsub_schema.py` is unconfirmed.
