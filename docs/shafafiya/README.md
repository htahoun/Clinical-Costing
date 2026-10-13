# docs/shafafiya — official DOH Shafafiya Cost.Submission artefacts

Place here (they are referenced by the scripts):

| File | Used by |
|---|---|
| `CostSubmission.xsd` | the served XSD (known to be truncated) — keep for reference |
| `CostSubmission_repaired.xsd` | **the one the scripts validate against** (`make_xml_final.py`, `validate_xml.py`, `audit_mysmyle_xml.py` default to this path) |
| technical PDFs (Shafafiya web services guide, ADCCDC data dictionary / code lists) | specialty codes, clinician ID rules, SOAP operation names |
| a validated sample XML | regression reference for `audit_mysmyle_xml.py` |
| the WSDL of `v3/webservices.asmx` | confirm the SOAP namespace/action in `submit_shafafiya.py` |

Present (13 Oct 2026), downloaded from doh.gov.ae (PDF + extracted text):

| File | What it is |
|---|---|
| `Abu_Dhabi_Clinical_Costing_Standard_v2_1_2026-01-27.*` | the Standard (cost types App.1, cost buckets App.2, allocation statistics App.3) |
| `Abu_Dhabi_Clinical_Costing_Guidelines_v2_2026-01-27.*` | the Guidelines v2 (cost ledger, drivers, prostheses at actual cost) |
| `Clinical_Costing_Cost_Data_Collection_Technical_Document_SY2_v1.00.*` | **SY2 technical document**: XML schema sample, field validations (pp. 11-19), portal validations (p. 20), DOH internal quality checks for ED / inpatient / outpatient encounters (pp. 22-25), Feb 2026 change log (Dental and Home buckets and OH Depreciation added, Supplies removed, ActivityClinicianID and ServiceCode added) |
| `Clinical_Costing_FAQ_v1.0.*` | the FAQ (Q4.3 small dental clinic may consolidate admin; Q4.4 ask DOH when a service does not map) |
| `Quick_Start_Guide_Clinical_Costing_v10_FINAL.*` | the Quick Start Guide |

Status (10 Oct 2026): the files below were part of the previous project and are
not yet in this repository. The DOH hosts (`shafafiya.doh.gov.ae`,
`shafafiyapte.doh.gov.ae`, `www.shafafiya.org`) are not reachable from the
cloud sandbox, so they must be uploaded by the Finance Director. Until the XSD
is here every verification table shows `NO XSD` and the element order in
`scripts/costsub_schema.py` is unconfirmed.
