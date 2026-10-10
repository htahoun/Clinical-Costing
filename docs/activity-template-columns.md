# Input column specification

The readers in `scripts/build2_lib.py` match headings tolerantly (case,
spaces and punctuation ignored). Any one of the accepted spellings works.

## Activity master (one row = one unit of one CPT code inside one claim)

| Canonical | Required | Accepted headings | Used for |
|---|---|---|---|
| claim_id | yes | Claim ID, ClaimID, Claim.ID, Claim Ref, Claim Reference, Claim No | `Claim.ID` |
| serial | yes | Serial, Activity Serial, Serial No, Encounter Serial, Activity ID, Encounter ID | `Encounter.ID` = first line's serial (trailing `B` stripped) |
| patient_id | yes | FileNo., FileNo, File No, Patient ID, PatientID, MRN | `PatientID` |
| clinician | yes | Doctor No, DoctorNo, Clinician, Clinician ID, ActivityClinicianID | `ActivityClinicianID` |
| start | yes | StartDateTime, Start Date Time, Encounter Start, Start, Date | month allocation + `StartDateTime` |
| code | yes | Code, CPT, CPT Code, Activity Code, Service Code | allocation lookup |
| specialty | no | Speciality, Specialty | `Specialty` (blank → clinician's modal) |
| end | no | EndDateTime, End Date Time, Encounter End, End | `EndDateTime` (defaults to start) |
| units | no | Units, Quantity, Qty | multiplies Cost/Unit (default 1) |
| diagnosis | no | Diagnosis, Principal Diagnosis, Diagnosis Code, ICD, ICD Code | principal diagnosis |
| theatre | no | TheatreAttendanceFlag, Theatre | procedure gating |
| procedure / duration | no | Procedure, Procedure Code / Duration | only when theatre = 1 |
| enc_type | no | Type, Encounter Type | `Type` (default 1) |
| payer, net, source | no | Payer / Net, Amount / Source | provenance and audit only |

Dates: `dd/mm/yyyy HH:MM` preferred; Excel datetimes and ISO dates accepted.

## Allocation table (one row per CPT code)

`Code`, `Cost/Unit` (or Cost per Unit / Unit Cost), and the 17 percentage
columns in this order: Anaesthesia, Imaging, Physician, OP, OtherLab,
Prosthesis, Supplies, SWNurs, SWDoc, SWAllied, SWNonClin, MS, Hotel, GS,
Depreciation, OHF, OHC. Percentages sum to 100 (fractions summing to 1 are
also accepted). The locked fixes in `docs/CLINICAL_COSTING.md` §3 are applied
automatically and logged.

## Ledger

`scripts/gl_new.py`, one figure per month, AED 2 dp. Change it only on a new
ledger from accounting and record the change in the master reference.
