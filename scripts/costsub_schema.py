"""
Shafafiya Cost.Submission — Generation-2 schema map (single source of truth).

Every schema-dependent name (element names, element order, bucket and cost-type
lists, enumerations, formats) lives here so that aligning the generator to the
official XSD is a one-file change.

Confirmed 10 Oct 2026 against the twelve FY2025 files that Shafafiya accepted
and DOH evaluated (input/source/mysmyle_zip): nested money groups, Start/End,
PatientID inside Encounter, Procedure with Type, EDTriageLevel,
CriticalCareTime, VentilationTime, no TotalCost element.

Nothing in this file is costing logic. See docs/CLINICAL_COSTING.md §5.
"""

FACILITY_ID = "MF4159"
RECEIVER_ID = "HAAD"            # the only enumerated value
DATE_FMT = "%d/%m/%Y %H:%M"     # dd/mm/yyyy HH:MM, Gulf Standard Time (GMT+4)
TZ_OFFSET_HOURS = 4
MAX_FILE_BYTES = 5 * 1024 * 1024
DISPOSITIONS = ("PTE_SUBMIT", "PTE_VALIDATE_ONLY", "PRODUCTION")

ROOT = "Cost.Submission"
HEADER_ORDER = ["SenderID", "ReceiverID", "TransactionDate", "RecordCount", "DispositionFlag"]

# Generation 2 (current XSD): 16 buckets, no Supplies.
BUCKETS = ["Allied", "Anaesthesia", "Dental", "ED", "Home", "ICU", "Imaging",
           "Laboratory", "Physician", "OP", "OR", "Other", "Pharmacy",
           "Prosthesis", "SPS", "Ward"]

# Generation 2 cost types, in schema order.
COST_TYPES = ["SWNurs", "SWDoc", "SWAllied", "SWNonClin", "Laboratory", "Imaging",
              "Pharmacy", "Prostheses", "MS", "Hotel", "GS", "Depreciation",
              "OHF", "OHDepreciation", "OHC", "NABMs"]
OVERHEAD_COST_TYPES = ["OHF", "OHDepreciation", "OHC"]

# Money groups as written in the XML: <CostBucketDirect><Allied>…</Allied>…
MONEY_GROUPS = {
    "CostBucketDirect": BUCKETS,
    "CostBucketOverheads": BUCKETS,
    "CostTypeDetails": COST_TYPES,
}
# Internal flat keys = group name + child name (what costsub_reader produces).
DIRECT_ELEMENTS = [f"CostBucketDirect{b}" for b in BUCKETS]
OVERHEAD_ELEMENTS = [f"CostBucketOverheads{b}" for b in BUCKETS]
COST_TYPE_ELEMENTS = [f"CostTypeDetails{t}" for t in COST_TYPES]
MONEY_ELEMENTS = DIRECT_ELEMENTS + OVERHEAD_ELEMENTS + COST_TYPE_ELEMENTS

# Claim and Encounter element order (as accepted by Shafafiya).
CLAIM_ORDER = ["ID", "ProviderID", "Encounter"]
ENCOUNTER_ORDER = [
    "FacilityID", "ID", "ActivityClinicianID", "PatientID", "Type", "Start", "End",
    "StartType", "Specialty", "SubSpecialty", "DRGCode", "ServiceCode", "EndType",
    "Diagnosis", "Procedure", "EDTriageLevel", "OPAttendanceType",
    "TheatreAttendanceFlag", "CriticalCareTime", "VentilationTime",
    "CostBucketDirect", "CostBucketOverheads", "CostTypeDetails",
]
DIAGNOSIS_ORDER = ["Type", "Code"]
PROCEDURE_ORDER = ["Type", "Code", "Duration"]
PRINCIPAL = "Principal"

# Allocation-table column -> internal flat element (TMAP). Supplies is reported
# in the Dental bucket (what DOH accepted in the SY02 files); switch with
# --supplies-bucket if DOH ever rules otherwise.
TMAP_DIRECT = {
    "Anaesthesia": "CostBucketDirectAnaesthesia",
    "Imaging": "CostBucketDirectImaging",
    "Physician": "CostBucketDirectPhysician",
    "OP": "CostBucketDirectOP",
    "OtherLab": "CostBucketDirectLaboratory",
    "Prosthesis": "CostBucketDirectProsthesis",
    "Supplies": "CostBucketDirectSupplies",   # virtual Gen1 name, remapped by bucket_map()
}
TMAP_TYPES = {
    "SWNurs": "CostTypeDetailsSWNurs",
    "SWDoc": "CostTypeDetailsSWDoc",
    "SWAllied": "CostTypeDetailsSWAllied",
    "SWNonClin": "CostTypeDetailsSWNonClin",
    "MS": "CostTypeDetailsMS",
    "Hotel": "CostTypeDetailsHotel",
    "GS": "CostTypeDetailsGS",
    "Depreciation": "CostTypeDetailsDepreciation",
    "OHF": "CostTypeDetailsOHF",
    "OHC": "CostTypeDetailsOHC",
}
TMAP = {**TMAP_DIRECT, **TMAP_TYPES}


def bucket_map(supplies_bucket: str = "Dental") -> dict:
    """Gen1 virtual element -> Gen2 element. One-line DOH ruling switch."""
    if supplies_bucket not in BUCKETS:
        raise ValueError(f"supplies bucket must be one of {BUCKETS}")
    return {"CostBucketDirectSupplies": f"CostBucketDirect{supplies_bucket}"}


def split_key(flat: str) -> tuple[str, str]:
    """'CostBucketDirectDental' -> ('CostBucketDirect', 'Dental')."""
    for g in MONEY_GROUPS:
        if flat.startswith(g):
            return g, flat[len(g):]
    raise KeyError(flat)


def fmt_money_fils(fils: int) -> str:
    """Integer fils (AED/1000) -> '123.456' with exactly 3 dp."""
    sign = "-" if fils < 0 else ""
    fils = abs(fils)
    return f"{sign}{fils // 1000}.{fils % 1000:03d}"
