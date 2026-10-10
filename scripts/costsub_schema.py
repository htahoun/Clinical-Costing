"""
Shafafiya Cost.Submission — Generation-2 schema map (single source of truth).

Every schema-dependent name (element names, element order, bucket and cost-type
lists, enumerations, formats) lives here so that aligning the generator to the
official XSD (docs/shafafiya/CostSubmission_repaired.xsd) is a one-file change.

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

DIRECT_ELEMENTS = [f"CostBucketDirect{b}" for b in BUCKETS]
OVERHEAD_ELEMENTS = [f"CostBucketOverheads{b}" for b in BUCKETS]
COST_TYPE_ELEMENTS = [f"CostTypeDetails{t}" for t in COST_TYPES]
MONEY_ELEMENTS = DIRECT_ELEMENTS + OVERHEAD_ELEMENTS + COST_TYPE_ELEMENTS + ["TotalCost"]

# Claim and Encounter element order. Children of Diagnosis / Procedure follow.
CLAIM_ORDER = ["ID", "ProviderID", "PatientID", "Encounter"]
ENCOUNTER_ORDER = [
    "FacilityID", "ID", "Type", "StartDateTime", "EndDateTime", "StartType", "EndType",
    "ActivityClinicianID", "ServiceCode", "Specialty", "SubSpecialty", "DRGCode",
    "OPAttendanceType", "TheatreAttendanceFlag", "TriageCategory",
    "CriticalCareHours", "VentilationHours",
    "Diagnosis", "Procedure",
    *DIRECT_ELEMENTS, *OVERHEAD_ELEMENTS, *COST_TYPE_ELEMENTS, "TotalCost",
]
DIAGNOSIS_ORDER = ["Type", "Code"]
PROCEDURE_ORDER = ["Code", "Duration"]
PRINCIPAL = "Principal"

# Allocation-table column -> XML element (TMAP). Supplies is parked in `Other`
# pending the DOH ruling (Dental vs Other); switch with --supplies-bucket.
TMAP_DIRECT = {
    "Anaesthesia": "CostBucketDirectAnaesthesia",
    "Imaging": "CostBucketDirectImaging",
    "Physician": "CostBucketDirectPhysician",
    "OP": "CostBucketDirectOP",
    "OtherLab": "CostBucketDirectLaboratory",
    "Prosthesis": "CostBucketDirectProsthesis",
    "Supplies": "CostBucketDirectSupplies",   # virtual Gen1 name, remapped by BUCKET_MAP
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


def bucket_map(supplies_bucket: str = "Other") -> dict:
    """Gen1 virtual element -> Gen2 element. One-line DOH ruling switch."""
    if supplies_bucket not in BUCKETS:
        raise ValueError(f"supplies bucket must be one of {BUCKETS}")
    return {"CostBucketDirectSupplies": f"CostBucketDirect{supplies_bucket}"}


def fmt_money_fils(fils: int) -> str:
    """Integer fils (AED/1000) -> '123.456' with exactly 3 dp."""
    sign = "-" if fils < 0 else ""
    fils = abs(fils)
    return f"{sign}{fils // 1000}.{fils % 1000:03d}"
