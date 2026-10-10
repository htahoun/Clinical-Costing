"""
Implant procedure codes (Finance Director's list, input/source/Implant_Codes.xlsx,
10 Oct 2026) and the stock sub-categories whose cost they carry.

Used to place the Prostheses cost type on implant claims only (instead of
spreading it over every Dental-bucket claim). Each stock pool is distributed
over the claims that contain its codes, in proportion to units.
"""

# Stock sub-category group -> CPT codes that consume it.
IMPLANT_POOLS = {
    # Implant fixtures and healing / cover screws: consumed when the implant is placed or re-entered.
    "fixtures": {
        "79933",  # surgical implant installation, 1 tooth (fixture + transmucosal element)
        "79935",  # surgical re-entry, healing screw removal and placement of abutment
        "79936",  # surgical re-entry, healing screw removal and placement of abutment (additional)
    },
    # Abutments and implant impression tools: consumed at the abutment / crown / framework stage.
    "abutments": {
        "26101",  # mesostructure implant abutment
        "27155",  # implant crown impression (acrylic / composite / pre-fabricated metal base)
        "27205",  # implant-supported crown, 1 tooth
        "67115",  # provisional implant-supported bridge retainer
        "67155",  # provisional implant-supported bridge retainer with metal base
        "69821",  # upper fixed prosthodontic framework, osseo-integrated
        "69822",  # lower fixed prosthodontic framework, osseo-integrated
        "79935",  # re-entry with abutment placement also consumes an abutment
        "79936",
    },
    # Bone grafting materials (xenograft, synthetic, membranes)
    "bone_graft": {
        "42631",  # periodontal surgery, osseous graft per sextant
        "42702",  # guided tissue regeneration, resorbable membrane per site
        "72431",  # alveolar bone preservation, xenograft, first tooth
        "72439",  # alveolar bone preservation, xenograft, each additional tooth
        "79511",  # harvesting of intraoral tissue for grafting
    },
}
# Implant-related codes that consume no implant stock (surgery / removal only): identified as implant
# work for reporting, but carry no Prostheses cost type from the pools above.
IMPLANT_NO_STOCK = {"29311", "29319", "42211", "73111", "73121", "79961", "79962"}

IMPLANT_CODES = set().union(*IMPLANT_POOLS.values()) | IMPLANT_NO_STOCK

# Stock movement sub-categories (cogs_from_movement.py) -> pool
SUBCATEGORY_POOL = {
    "01- Implant Fixtures": "fixtures",
    "02- Healing & Cover Screw": "fixtures",
    "03- Abutments": "abutments",
    "04- Implant Impression Tools & Supplies": "abutments",
    "05- Bone Grafting Materials, Synthetic": "bone_graft",
}
