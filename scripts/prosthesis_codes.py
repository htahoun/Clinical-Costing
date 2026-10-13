"""
Prosthesis procedure codes — the team's revised list (input/source/
LAB_Prosthesis_Implants_Codes.xlsx, sheet "Prosthesis codes", 48 codes),
accepted by the Finance Director on 13 Oct 2026 with the ruling:

    The clinic has no laboratory. "Prosthesis" in the clinic's language IS the
    dental lab. Every Laboratory cost is zero; every prosthesis code carries
    the Prostheses cost type / Prosthesis bucket.

So the month's prosthesis cost (in-house CAD/CAM materials, external
laboratory invoices and — as SWAllied routed to the Prosthesis bucket — the
lab technicians' pay) is placed on the claims that bill these codes, by billed
units, the same mechanism as the implant stock pools (scripts/implant_codes.py).
Implant-supported restorations appear in both lists: they consume an abutment
AND lab work.
"""

# code -> prosthesis subtype (the team's classification)
PROSTHESIS_CODES = {
    # indirect restorations
    "25521": "Indirect restoration", "25531": "Indirect restoration",
    # fixed dental prostheses (crowns, veneers, pontics, retainers)
    "27110": "Fixed dental prosthesis", "27111": "Fixed dental prosthesis", "27201": "Fixed dental prosthesis",
    "27206": "Fixed dental prosthesis", "27221": "Fixed dental prosthesis", "27602": "Fixed dental prosthesis",
    "62502": "Fixed dental prosthesis", "62701": "Fixed dental prosthesis", "62702": "Fixed dental prosthesis",
    "62703": "Fixed dental prosthesis", "62704": "Fixed dental prosthesis", "67111": "Fixed dental prosthesis",
    "67201": "Fixed dental prosthesis",
    # provisional crown / retainer (chairside)
    "27121": "Provisional crown", "67121": "Provisional retainer",
    # implant-supported restorations and components
    "26101": "Implant abutment", "27115": "Implant-supported crown", "27125": "Implant-supported crown",
    "27155": "Implant-supported crown", "27205": "Implant-supported crown", "67115": "Implant-supported retainer",
    "67145": "Implant-supported retainer", "67155": "Implant-supported retainer", "67205": "Implant-supported retainer",
    "69821": "Implant-supported framework", "69822": "Implant-supported framework", "29341": "Implant component",
    "79933": "Dental implant device",
    # removable dental prostheses
    "51301": "Removable dental prosthesis", "51302": "Removable dental prosthesis", "51601": "Removable dental prosthesis",
    "51602": "Removable dental prosthesis", "51611": "Removable dental prosthesis", "51612": "Removable dental prosthesis",
    "51711": "Removable dental prosthesis", "52101": "Removable dental prosthesis", "52102": "Removable dental prosthesis",
    "52111": "Removable dental prosthesis", "52301": "Removable dental prosthesis", "52511": "Removable dental prosthesis",
    "58301": "Removable dental prosthesis",
    # stents
    "57602": "Prosthetic stent",
    # implantable graft materials and membranes
    "42631": "Implantable graft material", "42702": "Implantable membrane",
    "79353": "Implantable graft material", "79356": "Implantable graft material",
}
assert len(PROSTHESIS_CODES) == 48

# Codes the team reviewed and kept OUT of Prosthesis (sheet "Implant-related review"): repairs, removals,
# recementation, relines, guides and templates. With no Laboratory cost type, the ones the team marked
# "Laboratory" (02951, 02952, 03001, 03002, 27722, 55402) stay in the Dental bucket (MS / dentists' time).
NOT_PROSTHESIS = {
    "02951", "02952", "03001", "03002", "42211", "27721", "27722", "27801", "29101", "29102", "29109", "29301",
    "29302", "29304", "29309", "66221", "66231", "66239", "66241", "73224", "29311", "29319", "29351", "55402",
    "55501", "55509", "56211", "56221", "66211", "66222", "66243", "66741", "79961", "79962", "25764", "23602",
}
assert not set(PROSTHESIS_CODES) & NOT_PROSTHESIS
