"""
Ledger-driven bucket rules: the Finance Director's tuning grid of 13 Oct 2026
(replaces the 10 Oct 2026 shares; see GRID below).

Cost buckets (WHERE) are derived from the cost types (WHAT) by fixed shares,
instead of from the allocation table's bucket columns. The 10 Oct 2026 shares were:

  Pharmacy                -> Anaesthesia
  SWDoc                   -> 2 % Anaesthesia, 98 % Dental
  SWNurs                  -> 1 % Anaesthesia, 4 % Imaging, 95 % Dental
  Depreciation (direct)   -> 4 % Imaging (OPG / X-ray units), 96 % Dental
  Imaging supplies        -> Imaging
  Laboratory              -> Laboratory (zero: ruling 13 Oct 2026, the clinic has no laboratory)
  SWAllied                -> Dental (CSSD); the lab technicians' share is pre-assigned by the engine
                             to the Prosthesis bucket on the prosthesis-coded claims
  Prostheses              -> Prosthesis (implants AND all lab-fabricated prostheses)
  MS, GS, Hotel, SWNonClin, NABMs -> Dental
  OHF, OHDepreciation, OHC -> Dental, overhead component (CostBucketOverheads)

Because every dirham takes its bucket from its cost type, buckets equal cost
types on every claim by construction. The Physician bucket stays 0 (Appendix 2:
medical physicians; dentists belong to Dental Services).
"""
from __future__ import annotations

from decimal import Decimal

DIRECT, OVERHEAD = "direct", "overhead"

# Finance Director's tuning grid, 13 Oct 2026 (medical, finance, purchasing and accounts teams):
# cost type -> % to (Anaesthesia, Dental, Imaging, Prosthesis). Rows absent or all-zero in the grid
# (Laboratory, Hotel, NABMs) default to Dental. Overhead cost types keep the OVERHEAD component in
# whichever bucket they land (CostBucketOverheads/<bucket>).
GRID = {
    "SWNurs":         (5, 80, 5, 10),
    "SWDoc":          (3, 77, 5, 15),     # 13 Oct 2026 tuning: Anaesthesia 3 %
    "SWAllied":       (0, 80, 10, 10),    # applied to the CSSD remainder; the dental technicians' share is
                                          # pre-assigned to Prosthesis by the engine (ruling 13 Oct 2026)
    "SWNonClin":      (0, 100, 0, 0),
    "Laboratory":     (0, 100, 0, 0),     # zero cost type (no laboratory)
    "Imaging":        (0, 10, 90, 0),     # dentist's own time reading / taking X-rays counted in Dental
    "Pharmacy":       (100, 0, 0, 0),
    "Prostheses":     (0, 0, 0, 100),     # 13 Oct 2026 tuning: implants and lab work stay in Prosthesis (App.1)
    "MS":             (3, 87, 5, 5),      # 13 Oct 2026 tuning
    "Hotel":          (0, 100, 0, 0),
    "GS":             (0, 100, 0, 0),
    "Depreciation":   (5, 90, 5, 0),      # equipment; lab machines are not "prosthesis" (ruling)
    "OHF":            (0, 98, 2, 0),      # 13 Oct 2026 tuning: OPG room share of overheads by floor area
    "OHDepreciation": (0, 100, 0, 0),     # back office
    "OHC":            (0, 100, 0, 0),
    "NABMs":          (0, 100, 0, 0),
}
GRID_BUCKETS = ("Anaesthesia", "Dental", "Imaging", "Prosthesis")
OVERHEAD_TYPES = {"OHF", "OHDepreciation", "OHC"}
for _t, _row in GRID.items():
    assert sum(_row) == 100, (_t, _row)

RULES = {
    t: [(b, Decimal(p) / Decimal(100), OVERHEAD if t in OVERHEAD_TYPES else DIRECT) for b, p in zip(GRID_BUCKETS, row) if p]
    for t, row in GRID.items()
}
DEFAULT_LAB_SHARE = Decimal("0.45")   # lab technicians / (lab technicians + CSSD) when the month is unknown (engine pool)


def rules_for_month(lab_share: Decimal | None) -> dict[str, list[tuple[str, Decimal, str]]]:
    ls = lab_share if lab_share is not None else DEFAULT_LAB_SHARE
    out = {}
    for t, lst in RULES.items():
        row = []
        for b, sh, comp in lst:
            if sh == "LAB_SHARE":
                sh = ls
            elif sh == "1-LAB_SHARE":
                sh = Decimal(1) - ls
            row.append((b, sh, comp))
        out[t] = row
    return out
