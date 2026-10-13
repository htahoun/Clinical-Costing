"""
Ledger-driven bucket rules (Finance Director's rulings, 10 Oct 2026).

Cost buckets (WHERE) are derived from the cost types (WHAT) by fixed shares,
instead of from the allocation table's bucket columns:

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

# cost type -> [(bucket, share, component)]; shares sum to 1. "LAB_SHARE" is replaced per month.
RULES = {
    "SWDoc": [("Anaesthesia", Decimal("0.02"), DIRECT), ("Dental", Decimal("0.98"), DIRECT)],
    "SWNurs": [("Anaesthesia", Decimal("0.01"), DIRECT), ("Imaging", Decimal("0.04"), DIRECT), ("Dental", Decimal("0.95"), DIRECT)],
    "SWAllied": [("Dental", Decimal(1), DIRECT)],   # lab technicians' share pre-assigned to Prosthesis by the engine
    "SWNonClin": [("Dental", Decimal(1), DIRECT)],
    "Laboratory": [("Laboratory", Decimal(1), DIRECT)],
    "Imaging": [("Imaging", Decimal(1), DIRECT)],
    "Pharmacy": [("Anaesthesia", Decimal(1), DIRECT)],
    "Prostheses": [("Prosthesis", Decimal(1), DIRECT)],
    "MS": [("Dental", Decimal(1), DIRECT)],
    "Hotel": [("Dental", Decimal(1), DIRECT)],
    "GS": [("Dental", Decimal(1), DIRECT)],
    "Depreciation": [("Imaging", Decimal("0.04"), DIRECT), ("Dental", Decimal("0.96"), DIRECT)],
    "OHF": [("Dental", Decimal(1), OVERHEAD)],
    "OHDepreciation": [("Dental", Decimal(1), OVERHEAD)],
    "OHC": [("Dental", Decimal(1), OVERHEAD)],
    "NABMs": [("Dental", Decimal(1), DIRECT)],
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
