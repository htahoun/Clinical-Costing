"""
The governing monthly expenditure ledger (Arjun / accounting), FY2025.
Locked figures — see docs/CLINICAL_COSTING.md §6 item 6. Amounts in AED.

Every costed month must tie to these figures to the fils (3 dp) exactly.
"""
from decimal import Decimal

GL_2025 = {
    1: Decimal("511302.30"),
    2: Decimal("532341.65"),
    3: Decimal("606868.87"),
    4: Decimal("518123.67"),
    5: Decimal("488066.69"),
    6: Decimal("616494.09"),
    7: Decimal("480166.82"),
    8: Decimal("526990.10"),
    9: Decimal("509542.05"),
    10: Decimal("546725.98"),
    11: Decimal("493746.03"),
    12: Decimal("623033.73"),
}
GL_2025_TOTAL = Decimal("6453401.98")

GL = {2025: GL_2025}


def gl_for(year: int, month: int) -> Decimal:
    try:
        return GL[year][month]
    except KeyError:
        raise SystemExit(f"No ledger figure for {year}-{month:02d}; add it to scripts/gl_new.py")


def gl_fils(year: int, month: int) -> int:
    """Ledger figure as integer fils (AED x 1000)."""
    return int((gl_for(year, month) * 1000).to_integral_value())


if __name__ == "__main__":
    total = sum(GL_2025.values())
    print("FY2025 ledger by month (AED)")
    for m, v in GL_2025.items():
        print(f"  {m:02d}  {v:>14,.2f}")
    print(f"  total {total:>12,.2f}   expected {GL_2025_TOTAL:,.2f}   "
          f"{'OK' if total == GL_2025_TOTAL else 'MISMATCH'}")
