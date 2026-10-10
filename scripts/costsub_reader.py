"""
Tolerant reader for any Shafafiya Cost.Submission XML (ours or a platform's).

Walks the tree by local element names, so it copes with flat money elements
(<CostBucketDirectAnaesthesia>) and with nested groups
(<CostBucketDirect><Anaesthesia>), which it keys the same way.
"""
from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from lxml import etree

GROUP_PREFIXES = ("CostBucketDirect", "CostBucketOverheads", "CostTypeDetails")


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def fils(text) -> int:
    s = (text or "").strip().replace(",", "")
    if not s:
        return 0
    return int((Decimal(s) * 1000).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def read_submission(path: str | Path) -> dict:
    """
    Returns {"path", "size", "header": {...}, "claims": {claim_id: claim}, "errors": [...]}
    claim = {"ID","ProviderID","PatientID","encounters":[enc,...]}
    enc   = {"fields": {name: text}, "diagnoses": [(type, code)], "procedures": [(code, duration)],
             "money": {element: fils}, "money_text": {element: raw text}}
    """
    path = Path(path)
    out = {"path": str(path), "size": path.stat().st_size, "header": {}, "claims": {}, "errors": [],
           "claim_order": []}
    tree = etree.parse(str(path))
    root = tree.getroot()
    out["root"] = _local(root.tag)
    for el in root:
        name = _local(el.tag)
        if name == "Header":
            out["header"] = {_local(c.tag): (c.text or "").strip() for c in el}
        elif name == "Claim":
            claim = {"ID": "", "ProviderID": "", "PatientID": "", "encounters": []}
            for c in el:
                cn = _local(c.tag)
                if cn == "Encounter":
                    claim["encounters"].append(_read_encounter(c))
                else:
                    claim[cn] = (c.text or "").strip()
            cid = claim["ID"]
            if cid in out["claims"]:
                out["errors"].append(f"duplicate <Claim> ID {cid}")
                out["claims"][cid]["encounters"].extend(claim["encounters"])
            else:
                out["claims"][cid] = claim
                out["claim_order"].append(cid)
    return out


def _read_encounter(el) -> dict:
    enc = {"fields": {}, "diagnoses": [], "procedures": [], "money": {}, "money_text": {}}
    for c in el:
        n = _local(c.tag)
        if n == "Diagnosis":
            d = {_local(x.tag): (x.text or "").strip() for x in c}
            enc["diagnoses"].append((d.get("Type", ""), d.get("Code", "")))
        elif n == "Procedure":
            p = {_local(x.tag): (x.text or "").strip() for x in c}
            enc["procedures"].append((p.get("Code", ""), p.get("Duration", "")))
        elif n in GROUP_PREFIXES:                       # nested style
            for x in c:
                key = n + _local(x.tag)
                enc["money"][key] = fils(x.text)
                enc["money_text"][key] = (x.text or "").strip()
        elif n.startswith(GROUP_PREFIXES) or n == "TotalCost":
            enc["money"][n] = fils(c.text)
            enc["money_text"][n] = (c.text or "").strip()
        else:
            enc["fields"][n] = (c.text or "").strip()
    return enc


def encounter_total(enc: dict) -> int:
    """Sum of all money cells (buckets + cost types)."""
    return sum(enc["money"].values())


def claim_total(claim: dict) -> int:
    return sum(encounter_total(e) for e in claim["encounters"])


def file_total(sub: dict) -> int:
    return sum(claim_total(c) for c in sub["claims"].values())


def xsd_validate(xml_path: str | Path, xsd_path: str | Path | None) -> tuple[str, list[str]]:
    """Return ("PASS"|"FAIL"|"NO XSD", messages)."""
    if not xsd_path or not Path(xsd_path).exists():
        return "NO XSD", ["official XSD not present — drop it in docs/shafafiya/ (see README there)"]
    schema = etree.XMLSchema(etree.parse(str(xsd_path)))
    doc = etree.parse(str(xml_path))
    if schema.validate(doc):
        return "PASS", []
    return "FAIL", [f"line {e.line}: {e.message}" for e in schema.error_log][:50]


DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}$")
MONEY_RE = re.compile(r"^-?\d+\.\d{3}$")
