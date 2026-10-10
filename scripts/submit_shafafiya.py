#!/usr/bin/env python3
"""
Shafafiya transport: SOAP UploadTransaction (base64 file) — dry-run by default.

  python3 scripts/submit_shafafiya.py upload output/xml_final/*.xml --env pte            # writes envelopes only
  python3 scripts/submit_shafafiya.py upload output/xml_final/*.xml --env prod --send    # really sends

Credentials come from the environment: SHAFAFIYA_USER, SHAFAFIYA_PASS (never put
them in the repo). PTE logins carry the `test` suffix on the username — pass the
exact username DOH issued for each environment. Resubmission only on DOH request.

Endpoints and the operation name are from docs/CLINICAL_COSTING.md §5; confirm
the SOAP action and response fields against the WSDL (docs/shafafiya/) before a
live send — the DOH hosts are not reachable from every network.
"""
from __future__ import annotations

import argparse
import base64
import os
import sys
import urllib.request
from pathlib import Path

ENDPOINTS = {
    "pte": "https://shafafiyapte.doh.gov.ae/v3/webservices.asmx",
    "prod": "https://shafafiya.doh.gov.ae/v3/webservices.asmx",
}
NS = "http://www.eClaimLink.ae/"   # Shafafiya v3 web-service namespace; confirm against the WSDL


def envelope(user: str, pwd: str, name: str, content_b64: str) -> str:
    return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <UploadTransaction xmlns="{NS}">
      <login>{user}</login>
      <pwd>{pwd}</pwd>
      <fileContent>{content_b64}</fileContent>
      <fileName>{name}</fileName>
    </UploadTransaction>
  </soap:Body>
</soap:Envelope>"""


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["upload"])
    ap.add_argument("files", nargs="+")
    ap.add_argument("--env", choices=ENDPOINTS, default="pte")
    ap.add_argument("--send", action="store_true", help="actually POST (default: dry run, write envelopes)")
    ap.add_argument("--outdir", default="output/soap")
    args = ap.parse_args(argv)

    user, pwd = os.environ.get("SHAFAFIYA_USER", ""), os.environ.get("SHAFAFIYA_PASS", "")
    if args.send and not (user and pwd):
        raise SystemExit("set SHAFAFIYA_USER and SHAFAFIYA_PASS in the environment")
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from costsub_reader import read_submission
    url = ENDPOINTS[args.env]
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    for f in args.files:
        p = Path(f)
        disp = read_submission(p)["header"].get("DispositionFlag")
        if args.env == "prod" and disp != "PRODUCTION":
            raise SystemExit(f"{p.name}: DispositionFlag {disp} must be PRODUCTION for the live endpoint")
        if args.env == "pte" and disp == "PRODUCTION":
            raise SystemExit(f"{p.name}: PRODUCTION file aimed at the test endpoint")
        body = envelope(user or "<SHAFAFIYA_USER>", "<hidden>" if not args.send else pwd, p.name,
                        base64.b64encode(p.read_bytes()).decode())
        if not args.send:
            (outdir / f"{p.stem}.soap.xml").write_text(body)
            print(f"dry run: {p.name} ({disp}) -> {url}; envelope written to {outdir / (p.stem + '.soap.xml')}")
            continue
        req = urllib.request.Request(url, data=body.encode(), headers={
            "Content-Type": "text/xml; charset=utf-8", "SOAPAction": f"{NS}UploadTransaction"})
        with urllib.request.urlopen(req, timeout=300) as resp:
            text = resp.read().decode(errors="replace")
        (outdir / f"{p.stem}.response.xml").write_text(text)
        print(f"sent {p.name} -> HTTP {resp.status}; response saved to {outdir / (p.stem + '.response.xml')}")
        print(text[:800])
    return 0


if __name__ == "__main__":
    sys.exit(main())
