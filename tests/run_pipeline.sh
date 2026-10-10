#!/usr/bin/env bash
# End-to-end proof on synthetic data: build -> XML -> validate -> audit (self) -> triage sample.
set -euo pipefail
cd "$(dirname "$0")/.."
W=${1:-/tmp/cc_synth}
rm -rf "$W"; mkdir -p "$W"
python3 tests/make_synthetic.py "$W/in"
python3 scripts/build4_final.py --activity "$W/in/activity.xlsx" --alloc "$W/in/allocation.xlsx" --out "$W/costed.xlsx"
python3 scripts/make_xml_final.py --workbook "$W/costed.xlsx" --outdir "$W/xml" --disposition PTE_SUBMIT --zip
python3 scripts/validate_xml.py "$W"/xml/*.xml
python3 scripts/audit_mysmyle_xml.py "$W/xml/CostSubmission_MF4159_202501.xml" "$W/xml/CostSubmission_MF4159_202501.xml" | tail -4
printf 'Claim 12MF4159-%s rejected: schema validation failed, element CostBucketDirectOther is not expected\nError CC-104: Total cost does not match sum of cost types for claim 15MF4159-%s\n' \
  "$(grep -o '12MF4159-[0-9a-f]*' "$W/xml/CostSubmission_MF4159_202501.xml" | head -1 | cut -d- -f2)" \
  "$(grep -o '15MF4159-[0-9a-f]*' "$W/xml/CostSubmission_MF4159_202501.xml" | head -1 | cut -d- -f2)" > "$W/rejection.txt"
python3 scripts/triage_rejection.py "$W/rejection.txt" --xml "$W/xml" --out "$W/triage.xlsx"
python3 scripts/submit_shafafiya.py upload "$W/xml/CostSubmission_MF4159_202501.xml" --env pte --outdir "$W/soap"
echo "PIPELINE OK"
