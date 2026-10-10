# Clinical-Costing

Clinical costing management for **Vision Dental Clinic L.L.C (DOH licence MF4159)**:
the annual Abu Dhabi DOH Clinical Cost Data Collection (ADCCDC) submission through
Shafafiya, and everything needed to build, validate, audit, submit and — when DOH
rejects a file — triage and resubmit it.

| Where | What |
|---|---|
| `docs/CLINICAL_COSTING.md` | Master reference: locked method, conventions, fixes, audit standard, history, open items |
| `docs/REJECTION_PLAYBOOK.md` | How a DOH/Shafafiya rejection is handled, start to finish |
| `docs/activity-template-columns.md` | Columns the activity master and allocation table must carry |
| `docs/shafafiya/` | Official XSD / technical documents go here (not yet in this repository) |
| `scripts/` | The engines (costing, XML, validation, audit, triage, transport) |
| `tests/run_pipeline.sh` | End-to-end proof on synthetic data |
| `.claude/skills/clinical-costing/`, `.claude/agents/` | The Claude skill and agent for this work |

## Quick start

```bash
pip install -r requirements.txt
bash tests/run_pipeline.sh            # synthetic end-to-end proof
# real run (inputs in input/, outputs in output/):
python3 scripts/build4_final.py --activity input/activity.xlsx --alloc input/allocation.xlsx \
    --out output/Clinical_Costing_2025_MF4159_FINAL.xlsx
python3 scripts/make_xml_final.py --workbook output/Clinical_Costing_2025_MF4159_FINAL.xlsx \
    --outdir output/xml_final --disposition PRODUCTION --supplies-bucket Other --zip
python3 scripts/validate_xml.py output/xml_final/*.xml --expect-disposition PRODUCTION
python3 scripts/triage_rejection.py input/rejection/* --xml output/xml_final
```

The repository is private: inputs, outputs, submitted XMLs and rejection files are committed so every comparison is reproducible from GitHub.
