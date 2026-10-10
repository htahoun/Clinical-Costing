# Clinical-Costing — Vision Dental Clinic L.L.C (MF4159)

Abu Dhabi DOH Clinical Cost Data Collection (ADCCDC) workflow: scale-only costing
of claim activity to the monthly ledger, Generation-2 Shafafiya Cost.Submission
XML generation, validation, audit, rejection triage and transport.

- Master reference (locked decisions, conventions, history): `docs/CLINICAL_COSTING.md`. Read it first.
- Rejection procedure: `docs/REJECTION_PLAYBOOK.md`.
- Skill: `.claude/skills/clinical-costing/SKILL.md`. Agent: `.claude/agents/clinical-costing.md`.
- Engines: `scripts/` (see the repository map in the master reference). Proof run: `bash tests/run_pipeline.sh`.
- The repository is private: `input/` (uploads) and `output/` (builds, XMLs, rejection files) are committed so comparisons can be done in GitHub.
- Python 3.11+, `pip install -r requirements.txt`.
