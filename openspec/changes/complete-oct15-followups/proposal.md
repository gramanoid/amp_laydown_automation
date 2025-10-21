## Why
Visual fidelity checks flagged MainDataTable geometry drift from Template V4, and the October 15 pipeline snapshot left several follow-up activities outstanding.

## What Changes
- Align table column widths and placement with Template V4.
- Rebuild and verify the presentation output against the master template.
- Validate the new artifact hierarchy through representative pipeline runs and upgraded regression coverage.

## Impact
- Affected specs: presentation rendering, pipeline orchestration.
- Affected code: `amp_automation/presentation/assembly.py`, `amp_automation/presentation/tables.py`, `config/master_config.json`, `tools/visual_diff.py`, pipeline execution scripts, regression tests.
