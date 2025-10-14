# 1) Session Overview
- Migrating AMP laydown automation to match `Template_V4_FINAL_071025.pptx` with accurate data transfer from Lumina Excel.
- Phase B table redesign underway; campaign/media block structure implemented in assembly pipeline.
- Validation tooling (reconciliation CLI/tests) already in place to guard accuracy.

# 2) Work Completed
- Updated `D:\OneDrive - Publicis Groupe\work\(.) AMP Laydowns Automation\amp_automation\presentation\assembly.py` with new grouped table builder `_prepare_main_table_data_detailed` and preserved legacy logic for reference.
- Confirmed regression coverage via `python -m pytest tests/test_template_validation.py` (requires `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1`).

# 3) Current State
- Table assembly outputs grouped campaign blocks but cell styling still follows legacy formatting.
- Repo has outstanding modifications/untracked files; no commits yet.
- Template geometry/config metadata already aligned with new PPT master.

# 4) Purpose
- Ensure generated PowerPoint decks visually and structurally match the new master template while preserving data fidelity for client reporting.

# 5) Next Steps
- [ ] Implement per-cell styling logic (fonts, alignments, fills, dual-line text) matching template spec.
- [ ] Handle continuation-slide splitting to retain headers, banding, and totals across pages.
- [ ] Re-run `python -m pytest tests/test_template_validation.py` after changes.

# 6) Important Notes
- Never alter template aesthetics (colours, fonts, layout); all adjustments must clone existing shapes/placeholders.
- Percentage tolerances ±0.5%; missing metrics display as `–`.
- Use absolute Windows paths in configs/logs and avoid introducing secrets.

# 7) Session Metadata
- Date: Unknown (set to current date when resuming).
- Tools: Python 3.13.8, pytest 8.4.2, PowerPoint template in `D:\OneDrive - Publicis Groupe\work\(.) AMP Laydowns Automation\template\Template_V4_FINAL_071025.pptx`.
- Repository: `D:\OneDrive - Publicis Groupe\work\(.) AMP Laydowns Automation` on branch `main`.
