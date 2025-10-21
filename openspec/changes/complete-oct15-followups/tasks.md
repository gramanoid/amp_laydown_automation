## 1. Template Geometry Alignment
- [ ] 1.1 Capture Template V4 column widths and table bounds as constants shared across assembly/tables modules.
- [ ] 1.2 Update continuation slide layout logic to honor exact Template V4 geometry (position, width, row heights).
- [ ] 1.3 Regenerate a presentation via `python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx --template template/Template_V4_FINAL_071025.pptx`.
- [ ] 1.4 Run `tools/visual_diff.py` and confirm metrics trend toward zero; archive comparison artifacts.
- [ ] 1.5 Perform manual PowerPoint Review→Compare against the master template and capture sign-off.

## 2. Pipeline Hierarchy Validation
- [ ] 2.1 Execute `python scripts/run_pipeline_local.py` with representative workbooks to validate the 00-08 artifact hierarchy.
- [ ] 2.2 Replace `tests/test_placeholder.py` with regression coverage exercising the new hierarchy and continuation logic.
- [ ] 2.3 Populate `input/` with curated production samples or document ingestion workflow for operators.
- [ ] 2.4 Audit repository docs/configs for legacy path references and update as needed.
