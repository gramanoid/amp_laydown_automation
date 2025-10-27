Last verified on 2025-10-21

# Project Snapshot
- AMP deck generator clones `Template_V4_FINAL_071025.pptx`, binds Lumina Excel data, and now relies on a COM post-processor (`tools/PostProcessCampaignMerges.ps1`) to merge campaign blocks and monthly totals while enforcing 8.4 pt body row heights.
- Latest artefact: `output/presentations/run_20251021_185319/GeneratedDeck_20251021_MergedCells.pptx`, produced via CLI run plus COM post-pass.
- Structural validation still passes, and percentage formatting/font-sizing fixes are live; residual row-height spikes (14.4–21.6 pt) remain on some slides after repeated COM passes.

# Current Position
The COM script merges campaign-name cells vertically and monthly totals across the first three columns, but the monthly-total merge is non-idempotent—rerunning the script on the same deck leaves pre-existing spans that bubble into the next campaign block (e.g., FACES-CONDITION on Slides 2, 3, 5, 16). Row-height probes continue to report >8.4 pt for affected rows even though the enforcement loop runs.

# Now
- Implement a pre-pass in `tools/PostProcessCampaignMerges.ps1` that unmerges columns 1–3 for every data row before new campaign/monthly-total merges to guarantee idempotency.
- Re-run the post-processor on `output/presentations/run_20251021_185319/GeneratedDeck_20251021_MergedCells.pptx`, confirm campaign metric rows remain individual cells, and verify row heights revert to 8.4 pt.
- Execute the COM row-height probe end-to-end and capture logs showing 8.4 pt (±0.1 pt) across all data rows.

# Next
- Add automated validation (Python or COM) that scans all slides for unexpected column merges or row heights above tolerance.
- Regenerate Slide 1 visual diff artefacts once the table geometry stabilises and archive comparison evidence for review.

# Later
- Broaden regression coverage (pytest suite currently removed) and reintroduce smoke tests once table behaviour is stable.
- Enhance visual diff masking to reduce noise from legitimate data deltas and document the COM enforcement workflow in the runbook.

# Risks
- COM automation remains a single point of failure; missing Office runtimes block both post-processing and verification.
- Reprocessing decks without idempotent safeguards can reintroduce merged-cell artefacts and invalidate row-height enforcement.
- Visual diff variances persist until table geometry and legend parity stay locked for all slides.

# Environment/Runbook
1. Generate deck:  
   `python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx --template template/Template_V4_FINAL_071025.pptx --output "GeneratedDeck_{timestamp}.pptx"`
2. Validate structure:  
   `python tools/validate_structure.py output/presentations/<run>/GeneratedDeck_<ts>.pptx --excel template/BulkPlanData_2025_10_14.xlsx`
3. Run focused tests (plugins disabled):  
   `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests/test_tables.py tests/test_structural_validator.py` *(suite currently deleted; restore before running)*
4. Post-process table merges:  
   `& tools/PostProcessCampaignMerges.ps1 -PresentationPath <deck>`
5. Probe PowerPoint row heights (PowerPoint must be installed):  
   Execute COM snippet in `docs/21_10_25/` to log `table.Rows(idx).Height`; flag any value above 8.4 pt.

# How to validate this doc
- [ ] Unmerge-prepass implemented and committed.
- [ ] COM post-processor rerun shows correct campaign row structure on Slides 2/3/5/16.
- [ ] Row-height probe logs confirm 8.4 pt data rows across the latest deck.
- [ ] Automated merge/height validation script added to tooling.
- [ ] Visual diff artefacts refreshed and archived once metrics align.
