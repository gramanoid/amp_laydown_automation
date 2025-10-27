Last verified on 2025-10-21

# Project Snapshot
- Clone pipeline generates decks from `Template_V4_FINAL_071025.pptx` using Lumina Excel exports.
- Slide geometry, fonts, column widths, legend parity, and trailing blank rows now align with the template.
- Row heights are forced to 8.4 pt in generation; a COM post-pass now merges campaign blocks and monthly totals, yet PowerPoint still reports 14.4–21.6 pt on select rows in `output/presentations/run_20251021_185319/GeneratedDeck_20251021_MergedCells.pptx`.

# Current Position
Row-height mitigation adds an explicit 8.4 pt lock across body/subtotal rows (hRule = exact), normalises percentage labels plus font sizing, and introduces a COM post-processor to merge campaign name blocks and monthly totals, but the latest probe on `output/presentations/run_20251021_185319/GeneratedDeck_20251021_MergedCells.pptx` still returns 14.4–21.6 pt on rows 21–32. Structural checks continue to pass; variance is confined to runtime rendering.

# Now/Next/Later
- **Now:** Diagnose the residual 14.4–21.6 pt readings after the COM merge pass (validate `HeightRule` behaviour and inspect campaign blocks); rerun the probe once all data rows read 8.4 pt (±0.1 pt).
- **Next:** Capture updated Slide 1 visual diff artefacts and archive evidence once metrics hit template values.
- **Later:** Revisit data masking for visual diff, broaden regression coverage, and document the row-height enforcement approach.

# 2025-10-21 Session Notes
- Implemented `tools/PostProcessCampaignMerges.ps1` to merge campaign-name spans and monthly totals while enforcing fonts (header 7 pt, body 6 pt, totals 6.5 pt) and 8.4 pt row heights; generated `output/presentations/run_20251021_185319/GeneratedDeck_20251021_MergedCells.pptx` via CLI + post-pass.
- Confirmed decimal-suppression fix keeps body percentages integer-formatted; campaign name font sizes now consistent at 6 pt.
- Investigated persistent tall rows on Slide 2, tracing issue to non-idempotent monthly-total merge logic that carries pre-existing spans upward when the COM script reruns; reproduction observed on Slides 2, 3, 5, 16 in the latest deck.

## Immediate TODOs
- Add a pre-pass to `tools/PostProcessCampaignMerges.ps1` that unmerges columns 1–3 for every data row before applying campaign and monthly-total merges, ensuring the post-processor is idempotent.
- Re-run the COM post-processor on the latest deck and verify affected slides show separated metric rows beneath each campaign.
- Execute the COM row-height probe end-to-end and confirm all data rows report 8.4 pt ± 0.1.

## Longer-Term Follow-Ups
- Extend automated validation to flag unexpected column merges or row-height deviations across all slides.
- Regenerate Slide 1 visual diff artefacts once row heights stabilise and archive comparison evidence.
- Document the COM workflow, including idempotency safeguards, within the broader runbook and ensure future regressions trigger automated alerts.

# Risks
- Fallback shrink-to-fit auto-sizing is insufficient, so PowerPoint may keep expanding rows until we add post-processing.
- COM automation is required for verification; missing Office runtime blocks validation.
- Visual diff variances remain until we neutralise data deltas.

# Environment/Runbook
1. Generate a deck:  
   `python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx --template template/Template_V4_FINAL_071025.pptx --output "GeneratedDeck_{timestamp}.pptx"`
2. Validate structure:  
   `python tools/validate_structure.py output/presentations/<run>/GeneratedDeck_<ts>.pptx --excel template/BulkPlanData_2025_10_14.xlsx`
3. Run focused tests:  
   `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests/test_autopptx_fallback.py tests/test_tables.py tests/test_assembly_split.py tests/test_structural_validator.py`
4. Run COM post-processing merges:  
   `& tools/PostProcessCampaignMerges.ps1 -PresentationPath <deck>` merges campaign name blocks and monthly totals, reapplies fonts, and reasserts 8.4 pt targets.
5. Probe row heights (PowerPoint must be installed):  
   Use the COM snippet in `docs/21_10_25/` to log `table.Rows(idx).Height`; investigate any readings above 8.4 pt (latest run: `output/presentations/run_20251021_185319/GeneratedDeck_20251021_MergedCells.pptx` still shows 14.4–21.6 pt spikes).

### How to validate this doc
- [ ] Run the generation, validation, and test commands successfully.
- [ ] Confirm COM probe reports 8.4 pt across all Slide 1 body rows.
- [ ] Document root cause for COM-driven row expansion and applied mitigation.
- [ ] Verify current deck lives under `output/presentations/run_20251021_185319/`.
