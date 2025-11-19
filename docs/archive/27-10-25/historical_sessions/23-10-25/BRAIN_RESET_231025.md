Last prepared on 2025-10-23
Last verified on 2025-10-23

# Project Snapshot
- Clone pipeline converts Lumina Excel exports into decks that mirror `Template_V4_FINAL_071025.pptx`, preserving template geometry, fonts, and layout.
- Table styling, percentage formatting, and legend suppression continue to be governed by Python clone logic; COM post-processing remains essential for campaign/monthly merges and row-height enforcement.
- AutoPPTX stays disabled except for negative tests; structural scripts, visual diff, and PowerPoint COM probes provide validation coverage.

# Current Position
All prior runs were archived and the only trusted baseline is `output\presentations\baseline_20251022\GeneratedDeck_20251022_160155_baseline.pptx`. New sanitizer prototypes (`tools\SanitizePrimaryColumns.ps1`, `tools\RebuildCampaignMerges.ps1`) were created but have not produced a validated deck; attempts on 23 Oct corrupted campaign columns. Deck regeneration via CLI now stalls because of extremely verbose debug logging, so a fresh clean deck has not been produced and no row-height probe was captured today.

# Now / Next / Later
- **Now:**
  - Suppress or lower presentation assembly debug logging so CLI deck generation finishes within minutes and produces a verifiable run folder.
  - Regenerate a clean deck from the CLI, starting from the 20251022 baseline, and diff against the baseline to confirm geometry retention before any sanitisation.
  - Re-test sanitizer/merge workflow on a disposable copy and abort on first regression; capture results in `docs\23-10-25\logs\`.
- **Next:**
  - Implement a safe, deterministic merge rebuild (python-pptx or minimal COM) that never splits cells and documents merge rules.
  - Once the sanitizer+merge pipeline succeeds, rerun post-process + row-height probe, archiving CSV evidence under `docs\23-10-25\artifacts\`.
  - Produce visual diff / Zen MCP artefacts for the recovered deck and update the runbook with mitigation steps.
- **Later:**
  - Resume Slide 1 EMU/legend parity workstreams, including multi-slide PNG baselines.
  - Rehydrate pytest suites (`tests/test_tables.py`, `tests/test_structural_validator.py`, etc.) and execute smoke tests via `scripts/run_pipeline_local.py` across additional markets.
  - Design the no-campaign-splitting pagination strategy and prepare the corresponding OpenSpec change proposal.

# 2025-10-23 Session Notes
- Updated `tools\PostProcessCampaignMerges.ps1` with layout locks, span normalisation, and safeguarded retries; logs `02-04-postprocess_run_*.log` capture the resulting failures.
- Discovered the existing deck is irrecoverably merged (hundreds of-row spans); created `docs\23-10-25\logs\05-postprocess_hand_off.md` to guide the next agent.
- Manually terminated repeated hangs (`Stop-Process` on POWERPNT and worker `pwsh` instances) to leave a clean slate for follow-up work.

## Immediate TODOs
- [ ] Regenerate a fresh deck and log the new run under `docs\23-10-25\`.
- [ ] Sanitize columns 1-3 (unmerge + reset layout) across all slides before running COM merges.
- [ ] Run `tools\PostProcessCampaignMerges.ps1 -Verbose` on the sanitised deck and confirm no residual span warnings remain.
- [ ] Execute the COM row-height probe and store results in `docs\23-10-25\artifacts\`.
- [ ] Update `docs\23-10-25\logs\` with outcomes (success/failure) for each pass.

## Longer-Term Follow-Ups
- Add automated regression scripts catching rogue merges or row-height drift before decks ship.
- Complete Slide 1 geometry/legend parity fixes and refresh visual diff plus Zen MCP/Compare workflows for evidence capture.
- Rehydrate pytest suites and smoke tests so pipeline regressions surface automatically.
- Facilitate a Q&A-led discovery to design campaign pagination that prevents splits across slides, then raise an OpenSpec change once prerequisites clear.

# Risks
- PowerPoint COM instability (Unexpected HRESULT, lingering processes) blocks post-processing; failure to close sessions can corrupt decks or lock files.
- Without successful campaign merges, continuation slides present split Campaign columns, impacting visual parity.
- Absent automated tests mean regressions rely on manual inspection; rogue merges or row-height drift could resurface unnoticed.

# Environment / Runbook
1. Generate deck:
   `python -m amp_automation.cli.main --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx --template D:\Drive\projects\work\AMP Laydowns Automation\template\Template_V4_FINAL_071025.pptx --output D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx`
2. Validate structure:
   `python D:\Drive\projects\work\AMP Laydowns Automation\tools\validate_structure.py D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx`
3. Post-process merges (after COM fix):
   `& D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessCampaignMerges.ps1 -PresentationPath D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx`
4. Fix horizontal merges when required:
   `& D:\Drive\projects\work\AMP Laydowns Automation\tools\FixHorizontalMerges.ps1 -PresentationPath D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx`
5. Row-height probe (PowerPoint must be closed beforehand):
   Use the PowerPoint COM snippet in `docs\21_10_25\` to iterate `table.Rows(idx).Height`, log results for the latest deck, and store the output in `docs\23-10-25\artifacts\`.
6. Tests (once restored):
   `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\test_autopptx_fallback.py tests\test_tables.py tests\test_assembly_split.py tests\test_structural_validator.py`

## Important Notes
- Always close existing PowerPoint sessions (`Stop-Process -Name POWERPNT -Force`) before running COM automation; scripts currently do not auto-close windows when failures occur.
- Maintain absolute Windows paths in documentation; avoid introducing secrets into logs or artefacts.
- Horizontal merge allowlist remains limited to MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD; treat any other merged labels as regressions.
- Template EMUs, centered alignment, and font sizes must remain faithful to `Template_V4_FINAL_071025.pptx`; adjust scripts cautiously to preserve pixel parity.

## Session Metadata
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx`
- Key logs/artefacts: `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\03-deck_regeneration.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\02-postprocess_attempt.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\04-merged_cells_cleanup.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\merged_cells_cleanup_20251022.txt`
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 23-10-25 (DD-MM-YY).
- Outstanding checklist (carry forward):
  - [ ] Regenerate deck + validation logs for 23 Oct 2025 and capture run metadata.
  - [ ] Confirm COM probe reports 8.4 +/- 0.1 pt across all Slide 1 body rows.
  - [ ] Document root cause for legacy deck span explosion and mitigation strategy.
  - [ ] Reproduce visual diff / Zen MCP evidence once geometry stabilises.

## How to validate this doc
- Confirm the baseline PPTX exists at `output\presentations\baseline_20251022\GeneratedDeck_20251022_160155_baseline.pptx` and matches the template geometry.
- Run the CLI deck generation with logging at INFO to ensure completion within ~5 minutes and a new run directory containing a PPTX and log artifacts.
- Execute the sanitizer/merge workflow on a disposable copy and verify columns 1-3 remain unmerged before post-processing; inspect logs in `docs\23-10-25\logs\` for recorded outcomes.
