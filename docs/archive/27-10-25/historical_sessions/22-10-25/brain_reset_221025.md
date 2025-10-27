Last verified on 2025-10-22

# Project Snapshot
- Clone pipeline converts Lumina Excel exports into decks that mirror `Template_V4_FINAL_071025.pptx`, preserving template geometry, fonts, and layout.
- Table styling, percentage formatting, and legend suppression continue to be governed by Python clone logic; COM post-processing is still required for campaign/monthly merges and row-height enforcement.
- AutoPPTX remains disabled except for negative tests; validation relies on structural scripts, visual diff, and PowerPoint COM probes.

# Current Position
Horizontal three-column merges are controlled by `tools/FixHorizontalMerges.ps1` (JSON-guided), which now leaves only MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD spans. However, `tools/PostProcessCampaignMerges.ps1` repeatedly fails with `Unexpected HRESULT` during `Presentations.Open`, leaving PowerPoint sessions running and preventing campaign-name merges and row-height locks. The latest regenerated deck (`D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx`) has not been through a successful post-process run, so campaign columns remain unmerged and the row-height probe still reports historical offenders.

# Now / Next / Later
- **Now:** Stabilise `tools/PostProcessCampaignMerges.ps1` so it opens, processes, and closes PowerPoint without throwing HRESULT errors; rerun the script plus the row-height probe on the 2025-10-22 deck and capture artefacts under `docs\22-10-25\artifacts\`.
- **Next:** Add an automated regression check (python-pptx or COM) verifying that only the three allowed summary labels remain merged across columns 1-3, and ensure the pipeline fails fast if rogue merges return.
- **Later:** Resume Slide 1 EMU/legend parity work, refresh visual diff + Zen MCP evidence once geometry stabilises, and restore pytest coverage for summary tiles, legend behaviour, and clone-toggle-off scenarios.

# 2025-10-22 Session Notes
- Cleared `output\presentations\` and regenerated a fresh deck via `python -m amp_automation.cli.main ... --output output/presentations/run_20251022_115646/GeneratedDeck_20251022_132000.pptx` (561,713 bytes), preserving verbose logs for troubleshooting.
- Extended `tools/PostProcessCampaignMerges.ps1` with `Reset-HorizontalSpan`, refined vertical split logic, and path normalisation; observed that script currently opens PowerPoint but throws `Unexpected HRESULT` and leaves duplicate windows running.
- Verified `tools/FixHorizontalMerges.ps1` removes 53 rogue spans per `docs/22-10-25/merged_cells_analysis/merged_cells_fix_instructions.json`, retaining the 27 expected summary merges (MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD).
- Documented COM automation behaviour and mitigation attempts in `docs/22-10-25/logs/` (post-process attempts, regeneration log, merged-cells cleanup summary).

## Immediate TODOs
- [ ] Make `tools/PostProcessCampaignMerges.ps1` idempotent and resilient to COM Protected View / retry scenarios (consider explicit `Quit()` guards, back-off retries, and guaranteed `Stop-Process POWERPNT`).
- [ ] Rerun the campaign/monthly merge script on `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx` and confirm the deck retains merged campaign spans without splitting columns when errors occur.
- [ ] Execute the COM row-height probe against the updated deck and ensure all data rows land at 8.4 +/- 0.1 pt (capture output in `docs\22-10-25\artifacts\`).
- [ ] Archive troubleshooting outcomes plus COM logs in `docs\22-10-25\logs\` and update daily artefacts accordingly.

## Longer-Term Follow-Ups
- Add automated regression scripts catching rogue merges or row-height drift before decks ship.
- Complete Slide 1 geometry/legend parity fix and re-run visual diff plus Zen MCP/Compare workflows for evidence capture.
- Rehydrate pytest suites and smoke tests (e.g., `tests/test_tables.py`, `tests/test_structural_validator.py`) so pipeline regressions surface automatically.
- Facilitate a Q&A-led discovery to design campaign pagination that prevents campaigns from splitting across slides, then spin up a dedicated OpenSpec change once prerequisites are cleared.

# Risks
- PowerPoint COM instability (Unexpected HRESULT, lingering processes) blocks post-processing; failure to close sessions can corrupt decks or lock files.
- Without successful campaign merges, continuation slides present split Campaign columns, impacting visual parity.
- Absent automated tests mean regressions rely on manual inspection; future changes could reintroduce rogue merges or row-height drift unnoticed.

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
   Use the PowerPoint COM snippet in `docs\21_10_25\` to iterate `table.Rows(idx).Height`, log results for the latest deck, and store the output in `docs\22-10-25\artifacts\`.
6. Tests (once restored):
   `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\test_autopptx_fallback.py tests\test_tables.py tests\test_assembly_split.py tests\test_structural_validator.py`

## Important Notes
- Always close existing PowerPoint sessions (`Stop-Process -Name POWERPNT -Force`) before running COM automation; scripts currently do not auto-close windows when failures occur.
- Maintain absolute Windows paths in documentation; avoid introducing secrets into logs or artefacts.
- Horizontal merge allowlist is limited to MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD; any other merged labels should be treated as regressions.
- Template EMUs, centered alignment, and font sizes must remain faithful to `Template_V4_FINAL_071025.pptx`; adjust scripts cautiously to preserve pixel parity.

## Session Metadata
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx`
- Key logs/artefacts: `docs\22-10-25\logs\03-deck_regeneration.md`, `docs\22-10-25\logs\02-postprocess_attempt.md`, `docs\22-10-25\logs\04-merged_cells_cleanup.md`, `docs\22-10-25\artifacts\merged_cells_cleanup_20251022.txt`
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 22-10-25 (DD-MM-YY).
- Outstanding checklist (carry forward):
  - [ ] Run the generation, validation, and test commands successfully.
  - [ ] Confirm COM probe reports 8.4 +/- 0.1 pt across all Slide 1 body rows.
  - [ ] Document root cause for COM-driven row expansion and mitigation.
  - [ ] Verify latest deck lives under `output/presentations/run_20251021_185319/` (historical compliance check).
