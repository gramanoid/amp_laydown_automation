Last prepared on 2025-10-24
Last verified on 2025-10-24 (updated with COM prohibition)

---

# 🚨 CRITICAL ARCHITECTURE DECISION 🚨

**DO NOT use PowerPoint COM automation for bulk table operations.**

COM-based bulk operations are **PROHIBITED** due to catastrophic performance issues discovered on 24 Oct 2025:
- PowerShell COM: **10+ hours** (never completed)
- Python (python-pptx): **~10 minutes**
- Performance difference: **60x minimum**

**READ THIS:** `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`

**Approved Usage:**
- ✅ Python (python-pptx, aspose_slides) for ALL bulk operations
- ✅ COM only for file I/O, exports, features not in python-pptx

**Prohibited:**
- 🚫 COM loops over cells/rows/columns
- 🚫 COM bulk property changes
- 🚫 COM merge operations in loops

---

# Project Snapshot
- Clone pipeline converts Lumina Excel exports into decks that mirror `Template_V4_FINAL_071025.pptx`, preserving template geometry, fonts, and layout.
- Table styling, percentage formatting, and legend suppression are governed by Python clone logic.
- **Post-processing now uses Python (python-pptx)** for bulk operations - see `amp_automation/presentation/postprocess/`.
- AutoPPTX stays disabled except for negative tests; structural scripts, visual diff, and PowerPoint COM probes provide validation coverage.

# Current Position
Fresh decks exist in `output\presentations\run_20251024_121026` (debug-heavy) and `output\presentations\run_20251024_121350` (INFO baseline). Logging fixes are working, and the post-process script now emits stopwatch checkpoints, watchdog exits, and Trace-Command output through `docs\24-10-25\logs`. The sanitized deck continues to stall during COM row-height enforcement on slide 2 (rows plateau at ~14–16 pt), so campaign/monthly merges remain incomplete until we soften or bypass the height guard.

# Now / Next / Later
- **Now:**
  - Cap or bypass `Set-RowHeightExact` retries so stubborn rows (slide 2) no longer stall the merge loop; document the guard and rerun the repro harness.
  - Capture a fresh row-height probe for the sanitized deck and archive it under `docs\24-10-25\artifacts\`.
  - Decide whether to retain or compress the 65 MB debug log from `run_20251024_121026` now that INFO logs cover the latest run.
- **Next:**
  - Re-run `tools\PostProcessCampaignMerges.ps1 -Verbose` once the height guard is updated; verify watchdog logs stay empty and slide 2 merges succeed.
  - Update `docs\24-10-25\logs\` with merge/probe outcomes and note any residual anomalies.
  - Draft a deterministic merge rebuild concept (python-pptx or minimal COM) to reduce reliance on fragile COM loops.
- **Later:**
  - Resume Slide 1 EMU/legend parity work with visual diff + Zen MCP/Compare evidence capture.
  - Rehydrate pytest suites (`tests\test_tables.py`, `tests\test_structural_validator.py`, etc.) and smoke additional markets via `scripts\run_pipeline_local.py`.
  - Design a campaign pagination approach that prevents across-slide splits, then raise an OpenSpec proposal once sanitise/merge stability is proven.

# 2025-10-24 Session Notes
- CLI regeneration succeeded (`run_20251024_115954`, `run_20251024_121350`), producing INFO-level logs without the earlier DEBUG slowdowns.
- Post-process instrumentation now writes stopwatch, watchdog, and Trace-Command output to `docs\24-10-25\logs\05-postprocess_merges.log` and `05-postprocess_trace_*.log`.
- Slide 2 campaign rows remain at ~14–16 pt; the watchdog skips the slide after ~6 minutes, leaving merges incomplete until the height guard is relaxed.

## Immediate TODOs
- [x] Regenerated decks (`run_20251024_115954`, `run_20251024_121350`) and logged CLI details in `docs\24-10-25\logs\02-cli_regeneration.md`.
- [x] Instrumented `tools\PostProcessCampaignMerges.ps1` with stopwatch logging, watchdog exits, and COM tracing (see `docs\24-10-25\logs\05-postprocess_merges.log`).
- [x] Implement a height-guard cap in `Set-RowHeightExact` so slide 2 rows stop stalling the merge loop (see `docs\24-10-25\logs\06-plateau_fix_summary.md`).
- [x] Verified plateau detection working: logs show rows 12-24 plateaued at 14.4-15.6pt and early-exit engaged.
- [x] Generated fresh deck (`run_20251024_142119`) with proper timestamp naming and verified plateau detection in real full-deck scenario (see `docs\24-10-25\logs\11-plateau_verification_results.md`).
- [x] Executed COM row-height probe and stored results in `docs\24-10-25\artifacts\row_height_probe_20251024_142119.csv` (1372 rows probed, analysis in `row_height_probe_analysis.md`).
- [ ] Run full-deck post-processing without debug logging to measure end-to-end improvement.

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
   Use the PowerPoint COM snippet in `docs\21_10_25\` to iterate `table.Rows(idx).Height`, log results for the latest deck, and store the output in `docs\24-10-25\artifacts\`.
6. Tests (once restored):
   `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\test_autopptx_fallback.py tests\test_tables.py tests\test_assembly_split.py tests\test_structural_validator.py`

## Important Notes
- Always close existing PowerPoint sessions (`Stop-Process -Name POWERPNT -Force`) before running COM automation; scripts currently do not auto-close windows when failures occur.
- Maintain absolute Windows paths in documentation; avoid introducing secrets into logs or artefacts.
- Horizontal merge allowlist remains limited to MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD; treat any other merged labels as regressions.
- Template EMUs, centered alignment, and font sizes must remain faithful to `Template_V4_FINAL_071025.pptx`; adjust scripts cautiously to preserve pixel parity.

## Session Metadata
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251024_121350\AMP_Presentation_20251024_121350.pptx` (latest INFO-level run with logs).
- Key logs/artefacts: `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\03-deck_regeneration.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\02-postprocess_attempt.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\04-merged_cells_cleanup.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\merged_cells_cleanup_20251022.txt`
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 24-10-25 (DD-MM-YY).
- Outstanding checklist (carry forward):
  - [ ] Regenerate deck + validation logs for 24 Oct 2025 and capture run metadata.
  - [ ] Confirm COM probe reports 8.4 +/- 0.1 pt across all Slide 1 body rows.
  - [ ] Document root cause for legacy deck span explosion and mitigation strategy.
  - [ ] Reproduce visual diff / Zen MCP evidence once geometry stabilises.

## How to validate this doc
- Confirm the baseline PPTX exists at `output\presentations\baseline_20251022\GeneratedDeck_20251022_160155_baseline.pptx` and matches the template geometry.
- Run the CLI deck generation with logging at INFO to ensure completion within ~5 minutes and a new run directory containing a PPTX and log artifacts.
- Execute the sanitize/merge workflow on a disposable copy and verify columns 1-3 remain unmerged before post-processing; inspect logs in `docs\24-10-25\logs\` for recorded outcomes.



