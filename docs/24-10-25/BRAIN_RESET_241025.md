Last prepared on 2025-10-24
Last verified on 24-10-25 (updated with post-processing workflow completion)

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
**POST-PROCESSING WORKFLOW COMPLETE:** 8-step Python-based workflow finalized and validated (100% success rate, 76 slides, 0 failures). Workflow: unmerge-all → delete-carried-forward → merge-campaign → merge-monthly → merge-summary → fix-grand-total-wrap → remove-pound-totals → normalize-fonts. Column width adjustment attempted but reverted per requirements. Fresh deck generated successfully (`run_20251024_200957`, 88 slides, 556KB).

**Key Commits:**
- 4cd74e2: Updated openspec/project.md with completed work
- b4353ca: PowerShell integration checkpoint
- 1376406: Clarified Python CLI usage and merge scope

# Now / Next / Later
- **Now:**
  - ✅ **Cell merge implementation committed** (d3e2b98).
  - ✅ **Tested Python post-processing** - 88 slides in ~30 seconds (normalization works, merges redundant).
  - ✅ **Documented architecture** - merges belong in generation, not post-processing (8320c3f).
  - **Update PowerShell to use Python CLI** - For normalization operations only (not merges).
  - **End-to-end pipeline test** - generation (with merges) → Python normalization → validation.
- **Next:**
  - **Create OpenSpec proposal** documenting post-processing architecture (normalization focus).
  - **Simplify Python post-processing CLI** - Remove redundant merge operations OR repurpose for edge case repairs.
  - **Update ADR** - Clarify that COM prohibition applies to post-processing, generation merges are fine.
  - Resume Slide 1 EMU/legend parity work with visual diff + Zen MCP/Compare evidence capture.
- **Later:**
  - Rehydrate pytest suites (`tests\test_tables.py`, `tests\test_structural_validator.py`, etc.) and smoke additional markets via `scripts\run_pipeline_local.py`.
  - Design a campaign pagination approach that prevents across-slide splits, then raise an OpenSpec proposal once post-processing is stable.
  - Consider implementing span reset logic IF we decide to move merges to post-processing (unlikely).

# 2025-10-24 Session Notes
- CLI regeneration succeeded (`run_20251024_115954`, `run_20251024_121350`), producing INFO-level logs without the earlier DEBUG slowdowns.
- Post-process instrumentation now writes stopwatch, watchdog, and Trace-Command output to `docs\24-10-25\logs\05-postprocess_merges.log` and `05-postprocess_trace_*.log`.
- Slide 2 campaign rows remain at ~14–16 pt; the watchdog skips the slide after ~6 minutes, leaving merges incomplete until the height guard is relaxed.

## Immediate TODOs
- [x] Regenerated decks and logged CLI details.
- [x] Instrumented PowerShell scripts with logging and tracing.
- [x] Implemented plateau detection to prevent stalls.
- [x] Executed COM row-height probe.
- [x] **Documented COM prohibition across all key files**.
- [x] **Created comprehensive ADR** (`docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`).
- [x] **Implemented Python cell merge logic**.
- [x] **Committed documentation and foundation**.
- [x] **Generated fresh test decks**.
- [x] **Tested Python post-processing**.
- [x] **Documented merge architecture**.
- [x] **Updated PowerShell to call Python CLI** (`PostProcessNormalize.ps1`).
- [x] **End-to-end test** - generation → Python normalization → validation (PASSED).
- [x] **Created OpenSpec proposal** for post-processing architecture.
- [x] **Finalized 8-step post-processing workflow** with 100% success rate.
- [x] **Generated production-ready deck** (`run_20251024_200957` - 88 slides).

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
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251024_163905\presentations.pptx` (fresh deck for Python testing - 88 slides, 568KB).
- Latest commit: `3af7582` - "docs: prohibit COM bulk operations and add Python migration foundation" (32 files, 3,353 insertions, 1,127 deletions).
- Key documentation: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` (comprehensive ADR), `README.md` (COM prohibition warning), `docs/24-10-25.md` (daily changelog).
- Python module: `amp_automation/presentation/postprocess/` (cli.py, table_normalizer.py, cell_merges.py, span_operations.py).
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



