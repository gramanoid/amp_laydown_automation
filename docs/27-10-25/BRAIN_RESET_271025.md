Last prepared on 2025-10-27
Last verified on 27-10-25 (session start)

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
- **Post-processing uses Python (python-pptx)** for bulk operations - see `amp_automation/presentation/postprocess/`.
- AutoPPTX stays disabled except for negative tests; structural scripts, visual diff, and PowerPoint COM probes provide validation coverage.

# Current Position
**POST-PROCESSING WORKFLOW COMPLETE:** 8-step Python-based workflow finalized and validated (100% success rate, 76 slides, 0 failures). Workflow: unmerge-all → delete-carried-forward → merge-campaign → merge-media → merge-monthly → merge-summary → fix-grand-total-wrap → remove-pound-totals → normalize-fonts. Fresh deck generated successfully with all formatting improvements applied.

**Session 27-10-25 Work Completed:**
- ✅ Fixed timestamp generation to use local system time (Arabian Standard Time UTC+4) across all modules
- ✅ Implemented smart line breaking for campaign names (prevents mid-word breaks via _smart_line_break function)
- ✅ Added media channel vertical merging (TELEVISION, DIGITAL, OOH, OTHER, etc.)
- ✅ Corrected font sizes: 6pt body/bottom rows, 7pt BRAND TOTAL, 6pt campaign column
- ✅ Removed debug output from production code
- ⚠️ Campaign text wrapping issue identified: PowerPoint overriding explicit line breaks (documented in NOW_TASKS.md)

# Now / Next / Later
- **Now:**
  - [ ] **Fix campaign cell text wrapping** - Investigate column width increase or word-wrap disable (see NOW_TASKS.md)
  - [ ] **Slide 1 EMU/legend parity work** - Visual diff to compare generated vs template, fix geometry/legend discrepancies
  - [ ] **Test suite rehydration** - Fix/update `tests/test_tables.py`, `tests/test_structural_validator.py`
  - [ ] **Add regression tests** - Test merge correctness, font normalization, row formatting
- **Next:**
  - [ ] **Campaign pagination design** - Strategy to prevent campaign splits across slides
  - [ ] **Create OpenSpec proposal** - Document campaign pagination approach once design complete
  - [ ] **Python normalization expansion** - Consider row height normalization, cell margin/padding if needed
  - [ ] **Visual diff workflow** - Establish repeatable process with Zen MCP evidence capture
- **Later:**
  - [ ] **Automated regression scripts** - Catch rogue merges or row-height drift before decks ship
  - [ ] **Smoke test additional markets** - Validate pipeline with different data sets via `scripts/run_pipeline_local.py`
  - [ ] **Performance profiling** - Identify any bottlenecks in generation or post-processing pipeline

# 2025-10-27 Session Notes
- Completed formatting improvements: timestamp fix, smart line breaking, media merging, font corrections
- Latest deck: `run_20251027_193259` (generated with correct local timestamp 19:32:59 AST)
- Commits: d6f044a (timestamp fix), ace42e4 (5pt font attempt), 54df939 (media merging)
- Untracked: docs/NOW_TASKS.md (campaign wrapping issue documentation)

## Immediate TODOs
- [x] Fixed timestamp to use local system time across all modules
- [x] Implemented smart line breaking (_smart_line_break function)
- [x] Added media channel vertical merging
- [x] Corrected font sizes (6pt body, 7pt BRAND TOTAL)
- [ ] Fix campaign cell width to prevent PowerPoint word-wrap override (tomorrow's priority)

## Longer-Term Follow-Ups
- Complete campaign pagination design with Q&A-led discovery
- Expand Zen MCP/Compare workflows for evidence capture
- Add automated regression coverage for font/merge correctness
- Establish baseline metrics for pipeline performance monitoring

# Risks
- Without Slide 1 geometry verification, EMU/legend discrepancies may persist unnoticed
- Absent automated tests mean regressions rely on manual inspection
- Campaign pagination strategy delay may impact continuation slide quality
- Untracked diagnostic files may clutter repository if not organized

# Environment / Runbook
1. Generate deck:
   `python -m amp_automation.cli.main --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx --template D:\Drive\projects\work\AMP Laydowns Automation\template\Template_V4_FINAL_071025.pptx --output D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx`
2. Validate structure:
   `python D:\Drive\projects\work\AMP Laydowns Automation\tools\validate_structure.py D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx`
3. Post-process (Python normalization):
   `python -m amp_automation.presentation.postprocess.cli D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx postprocess-all`
4. Post-process (PowerShell wrapper):
   `& D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessNormalize.ps1 -PresentationPath D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx`
5. Tests (once restored):
   `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\test_autopptx_fallback.py tests\test_tables.py tests\test_assembly_split.py tests\test_structural_validator.py`

## Important Notes
- Always close existing PowerPoint sessions (`Stop-Process -Name POWERPNT -Force`) before running COM automation; scripts currently do not auto-close windows when failures occur.
- Maintain absolute Windows paths in documentation; avoid introducing secrets into logs or artifacts.
- Horizontal merge allowlist remains limited to MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD; treat any other merged labels as regressions.
- Template EMUs, centered alignment, and font sizes must remain faithful to `Template_V4_FINAL_071025.pptx`; adjust scripts cautiously to preserve pixel parity.

## Session Metadata
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251027_193259\AMP_Presentation_20251027_193259.pptx` (88 slides, with media merging and formatting improvements).
- Latest commit: `d6f044a` - "fix: use local system time for all timestamps instead of UTC"
- Key documentation: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` (comprehensive ADR), `README.md` (COM prohibition warning), `docs/NOW_TASKS.md` (campaign wrapping issue).
- Python module: `amp_automation/presentation/postprocess/` (cli.py, table_normalizer.py, cell_merges.py with _smart_line_break, unmerge_operations.py, span_operations.py).
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 27-10-25 (DD-MM-YY).
- Outstanding checklist (carry forward):
  - [ ] Fix campaign cell text wrapping (PowerPoint override issue)
  - [ ] Set up visual diff baseline for Slide 1 geometry comparison
  - [ ] Rehydrate pytest test suites with current pipeline state
  - [ ] Design campaign pagination to prevent across-slide splits
  - [ ] Document Zen MCP evidence capture workflow

## How to validate this doc
- Confirm the latest PPTX exists at `output\presentations\run_20251027_193259\AMP_Presentation_20251027_193259.pptx`
- Run the CLI deck generation with logging at INFO to ensure completion within ~5 minutes
- Execute Python post-processing on a test deck and verify 100% success rate with 0 failures
- Verify fonts: 6pt body/campaign/bottom rows, 7pt header/BRAND TOTAL
- Check media channel vertical merging (TELEVISION, DIGITAL, OOH, OTHER)
- Verify timestamps use local system time (Arabian Standard Time UTC+4)

Last verified on 27-10-25 (end of session)
