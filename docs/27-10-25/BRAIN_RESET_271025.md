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
**POST-PROCESSING WORKFLOW COMPLETE:** 8-step Python-based workflow finalized and validated (100% success rate, 76 slides, 0 failures). Workflow: unmerge-all → delete-carried-forward → merge-campaign → merge-monthly → merge-summary → fix-grand-total-wrap → remove-pound-totals → normalize-fonts. Fresh deck generated successfully (`run_20251024_200957`, 88 slides, 556KB).

**Session 27-10-25 Starting Point:**
- Fresh session with clean slate
- Production-ready pipeline available
- Ready to tackle next priorities

# Now / Next / Later
- **Now:**
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
- Session started fresh with context from 24-10-25
- No commits yet today
- Repository has untracked files from previous sessions (scripts/, tools/debug/, diagnostic utilities)

## Immediate TODOs
- [ ] Review Slide 1 geometry requirements from template
- [ ] Set up visual diff workflow with baseline comparison
- [ ] Identify which pytest suites need immediate attention
- [ ] Document test rehydration priorities

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
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251024_200957\AMP_Presentation_20251024_200957.pptx` (production deck - 88 slides, 556KB).
- Latest commit: `6c6f65e` - "docs: finalize 8-step post-processing workflow and end-of-session docs"
- Key documentation: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` (comprehensive ADR), `README.md` (COM prohibition warning), `docs/24-10-25.md` (previous session).
- Python module: `amp_automation/presentation/postprocess/` (cli.py, table_normalizer.py, cell_merges.py, unmerge_operations.py, span_operations.py).
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 27-10-25 (DD-MM-YY).
- Outstanding checklist (carry forward):
  - [ ] Set up visual diff baseline for Slide 1 geometry comparison
  - [ ] Rehydrate pytest test suites with current pipeline state
  - [ ] Design campaign pagination to prevent across-slide splits
  - [ ] Document Zen MCP evidence capture workflow

## How to validate this doc
- Confirm the production PPTX exists at `output\presentations\run_20251024_200957\AMP_Presentation_20251024_200957.pptx` and passes structural validation.
- Run the CLI deck generation with logging at INFO to ensure completion within ~5 minutes.
- Execute Python post-processing on a test deck and verify 100% success rate with 0 failures.
- Verify all fonts normalized to Verdana 6-7pt (except GRAND TOTAL at 6pt with special formatting).
