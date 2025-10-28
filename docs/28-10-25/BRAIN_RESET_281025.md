Last prepared on 2025-10-28

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
**VALIDATION SUITE COMPLETE & RECONCILIATION INVESTIGATION READY:**
- 8-step Python post-processing workflow validated (100% success rate)
- Structural validator enhanced to handle last-slide-only shapes (BRAND TOTAL, indicator shapes only on final slides)
- Comprehensive data validation suite implemented: 4 new modules + unified report generator (1,200+ lines of validation code)
- All validators tested on 144-slide production deck - PASS status
- Repository cleanup complete: tools reorganized, archives documented, logs restructured

**Session 27-10-25 Work Completed:**
- ✅ Fixed timestamp generation to use local system time (Arabian Standard Time UTC+4) across all modules
- ✅ Implemented smart line breaking for campaign names (prevents mid-word breaks via _smart_line_break function)
- ✅ Added media channel vertical merging (TELEVISION, DIGITAL, OOH, OTHER, etc.)
- ✅ Corrected font sizes: 6pt body/bottom rows, 7pt BRAND TOTAL, 6pt campaign column
- ✅ Campaign text wrapping issue RESOLVED: Removed hyphens at source + widened column to 1,000,000 EMU
- ✅ **Updated structural validator:** Handle last-slide-only shapes (QuarterBudget*, MediaShare*, FunnelShare*, FooterNotes)
- ✅ **Expanded data validation suite:**
  - data_accuracy.py (160 lines) - numerical accuracy validation
  - data_format.py (280 lines) - format/style validation (1,575 checks per deck)
  - data_completeness.py (170 lines) - required data presence validation
  - validation/utils.py (190 lines) - shared utilities
  - tools/validate_all_data.py (250 lines) - unified report generator
- ✅ **Fixed validator bugs:** Table cell indexing in data_accuracy.py, metadata filtering in reconciliation.py
- ✅ **Repository cleanup (Tier 6):** Tools reorganized, archive documentation, logs restructured

# Now / Next / Later
- **Now (28-10-25) - UPDATED with Actual Status:**
  - [ ] **Test suite rehydration** - Restore `tests/test_tables.py`, `tests/test_structural_validator.py` with current pipeline assertions
  - [ ] **Add regression tests** - Campaign merging, font normalization, row height consistency test coverage
  - [ ] **Document campaign pagination approach** - Specify max_rows=40 strategy and continuation slide handling

- **Next (Deprioritized/Optional):**
  - [ ] **Slide 1 EMU/legend parity work** - Visual diff to compare template vs generated (archived, lower priority)
  - [ ] **Visual diff workflow** - Establish repeatable process with Zen MCP evidence capture (Phase 4+ work)
  - [ ] **Automated regression scripts** - Catch rogue merges or row-height drift before decks ship
  - [ ] **Python normalization expansion** - Consider row height normalization, cell margin/padding if needed
  - [ ] **Smoke test additional markets** - Validate pipeline with different data sets

- **✅ Completed (27 Oct Evening - All Fixed Before Session End):**
  - [x] **Reconciliation data source investigation** - Fixed in commit e27af1e; 100% pass rate (630/630 records). Root cause: case-sensitivity + pagination format + market code mapping. Solution: Added normalization functions for market/brand matching.
  - [x] **Fix campaign cell text wrapping** - Removed hyphens + widened column to 1,000,000 EMU (commit 395025b)
  - [x] **Fix structural validator** - Updated to handle last-slide-only shapes (BRAND TOTAL on final slides) (commit 6e83fae)
  - [x] **Expand data validation suite** - 1,200+ lines across 5 modules, all tests PASS (commits 203a90e, 28a74f0)
  - [x] **Repository cleanup (Tier 6)** - Tools reorganized, archives documented, logs restructured (commits 2861d70, c6d42f4, c161e58)
  - [x] **Campaign pagination enhancement** - Max_rows=40 strategy verified on 144-slide deck (commit 951bb14)

# 2025-10-28 Session Notes

**Session Context (28-10-25):**
- **Status:** All critical issues from 27-10-25 have been resolved and committed
- Reconciliation validator now passes 100% (630/630 records) - fixed in commit e27af1e
- All formatting improvements verified: timestamp, smart line breaking, media merging, fonts
- Data validation suite operational with comprehensive coverage (1,200+ lines)
- Production deck ready: `run_20251027_215710` (144 slides, all improvements)
- Next priorities: Test suite rehydration + regression test coverage

# Immediate TODOs (28-10-25)
- [ ] Test suite rehydration - restore test_tables.py and test_structural_validator.py
- [ ] Add regression tests - campaign merging, font normalization, row height coverage
- [ ] Document campaign pagination approach - max_rows=40 strategy + continuation slides

# Risks
- Test suite absence means regressions detected late (manual inspection only)
- Slide 1 geometry parity unverified (EMU/legend discrepancies possible but lower priority)
- Python normalization expansion needs assessment (row height, cell padding not yet addressed)
- Additional market testing not yet performed (pipeline validated only on current data)

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
6. Validate all data:
   `python D:\Drive\projects\work\AMP Laydowns Automation\tools\validate_all_data.py D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx`

## Important Notes
- Always close existing PowerPoint sessions (`Stop-Process -Name POWERPNT -Force`) before running COM automation
- Maintain absolute Windows paths in documentation; avoid introducing secrets into logs or artifacts
- Horizontal merge allowlist: MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD (treat other merges as regressions)
- Template EMUs, centered alignment, and font sizes must remain faithful to `Template_V4_FINAL_071025.pptx`
- Reconciliation validator failing at scale - investigate Excel data source mapping

## Session Metadata
- Previous deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251027_215710\AMP_Presentation_20251027_215710.pptx` (144 slides)
- Latest commit: `d655002` - "docs: session end closure for 27-10-25 - reconciliation validation complete"
- Key documentation: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`, `docs/NOW_TASKS.md`, `openspec/project.md`
- Python post-processing module: `amp_automation/presentation/postprocess/`
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 28-10-25 (DD-MM-YY)
- Outstanding checklist (carry forward):
  - [ ] Reconciliation data source investigation (630/631 failing checks)
  - [ ] Slide 1 geometry parity setup with visual diff baseline
  - [ ] Test suite rehydration with current pipeline state
  - [ ] Campaign pagination strategy refinement
  - [ ] Zen MCP evidence capture workflow documentation

## How to validate this doc
- Confirm the latest PPTX exists at `output\presentations\run_20251027_215710\AMP_Presentation_20251027_215710.pptx`
- Run the CLI deck generation with logging at INFO to ensure completion within ~5 minutes
- Execute Python post-processing on a test deck and verify 100% success rate
- Verify fonts: 6pt body/campaign/bottom rows, 7pt header/BRAND TOTAL
- Check media channel vertical merging (TELEVISION, DIGITAL, OOH, OTHER)
- Verify timestamps use local system time (Arabian Standard Time UTC+4)
- Run reconciliation validation and document findings

Last verified on 28-10-25 (session start)
