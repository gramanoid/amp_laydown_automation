Last prepared on 2025-10-28
Last verified on 28-10-25 (End of Session - All Tasks Complete)

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
**FORMATTING & OUTPUT IMPROVEMENTS COMPLETE:**
- **6 formatting enhancements delivered:** Bold TOTAL/GRPs columns, merged % cells with bold styling, smart quarterly budget formatting, evenly distributed quarterly boxes, AMP_Laydowns_ddmmyy output naming, DD-MM-YY footer dates
- **Output standardization:** All presentations now use consistent `AMP_Laydowns_ddmmyy.pptx` naming pattern (dynamically extracted from Excel filename)
- **Footer dates dynamic:** Source date in footer automatically updates based on Excel file date (YYYY_MM_DD pattern → DD-MM-YY display)
- **Percentage cells:** Column 17 cells now fully merged with bold styling applied
- **Production deck:** `run_20251028_163719/AMP_Laydowns_281025.pptx` (127 slides, all improvements validated)

**Session 28-10-25 Work Completed:**
- ✅ **Point 1:** Make TOTAL and GRPs columns (15, 16) bold in all cells (assembly.py:590, tables.py:808-809)
- ✅ **Point 2:** Merge percentage cells vertically (col 17) like campaign merging with boundary detection at gray MONTHLY TOTAL rows (cell_merges.py:448-525)
- ✅ **Point 3:** Merge until row before gray MONTHLY TOTAL (working via boundary detection logic)
- ✅ **Point 4:** Campaign percentage calculation already correct (assembly.py:1234-1239)
- ✅ **Point 5:** GRAND TOTAL shows 100% already implemented (assembly.py:1319)
- ✅ **Point 6:** Fixed MONTHLY TOTAL label pound symbol (£) - was being removed by postprocess normalization, fixed in table_normalizer.py:416 (now only removes from numeric cells, preserves in label)
- ✅ **Bold percentage cells:** Updated _apply_cell_styling() to apply bold to merged cells when no new text provided (cell_merges.py:744-756)
- ✅ **Smart quarterly formatting:** Values >= 1000K display as M (1211K → 1.2M), values < 1000K stay as K (300K → 300K) (assembly.py:558-586)
- ✅ **Quarterly box distribution:** Evenly spaced with equal gaps (1.289" between Q1-Q2, Q2-Q3, Q3-Q4)
- ✅ **Output filename standardization:** AMP_Laydowns_{timestamp}.pptx with %d%m%y format (master_config.json, cli/main.py:252-270)
- ✅ **Footer date format:** Changed from DDMMYY to DD-MM-YY with dynamic extraction from Excel filename YYYY_MM_DD pattern

# Now / Next / Later
- **Now (28-10-25) - COMPLETED:**
  - [x] **Test suite rehydration** - Restored `tests/test_tables.py`, `tests/test_structural_validator.py` with 16 regression tests (commit 9d4eca9)
  - [x] **Add regression tests** - 8 formatting tests, 3 structural validator tests, 5 footer date extraction tests (commit 9d4eca9)
  - [x] **Fix validate_structure.py** - Corrected PROJECT_ROOT calculation bug (commit e4dbbd5)

- **📦 Future Work (Archived - Deprioritized):**
  - 🗂️ **Slide 1 EMU/legend parity work** - Lower priority; visual diff workflow would be Phase 4+ effort
  - 🗂️ **Visual diff workflow** - Establish repeatable process with Zen MCP evidence capture (Phase 4+ work)
  - 🗂️ **Automated regression scripts** - Catch rogue merges or row-height drift before decks ship
  - 🗂️ **Python normalization expansion** - Consider row height normalization, cell margin/padding if needed
  - 🗂️ **Smoke test additional markets** - Validate pipeline with different data sets

  *Rationale: All 6 formatting improvements from 28-10-25 are complete and tested. These items are nice-to-have optimizations that can be addressed in future iterations without blocking current production use.*

- **✅ Completed (27 Oct Evening - All Fixed Before Session End):**
  - [x] **Reconciliation data source investigation** - Fixed in commit e27af1e; 100% pass rate (630/630 records). Root cause: case-sensitivity + pagination format + market code mapping. Solution: Added normalization functions for market/brand matching.
  - [x] **Fix campaign cell text wrapping** - Removed hyphens + widened column to 1,000,000 EMU (commit 395025b)
  - [x] **Fix structural validator** - Updated to handle last-slide-only shapes (BRAND TOTAL on final slides) (commit 6e83fae)
  - [x] **Expand data validation suite** - 1,200+ lines across 5 modules, all tests PASS (commits 203a90e, 28a74f0)
  - [x] **Repository cleanup (Tier 6)** - Tools reorganized, archives documented, logs restructured (commits 2861d70, c6d42f4, c161e58)
  - [x] **Campaign pagination enhancement** - Max_rows=40 strategy verified on 144-slide deck (commit 951bb14)

# 2025-10-28 Session Notes

**Session Context (28-10-25 - Evening):**
- **Status:** All 6 formatting/output improvement requests delivered and committed
- **User Requirements:** One-by-one implementation with approval gates at each step (critical methodology)
- All implementations verified in fresh decks; 6 commits pushed to main
- Production deck ready: `run_20251028_163719/AMP_Laydowns_281025.pptx` (127 slides, all improvements)
- Next priorities: Test suite rehydration + regression test coverage

**Commits Today (28-10-25):**
1. bcbb026 - Fix: apply bold styling to merged percentage cells
2. 0bdfb16 - Feat: smart quarterly budget formatting and dimension updates
3. 83516cb - Fix: evenly distribute quarterly budget boxes horizontally
4. 7b9754c - Feat: standardize output filename to AMP_Laydowns_ddmmyy format
5. 80c997b - Fix: change footer source date format to DD-MM-YY

# Session 28-10-25 Completion Status

✅ **ALL TASKS COMPLETED:**
- [x] 6 formatting improvements (bold columns, merged percentage cells, quarterly formatting, output naming, footer dates)
- [x] Test suite rehydration (16 regression tests across test_tables.py and test_structural_validator.py)
- [x] validate_structure.py PROJECT_ROOT bug fix
- [x] Comprehensive documentation updates

# Risks & Mitigations

**Mitigated:**
- ✅ Test suite regression coverage - 16 tests now in place (commit 9d4eca9)
- ✅ validate_structure.py bug - PROJECT_ROOT fixed (commit e4dbbd5)

**Archived (Future Work):**
- 🗂️ Slide 1 geometry parity - Lower priority; deferred to Phase 4+ work
- 🗂️ Python normalization expansion - Deferred; only required if new normalization issues surface
- 🗂️ Additional market testing - Deferred; pipeline validated on current production data

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
- Archived items:
  - 🗂️ Reconciliation data source investigation (COMPLETED 27-10-25: 630/630 passing)
  - 🗂️ Slide 1 geometry parity (Deferred to Phase 4+)
  - 🗂️ Test suite rehydration (COMPLETED 28-10-25: 16 regression tests)
  - 🗂️ Visual diff workflow (Deferred to Phase 4+)
  - 🗂️ Automated regression scripts (Deferred to Phase 4+)

## How to validate this doc
- Confirm the latest PPTX exists at `output\presentations\run_20251028_163719\AMP_Laydowns_281025.pptx`
- Run the CLI deck generation with logging at INFO to ensure completion within ~5 minutes
- Execute Python post-processing on a test deck and verify 100% success rate
- Verify fonts: 6pt body/campaign/bottom rows, 7pt header/BRAND TOTAL
- Check media channel vertical merging (TELEVISION, DIGITAL, OOH, OTHER)
- Verify timestamps use local system time (Arabian Standard Time UTC+4)
- Run test suite: `pytest tests/test_tables.py tests/test_structural_validator.py -v --tb=short`
- Verify 16 regression tests pass (8 formatting, 3 structural, 5 footer date extraction)

## Session Completion Summary

**Session 28-10-25 Evening:**
- ✅ 6 formatting improvements implemented and tested
- ✅ Test suite rehydrated with 16 comprehensive regression tests
- ✅ validate_structure.py PROJECT_ROOT bug fixed
- ✅ Documentation fully updated and archived
- ✅ All code committed to main branch
- ✅ Production deck ready: `run_20251028_163719/AMP_Laydowns_281025.pptx` (127 slides)

**Archived/Deferred (Phase 4+ work):**
- Slide 1 EMU/legend parity
- Visual diff workflow
- Automated regression scripts (shell/Python)
- Python normalization expansion
- Smoke tests with additional markets

Last verified on 28-10-25 (session completion - all tasks archived/completed)
