# Master Consolidated TODO List - 27 Oct 2025

**Generated:** 2025-10-27 14:05
**Fresh Deck:** `output/presentations/run_20251027_135302/presentations.pptx` (88 slides, 565KB)
**Status:** 4 Active OpenSpec changes, 39 pending tasks total

---

## 📊 Quick Overview

| Priority | Count | Estimated Hours |
|----------|-------|-----------------|
| 🔴 CRITICAL | 8 | 10-12 hours |
| 🟠 HIGH | 12 | 15-18 hours |
| 🟡 MEDIUM | 12 | 12-15 hours |
| 🟢 LOW | 7 | 6-8 hours |
| **TOTAL** | **39** | **43-53 hours** |

### OpenSpec Changes
- ✅ `update-table-styling-continuations` - ARCHIVED (2025-10-21)
- ⚠️ `adopt-template-cloning-pipeline` - 1 task remaining (88% complete)
- ⚠️ `complete-oct15-followups` - 6 tasks remaining (14% complete)
- ⚠️ `clarify-postprocessing-architecture` - 8 tasks remaining (Phase 1 complete)
- 🆕 `implement-campaign-pagination` - 17 tasks (NEW - Option A selected)

---

## 🔴 CRITICAL PRIORITY (Must Do First - 10-12 hours)

### Visual Parity & Quality Assurance
**Why Critical:** Business requirement for client-facing decks, blocks archival of major OpenSpec work

1. **⏰ 2h** - Capture Template V4 geometry constants (`complete-oct15-followups` 1.1)
   - Extract EMU values: column widths, table bounds, row heights
   - Store in `assembly.py`/`tables.py` or config constants
   - **Output:** Constants module with template measurements
   - **Blocks:** Continuation slide alignment (1.2)

2. **⏰ 1.5h** - Update continuation slide layout logic (`complete-oct15-followups` 1.2)
   - Apply geometry constants from 1.1
   - Update `amp_automation/presentation/assembly.py`
   - **Dependencies:** Task 1 must complete first
   - **Output:** Pixel-perfect continuation slides

3. **⏰ 1h** - Run visual_diff.py validation (`complete-oct15-followups` 1.4)
   - Execute: `python tools/visual_diff.py` on `run_20251027_135302` vs template
   - Analyze pixel differences and geometry deviations
   - **Output:** `docs/27-10-25/artifacts/visual_diff_results.md`

4. **⏰ 0.5h** - Manual PowerPoint Review→Compare (`complete-oct15-followups` 1.5 + `adopt-template-cloning` 4.4)
   - Use PowerPoint Review > Compare on Slide 1
   - Screenshot differences
   - **Output:** `docs/27-10-25/artifacts/powerpoint_compare_signoff.md`

5. **⏰ 0.5h** - Archive visual diff findings (`adopt-template-cloning` 4.4)
   - Document findings from tasks 3-4
   - Create comparison report
   - **Output:** `docs/27-10-25/artifacts/visual_parity_complete.md`
   - **Enables:** Archival of `adopt-template-cloning-pipeline` OpenSpec change

### Post-Processing Validation
**Why Critical:** Validates 24 Oct architecture work (60x performance improvement)

6. **⏰ 1.5h** - End-to-end post-processing test (`clarify-postprocessing` Phase 2)
   - Apply: `py -m amp_automation.presentation.postprocess.cli run_20251027_135302/presentations.pptx postprocess-all`
   - Run structural validation
   - Verify: 76 slides, 0 failures, 100% success, fonts normalized
   - **Output:** `docs/27-10-25/artifacts/postprocessing_e2e_test.md`

7. **⏰ 1h** - Update PowerShell scripts (`clarify-postprocessing` Phase 2)
   - Deprecate `tools/PostProcessCampaignMerges.ps1`
   - Add deprecation notices to old COM scripts
   - Document migration to `PostProcessNormalize.ps1`
   - **Output:** Updated scripts with deprecation warnings

8. **⏰ 1h** - Update COM prohibition ADR (`clarify-postprocessing` Phase 2)
   - Clarify generation-time vs post-processing scope
   - Add "when to use COM vs python-pptx" guidance
   - Update `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`
   - **Output:** ADR with clarified scope and decision matrix

---

## 🟠 HIGH PRIORITY (Do Next - 15-18 hours)

### Test Suite & Regression Coverage
**Why High:** Prevents regressions, enables automated quality gates

9. **⏰ 2h** - Rehydrate test_tables.py (`BRAIN_RESET` NOW)
   - Fix broken tests without modifying core logic
   - Update fixtures, paths, expected values
   - **Output:** `tests/test_tables.py` passing

10. **⏰ 2h** - Rehydrate test_structural_validator.py (`BRAIN_RESET` NOW)
    - Fix validation test suite
    - Update structural expectations
    - **Output:** `tests/test_structural_validator.py` passing

11. **⏰ 3h** - Add regression tests for post-processing (`BRAIN_RESET` NOW + `complete-oct15-followups` 2.2)
    - Test merge correctness: campaign vertical, monthly/summary horizontal
    - Test font normalization: Verdana 6-7pt coverage
    - Test row formatting: MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD
    - Replace `tests/test_placeholder.py` with real tests
    - **Output:** `tests/test_post_processing.py` or extended `test_tables.py`

12. **⏰ 1h** - Pipeline hierarchy validation (`complete-oct15-followups` 2.1)
    - Execute: `python scripts/run_pipeline_local.py`
    - Validate 00-08 artifact hierarchy
    - **Output:** `docs/27-10-25/artifacts/pipeline_validation.md`

### Campaign Pagination - Analysis Phase
**Why High:** Foundation for implementation, quick data analysis

13. **⏰ 1.5h** - Analyze campaign size distribution (`implement-campaign-pagination` 1.1)
    - Read `BulkPlanData_2025_10_14.xlsx`
    - Compute: row count distribution, % >32 rows, avg by market, max size
    - **Output:** `docs/27-10-25/artifacts/campaign_size_analysis.md`

14. **⏰ 1h** - Design smart pagination algorithm (`implement-campaign-pagination` 1.2)
    - Define lookahead logic
    - Handle edge cases (exactly 32, >32, first/last campaign)
    - **Output:** Algorithm documented in `design.md`

15. **⏰ 0.5h** - Update configuration schema (`implement-campaign-pagination` 1.3)
    - Add `features.smart_campaign_pagination` (default false)
    - Add `presentation.table.min_rows_for_fresh_slide` (default 5)
    - **Output:** Updated `config/master_config.json`

### Documentation & Cleanup
**Why High:** Enables team productivity and knowledge transfer

16. **⏰ 1h** - Document COM vs python-pptx guidance (`clarify-postprocessing` Phase 2)
    - Create decision matrix in README, AGENTS.md
    - Reference ADR for details
    - **Output:** Updated README with guidance

17. **⏰ 1h** - Audit legacy path references (`complete-oct15-followups` 2.4)
    - Review all docs/ and config/ files
    - Update outdated paths
    - **Output:** Updated documentation

18. **⏰ 1h** - Commit and document session progress
    - Commit all completed work from today
    - Update `docs/27-10-25/27-10-25.md` with progress
    - Update BRAIN_RESET with checked items
    - **Output:** Clean git history, updated docs

---

## 🟡 MEDIUM PRIORITY (Next Session - 12-15 hours)

### Campaign Pagination - Implementation
**Why Medium:** Improves UX but not blocking critical work

19. **⏰ 2h** - Implement campaign lookahead logic (`implement-campaign-pagination` 2.1)
    - Add `should_start_campaign_on_fresh_slide()` function
    - Add `calculate_campaign_row_count()` helper
    - Unit test both functions
    - **Location:** `amp_automation/presentation/assembly.py`

20. **⏰ 2h** - Update slide creation logic (`implement-campaign-pagination` 2.2)
    - Modify `build_presentation()` to check lookahead
    - Track remaining capacity per slide
    - Ensure CARRIED FORWARD and GRAND TOTAL added correctly
    - **Output:** Updated assembly logic

21. **⏰ 1h** - Handle >32 row campaigns (`implement-campaign-pagination` 2.3)
    - Ensure large campaigns start on fresh slide then split
    - Verify continuation slides maintain formatting
    - Test MONTHLY TOTAL placement
    - **Output:** Edge case handling complete

22. **⏰ 0.5h** - Wire configuration toggle (`implement-campaign-pagination` 2.4)
    - Read `features.smart_campaign_pagination` from config
    - Log when smart pagination triggers
    - Default: off for backward compatibility
    - **Output:** Feature toggle functional

### Campaign Pagination - Testing
**Why Medium:** Ensures implementation correctness

23. **⏰ 1.5h** - Update existing tests (`implement-campaign-pagination` 3.1)
    - Modify `tests/test_assembly_split.py`
    - Ensure backward compatibility tests (feature off)
    - Add fixtures for various campaign sizes
    - **Output:** Updated test suite

24. **⏰ 1.5h** - Add new test coverage (`implement-campaign-pagination` 3.2)
    - Test small campaigns (5-10 rows)
    - Test medium campaigns (15-30 rows)
    - Test large campaigns (>32 rows)
    - Test edge case: exactly 32 rows
    - Test config toggle on/off
    - **Output:** Comprehensive test coverage

25. **⏰ 1h** - Regenerate test decks (`implement-campaign-pagination` 3.3)
    - Generate with feature ON and OFF
    - Compare slide counts and distributions
    - Verify no mid-campaign splits (unless >32)
    - **Output:** Test decks for comparison

26. **⏰ 0.5h** - Run structural validation (`implement-campaign-pagination` 3.4)
    - Validate both ON/OFF decks
    - Ensure requirements met
    - Verify CARRIED FORWARD and GRAND TOTAL correct
    - **Output:** Validation passing

### Campaign Pagination - Documentation
**Why Medium:** Required for feature completion

27. **⏰ 0.5h** - Update documentation (`implement-campaign-pagination` 4.1)
    - Update README with smart pagination feature
    - Update AGENTS.md with config guidance
    - Document examples
    - **Output:** Updated documentation

28. **⏰ 0.5h** - Visual inspection (`implement-campaign-pagination` 4.2)
    - Manually review generated deck
    - Check campaign boundaries
    - Verify no unexpected splits
    - **Output:** Visual sign-off

29. **⏰ 0.5h** - Performance validation (`implement-campaign-pagination` 4.3)
    - Measure generation time impact
    - Measure slide count increase
    - Document trade-offs
    - **Output:** Performance report

30. **⏰ 0.5h** - Create before/after comparison (`implement-campaign-pagination` 4.4)
    - Generate both versions
    - Document differences
    - **Output:** `docs/27-10-25/artifacts/campaign_pagination_comparison.md`

---

## 🟢 LOW PRIORITY (Future Sessions - 6-8 hours)

### Cleanup & Maintenance
**Why Low:** Nice to have, not blocking

31. **⏰ 0.5h** - Populate input/ directory (`complete-oct15-followups` 2.3)
    - Add 2-3 curated production Excel files
    - Document ingestion workflow
    - **Output:** `input/` with samples, operator guide

32. **⏰ 1h** - Audit and deprecate PowerShell scripts (`clarify-postprocessing` Phase 3)
    - List all COM-based PowerShell scripts
    - Categorize: keep, deprecate, archive
    - Add deprecation notices
    - **Output:** Updated scripts with notices

33. **⏰ 1h** - Expand Python normalization coverage (`clarify-postprocessing` Phase 3)
    - Add row height normalization (if needed)
    - Add cell margin/padding normalization
    - **Output:** Enhanced `table_normalizer.py`

34. **⏰ 1.5h** - Add merge correctness regression tests (`clarify-postprocessing` Phase 3)
    - Test generation creates expected merges
    - Test continuation slide merges
    - Test edge cases
    - **Output:** `tests/test_merge_regression.py`

35. **⏰ 1h** - Create migration guide (`clarify-postprocessing` Phase 3)
    - Document PowerShell → Python transition
    - Provide examples for common operations
    - **Output:** `docs/MIGRATION_POWERSHELL_TO_PYTHON.md`

36. **⏰ 0.5h** - Smoke test additional markets (`BRAIN_RESET` Later)
    - Run `scripts/run_pipeline_local.py` with different datasets
    - Validate across markets
    - **Output:** Multi-market validation report

37. **⏰ 1h** - Performance profiling (`BRAIN_RESET` Later)
    - Identify bottlenecks in generation/post-processing
    - Measure timing for each phase
    - **Output:** Performance profile report

### Repository Cleanup
**Why Low:** Housekeeping

38. **⏰ 0.5h** - Organize untracked diagnostic files
    - Review files in `scripts/`, `tools/debug/`
    - Archive or delete obsolete files
    - Commit useful utilities
    - **Output:** Clean repository

39. **⏰ 0.5h** - Archive completed OpenSpec changes
    - Archive `adopt-template-cloning-pipeline` (after task 5 complete)
    - Archive `complete-oct15-followups` (after all tasks complete)
    - Archive `clarify-postprocessing-architecture` (after Phase 2 complete)
    - Update `openspec/specs/` with new capabilities
    - **Output:** Archived changes, updated specs

---

## 📋 Execution Roadmap

### Session 1: Visual Parity & Quality (3-4 hours) ⚡ START HERE
**Goal:** Complete CRITICAL visual parity work, unblock archival

- [ ] Task 1: Capture Template V4 geometry constants (2h)
- [ ] Task 2: Update continuation slide layout (1.5h)
- [ ] Task 3: Run visual_diff.py (1h)
- [ ] Task 4: PowerPoint Compare (0.5h)
- [ ] Task 5: Archive findings (0.5h)
- [ ] Task 18: Commit progress (1h)

**Outcome:** Visual parity validated, `adopt-template-cloning-pipeline` ready to archive

---

### Session 2: Post-Processing & Testing (4-5 hours)
**Goal:** Validate 24 Oct work, rehydrate test suites

- [ ] Task 6: E2E post-processing test (1.5h)
- [ ] Task 7: Update PowerShell scripts (1h)
- [ ] Task 8: Update COM ADR (1h)
- [ ] Task 9: Rehydrate test_tables.py (2h)
- [ ] Task 10: Rehydrate test_structural_validator.py (2h)
- [ ] Task 11: Add regression tests (3h - can split across sessions)

**Outcome:** Post-processing validated, test suite functional

---

### Session 3: Campaign Pagination Analysis (3-4 hours)
**Goal:** Complete analysis and design phase

- [ ] Task 13: Campaign size analysis (1.5h)
- [ ] Task 14: Algorithm design (1h)
- [ ] Task 15: Config schema update (0.5h)
- [ ] Task 16: COM guidance docs (1h)
- [ ] Task 17: Audit legacy paths (1h)
- [ ] Task 12: Pipeline validation (1h)

**Outcome:** Campaign pagination ready for implementation

---

### Session 4: Campaign Pagination Implementation (6-7 hours)
**Goal:** Implement smart pagination feature

- [ ] Tasks 19-22: Implementation (5.5h total)
- [ ] Tasks 23-26: Testing (4.5h total)

**Outcome:** Feature implemented and tested

---

### Session 5: Campaign Pagination Finalization (2-3 hours)
**Goal:** Document and validate feature

- [ ] Tasks 27-30: Documentation & validation (2h total)
- [ ] Task 18: Commit and document (1h)
- [ ] Task 39: Archive completed OpenSpec changes (0.5h)

**Outcome:** Feature complete and documented

---

### Session 6: Cleanup & Low Priority (6-8 hours)
**Goal:** Technical debt and future improvements

- [ ] Tasks 31-38: Cleanup and maintenance (6-8h total)

**Outcome:** Repository clean, technical debt addressed

---

## 🎯 Success Metrics

### Immediate (Sessions 1-2)
- ✅ Visual diff shows <0.5% geometry deviation
- ✅ Post-processing: 76 slides, 0 failures, 100% success
- ✅ Test suites passing: `test_tables.py`, `test_structural_validator.py`
- ✅ 1 OpenSpec change archived (`adopt-template-cloning-pipeline`)

### Short-term (Sessions 3-5)
- ✅ Campaign pagination implemented and tested
- ✅ 0 campaigns <32 rows split across slides
- ✅ Slide count increase <15%
- ✅ Generation time increase <5%
- ✅ 2-3 more OpenSpec changes archived

### Long-term (Session 6+)
- ✅ All 4 OpenSpec changes completed and archived
- ✅ Comprehensive test coverage (>80%)
- ✅ Clean repository (no untracked diagnostic files)
- ✅ All documentation up to date

---

## 📊 Progress Tracking

### Completed Today (27 Oct 2025)
- ✅ Cleared output folder
- ✅ Generated fresh deck: `run_20251027_135302` (88 slides)
- ✅ Archived `update-table-styling-continuations` OpenSpec change
- ✅ Created `implement-campaign-pagination` OpenSpec change (Option A)
- ✅ Analyzed all OpenSpec docs and consolidated tasks
- ✅ Created master todolist (this document)

### Next Immediate Actions
1. Start Session 1: Visual Parity work (Task 1: Capture geometry constants)
2. Commit today's work before ending session
3. Update `docs/27-10-25/27-10-25.md` with end-of-day summary

---

## 📁 Key Files & Commands

### Generate Deck
```bash
py -m amp_automation.cli.main --excel "template\BulkPlanData_2025_10_14.xlsx" --template "template\Template_V4_FINAL_071025.pptx" --output "output\presentations"
```

### Validate Structure
```bash
py tools\validate_structure.py "output\presentations\run_YYYYMMDD_HHMMSS\presentations.pptx" --excel "template\BulkPlanData_2025_10_14.xlsx"
```

### Post-Process
```bash
py -m amp_automation.presentation.postprocess.cli "output\presentations\run_YYYYMMDD_HHMMSS\presentations.pptx" postprocess-all
```

### Run Tests
```bash
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD="1"; py -m pytest tests\
```

### Visual Diff
```bash
py tools\visual_diff.py --template "template\Template_V4_FINAL_071025.pptx" --generated "output\presentations\run_YYYYMMDD_HHMMSS\presentations.pptx"
```

---

## 🔗 Related Documentation

- **Active OpenSpec Changes:** `openspec/changes/`
- **Archived Changes:** `openspec/changes/archive/`
- **Session Docs:** `docs/27-10-25/`
- **Brain Reset:** `docs/27-10-25/BRAIN_RESET_271025.md`
- **COM Prohibition ADR:** `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`
- **Bootstrap Prompt:** `docs/27-10-25/artifacts/01-fresh-start_bootstrap_prompt.md`
- **Pending Tasks Analysis:** `docs/27-10-25/artifacts/02-pending_tasks_summary.md`
- **Active Tasks Breakdown:** `docs/27-10-25/artifacts/03-active_openspec_tasks.md`
- **Campaign Pagination Analysis:** `docs/27-10-25/artifacts/04-campaign_pagination_task.md`

---

**Last Updated:** 2025-10-27 14:05
**Next Review:** After Session 1 completion
