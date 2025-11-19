# Pending Tasks Summary - 27 Oct 2025

Generated: 2025-10-27 13:53
Fresh deck: `output/presentations/run_20251027_135302/presentations.pptx` (88 slides, 565KB)

---

## CRITICAL FINDINGS

**No specs exist in `openspec/specs/`** - All OpenSpec changes are still in proposal/implementation stage and have NOT been archived. This indicates several completed initiatives need formal archival.

---

## OpenSpec Change Status Overview

### 1. ✅ MOSTLY COMPLETE - `adopt-template-cloning-pipeline`
**Status:** Implementation complete, archival pending

**Completed Tasks (7/8):**
- [x] 1.1-1.4: Clone-based rendering pipeline fully implemented
- [x] 2.1-2.3: Verification and regression safety nets in place
- [x] 3.1-3.2: Output & packaging complete
- [x] 4.1-4.3: Visual parity closure (background/fills, footer alignment, legend RGB)

**PENDING (1/8):**
- [ ] **4.4: Export template/generated decks, rerun visual diff, perform Zen MCP + PowerPoint Review > Compare, archive findings**
  - Status: Blocked on multi-slide template imagery
  - Priority: HIGH (foundational for quality assurance)
  - Files: `tools/visual_diff.py`, `Template_V4_FINAL_071025.pptx`, generated deck

**Action Required:**
1. Run visual diff on fresh deck (`run_20251027_135302`) vs template
2. Perform Zen MCP Compare evidence capture
3. Archive findings in `docs/27-10-25/artifacts/`
4. **Archive this OpenSpec change** once visual diff complete

---

### 2. ⚠️ PARTIALLY COMPLETE - `complete-oct15-followups`
**Status:** 0/7 tasks complete, HIGH PRIORITY

**All Tasks PENDING:**
- [ ] **1.1: Capture Template V4 column widths and table bounds as constants**
  - Location: `assembly.py`, `tables.py` modules
  - Impact: Ensures geometry consistency across regenerations

- [ ] **1.2: Update continuation slide layout to honor exact Template V4 geometry**
  - Files: `amp_automation/presentation/assembly.py`
  - Critical for: Position, width, row heights on continuation slides

- [ ] **1.3: Regenerate presentation** ✅ **JUST COMPLETED**
  - Fresh deck: `run_20251027_135302/presentations.pptx` (88 slides)

- [ ] **1.4: Run visual_diff.py and confirm metrics trend toward zero**
  - Tool: `tools/visual_diff.py`
  - Archive: Comparison artifacts in `docs/27-10-25/artifacts/`

- [ ] **1.5: Manual PowerPoint Review→Compare sign-off**
  - Evidence: Zen MCP screenshots or Compare output

- [ ] **2.1: Execute pipeline with representative workbooks to validate 00-08 artifact hierarchy**
  - Script: `python scripts/run_pipeline_local.py`
  - Validates: Full pipeline orchestration

- [ ] **2.2: Replace tests/test_placeholder.py with regression coverage**
  - New tests: Continuation logic, hierarchy validation
  - Files: `tests/test_placeholder.py` → new test suite

- [ ] **2.3: Populate input/ with curated production samples**
  - Document: Ingestion workflow for operators

- [ ] **2.4: Audit repository docs/configs for legacy path references**
  - Update: All documentation and config files

**Priority Assessment:**
- **IMMEDIATE:** Tasks 1.1, 1.2, 1.4, 1.5 (template geometry and visual parity)
- **SHORT-TERM:** Tasks 2.1, 2.2 (pipeline validation and tests)
- **MEDIUM-TERM:** Tasks 2.3, 2.4 (documentation and workflow)

---

### 3. ✅ COMPLETE - `update-table-styling-continuations`
**Status:** All 10/10 tasks complete

**Recent Completion (21 Oct 2025):**
- [x] All implementation tasks (1.1-1.5)
- [x] Validation (2.1)
- [x] Slide 1 parity follow-up (3.1-3.4)

**Action Required:**
- **Archive this OpenSpec change** to `openspec/changes/archive/2025-10-21-update-table-styling-continuations/`
- Update `openspec/specs/` if new capabilities were added

---

### 4. ⚠️ IN PROGRESS - `clarify-postprocessing-architecture`
**Status:** Phase 1 complete (8/8), Phase 2 pending (4 tasks), Phase 3 future (4 tasks)

**✅ Phase 1: Discovery & Documentation (COMPLETED 24 Oct 2025)**
- [x] Implement Python cell merge operations (commit d3e2b98)
- [x] Test Python post-processing on 88-slide deck
- [x] Analyze architecture and document findings (commit 8320c3f)
- [x] Update project documentation (commit 3c54b1b)
- [x] Create OpenSpec proposal

**⏭️ Phase 2: Integration & Testing (PENDING - NEXT SESSION)**
- [ ] **Update PowerShell scripts to call Python CLI**
  - Files: `tools/PostProcessCampaignMerges.ps1`
  - Action: Modify to use Python for normalization, document merge operations as edge-case repairs
  - Status: **PARTIALLY DONE** - `PostProcessNormalize.ps1` exists but old script needs deprecation

- [ ] **Run end-to-end pipeline test**
  - Generate fresh deck ✅ **JUST COMPLETED** (`run_20251027_135302`)
  - Apply Python normalization
  - Run structural validation
  - Verify merge correctness and table formatting

- [ ] **Update COM prohibition ADR**
  - File: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`
  - Clarify: COM restriction scope (post-processing vs generation)

- [ ] **Add guidance on when to use COM vs python-pptx**
  - Location: ADR, README, AGENTS.md

**⏭️ Phase 3: Cleanup & Standardization (FUTURE)**
- [ ] Audit and deprecate redundant PowerShell scripts
- [ ] Expand Python normalization coverage (row height, margins/padding)
- [ ] Add regression tests for merge correctness
- [ ] Create migration guide

**Success Metrics:**
- ✅ Python implementation completed and committed
- ✅ Architecture discovery documented
- ✅ Performance validated (<1 minute for 88 slides)
- ⏭️ End-to-end test passes without errors
- ⏭️ PowerShell scripts updated or deprecated
- ⏭️ Documentation reflects new architecture

---

## IMMEDIATE ACTION ITEMS (Prioritized)

### Priority 1: Visual Parity & Quality Assurance
1. **Run visual diff on fresh deck** (`run_20251027_135302`)
   - Tool: `tools/visual_diff.py`
   - Compare: Generated vs `Template_V4_FINAL_071025.pptx`
   - Blockers: Need to understand visual_diff.py capabilities

2. **Perform Zen MCP + PowerPoint Review > Compare**
   - Evidence: Screenshot Slide 1 comparison
   - Archive: `docs/27-10-25/artifacts/visual_diff_results.md`

3. **Template geometry constants capture** (`complete-oct15-followups` task 1.1)
   - Extract: Column widths, table bounds from Template V4
   - Store: Constants in `assembly.py`/`tables.py` or config

4. **Continuation slide geometry alignment** (`complete-oct15-followups` task 1.2)
   - Update: `amp_automation/presentation/assembly.py`
   - Ensure: Position, width, row heights match template exactly

### Priority 2: Post-Processing Pipeline Validation
5. **End-to-end pipeline test** (`clarify-postprocessing-architecture` Phase 2)
   - Fresh deck: ✅ Generated (`run_20251027_135302`)
   - Apply: Python post-processing (`postprocess-all`)
   - Validate: Structural validation + font checks

6. **Deprecate old PowerShell COM scripts**
   - Audit: `tools/PostProcessCampaignMerges.ps1` and related
   - Update: Documentation to use `PostProcessNormalize.ps1`
   - Archive: Old scripts with deprecation notices

### Priority 3: Testing & Regression Coverage
7. **Rehydrate pytest test suites** (from BRAIN_RESET)
   - Files: `tests/test_tables.py`, `tests/test_structural_validator.py`
   - Fix: Broken tests without modifying core logic

8. **Add regression tests** (`complete-oct15-followups` task 2.2)
   - Replace: `tests/test_placeholder.py`
   - Cover: Merge correctness, continuation logic, font normalization

### Priority 4: Pipeline Orchestration
9. **Execute full pipeline validation** (`complete-oct15-followups` task 2.1)
   - Script: `python scripts/run_pipeline_local.py`
   - Validate: 00-08 artifact hierarchy

10. **Populate input/ directory** (`complete-oct15-followups` task 2.3)
    - Add: Curated production samples
    - Document: Operator ingestion workflow

### Priority 5: Documentation & Archival
11. **Update COM prohibition ADR** (`clarify-postprocessing-architecture` Phase 2)
    - Clarify: Generation-time vs post-processing scope
    - Add: When to use COM vs python-pptx guidance

12. **Archive completed OpenSpec changes**
    - Archive: `update-table-styling-continuations` (all tasks complete)
    - Archive: `adopt-template-cloning-pipeline` (after visual diff task 4.4)
    - Update: `openspec/specs/` with new capabilities

13. **Audit legacy path references** (`complete-oct15-followups` task 2.4)
    - Review: All docs/ and config/ files
    - Update: Outdated paths and references

---

## BLOCKERS & DEPENDENCIES

### Active Blockers
1. **Visual diff multi-slide imagery** - Blocks `adopt-template-cloning-pipeline` task 4.4
   - Need: Investigation into `tools/visual_diff.py` capabilities
   - Alternative: Manual PowerPoint Review > Compare workflow

### Dependencies
1. **Visual parity verification** (Priority 1) → Required before archiving `adopt-template-cloning-pipeline`
2. **Template geometry constants** (Priority 1) → Required for continuation slide alignment
3. **Post-processing validation** (Priority 2) → Required before archiving `clarify-postprocessing-architecture`
4. **Test suite rehydration** (Priority 3) → Required for regression coverage

---

## RECOMMENDED NEXT STEPS

### Option A: Complete Visual Parity (Highest Impact)
**Timeline:** 2-3 hours
1. Investigate `tools/visual_diff.py` functionality
2. Run visual diff on `run_20251027_135302` vs template
3. Perform manual PowerPoint Review > Compare for Slide 1
4. Document findings and archive evidence
5. Archive `adopt-template-cloning-pipeline` OpenSpec change

**Benefits:**
- Unblocks archival of major OpenSpec change
- Establishes visual diff baseline for future work
- Validates template fidelity (business-critical)

### Option B: Complete Post-Processing Validation (Quick Win)
**Timeline:** 1-2 hours
1. Run Python post-processing on fresh deck
2. Validate structural correctness and font normalization
3. Update PowerShell script deprecation notices
4. Update COM prohibition ADR with clarifications
5. Complete Phase 2 of `clarify-postprocessing-architecture`

**Benefits:**
- Validates recent architecture work (24 Oct session)
- Quick completion of in-progress OpenSpec change
- Establishes post-processing baseline

### Option C: Systematic Approach (Complete Oct15 Followups)
**Timeline:** 4-6 hours
1. Capture Template V4 geometry constants (1-2 hours)
2. Update continuation slide layout logic (1-2 hours)
3. Run visual diff validation (1 hour)
4. Execute pipeline validation (1 hour)
5. Document and archive results

**Benefits:**
- Addresses oldest pending work (Oct 15 followups)
- Comprehensive completion of high-priority change
- Sets foundation for continuation slide quality

---

## SUMMARY STATISTICS

- **Total OpenSpec Changes:** 4
- **Complete (pending archival):** 1 (`update-table-styling-continuations`)
- **In Progress:** 3
- **Total Pending Tasks:** 16 (across all changes)
- **High Priority Tasks:** 6 (visual parity, geometry constants, post-processing validation)
- **Blockers:** 1 (visual diff multi-slide imagery)

**Fresh Deck Generated:**
- Path: `output/presentations/run_20251027_135302/presentations.pptx`
- Size: 565KB (88 slides)
- Status: Ready for validation and testing

---

**Recommendation:** Start with **Option A (Visual Parity)** to unblock the largest completed work and establish quality baselines, then move to **Option B (Post-Processing)** for a quick win validating recent work.
