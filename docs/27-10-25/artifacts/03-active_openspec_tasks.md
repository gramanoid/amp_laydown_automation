# Active OpenSpec Tasks - 27 Oct 2025

Generated: 2025-10-27 13:58
Archived: `update-table-styling-continuations` → `archive/2025-10-21-update-table-styling-continuations/`

---

## 1. adopt-template-cloning-pipeline

**Status:** 7/8 tasks complete (88%)
**Priority:** HIGH - One critical task remaining

### ✅ Completed Tasks (7)

#### 1. Clone-Based Rendering Pipeline
- [x] 1.1 Analyze master slide structure (table/tiles/legends) and document shape IDs required for cloning
- [x] 1.2 Implement cloning helpers to duplicate target shapes onto new slides while preserving geometry and styling
- [x] 1.3 Replace manual table construction with data population into cloned table cells/shapes
- [x] 1.4 Wire configuration toggle to enable/disable clone-based workflow and phase out AutoPPTX adapter once parity confirmed
  - Implemented via `features.clone_pipeline_enabled` (default `true`) in `master_config.json`
  - Regression coverage: `tests/test_autopptx_fallback.py` exercises AutoPPTX path

#### 2. Verification & Regression Safety Nets
- [x] 2.1 Extend visual diff runner to compare multiple slides per deck and surface geometry mismatches
  - Implemented COM-based export + PIL metrics
  - Note: Needs follow-up to resolve PowerPoint export failures on regenerated deck
- [x] 2.2 Add unit/integration tests covering clone pipeline, including fixture slides and structural assertions
  - Current coverage: `tests/test_tables.py`, `tests/test_assembly_split.py` pass
  - Follow-up tracked: add tile-format tests
- [x] 2.3 Update documentation/logging to reflect new workflow and capture validation steps for operators
  - Structural contract captured in docs/17_10_25 + docs/20_10_25
  - Enforced via `config/structural_contract.json` + `tools/validate_structure.py`
  - Outstanding issues: media ordering, slide-level GRAND TOTAL, footnote date

#### 3. Output & Packaging
- [x] 3.1 Flatten run output structure to avoid nested `.../run_<ts>/output/<file>` when `--output` includes a path
- [x] 3.2 Diagnose and eliminate PowerPoint "Repair" prompt (inspect low-shape slides; adjust XML insertion point if needed)

### ⏭️ PENDING Tasks (1)

#### 4. Visual Parity Closure
- [ ] **4.4 Export template/generated decks, rerun visual diff, perform PowerPoint Review > Compare, and archive findings**
  - **Files:**
    - `tools/visual_diff.py` - Visual diff tool (needs investigation for multi-slide capability)
    - `template/Template_V4_FINAL_071025.pptx` - Master template
    - `output/presentations/run_20251027_135302/presentations.pptx` - Fresh generated deck
  - **Steps:**
    1. Run `tools/visual_diff.py` on fresh deck vs template
    2. Perform manual PowerPoint Review > Compare on Slide 1 (critical geometry check)
    3. Document geometry deviations (target: <0.5%)
    4. Archive findings in `docs/27-10-25/artifacts/visual_parity_results.md`
  - **Blockers:** None (previously blocked on multi-slide template imagery - investigate visual_diff.py capabilities)
  - **Acceptance:** Visual diff shows <0.5% geometry deviation from template, PowerPoint Compare evidence captured
  - **Priority:** HIGH - Required before archiving this OpenSpec change

---

## 2. complete-oct15-followups

**Status:** 1/7 tasks complete (14%)
**Priority:** HIGH - Template geometry and validation work

### ⏭️ PENDING Tasks (7)

#### 1. Template Geometry Alignment (4 tasks)
- [ ] **1.1 Capture Template V4 column widths and table bounds as constants shared across assembly/tables modules**
  - **Files:**
    - `amp_automation/presentation/assembly.py`
    - `amp_automation/presentation/tables.py`
    - `config/master_config.json` (or new constants file)
  - **Action:** Extract exact EMU values from Template V4 for:
    - Column widths (all columns A-L)
    - Table outer bounds (x, y, width, height)
    - Row heights (header, body, MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD)
  - **Deliverable:** Constants module or config section with template geometry
  - **Priority:** HIGH - Foundation for continuation slide alignment

- [ ] **1.2 Update continuation slide layout logic to honor exact Template V4 geometry (position, width, row heights)**
  - **Files:** `amp_automation/presentation/assembly.py` (continuation slide generation)
  - **Action:** Apply Template V4 constants captured in 1.1 to continuation slide builder
  - **Dependencies:** Requires 1.1 completion
  - **Priority:** HIGH - Ensures pixel-perfect continuation slides

- [ ] **1.3 Regenerate a presentation via CLI**
  - ✅ **COMPLETED:** Fresh deck generated at `output/presentations/run_20251027_135302/presentations.pptx`
  - Command used: `py -m amp_automation.cli.main --excel template\BulkPlanData_2025_10_14.xlsx --template template\Template_V4_FINAL_071025.pptx --output output\presentations`

- [ ] **1.4 Run tools/visual_diff.py and confirm metrics trend toward zero; archive comparison artifacts**
  - **Files:** `tools/visual_diff.py`
  - **Input:** `run_20251027_135302/presentations.pptx` vs `template/Template_V4_FINAL_071025.pptx`
  - **Action:**
    1. Execute visual diff tool
    2. Analyze metrics (pixel differences, geometry deviations)
    3. Confirm trending toward zero (improvement over previous runs)
    4. Archive results: `docs/27-10-25/artifacts/visual_diff_oct15_followup.md`
  - **Dependencies:** Requires 1.3 completion (✅ done)
  - **Priority:** HIGH - Validates geometry alignment work

- [ ] **1.5 Perform manual PowerPoint Review→Compare against the master template and capture sign-off**
  - **Files:**
    - Template: `template/Template_V4_FINAL_071025.pptx`
    - Generated: `output/presentations/run_20251027_135302/presentations.pptx`
  - **Action:**
    1. Open both files in PowerPoint
    2. Use Review > Compare feature
    3. Focus on Slide 1 geometry (table position, column widths, row heights)
    4. Screenshot differences
    5. Document sign-off in `docs/27-10-25/artifacts/powerpoint_compare_signoff.md`
  - **Dependencies:** Requires 1.4 completion
  - **Priority:** HIGH - Business sign-off for visual parity

#### 2. Pipeline Hierarchy Validation (3 tasks)
- [ ] **2.1 Execute python scripts/run_pipeline_local.py with representative workbooks to validate the 00-08 artifact hierarchy**
  - **Files:** `scripts/run_pipeline_local.py`
  - **Action:**
    1. Verify script exists and is executable
    2. Run with test workbooks
    3. Validate artifact hierarchy (00-raw, 01-normalized, ..., 08-final)
    4. Document results in `docs/27-10-25/artifacts/pipeline_validation.md`
  - **Priority:** MEDIUM - Full pipeline orchestration validation

- [ ] **2.2 Replace tests/test_placeholder.py with regression coverage exercising the new hierarchy and continuation logic**
  - **Files:**
    - `tests/test_placeholder.py` (to replace)
    - New: `tests/test_continuation_slides.py` or extend `tests/test_assembly_split.py`
  - **Action:**
    1. Create regression tests for:
       - Continuation slide generation (split logic)
       - Template geometry adherence
       - Cell merges on continuation slides
       - CARRIED FORWARD row propagation
    2. Remove or update `test_placeholder.py`
  - **Priority:** MEDIUM - Automated regression coverage

- [ ] **2.3 Populate input/ with curated production samples or document ingestion workflow for operators**
  - **Files:**
    - Create: `input/` directory with sample workbooks
    - Document: `docs/operator_guide.md` or similar
  - **Action:**
    1. Curate 2-3 representative production Excel files
    2. Add to `input/` directory
    3. Document ingestion workflow:
       - File naming conventions
       - Column mapping expectations
       - Validation steps
  - **Priority:** LOW - Operator documentation

- [ ] **2.4 Audit repository docs/configs for legacy path references and update as needed**
  - **Files:** All `docs/**/*.md`, `config/**/*.json`, `README.md`, `AGENTS.md`
  - **Action:**
    1. Search for outdated path references (old output structure, deprecated scripts)
    2. Update to current conventions
    3. Remove references to archived/deprecated features
  - **Priority:** LOW - Documentation cleanup

---

## 3. clarify-postprocessing-architecture

**Status:** Phase 1 complete (8/8), Phase 2 pending (4 tasks), Phase 3 future (4 tasks)
**Priority:** MEDIUM - Validation and cleanup work

### ✅ Phase 1: Discovery & Documentation (COMPLETED 24 Oct 2025)
- [x] Implement Python cell merge operations (commit d3e2b98)
- [x] Test Python post-processing on 88-slide deck (30 seconds, 60x faster than COM)
- [x] Analyze architecture and document findings (commit 8320c3f)
- [x] Update project documentation (commit 3c54b1b)
- [x] Create OpenSpec proposal

### ⏭️ Phase 2: Integration & Testing (PENDING - NEXT SESSION)

- [ ] **Update PowerShell scripts to call Python CLI**
  - **Files:**
    - `tools/PostProcessCampaignMerges.ps1` (to deprecate or update)
    - `tools/PostProcessNormalize.ps1` (already exists, verify usage)
  - **Action:**
    1. Audit all PowerShell scripts in `tools/` for COM-based post-processing
    2. Update scripts to call Python CLI: `py -m amp_automation.presentation.postprocess.cli`
    3. Add deprecation notices to old COM-based scripts
    4. Document merge operations as edge-case repair tools (not primary workflow)
  - **Status:** PARTIALLY DONE - `PostProcessNormalize.ps1` exists, old scripts need deprecation
  - **Priority:** MEDIUM - Cleanup and standardization

- [ ] **Run end-to-end pipeline test**
  - **Files:**
    - Fresh deck: ✅ `output/presentations/run_20251027_135302/presentations.pptx`
    - Python CLI: `amp_automation/presentation/postprocess/cli.py`
    - Validator: `tools/validate_structure.py`
  - **Action:**
    1. Apply Python post-processing: `py -m amp_automation.presentation.postprocess.cli output/presentations/run_20251027_135302/presentations.pptx postprocess-all`
    2. Run structural validation: `py tools/validate_structure.py output/presentations/run_20251027_135302/presentations.pptx --excel template/BulkPlanData_2025_10_14.xlsx`
    3. Verify merge correctness (campaign vertical, monthly/summary horizontal)
    4. Verify table formatting (fonts, cell fills, row heights)
    5. Document results: `docs/27-10-25/artifacts/postprocessing_e2e_test.md`
  - **Expected Results:**
    - 76 slides processed
    - 0 failures
    - 100% success rate
    - All fonts: Verdana 6-7pt (GRAND TOTAL 6pt with zero margins)
  - **Priority:** HIGH - Validates 24 Oct architecture work

- [ ] **Update COM prohibition ADR**
  - **Files:** `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`
  - **Action:**
    1. Add section clarifying COM scope:
       - ✅ Generation-time merges (acceptable - not bulk operations)
       - 🚫 Post-processing bulk operations (prohibited)
    2. Add guidance: When to use COM vs python-pptx
       - COM: File I/O, exports, single operations
       - python-pptx: All bulk table operations, loops over cells/rows
    3. Update examples with generation vs post-processing context
  - **Priority:** MEDIUM - Documentation clarity

- [ ] **Add guidance on when to use COM vs python-pptx**
  - **Files:**
    - `README.md` - Quick reference section
    - `AGENTS.md` - OpenSpec instructions
    - `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` - Detailed guidance
  - **Action:**
    1. Create decision matrix:
       - Use COM: File operations, single-shape edits, features not in python-pptx
       - Use python-pptx: Bulk operations, loops, table manipulations
    2. Add examples to each documentation file
    3. Reference ADR for detailed rationale
  - **Priority:** LOW - Developer guidance

### ⏭️ Phase 3: Cleanup & Standardization (FUTURE)

- [ ] **Audit and deprecate redundant PowerShell scripts**
  - **Files:** All `tools/*.ps1` scripts
  - **Action:**
    1. List all PowerShell scripts with COM automation
    2. Categorize: Keep (necessary COM), Deprecate (replaced by Python), Archive (obsolete)
    3. Add deprecation notices to headers
    4. Update runbooks to reference new Python CLI
  - **Priority:** LOW - Technical debt cleanup

- [ ] **Expand Python normalization coverage**
  - **Files:** `amp_automation/presentation/postprocess/table_normalizer.py`
  - **Action:**
    1. Add row height normalization (if needed based on validation results)
    2. Add cell margin/padding normalization
    3. Add font consistency checks (already have normalize_table_fonts)
    4. Document new operations in CLI help text
  - **Priority:** LOW - Future enhancements

- [ ] **Add regression tests for merge correctness**
  - **Files:**
    - New: `tests/test_post_processing.py`
    - Or extend: `tests/test_tables.py`
  - **Action:**
    1. Test generation creates expected merges (campaign vertical, monthly/summary horizontal)
    2. Test merge behavior on continuation slides
    3. Test edge cases (single-row campaigns, etc.)
    4. Test no rogue merges (only allowlist: MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD)
  - **Priority:** LOW - Automated regression coverage

- [ ] **Create migration guide**
  - **Files:** New: `docs/MIGRATION_POWERSHELL_TO_PYTHON.md`
  - **Action:**
    1. Document transition from PowerShell COM to Python
    2. Provide examples for common operations:
       - Font normalization: PowerShell COM vs Python
       - Cell merges: PowerShell COM vs Python
       - Table formatting: PowerShell COM vs Python
    3. Update README with migration guide link
    4. Update AGENTS.md with workflow changes
  - **Priority:** LOW - Developer onboarding

---

## SUMMARY

### Task Counts by OpenSpec Change
- **adopt-template-cloning-pipeline:** 1 pending (HIGH priority - visual diff)
- **complete-oct15-followups:** 7 pending (4 HIGH, 2 MEDIUM, 1 LOW)
- **clarify-postprocessing-architecture:** 8 pending (1 HIGH in Phase 2, 4 MEDIUM in Phase 2, 3 LOW in Phase 3)

### Total: 16 pending tasks

### High Priority Tasks (5)
1. Visual diff + PowerPoint Compare (adopt-template-cloning-pipeline 4.4)
2. Capture Template V4 geometry constants (complete-oct15-followups 1.1)
3. Update continuation slide layout (complete-oct15-followups 1.2)
4. Run visual_diff.py validation (complete-oct15-followups 1.4)
5. End-to-end post-processing test (clarify-postprocessing-architecture Phase 2)

### Medium Priority Tasks (5)
6. Manual PowerPoint Compare sign-off (complete-oct15-followups 1.5)
7. Pipeline hierarchy validation (complete-oct15-followups 2.1)
8. Regression test coverage (complete-oct15-followups 2.2)
9. Update PowerShell scripts (clarify-postprocessing-architecture Phase 2)
10. Update COM prohibition ADR (clarify-postprocessing-architecture Phase 2)

### Low Priority Tasks (6)
11. Populate input/ directory (complete-oct15-followups 2.3)
12. Audit legacy path references (complete-oct15-followups 2.4)
13. COM vs python-pptx guidance (clarify-postprocessing-architecture Phase 2)
14-16. Phase 3 cleanup tasks (clarify-postprocessing-architecture)

---

## RECOMMENDED EXECUTION ORDER

### Session 1: Visual Parity & Geometry (3-4 hours)
1. Capture Template V4 geometry constants (1.1) - 1 hour
2. Update continuation slide layout (1.2) - 1 hour
3. Run visual_diff.py validation (1.4) - 30 min
4. Perform PowerPoint Compare (1.5) - 30 min
5. Visual diff for adopt-template-cloning (4.4) - 1 hour

**Outcome:** Complete visual parity verification, unblock archival of adopt-template-cloning-pipeline

### Session 2: Post-Processing Validation (2-3 hours)
1. Run end-to-end post-processing test - 1 hour
2. Update PowerShell scripts deprecation - 30 min
3. Update COM prohibition ADR - 30 min
4. Document results and findings - 30 min

**Outcome:** Validate 24 Oct architecture work, complete Phase 2 of clarify-postprocessing-architecture

### Session 3: Testing & Pipeline Validation (3-4 hours)
1. Pipeline hierarchy validation (2.1) - 1 hour
2. Regression test coverage (2.2) - 2 hours
3. Documentation updates - 1 hour

**Outcome:** Establish automated regression coverage, validate full pipeline orchestration

### Session 4: Cleanup & Documentation (2-3 hours)
1. Populate input/ directory (2.3)
2. Audit legacy path references (2.4)
3. Phase 3 cleanup tasks (if time permits)

**Outcome:** Complete documentation, cleanup technical debt
