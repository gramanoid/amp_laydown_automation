# Bootstrap Prompt for 28-10-25 Session

## ASSUME
- The reconciliation validation failures (630/631 checks) indicate a systematic Excel data source mismatch rather than validator code issues
- Slide 1 geometry parity can be established using visual diff + PowerPoint Compare (Zen MCP) without manual layout work
- Test suite rehydration and campaign pagination are independent tasks suitable for parallel work
- All 27-10-25 commits are verified and production-ready

## MISSING
- Direct evidence of Excel market/brand name mapping discrepancies (requires debug output analysis)
- Baseline visual diff report for Slide 1 (template vs generated deck)
- Current test failure reasons and remediation steps
- Campaign pagination design specification (Phase 3-4 cancelled, but approach needs documentation)

---

## Zero-Question Rule

**When uncertain about direction, do NOT ask the user.** Instead:
1. Propose two concrete options with pros/cons
2. Choose the option with lower risk + faster feedback cycle
3. Document the assumption and proceed

**Example:** If reconciliation findings are ambiguous, propose investigating Excel column mapping first (faster, data-driven) vs investigating validator assumptions (slower, speculative).

---

## Alignment Check

**One-line overview:**
AMP Automation converts Lumina Excel exports into pixel-accurate PowerPoint decks mirroring a template while binding financial/media metrics; validation infrastructure is complete, now focusing on data source reconciliation and geometry parity.

**Three-bullet summary:**

- **What:** 8-step Python post-processing pipeline (unmerge → delete-CF → merge campaigns/media/monthly/summary → normalize fonts) validated on 144-slide deck; comprehensive data validation suite (4 modules + unified report) passing all checks
- **NOW Tasks:** Reconcile 630/631 validation failures against Excel market/brand mapping; establish Slide 1 geometry baseline with visual diff; rehydrate test suites for regression coverage
- **Biggest Risk:** Reconciliation failures are systematic (not stochastic), pointing to undiagnosed Excel data source mismatch. Mitigation: Analyze debug output immediately to isolate market/brand name discrepancy, document findings in `docs/NOW_TASKS.md`

---

## Brain Reset Digest

### Session Overview
**Date:** 28-10-25 (UAE/AST, UTC+04)
**Duration:** Fresh start, carrying forward 27-10-25 completions
**Scope:** Reconciliation investigation, Slide 1 geometry parity, test suite rehydration
**Last Updated:** 2025-10-28 00:00 AST

### Work Completed (27-10-25 Session)

#### Formatting Phase
✅ Timestamp fix: Local system time (AST) across `cli/main.py`, `utils/logging.py`, `assembly.py`
✅ Smart line breaking: `_smart_line_break()` for campaign names (dash handling, word-count-based splits)
✅ Media channel merging: Vertical cell merging for TELEVISION, DIGITAL, OOH, OTHER
✅ Font corrections: 6pt body/bottom rows, 7pt BRAND TOTAL, 6pt campaign column
✅ Campaign text wrapping: Hyphens removed + column widened to 1,000,000 EMU

#### Validation Phase
✅ Structural validator enhanced: Last-slide-only shapes support (BRAND TOTAL on final slides only)
✅ Data validation suite expanded: 1,200+ lines across 5 modules
- `data_accuracy.py` (160 lines): Numerical accuracy
- `data_format.py` (280 lines): Format/style validation (1,575 checks/deck)
- `data_completeness.py` (170 lines): Required data presence
- `validation/utils.py` (190 lines): Shared utilities
- `tools/validate_all_data.py` (250 lines): Unified report generator

✅ Validator bugs fixed: Table cell indexing, metadata filtering
✅ All validators tested on 144-slide production deck - **PASS status**

#### Repository Cleanup (Tier 6)
✅ Tools reorganized: `validate/` and `verify/` subdirectories created
✅ Archive documentation: `tools/archive/README_ARCHIVE.md`, historical session docs
✅ Logs restructured: 196 production logs reorganized by date (2025-10-14 through 2025-10-27)

### Current State

**What's Working:**
- 8-step post-processing workflow: 100% success rate on 144-slide deck
- All formatting improvements committed and verified
- Structural validator handles final-slide indicators correctly
- Data validation suite running with comprehensive coverage
- Production deck available: `run_20251027_215710` (144 slides, all improvements)

**What's Broken:**
- Reconciliation validation: 630/631 checks failing (likely Excel data source mismatch)
- Test suites need rehydration (`tests/test_tables.py`, `tests/test_structural_validator.py`)
- Slide 1 geometry parity not yet established (no visual diff baseline)

**What's Outstanding:**
- [ ] Reconciliation data source investigation (Excel market/brand mapping analysis)
- [ ] Slide 1 EMU/legend parity verification (visual diff + PowerPoint Compare)
- [ ] Test suite rehydration with current pipeline state
- [ ] Campaign pagination design refinement

### Purpose

Enable rapid iteration on presentation quality by:
1. **Validating data integrity** at multiple levels (accuracy, format, completeness, reconciliation)
2. **Automating formatting** through Python post-processing (no COM, 60x faster than PowerShell)
3. **Establishing visual parity** between generated and template decks via automated diff tooling
4. **Preventing regressions** through comprehensive test coverage

### Next Steps

**Immediate (28-10-25):**
1. Investigate reconciliation failures: Analyze debug output from Excel market/brand mapping
   - Compare Excel column indices (campaign_name=83, funnel_stage=95) against presentation values
   - Identify systematic mismatches (case sensitivity? extra whitespace? missing markets?)
   - Document findings in `docs/NOW_TASKS.md` with recommendations
2. Establish Slide 1 geometry baseline: Run visual diff (`tools/visual_diff.py`) on template vs latest deck
   - Capture PowerPoint Compare evidence (Zen MCP)
   - Flag EMU/legend discrepancies for remediation

**Near-term (28-10-25 evening/29-10-25):**
1. Test suite rehydration: Fix `test_tables.py` and `test_structural_validator.py` for current pipeline
2. Add regression tests: Campaign merging, font normalization, row height consistency
3. Refine campaign pagination: Document design approach to prevent across-slide splits

**Later:**
1. Establish visual diff workflow with evidence capture (Zen MCP)
2. Implement automated regression detection (catch merge/font regressions pre-ship)
3. Expand Python normalization (row height, cell margins/padding if needed)
4. Validate pipeline with additional market data sets

### Important Notes

**Critical Architecture (DO NOT VIOLATE):**
- 🚫 **COM prohibited** for bulk table operations (60x performance penalty, 10+ hours vs 10 minutes)
- ✅ **Python required** for all bulk merging, styling, and normalization
- ✅ COM only for: File I/O, exports, features unavailable in python-pptx

**Horizontal Merge Allowlist:**
- MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD (all other merges = regression)

**Template Fidelity:**
- Never modify colors, fonts, positions outside clone operations
- Maintain EMU geometry and centered alignment faithful to `Template_V4_FINAL_071025.pptx`
- Font sizes: 6pt body/bottom, 7pt BRAND TOTAL, 6pt campaign column (non-negotiable)

**Data Quality:**
- Percent stats tolerance ~0.5%; use dash (`-`) for missing metrics
- Lumina Excel: Fixed column indices (campaign=83, funnel=95, etc.)
- Template V4: Up to 32 body rows per slide + carried totals + slide GRAND TOTAL on continuation slides

**Reconciliation Context:**
- Reconciliation validator comparing Excel data source against generated presentation values
- 630/631 failures suggest systematic mismatch (not random errors)
- Likely causes: Excel column mapping incorrect, market/brand names have case/whitespace differences, or validator has incorrect assumptions

### Session Metadata
- **Latest Deck:** `output/presentations/run_20251027_215710/AMP_Presentation_20251027_215710.pptx` (144 slides)
- **Latest Commit:** `d655002` - "docs: session end closure for 27-10-25 - reconciliation validation complete"
- **Key Files:**
  - `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` (READ THIS FIRST if making architecture decisions)
  - `docs/NOW_TASKS.md` (High-priority item tracking)
  - `openspec/project.md` (Project conventions and tech stack)
  - `AGENTS.md` (OpenSpec workflow instructions)
  - `amp_automation/presentation/postprocess/` (Post-processing modules)
  - `amp_automation/validation/` (Data validation modules)
- **Timezone:** Abu Dhabi/Dubai (UTC+04)
- **Template Path:** `template/Template_V4_FINAL_071025.pptx`
- **Excel Source:** `template/BulkPlanData_2025_10_14.xlsx` (Lumina export)

---

## Workflow Directive

### Plan → Change → Test → Document → Commit Loop

**For any task (new feature, bug fix, investigation):**

1. **Plan** (5-10 min):
   - Use `TodoWrite` to outline discrete steps
   - Identify affected files and modules
   - Flag any architectural decisions (use OpenSpec if breaking change)
   - Verify approach aligns with constraints (COM prohibition, merge allowlist, etc.)

2. **Change** (implement):
   - Write code following project style (`snake_case` functions, type hints, Google docstrings)
   - Use Python (python-pptx) for bulk operations; never COM loops
   - Follow 8-step post-processing workflow order (unmerge → delete-CF → merge-* → normalize)
   - Update validators and test stubs as needed

3. **Test** (validation):
   - Generate test deck: `python -m amp_automation.cli.main ...`
   - Validate structure: `python tools/validate_structure.py ...`
   - Run data validation: `python tools/validate_all_data.py ...`
   - Spot-check fonts, merges, and layout visually if needed
   - Run pytest suite: `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests/...`

4. **Document** (update context):
   - Update relevant docs: `docs/NOW_TASKS.md`, session brain reset, commit message
   - For investigation findings, add to `docs/NOW_TASKS.md` with clear conclusions
   - Link related OpenSpec changes or ADRs if applicable
   - Include file paths in documentation (e.g., `assembly.py:1222`)

5. **Commit** (save):
   - Write detailed message (problem + solution, not just "fix: ...")
   - Use conventional prefixes: `fix:`, `feat:`, `docs:`, `refactor:`
   - For architecture decisions, reference OpenSpec change ID
   - Push immediately unless experimental (then create branch)

### Command Conventions (Windows PowerShell)

```powershell
# Generate deck
python -m amp_automation.cli.main --excel "D:\...\BulkPlanData_2025_10_14.xlsx" --template "D:\...\Template_V4_FINAL_071025.pptx" --output "D:\...\output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx"

# Validate structure
python "D:\...\tools\validate_structure.py" "D:\...\output\...\GeneratedDeck.pptx" --excel "D:\...\BulkPlanData_2025_10_14.xlsx"

# Post-process (Python)
python -m amp_automation.presentation.postprocess.cli "D:\...\GeneratedDeck.pptx" postprocess-all

# Validate all data
python "D:\...\tools\validate_all_data.py" "D:\...\GeneratedDeck.pptx" --excel "D:\...\BulkPlanData_2025_10_14.xlsx"

# Run tests
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1; python -m pytest tests\test_tables.py tests\test_structural_validator.py -v

# Close PowerPoint (if needed)
Stop-Process -Name POWERPNT -Force
```

### Inherited Guardrails

✅ **RESPECT THESE:**
- All 27-10-25 session completions are validated and production-ready
- Do NOT re-implement features already marked [x] in BRAIN_RESET
- Do NOT modify test fixtures, security code, or `.gitignore` without explicit request
- Do NOT commit secrets, large generated files, or logs
- Always run validators after generating new decks
- Always close PowerPoint sessions before running COM automation
- Maintain absolute Windows paths in documentation and config files

✅ **RESPECT THIS CONTEXT:**
- Reconciliation is likely due to Excel data source mapping (not validator code bugs)
- Slide 1 geometry parity is achievable with visual diff + PowerPoint Compare (Zen MCP)
- Test suite rehydration is straightforward (update assertions for current pipeline output)
- Campaign pagination design is Phase 3-4 work; focus on enabling logic, not aggressive optimization

---

## How to Use This Bootstrap

1. **Read in order:** Alignment Check → Brain Reset Digest → Workflow Directive → Guardrails
2. **Before starting work:** Create a TodoList via `TodoWrite` breaking down your task
3. **When stuck:** Consult ARCHITECTURE_DECISION_COM_PROHIBITION.md and openspec/project.md
4. **When making decisions:** Use Zero-Question Rule (propose + choose + document assumption)
5. **When done:** Update session docs, commit with detail, and push

---

**STATUS:** Ready for 28-10-25 session start
**Prepared on:** 2025-10-28 00:00 AST
**Key Contact:** See `AGENTS.md` for workflow instructions
