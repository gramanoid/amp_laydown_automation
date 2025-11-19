# Fresh Start Bootstrap Prompt - 27 Oct 2025

Generated: 2025-10-27
Source documents:
- `D:\Drive\projects\work\AMP Laydowns Automation\docs\24-10-25\BRAIN_RESET_241025.md` (Last Updated: 2025-10-24)
- `D:\Drive\projects\work\AMP Laydowns Automation\docs\24-10-25\24-10-25.md` (Last Updated: 2025-10-24)
- `D:\Drive\projects\work\AMP Laydowns Automation\AGENTS.md` (Last Updated: 2025-10-21)
- `D:\Drive\projects\work\AMP Laydowns Automation\openspec\project.md` (Last Updated: 2025-10-24)
- `D:\Drive\projects\work\AMP Laydowns Automation\openspec\AGENTS.md` (Last Updated: 2025-10-21)

---

## Assumptions & Missing Information

**ASSUME:**
- Production deck from 24 Oct 2025 (`run_20251024_200957`) is available and validated
- Python post-processing pipeline is stable and production-ready
- Template file `Template_V4_FINAL_071025.pptx` remains unchanged
- Excel data source `BulkPlanData_2025_10_14.xlsx` is current baseline

**MISSING:**
- Visual diff baseline not yet established for Slide 1 geometry verification
- Test suite status unclear (needs assessment of which tests are broken/outdated)
- Campaign pagination requirements not fully specified (need Q&A discovery)
- Zen MCP evidence capture workflow not documented

---

## Zero-Question Rule

When uncertain about implementation choices, you MUST:
1. Propose TWO concrete options with pros/cons
2. Choose the option that best aligns with project conventions (see below)
3. Document your assumption in session notes with `ASSUME:` prefix
4. Proceed with implementation

DO NOT ask the user questions during execution unless:
- The choice has irreversible consequences (data loss, security implications)
- Multiple approaches are equally valid and user preference is critical
- Clarification is needed to avoid rework of substantial scope

---

## Alignment Check

### Project Overview
AMP Laydowns Automation: Clone-based generation of pixel-accurate PowerPoint presentations from Lumina Excel exports, mirroring `Template_V4_FINAL_071025.pptx` while preserving financial and media metrics.

### Project Summary
- **Mission:** Automate AMP laydown deck generation using Python (python-pptx) for template cloning and table assembly
- **Current State:** Post-processing workflow complete (8-step Python pipeline: unmerge → delete-carried-forward → merge operations → format fixes → font normalization), production deck validated
- **Tech Stack:** Python 3.13, python-pptx for bulk operations, PowerShell wrappers for integration, PowerPoint COM ONLY for file I/O (not bulk table ops)

### NOW Tasks (with acceptance criteria)
1. **Slide 1 EMU/legend parity work**
   - Acceptance: Visual diff shows <0.5% geometry deviation from template
   - Files: `amp_automation/presentation/assembly.py` (likely), template comparison baseline
   - Evidence: Zen MCP Compare screenshots or visual diff artifacts

2. **Test suite rehydration**
   - Acceptance: `tests/test_tables.py`, `tests/test_structural_validator.py` pass without modifications to core logic
   - Files: `tests/test_tables.py:*`, `tests/test_structural_validator.py:*`, fixture files
   - Evidence: pytest output showing PASSED status, coverage report if available

3. **Add regression tests for merge correctness**
   - Acceptance: Tests verify campaign vertical merges, monthly/summary horizontal merges, no rogue merges
   - Files: New test file `tests/test_post_processing.py` or additions to `tests/test_tables.py`
   - Evidence: Test coverage for all 8 post-processing operations

### Biggest Risk + Mitigation
**Risk:** Visual parity regressions go undetected without automated visual diff workflow, causing client-facing deck quality issues.

**Mitigation:**
1. Establish visual diff baseline using `run_20251024_200957` vs template
2. Document Zen MCP evidence capture process in `docs/27-10-25/`
3. Add visual diff validation to runbook (`docs/27-10-25/BRAIN_RESET_271025.md`)
4. Create regression test for critical geometry values (row heights, column widths, EMU coordinates)

---

## Brain Reset Digest

### Session Overview
Continuing AMP Laydowns Automation project with focus on quality assurance and test coverage. Previous session (24 Oct 2025) completed major milestone: Python-based post-processing pipeline replacing PowerShell COM automation (60x performance improvement). Current session priorities: Slide 1 geometry verification, test suite restoration, regression test coverage.

### Work Completed (Previous Session - 24 Oct 2025)
- ✅ Implemented complete Python post-processing pipeline (`amp_automation/presentation/postprocess/`)
- ✅ 8-step workflow finalized: unmerge-all → delete-carried-forward → merge-campaign → merge-monthly → merge-summary → fix-grand-total-wrap → remove-pound-totals → normalize-fonts
- ✅ Production deck generated and validated: `run_20251024_200957` (88 slides, 556KB, 100% success rate)
- ✅ PowerShell integration completed: `tools/PostProcessNormalize.ps1` wrapper
- ✅ Architecture documented: COM prohibition ADR (`docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`)
- ✅ End-to-end pipeline validated: generation → Python normalization → structural validation (all PASSED)
- ✅ OpenSpec proposal created: `openspec/changes/clarify-postprocessing-architecture/`

**Key Discovery:** Clone pipeline (assembly.py:629,649) creates cell merges during generation, NOT post-processing. Post-processing focuses on normalization and edge case fixes.

### Current State (Session Start - 27 Oct 2025)
- **Position:** Fresh session, clean slate, production-ready baseline available
- **Branch:** main (based on git status)
- **Latest Commit:** 6c6f65e - "docs: finalize 8-step post-processing workflow and end-of-session docs"
- **Dirty Files:** 1 modified (`docs/24-10-25/end_day/summary.md`), multiple untracked files (diagnostic tools, session artifacts)
- **Blockers:** None at session start

### Purpose
Automate Annual Marketing Plan laydown decks by converting standardized Lumina Excel exports into pixel-accurate PowerPoint presentations that mirror the `Template_V4_FINAL_071025.pptx` master while preserving financial and media metrics.

**Critical Constraint:** Visual fidelity is business-critical for client-facing AMP decks. Template aesthetics (colors, fonts, positions, geometry) must remain unchanged.

### Next Steps (NOW tasks - unchecked from BRAIN_RESET)
1. **Slide 1 EMU/legend parity work**
   - Set up visual diff baseline for Slide 1 geometry comparison
   - Compare generated deck Slide 1 vs template Slide 1
   - Document discrepancies and fix in `amp_automation/presentation/assembly.py` or related modules
   - Capture evidence via Zen MCP Compare or visual diff artifacts

2. **Test suite rehydration**
   - Assess current pytest suite status: `tests/test_tables.py`, `tests/test_structural_validator.py`, `tests/test_assembly_split.py`, `tests/test_autopptx_fallback.py`
   - Fix broken tests without modifying core logic (update fixtures, paths, expected values)
   - Document test rehydration priorities in session notes

3. **Add regression tests**
   - Test merge correctness: campaign vertical merges, monthly/summary horizontal merges
   - Test font normalization: Verdana 6-7pt coverage, GRAND TOTAL special formatting
   - Test row formatting: MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD rows
   - Create `tests/test_post_processing.py` or extend existing test files

### Important Notes
- **COM Prohibition:** PowerPoint COM automation for bulk table operations is PROHIBITED (10+ hours vs Python's 10 minutes - 60x difference). COM only for file I/O, exports, features not in python-pptx.
- **Merge Architecture:** Cell merges created during generation (`amp_automation/presentation/assembly.py:629,649`), not post-processing. Post-processing handles normalization and edge cases.
- **Horizontal Merge Allowlist:** MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD only. Any other merged labels are regressions.
- **Font Standards:** Verdana 6pt (GRAND TOTAL with zero margins), Verdana 7pt (other bottom rows), Verdana 6pt with dashes in empty cells to prevent Calibri 18pt reversion.
- **Template Fidelity:** EMUs, centered alignment, font sizes must remain faithful to `Template_V4_FINAL_071025.pptx`. Adjust scripts cautiously.

### Session Metadata
- **Timezone:** Abu Dhabi/Dubai (UTC+04)
- **Today's Date:** 27-10-25 (DD-MM-YY)
- **Session Started:** 2025-10-27
- **Project Root:** `D:\Drive\projects\work\AMP Laydowns Automation`
- **Latest Baseline Deck:** `output/presentations/run_20251024_200957/AMP_Presentation_20251024_200957.pptx` (88 slides, 556KB)
- **Latest Commit:** 6c6f65e (24 Oct 2025)
- **Key Modules:**
  - CLI: `amp_automation/cli/main.py`
  - Assembly: `amp_automation/presentation/assembly.py` (creates merges during generation)
  - Post-processing: `amp_automation/presentation/postprocess/` (cli.py, table_normalizer.py, cell_merges.py, unmerge_operations.py)
  - Validation: `tools/validate_structure.py`
  - PowerShell wrapper: `tools/PostProcessNormalize.ps1`

### Outstanding Checklist (Carry Forward from BRAIN_RESET)
- [ ] Set up visual diff baseline for Slide 1 geometry comparison
- [ ] Rehydrate pytest test suites with current pipeline state
- [ ] Design campaign pagination to prevent across-slide splits
- [ ] Document Zen MCP evidence capture workflow

---

## Workflow Directive

Follow the **Plan → Change → Test → Document → Commit** loop:

### Plan
- Review NOW tasks in BRAIN_RESET and prioritize based on user guidance
- Break down tasks into small, testable increments
- Identify files requiring changes (use absolute paths: `D:\Drive\projects\work\AMP Laydowns Automation\...`)
- Check for OpenSpec change proposals if task involves new features or architecture changes

### Change
- Use typed Python with dataclasses (`slots=True`), snake_case, module loggers
- Maintain Black-compatible formatting
- Preserve existing conventions: pathlib `Path`, f-strings, config-driven constants
- **NEVER modify template aesthetics** (colors, fonts, positions) outside clone operations
- **RESPECT COM PROHIBITION:** Use python-pptx for all bulk table operations
- **NO SECRETS** in logs or code (absolute paths OK, credentials/API keys NOT OK)

### Test
- Run structural validation: `python tools\validate_structure.py <deck_path> --excel template\BulkPlanData_2025_10_14.xlsx`
- Execute pytest suites (once rehydrated): `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\`
- Verify post-processing: `python -m amp_automation.presentation.postprocess.cli <deck> postprocess-all` (expect 100% success, 0 failures)
- Visual diff for geometry changes (manual or automated)
- Close PowerPoint sessions before COM operations: `Stop-Process -Name POWERPNT -Force`

### Document
- Update session notes in `docs\27-10-25\27-10-25.md` with progress checkpoints
- Add artifacts to `docs\27-10-25\artifacts\` (logs, analysis, evidence)
- Update BRAIN_RESET TODOs: check off completed items, add new discoveries
- Maintain absolute Windows paths in documentation
- Document assumptions with `ASSUME:` prefix

### Commit
- Use conventional commit messages: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`
- Reference OpenSpec change IDs for significant work (e.g., `feat: implement campaign pagination (adopt-template-cloning-pipeline)`)
- Keep commits atomic and focused
- Run `git status` before committing to verify clean working tree
- **DO NOT commit secrets, generated decks (>500KB), or verbose logs (>1MB)**

### Guardrails (Inherited from Previous Sessions)
- ✅ Use absolute paths when referencing files in documentation or logs
- ✅ Run tests/linters before committing (once test suite is rehydrated)
- ✅ Preserve template geometry and styling (no aesthetic changes)
- ✅ Honor horizontal merge allowlist (MONTHLY TOTAL, GRAND TOTAL, CARRIED FORWARD only)
- ✅ Use Python (python-pptx) for bulk operations, not PowerPoint COM
- ✅ Close PowerPoint sessions before COM automation
- ✅ Fail fast if validator or visual diff thresholds not met
- ✅ Keep repo clean before pushing (no untracked diagnostic files >100 lines)

---

## Quick Reference

### Key Commands
```powershell
# Generate deck
python -m amp_automation.cli.main --excel template\BulkPlanData_2025_10_14.xlsx --template template\Template_V4_FINAL_071025.pptx --output output\presentations\run_YYYYMMDD_HHMMSS\GeneratedDeck_TIMESTAMP.pptx

# Validate structure
python tools\validate_structure.py output\presentations\run_YYYYMMDD_HHMMSS\deck.pptx --excel template\BulkPlanData_2025_10_14.xlsx

# Post-process (Python CLI)
python -m amp_automation.presentation.postprocess.cli output\presentations\run_YYYYMMDD_HHMMSS\deck.pptx postprocess-all

# Post-process (PowerShell wrapper)
.\tools\PostProcessNormalize.ps1 -PresentationPath output\presentations\run_YYYYMMDD_HHMMSS\deck.pptx

# Run tests (once rehydrated)
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD="1"; python -m pytest tests\

# Close PowerPoint before COM operations
Stop-Process -Name POWERPNT -Force
```

### Key Files & Locations
- **Template:** `template\Template_V4_FINAL_071025.pptx`
- **Excel Data:** `template\BulkPlanData_2025_10_14.xlsx`
- **Latest Deck:** `output\presentations\run_20251024_200957\AMP_Presentation_20251024_200957.pptx`
- **CLI Entry:** `amp_automation\cli\main.py`
- **Assembly Logic:** `amp_automation\presentation\assembly.py` (creates merges at lines 629, 649)
- **Post-processing:** `amp_automation\presentation\postprocess\cli.py` (8-step workflow)
- **Validation:** `tools\validate_structure.py`
- **ADR:** `docs\ARCHITECTURE_DECISION_COM_PROHIBITION.md`
- **OpenSpec:** `openspec\project.md`, `openspec\AGENTS.md`
- **Tests:** `tests\test_tables.py`, `tests\test_structural_validator.py`

### Project Conventions (from openspec/project.md)
- Python 3.13.x runtime with `from __future__ import annotations`
- Typed, snake_case functions; dataclasses with `slots=True`
- Module loggers under `amp_automation.*` (no bare `print`)
- Pathlib `Path`, f-strings, config-driven constants
- Black-compatible formatting, concise inline comments
- CLI orchestrates runs via config
- Trunk-based git workflow (main), short-lived branches, rebase before merge

---

**Ready to proceed. Priorities: Slide 1 geometry verification → Test suite rehydration → Regression test coverage.**
