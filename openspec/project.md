# Project Context

## Immediate Next Steps (27 Oct 2025)
Last verified on 27-10-25 (end of session)

**COMPLETED (24-27 Oct 2025):**
1. ✅ **Python post-processing migration complete:** Full cell merge logic implemented in `cell_merges.py`. PowerShell COM scripts replaced with Python CLI. Performance: ~30 seconds for 88 slides (vs 10+ hours COM).
2. ✅ **8-step workflow finalized:** unmerge-all → delete-carried-forward → merge-campaign → merge-media → merge-monthly → merge-summary → fix-grand-total-wrap → remove-pound-totals → normalize-fonts. Validated: 100% success rate.
3. ✅ **Timestamp fix (27 Oct):** Local system time (Arabian Standard Time UTC+4) used across all modules (cli.py, logging.py, assembly.py).
4. ✅ **Media channel merging (27 Oct):** Vertical cell merging for TELEVISION, DIGITAL, OOH, OTHER media channels.
5. ✅ **Font corrections (27 Oct):** 6pt body/campaign/bottom rows, 7pt header/BRAND TOTAL.
6. ✅ **Smart line breaking (27 Oct):** _smart_line_break function implemented (dash handling, word-count-based splitting).
7. ✅ **Production deck generated:** `run_20251027_193259` (88 slides) with all formatting improvements.

**CURRENT PRIORITIES:**
1. **Fix campaign cell text wrapping:** PowerPoint overriding explicit line breaks. Solutions: widen column A, disable word-wrap, or conditional font size. See `docs/NOW_TASKS.md`.
2. **Slide 1 EMU/legend parity:** Visual diff to compare generated vs template. Fix any geometry/legend discrepancies.
3. **Test suite rehydration:** Fix/update `tests/test_tables.py`, `tests/test_structural_validator.py`. Add regression tests for merge correctness.
4. **Campaign pagination design:** Design strategy to prevent campaign splits across slides. Create OpenSpec proposal once design is complete.

## Purpose
Automate Annual Marketing Plan laydown decks by converting standardized Lumina Excel exports into pixel-accurate PowerPoint presentations that mirror the `Template_V4_FINAL_071025.pptx` master while preserving financial and media metrics.

## Tech Stack
- Python 3.13.x runtime with `from __future__ import annotations`
- Data processing: pandas, numpy, openpyxl (via pandas)
- Presentation generation: python-pptx plus template-clone helpers
- **Post-processing: python-pptx (bulk operations), PowerShell wrapper (`tools/PostProcessNormalize.ps1`)**
- PowerPoint COM: **ONLY for file I/O and generation-time merges** (NOT bulk post-processing)
- Tooling: pytest 8.x (`PYTEST_DISABLE_PLUGIN_AUTOLOAD=1`), Zen MCP
- **PROHIBITED:** PowerPoint COM for bulk post-processing table operations (see `docs/24-10-25/ARCHITECTURE_DECISION_COM_PROHIBITION.md`)

## Project Conventions

### Code Style
- Typed, snake_case functions; dataclasses (`slots=True`); module loggers under `amp_automation.*`
- Pathlib `Path`, f-strings, config-driven constants; no bare `print`
- Black-compatible formatting and concise inline comments

### Architecture Patterns
- CLI (`amp_automation.cli.main`) orchestrates runs via config
- Data ingestion normalizes Lumina exports using configured column indices
- **Presentation assembly:** Clones template shapes, enforces styling, **creates cell merges during generation** (assembly.py with _smart_line_break for campaign names)
- **Post-processing:** Python-based normalization (`amp_automation/presentation/postprocess/`) via CLI or PowerShell wrapper. Includes media channel vertical merging, font normalization, and smart text formatting.
- Validation: `tools/validate_structure.py`, `tools/visual_diff.py`, Zen MCP + PowerPoint Review > Compare
- Change management via OpenSpec (`openspec/changes/*`)

### Key Architecture Decisions
- **Cell merges created during generation, NOT post-processing** (discovered 24 Oct 2025)
- **COM prohibited for bulk post-processing operations** (60x performance penalty)
- **Python (python-pptx) required for all bulk table operations**
- See: `openspec/changes/clarify-postprocessing-architecture/` and `docs/24-10-25/ARCHITECTURE_DECISION_COM_PROHIBITION.md`

### Testing Strategy
- Targeted pytest suites (`tests/test_tables.py`, `tests/test_assembly_split.py`, `tests/test_autopptx_fallback.py`, `tests/test_structural_validator.py`)
- Lightweight fixtures using python-pptx scratch slides
- Update tests whenever template geometry changes; enforce ~0.5% tolerance and dash placeholders for missing metrics

### Git Workflow
- Trunk-based (main); short-lived branches, rebase before merge
- Use OpenSpec change IDs in commits for significant work
- Keep repo clean (no generated decks/logs) before pushing

## Domain Context
- Lumina Excel: fixed column indices (campaign names 83, funnel stage 95, etc.); ingestion collapses to market/brand/year
- Template V4 geometry: up to 32 body rows per slide (plus carried-forward + slide GRAND TOTAL), summary tiles, footers, legend
- Continuation slides must retain headers, carried totals, and now slide-level GRAND TOTAL rows
- Visual fidelity is business-critical for client-facing AMP decks

## Important Constraints
- Never modify template aesthetics (colors, fonts, positions) outside of clone operations
- Percent stats tolerance ~0.5%; use dash (`-`) for missing metrics
- Operate under Windows path semantics; avoid sensitive info in logs/config
- Fail fast if validator or visual diff thresholds are not met

## External Dependencies
- Template: `template/Template_V4_FINAL_071025.pptx`
- Lumina Excel exports (column mapping per config)
- Python packages: pandas, numpy, python-pptx, openpyxl
- Zen MCP server (`temp/zen-mcp-server`), PowerPoint COM (file I/O only), OpenSpec CLI
