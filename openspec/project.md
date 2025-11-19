# Project Context

## Immediate Next Steps (29 Oct 2025)
Last verified on 29-10-25 (session complete - accuracy validation)

**COMPLETED (29 Oct 2025):**
1. ✅ **Automated Excel transformations:** Expert campaign exclusion (1,158 rows), geography normalization (11 rules), Panadol brand splitting (Pain/C&F)
2. ✅ **Media split percentages:** Color-coded display in MONTHLY TOTAL rows (TV 38% • DIG 34% • OOH 25% • OTH 3%)
3. ✅ **Comprehensive accuracy validation:** Horizontal/vertical totals validation with K/M parsing, 1% tolerance
4. ✅ **GRP metrics data consistency fix (CRITICAL):** Aligned transformation cache with processed data (Panadol + geography)
5. ✅ **M-suffix formatting precision:** Changed to 2 decimals for millions (£2.84M vs £3M)
6. ✅ **Title formatting fixes:** Black color, left-aligned, 8" width, aligned with table margin
7. ✅ **Production deck validated:** `run_20251029_132708` (120 slides) - Zero large errors (>£5K)

**COMPLETED (24-28 Oct 2025):**
1. ✅ Python post-processing migration, 8-step workflow, timestamp fixes
2. ✅ Media channel merging, font corrections, smart line breaking
3. ✅ Campaign text wrapping resolution, structural validator enhancements
4. ✅ Data validation suite expansion (1,200+ lines across 5 modules)
5. ✅ Test suite rehydration (16 regression tests: 8 formatting, 3 structural, 5 footer)

**DEFERRED (Phase 4+):**
1. Slide 1 EMU/legend parity (P2)
2. Visual diff workflow enhancement (P3)
3. Python normalization expansion (P3)
4. Automated regression scripts (P3)
5. Reconciliation data source investigation (P3 - all tiles showing "expected data missing")

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

---

Last verified on 29-10-25
