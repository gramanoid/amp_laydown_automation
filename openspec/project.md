# Project Context

## Immediate Next Steps (24 Oct 2025)
Last verified on 24-10-25
1. **Complete Python post-processing migration:** Implement full cell merge logic (`cell_merges.py`) and span reset operations (`span_operations.py`) to replace deprecated PowerShell COM scripts. Target: <10 minutes for 88-slide deck.
2. **Full pipeline testing:** End-to-end test of generation → Python post-processing → validation. Verify visual parity with baseline decks and measure performance against <20 minute target.
3. **PowerShell integration:** Update `PostProcessCampaignMerges.ps1` to call Python CLI, retaining COM only for file I/O operations.
4. **Documentation & enforcement:** Ensure all docs reference COM prohibition (`docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`). Add code review requirements for any COM usage.
5. **Campaign pagination discovery:** Once post-processing is stable, design no-campaign-splitting strategy and capture OpenSpec change proposal.

## Purpose
Automate Annual Marketing Plan laydown decks by converting standardized Lumina Excel exports into pixel-accurate PowerPoint presentations that mirror the `Template_V4_FINAL_071025.pptx` master while preserving financial and media metrics.

## Tech Stack
- Python 3.13.x runtime with `from __future__ import annotations`
- Data processing: pandas, numpy, openpyxl (via pandas)
- Presentation generation: python-pptx plus template-clone helpers
- Post-processing: python-pptx (bulk operations), PowerPoint COM ONLY for file I/O
- Tooling: pytest 8.x (`PYTEST_DISABLE_PLUGIN_AUTOLOAD=1`), Zen MCP
- **PROHIBITED:** PowerPoint COM for bulk table operations (see `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`)

## Project Conventions

### Code Style
- Typed, snake_case functions; dataclasses (`slots=True`); module loggers under `amp_automation.*`
- Pathlib `Path`, f-strings, config-driven constants; no bare `print`
- Black-compatible formatting and concise inline comments

### Architecture Patterns
- CLI (`amp_automation.cli.main`) orchestrates runs via config
- Data ingestion normalizes Lumina exports using configured column indices
- Presentation assembly clones template shapes, enforces styling, drives summary tiles, handles continuation logic
- Validation: `tools/visual_diff.py`, `tools/validate_structure.py`, Zen MCP + PowerPoint Review > Compare once imagery aligns
- Change management via OpenSpec (`openspec/changes/*`)

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
