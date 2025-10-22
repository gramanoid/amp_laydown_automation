# Project Context

# Project Context

## Immediate Next Steps (21 Oct 2025)
1. **Slide 1 parity closure:** Reset table row heights/column widths to template EMUs, force center alignment, and drop the synthesized legend shapes so Slide 1 geometrically matches the master.
2. **Baseline imagery & visual diff:** Export multi-slide PNGs for the template and regenerated deck, rerun full-deck `tools/visual_diff.py`, and execute Zen MCP + PowerPoint Review > Compare with archived evidence.
3. **Regression & smoke:** Add pytest coverage for summary-tile typography, legend rebuild behaviour, and clone-toggle-off path, then run `scripts/run_pipeline_local.py` on additional markets to validate the 32-row contract once visuals are signed off.
4. **Campaign pagination discovery:** Once current blockers clear, run a focused Q&A to design a no-campaign-splitting strategy so each campaign stays on a single slide and capture an OpenSpec change proposal for delivery.

## Purpose
Automate Annual Marketing Plan laydown decks by converting standardized Lumina Excel exports into pixel-accurate PowerPoint presentations that mirror the `Template_V4_FINAL_071025.pptx` master while preserving financial and media metrics.

## Tech Stack
- Python 3.13.x runtime with `from __future__ import annotations`
- Data processing: pandas, numpy, openpyxl (via pandas)
- Presentation generation: python-pptx plus template-clone helpers
- Tooling: pytest 8.x (`PYTEST_DISABLE_PLUGIN_AUTOLOAD=1`), PowerPoint COM automation, Zen MCP

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
- Zen MCP server (`temp/zen-mcp-server`), PowerPoint COM, OpenSpec CLI
