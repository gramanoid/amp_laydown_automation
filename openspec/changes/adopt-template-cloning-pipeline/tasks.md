## 1. Clone-Based Rendering Pipeline
- [x] 1.1 Analyze master slide structure (table/tiles/legends) and document shape IDs required for cloning.
- [x] 1.2 Implement cloning helpers to duplicate target shapes onto new slides while preserving geometry and styling.
- [x] 1.3 Replace manual table construction with data population into cloned table cells/shapes.
- [x] 1.4 Wire configuration toggle to enable/disable clone-based workflow and phase out AutoPPTX adapter once parity confirmed.
  - Implemented via `features.clone_pipeline_enabled` (default `true`) in `master_config.json`, wiring the CLI path to fall back on AutoPPTX when disabled.
  - Regression coverage: `tests/test_autopptx_fallback.py` exercises the AutoPPTX path to prevent year-filter regressions.

## 2. Verification & Regression Safety Nets
- [x] 2.1 Extend visual diff runner to compare multiple slides per deck and surface geometry mismatches.
  - Implemented COM-based export + PIL metrics; needs follow-up to resolve PowerPoint export failures on regenerated deck.
- [x] 2.2 Add unit/integration tests covering clone pipeline, including fixture slides and structural assertions (cell positions, fills, fonts).
  - Current coverage: `tests/test_tables.py`, `tests/test_assembly_split.py` pass; add tile-format tests (follow-up tracked outside spec).
- [x] 2.3 Update documentation/logging to reflect new workflow and capture validation steps for operators.
  - Structural contract captured in docs/17_10_25 + docs/20_10_25 and enforced via `config/structural_contract.json` + `tools/validate_structure.py`; structural issues identified (media ordering, slide-level `GRAND TOTAL`, footnote date) remain open follow-ups.

## 3. Output & Packaging
- [x] 3.1 Flatten run output structure to avoid nested `.../run_<ts>/output/<file>` when `--output` includes a path.
- [x] 3.2 Diagnose and eliminate PowerPoint "Repair" prompt (inspect low-shape slides; adjust XML insertion point if needed).

## 4. Visual Parity Closure (in flight)
- [x] 4.1 Reset cloned table background/alternating fills to template greys; ensure highlight bars only color intended cells (implemented in `tables.py::style_table_cell`).
- [x] 4.2 Realign footer text box and line spacing using measurement bundle so it clears GRAND TOTAL (footer now cloned with `_apply_configured_position`).
- [x] 4.3 Force legend chip RGB values and summary tile typography to match template (legend rebuild + `fonts.legend_family` control landed).
- [ ] 4.4 Export template/generated decks, rerun `tools/visual_diff.py`, perform Zen MCP + PowerPoint Review > Compare, and archive findings (blocked on multi-slide template imagery).
