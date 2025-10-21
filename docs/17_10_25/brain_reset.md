# Brain Reset Snapshot – 20 Oct 2025

Use this document to regain full operational context after a cold start. It captures the project mandate, technical architecture, latest execution state, validations performed, toolchain status (including Zen MCP), and the precise backlog that remains.

## Mission Overview
- **Objective:** Produce AMP campaign decks that are indistinguishable from Template V4 by cloning template shapes and populating data—fully retiring the legacy manual constructors.
- **Primary Success Criteria:**
  1. Generated PPTX opens in PowerPoint with zero repair prompts.
  2. Visual pixel diff returns zero (or sub-threshold) variance across all slides.
  3. Regression suite covers tables, media/funnel tiles, layout hierarchy, and data binding edge cases.

## Architecture & Implementation Conventions
- Geometry constants live in `amp_automation/presentation/template_geometry.py`; treat these as source of truth for placement.
- Cloning API surface (all in `amp_automation/presentation/template_clone.py`):
  - `clone_template_table(template_slide, target_slide, shape_name)`
  - `clone_template_shape(template_slide, target_slide, shape_name)`
  - `_clone_element(element, source_part, target_part)` handles XML deep-copy, appends to `p:spTree`, and remaps relationship IDs (`r:embed`, `r:link`, `r:id`). Never insert before `p:extLst` again.
- Population happens post-clone only; **never** mutate geometry.
- `assembly.py` orchestrates slide build using clone helpers and delegates data formatting to `tables.py` / `template_geometry.py` utilities.
- Validation helpers:
  - `tools/visual_diff.py` exports both template & generated decks via PowerPoint COM, compares PNGs with PIL, and writes metrics + diffs.
  - `tools/inspect_generated_deck.py` re-saves decks to surface latent corruption, outputs shape & rel diagnostics.

## Completed Work (Latest Run: `run_20251020_124516`)
1. Flattened CLI output resolution—`--output` now treated as filename inside run directory (`output/presentations/run_<ts>/`).
2. Eliminated PowerPoint repair prompt by appending cloned shapes to `p:spTree` and remapping relationship IDs during clone.
3. Hardened table border styling (`apply_table_borders`) to emit spec-compliant line definitions (removed unsupported cap/alignment/headEnd attributes). Resulting decks now open via PowerPoint COM without errors.
4. Regenerated deck `GeneratedDeck_Task11_fixed.pptx` (114 slides) under `output/presentations/run_20251020_124516/`; COM automation and python-pptx both confirm slide count = 114.
5. Executed targeted pytest suite: `python -m pytest tests/test_tables.py tests/test_assembly_split.py` → all passing.
6. Ran `tools/visual_diff.py` against the new run. COM export succeeded, producing 114 generated PNGs and 1 template PNG; diff currently fails threshold because template baseline only contains the single master slide.
7. Zen MCP server vendor repo cloned locally at `temp/zen-mcp-server` (also installed via `pip install -e temp/zen-mcp-server`), mirroring the OpenRouter allow-list (`claude-sonnet-4.5`, `grok-4`, `gemini-2.5-pro`, `gpt-5`). Custom FastMCP adapter (`mcp_app.py`) enables `mcp run` CLI invocation.
8. Manual numpy/PIL diff confirms Slide1 PNG parity (max diff = 0). Additional template imagery still required for the remaining 113 slides.
9. Added `features.clone_pipeline_enabled` configuration flag (default `true`) so operators can fall back to the legacy AutoPPTX workflow when required.
10. Legacy AutoPPTX fallback now reuses the detailed clone table builder (year-aware) and is exercised by `tests/test_autopptx_fallback.py` to guard the configuration toggle path.
11. Refactored `tools/visual_diff.py` to accept explicit `--reference`/`--generated` arguments, export caching, slide limits, and tightened thresholds; zero-diff metrics captured for `GeneratedDeck_Task11_fixed.pptx` when using the run itself as the provisional reference (summary: `output/visual_diff/diffs/GeneratedDeck_Task11_fixed_vs_GeneratedDeck_Task11_fixed/diff_summary.json`).
12. Added `config/structural_contract.json` + `tools/validate_structure.py` to codify the structural checklist; pytest `tests/test_structural_validator.py` confirms current deck breaches (media ordering, missing grand totals, footnote date).

### Structural Validation Contract (20 Oct 2025)
The current baseline focuses on structural fidelity (slide geometry, styling, layout). Use the checklist below when reviewing generated decks prior to declaring a golden 114-slide reference:

1. **Slide Frame:** Title bar format `[MARKET] – [BRAND] (YY)` with template background, border, and Verdana typography. No deviations in color or placement.
2. **Column Layout:** Headers fixed as `CAMPAIGN | MEDIA | METRICS | JAN … DEC | TOTAL | GRPs | %`. Column widths may flex to avoid wrapping, but labels and dash placeholders remain identical.
3. **Media Blocks:** Only media with spend are rendered, preserving template ordering (TV → Digital → OOH → Radio → others). Entirely empty rows are removed; empty cells use the template grey dash.
4. **Metric Rows:** Media-specific metric rows strictly follow the template (TV: £000/GRPs/Reach@1+/OTS@3+, Digital: £000 + channel reach rows, other media: spend only). No additional metrics allowed.
5. **Styling:** Apply client RGB palette, Verdana fonts, alternating fills, and text colors exactly. Numeric formatting uses `K/M` suffixes to keep values on a single line where possible.
6. **Subtotals/Totals:** “MONTHLY TOTAL (£ 000)” appears after the final media row of each campaign (even across slide splits). A slide-level “GRAND TOTAL” closes every slide. Quarter tiles are present and compute from on-slide data.
7. **Footer Tiles & Footnote:** TV/DIG/OTHER tiles plus AWA/CON/PUR cards appear on every slide with fixed colors. Footnote required; source line reads `DDMMYY Lumina Export`, pulling the date from the Excel filename.
8. **Dynamic vs Static:** Structural elements above are immutable. Campaign order, row counts, and values are data-driven. Media sections disappear only when entirely empty; columns never do.
9. **Other Elements:** No additional logos, icons, animations, or transitions beyond template defaults.

## In-Progress & Blocked Items
1. **Template Baseline for Visual Diff**
   - `tools/visual_diff.py` now exports the entire generated deck, but the template export still yields a single master-slide PNG. Current workaround uses the generated deck as its own baseline (zero diff) to validate the tooling; still need a genuine 114-slide reference set (curated template or golden run) before thresholds can assert parity.
2. **Structural Validator Follow-ups**
   - `tools/validate_structure.py` flags structural gaps in `GeneratedDeck_Task11_fixed.pptx`: campaigns render Digital blocks before Television, campaign tables omit per-slide `GRAND TOTAL`, and the footnote still carries placeholder dates. These must be resolved before declaring golden parity.

2. **Zen MCP Visual Analysis Workflow**
   - Server reachable via `mcp run mcp_app.py:app` (FastMCP). Inspector command stalls due to hosting requirements; revisit once slide images are available.
   - Next: supply paired template/generated PNGs to Zen `chat`/`thinkdeep` for LLM-based comparison once matching template imagery exists.

3. **Manual QA Sign-off**
   - Must run PowerPoint Review → Compare between template and generated deck; capture screenshots, annotate differences (expected: none).

4. **Regression Expansion**
   - Add tests covering media/funnel tile typography, legend alignment, and prevention of duplicated cloned shapes across slides.
   - Extend integration coverage to ensure pipeline splits maintain hierarchical ordering when multiple campaigns exist.

5. **Pipeline Hierarchy Validation**
   - Run `scripts/run_pipeline_local.py` on representative Excel inputs (multi-market, multi-campaign) to ensure stage ordering and slide counts remain correct.

## Operational Commands & Paths
- Generate deck:
  ```bash
  python -m amp_automation.cli.main \
    --excel template/BulkPlanData_2025_10_14.xlsx \
    --template template/Template_V4_FINAL_071025.pptx \
    --output GeneratedDeck_Task11_fixed.pptx
  ```
- Visual diff baseline (once exports succeed):
  ```bash
  python tools/visual_diff.py \
    --generated output/presentations/run_20251020_124516/GeneratedDeck_Task11_fixed.pptx \
    --reference template/Template_V4_FINAL_071025.pptx
  ```
- Manual diff helper (current workaround): `output/visual_diff/manual_diff/`
- Tests: `pytest tests/test_tables.py tests/test_assembly_split.py tests/test_autopptx_fallback.py`
- Visual diff: `python tools/visual_diff.py --reference <baseline.pptx> --generated <new_run.pptx> [--keep-exports --max-slides N]`
- Zen MCP launch:
  ```powershell
  & "python" -m zen_mcp_server.server
  ```

## Risks & Considerations
- PowerPoint automation is unstable in headless PowerShell—may require running from Windows desktop session or using `comtypes` with elevated privileges.
- Template PPTX currently exports only Slide1 PNG because file contains one master slide; ensure template export path is correct before diffing.
- Keep relationship remapping logic under test; any missing `r:id` update can reintroduce repair prompts.
- Do not alter template geometry constants without revalidating every slide.

## Next Session Checklist
1. Resolve COM export blocker; once PNGs exist for full deck, rerun visual diff + capture metrics JSON.
2. Use Zen MCP `chat` or `thinkdeep` with exported images to log LLM validation.
3. Perform PowerPoint Review → Compare and note results in work log.
4. Implement regression tests (tiles + duplication).
5. Run `scripts/run_pipeline_local.py` smoke tests; log timings and output structure.
6. Update documentation and OpenSpec once above steps complete.
