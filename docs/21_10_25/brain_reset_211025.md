# Brain Reset – 21 Oct 2025  
_Use this after a cold start to regain full situational awareness_

---
## Mission Snapshot
- **Goal:** Produce clone-generated AMP laydown decks that are pixel-identical to `template/Template_V4_FINAL_071025.pptx` while binding Lumina data without triggering PowerPoint repair prompts.
- **Environment:** Windows (PowerShell shell), Python 3.13, `python-pptx` for clone pipeline, COM exports for visual diff. AutoPPTX path kept only for negative tests (`features.clone_pipeline_enabled = true`).
- **Current Baseline Deck:** `output/presentations/run_20251021_102357/GeneratedDeck_20251021_102357.pptx` generated with `python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx --template template/Template_V4_FINAL_071025.pptx --output "GeneratedDeck_{timestamp}.pptx"`.
- **Latest Commit:** `7f7326b` (“chore: snapshot before slide1 fix plan”) is pushed to `origin/main`.

---
## Technical State (21 Oct)
### Clone Pipeline
- Slide geometry, fonts, fills, summary tiles, footer positioning now clone directly from the template (`amp_automation/presentation/assembly.py`). Verdana fonts enforced via the new `presentation.fonts.table_family` config entry.
- Legend reconstruction logic synthesizes chips/text when template groups are missing; Slide 1 still retains the synthesized legend (to be pruned for parity).
- Structural contract codified in `config/structural_contract.json`; validator script `tools/validate_structure.py` passes on the latest deck, with `run_20251020_124516` acting as the failing fixture (grand total row blanked intentionally). Targeted pytest suite (`tests/test_autopptx_fallback.py`, `test_tables.py`, `test_assembly_split.py`, `test_structural_validator.py`) is green.

### Visual QA
- Single-slide diff (`tools/visual_diff.py --max-slides 1`) compares the latest deck against the template and still reports high variance (mean 195.45 / RMS 213.45) because alignment/row-height discrepancies remain.
- Multi-slide template baseline not yet captured; Zen MCP + PowerPoint Review > Compare evidence pending.
- Fresh forensic bundle from the external analyst (now using the correct decks) lives under `docs/21_10_25/differential_visual_analysis/` (README, Quick Reference, LLM Fix Instructions, Complete Forensic Analysis, Technical Methodology, plus `local_diff_report.json` and `slide1_211025.pptx`).

### Key Measurements (from template)
- Slide dimensions: 9 144 000 EMU × 5 143 500 EMU (10.0″ × 5.625″).
- Table position: left 163 582 EMU, top 638 117 EMU, width 8 531 095 EMU, height 3 766 424 EMU.
- Column widths (EMU): `[812 364, 729 251, 831 384, 338 274, 400 567, 400 567, 400 567, 414 770, 415 954, 465 506, 437 595, 443 865, 400 567, 437 595, 352 043, 449 092, 400 567, 400 567]`.
- Row heights: row 0 = 161 729 EMU (header), rows 1‑33 = 99 205 EMU (body), row 34 = 0.
- Alignment: template centers every table cell.

---
## Delta Summary (Template vs Latest Deck)
1. **Row Heights:** Header currently 127 101 EMU, body rows 107 899 EMU → needs reset to 161 729 / 99 205 EMU.
2. **Text Alignment:** Generated deck still Left/Right; required alignment is `PP_ALIGN.CENTER` across all cells.
3. **Legend Shapes:** Template slide 1 has 12 shapes; clone slide keeps 20 (extra legend shapes `TelevisionLegend*`, `DigitalLegend*`, etc.). They must be removed on slides where the template omits the legend.
4. **Column Widths:** Off by ≤ 252 EMU in total; snap to template width list to eliminate drift.
5. **Visual Evidence:** Need multi-slide template PNG baseline + Zen MCP / PowerPoint Compare run (currently blocked).

---
## Active Focus Items (Alex)
1. **Slide 1 polish:** Correct alternating greys, enforce Verdana rendering, and fix footer spacing (row height + alignment changes are part of this).
2. **Visual QA evidence:** Capture multi-slide template baseline, rerun `tools/visual_diff.py`, then execute Zen MCP + PowerPoint Review > Compare and archive outputs.
3. **Regression coverage:** Extend pytest for summary-tile typography, legend rebuild behaviour, and clone-toggle-off path.
4. **Pipeline smoke:** Run `scripts/run_pipeline_local.py` on additional markets (post-visual sign-off) to validate the 32-row contract.
5. **Baseline alignment:** Confirm final slide count (current clone deck = 88 slides) with Alex after styling parity is achieved.

---
## Immediate Implementation Plan
1. **Table geometry alignment:**
   - Update `tables.py::style_table_cell` to apply template row heights and enforce `PP_ALIGN.CENTER` for every paragraph.
   - Snap column widths to the template’s EMU list after table creation.
2. **Legend pruning:**
   - In assembly, only clone legend groups when the template slide includes them; delete shapes 12‑19 on slides that should not show the legend (Slide 1).
3. **Verification loop:**
   - Regenerate deck (`python -m amp_automation.cli.main --excel ... --output "GeneratedDeck_{timestamp}.pptx"`).  
   - Run `tools/visual_diff.py --max-slides 1`, `tools/validate_structure.py`, and the pytest bundle.  
   - Archive updated diff outputs under a new run directory and log results in `docs/21_10_25/21_10_25.md`.
4. **Baseline capture:** Export multi-slide PNGs for the template and generated deck (PowerPoint COM export), then schedule Zen MCP + PowerPoint Review > Compare run with evidence stored under `docs/21_10_25/`.

---
## Project Management Overview
- **Backlog:** All items tracked in this brain reset plus OpenSpec project overview. Outstanding tasks remain in the “Visual QA evidence”, “Regression coverage”, and “Pipeline smoke” buckets.
- **Documentation:** Daily logs exist per date (`docs/20_10_25/20_10_25.md`, `docs/21_10_25/21_10_25.md`). Differential assets are centralized under `docs/21_10_25/differential_visual_analysis/`. Measurement pack is under `docs/coordinates_measurements/`.
- **Testing:** Keep `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests/test_autopptx_fallback.py tests/test_tables.py tests/test_assembly_split.py tests/test_structural_validator.py` as the quick verification suite. Additional tests will be added for typography/legend/clone toggle once styling fixes land.
- **Source Control:** Work pushed to `origin/main`. Use new commits for Slide 1 fixes; ensure generated decks/visual outputs stay out of git.

---
## Quick Commands Reference
```bash
# Generate deck
python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx \
  --template template/Template_V4_FINAL_071025.pptx \
  --output "GeneratedDeck_{timestamp}.pptx"

# Structural validation
python tools/validate_structure.py output/presentations/<run>/GeneratedDeck_<ts>.pptx \
  --excel template/BulkPlanData_2025_10_14.xlsx

# Visual diff (Slide 1 probe)
python tools/visual_diff.py \
  --reference template/Template_V4_FINAL_071025.pptx \
  --generated output/presentations/<run>/GeneratedDeck_<ts>.pptx \
  --keep-exports --max-slides 1 \
  --output-dir output/visual_diff/<run>_slide1

# Focused pytest suite
PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest \
  tests/test_autopptx_fallback.py tests/test_tables.py \
  tests/test_assembly_split.py tests/test_structural_validator.py
```

---
## Assets & Locations
- Latest deck: `output/presentations/run_20251021_102357/GeneratedDeck_20251021_102357.pptx`
- Diff exports: `output/visual_diff/run_20251021_102357_slide1/`
- Forensic bundle: `docs/21_10_25/differential_visual_analysis/`
- Measurement pack: `docs/coordinates_measurements/`
- OpenSpec project overview: `openspec/project.md`
- Current daily log: `docs/21_10_25/21_10_25.md`

---
## Next Session Kick-off Checklist
1. Re-read this brain reset and `docs/21_10_25/21_10_25.md` to refresh context.
2. Implement row-height + alignment + column-width + legend pruning fixes.
3. Regenerate deck and rerun validation/diff/tests.
4. Capture multi-slide template baseline and execute Zen MCP + Compare.
5. Update logs + OpenSpec with outcomes and adjust remaining backlog items.

Once these are complete we can proceed to regression coverage and the multi-market smoke runs before final sign-off with Alex.
