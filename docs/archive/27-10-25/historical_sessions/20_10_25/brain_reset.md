# Brain Reset Snapshot - 20 Oct 2025 (Alex-approved baseline)

> Read this first after a cold start. It captures the mission, current technical state, risks, and the live backlog so you can resume instantly.

---

## Mission & Success Criteria
- **Goal:** Generate AMP laydown decks that are pixel-identical to `template/Template_V4_FINAL_071025.pptx` while binding Lumina Excel data without corruption.
- **Ship criteria:**
  1. PowerPoint opens/closes generated decks with zero repair prompts.
  2. `tools/validate_structure.py` passes (media ordering, campaign totals, slide grand totals, footnotes, tiles, legend).
  3. Visual diff vs the approved template stays under thresholds, corroborated by Zen MCP plus PowerPoint Review > Compare evidence.
  4. Targeted regression suites cover the clone pipeline, typography, continuation logic, and the AutoPPTX toggle.

**Stakeholder:** Alex (sole approver). All reviews/sign-offs route to him.

---

## Current Snapshot (20 Oct 2025 evening)
- **Clone-only pipeline:** `tooling.autopptx.enabled` remains `false`; clone pipeline is the production path. Latest good deck is `output/presentations/run_20251021_102357/GeneratedDeck_20251021_102357.pptx` (88 slides, clone-only; prior runs used the `_Task11_fixed` legacy name).
- **Styling progress:** Table headers/body fills now match Template V4 theme colours, media-month cells map to the measurement RGB palette, and summary tiles/footer reuse template text frames. Legend tokens are rebuilt when template groups are missing, with colour chips and Verdana 6 pt labels positioned per measurements.
- **Validation state:** `tools/validate_structure.py` passes on the new deck; the regression fixture at `run_20251020_124516` is intentionally left with missing grand totals for negative tests. Pytest panel (`test_autopptx_fallback`, `test_tables`, `test_assembly_split`, `test_structural_validator`) is green.
- **Visual parity:** Slide 1 still fails diff thresholds (mean~195, RMS~213) because we only have a single-slide template export and residual grey/anti-aliasing drift. Multi-slide template imagery is still required before Zen MCP / Review > Compare can be executed.
- **Tooling:** Visual diff CLI (`tools/visual_diff.py`) and Zen MCP server (`temp/zen-mcp-server`) are installed. Manual scripts (`scratch/fill_probe.py`) confirm header and media fill colours post-styling.

---

## Progress Log (key milestones)
1. **Template cloning foundation (earlier in Oct):** Relationship remapping + clone helpers eliminated repair prompts; structural contract codified in `config/structural_contract.json` + validator tests.
2. **Row-limit & splitting update (20 Oct morning):** Raised `MAX_ROWS_PER_SLIDE` to 32, removed forced one-campaign-per-slide split.
3. **Measurement alignment (20 Oct):** Audited `docs/coordinates_measurements/` and aligned geometry constants.
4. **AutoPPTX sunset (20 Oct 17:15):** Configured clone-only default; retained AutoPPTX path as opt-in via config/test.
5. **Styling parity push (20 Oct evening):** Recreated template fills, cloned summary tiles/footer, added legend reconstruction, and introduced `fonts.legend_family` control.

---

## Work Breakdown

### Done
- Clone pipeline stable; AutoPPTX disabled by default with regression coverage for manual toggles.
- Table/legend/summary tiles restyled to align with measurement pack; footer coordinates corrected.
- Structural validator passes on the 88-slide deck; pytest suite green.
- Visual diff tooling functional; exports cached for Slide 1.

### In Progress
- Styling polish (alternating greys, text rendering, footer spacing) to land pixel parity on Slide 1 before scaling to all slides. Verdana body/header fonts now match the template and base columns/totals carry the BACKGROUND_1 grey; awaiting refreshed captures to confirm footer spacing and residual diff noise.
- Capturing a multi-slide template baseline to unblock full-deck visual diff and Zen MCP workflows.

### Pending
- [ ] **Focus 1 – Slide 1 polish:** Finish alternating greys, Verdana rendering, and footer spacing using the measurement pack plus refreshed 125% captures.
- [ ] **Focus 2 – Visual QA evidence:** Capture a multi-slide template baseline, rerun `tools/visual_diff.py`, then execute Zen MCP + PowerPoint Review > Compare and archive the outputs.
- [ ] **Focus 3 – Regression coverage:** Extend pytest coverage for summary-tile typography, legend rebuild behaviour, and the clone-toggle-off path.
- [ ] **Focus 4 – Pipeline smoke:** After visuals pass, run `scripts/run_pipeline_local.py` on additional markets to validate the 32-row contract.
- [ ] **Focus 5 – Baseline alignment:** Align with Alex on the final baseline deck (current clone build is 88 slides) once styling parity is signed off.
- [ ] **Documentation/OpenSpec:** Keep this brain reset, daily log, OpenSpec project, and change files synchronised with progress and Alex's tracker.

---

## Risk Register
- **Visual diff outstanding:** Without a multi-slide template reference, diff metrics cannot prove parity; styling tweaks may still be required.
- **Anti-aliasing variance:** Even with correct fills, font rendering differences could keep diff metrics above threshold; may need to experiment with export DPI or anti-aliasing settings.
- **Regression gaps:** Tests do not yet cover the new summary-tile/legend styling paths; risk of future regressions if we skip coverage.
- **Multi-market assurance:** Updated styling not yet exercised on larger Excel inputs; potential edge cases in splitting/styling across datasets.

---

## Assets & References
- **Decks:**
  - Latest passing clone run: `output/presentations/run_20251021_102357/GeneratedDeck_20251021_102357.pptx`
  - Legacy failure fixture: `output/presentations/run_20251020_124516/GeneratedDeck_Task11_fixed.pptx`
- **Visual diagnostics:** `output/visual_diff/run_20251021_102357_slide1/exports/`
- **Template snapshots:** `output/presentations/run_20251020_164703/{template_201025.png, output_201025.png, Screenshot 2025-10-20 170126.png}`
- **Measurements:** `docs/coordinates_measurements/{complete_measurements_with_colors.json, FINAL_VERIFIED_MEASUREMENTS.txt, coordinate_guide.txt}`
- **Key commands:**
  - Generate: `python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx --template template/Template_V4_FINAL_071025.pptx --output 'GeneratedDeck_{timestamp}.pptx'`
  - Structural validator: `python tools/validate_structure.py <deck> --excel template/BulkPlanData_2025_10_14.xlsx`
  - Visual diff (Slide probe): `python tools/visual_diff.py --reference template/Template_V4_FINAL_071025.pptx --generated <deck> --keep-exports --max-slides 1`
  - Tests: `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests/test_autopptx_fallback.py tests/test_tables.py tests/test_assembly_split.py tests/test_structural_validator.py`
  - Zen MCP server: `python -m zen_mcp_server.server`

Keep this file updated at the end of each session so the next reset starts from a fully informed state.
