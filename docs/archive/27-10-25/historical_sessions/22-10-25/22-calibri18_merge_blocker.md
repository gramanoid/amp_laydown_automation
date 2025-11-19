### Calibri 18 / Row-Height / Merge Blocker – 22 Oct 2025 16:52 (UTC+04)

**Scope:** `output\presentations\run_20251022_160155\GeneratedDeck_20251022_160155.pptx` (post-process target) and regenerated fallback `output\presentations\run_20251022_164818\GeneratedDeck_20251022_164818.pptx`.

---

#### Current Symptoms
- PowerPoint still renders many ostensibly blank cells as **Calibri 18 pt**, inflating row heights to 14.4–21.6 pt even though the generator emits Verdana 6 pt dashes/ZWSPs.
- `tools/PostProcessCampaignMerges.ps1` cannot complete: `Cell.Merge : Invalid request. Cannot merge cells of different sizes.` surfaces for campaign/monthly ranges; the COM session hangs until PowerPoint is closed manually.
- Failed post-process runs leave the deck with large portions of the Campaign column cleared out (blank grey bands captured in user screenshot).
- Row-height probe (`tools/ProbeRowHeights.ps1`) against the unprocessed deck reports **563/1372 rows** outside the 8.3–8.5 pt tolerance (headers at 12.735 pt; data rows up to 21.6 pt).

---

#### Timeline & Attempts (22 Oct 2025)
1. **Generator-side enforcement (succeeded but insufficient):**
   - Added Verdana 6 pt / zero-width-space logic in `amp_automation/presentation/tables.py`.
   - Python audit (`docs/22-10-25/artifacts/dash_font_check_20251022_1608.txt`) confirmed 24 696 blank/dash cells styled correctly in the baked PPTX.
2. **COM post-process retry (failed):**
   - Updated `tools/PostProcessCampaignMerges.ps1` to normalize blanks, but the first run timed out at 600 s with slide 2 merge failure; PowerPoint left open until the user killed it.
   - Partial execution erased campaign labels (see screenshot).
3. **Row-height probe (post-failure):**
   - Output `docs/22-10-25/artifacts/row_height_probe_20251022_1617.csv` shows persistent outliers despite generator fix.
4. **Emergency revert for user QA:**
   - Regenerated clean baseline (`run_20251022_164818/GeneratedDeck_20251022_164818.pptx`) so manual review can continue while automation is repaired (`docs/22-10-25/logs/20-revert_regeneration_plan_20251022_1650.md` + `21-deck_regeneration_20251022_1648.md`).

---

#### Findings & Hypotheses
- **Blanks revert to Calibri 18 during COM merges:** PowerPoint re-initialises text frames when cells are split/merged; without an immediate post-merge restyle, the default table theme (Calibri 18 pt) returns, inflating row height.
- **`Cannot merge cells of different sizes`** occurs because row heights diverge before the merge loop executes. Once a merge attempt fails, subsequent splits propagate blank formatting, wiping campaign text.
- The current script enforces Verdana before merges, but PowerPoint’s merge operations discard the fonts; the final `Apply-BlankCellFormatting` pass never runs when the script aborts mid-slide.

---

#### Impact
- We cannot deliver post-processed decks: campaign/monthly merges fail, row heights exceed tolerance, and manual intervention is required to repair corrupted slides.
- Any attempt to re-run the script risks blanking large sections again, so the regenerated deck must stay untouched until the automation is fixed.

---

#### Immediate Next Steps
1. **Add a guarded backup/restore flow** inside `PostProcessCampaignMerges.ps1` (duplicate presentation before mutations; rollback on failure).
2. **Make merges idempotent with deterministic sizing:**
   - Pre-set row heights *after every split* and *before* merging.
   - After each merge, call a new helper to restyle blank/dash cells within the merged range (fonts + ZWSP).
3. **Short-circuit on merge exceptions** and log the failing slide without clearing text; capture diagnostics rather than continuing blindly.
4. **Re-run row-height probe** only after the script completes successfully; compare to generator baseline to confirm 8.4 ± 0.1 pt compliance.

---

#### Outstanding Questions / Assumptions
- Assumed there is no accessible older deck snapshot beyond the regenerated run; if archives surface, they could speed QA.
- Need confirmation that downstream consumers tolerate zero-width spaces (so far no negative reports).
- Pending decision: whether to move the Verdana enforcement entirely into the COM script and remove it from generation to avoid double-handling.

---

#### Related Artefacts
- Logs: `docs/22-10-25/logs/17-font_enforcement_results_20251022_1609.md`, `19-postprocess_run_20251022_1626.md`, `20-revert_regeneration_plan_20251022_1650.md`, `21-deck_regeneration_20251022_1648.md`.
- CSVs: `docs/22-10-25/artifacts/row_height_probe_20251022_1617.csv`, `dash_font_check_20251022_1608.txt`.
- Scripts touched: `amp_automation/presentation/tables.py`, `tools/PostProcessCampaignMerges.ps1`, `tools/ProbeRowHeights.ps1`.
