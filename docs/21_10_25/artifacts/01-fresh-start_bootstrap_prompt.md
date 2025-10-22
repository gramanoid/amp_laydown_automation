MISSING: D:\Drive\projects\work\AMP Laydowns Automation\openspec\README.md (no primary spec entry located)

Zero-question rule: Do not ask the user anything. When uncertain, lay out two options with pros/cons, choose one, and state the assumption you are making before continuing.

**Alignment Check — Overview:** Automate AMP media laydown decks so generated PowerPoint files mirror `Template_V4_FINAL_071025.pptx` while binding Lumina Excel data and enforcing table geometry via COM tooling.

- **Project Summary:** Generator outputs decks such as `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251021_185319\GeneratedDeck_20251021_MergedCells.pptx`, with COM post-processing (`tools/PostProcessCampaignMerges.ps1`) merging campaign blocks, monthly totals, and normalising fonts/percentages to match the template.
- **NOW Tasks (Acceptance):** Implement a column 1–3 unmerge pre-pass in `tools/PostProcessCampaignMerges.ps1`, rerun it on the latest deck, and confirm the COM probe reports 8.4 pt ± 0.1 pt for every data row across slides (capture logs/screenshots).
- **Biggest Risk & Mitigation:** COM automation keeps prior merged geometry, causing random row inflation; mitigate by resetting table cells before merges and adding automated validation to flag unexpected spans early.

## Brain Reset Digest
### Session Overview
Automation currently relies on a COM post-pass to achieve pixel-parity with the AMP laydown template; 21 Oct 2025 efforts centred on resolving residual row-height drift after campaign/monthly total merges.

### Work Completed
- Authored and executed `tools/PostProcessCampaignMerges.ps1` to merge campaign-name spans and monthly totals while enforcing fonts (header 7 pt, body 6 pt, totals 6.5 pt) and reapplying 8.4 pt row heights.
- Normalised percentage formatting and campaign typography in the generated deck (`GeneratedDeck_20251021_MergedCells.pptx`) with structural validation remaining green.
- Investigated lingering tall rows, tracing them to non-idempotent monthly-total merges that persist across repeated COM executions (observed on Slides 2, 3, 5, 16).

### Current State
Row-height probes still report 14.4–21.6 pt on select rows despite enforcement loops because reruns inherit old merges. Structural checks and template geometry alignment remain intact; visual diff artefacts are pending until row heights stabilise.

### Purpose
Deliver cloned AMP laydown decks that match `Template_V4_FINAL_071025.pptx` pixel-for-pixel while binding Lumina Excel data, backed by deterministic table geometry and documented automation (per OpenSpec guardrails).

### Next Steps
- [ ] Unmerge-prepass implemented and committed.
- [ ] COM post-processor rerun shows correct campaign row structure on Slides 2/3/5/16.
- [ ] Row-height probe logs confirm 8.4 pt data rows across the latest deck.
- [ ] Automated merge/height validation script added to tooling.
- [ ] Visual diff artefacts refreshed and archived once metrics align.
- Add automated validation (Python/COM) scanning all slides for rogue merges or oversized rows.
- Regenerate Slide 1 visual diff artefacts and archive evidence after geometry stabilises.
- Restore the pytest suite and smoke tests to cover future regressions.

### Important Notes
- Last verified on 2025-10-21 per `BRAIN_RESET_211025.md`; snapshots share the same timestamp.
- COM automation is mandatory for post-processing and probing; ensure PowerPoint is installed with trusted VBA/COM access.
- Follow OpenSpec conventions (`openspec/AGENTS.md`): use verb-led change IDs, validate with `openspec validate <id> --strict`, and keep specs/tasks in sync.
- Historical tests under `tests/` are currently deleted; reinstate before relying on pytest instructions in runbooks.

### Session Metadata
- Documents consulted: `D:\Drive\projects\work\AMP Laydowns Automation\docs\21_10_25\BRAIN_RESET_211025.md`, `...\BRAIN_RESET.md`, `...\SNAPSHOT_2025-10-21.md`, and `D:\Drive\projects\work\AMP Laydowns Automation\openspec\AGENTS.md`.
- Latest artefact: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251021_185319\GeneratedDeck_20251021_MergedCells.pptx`.
- Timezone basis: Abu Dhabi/Dubai (UTC+04) — today is 21-10-25 (DD-MM-YY).

**Workflow Directive:** Operate in the Plan → Change → Test → Document → Commit loop. Reference runbook commands (`python -m amp_automation.cli.main`, `tools/validate_structure.py`, `& tools/PostProcessCampaignMerges.ps1`, COM probe scripts) during execution, respect inherited checklists, keep all paths absolute (e.g., `D:\Drive\projects\work\AMP Laydowns Automation\...`), avoid exposing secrets, and restore automated tests/linters whenever instructions call for them before staging or pushing.
