ASSUME: No new presentation run exists beyond `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx`.
ASSUME: Abu Dhabi/Dubai (UTC+04) is the authoritative timezone; today is 24-10-25 (DD-MM-YY).
MISSING: Sanitized deck logs and row-height probe outputs for 24 Oct 2025 (CLI regeneration pending).

Zero-question rule: Do not ask the user for clarification. When uncertainty arises, draft two viable options with pros/cons, select the better fit, proceed under that assumption, and record the choice in the handoff log.

Alignment Check —
Project overview: Automated AMP laydown decks must match `Template_V4_FINAL_071025.pptx` while binding Lumina Excel data.
- Project summary: Clone-based PPTX generator with COM post-processing enforces template geometry, typography, and merge rules documented in `D:\Drive\projects\work\AMP Laydowns Automation\openspec\project.md`.
- Now tasks (acceptance criteria): Generate a fresh CLI deck with debug logging reduced (`docs\24-10-25\logs\` contains new run metadata), sanitize columns 1-3 without corrupting campaign spans, execute `tools\PostProcessCampaignMerges.ps1` plus COM row-height probe, and archive outputs in `docs\24-10-25\artifacts\`.
- Biggest risk + mitigation: PowerPoint COM instability may corrupt decks; pre-emptively stop lingering POWERPNT processes, run scripts serially, and capture failure logs for rollback.

Brain Reset Digest
Last Updated: 2025-10-24 (`D:\Drive\projects\work\AMP Laydowns Automation\docs\24-10-25\BRAIN_RESET_241025.md`)
Session Overview:
- Clone pipeline mirrors `Template_V4_FINAL_071025.pptx` using Lumina Excel exports with python-pptx plus COM post-processing; AutoPPTX remains disabled.

Work Completed:
- Baseline deck `output\presentations\baseline_20251022\GeneratedDeck_20251022_160155_baseline.pptx` preserved for geometry validation.
- Sanitizer prototypes (`tools\SanitizePrimaryColumns.ps1`, `tools\RebuildCampaignMerges.ps1`) authored alongside handoff notes `docs\23-10-25\logs\05-postprocess_hand_off.md`.
- Environment tidied by terminating stray PowerPoint sessions on 23 Oct 2025.

Current State:
- CLI regeneration blocks on prolific DEBUG logging; no 24 Oct run folder exists yet.
- Sanitizer/merge workflow still unvalidated; prior attempts corrupted campaign columns when fed legacy decks.
- Row-height probes and visual diff evidence remain missing for a clean deck.

Purpose:
- Deliver pixel-faithful AMP laydown decks by converting Lumina Excel exports while preserving 8.4 pt row heights, centered alignment, and restricted merge allowances.

Next Steps:
- Now: Suppress presentation assembly logging, regenerate a clean deck from the 20251022 baseline with run metadata stored under `docs\24-10-25\logs\`, and sanitize columns 1-3 on a disposable copy before merging.
- Next: Execute `tools\PostProcessCampaignMerges.ps1`, capture COM row-height probe output in `docs\24-10-25\artifacts\`, update logs with pass/fail notes, and draft deterministic merge rebuild guidance.
- Later: Resume Slide 1 EMU/legend parity, refresh visual diff + Zen MCP evidence, restore pytest suites, and design no-campaign-splitting pagination plus OpenSpec change proposal.
- Immediate TODOs:
  - [ ] Regenerate a fresh deck and log the new run under `docs\24-10-25\`.
  - [ ] Sanitize columns 1-3 (unmerge + reset layout) across all slides before running COM merges.
  - [ ] Run `tools\PostProcessCampaignMerges.ps1 -Verbose` on the sanitised deck and confirm no residual span warnings remain.
  - [ ] Execute the COM row-height probe and store results in `docs\24-10-25\artifacts\`.
  - [ ] Update `docs\24-10-25\logs\` with outcomes (success/failure) for each pass.
- Longer-Term Follow-Ups:
  - Add automated regression scripts catching rogue merges or row-height drift before decks ship.
  - Complete Slide 1 geometry/legend parity fixes and refresh visual diff plus Zen MCP/Compare workflows for evidence capture.
  - Rehydrate pytest suites and smoke tests so pipeline regressions surface automatically.
  - Facilitate a Q&A-led discovery to design campaign pagination that prevents splits across slides, then raise an OpenSpec change once prerequisites clear.
- Outstanding checklist (carry forward):
  - [ ] Regenerate deck + validation logs for 24 Oct 2025 and capture run metadata.
  - [ ] Confirm COM probe reports 8.4 +/- 0.1 pt across all Slide 1 body rows.
  - [ ] Document root cause for legacy deck span explosion and mitigation strategy.
  - [ ] Reproduce visual diff / Zen MCP evidence once geometry stabilises.

Important Notes:
- Close POWERPNT before automation (`Stop-Process -Name POWERPNT -Force`); scripts do not self-clean.
- Keep documentation on absolute Windows paths; never include secrets in logs.
- Only MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD may remain merged across columns 1-3.
- Template EMUs, centered alignment, and font sizes must match `Template_V4_FINAL_071025.pptx`.

Session Metadata:
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx`.
- Key logs: `docs\22-10-25\logs\03-deck_regeneration.md`, `docs\22-10-25\logs\02-postprocess_attempt.md`, `docs\22-10-25\logs\04-merged_cells_cleanup.md`, `docs\23-10-25\logs\05-postprocess_hand_off.md`.
- Artifacts: `docs\22-10-25\artifacts\merged_cells_cleanup_20251022.txt`.
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 24-10-25 (DD-MM-YY).

Workflow Directive:
Follow the `Plan → Change → Test → Document → Commit` loop. Plan by reviewing `docs\24-10-25\BRAIN_RESET_241025.md` and relevant OpenSpec entries; stage changes with deterministic scripts (e.g., `python -m amp_automation.cli.main`, `tools\PostProcessCampaignMerges.ps1`) while capturing outputs under `docs\24-10-25\`. Test using `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest …` and any COM probes before sharing results. Document every run in the daily log and update checklists. Commit only after validations pass, referencing change IDs when applicable. Honour guardrails: stay within `D:\Drive\projects\work\AMP Laydowns Automation`, retain absolute paths, avoid secrets, and run required linters/tests prior to handoff.
