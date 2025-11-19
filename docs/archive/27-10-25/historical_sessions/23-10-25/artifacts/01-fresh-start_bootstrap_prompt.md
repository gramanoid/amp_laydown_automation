ASSUME: Seed date defaults to 23-10-25 (Abu Dhabi/Dubai, UTC+04) because no explicit seed token was supplied.
MISSING: Root cause for `D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessCampaignMerges.ps1` throwing `Unexpected HRESULT` during `Presentations.Open` remains undocumented.
MISSING: Verified 8.4 +/- 0.1 pt row-height probe results for `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx` are not yet captured.

Zero-question rule: Do not ask the user anything. When uncertain, outline two viable options with pros/cons, choose the path you recommend, execute it, and record the assumption inline.

Alignment Check:
One-line overview: Automate pixel-accurate AMP laydown decks from Lumina Excel data using template cloning plus COM post-processing within `D:\Drive\projects\work\AMP Laydowns Automation`.
- Project summary: Daily context lives in `D:\Drive\projects\work\AMP Laydowns Automation\docs\23-10-25\23-10-25.md`; focus on stabilising `tools\PostProcessCampaignMerges.ps1`, keeping horizontal merge governance via `tools\FixHorizontalMerges.ps1`, and preserving deck fidelity recorded in `docs\22-10-25\` logs.
- NOW tasks (acceptance criteria): Harden the COM post-process script so it opens, processes, and closes PowerPoint cleanly; rerun campaign/monthly merges plus the row-height probe on `output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx` with artefacts stored under `docs\23-10-25\artifacts\` and `docs\23-10-25\logs\`; confirm every data row reports 8.4 +/- 0.1 pt.
- Biggest risk + mitigation: COM automation instability (`Unexpected HRESULT`, lingering POWERPNT.exe) can corrupt decks and block evidence capture—pre-empt with explicit `Stop-Process POWERPNT`, resilient retries, and prompt archival of COM logs for post-mortem analysis.

Brain Reset Digest
Last Updated: 2025-10-23 (`D:\Drive\projects\work\AMP Laydowns Automation\docs\23-10-25\BRAIN_RESET_231025.md`)
#### Session Overview
- Clone pipeline mirrors `Template_V4_FINAL_071025.pptx`; AutoPPTX stays disabled apart from negative tests.
- Today extends the 2025-10-22 COM hardening effort while staging fresh artefact/log destinations under `docs\23-10-25\`.
#### Work Completed
- Horizontal merge cleanup via `tools\FixHorizontalMerges.ps1` leaves only MONTHLY TOTAL / GRAND TOTAL / CARRIED FORWARD spans per `docs\22-10-25\merged_cells_analysis`.
- Geometry resets and path normalisation already landed in `tools\PostProcessCampaignMerges.ps1`, though the script still faults during `Presentations.Open`.
- Regenerated deck `output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx` (561,713 bytes) awaits successful post-processing.
#### Current State
- `tools/PostProcessCampaignMerges.ps1` throws `Unexpected HRESULT`, leaving orphaned PowerPoint sessions and preventing campaign/monthly merges plus row-height locks.
- No fresh row-height probe exists for the 2025-10-22 deck, so Slide 1 parity, visual diff, and Zen MCP evidence remain blocked.
- Outstanding compliance check still references `output/presentations/run_20251021_185319/` for historical hygiene.
#### Purpose
- Deliver pixel-accurate AMP laydown decks with enforced row heights, template EMUs, and validated merges, ensuring evidence is archived for every critical run.
#### Next Steps
- [ ] Make `tools/PostProcessCampaignMerges.ps1` idempotent and resilient to Protected View/retry scenarios (explicit `Quit()`, back-off retries, guaranteed `Stop-Process POWERPNT`).
- [ ] Rerun the campaign/monthly merge script on `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx` and confirm campaign spans remain intact even when errors occur.
- [ ] Execute the COM row-height probe against the updated deck and ensure all data rows land at 8.4 +/- 0.1 pt, storing output in `docs\23-10-25\artifacts\`.
- [ ] Archive COM troubleshooting results in `docs\23-10-25\logs\` and update the daily artefacts accordingly.
- [ ] Run the generation, validation, and test commands successfully.
- [ ] Confirm COM probe reports 8.4 +/- 0.1 pt across all Slide 1 body rows.
- [ ] Document root cause for COM-driven row expansion and mitigation.
- [ ] Verify latest deck lives under `output/presentations/run_20251021_185319/` (historical compliance check).
#### Important Notes
- Always close lingering POWERPNT.exe instances (`Stop-Process -Name POWERPNT -Force`) before COM automation.
- Maintain absolute Windows paths in documentation; avoid secrets in artefacts.
- Restrict horizontal merges to MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD labels; treat others as regressions.
- Preserve template EMUs, centered alignment, and font sizing dictated by `Template_V4_FINAL_071025.pptx`.
#### Session Metadata
- Latest deck: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_115646\GeneratedDeck_20251022_132000.pptx`
- Key references: `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\02-postprocess_attempt.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\03-deck_regeneration.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\04-merged_cells_cleanup.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\merged_cells_cleanup_20251022.txt`
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 23-10-25 (DD-MM-YY).

Workflow Directive:
Follow `Plan → Change → Test → Document → Commit` rigorously. Plan using `D:\Drive\projects\work\AMP Laydowns Automation\docs\23-10-25\BRAIN_RESET_231025.md` and `openspec\project.md`; adjust code in place; test with `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\test_tables.py tests\test_structural_validator.py` plus deck validation commands; document outcomes under `docs\23-10-25\`; only commit once artefacts, logs, and checklists reflect reality. Honor inherited guardrails: absolute path references, no secrets, respect OpenSpec change management, and keep the COM automation/TODO checklists current.
