ASSUME: Treating D:\Drive\projects\work\AMP Laydowns Automation\openspec\project.md as the primary spec entry because README.md is absent.
MISSING: D:\Drive\projects\work\AMP Laydowns Automation\openspec\README.md (primary spec entry not present).

Zero-question rule: Do not ask the user anything; when uncertain, outline two actionable options with pros/cons, choose one, and record the chosen assumption in D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\.

Alignment Check - Generate AMP laydown decks that clone Template_V4_FINAL_071025.pptx with Lumina Excel data plus COM tooling.
- **Project Summary:** Pipeline run `python -m amp_automation.cli.main --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx --template D:\Drive\projects\work\AMP Laydowns Automation\template\Template_V4_FINAL_071025.pptx --output D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_092659\GeneratedDeck_20251022_092659.pptx` builds the latest deck, followed by `tools\FixHorizontalMerges.ps1` (JSON-guided) to leave only MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD spans across columns 0-2.
- **NOW Tasks (Acceptance):** Refactor `D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessCampaignMerges.ps1` to unmerge entire campaign blocks (rows and columns) before reapplying merges, rerun it on `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_092659\GeneratedDeck_20251022_092659.pptx`, confirm the row-height probe (`D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\row_height_probe_20251022.txt`) reports 0 rows outside 8.4 +/- 0.1 pt, and capture validation artefacts plus COM logs under `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\`.
- **Biggest Risk & Mitigation:** PowerPoint COM still raises "Cannot merge cells of different sizes" and leaves partial merges that inflate row heights; mitigate by scripting a full geometry reset, killing lingering `POWERPNT` processes before runs, and adding automated sanity checks so only the three summary labels remain merged.

## Brain Reset Digest
### Session Overview
22 Oct 2025 focused on regenerating a clean presentation, auditing JSON-described horizontal merges, and diagnosing why campaign-level COM merges keep failing despite the new split/remerge tooling.

### Work Completed
- Cleared `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\` then regenerated `GeneratedDeck_20251022_092659.pptx` (561,713 bytes) via the CLI command captured in `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\03-deck_regeneration.md`.
- Ran `D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessCampaignMerges.ps1` against `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251021_185319\GeneratedDeck_20251021_MergedCells.pptx`; PowerPoint COM threw repeated `Cell.Merge` size mismatches and the post-run probe found 364 of 1,429 rows outside 8.4 +/- 0.1 pt (`D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\02-postprocess_attempt.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\row_height_probe_20251022.txt`).
- Authored and executed `D:\Drive\projects\work\AMP Laydowns Automation\tools\FixHorizontalMerges.ps1` using `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\merged_cells_analysis\merged_cells_fix_instructions.json`; retained 27 allowed summary spans and removed 53 rogue or empty merges (see `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\04-merged_cells_cleanup.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\merged_cells_cleanup_20251022.txt`).

### Current State
Horizontal merges across the first three columns now match the JSON specification, but vertical campaign merges remain brittle - COM continues to flag geometry mismatches when the deck already contains merged rows, leaving 364 rows above tolerance (earlier notes cited 692 offenders, underscoring the need to stabilise the probe). Slide 1 still requires table EMU/legend parity work, and the PostProcess script must become idempotent before reruns stop compounding errors.

### Purpose
Deliver pixel-perfect AMP laydown decks that mirror `Template_V4_FINAL_071025.pptx` while binding Lumina Excel metrics, using deterministic clone logic plus COM enforcement in line with OpenSpec guardrails.

### Next Steps
- Expand `D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessCampaignMerges.ps1` to split entire campaign and monthly blocks (rows + columns) before merging, skipping any block that lacks uniform geometry.
- Rerun the campaign merge script on `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_092659\GeneratedDeck_20251022_092659.pptx`, immediately follow with the row-height probe, and expect 0 rows beyond 8.4 +/- 0.1 pt (log outputs to `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\`).
- Automate verification that only MONTHLY TOTAL, GRAND TOTAL, and CARRIED FORWARD remain merged (e.g., python-pptx smoke test) so regressions surface without manual JSON audits.
- Refresh Slide 1 visual diff and Zen MCP/Compare evidence once geometry stabilises, archiving artefacts under `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\`.
- Ensure PowerPoint is closed via `Stop-Process -Name POWERPNT -Force` (as needed) before running COM automation to avoid stale sessions.

### Important Notes
- The merged-cells analysis JSON (`D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\merged_cells_analysis\merged_cells_fix_instructions.json`) distinguishes legitimate "monthly merges" (summary labels) from empty spans; only those labels should stay merged across columns 0-2.
- AutoPPTX remains disabled beyond negative tests; template EMU dimensions and centered alignment must be preserved when resetting row heights.
- Visual diff work still depends on the slide-level legend removal and EMU sync outlined in `D:\Drive\projects\work\AMP Laydowns Automation\openspec\project.md`.
- Maintain absolute Windows paths in documentation and keep secrets out of logs; respect OpenSpec change-ID workflow before committing.

### Session Metadata
- Last Updated (`D:\Drive\projects\work\AMP Laydowns Automation\docs\21_10_25\BRAIN_RESET.md`): 2025-10-21.
- Key references: `D:\Drive\projects\work\AMP Laydowns Automation\docs\21_10_25\21_10_25.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\21_10_25\BRAIN_RESET.md`, `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\merged_cells_analysis\merged_cells_fix_instructions.json`, `D:\Drive\projects\work\AMP Laydowns Automation\openspec\AGENTS.md`, `D:\Drive\projects\work\AMP Laydowns Automation\openspec\project.md`.
- Latest deck artefacts: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_092659\GeneratedDeck_20251022_092659.pptx` (clean baseline) and `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251021_185319\GeneratedDeck_20251021_MergedCells.pptx` (postprocess trial).
- Timezone anchor: Abu Dhabi/Dubai (UTC+04); today is 22-10-25 (DD-MM-YY).
- Outstanding Checklist:
  - [ ] Run the generation, validation, and test commands successfully.
  - [ ] Confirm COM probe reports 8.4?pt across all Slide?1 body rows.
  - [ ] Document root cause for COM-driven row expansion and applied mitigation.
  - [ ] Verify current deck lives under `output/presentations/run_20251021_185319/`.

Workflow Directive: Follow the Plan -> Change -> Test -> Document -> Commit loop. Plan by reviewing `D:\Drive\projects\work\AMP Laydowns Automation\openspec\AGENTS.md`, `D:\Drive\projects\work\AMP Laydowns Automation\openspec\project.md`, current change folders, and the JSON/PowerShell artefacts under `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\`. Change within `D:\Drive\projects\work\AMP Laydowns Automation\` by updating clone scripts or COM helpers while honouring template geometry and keeping AutoPPTX disabled. Test each iteration using the CLI generation command noted above, `python D:\Drive\projects\work\AMP Laydowns Automation\tools\validate_structure.py D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251022_092659\GeneratedDeck_20251022_092659.pptx --excel D:\Drive\projects\work\AMP Laydowns Automation\template\BulkPlanData_2025_10_14.xlsx`, `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests\test_autopptx_fallback.py tests\test_tables.py tests\test_assembly_split.py tests\test_structural_validator.py`, plus the COM scripts `D:\Drive\projects\work\AMP Laydowns Automation\tools\PostProcessCampaignMerges.ps1` and `D:\Drive\projects\work\AMP Laydowns Automation\tools\FixHorizontalMerges.ps1` with updated logs. Document every run in `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\logs\` and capture artefacts in `D:\Drive\projects\work\AMP Laydowns Automation\docs\22-10-25\artifacts\`. Commit only after validations pass, absolute paths and inherited checklists are satisfied, and no secrets or unintended files (like regenerated decks) leave the workspace.
