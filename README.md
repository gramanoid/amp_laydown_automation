# AMP Laydowns Automation
**Last Updated:** 24-10-25

## Purpose
Automates PowerPoint presentation generation for Advertising Media Planning (AMP) laydowns. Converts Lumina Excel exports into decks that mirror `Template_V4_FINAL_071025.pptx` while preserving template geometry, fonts, and layout.

**CRITICAL:** PowerPoint COM automation for bulk operations is PROHIBITED. Performance testing (24 Oct 2025) proved COM takes 10+ hours vs Python's 10 minutes (60x difference). See `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`.

## Contents
- `amp_automation/` - Python pipeline (CLI, data processing, presentation generation)
- `amp_automation/presentation/postprocess/` - Python post-processing (normalization, merges)
- `tools/` - Validation scripts and DEPRECATED PowerShell COM scripts
- `template/` - Excel data and PowerPoint template (V4 FINAL)
- `docs/` - Architecture decisions, daily logs, project status
- `openspec/` - Change proposals and project context

## Usage
Generate deck:
```bash
python -m amp_automation.cli.main --excel template/BulkPlanData_2025_10_14.xlsx \
  --template template/Template_V4_FINAL_071025.pptx --output output/presentations
```

Post-process (Python):
```bash
python -m amp_automation.presentation.postprocess.cli \
  --presentation-path output/presentations/run_*/presentations.pptx \
  --operations normalize --verbose
```

Validate:
```bash
python tools/validate_structure.py output/presentations/run_*/presentations.pptx \
  --excel template/BulkPlanData_2025_10_14.xlsx
```

## Dependencies
- Python 3.9+, python-pptx 1.0.2, pandas, openpyxl
- PowerShell 7+ and PowerPoint (legacy scripts only)

## Testing & Validation
```bash
pytest tests/  # Unit tests
python tools/validate_structure.py  # Structural validation
```

Full pipeline: generation → Python post-processing → validation. Target: <20 minutes for 88 slides.

## Notes
- Use Python (python-pptx) for ALL bulk table operations
- COM permitted ONLY for file I/O, exports, features unavailable in python-pptx
- Performance targets: deck generation <5min, normalization <5min, merges <10min
- See `docs/24-10-25/BRAIN_RESET_241025.md` for current project status
