# AMP Laydowns Automation
**Last Updated:** 27-10-25

## Purpose
Automates PowerPoint presentation generation for Advertising Media Planning (AMP) laydowns. Converts Lumina Excel exports into decks that mirror `Template_V4_FINAL_071025.pptx` while preserving template geometry, fonts, and layout.

**CRITICAL:** PowerPoint COM automation for bulk operations is PROHIBITED. Performance testing (24 Oct 2025) proved COM takes 10+ hours vs Python's 10 minutes (60x difference). See `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`.

## Contents
- `amp_automation/` - Python pipeline (CLI, data processing, presentation generation)
- `amp_automation/presentation/postprocess/` - Python post-processing (normalization, merges)
- `tools/validate/` - Data and structural validation tools
- `tools/verify/` - Verification and post-processing checks
- `docs/archive/` - Deprecated scripts and historical session documentation
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
python tools/validate/validate_all_data.py output/presentations/run_*/presentations.pptx \
  --excel template/BulkPlanData_2025_10_14.xlsx
```

## Dependencies
- Python 3.9+, python-pptx 1.0.2, pandas, openpyxl
- See `requirements.txt` for full package list

## Testing & Validation
```bash
pytest tests/  # Unit tests
python tools/validate/validate_all_data.py  # Comprehensive validation (accuracy, format, completeness, reconciliation)
python tools/validate/validate_structure.py  # Structural contract validation
python tools/verify/verify_deck_fonts.py  # Font verification
```

Full pipeline: generation → Python post-processing → validation. Target: <20 minutes for 88 slides.

## Notes
SKIPPED: Section content remains current as of 27-10-25
- Use Python (python-pptx) for ALL bulk table operations
- COM permitted ONLY for file I/O, exports, features unavailable in python-pptx
- Performance targets: deck generation <5min, normalization <5min, merges <10min
- See `docs/27-10-25/BRAIN_RESET_271025.md` for current project status

Last verified on 27-10-25
