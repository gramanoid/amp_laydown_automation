# Project-Level Claude Code Configuration

**Project:** AMP Laydowns Automation
**Created:** 29-10-25
**Last Updated:** 29-10-25

---

## Overview

Automates Annual Marketing Plan laydown decks by converting standardized Lumina Excel exports into pixel-accurate PowerPoint presentations that mirror the master template while preserving financial and media metrics.

**Tech Stack:** Python 3.13.x, pandas, numpy, python-pptx, openpyxl, pytest 8.x
**Architecture:** CLI-driven template cloning with Python post-processing pipeline, validation suite, OpenSpec change management

---

## Project-Specific Overrides

These settings override the global CLAUDE.md defaults for this project only.

### Custom Time Limits

```
CLAUDE_MAX_SESSION_SECONDS=14400  # 4 hours (extended for complex deck generation tasks)
CLAUDE_TIMEZONE=Asia/Dubai        # Arabian Standard Time UTC+4
```

### Protected Files (Never Modify)

Files that should never be edited by Claude Code:

- template/Template_V4_FINAL_071025.pptx
- template/BulkPlanData_*.xlsx
- output/presentations/**/*.pptx (generated decks)
- docs/*/logs/*.json (historical session logs)
- docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md
- .venv/** (virtual environment)

### Required Tests

Tests that MUST pass before any commit:

- Unit tests: `pytest -m unit`
- Integration tests: `pytest -m integration`
- All tests: `pytest`
- Structural validation: `python tools/validate/validate_structure.py`

---

## Guardrails

### Forbidden Actions

- Do NOT modify template/Template_V4_FINAL_071025.pptx (master template)
- Do NOT delete output/presentations/ folders without explicit approval
- Do NOT commit .env files, credentials, or Excel source data
- Do NOT use PowerPoint COM for bulk table operations (see ARCHITECTURE_DECISION_COM_PROHIBITION.md)

### Code Style

- Python: Type hints required, Google-style docstrings, snake_case, dataclasses with slots=True
- Testing: pytest with markers (@pytest.mark.unit, @pytest.mark.integration, @pytest.mark.regression)
- Formatting: Black-compatible (88-120 char line length)
- See AGENTS.md for detailed build/test/style commands

### Security Considerations

- No hardcoded credentials in code
- All secrets must use `.env` file
- Excel data files excluded from git (.gitignore)
- Generated decks excluded from git (output/presentations/)

---

## Integration Notes

- **AGENTS.md:** Refer to project's AGENTS.md for build commands, test commands, and deployment procedures
- **OpenSpec:** Project uses OpenSpec for specifications and change proposals (openspec/changes/)
- **Daily State:** Use docs/{DD-MM-YY}/PROJECT_STATUS_{DD-MM-YY}.md for session tracking
- **Task Marking:** Mark completed tasks with [x] in openspec/changes/{feature}/tasks.md

---

## Reference

- Global CLAUDE.md: `C:\Users\alexg\.claude\CLAUDE.md`
- Project AGENTS.md: `D:\Drive\projects\work\AMP Laydowns Automation\AGENTS.md`
- OpenSpec: `D:\Drive\projects\work\AMP Laydowns Automation\openspec\`
