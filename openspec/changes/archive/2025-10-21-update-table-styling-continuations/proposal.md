## Why
Current Template V4 migration still renders tables with legacy cell styling and lacks continuation-slide parity, risking visual drift and misaligned totals in client decks.

## What Changes
- Implement per-cell styling that matches Template V4 fonts, alignment, fills, wrapped text, and dual-line media labels.
- Enhance continuation-slide generation to retain headers, banding, carried subtotals, and summary tiles across split tables.
- Extend reconciliation coverage to assert new styling metadata where detectable and keep tolerance rules intact.

## Impact
- Affected specs: `presentation`
- Affected code: `amp_automation/presentation/assembly.py`, `amp_automation/presentation/tables.py`, `config/master_config.json`, `tests/test_template_validation.py`
