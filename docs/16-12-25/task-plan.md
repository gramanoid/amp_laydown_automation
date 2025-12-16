# Task Plan - 16-12-25

## Overview
Continuing from 15-12-25 session. Product split implementation complete, pending commit and validation.

## Next Actions (Priority Order)

### 1. Commit Pending Changes [HIGH - Blocking]
- **Why:** 5 modified files with significant changes (+1023/-732 lines) sitting uncommitted
- **Entry:** `git add` + `git commit`
- **Risk:** None - changes already validated in previous session
- **Files:**
  - `amp_automation/presentation/assembly.py` (+222 lines - product split logic)
  - `config/master_config.json` (restructured - product splits for 3 brands)
  - `amp_automation/data/ingestion.py` (+28 lines - supporting changes)
  - `docs/15-12-25/` (session docs + retrospective)

### 2. Run Test Suite [HIGH - Validation]
- **Why:** Confirm no regressions after product split changes
- **Entry:** `pytest` or `pytest -m unit`
- **Risk:** Low - validated manually in 15-12-25

### 3. Update Structural Validator [MEDIUM - Tech Debt]
- **Why:** Contract needs update (GRAND TOTAL -> BRAND TOTAL)
- **Entry:** `tools/validate/validate_structure.py`
- **Risk:** May need additional investigation

## Blockers
- None identified

## Recommended Commands
```bash
# Step 1: Review changes
git diff --stat

# Step 2: Commit
git add -A && git commit -m "feat: add product splits for Parodontax, Sensodyne, Sensodyne Pronamel"

# Step 3: Test
pytest
```

## Generated
2025-12-16T14:32:00Z
