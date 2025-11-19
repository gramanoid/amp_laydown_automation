# Agent Guide: AMP Laydowns Automation

**Last verified on 29-10-25**

<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

## Quick Project Recap (29 Oct 2025)
Last verified on 29-10-25
- **Mission:** Clone-based generation of AMP laydown decks that match `Template_V4_FINAL_071025.pptx` pixel-for-pixel while binding Lumina Excel data.
- **Latest Baseline Deck:** `output/presentations/run_20251029_132708/AMP_Laydowns_291025.pptx` (120 slides with accuracy validation).
- **CRITICAL ARCHITECTURE DECISION (24 Oct 2025):** PowerPoint COM automation for bulk operations is PROHIBITED due to catastrophic performance issues (10+ hours vs Python's 10 minutes - 60x difference). See `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`.
- **Session 29 Oct Status:**
  - ✅ 3 automated Excel transformations (Expert exclusion, geography normalization, Panadol splitting)
  - ✅ Media split percentages with color coding
  - ✅ Comprehensive accuracy validation system
  - ✅ M-suffix formatting improvements (2 decimals for millions)
  - ✅ GRP metrics data consistency fixes
  - 🗂️ **Deferred (Phase 4+):** Slide 1 geometry parity, visual diff workflow, automated regression scripts
- **Production Status:** All accuracy improvements validated. Zero large errors (>£5K).

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->

---

## Role

Technical implementation agent for AMP Laydowns automation pipeline. Responsible for:
- Excel data ingestion and transformation
- PowerPoint deck generation via python-pptx template cloning
- Post-processing pipeline execution (8 steps: normalization, merges, styling)
- Accuracy validation and testing
- Session documentation and change management

---

## Tools & Capabilities

### Core Commands

**Generation:**
```bash
# Generate deck from Excel data
uv run python -m amp_automation.cli.main \
  --excel template/BulkPlanData_2025_10_28.xlsx

# Alternative: direct module invocation
python -m amp_automation.cli.main --excel [path]
```

**Testing:**
```bash
# Run all tests
pytest

# Run by marker
pytest -m unit           # Unit tests only
pytest -m integration    # Integration tests (use real data)
pytest -m regression     # Regression tests
pytest -m slow           # Slow tests (>5s)

# Run specific test files
pytest tests/test_accuracy_validation.py -v
pytest tests/test_tables.py tests/test_structural_validator.py -v
```

**Validation:**
```bash
# Structural validation
python tools/validate/validate_structure.py

# Accuracy validation (inline in tests)
pytest tests/test_accuracy_validation.py::test_latest_deck_accuracy
pytest tests/test_accuracy_validation.py::test_production_deck_accuracy
```

**OpenSpec:**
```bash
# List active changes and specs
openspec list
openspec list --specs

# Validate changes
openspec validate [change-id] --strict

# Archive completed work
openspec archive [change-id] --yes
```

### File Structure

```
amp_automation/
├── cli/main.py                    # CLI entry point
├── data/ingestion.py              # Excel loading + transformations
├── presentation/
│   ├── assembly.py                # Deck generation orchestration
│   ├── tables.py                  # Table creation and styling
│   ├── postprocess/               # 8-step post-processing
│   │   ├── cell_merges.py         # Cell merge operations
│   │   ├── table_normalizer.py   # Content normalization
│   │   └── ...
│   └── template_clone.py          # Template cloning utilities
├── validation/
│   └── accuracy_validator.py     # Comprehensive accuracy checks
└── config.py                      # Configuration loading

tests/
├── test_accuracy_validation.py   # Accuracy validation tests
├── test_tables.py                # Table generation tests
└── test_structural_validator.py  # Structural contract tests

tools/
├── validate/
│   ├── validate_structure.py     # Structural validator
│   └── validate_all_data.py      # Legacy comprehensive validator
└── verify/
    └── verify_deck_fonts.py      # Font verification

template/
├── Template_V4_FINAL_071025.pptx # Master template (PROTECTED)
└── BulkPlanData_2025_10_28.xlsx  # Latest Excel data (PROTECTED)

openspec/
├── project.md                    # Project conventions
├── AGENTS.md                     # OpenSpec workflow guide
├── DEFERRED.md                   # Phase 4+ deferred items
├── specs/                        # Current capabilities
└── changes/                      # Active proposals
```

---

## Guardrails

### Forbidden Actions

**NEVER modify these files:**
- `template/Template_V4_FINAL_071025.pptx` (master template)
- `template/BulkPlanData_*.xlsx` (source data)
- `output/presentations/**/*.pptx` (generated decks)
- `docs/*/logs/*.json` (session logs)
- `.venv/**` (virtual environment)
- `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` (architecture decision record)

**NEVER use PowerPoint COM for:**
- Bulk table operations
- Cell merging
- Content normalization
- Styling operations
- Anything that can be done with python-pptx

**COM permitted ONLY for:**
- File I/O (open/save/close)
- PDF export
- Features unavailable in python-pptx (e.g., built-in animations)

### Code Standards

**Python:**
- Type hints required for all functions
- Google-style docstrings for public APIs
- snake_case for functions/variables
- PascalCase for classes
- UPPER_SNAKE_CASE for constants
- dataclasses with slots=True for performance
- Line length: 88-120 characters

**Testing:**
- Use pytest markers: @pytest.mark.unit, @pytest.mark.integration, @pytest.mark.regression
- Test file naming: test_*.py
- Function naming: test_*
- Minimum 70% coverage for critical paths
- Always test edge cases and error conditions

**Commits:**
- Format: `<type>: <description>` (e.g., "feat: add accuracy validation")
- Types: feat, fix, docs, test, refactor, chore
- Include file references when relevant
- Add `🤖 Generated with [Claude Code](https://claude.com/claude-code)` footer
- Add `Co-Authored-By: Claude <noreply@anthropic.com>` trailer

### Security

- No hardcoded credentials (use .env)
- All secrets in .gitignore
- Excel data files excluded from git
- Generated decks excluded from git
- Never commit sensitive client data

---

## Workflow

### Standard Session Flow

1. **Initialize** - Read CLAUDE.md, AGENTS.md, openspec/project.md
2. **Check context** - Run `openspec list` to see active changes
3. **Plan** - Create proposal for new features (see openspec/AGENTS.md)
4. **Implement** - Follow tasks.md checklist, mark completed with [x]
5. **Test** - Run pytest with appropriate markers
6. **Validate** - Run structural and accuracy validators
7. **Document** - Update PROJECT_STATUS and session log
8. **Commit** - Use conventional commit format

### Data Transformation Pipeline

**Location:** `amp_automation/data/ingestion.py`

**Current transformations (lines 107-174):**
1. Expert campaign exclusion (Plan Name contains "expert")
2. Geography normalization (11 mapping rules: FWA→FSA, GINE→GNE, etc.)
3. Panadol brand splitting (Pain vs C&F based on Product Business column)

**When adding transformations:**
- Apply to both main pipeline AND `get_month_specific_tv_metrics()` cache (lines 414-460)
- Log transformation results (rows affected, splits created)
- Test with production data
- Update ingestion.py docstrings

### Accuracy Validation

**Location:** `amp_automation/validation/accuracy_validator.py`

**Validates:**
- Horizontal totals (row sums across months = TOTAL column)
- Vertical totals (MONTHLY TOTAL rows = sum of rows above)
- Cell data mapping (PowerPoint matches Excel source)
- GRP metrics accuracy

**Tolerance:** 1% for rounding, ignores <£1 differences

**Test usage:**
```python
from amp_automation.validation.accuracy_validator import validate_deck_accuracy

report = validate_deck_accuracy(pptx_path)
print(report.summary())
assert report.passed, f"Validation failed with {report.error_count} errors"
```

### Presentation Generation

**Key functions:**
- `_build_campaign_block()` - Assembles campaign data
- `_build_campaign_monthly_total_row()` - Creates TOTAL rows with media splits
- `_build_budget_row()` - Formats budget rows
- `_format_total_budget()` - Formats totals (K/M suffix with precision)

**Post-processing:** 8-step Python pipeline
1. Cell merges (campaign names, media cells, MONTHLY TOTAL rows)
2. Content normalization
3. Font consistency
4. Border styling
5. Color fills
6. Alignment
7. Validation
8. Cleanup

---

## Escalation

**When to ask for approval:**
- Breaking changes to table structure
- New data transformations affecting >100 rows
- Architecture decisions (new dependencies, patterns)
- Changes to template geometry or layout
- Modifications to protected files

**When to create OpenSpec proposal:**
- New features or capabilities
- Breaking API/schema changes
- Architecture or pattern changes
- Performance optimizations changing behavior
- Security pattern updates

**When to proceed directly:**
- Bug fixes restoring intended behavior
- Typos, formatting, comments
- Non-breaking dependency updates
- Configuration changes
- Tests for existing behavior

**Blocked by:**
- Missing clarification on ambiguous requirements
- Need for user approval on design decisions
- External dependencies or data not available
- Technical limitations requiring architecture review

---

Last verified on 29-10-25
