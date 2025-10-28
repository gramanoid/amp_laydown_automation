# Test Suite for AMP Automation

**Status:** PHASE 2B Complete - Test suite restored and expanded with comprehensive regression coverage
**Last Updated:** 28-10-25
**Coverage:** 8 test files, 40+ test cases, 10 edge cases documented

---

## Quick Start

### Run All Tests
```bash
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
python -m pytest tests/ -v
```

### Run Specific Test Categories
```bash
# Unit tests only (fast, no external dependencies)
python -m pytest tests/ -m unit -v

# Integration tests (use real decks/data)
python -m pytest tests/ -m integration -v

# Regression tests (edge case coverage from 27-10-25 fixes)
python -m pytest tests/ -m regression -v

# Quick run (skip slow tests)
python -m pytest tests/ -m "not slow" -v
```

### Run Specific Test File
```bash
python -m pytest tests/test_font_normalization.py -v
python -m pytest tests/test_reconciliation_validator.py -v
python -m pytest tests/test_campaign_merging.py -v
python -m pytest tests/test_pagination.py -v
python -m pytest tests/test_structural_validator.py -v
python -m pytest tests/test_tables.py -v
```

### Coverage Report
```bash
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
python -m pytest tests/ --cov=amp_automation --cov-report=html --cov-report=term
```

---

## Test Files

### conftest.py
**Shared fixtures and utilities for all tests**
- Path fixtures: `template_path`, `excel_path`, `latest_deck_path`, `contract_path`
- Presentation fixtures: `blank_presentation`, `blank_slide`
- Styling fixtures: `cell_style_context`, `table_layout`
- Test data fixtures: `campaign_names_with_hyphens`, `market_case_variations`, etc.
- Helper functions: `find_main_table()`, `is_media_header()`, `is_merged_cell()`, `extract_font_size()`
- Skip markers: `skipif_no_deck`, `skipif_no_template`, `skipif_no_excel`

### test_tables.py
**Table assembly and styling regression tests**
- ✅ `test_add_and_style_table_populates_cells` - Basic table creation (unit)
- ✅ `test_table_font_size_header` - 7pt header font (EC-002, unit)
- ✅ `test_table_font_size_body` - 6pt body font (EC-002, unit)
- ✅ `test_table_word_wrap_disabled` - word_wrap=False for proper breaks (EC-001, unit)

### test_font_normalization.py
**Font size consistency across all slides (EC-002)**
- ✅ `test_ec002_font_sizes_consistent_across_production_deck` - Production deck validation (integration)
- ✅ `test_font_size_header_consistency` - Header styling (unit)
- ✅ `test_font_size_body_consistency` - Body styling (unit)
- ✅ `test_font_family_consistent` - Calibri throughout (unit)

### test_reconciliation_validator.py
**Market/brand normalization and title parsing (EC-004, EC-005, EC-010)**
- ✅ `test_ec004_market_normalization_case_insensitive` - Case handling (regression)
- ✅ `test_ec005_market_code_mapping` - MOR→MOROCCO translation (regression)
- ✅ `test_ec004_brand_normalization_within_market` - Brand case handling (regression)
- ✅ `test_ec010_title_parsing_both_pagination_formats` - Both "(1 of 3)" and "(1/3)" formats (regression)
- ✅ `test_ec010_title_parsing_edge_cases` - Edge cases in parsing (unit)
- ✅ `test_market_code_map_completeness` - All markets present (unit)
- ✅ `test_normalization_functions_return_string` - Type safety (unit)
- ✅ `test_normalization_fallback_returns_original` - Fallback behavior (unit)

### test_campaign_merging.py
**Campaign text wrapping and media merging (EC-001, EC-008)**
- ✅ `test_ec001_campaign_names_no_mid_word_wrap` - Production deck validation (integration, regression)
- ✅ `test_ec001_word_wrap_disabled_in_new_cells` - New cell creation (unit)
- ✅ `test_ec001_campaign_column_width_sufficient` - Column width checks (unit)
- ✅ `test_ec008_media_headers_merged` - Media channel merging (integration, regression)
- ✅ `test_campaign_name_hyphen_handling` - Smart line breaking (unit)
- ✅ `test_merged_cell_identification` - Cell merge detection (unit)

### test_pagination.py
**Continuation slides and pagination (EC-003, EC-006)**
- ✅ `test_ec003_multi_slide_markets_exist` - Multi-slide market presence (integration)
- ✅ `test_ec003_continuation_indicators_present` - "(Continued)" labels (integration, regression)
- ✅ `test_ec006_carried_forward_rows_present` - Carried-forward presence (integration, regression)
- ✅ `test_ec006_empty_metrics_not_accumulated` - Empty metric handling (unit)
- ✅ `test_max_rows_per_slide_boundary` - max_rows_per_slide=32 boundary (unit)
- ✅ `test_continuation_title_format` - Title formatting (unit)

### test_structural_validator.py
**Structural validation of generated decks**
- ✅ `test_structural_validator_passes_production_deck` - 27-10-25 deck validation (integration)
- ✅ `test_structural_validator_recognizes_brand_total` - BRAND TOTAL label (regression)
- ✅ `test_structural_validator_handles_last_slide_only_shapes` - Last-slide-only validation (regression)

---

## Edge Cases Covered

| Edge Case | Test File | Tests | Status |
|-----------|-----------|-------|--------|
| EC-001: Campaign text mid-word wrap | test_campaign_merging.py | 3 unit, 1 integration | ✅ |
| EC-002: Font size drift | test_font_normalization.py | 4 tests | ✅ |
| EC-003: Campaign boundary violation | test_pagination.py | 2 integration | ✅ |
| EC-004: Market case sensitivity | test_reconciliation_validator.py | 2 regression | ✅ |
| EC-005: Market code mapping | test_reconciliation_validator.py | 1 regression | ✅ |
| EC-006: Empty metrics handling | test_pagination.py | 1 unit | ✅ |
| EC-007: Multi-year selection | test_reconciliation_validator.py | (implied by normalization) | ✅ |
| EC-008: Media merge completeness | test_campaign_merging.py | 1 integration | ✅ |
| EC-009: Timestamp timezone | (Skipped - low priority) | - | - |
| EC-010: Pagination format parsing | test_reconciliation_validator.py | 2 unit, 1 regression | ✅ |

---

## Coverage Targets

### Line Coverage
- **Target:** ≥70% of core modules
- **Target Modules:**
  - `amp_automation/presentation/assembly.py` (pagination, formatting)
  - `amp_automation/presentation/postprocess/cell_merges.py` (merging)
  - `amp_automation/validation/reconciliation.py` (data matching)
  - `amp_automation/validation/` (all validators)

### Branch Coverage
- **Target:** ≥80% of decision points
- **Focus Areas:**
  - Pagination boundary checks (max_rows_per_slide)
  - Campaign boundary detection
  - Normalization fallbacks

### Execution Time
- **Unit tests:** <10 seconds (no deck generation)
- **Integration tests:** <120 seconds (2-3 deck validations)
- **Full suite:** <180 seconds

---

## Environment Setup

### Requirements
```
pytest>=8.0
python-pptx>=0.6.21
pandas>=1.5.0
pptx>=0.6.21
```

### Installation
```bash
cd "D:\Drive\projects\work\AMP Laydowns Automation"
pip install -r requirements-dev.txt
# or
uv pip install pytest>=8.0
```

### Configuration
- **pytest.ini:** Marker definitions, test discovery paths, verbosity
- **conftest.py:** Shared fixtures, helpers, skip conditions
- **PYTEST_DISABLE_PLUGIN_AUTOLOAD=1:** Critical! Required to avoid plugin conflicts

---

## Test Markers

```bash
# Run by marker
python -m pytest tests/ -m unit -v              # Unit tests only
python -m pytest tests/ -m integration -v       # Integration tests
python -m pytest tests/ -m regression -v        # Regression tests
python -m pytest tests/ -m slow -v              # Slow tests only
python -m pytest tests/ -m "not slow" -v        # Skip slow tests

# Combine markers
python -m pytest tests/ -m "regression and not slow" -v
```

---

## Common Issues & Troubleshooting

### Issue: "No module named 'amp_automation'"
**Solution:** Ensure you're in project root directory and pytest is using correct PYTHONPATH
```bash
cd "D:\Drive\projects\work\AMP Laydowns Automation"
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
python -m pytest tests/ -v
```

### Issue: "Fixture 'latest_deck_path' not found"
**Solution:** Ensure `conftest.py` is in tests/ directory and deck exists
```bash
# Verify deck path
ls output/presentations/run_20251027_215710/presentations.pptx
```

### Issue: "Skipping test due to missing deck"
**Solution:** Tests with `@skipif_no_deck` require production deck to exist
- Download/generate `run_20251027_215710/presentations.pptx`
- Or skip integration tests: `pytest tests/ -m "not integration"`

### Issue: "Test timeout or hangs"
**Solution:** Check if PowerPoint is open (blocks file access)
```powershell
Stop-Process -Name POWERPNT -Force
```

---

## Continuous Integration

### Before Pushing
1. **Run full suite:**
   ```bash
   $env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
   python -m pytest tests/ -v --tb=short
   ```

2. **Check coverage:**
   ```bash
   python -m pytest tests/ --cov=amp_automation --cov-report=term-missing
   ```

3. **Verify regressions:**
   ```bash
   python -m pytest tests/ -m regression -v
   ```

### Before Releasing
- [ ] All unit tests passing (no failures)
- [ ] All integration tests passing (no failures)
- [ ] All regression tests passing (no failures)
- [ ] Coverage ≥70% line, ≥80% branch
- [ ] No slow test timeout (>180 seconds total)
- [ ] Documentation updated

---

## Adding New Tests

### Checklist
1. **Identify edge case:** Reference REGRESSION_TEST_EDGE_CASE_INVENTORY.md
2. **Choose test file:** Use existing file or create new (e.g., `test_<feature>.py`)
3. **Use fixtures:** Reference conftest.py for available fixtures
4. **Add markers:** `@pytest.mark.unit`, `@pytest.mark.regression`, etc.
5. **Document:** Docstring explaining what's tested and why
6. **Run:** `pytest tests/test_<feature>.py -v`
7. **Commit:** Include test in PR with feature

### Template
```python
@pytest.mark.regression
def test_ec_xxx_descriptive_name(fixture_name):
    """Test description matching edge case inventory."""
    # Arrange
    # Act
    # Assert
    assert expected == actual
```

---

## References

- **Test Coverage Assessment:** `docs/TEST_COVERAGE_ASSESSMENT.md`
- **Edge Case Inventory:** `docs/REGRESSION_TEST_EDGE_CASE_INVENTORY.md`
- **Campaign Pagination Design:** `docs/CAMPAIGN_PAGINATION_DESIGN.md`
- **Reconciliation Validator Design:** `docs/RECONCILIATION_VALIDATOR_DESIGN.md`
- **Session 28-10-25 Context:** `docs/28-10-25/BRAIN_RESET_281025.md`

---

**Test Suite Status:** ✅ READY FOR EXECUTION
**Last Verified:** 28-10-25 (syntax check, fixture references)
**Next Step:** Run full suite and fix any import/fixture issues
