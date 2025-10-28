# Test Coverage Assessment & Regression Test Plan

**Last Updated:** 28-10-25
**Status:** ASSESSMENT COMPLETE (Ready for test restoration)
**Test Framework:** pytest 8.x with PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
**Coverage Target:** Critical paths + regression vectors from 27-10-25 work

---

## Executive Summary

Based on 27-10-25 session completions, we need regression test coverage for:

1. **Campaign Cell Merging** - Verify horizontal merge correctness (CAMPAIGN_NAME only)
2. **Font Normalization** - Verify 6pt/7pt rules applied consistently
3. **Row Height & Continuation Handling** - Verify carried-forward rows and continuation slides
4. **Structural Validation** - Verify last-slide-only shapes (BRAND TOTAL, indicators)
5. **Data Reconciliation** - Verify market/brand normalization and case-insensitive matching

**Critical Gap:** Test suites currently missing from repo (archived in commit 951bb14). Restoration required before Phase 2B.

---

## What Was Fixed in 27-10-25 (Must Prevent Regression)

### 1. Campaign Cell Text Wrapping
**Commit:** 395025b - "fix: resolve campaign text wrapping with hyphens removal"
**What Changed:**
- Removed hyphens from campaign names at source (assembly.py data preprocessing)
- Widened campaign name column to 1,000,000 EMU (maximum)
- Disabled word_wrap in cell_merges.py:612

**Regression Risk:** HIGH
- If word_wrap is re-enabled, campaign names break mid-word: "FACES-CONDITION" → "FACES-\nCONDITIO\nN"
- If column width reverts, overflow text truncates
- If hyphens reintroduced, smart line breaking doesn't trigger

**Test Coverage Needed:**
```python
def test_campaign_name_wrapping():
    """Verify campaign names don't wrap mid-word, and smart line breaks work."""
    # Generate test deck with known campaign names containing hyphens
    # Verify: word_wrap=False in cell styles
    # Verify: campaign names use \n only (not PowerPoint auto-wrap)
    # Verify: Campaign name column width >= 1,000,000 EMU
```

---

### 2. Font Size Corrections
**Commit:** ace42e4 (earlier work) + 27-10-25 consistency pass
**What Changed:**
- Body rows: 6pt (assembly.py cell styling)
- Campaign column: 6pt (assembly.py:672)
- Header row: 7pt
- BRAND TOTAL row: 7pt
- MONTHLY TOTAL: 6pt
- GRAND TOTAL: 7pt
- Bottom utility rows: 6pt

**Regression Risk:** CRITICAL
- Font size changes hard to spot visually (5pt vs 6pt barely distinguishable)
- Font normalization missing from some rows = inconsistent deck appearance
- Template V4 requires strict adherence (client-facing quality metric)

**Test Coverage Needed:**
```python
def test_font_sizes_body_rows():
    """Verify 6pt applied to body/campaign/bottom rows."""
    # Generate test deck
    # Iterate all body rows (rows 1 to N-2, excluding header/totals)
    # Verify: font.size == Pt(6)

def test_font_sizes_header_and_totals():
    """Verify 7pt applied to header and total rows."""
    # Verify header (row 0): font.size == Pt(7)
    # Verify GRAND TOTAL (last row): font.size == Pt(7)
    # Verify BRAND TOTAL on continuation slides: font.size == Pt(7)

def test_font_sizes_monthly_total():
    """Verify MONTHLY TOTAL row is 6pt."""
    # Find "MONTHLY TOTAL (￡ 000)" rows
    # Verify: font.size == Pt(6)
```

---

### 3. Structural Validator - Last-Slide-Only Shapes
**Commit:** 6e83fae - "fix: update structural validator for last-slide-only shapes"
**What Changed:**
- Updated contract: QuarterBudget*, MediaShare*, FunnelShare*, FooterNotes now "last_slide_only_shapes"
- Validator only checks indicators/footer on final slides (where BRAND TOTAL appears)
- grand_total_label changed from "GRAND TOTAL" to "BRAND TOTAL"

**Regression Risk:** HIGH
- If validator reverts to old logic, will flag "missing" indicators on continuation slides (false positives)
- If grand_total_label not updated, reconciliation matching fails

**Test Coverage Needed:**
```python
def test_structural_validator_last_slide_only():
    """Verify validator handles last-slide-only shapes correctly."""
    # Generate multi-slide deck
    # Run structural validator
    # Verify: No false positives for indicators on continuation slides
    # Verify: BRAND TOTAL label recognized (not GRAND TOTAL)

def test_structural_validator_final_slide():
    """Verify indicators present on final slide."""
    # Verify: QuarterBudget shapes exist on last slide
    # Verify: MediaShare shapes exist on last slide
    # Verify: FunnelShare shapes exist on last slide
    # Verify: FooterNotes exist on last slide
```

---

### 4. Data Reconciliation - Market/Brand Normalization
**Commit:** e27af1e - "fix: resolve reconciliation data source issue with case-insensitive matching"
**What Changed:**
- Added `_normalize_market_name()` with MARKET_CODE_MAP translation (MOR→MOROCCO, etc.)
- Added `_normalize_brand_name()` for case-insensitive brand matching
- Fixed `_parse_title_tokens()` regex to handle both "(n of m)" and "(n/m)" pagination formats
- Reconciliation now passes 100% (630/630 records)

**Regression Risk:** CRITICAL
- If normalization functions removed, reconciliation failures = data validation alerts
- Excel data source mismatch = undetected; could cause silent data quality issues

**Test Coverage Needed:**
```python
def test_reconciliation_market_normalization():
    """Verify market names normalized correctly."""
    # Test cases: "SOUTH AFRICA" (exact), "south africa" (lowercase), "Morocco" (MOR code)
    # Verify: Normalized to canonical form
    # Verify: MARKET_CODE_MAP applied correctly

def test_reconciliation_brand_normalization():
    """Verify brands normalized case-insensitively."""
    # Test cases: "Fanta", "FANTA", "fAnTa"
    # Verify: All match same canonical brand
    # Verify: Within correct market only

def test_reconciliation_pagination_parsing():
    """Verify title parsing handles both pagination formats."""
    # Test: "MARKET (1 of 3)" → (1, 3)
    # Test: "MARKET (1/3)" → (1, 3)
    # Verify: Both formats recognized
```

---

### 5. Media Channel Merging
**Commit:** 54df939 - "feat: add media channel vertical merging"
**What Changed:**
- Vertical cell merging added for TELEVISION, DIGITAL, OOH, OTHER media headers
- Merge logic in cell_merges.py:merge_media_channels()
- Applied during post-processing (Step 4)

**Regression Risk:** MEDIUM
- If merge logic broken, media headers appear on every row (visual clutter)
- Merge incorrectness affects row count calculations

**Test Coverage Needed:**
```python
def test_media_channel_merging():
    """Verify media headers merged vertically."""
    # Generate test deck with all media types
    # For each media type (TV, DIGITAL, OOH, OTHER):
    #   - Count merged cells in media column
    #   - Verify: All sub-rows belong to single media header
    #   - Verify: No orphaned or partial merges
```

---

### 6. Smart Line Breaking for Campaign Names
**Commit:** 54df939 (integrated into assembly.py)
**What Changed:**
- `_smart_line_break()` function added to handle campaign names with dashes
- Splits on word boundaries, not mid-word
- Respects max character count per line

**Regression Risk:** MEDIUM
- If function removed/broken, campaign names revert to PowerPoint auto-wrap
- Auto-wrap breaks mid-word (visible quality degradation)

**Test Coverage Needed:**
```python
def test_smart_line_breaking_campaign_names():
    """Verify campaign names split on word boundaries, not mid-word."""
    # Test: "FACES-CONDITION" → "FACES\nCONDITION" (2 lines)
    # Test: "LONG-CAMPAIGN-NAME-HERE" → Smart split respecting width
    # Verify: No dashes at line breaks
    # Verify: Text fits within column width
```

---

### 7. Timestamp Generation (Local System Time)
**Commit:** d6f044a - "fix: use local system time for all timestamps instead of UTC"
**What Changed:**
- Timestamps now use Arabian Standard Time (UTC+4)
- Applied in: cli/main.py, utils/logging.py, assembly.py

**Regression Risk:** LOW (observable, not functional)
- If reverts to UTC, timestamps off by 4 hours
- Not data-critical but affects audit trails

**Test Coverage Needed:**
```python
def test_timestamp_local_system_time():
    """Verify timestamps use local system time (AST)."""
    # Generate deck
    # Extract timestamp from deck metadata
    # Verify: Offset matches system timezone (UTC+4)
    # Verify: Time reasonable (within last hour)
```

---

## Test File Organization

### Proposed Structure
```
tests/
├── __init__.py
├── conftest.py                        # Shared fixtures, temp deck generation
├── test_campaign_merging.py          # Campaign cell merge regression tests
├── test_font_normalization.py        # Font size regression tests
├── test_continuation_handling.py     # Row height, carried-forward regression tests
├── test_structural_validation.py     # Structural validator tests (restored)
├── test_reconciliation.py            # Reconciliation validator tests
├── test_assembly_split.py            # Pagination & table splitting tests
├── test_tables.py                    # Core table generation tests (restored)
└── conftest_fixtures.py              # Example test decks, template refs
```

---

## Test Fixtures & Setup

### Required Fixtures (conftest.py)
```python
import pytest
from pathlib import Path
from pptx import Presentation

@pytest.fixture
def template_path():
    """Path to Template_V4_FINAL_071025.pptx"""
    return Path(__file__).parent.parent / "template" / "Template_V4_FINAL_071025.pptx"

@pytest.fixture
def excel_path():
    """Path to BulkPlanData_2025_10_14.xlsx"""
    return Path(__file__).parent.parent / "template" / "BulkPlanData_2025_10_14.xlsx"

@pytest.fixture
def test_deck(template_path, excel_path, tmp_path):
    """Generate fresh test deck before each test."""
    # Use amp_automation.cli.main to generate
    # Return path to generated PPTX in tmp_path
    pass

@pytest.fixture
def sample_campaign_names():
    """Sample campaign names for text wrapping tests."""
    return [
        "FACES-CONDITION",
        "LONG-CAMPAIGN-NAME",
        "SHORT",
        "MULTI-WORD-CAMPAIGN-NAME-WITH-HYPHENS",
    ]
```

### Test Execution Environment
```bash
# Before running tests
Stop-Process -Name POWERPNT -Force  # Close any PowerPoint instances

# Run test suite
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
python -m pytest tests/ -v --tb=short -k "not slow"

# Run with coverage
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD=1
python -m pytest tests/ --cov=amp_automation --cov-report=html
```

---

## Regression Test Categories

### Category 1: Cell Formatting (test_campaign_merging.py)
- Campaign cell merge correctness
- Media channel vertical merge correctness
- Row height consistency
- Cell alignment (centered)

### Category 2: Font Normalization (test_font_normalization.py)
- Body row fonts (6pt)
- Header/total fonts (7pt)
- Font consistency across continuation slides
- Font consistency across media channels

### Category 3: Table Splitting & Continuation (test_continuation_handling.py)
- Pagination respects max_rows_per_slide
- Campaign boundaries not crossed
- Carried-forward row accumulation
- Slide-level GRAND TOTAL correctness
- Continuation indicator in titles

### Category 4: Structural Validation (test_structural_validation.py)
- Last-slide-only shape detection
- BRAND TOTAL vs GRAND TOTAL distinction
- Indicator presence on final slides only
- Footer shape validation

### Category 5: Data Reconciliation (test_reconciliation.py)
- Market name normalization
- Brand name normalization
- Pagination format parsing (both formats)
- Data accuracy validation
- Data completeness validation

### Category 6: Assembly & Splitting (test_assembly_split.py)
- Smart line breaking on campaign names
- Campaign boundary detection
- Media block length calculation
- Table data chunking logic

---

## Metrics & Success Criteria

### Coverage Targets
- **Line Coverage:** ≥70% of core modules (assembly.py, cell_merges.py, reconciliation.py)
- **Function Coverage:** 100% of public functions tested
- **Branch Coverage:** ≥80% (decision points in pagination, splitting)

### Test Execution Time
- **Unit tests:** <5 seconds (no deck generation)
- **Integration tests:** <60 seconds (1-2 deck generations)
- **Full suite:** <120 seconds

### Failure Criteria
- Any regression test FAIL = blocker (must fix before pushing)
- Coverage drop >5% = warning (investigate)
- Execution time >150% baseline = performance regression (investigate)

---

## Edge Cases Requiring Tests

### Edge Case 1: Single-Slide Markets
**Scenario:** Market with <32 body rows
**Expectation:** Single slide, no continuation indicator, no carried-forward row
**Test:**
```python
def test_single_slide_market():
    # Verify: is_split == False
    # Verify: No " (Continued)" in title
    # Verify: No "CARRIED FORWARD" row
```

### Edge Case 2: Multi-Slide Markets
**Scenario:** Market with >32 body rows requiring continuation
**Expectation:** Multiple slides, all except last have " (Continued)", all have carried-forward
**Test:**
```python
def test_multi_slide_market():
    # Verify: is_split == True
    # Verify: All slides except last have " (Continued)"
    # Verify: All continuation slides have "CARRIED FORWARD" row
    # Verify: Only final slide has slide-level "GRAND TOTAL"
```

### Edge Case 3: Campaign Boundary Preservation
**Scenario:** Campaign that exceeds max_rows when combined with previous
**Expectation:** Campaign moved to next slide (not split mid-campaign)
**Test:**
```python
def test_campaign_boundary_not_crossed():
    # Create scenario where campaign_size + current_body_count > MAX_ROWS
    # Verify: Campaign moved to new slide, not split
```

### Edge Case 4: Media Channel Grouping
**Scenario:** Multiple campaigns within single media type
**Expectation:** Media header + all sub-campaigns + metrics stay together on same slide
**Test:**
```python
def test_media_channel_grouping():
    # Verify: TELEVISION header + all TV campaigns on same slide
    # Verify: Not split across continuation boundaries
```

### Edge Case 5: Empty Metrics
**Scenario:** Campaign with no monthly data (all dashes)
**Expectation:** Still included in table, CARRIED FORWARD excludes from accumulation
**Test:**
```python
def test_empty_metrics_handling():
    # Verify: Campaign row present
    # Verify: Metrics show dashes (not NULL)
    # Verify: Not counted in carried-forward total
```

---

## Blocked Test Scenarios (Phase 4+)

These require additional work before testing:

- **Row Height Normalization:** Auto-adjust based on text length (Phase 4)
- **Cell Margin/Padding:** Expand if needed for footer readability (Phase 4)
- **Visual Diff Automation:** Establish Slide 1 baseline, compare EMU/legend (Phase 4)
- **Smart Pagination:** Enable row-level splitting if Phase 3 data requires (Phase 4)

---

## Defect Tracking & Regression Prevention

### How to Report Regression
1. Run full test suite: `pytest tests/ -v`
2. If test FAILS, identify: Which module? Which commit introduced it?
3. File issue with: test output + git blame result + proposed fix
4. Don't commit until test passes

### How to Prevent Regression
1. **Before coding:** Create failing test for desired behavior
2. **After coding:** Make test pass
3. **Before pushing:** Full test suite green
4. **After merging:** Monitor for new failures on downstream branches

### Test-Driven Development Workflow
```
1. Identify bug or missing feature
2. Write test that demonstrates issue (RED)
3. Implement minimal fix (GREEN)
4. Refactor for clarity (REFACTOR)
5. Commit with test coverage
6. Push to main
```

---

## References

- **Test Framework:** pytest 8.x (use `pytest.ini` config)
- **Python:** 3.10+ (from __future__ annotations)
- **Modules Under Test:**
  - amp_automation/presentation/assembly.py (pagination, formatting)
  - amp_automation/presentation/postprocess/cell_merges.py (merging)
  - amp_automation/validation/reconciliation.py (data matching)
  - amp_automation/validation/ (all validators)
- **Fixtures:** See conftest.py (to be created during restoration)

---

**Assessment Status:** ✅ COMPLETE
**Ready for Test Restoration:** YES
**Test Framework:** pytest 8.x
**Estimated Restoration Time:** 4-6 hours (including updates for 27-10-25 changes)
