# Regression Test Edge Case Inventory

**Last Updated:** 28-10-25
**Status:** DOCUMENTED (Identified from 27-10-25 fixes)
**Scope:** Campaign pagination, font normalization, reconciliation, structural validation

---

## Overview

This document catalogs edge cases discovered during 27-10-25 session that require regression test coverage. Each case identifies:
- **Risk:** Why it matters (HIGH/MEDIUM/LOW)
- **Root Cause:** What breaks if not tested
- **Test Strategy:** How to detect regression
- **Prevention:** Configuration or guardrail

---

## CRITICAL EDGE CASES (HIGH RISK)

### EC-001: Campaign Text Wrapping Mid-Word
**Risk:** HIGH - Visual quality degradation, client-visible bug
**Triggered By:** Commit 395025b - word wrap fix
**Root Cause:** PowerPoint auto-wrap breaks mid-word when word_wrap=True and column width insufficient

**Scenario:**
- Campaign name: "FACES-CONDITION" (21 characters)
- Column width: < 1,000,000 EMU
- word_wrap property: true
- Expected: "FACES\nCONDITION" (smart line break)
- Bug Output: "FACES-CONDITIO\nN" or "FACES-CONDI\nTION"

**Test:**
```python
def test_ec001_campaign_text_wrap_mid_word():
    """Verify campaign names don't split mid-word."""
    deck = generate_test_deck(campaigns=["FACES-CONDITION", "LONG-CAMPAIGN-NAME"])

    for slide in deck.slides:
        table = find_main_table(slide)
        if table:
            for row_idx in range(1, len(table.rows)):
                cell = table.cell(row_idx, 0)  # Campaign name column

                # Check word wrap disabled
                text_frame = cell.text_frame
                assert not text_frame.word_wrap, f"Row {row_idx}: word_wrap should be False"

                # Check no mid-word breaks (hyphens not at line ends)
                for para in text_frame.paragraphs:
                    for run in para.runs:
                        assert not run.text.endswith("-"), \
                            f"Row {row_idx}: Campaign name ends with hyphen (mid-word break)"

                # Check column width sufficient
                assert table.columns[0].width >= 1000000, \
                    f"Campaign column too narrow: {table.columns[0].width} EMU"
```

**Prevention:**
- Always set `text_frame.word_wrap = False` in cell_merges.py
- Keep campaign column width ≥ 1,000,000 EMU (assembly.py:1222)
- Test with campaigns containing hyphens (FACES-CONDITION)

---

### EC-002: Font Size Drift Across Slides
**Risk:** HIGH - Subtle but visible, hard to detect manually
**Triggered By:** Commits ace42e4 (font fix) + 27-10-25 consistency
**Root Cause:** Font normalization skipped on some rows due to conditional logic bugs

**Scenario:**
- Slide 1 body rows: 6pt ✓
- Slide 2 body rows: 6pt ✓
- Slide 3 body rows: 5pt or 8pt ✗ (accidental revert)
- Continuation slide campaign rows: 6pt ✓
- Continuation slide body rows: 7pt ✗ (wrong size)

**Test:**
```python
def test_ec002_font_size_consistency_across_slides():
    """Verify font sizes consistent across all slides."""
    deck = generate_test_deck(num_slides=5)  # Force multi-slide deck

    font_size_map = {
        "header": Pt(7),           # Row 0
        "body": Pt(6),            # Rows 1 to N-2
        "monthly_total": Pt(6),   # "MONTHLY TOTAL"
        "brand_total": Pt(7),     # Continuation: BRAND TOTAL
        "campaign": Pt(6),        # Campaign name rows
    }

    for slide_idx, slide in enumerate(deck.slides):
        table = find_main_table(slide)
        if not table:
            continue

        for row_idx, row in enumerate(table.rows):
            # Identify row type
            row_label = str(table.cell(row_idx, 0).text).strip().upper()

            if row_idx == 0:
                expected_size = font_size_map["header"]
                row_type = "header"
            elif "MONTHLY TOTAL" in row_label:
                expected_size = font_size_map["monthly_total"]
                row_type = "monthly_total"
            elif "BRAND TOTAL" in row_label:
                expected_size = font_size_map["brand_total"]
                row_type = "brand_total"
            elif row_idx == len(table.rows) - 1:
                expected_size = font_size_map["header"]  # Last row is GRAND TOTAL
                row_type = "grand_total"
            else:
                expected_size = font_size_map["body"]
                row_type = "body"

            # Check all cells in row
            for col_idx in range(len(table.columns)):
                cell = table.cell(row_idx, col_idx)
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        actual_size = run.font.size
                        assert actual_size == expected_size, \
                            f"Slide {slide_idx}, Row {row_idx} ({row_type}), Col {col_idx}: " \
                            f"Expected {expected_size}, got {actual_size}"
```

**Prevention:**
- Always apply font sizes via post-processing (table_normalizer.py)
- Test full suite of decks (single-slide, multi-slide, multi-slide with continuation)
- Verify font consistency before releasing deck

---

### EC-003: Campaign Boundary Violated During Pagination
**Risk:** HIGH - Data integrity, incomplete campaign rows
**Triggered By:** Pagination logic (_split_table_data_by_campaigns)
**Root Cause:** Row block calculation doesn't respect campaign boundaries

**Scenario:**
- Campaign 1: 15 rows (media blocks + metrics)
- Campaign 2: 20 rows
- MAX_ROWS_PER_SLIDE = 32
- Slide 1 fits: Campaign 1 (15 rows) + Campaign 2 partial (17 rows) = 32
- Bug Output: Campaign 2 split across Slide 1 & Slide 2 ✗
- Expected: Campaign 2 moved entirely to Slide 2

**Test:**
```python
def test_ec003_campaign_boundary_not_violated():
    """Verify campaigns not split across continuation slides."""
    # Generate deck with markets containing campaigns > 1 slide worth
    deck = generate_test_deck(
        market_configs=[
            {"market": "TestMarket", "brands": ["TestBrand"]},
            # Create data forcing multi-slide scenario
        ]
    )

    for slide_idx, slide in enumerate(deck.slides):
        table = find_main_table(slide)
        if not table:
            continue

        # Track campaigns in this slide
        campaigns_seen = set()
        for row_idx in range(1, len(table.rows) - 1):  # Skip header & GRAND TOTAL
            cell = table.cell(row_idx, 0)
            row_label = str(cell.text).strip()

            # Campaign rows have names (not empty, not media headers, not metrics)
            if row_label and row_label != "-" and not is_media_header(row_label):
                campaigns_seen.add(row_label)

        # Next slide: verify different campaigns or no split
        if slide_idx < len(deck.slides) - 1:
            next_slide = deck.slides[slide_idx + 1]
            next_table = find_main_table(next_slide)
            if next_table:
                campaigns_next = set()
                for row_idx in range(1, len(next_table.rows) - 1):
                    cell = next_table.cell(row_idx, 0)
                    row_label = str(cell.text).strip()
                    if row_label and row_label != "-":
                        campaigns_next.add(row_label)

                # Verify no campaign overlap (no split)
                overlap = campaigns_seen & campaigns_next
                assert not overlap, f"Campaign split across slides: {overlap}"
```

**Prevention:**
- Test with decks containing markets requiring multiple slides
- Verify campaign boundaries identified correctly (assembly.py:2322)
- Always test multi-slide pagination scenarios

---

### EC-004: Reconciliation False Negatives (Unmatched Data)
**Risk:** HIGH - Silent data quality failures
**Triggered By:** Commit e27af1e - normalization fix
**Root Cause:** Case sensitivity or code mapping not applied

**Scenario:**
- Excel data: "south africa" (lowercase), "MOR" (market code)
- PPT deck: "SOUTH AFRICA" (uppercase), "MOROCCO" (display name)
- Bug: No match found, reconciliation FAILS for entire market
- Before fix: 630/631 checks failing
- After fix: 630/630 checks passing

**Test:**
```python
def test_ec004_reconciliation_handles_case_mismatch():
    """Verify reconciliation normalizes case differences."""
    # Create Excel data with mixed case
    df = pd.DataFrame([
        {"Country": "south africa", "Brand": "fanta", "Year": 2025, ...},
        {"Country": "SOUTH AFRICA", "Brand": "FANTA", "Year": 2025, ...},
        {"Country": "South Africa", "Brand": "Fanta", "Year": 2025, ...},
    ])

    # Create deck with uppercase
    deck = generate_test_deck(markets=["SOUTH AFRICA"], brands=["FANTA"])

    # Reconcile
    results = generate_reconciliation_report(deck_path, df=df)

    # Verify matches found
    assert len(results) > 0, "No reconciliation results (data not found)"
    assert all(r.passed for r in results), "Reconciliation failures due to case mismatch"

def test_ec005_reconciliation_handles_market_code():
    """Verify market code mapping (MOR → MOROCCO)."""
    df = pd.DataFrame([
        {"Country": "MOR", "Brand": "Sprite", "Year": 2025, ...},
    ])

    deck = generate_test_deck(markets=["MOROCCO"], brands=["Sprite"])
    results = generate_reconciliation_report(deck_path, df=df)

    # Verify match found despite code/display name difference
    assert any(r.market == "MOROCCO" for r in results), "Market code not translated"
```

**Prevention:**
- Always test with case-mismatched data (create test Excel with mixed case)
- Test market code translations (add "MOR" to Excel, expect MOROCCO in deck)
- Verify reconciliation pass rate = 100% before releasing

---

## MEDIUM RISK EDGE CASES

### EC-006: Empty Campaign Metrics
**Risk:** MEDIUM - Affects carried-forward calculations
**Triggered By:** Reconciliation accumulation logic
**Root Cause:** Empty metrics (dashes) included in accumulation

**Scenario:**
- Campaign with no monthly data (all "-" dashes)
- Carried-forward row sums campaigns
- Bug: Empty campaign adds 0 to total (correct)
- But: Row still included in carried-forward, suggests data present when absent

**Test:**
```python
def test_ec006_empty_campaign_metrics():
    """Verify campaigns with no data handled correctly."""
    # Campaign with empty metrics (all dashes)
    data = [
        {"campaign": "NoData", "jan": "-", "feb": "-", ...},
        {"campaign": "HasData", "jan": "1000", "feb": "1000", ...},
    ]

    # Generate deck with this data
    deck = generate_test_deck_from_data(data)

    # Verify:
    # 1. NoData campaign row present
    # 2. All metrics show dashes
    # 3. Carried-forward excludes NoData (only HasData counted)

    for slide in deck.slides:
        table = find_main_table(slide)
        # ... verify NoData row present with all dashes
        # ... verify carried-forward sums only HasData
```

**Prevention:**
- Test with campaigns containing no data
- Verify carried-forward calculations skip empty metrics

---

### EC-007: Multi-Year Market Data
**Risk:** MEDIUM - Reconciliation ambiguity
**Triggered By:** Reconciliation year selection logic
**Root Cause:** Multiple years available, wrong year selected

**Scenario:**
- Excel data: "SOUTH AFRICA / Fanta" in years 2024, 2025, 2026
- PPT slide: "SOUTH AFRICA / Fanta" (no year indicator)
- Deck generated with 2025 data
- Bug: Reconciliation picks 2024 or 2026 instead of 2025

**Test:**
```python
def test_ec007_reconciliation_multi_year_selection():
    """Verify correct year selected when multiple available."""
    # Excel data with 3 years
    df = pd.DataFrame([
        {"Country": "SOUTH AFRICA", "Brand": "Fanta", "Year": 2024, "Jan": 100, ...},
        {"Country": "SOUTH AFRICA", "Brand": "Fanta", "Year": 2025, "Jan": 200, ...},
        {"Country": "SOUTH AFRICA", "Brand": "Fanta", "Year": 2026, "Jan": 300, ...},
    ])

    # Deck with 2025 data (summary tiles = 2025 values)
    deck = generate_test_deck_with_data(df.query("Year == 2025"))

    results = generate_reconciliation_report(deck_path, df=df)

    # Verify correct year selected
    for r in results:
        assert r.year == 2025, f"Wrong year selected: {r.year}"
        assert r.passed, f"Reconciliation failed for selected year"
```

**Prevention:**
- Test with multi-year data
- Verify year selection score metric (most passes, smallest diff)

---

### EC-008: Media Channel Merge Incompleteness
**Risk:** MEDIUM - Visual inconsistency
**Triggered By:** Commit 54df939 - media merging
**Root Cause:** Not all media headers merged (partial merge)

**Scenario:**
- TELEVISION header: merged (rows 2-4)
- DIGITAL header: NOT merged (appears on every row)
- OOH header: merged (rows 8-10)
- Result: Inconsistent appearance

**Test:**
```python
def test_ec008_media_channel_merging_complete():
    """Verify all media headers merged vertically."""
    deck = generate_test_deck()

    for slide in deck.slides:
        table = find_main_table(slide)
        if not table:
            continue

        # Track media columns (usually column 1)
        media_col = 1
        current_media = None
        merge_start = None

        for row_idx, row in enumerate(table.rows):
            media_cell = table.cell(row_idx, media_col)
            media_label = str(media_cell.text).strip()

            if media_label in ["TELEVISION", "DIGITAL", "OOH", "OTHER"]:
                if current_media and current_media != media_label:
                    # Previous media ended, verify merge
                    assert _is_merged(table, media_col, merge_start, row_idx - 1), \
                        f"Media '{current_media}' not fully merged"

                current_media = media_label
                merge_start = row_idx

        # Check final media type
        if current_media:
            assert _is_merged(table, media_col, merge_start, len(table.rows) - 2), \
                f"Media '{current_media}' not fully merged"
```

**Prevention:**
- Test with all media types present
- Verify merge spans all sub-rows

---

## LOW RISK EDGE CASES

### EC-009: Timestamp Timezone Offset
**Risk:** LOW - Observable but not data-critical
**Triggered By:** Commit d6f044a - local time fix
**Root Cause:** Timestamp uses UTC instead of AST (+4)

**Scenario:**
- System timezone: UTC+4 (Arabia Standard Time)
- Timestamp in deck metadata: 2025-10-28 10:00:00 (UTC)
- Expected: 2025-10-28 14:00:00 (AST)
- Discrepancy: 4-hour offset

**Test:**
```python
def test_ec009_timestamp_local_system_time():
    """Verify timestamps use local system time (AST)."""
    import datetime
    import zoneinfo

    deck = generate_test_deck()

    # Extract timestamp from deck (usually in metadata or subtitle)
    deck_time = extract_timestamp_from_deck(deck)
    current_time = datetime.datetime.now()
    current_tz = zoneinfo.ZoneInfo("Asia/Dubai")  # AST
    current_ast = current_time.astimezone(current_tz)

    # Verify offset matches
    offset_expected = current_ast.utcoffset()
    offset_actual = deck_time.utcoffset()

    assert offset_actual == offset_expected, \
        f"Timestamp offset mismatch: expected {offset_expected}, got {offset_actual}"
```

**Prevention:**
- Verify timestamps use local system time setting
- Test on system with UTC+4 timezone set

---

### EC-010: Pagination Marker Format Variations
**Risk:** LOW - Edge case in title parsing
**Triggered By:** Commit e27af1e - pagination parsing fix
**Root Cause:** Only "(n of m)" format supported, not "(n/m)"

**Scenario:**
- Standard format: "SOUTH AFRICA / Fanta (1 of 3)"
- Alternative format: "SOUTH AFRICA / Fanta (1/3)"
- Bug: Alternative format not recognized, market/brand not parsed

**Test:**
```python
def test_ec010_pagination_format_variations():
    """Verify both pagination formats recognized."""
    titles = [
        "SOUTH AFRICA / Fanta (1 of 3)",
        "SOUTH AFRICA / Fanta (1/3)",
        "SOUTH AFRICA / Fanta (continued)",  # Bonus: no number
    ]

    for title in titles:
        market, brand = _parse_title_tokens(title)
        assert market == "SOUTH AFRICA", f"Market not parsed from '{title}'"
        assert brand == "Fanta", f"Brand not parsed from '{title}'"
```

**Prevention:**
- Test with both pagination formats
- Add new formats as discovered (update regex in reconciliation.py:503)

---

## Test Matrix

| Edge Case | Single-Slide | Multi-Slide | No Data | Multi-Year | High Volume |
|-----------|--------------|-------------|---------|-----------|-------------|
| EC-001 | ✓ | ✓ |  | | ✓ |
| EC-002 | ✓ | ✓ | | | ✓ |
| EC-003 | | ✓ | | | ✓ |
| EC-004 | ✓ | ✓ | ✓ | ✓ | |
| EC-005 | ✓ | ✓ | | ✓ | |
| EC-006 | ✓ | | ✓ | | |
| EC-007 | | ✓ | | ✓ | |
| EC-008 | ✓ | ✓ | | | |
| EC-009 | ✓ | ✓ | | | |
| EC-010 | ✓ | ✓ | | | |

---

## Test Data Requirements

### Minimal Test Dataset
```python
# Small dataset for unit tests
min_data = {
    "SOUTH AFRICA": ["Fanta", "Sprite"],  # 2 brands
    "EGYPT": ["Fanta"],                    # 1 brand
    # Single year: 2025
    # Single funnel stage: Awareness
    # Single media type: TV
}
```

### Comprehensive Test Dataset
```python
# Full dataset for integration tests
full_data = {
    "SOUTH AFRICA": ["Fanta", "Sprite", "Coca-Cola"],
    "EGYPT": ["Fanta"],
    "MOROCCO": ["Fanta", "Sprite"],
    "KENYA": ["Fanta"],
    # Multiple years: 2024, 2025, 2026
    # Multiple funnel stages: Awareness, Consideration, Preference, Purchase
    # Multiple media types: TV, Digital, OOH, Other
    # Large campaigns: >32 rows per slide (force pagination)
    # Mixed case brands: "fanta", "FANTA", "Fanta"
    # Market codes: "MOR", "KSA", "south africa" (lowercase)
    # Empty metrics: Some campaigns with no data
}
```

---

## Regression Prevention Checklist

Before each release:

- [ ] EC-001: Campaign names with hyphens tested (word wrap off)
- [ ] EC-002: Font sizes verified across all slides (6pt/7pt)
- [ ] EC-003: Multi-slide markets verified (no campaign splits)
- [ ] EC-004: Case mismatch reconciliation tested
- [ ] EC-005: Market code translation tested (MOR→MOROCCO)
- [ ] EC-006: Empty metrics campaigns tested
- [ ] EC-007: Multi-year reconciliation tested
- [ ] EC-008: All media channel merges verified
- [ ] EC-009: Timestamp timezone verified
- [ ] EC-010: Both pagination formats tested

---

**Document Status:** ✅ COMPLETE
**Ready for Testing:** YES
**Test Count:** 10 critical + medium cases, 3 categories
**Estimated Coverage:** 80%+ of regression vectors
