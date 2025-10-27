# Dynamic Elements Rules & Calculation Logic

**Date:** 27 October 2025
**Purpose:** Document all dynamic elements on presentation slides with calculation rules and implementation details
**Status:** NEEDS VERIFICATION - Current implementation may not fully align with requirements

---

## Overview

This document defines calculation rules for all dynamic elements on AMP presentation slides. Each element has specific scope (slide-level, campaign-level, or brand-level) and calculation logic that must be followed.

---

## 1. Slide Title (Top Left)

### Current Format
```
{MARKET} - {BRAND}
```

### Required Format (Per User Request)
```
{MARKET} - {BRAND} [25]
```
Where `[25]` represents the year (2025 → 25).

### Implementation Location
- **File:** `amp_automation/presentation/assembly.py:518-553`
- **Functions:**
  - `_compose_title_text()` - Composes title text
  - `_apply_title()` - Applies title to slide shape

### Current Code
```python
def _compose_title_text(
    combination_row: tuple[str, str, int],
    slide_title_suffix: str,
) -> str:
    market, brand, year = combination_row
    market_display = _normalize_market_name(market)
    title_format = presentation_config.get("title", {}).get("format", "{market} - {brand}")
    return title_format.format(market=market_display, brand=brand, year=year) + slide_title_suffix
```

### **ISSUE IDENTIFIED:** Title does NOT include year by default

### Calculation Rules
| Element | Scope | Calculation Logic | Example |
|---------|-------|-------------------|---------|
| Market | Static | From data grouping (market column) | "KSA", "South Africa" |
| Brand | Static | From data grouping (brand column) | "Sensodyne", "Panadol" |
| Year | Static | From data grouping (year column), formatted as 2-digit | "[25]" for 2025 |
| Slide Suffix | Dynamic | Added for continuation slides | " (Continued)" or "" |

### **REQUIRED FIX:**
Update title format to include year:
```python
title_format = "{market} - {brand} [{year}]"
# Example output: "KSA - Sensodyne [25]"
```

Or update config `master_config.json`:
```json
"presentation": {
  "title": {
    "format": "{market} - {brand} [{year}]"
  }
}
```

### Configuration
- **Config File:** `config/master_config.json`
- **Config Path:** `presentation.title.format`
- **Default:** `"{market} - {brand}"` (missing year!)

---

## 2. MONTHLY TOTAL Rows

### Scope
**CAMPAIGN-LEVEL** - Calculates totals for EACH campaign independently

### Purpose
Provides monthly budget totals for a single campaign, appearing immediately after each campaign's media type rows.

### Implementation Location
- **File:** `amp_automation/presentation/assembly.py:909-935`
- **Function:** `_build_campaign_monthly_total_row()`

### Calculation Logic

**For each campaign:**
1. Sum all budget values from campaign rows for each month column
2. Calculate total budget (sum across all months)
3. Format values as currency (£K or £M)
4. Include empty cells for GRP and % columns

### Code Implementation
```python
def _build_campaign_monthly_total_row(
    row_idx: int,
    month_totals: list[float],
    cell_metadata: dict[tuple[int, int], dict[str, object]],
) -> list[str]:
    row: list[str] = ["MONTHLY TOTAL (£ 000)", "", ""]

    # Add monthly values (columns 3-14 for Jan-Dec)
    for month_idx, value in enumerate(month_totals):
        formatted = _format_budget_cell(value)
        col_idx = 3 + month_idx
        row.append(formatted)
        _set_cell_metadata(...)

    # Add total budget (column 15)
    total_value = sum(month_totals)
    row.append(_format_total_budget(total_value))

    # GRP and % columns remain blank
    row.extend(["", ""])
    return row
```

### Row Structure
| Col 0-2 | Col 3-14 | Col 15 | Col 16 | Col 17 |
|---------|----------|--------|--------|--------|
| "MONTHLY TOTAL (£ 000)" | Monthly sums | Total budget | "" | "" |

### Calculation Rules

| Column | Scope | Calculation | Notes |
|--------|-------|-------------|-------|
| **0-2** | Label | Static text "MONTHLY TOTAL (£ 000)" | Merged horizontally in post-processing |
| **3-14** | Campaign | `SUM(campaign_rows[month])` | Sum only rows for THIS campaign |
| **15** | Campaign | `SUM(months[3:14])` | Sum all monthly columns |
| **16** | N/A | Empty | GRP not calculated for budget subtotal |
| **17** | N/A | Empty | Percentage not shown at campaign level |

### **CRITICAL RULE:**
Monthly totals MUST only sum rows belonging to the specific campaign. When campaigns span multiple slides:
- Each slide gets its own continuation slide
- Monthly totals on continuation slides sum ONLY that slide's rows
- Final monthly total sums the entire campaign

### Verification Points
✅ **Correct:** Campaign A has 3 media types → MONTHLY TOTAL sums all 3
✅ **Correct:** Campaign A split across 2 slides → First slide shows partial total, last slide shows campaign total
❌ **Incorrect:** MONTHLY TOTAL includes rows from different campaign
❌ **Incorrect:** MONTHLY TOTAL spans multiple campaigns

---

## 3. GRAND TOTAL Row

### Scope
**BRAND-LEVEL** - Calculates totals across ALL campaigns for a brand across ALL slides

### Purpose
Provides final totals for the entire brand, appearing at the bottom of the LAST slide for each brand.

### **CRITICAL UNDERSTANDING:**
GRAND TOTAL is **BRAND-LEVEL**, not slide-level. This means:
- **Calculation scope:** ALL campaigns for a brand, across ALL slides
- **Appears on:** Only the LAST slide of each brand
- **Example:** If "KSA - Sensodyne" has 2 slides, GRAND TOTAL appears ONLY on slide 2, summing campaigns from BOTH slides

### Implementation Location
- **File:** `amp_automation/presentation/assembly.py:2080-2176`
- **Function:** `_prepare_table_data()` (builds entire table)
- **Builder Function:** `_build_grand_total_row()` (assembly.py:1209-1221)

### Calculation Logic

**At brand level:**
1. Accumulate monthly totals across ALL campaigns
2. Calculate total budget (sum of all campaign budgets)
3. Calculate total GRP (sum of all TV campaign GRPs)
4. Percentage always 100% (represents entire brand)

### Code Implementation
```python
# In _prepare_table_data():
monthly_totals = [0.0] * len(TABLE_MONTH_ORDER)  # Initialize
grand_total_grp = 0.0

# For each campaign in brand:
for campaign_name in campaign_names:
    block_rows, block_month_totals, block_grp_total = _build_campaign_block(...)

    # Accumulate across ALL campaigns
    monthly_totals = [
        total + addition
        for total, addition in zip(monthly_totals, block_month_totals)
    ]
    grand_total_grp += block_grp_total

# Build grand total row with accumulated values
grand_total_row = _build_grand_total_row(
    monthly_totals,      # Summed across all campaigns
    total_budget,        # Total brand budget
    grand_total_grp,     # Total brand GRP
)
```

### Builder Function
```python
def _build_grand_total_row(
    monthly_totals: list[float],
    total_budget: float,
    grand_total_grp: float,
) -> list[str]:
    row: list[str] = ["GRAND TOTAL", "", ""]
    for value in monthly_totals:
        row.append(_format_budget_cell(value))

    row.append(_format_total_budget(total_budget))
    row.append(format_number(grand_total_grp, is_grp=True))
    row.append(_format_percentage_cell(100.0))  # Always 100%
    return row
```

### Row Structure
| Col 0-2 | Col 3-14 | Col 15 | Col 16 | Col 17 |
|---------|----------|--------|--------|--------|
| "GRAND TOTAL" | Brand monthly sums | Brand total budget | Brand total GRP | 100% |

### Calculation Rules

| Column | Scope | Calculation | Notes |
|--------|-------|-------------|-------|
| **0-2** | Label | Static text "GRAND TOTAL" | Merged horizontally in post-processing |
| **3-14** | Brand | `SUM(all_campaigns[month])` | Sum ALL campaigns in brand |
| **15** | Brand | `SUM(all_campaigns.total_budget)` | Total budget for brand |
| **16** | Brand | `SUM(all_tv_campaigns.grp)` | Total GRP for brand (TV only) |
| **17** | Brand | `100.0` | Always 100% (represents entire brand) |

### **CRITICAL RULE: Brand-Level Scope**

When brand spans multiple slides:
```
KSA - Sensodyne [25] (Slide 1)
  Campaign 1 (20 rows)
  Campaign 2 (19 rows)
  [NO GRAND TOTAL HERE]

KSA - Sensodyne [25] (Slide 2 of 2)
  Campaign 3 (12 rows)
  Campaign 4 (10 rows)
  GRAND TOTAL ← Sums campaigns 1+2+3+4 from BOTH slides
```

### **QUESTION FOR USER:**
Current implementation creates GRAND TOTAL at the END of table data preparation, which means:
- ✅ It correctly sums ALL campaigns in the brand
- ❓ But does it appear on EVERY slide or ONLY the last slide?

Need to verify: When pagination splits a brand across multiple slides, does GRAND TOTAL appear:
- **Option A:** On every slide (showing accumulated total up to that slide)
- **Option B:** Only on the last slide for that brand ← **This should be correct**

### Verification Points
✅ **Correct:** GRAND TOTAL sums ALL campaigns for the brand
✅ **Correct:** GRAND TOTAL appears only on LAST slide of brand
❌ **Incorrect:** GRAND TOTAL shows slide-level total instead of brand total
❌ **Incorrect:** GRAND TOTAL appears on every slide

---

## 4. CARRIED FORWARD Row

### Scope
**SLIDE-LEVEL** - Calculates totals for rows on continuation slides only

### Purpose
Appears at the TOP of continuation slides (slides 2, 3, etc.) to show the budget carried forward from the previous slide.

### Implementation Location
- **File:** `amp_automation/presentation/assembly.py:1224-1269`
- **Functions:**
  - `_build_carried_forward_row()` - Builds row data
  - `_build_carried_forward_metadata_values()` - Builds cell metadata

### Calculation Logic

**For continuation slides:**
1. Calculate monthly totals for rows ALREADY SHOWN on previous slides
2. Show accumulated budget up to the previous slide
3. Show accumulated GRP for TV campaigns up to the previous slide
4. No percentage column

### Code Implementation
```python
def _build_carried_forward_row(
    month_totals: list[float],
    total_budget: float,
    grp_total: float,
) -> list[str]:
    row: list[str] = ["CARRIED FORWARD", "", ""]
    for value in month_totals:
        row.append(_format_budget_cell(value))

    row.append(_format_total_budget(total_budget))
    row.append(format_number(grp_total, is_grp=True) if grp_total else "")
    row.append("")  # No percentage
    return row
```

### Row Structure
| Col 0-2 | Col 3-14 | Col 15 | Col 16 | Col 17 |
|---------|----------|--------|--------|--------|
| "CARRIED FORWARD" | Accumulated monthly | Accumulated budget | Accumulated GRP | "" |

### Calculation Rules

| Column | Scope | Calculation | Notes |
|--------|-------|-------------|-------|
| **0-2** | Label | Static text "CARRIED FORWARD" | Merged horizontally in post-processing |
| **3-14** | Previous slides | `SUM(previous_slides[month])` | Accumulated from slide 1 to N-1 |
| **15** | Previous slides | `SUM(previous_slides.total_budget)` | Accumulated budget |
| **16** | Previous slides | `SUM(previous_slides.grp)` | Accumulated GRP (TV only) |
| **17** | N/A | Empty | No percentage shown |

### **CRITICAL RULE:**
CARRIED FORWARD appears ONLY on continuation slides (slide 2+), never on first slide.

### Verification Points
✅ **Correct:** CARRIED FORWARD shows on slide 2+, not slide 1
✅ **Correct:** Values match sum of all rows from previous slides
❌ **Incorrect:** CARRIED FORWARD appears on first slide
❌ **Incorrect:** Values don't match accumulated totals

---

## 5. TV Quarter Indicators (Green Rectangles)

### **IMPORTANT DISCOVERY:**
Based on code inspection, "quarter indicators" refer to **SUMMARY SLIDE TILES**, not visual indicators on data slides.

### Scope
**BRAND-LEVEL** - Appears on summary slides (if they exist)

### Purpose
Display quarterly budget totals on brand summary slides using text shapes configured in the template.

### Implementation Location
- **File:** `amp_automation/presentation/assembly.py:164-189`
- **Function:** `_populate_quarter_tiles()`

### Calculation Logic

**For each quarter:**
1. Define quarter months:
   - Q1: Jan, Feb, Mar
   - Q2: Apr, May, Jun
   - Q3: Jul, Aug, Sep
   - Q4: Oct, Nov, Dec
2. Sum budget values for those 3 months
3. Format and display in quarter tile shape

### Code Implementation
```python
def _populate_quarter_tiles(slide, template_slide, subset):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    month_values = {month: float(subset[month].sum()) for month in months}

    for quarter_key, config in SUMMARY_TILE_CONFIG.get("quarter_budgets", {}).items():
        shape_name = config.get("shape")
        shape = next((s for s in slide.shapes if getattr(s, "name", "") == shape_name ...), None)

        quarter_months = {
            "q1": ("Jan", "Feb", "Mar"),
            "q2": ("Apr", "May", "Jun"),
            "q3": ("Jul", "Aug", "Sep"),
            "q4": ("Oct", "Nov", "Dec"),
        }.get(quarter_key.lower(), ())

        value = sum(month_values.get(month, 0.0) for month in quarter_months)
        formatted = _format_tile_value(config, value)
        prefix = config.get("prefix", "")

        _set_shape_text(shape, template_shape, f"{prefix}{formatted}")
```

### Calculation Rules

| Quarter | Months | Calculation | Scope |
|---------|--------|-------------|-------|
| **Q1** | Jan, Feb, Mar | `SUM(Jan, Feb, Mar)` | Brand-level |
| **Q2** | Apr, May, Jun | `SUM(Apr, May, Jun)` | Brand-level |
| **Q3** | Jul, Aug, Sep | `SUM(Jul, Aug, Sep)` | Brand-level |
| **Q4** | Oct, Nov, Dec | `SUM(Oct, Nov, Dec)` | Brand-level |

### **USER CLARIFICATION NEEDED:**

The user mentioned "quarter green rectangles" that need to "calculate at brand level, not at slide or campaign level." However:

1. **If referring to summary slide tiles:** Current implementation is correct (brand-level)
2. **If referring to visual indicators on DATA slides:** No such implementation exists in the code

**Questions for user:**
- Are you referring to shapes/rectangles that appear on the DATA slides (table slides)?
- Or are you referring to the quarter budget tiles on SUMMARY slides?
- Should there be visual indicators (green rectangles) showing which quarters have TV activity?

### Configuration
- **Config File:** `config/master_config.json`
- **Config Path:** `presentation.summary_tiles.quarter_budgets`

---

## 6. Slide-Level GRAND TOTAL (For Split Slides)

### **POTENTIAL ISSUE:**
When a brand spans multiple slides with smart pagination, does the system create:
- **Option A:** Per-slide GRAND TOTAL (showing accumulated total up to that slide)
- **Option B:** Only one final GRAND TOTAL on the last slide

### Implementation Location
- **File:** `amp_automation/presentation/assembly.py:2397-2420`
- **Context:** `_split_table_data_by_campaigns()` function

### Code Analysis
```python
# When creating per-slide GRAND TOTAL:
per_slide_row = _build_grand_total_row(slide_months, slide_total, slide_grp)
slide_data.append(per_slide_row)

per_slide_metadata = _build_grand_total_metadata_values(
    slide_months, slide_total, slide_grp
)
```

### **QUESTION FOR USER:**
Current code suggests per-slide GRAND TOTAL is created when splitting. Is this correct behavior, or should it be:
- GRAND TOTAL only on LAST slide of brand
- Intermediate slides show NO GRAND TOTAL or CARRIED FORWARD continuation

---

## Summary of Calculation Scopes

| Element | Scope | Appears On | Calculates |
|---------|-------|------------|------------|
| **Title** | Slide | Every slide | Market + Brand + Year |
| **MONTHLY TOTAL** | Campaign | After each campaign | Sum of campaign rows |
| **CARRIED FORWARD** | Accumulated | Continuation slides (2+) | Sum of previous slides |
| **GRAND TOTAL** | Brand | Last slide of brand | Sum of ALL campaigns in brand |
| **Quarter Tiles** | Brand | Summary slides | Sum of Q1/Q2/Q3/Q4 months |

---

## Critical Rules to Follow

### Rule 1: Calculation Scope Hierarchy
```
BRAND LEVEL (highest)
  ├─ GRAND TOTAL (all campaigns across all slides)
  ├─ Quarter Tiles (Q1-Q4 across all campaigns)
  │
  ├─ SLIDE LEVEL (intermediate)
  │    └─ CARRIED FORWARD (accumulated from previous slides)
  │
  └─ CAMPAIGN LEVEL (lowest)
       └─ MONTHLY TOTAL (single campaign only)
```

### Rule 2: Never Mix Scopes
❌ **WRONG:** MONTHLY TOTAL includes data from multiple campaigns
✅ **CORRECT:** MONTHLY TOTAL only sums rows for that specific campaign

❌ **WRONG:** GRAND TOTAL only shows current slide total
✅ **CORRECT:** GRAND TOTAL sums ALL campaigns in brand across ALL slides

### Rule 3: Slide Placement
- **MONTHLY TOTAL:** After every campaign
- **CARRIED FORWARD:** Top of continuation slides (slide 2+)
- **GRAND TOTAL:** Bottom of LAST slide for each brand
- **Quarter Tiles:** On summary slides (if exist)

### Rule 4: Title Format
Current: `{market} - {brand}`
Required: `{market} - {brand} [{year}]`
Example: `KSA - Sensodyne [25]`

---

## Action Items

### 1. Fix Title Format ✅ **TO DO**
Add year to title format in either:
- Code: `amp_automation/presentation/assembly.py:524`
- Config: `config/master_config.json` → `presentation.title.format`

### 2. Verify GRAND TOTAL Behavior ❓ **NEEDS USER CLARIFICATION**
Questions:
- Does GRAND TOTAL appear on every slide or only the last slide of a brand?
- Is current behavior correct?

### 3. Clarify Quarter Indicators ❓ **NEEDS USER CLARIFICATION**
Questions:
- Are "green rectangles" referring to summary slide tiles or data slide indicators?
- Should there be visual indicators on data slides showing TV activity by quarter?
- Current implementation only populates summary slide tiles - is this correct?

### 4. Document Post-Processing Rules ✅ **TO DO**
Post-processing operations (cell merges, font normalization) have their own rules:
- Campaign column merge (vertical, column 0)
- MONTHLY TOTAL merge (horizontal, columns 0-2)
- GRAND TOTAL merge (horizontal, columns 0-2)

---

## References

### Code Files
- **Main Assembly:** `amp_automation/presentation/assembly.py`
- **Post-Processing:** `amp_automation/presentation/postprocess/cell_merges.py`
- **Configuration:** `config/master_config.json`

### Key Functions
- `_compose_title_text()` - Line 518
- `_build_campaign_monthly_total_row()` - Line 909
- `_build_grand_total_row()` - Line 1209
- `_build_carried_forward_row()` - Line 1224
- `_populate_quarter_tiles()` - Line 164
- `_prepare_table_data()` - Line 2034 (main data preparation)

---

**Document Status:** DRAFT - Requires user verification and clarification
**Next Steps:** User to review and clarify questions marked ❓
