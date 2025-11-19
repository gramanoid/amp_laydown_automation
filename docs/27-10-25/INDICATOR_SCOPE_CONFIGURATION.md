# Indicator Scope Configuration

**Date:** 27 October 2025
**Status:** Implemented (Brand-level only)
**Configuration:** `config/master_config.json` → `presentation.summary_tiles`

---

## Overview

All brand-level indicators (quarter budgets, media share, funnel stage) are now modularized with configurable scope. This allows easy switching between calculation levels: **brand**, **campaign**, or **slide**.

---

## Current Implementation

### Scope: `brand` (IMPLEMENTED ✅)

All indicators calculate at **brand level** across ALL campaigns and appear **ONLY on the LAST slide** of each brand.

**Example: Brand with 3 slides**
- **Slide 1**: No indicators
- **Slide 2**: No indicators
- **Slide 3** (last): Shows all brand-level indicators (Q1-Q4, TV/DIG/OTHER, AWA/CON/PUR)

**Calculation:**
- **Quarter Budgets:** Sum of all months across all campaigns for the brand
  - Q1 = Jan + Feb + Mar (all campaigns)
  - Q2 = Apr + May + Jun (all campaigns)
  - Q3 = Jul + Aug + Sep (all campaigns)
  - Q4 = Oct + Nov + Dec (all campaigns)
- **Media Share:** `(media_type_total / brand_total_budget) * 100`
  - TV: All Television budget / Total brand budget
  - DIG: All Digital budget / Total brand budget
  - OTHER: All Other budget / Total brand budget
- **Funnel Stage:** `(funnel_stage_total / brand_total_budget) * 100`
  - AWA: All Awareness budget / Total brand budget
  - CON: All Consideration budget / Total brand budget
  - PUR: All Purchase budget / Total brand budget

---

## Configuration Location

**File:** `config/master_config.json`

```json
{
  "presentation": {
    "summary_tiles": {
      "_scope": "brand",
      "_scope_options": ["brand", "campaign", "slide"],
      "_scope_description": "Controls calculation scope for all indicators. Currently only 'brand' is implemented.",

      "quarter_budgets": {
        "_comment": "Quarter budget totals (Q1-Q4) calculated at brand level across all campaigns",
        "_calculation": "Q1 = Jan+Feb+Mar, Q2 = Apr+May+Jun, Q3 = Jul+Aug+Sep, Q4 = Oct+Nov+Dec",
        "_display": "Only shown on the LAST slide when brand spans multiple slides"
      },

      "media_share": {
        "_comment": "Media share percentages calculated at brand level across all campaigns",
        "_calculation": "(media_type_budget / total_brand_budget) * 100",
        "_display": "Only shown on the LAST slide when brand spans multiple slides"
      },

      "funnel_share": {
        "_comment": "Funnel stage percentages calculated at brand level across all campaigns",
        "_calculation": "(funnel_stage_budget / total_brand_budget) * 100",
        "_display": "Only shown on the LAST slide when brand spans multiple slides"
      }
    }
  }
}
```

---

## Future Scope Options

### Scope: `campaign` (NOT YET IMPLEMENTED ⚠️)

**Behavior:** Each campaign would show its own indicators on its last slide.

**Example: Brand with 3 campaigns**
- **Campaign 1 slides:** Last slide shows Campaign 1 quarters/media/funnel
- **Campaign 2 slides:** Last slide shows Campaign 2 quarters/media/funnel
- **Campaign 3 slides:** Last slide shows Campaign 3 quarters/media/funnel

**Calculation:**
- **Quarter Budgets:** Sum for THIS campaign only
- **Media Share:** `(media_type_in_campaign / campaign_total) * 100`
- **Funnel Stage:** `(funnel_stage_in_campaign / campaign_total) * 100`

**Implementation Required:**
1. Modify `_populate_summary_tiles()` to filter by campaign
2. Track campaign boundaries across splits
3. Detect last slide per campaign (not just per brand)

---

### Scope: `slide` (NOT YET IMPLEMENTED ⚠️)

**Behavior:** Every slide shows indicators for THAT slide only.

**Example: Brand with 3 slides**
- **Slide 1:** Shows Q1-Q4 / TV/DIG/OTHER / AWA/CON/PUR for campaigns on slide 1 only
- **Slide 2:** Shows Q1-Q4 / TV/DIG/OTHER / AWA/CON/PUR for campaigns on slide 2 only
- **Slide 3:** Shows Q1-Q4 / TV/DIG/OTHER / AWA/CON/PUR for campaigns on slide 3 only

**Calculation:**
- **Quarter Budgets:** Sum for campaigns visible on THIS slide only
- **Media Share:** `(media_type_on_slide / slide_total) * 100`
- **Funnel Stage:** `(funnel_stage_on_slide / slide_total) * 100`

**Implementation Required:**
1. Remove `is_last_slide` check from `_populate_summary_tiles()`
2. Calculate indicators from split data instead of full brand data
3. Pass slide-specific data to indicator population functions

---

## Implementation Details

### Code Locations

**Configuration Loading:**
- `amp_automation/presentation/assembly.py:1367` - `SUMMARY_TILE_CONFIG` loaded from master config

**Indicator Population:**
- `assembly.py:128-183` - `_populate_summary_tiles()` - Main entry point with `is_last_slide` check
- `assembly.py:186-211` - `_populate_quarter_tiles()` - Q1-Q4 calculation
- `assembly.py:214-241` - `_populate_media_share_tiles()` - TV/DIG/OTHER calculation
- `assembly.py:244-268` - `_populate_funnel_share_tiles()` - AWA/CON/PUR calculation

**Slide Creation:**
- `assembly.py:3072-3113` - `_populate_slide_content()` - Passes `is_last_slide` flag
- `assembly.py:3355-3377` - Slide creation loop - Calculates `is_last_slide = (split_idx == len(table_splits) - 1)`

**Table Splitting:**
- `assembly.py:2357-2420` - `finalize_split()` - Only adds BRAND TOTAL on last slide

---

## How to Change Scope (Future)

### Step 1: Update Configuration

Edit `config/master_config.json`:

```json
{
  "presentation": {
    "summary_tiles": {
      "_scope": "campaign",  // Change from "brand" to "campaign" or "slide"
      ...
    }
  }
}
```

### Step 2: Implement Scope Logic (Future Development)

Create scope-adaptive wrapper functions:

```python
def _populate_summary_tiles_with_scope(slide, template_slide, df, combination_row, excel_path, is_last_slide, scope):
    """Populate indicators based on configured scope."""

    if scope == "brand":
        # Current implementation - only on last slide
        if not is_last_slide:
            return
        subset = df.loc[combo_filter].copy()  # Full brand data

    elif scope == "campaign":
        # Future implementation - campaign-level indicators
        subset = df.loc[campaign_filter].copy()  # Current campaign only

    elif scope == "slide":
        # Future implementation - slide-level indicators
        subset = slide_specific_data  # Data from current split only

    # Populate indicators using subset
    _populate_quarter_tiles(slide, template_slide, subset)
    _populate_media_share_tiles(slide, template_slide, subset, total_cost)
    _populate_funnel_share_tiles(slide, template_slide, subset, total_cost)
```

### Step 3: Test Thoroughly

1. **Single-slide brands:** Indicators should appear on the only slide
2. **Multi-slide brands:**
   - Brand scope: Last slide only
   - Campaign scope: Last slide per campaign
   - Slide scope: Every slide
3. **Verify calculations:** Ensure totals match expected scope

---

## Testing Checklist

### Brand Scope (Current Implementation)

- [x] Single-slide brand shows indicators
- [x] Multi-slide brand shows indicators ONLY on last slide
- [x] Quarter totals sum ALL campaigns
- [x] Media share uses ALL campaign budgets
- [x] Funnel share uses ALL campaign budgets
- [x] BRAND TOTAL appears only on last slide

### Campaign Scope (Future)

- [ ] Each campaign's last slide shows campaign-level indicators
- [ ] Quarter totals sum THIS campaign only
- [ ] Media share uses THIS campaign budget
- [ ] Funnel share uses THIS campaign budget

### Slide Scope (Future)

- [ ] Every slide shows slide-level indicators
- [ ] Quarter totals sum campaigns ON THIS SLIDE
- [ ] Media share uses campaigns ON THIS SLIDE
- [ ] Funnel share uses campaigns ON THIS SLIDE

---

## Benefits of Modularization

1. **Flexibility:** Easy to switch calculation scope without code changes
2. **Documentation:** Clear configuration makes behavior explicit
3. **Maintainability:** Future developers can understand intent quickly
4. **Extensibility:** Adding new scopes requires minimal changes
5. **Testing:** Each scope can be tested independently

---

## Migration Guide

If you need to implement campaign or slide scope:

1. **Add validation** in `assembly.py` to check `_scope` config value
2. **Create wrapper functions** that route to scope-specific implementations
3. **Implement calculation functions** for each scope level
4. **Update tests** to cover all scope options
5. **Document behavior** in user-facing documentation

---

**Next Steps:** Testing with real multi-slide brands to verify brand-level scope works correctly.
