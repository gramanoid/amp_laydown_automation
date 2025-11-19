# Implementation Plan: Dynamic Elements Fix

**Date:** 27 October 2025
**Status:** IN PROGRESS
**Branch:** (To be created)

---

## Overview

Comprehensive fix for all dynamic elements on presentation slides to ensure correct calculation scopes and consistent display across multi-slide brands.

---

## Requirements Summary

### 1. ✅ Title - No Change Needed
- Already displays: `{MARKET} - {BRAND} ({YEAR})`
- Example: "RSA - SENSODYNE (25)"

### 2. ❌ MONTHLY TOTAL - Add GRP Column
**Current:** GRP column is blank
**Required:** Sum all GRP values for the campaign and display in MONTHLY TOTAL row

**Code Location:** `assembly.py:909-935` - `_build_campaign_monthly_total_row()`

**Change:**
```python
# Current (line 934):
row.extend(["", ""])  # GRP and % blank

# Required:
row.append(format_number(campaign_grp_total, is_grp=True))  # GRP total
row.append("")  # % remains blank
```

### 3. 🔄 GRAND TOTAL → BRAND TOTAL
**Changes Required:**
- Rename "GRAND TOTAL" to "BRAND TOTAL"
- Add special styling (different background color or border)
- Only appears on LAST slide of each brand
- Add explanatory note below table

**Code Locations:**
- `assembly.py:1209-1221` - `_build_grand_total_row()` - Change label
- `assembly.py:2160-2165` - Main table data preparation - Ensure only on last slide
- Post-processing styling updates

**Implementation:**
1. Change label text from "GRAND TOTAL" to "BRAND TOTAL"
2. Add metadata flag for special styling
3. Implement logic to only add BRAND TOTAL on final slide
4. Add note shape below table (if not in template)

### 4. 🟢 Quarter Boxes (Q1-Q4) - Brand Level on LAST SLIDE ONLY
**Current Behavior:** Only on summary slides (assembly.py:164)
**Required:** On LAST DATA slide of each brand with brand-level totals

**Quarter Definitions:**
- Q1 = Jan + Feb + Mar
- Q2 = Apr + May + Jun
- Q3 = Jul + Aug + Sep
- Q4 = Oct + Nov + Dec

**Implementation Strategy:**
1. Calculate brand-level quarter totals ONCE per brand
2. Store in brand context/metadata
3. Add quarter shapes to LAST slide only for that brand
4. Same slide that gets BRAND TOTAL row

**Code Changes:**
- New function: `_calculate_brand_quarter_totals(subset)` → returns Q1-Q4 dict
- Modify: `_create_single_slide()` to add quarter shapes when `is_last_slide=True`
- Update: Template shape names for quarter boxes

### 5. 📊 Media Share Boxes - Brand Level on LAST SLIDE ONLY
**Displays:**
- TV: X%
- DIG.: X%
- OTHER: X%

**Calculation:** `(media_type_budget / total_brand_budget) * 100`

**Implementation Strategy:**
1. Calculate brand-level media share ONCE per brand
2. Store in brand context/metadata
3. Add media share shapes to LAST slide only
4. Same slide that gets BRAND TOTAL row

**Code Changes:**
- New function: `_calculate_brand_media_share(subset)` → returns TV/DIG/OTHER %
- Modify: `_create_single_slide()` to add media share shapes when `is_last_slide=True`
- Update: Template shape names for media boxes

### 6. 🎯 Funnel Stage Boxes - Brand Level on LAST SLIDE ONLY
**Displays:**
- AWA: X% (Awareness)
- CON: X% (Consideration)
- PUR: X% (Purchase)

**Calculation:** `(funnel_stage_budget / total_brand_budget) * 100`

**Implementation Strategy:**
1. Calculate brand-level funnel share ONCE per brand
2. Store in brand context/metadata
3. Add funnel stage shapes to LAST slide only
4. Same slide that gets BRAND TOTAL row

**Code Changes:**
- New function: `_calculate_brand_funnel_share(subset)` → returns AWA/CON/PUR %
- Modify: `_create_single_slide()` to add funnel shapes when `is_last_slide=True`
- Update: Template shape names for funnel boxes

### 7. 🔤 Campaign Text Wrapping - Full Words Only
**Current:** May break words mid-word
**Required:** Only wrap on full words (space boundaries)

**Code Location:** `postprocess/cell_merges.py:21-107` - `merge_campaign_cells()`

**Implementation:**
1. Set text frame word wrap to True
2. Ensure no forced line breaks mid-word
3. May need to adjust cell width or font size for long campaign names

**Python-pptx Properties:**
```python
text_frame.word_wrap = True  # Ensure enabled
# Text will automatically wrap on word boundaries
```

### 8. 🗑️ Remove CARRIED FORWARD Logic
**Current:** CARRIED FORWARD rows exist in code
**Required:** Completely remove this feature

**Code Locations:**
- `assembly.py:1224-1269` - `_build_carried_forward_row()` - DELETE
- `assembly.py:1239-1269` - `_build_carried_forward_metadata_values()` - DELETE
- Any references to "CARRIED FORWARD" in splitting logic
- Post-processing: `cell_merges.py:282-293` - `is_carried_forward()` - DELETE

**Search Pattern:** `CARRIED FORWARD|carried.forward|CarriedForward`

### 9. 🧩 Modularization - Configurable Scope
**Requirement:** All brand-level indicators should be modularized so scope can be easily changed

**Design:**
```python
# New configuration structure
INDICATOR_SCOPE_CONFIG = {
    "quarter_boxes": "brand",      # or "campaign", "slide"
    "media_share": "brand",        # or "campaign", "slide"
    "funnel_stage": "brand",       # or "campaign", "slide"
}

# Wrapper functions that adapt based on scope
def _calculate_quarter_totals(data, scope, context):
    if scope == "brand":
        return _calculate_brand_quarter_totals(data)
    elif scope == "campaign":
        return _calculate_campaign_quarter_totals(data)
    elif scope == "slide":
        return _calculate_slide_quarter_totals(data)
```

---

## Implementation Order

### Phase 1: Quick Fixes (30 min)
1. ✅ Verify title (already correct)
2. ❌ Add GRP to MONTHLY TOTAL
3. 🗑️ Remove CARRIED FORWARD logic
4. 🔤 Fix campaign text wrapping

### Phase 2: BRAND TOTAL Rename (45 min)
1. Change label from GRAND TOTAL to BRAND TOTAL
2. Add special styling
3. Ensure only on last slide
4. Add explanatory note

### Phase 3: Brand-Level Indicators (2-3 hours)
1. 🟢 Implement brand-level quarter boxes
2. 📊 Implement brand-level media share boxes
3. 🎯 Implement brand-level funnel stage boxes
4. Test multi-slide brand scenarios

### Phase 4: Modularization (1 hour)
1. Create configuration for scope selection
2. Implement scope-adaptive wrapper functions
3. Add validation and error handling
4. Document configuration options

### Phase 5: Testing & Verification (1 hour)
1. Generate test deck with multi-slide brands
2. Verify all indicators show correct brand-level totals
3. Verify indicators appear on every slide
4. Verify BRAND TOTAL only on last slide
5. Run post-processing pipeline
6. Manual visual inspection

---

## Code Structure Changes

### New Functions to Create

```python
# Brand-level calculations (run once per brand)
def _calculate_brand_quarter_totals(subset: pd.DataFrame) -> dict[str, float]:
    """Calculate Q1-Q4 totals for entire brand."""
    pass

def _calculate_brand_media_share(subset: pd.DataFrame) -> dict[str, float]:
    """Calculate TV/DIG/OTHER percentages for entire brand."""
    pass

def _calculate_brand_funnel_share(subset: pd.DataFrame) -> dict[str, float]:
    """Calculate AWA/CON/PUR percentages for entire brand."""
    pass

# Slide population (run for every slide)
def _populate_brand_indicators(slide, template_slide, brand_context):
    """Add quarter boxes, media share, and funnel stage to slide."""
    _populate_quarter_boxes(slide, template_slide, brand_context["quarters"])
    _populate_media_share_boxes(slide, template_slide, brand_context["media_share"])
    _populate_funnel_stage_boxes(slide, template_slide, brand_context["funnel_share"])

# Modular wrappers
def _calculate_indicators_by_scope(data, scope_config, context):
    """Calculate indicators based on configured scope (brand/campaign/slide)."""
    pass
```

### Modified Functions

```python
# Update to include GRP in MONTHLY TOTAL
def _build_campaign_monthly_total_row(
    row_idx: int,
    month_totals: list[float],
    campaign_grp_total: float,  # NEW parameter
    cell_metadata: dict[tuple[int, int], dict[str, object]],
) -> list[str]:
    # ... existing code ...
    row.append(format_number(campaign_grp_total, is_grp=True))  # NEW
    row.append("")  # % remains blank
    return row

# Update to use "BRAND TOTAL" label
def _build_brand_total_row(  # Renamed from _build_grand_total_row
    monthly_totals: list[float],
    total_budget: float,
    brand_total_grp: float,
) -> list[str]:
    row: list[str] = ["BRAND TOTAL", "", ""]  # Changed label
    # ... rest of implementation
    return row

# Update main creation flow
def _create_presentation_for_combination(
    prs, template_slide, combination_row, subset, excel_path
):
    # Calculate brand-level indicators ONCE
    brand_context = {
        "quarters": _calculate_brand_quarter_totals(subset),
        "media_share": _calculate_brand_media_share(subset),
        "funnel_share": _calculate_brand_funnel_share(subset),
    }

    # Split into slides
    slides_data = _split_table_data_by_campaigns(...)

    # Create each slide
    for slide_idx, slide_data in enumerate(slides_data):
        is_last_slide = (slide_idx == len(slides_data) - 1)

        slide = _create_single_slide(...)

        # Add brand-level indicators ONLY on last slide
        if is_last_slide:
            _populate_brand_indicators(slide, template_slide, brand_context)
            # Also add BRAND TOTAL row with special styling
```

---

## Template Requirements

### Shape Names Needed

**Quarter Boxes:**
- `Q1_Box` - Q1 budget total
- `Q2_Box` - Q2 budget total
- `Q3_Box` - Q3 budget total
- `Q4_Box` - Q4 budget total

**Media Share Boxes:**
- `TV_Share_Box` - TV percentage
- `DIG_Share_Box` - Digital percentage
- `OTHER_Share_Box` - Other media percentage

**Funnel Stage Boxes:**
- `AWA_Box` - Awareness percentage
- `CON_Box` - Consideration percentage
- `PUR_Box` - Purchase percentage

**Notes:**
- `Brand_Total_Note` - Explanatory note for BRAND TOTAL (optional)

### Template Verification Script
```python
# Check if template has required shapes
required_shapes = [
    "Q1_Box", "Q2_Box", "Q3_Box", "Q4_Box",
    "TV_Share_Box", "DIG_Share_Box", "OTHER_Share_Box",
    "AWA_Box", "CON_Box", "PUR_Box"
]

template_prs = Presentation("template/Template_V4_FINAL_071025.pptx")
template_slide = template_prs.slides[0]  # Adjust index

for shape_name in required_shapes:
    shape = next((s for s in template_slide.shapes if getattr(s, "name", "") == shape_name), None)
    if not shape:
        print(f"❌ Missing: {shape_name}")
    else:
        print(f"✅ Found: {shape_name}")
```

---

## Configuration Updates

### master_config.json

```json
{
  "presentation": {
    "indicators": {
      "quarter_boxes": {
        "scope": "brand",
        "enabled": true,
        "shapes": {
          "q1": "Q1_Box",
          "q2": "Q2_Box",
          "q3": "Q3_Box",
          "q4": "Q4_Box"
        }
      },
      "media_share": {
        "scope": "brand",
        "enabled": true,
        "shapes": {
          "tv": "TV_Share_Box",
          "digital": "DIG_Share_Box",
          "other": "OTHER_Share_Box"
        }
      },
      "funnel_stage": {
        "scope": "brand",
        "enabled": true,
        "shapes": {
          "awareness": "AWA_Box",
          "consideration": "CON_Box",
          "purchase": "PUR_Box"
        }
      }
    },
    "table": {
      "brand_total": {
        "label": "BRAND TOTAL",
        "only_last_slide": true,
        "special_styling": {
          "background_color": [200, 200, 200],
          "font_bold": true,
          "border_color": [0, 0, 0],
          "border_width": 2
        },
        "note_text": "* Brand Total represents sum of all campaigns across multiple slides"
      }
    }
  }
}
```

---

## Testing Checklist

### Single-Slide Brand
- [ ] Quarter boxes show correct totals
- [ ] Media share shows correct percentages
- [ ] Funnel stage shows correct percentages
- [ ] BRAND TOTAL appears at bottom
- [ ] MONTHLY TOTAL includes GRP values

### Multi-Slide Brand (e.g., 3 slides)
- [ ] Quarter boxes appear ONLY on slide 3 (last) with brand-level totals
- [ ] Media share appears ONLY on slide 3 (last) with brand-level percentages
- [ ] Funnel stage appears ONLY on slide 3 (last) with brand-level percentages
- [ ] BRAND TOTAL appears ONLY on slide 3 (last)
- [ ] BRAND TOTAL sums ALL campaigns from all 3 slides
- [ ] MONTHLY TOTAL on each slide shows GRP for that campaign only
- [ ] Slides 1-2 have NO quarter boxes, NO media share, NO funnel stage, NO BRAND TOTAL

### Post-Processing
- [ ] Campaign cells merge vertically
- [ ] MONTHLY TOTAL cells merge horizontally
- [ ] BRAND TOTAL cells merge horizontally
- [ ] Campaign text wraps on full words only
- [ ] Font normalization works correctly
- [ ] No CARRIED FORWARD rows exist

### Visual Verification
- [ ] All boxes are visible and correctly positioned
- [ ] Text is readable and properly formatted
- [ ] BRAND TOTAL has special styling
- [ ] Campaign names don't break mid-word

---

## Rollback Plan

If issues arise:
1. Create feature branch before changes
2. Commit after each phase
3. Test after each phase
4. Can revert to any previous phase if needed

**Branch naming:** `fix/brand-level-indicators`

---

## Time Estimate

| Phase | Duration | Cumulative |
|-------|----------|------------|
| Phase 1: Quick Fixes | 30 min | 30 min |
| Phase 2: BRAND TOTAL | 45 min | 1h 15min |
| Phase 3: Brand Indicators | 2-3 hours | 3h 15min - 4h 15min |
| Phase 4: Modularization | 1 hour | 4h 15min - 5h 15min |
| Phase 5: Testing | 1 hour | 5h 15min - 6h 15min |

**Total:** 5-6 hours

---

## Success Criteria

1. ✅ MONTHLY TOTAL includes GRP column with correct campaign GRP sum
2. ✅ BRAND TOTAL (renamed) appears only on last slide with special styling
3. ✅ Quarter boxes show brand-level totals on every slide
4. ✅ Media share boxes show brand-level percentages on every slide
5. ✅ Funnel stage boxes show brand-level percentages on every slide
6. ✅ Campaign text wraps on full words only
7. ✅ No CARRIED FORWARD logic remains
8. ✅ All indicators are modularized with configurable scope
9. ✅ Multi-slide brands display correctly
10. ✅ Post-processing completes without errors

---

**Next Steps:** User approval to proceed with implementation
