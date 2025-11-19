# Full Implementation Plan - Dynamic Elements Fix (REVISED)

**Date:** 27 October 2025
**Status:** AWAITING APPROVAL
**Estimated Time:** 5-6 hours
**Branch:** `fix/brand-level-indicators`

---

## Executive Summary

Comprehensive fix for all dynamic elements on presentation slides with the following key changes:

1. ✅ **Title** - Already correct, no changes needed
2. ❌ **MONTHLY TOTAL** - Add GRP column calculation
3. 🔄 **GRAND TOTAL → BRAND TOTAL** - Rename, special styling, last slide only
4. 🟢 **Quarter Boxes** - Brand-level, last slide only
5. 📊 **Media Share Boxes** - Brand-level, last slide only
6. 🎯 **Funnel Stage Boxes** - Brand-level, last slide only
7. 🔤 **Campaign Text Wrapping** - Full words only
8. 🗑️ **Remove CARRIED FORWARD** - Complete removal
9. 🧩 **Modularization** - Configurable scope for all indicators

### Key Design Decision: **LAST SLIDE ONLY**
All brand-level indicators (BRAND TOTAL, quarter boxes, media share, funnel stage) appear **ONLY on the LAST slide** of each brand.

---

## Visual Example

### Multi-Slide Brand: "KSA - Sensodyne" (3 slides)

#### **Slide 1 & 2** (Intermediate slides)
```
┌──────────────────────────────────────────────┐
│ KSA - SENSODYNE (25)                         │
├──────────────────────────────────────────────┤
│ CAMPAIGN    │ MEDIA  │ METRICS │ JAN │ ... │
│ Clinical    │ TV     │ £ 000   │ 69K │ ... │
│ White       │        │ GRPs    │ 861 │ ... │
│             │ Digital│ £ 000   │ 44K │ ... │
│ MONTHLY TOTAL (£ 000)  │ 69K │ ... │ 2M  │
│                        │     │ ... │ 456 │ ← GRP
│                                              │
│ [More campaigns...]                          │
│                                              │
│ ❌ NO BRAND TOTAL                            │
│ ❌ NO Quarter Boxes                          │
│ ❌ NO Media Share                            │
│ ❌ NO Funnel Stage                           │
└──────────────────────────────────────────────┘
```

#### **Slide 3** (LAST slide - has everything)
```
┌──────────────────────────────────────────────┐
│ KSA - SENSODYNE (25)                         │
├──────────────────────────────────────────────┤
│ CAMPAIGN    │ MEDIA  │ METRICS │ JAN │ ... │
│ Feel        │ TV     │ £ 000   │ 113K│ ... │
│ Familiar    │        │ GRPs    │ 1.3K│ ... │
│             │ Digital│ £ 000   │ 343K│ ... │
│ MONTHLY TOTAL (£ 000)  │ 113K│ ... │ 814K │
│                        │     │ ... │ 4.9K │ ← GRP
│                                              │
│ ✅ BRAND TOTAL*        │ 184K│ ... │ 4M   │
│                        │     │ ... │ 100% │
│                                              │
│ ✅ Q1: £659K   Q2: £1,211K   Q3: £1,828K   │
│    Q4: £1,010K                              │
│                                              │
│ ✅ TV: 0%   DIG: 48%   OTHER: 2%            │
│                                              │
│ ✅ AWA: 82%   CON: 14%   PUR: 4%            │
│                                              │
│ * Brand Total = Sum of all campaigns        │
└──────────────────────────────────────────────┘
```

---

## Detailed Changes by Component

---

### 1. ✅ TITLE - No Changes Needed

**Current Implementation:** Already displays `{MARKET} - {BRAND} ({YEAR})`

**Example:** "KSA - SENSODYNE (25)" ✅

**Status:** ✅ VERIFIED - No action needed

---

### 2. ❌ MONTHLY TOTAL - Add GRP Column

**File:** `amp_automation/presentation/assembly.py`
**Function:** `_build_campaign_monthly_total_row()` (lines 909-935)
**Time Estimate:** 20 minutes

#### Current Code (Lines 909-935)
```python
def _build_campaign_monthly_total_row(
    row_idx: int,
    month_totals: list[float],
    cell_metadata: dict[tuple[int, int], dict[str, object]],
) -> list[str]:
    row: list[str] = ["MONTHLY TOTAL (£ 000)", "", ""]

    # Add monthly budget totals (columns 3-14)
    for month_idx, value in enumerate(month_totals):
        formatted = _format_budget_cell(value)
        col_idx = 3 + month_idx
        row.append(formatted)
        _set_cell_metadata(...)

    # Add total budget (column 15)
    total_value = sum(month_totals)
    total_col_idx = 3 + len(TABLE_MONTH_ORDER)
    total_formatted = _format_total_budget(total_value)
    row.append(total_formatted)
    _set_cell_metadata(...)

    # ❌ PROBLEM: GRP and % columns left blank
    row.extend(["", ""])
    return row
```

#### New Code
```python
def _build_campaign_monthly_total_row(
    row_idx: int,
    month_totals: list[float],
    campaign_grp_total: float,  # ✅ NEW PARAMETER
    cell_metadata: dict[tuple[int, int], dict[str, object]],
) -> list[str]:
    row: list[str] = ["MONTHLY TOTAL (£ 000)", "", ""]

    # Add monthly budget totals (columns 3-14)
    for month_idx, value in enumerate(month_totals):
        formatted = _format_budget_cell(value)
        col_idx = 3 + month_idx
        row.append(formatted)
        _set_cell_metadata(...)

    # Add total budget (column 15)
    total_value = sum(month_totals)
    total_col_idx = 3 + len(TABLE_MONTH_ORDER)
    total_formatted = _format_total_budget(total_value)
    row.append(total_formatted)
    _set_cell_metadata(...)

    # ✅ NEW: Add GRP total (column 16)
    grp_col_idx = total_col_idx + 1
    grp_formatted = format_number(campaign_grp_total, is_grp=True)
    row.append(grp_formatted if grp_formatted else "")
    _set_cell_metadata(
        cell_metadata,
        row_idx,
        grp_col_idx,
        campaign_grp_total,
        "GRPs",
        not is_empty_formatted_value(grp_formatted),
    )

    # % column remains blank
    row.append("")
    return row
```

#### Changes to Calling Code
**File:** `amp_automation/presentation/assembly.py`
**Function:** `_build_campaign_block()` (around line 1199)

```python
# Current:
monthly_total_row = _build_campaign_monthly_total_row(
    row_idx,
    block_month_totals,
    cell_metadata,
)

# New:
monthly_total_row = _build_campaign_monthly_total_row(
    row_idx,
    block_month_totals,
    block_grp_total,  # ✅ Pass campaign GRP total
    cell_metadata,
)
```

**Impact:**
- MONTHLY TOTAL now shows campaign GRP sum in column 16
- Column 17 (%) remains blank (correct behavior)

---

### 3. 🔄 GRAND TOTAL → BRAND TOTAL

**Files:**
- `amp_automation/presentation/assembly.py` (label change)
- `amp_automation/presentation/postprocess/cell_merges.py` (merge logic update)

**Time Estimate:** 45 minutes

#### 3.1. Rename Label

**File:** `assembly.py`
**Function:** `_build_grand_total_row()` (line 1209)

```python
# Current (line 1214):
row: list[str] = ["GRAND TOTAL", "", ""]

# New:
row: list[str] = ["BRAND TOTAL", "", ""]
```

#### 3.2. Update Post-Processing

**File:** `postprocess/cell_merges.py`
**Function:** `merge_summary_cells()` (line 173)

```python
# Current (line 200):
if (is_grand_total(cell_text) or is_carried_forward(cell_text)) and _has_gray_background(cell):

# New:
if (is_brand_total(cell_text) or is_carried_forward(cell_text)) and _has_gray_background(cell):
```

Add new helper function:
```python
def is_brand_total(cell_text: str) -> bool:
    """Check if cell text represents a BRAND TOTAL row."""
    normalized = normalize_label(cell_text)
    return "BRAND" in normalized and "TOTAL" in normalized
```

Update existing function to also check for "BRAND TOTAL":
```python
def is_grand_total(cell_text: str) -> bool:
    """Check if cell text represents a GRAND/BRAND TOTAL row."""
    normalized = normalize_label(cell_text)
    return ("GRAND" in normalized or "BRAND" in normalized) and "TOTAL" in normalized
```

#### 3.3. Special Styling

**Options for styling:**
- **Option A:** Darker gray background (RGB: 180, 180, 180 instead of 217, 217, 217)
- **Option B:** Bold text + thicker border
- **Option C:** Different color entirely (light blue/yellow highlight)

**Recommended:** Option A (darker gray) - most professional

**Implementation in `postprocess/cell_merges.py`:**
```python
# In merge_summary_cells():
if is_brand_total(cell_text):
    # Apply special styling
    merged_cell = table.cell(row_idx, 0)

    # Darker background for BRAND TOTAL
    fill = merged_cell.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(180, 180, 180)  # Darker gray

    # Bold text
    for paragraph in merged_cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.bold = True
```

#### 3.4. Last Slide Only Logic

**File:** `assembly.py`
**Function:** Main table creation flow (around line 2160)

**Current behavior:** GRAND TOTAL added to every brand's table data

**New behavior:** BRAND TOTAL only added when creating LAST slide

```python
# In _prepare_table_data():
# Build table with all campaigns
for campaign_name in campaign_names:
    # ... build campaign blocks
    monthly_totals = [total + addition for ...]
    grand_total_grp += block_grp_total

# ❌ OLD: Always add GRAND TOTAL
# grand_total_row = _build_grand_total_row(monthly_totals, total_budget, grand_total_grp)
# table_rows.append(grand_total_row)

# ✅ NEW: Don't add here - will be added during slide creation
# Return brand totals in metadata for later use
return table_rows, cell_metadata, {
    "monthly_totals": monthly_totals,
    "total_budget": total_budget,
    "brand_grp_total": grand_total_grp,
}
```

Then in slide creation:
```python
# In _create_presentation_for_combination():
for slide_idx, slide_data in enumerate(slides_data):
    is_last_slide = (slide_idx == len(slides_data) - 1)

    # Create slide with data
    slide = _create_single_slide(...)

    # Add BRAND TOTAL only on last slide
    if is_last_slide:
        brand_total_row = _build_brand_total_row(
            brand_metadata["monthly_totals"],
            brand_metadata["total_budget"],
            brand_metadata["brand_grp_total"],
        )
        # Add row to table on last slide
        _add_brand_total_to_table(table, brand_total_row)
```

#### 3.5. Explanatory Note

**Add text shape below table on last slide:**
```python
if is_last_slide:
    # Add note shape
    note_shape = slide.shapes.add_textbox(
        left=Inches(0.5),
        top=Inches(5.0),
        width=Inches(9.0),
        height=Inches(0.3)
    )
    text_frame = note_shape.text_frame
    text_frame.text = "* Brand Total represents the sum of all campaigns across all slides for this brand"

    # Style note
    for paragraph in text_frame.paragraphs:
        paragraph.font.size = Pt(8)
        paragraph.font.italic = True
        paragraph.alignment = PP_ALIGN.LEFT
```

---

### 4. 🟢 Quarter Boxes (Q1-Q4) - Last Slide Only

**Time Estimate:** 1 hour

#### 4.1. Calculate Brand-Level Quarter Totals

**New Function:**
```python
def _calculate_brand_quarter_totals(subset: pd.DataFrame) -> dict[str, float]:
    """
    Calculate Q1-Q4 budget totals for entire brand.

    Args:
        subset: DataFrame with all data for the brand

    Returns:
        Dictionary with Q1-Q4 totals
    """
    # Month columns in data
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    # Calculate total for each month
    month_totals = {}
    for month in months:
        if month in subset.columns:
            month_totals[month] = float(subset[month].sum() or 0.0)
        else:
            month_totals[month] = 0.0

    # Group into quarters
    quarters = {
        "Q1": sum(month_totals.get(m, 0.0) for m in ["Jan", "Feb", "Mar"]),
        "Q2": sum(month_totals.get(m, 0.0) for m in ["Apr", "May", "Jun"]),
        "Q3": sum(month_totals.get(m, 0.0) for m in ["Jul", "Aug", "Sep"]),
        "Q4": sum(month_totals.get(m, 0.0) for m in ["Oct", "Nov", "Dec"]),
    }

    return quarters
```

#### 4.2. Add Quarter Boxes to Last Slide

**New Function:**
```python
def _populate_quarter_boxes(slide, template_slide, quarters: dict[str, float]):
    """
    Add quarter budget boxes to slide.

    Args:
        slide: Current slide
        template_slide: Template reference
        quarters: Dict with Q1-Q4 totals
    """
    quarter_shapes = {
        "Q1": "Q1_Box",
        "Q2": "Q2_Box",
        "Q3": "Q3_Box",
        "Q4": "Q4_Box",
    }

    for quarter_key, shape_name in quarter_shapes.items():
        # Find shape by name
        shape = next(
            (s for s in slide.shapes if getattr(s, "name", "") == shape_name),
            None
        )

        if not shape:
            logger.warning(f"Quarter box shape '{shape_name}' not found")
            continue

        # Format value
        value = quarters.get(quarter_key, 0.0)
        formatted = _format_currency_short(value)

        # Set text
        text_frame = shape.text_frame
        text_frame.clear()
        paragraph = text_frame.paragraphs[0]
        paragraph.text = f"{quarter_key}: {formatted}"

        # Style
        paragraph.font.name = "Verdana"
        paragraph.font.size = Pt(10)
        paragraph.font.bold = True
        paragraph.alignment = PP_ALIGN.CENTER
```

#### 4.3. Integration in Main Flow

```python
# In _create_presentation_for_combination():

# Calculate brand-level indicators ONCE
brand_indicators = {
    "quarters": _calculate_brand_quarter_totals(subset),
    "media_share": _calculate_brand_media_share(subset),
    "funnel_stage": _calculate_brand_funnel_share(subset),
}

# Create slides
for slide_idx, slide_data in enumerate(slides_data):
    is_last_slide = (slide_idx == len(slides_data) - 1)

    slide = _create_single_slide(...)

    # Add indicators only on last slide
    if is_last_slide:
        _populate_quarter_boxes(slide, template_slide, brand_indicators["quarters"])
```

---

### 5. 📊 Media Share Boxes - Last Slide Only

**Time Estimate:** 45 minutes

#### 5.1. Calculate Brand-Level Media Share

**New Function:**
```python
def _calculate_brand_media_share(subset: pd.DataFrame) -> dict[str, float]:
    """
    Calculate media type percentages for entire brand.

    Args:
        subset: DataFrame with all data for the brand

    Returns:
        Dictionary with TV, Digital, Other percentages
    """
    total_budget = float(subset["Total Cost"].sum() or 0.0)

    if total_budget == 0:
        return {"TV": 0.0, "Digital": 0.0, "Other": 0.0}

    # Group by media type
    media_column = "Mapped Media Type"
    if media_column not in subset.columns:
        logger.warning(f"Column '{media_column}' not found")
        return {"TV": 0.0, "Digital": 0.0, "Other": 0.0}

    media_totals = subset.groupby(media_column)["Total Cost"].sum()

    # Calculate percentages
    tv_budget = float(media_totals.get("Television", 0.0))
    digital_budget = float(media_totals.get("Digital", 0.0))
    other_budget = total_budget - tv_budget - digital_budget

    return {
        "TV": (tv_budget / total_budget) * 100.0,
        "Digital": (digital_budget / total_budget) * 100.0,
        "Other": (other_budget / total_budget) * 100.0,
    }
```

#### 5.2. Add Media Share Boxes to Last Slide

**New Function:**
```python
def _populate_media_share_boxes(slide, template_slide, media_share: dict[str, float]):
    """
    Add media share percentage boxes to slide.

    Args:
        slide: Current slide
        template_slide: Template reference
        media_share: Dict with TV/Digital/Other percentages
    """
    media_shapes = {
        "TV": "TV_Share_Box",
        "Digital": "DIG_Share_Box",
        "Other": "OTHER_Share_Box",
    }

    for media_type, shape_name in media_shapes.items():
        shape = next(
            (s for s in slide.shapes if getattr(s, "name", "") == shape_name),
            None
        )

        if not shape:
            logger.warning(f"Media share box '{shape_name}' not found")
            continue

        # Format percentage
        percentage = media_share.get(media_type, 0.0)

        # Set text
        text_frame = shape.text_frame
        text_frame.clear()
        paragraph = text_frame.paragraphs[0]

        # Format display label
        display_label = "DIG." if media_type == "Digital" else media_type.upper()
        paragraph.text = f"{display_label}: {percentage:.0f}%"

        # Style
        paragraph.font.name = "Verdana"
        paragraph.font.size = Pt(9)
        paragraph.alignment = PP_ALIGN.CENTER
```

---

### 6. 🎯 Funnel Stage Boxes - Last Slide Only

**Time Estimate:** 45 minutes

#### 6.1. Calculate Brand-Level Funnel Share

**New Function:**
```python
def _calculate_brand_funnel_share(subset: pd.DataFrame) -> dict[str, float]:
    """
    Calculate funnel stage percentages for entire brand.

    Args:
        subset: DataFrame with all data for the brand

    Returns:
        Dictionary with Awareness/Consideration/Purchase percentages
    """
    total_budget = float(subset["Total Cost"].sum() or 0.0)

    if total_budget == 0:
        return {"Awareness": 0.0, "Consideration": 0.0, "Purchase": 0.0}

    # Group by funnel stage
    funnel_column = "Funnel Stage"
    if funnel_column not in subset.columns:
        logger.warning(f"Column '{funnel_column}' not found")
        return {"Awareness": 0.0, "Consideration": 0.0, "Purchase": 0.0}

    funnel_totals = subset.groupby(funnel_column)["Total Cost"].sum()

    # Calculate percentages
    awareness = float(funnel_totals.get("Awareness", 0.0))
    consideration = float(funnel_totals.get("Consideration", 0.0))
    purchase = float(funnel_totals.get("Purchase", 0.0))

    return {
        "Awareness": (awareness / total_budget) * 100.0,
        "Consideration": (consideration / total_budget) * 100.0,
        "Purchase": (purchase / total_budget) * 100.0,
    }
```

#### 6.2. Add Funnel Stage Boxes to Last Slide

**New Function:**
```python
def _populate_funnel_stage_boxes(slide, template_slide, funnel_share: dict[str, float]):
    """
    Add funnel stage percentage boxes to slide.

    Args:
        slide: Current slide
        template_slide: Template reference
        funnel_share: Dict with Awareness/Consideration/Purchase percentages
    """
    funnel_shapes = {
        "Awareness": "AWA_Box",
        "Consideration": "CON_Box",
        "Purchase": "PUR_Box",
    }

    for stage, shape_name in funnel_shapes.items():
        shape = next(
            (s for s in slide.shapes if getattr(s, "name", "") == shape_name),
            None
        )

        if not shape:
            logger.warning(f"Funnel stage box '{shape_name}' not found")
            continue

        # Format percentage
        percentage = funnel_share.get(stage, 0.0)

        # Set text
        text_frame = shape.text_frame
        text_frame.clear()
        paragraph = text_frame.paragraphs[0]

        # Format display label (3-letter abbreviation)
        display_label = stage[:3].upper()
        paragraph.text = f"{display_label}: {percentage:.0f}%"

        # Style
        paragraph.font.name = "Verdana"
        paragraph.font.size = Pt(9)
        paragraph.alignment = PP_ALIGN.CENTER
```

---

### 7. 🔤 Campaign Text Wrapping - Full Words Only

**File:** `amp_automation/presentation/postprocess/cell_merges.py`
**Function:** `_apply_cell_styling()` (around line 400)
**Time Estimate:** 15 minutes

#### Current Code
```python
def _apply_cell_styling(cell, text, font_size, bold, center_align, vertical_center):
    """Apply styling to merged cell."""
    text_frame = cell.text_frame
    text_frame.clear()
    # ... rest of styling
```

#### Updated Code
```python
def _apply_cell_styling(cell, text, font_size, bold, center_align, vertical_center):
    """Apply styling to merged cell."""
    text_frame = cell.text_frame
    text_frame.clear()

    # ✅ Enable word wrap (full words only, no mid-word breaks)
    text_frame.word_wrap = True

    # ... rest of styling (paragraphs, fonts, alignment)
```

**Note:** Python-pptx's `word_wrap=True` automatically wraps on word boundaries (spaces, hyphens). This is the correct behavior.

**Additional safeguard:**
```python
# If campaign name is extremely long and still doesn't fit:
if len(text) > 50:  # Arbitrary threshold
    # Consider abbreviating or using smaller font
    logger.warning(f"Campaign name very long: {text}")
```

---

### 8. 🗑️ Remove CARRIED FORWARD Logic

**Time Estimate:** 30 minutes

#### Files to Modify

**1. `assembly.py` - Remove builder functions**
```python
# DELETE ENTIRE FUNCTIONS (lines 1224-1269):
# - _build_carried_forward_row()
# - _build_carried_forward_metadata_values()
```

**2. `assembly.py` - Remove from splitting logic**

Search for "CARRIED FORWARD" and remove all references:
```bash
grep -n "CARRIED FORWARD\|carried.forward" assembly.py
```

Remove any code that:
- Builds CARRIED FORWARD rows
- Adds CARRIED FORWARD to split slides
- Tracks carried forward totals

**3. `postprocess/cell_merges.py` - Remove merge logic**
```python
# In merge_summary_cells() (line 200):
# OLD:
if (is_grand_total(cell_text) or is_carried_forward(cell_text)) and _has_gray_background(cell):

# NEW:
if is_brand_total(cell_text) and _has_gray_background(cell):
```

```python
# DELETE ENTIRE FUNCTION (lines 282-293):
# def is_carried_forward(cell_text: str):
```

**4. Search entire codebase**
```bash
grep -r "CARRIED FORWARD" amp_automation/
grep -r "carried.forward" amp_automation/
```

Remove all occurrences.

---

### 9. 🧩 Modularization - Configurable Scope

**Time Estimate:** 1 hour

#### 9.1. Configuration Structure

**File:** `config/master_config.json`

Add new section:
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
    }
  }
}
```

#### 9.2. Scope-Adaptive Wrapper Functions

**New File:** `amp_automation/presentation/indicators.py`

```python
"""
Brand-level indicator calculations with configurable scope.
"""

import logging
from typing import Literal
import pandas as pd

logger = logging.getLogger(__name__)

IndicatorScope = Literal["brand", "campaign", "slide"]

def calculate_quarter_totals(
    data: pd.DataFrame,
    scope: IndicatorScope,
    context: dict = None
) -> dict[str, float]:
    """
    Calculate quarter totals based on configured scope.

    Args:
        data: DataFrame with data
        scope: "brand", "campaign", or "slide"
        context: Additional context (e.g., campaign name for campaign scope)

    Returns:
        Dictionary with Q1-Q4 totals
    """
    if scope == "brand":
        return _calculate_brand_quarter_totals(data)
    elif scope == "campaign":
        return _calculate_campaign_quarter_totals(data, context.get("campaign_name"))
    elif scope == "slide":
        return _calculate_slide_quarter_totals(data)
    else:
        raise ValueError(f"Invalid scope: {scope}")

def calculate_media_share(
    data: pd.DataFrame,
    scope: IndicatorScope,
    context: dict = None
) -> dict[str, float]:
    """Calculate media share based on configured scope."""
    if scope == "brand":
        return _calculate_brand_media_share(data)
    elif scope == "campaign":
        return _calculate_campaign_media_share(data, context.get("campaign_name"))
    elif scope == "slide":
        return _calculate_slide_media_share(data)
    else:
        raise ValueError(f"Invalid scope: {scope}")

def calculate_funnel_share(
    data: pd.DataFrame,
    scope: IndicatorScope,
    context: dict = None
) -> dict[str, float]:
    """Calculate funnel stage share based on configured scope."""
    if scope == "brand":
        return _calculate_brand_funnel_share(data)
    elif scope == "campaign":
        return _calculate_campaign_funnel_share(data, context.get("campaign_name"))
    elif scope == "slide":
        return _calculate_slide_funnel_share(data)
    else:
        raise ValueError(f"Invalid scope: {scope}")

# Private implementation functions
def _calculate_brand_quarter_totals(data: pd.DataFrame) -> dict[str, float]:
    """Calculate quarters at brand level."""
    # Implementation from section 4.1
    pass

def _calculate_campaign_quarter_totals(data: pd.DataFrame, campaign_name: str) -> dict[str, float]:
    """Calculate quarters at campaign level."""
    campaign_data = data[data["Campaign Name"] == campaign_name]
    return _calculate_brand_quarter_totals(campaign_data)

def _calculate_slide_quarter_totals(data: pd.DataFrame) -> dict[str, float]:
    """Calculate quarters at slide level (data already filtered)."""
    return _calculate_brand_quarter_totals(data)

# Similar implementations for media_share and funnel_share...
```

#### 9.3. Usage in Main Code

```python
# In _create_presentation_for_combination():

# Load config
indicator_config = presentation_config.get("indicators", {})
quarter_scope = indicator_config.get("quarter_boxes", {}).get("scope", "brand")
media_scope = indicator_config.get("media_share", {}).get("scope", "brand")
funnel_scope = indicator_config.get("funnel_stage", {}).get("scope", "brand")

# Calculate indicators using configured scope
from amp_automation.presentation.indicators import (
    calculate_quarter_totals,
    calculate_media_share,
    calculate_funnel_share
)

brand_indicators = {
    "quarters": calculate_quarter_totals(subset, quarter_scope),
    "media_share": calculate_media_share(subset, media_scope),
    "funnel_stage": calculate_funnel_share(subset, funnel_scope),
}
```

**Benefit:** Can easily change scope in config without code changes:
```json
{
  "quarter_boxes": {
    "scope": "campaign"  // Change to campaign-level
  }
}
```

---

## Implementation Phases

### **Phase 1: Quick Fixes** (30 minutes)
1. ✅ Verify title (already correct)
2. ❌ Add GRP to MONTHLY TOTAL
3. 🗑️ Remove CARRIED FORWARD logic
4. 🔤 Fix campaign text wrapping

**Deliverable:** MONTHLY TOTAL shows GRP, no CARRIED FORWARD, text wraps correctly

---

### **Phase 2: BRAND TOTAL** (45 minutes)
1. Rename GRAND TOTAL → BRAND TOTAL
2. Add special styling (darker gray)
3. Ensure only on last slide
4. Add explanatory note

**Deliverable:** BRAND TOTAL appears only on last slide with distinct styling

---

### **Phase 3: Brand-Level Indicators** (2-3 hours)
1. 🟢 Implement quarter boxes calculation and display
2. 📊 Implement media share calculation and display
3. 🎯 Implement funnel stage calculation and display
4. Test with multi-slide brands

**Deliverable:** All indicators appear on last slide with correct brand-level totals

---

### **Phase 4: Modularization** (1 hour)
1. Create configuration structure
2. Implement scope-adaptive wrappers
3. Add validation and error handling
4. Document configuration options

**Deliverable:** Indicators can be reconfigured to different scopes via config

---

### **Phase 5: Testing & Documentation** (1 hour)
1. Generate test deck with multi-slide brands
2. Run post-processing pipeline
3. Visual verification of all changes
4. Update documentation

**Deliverable:** Fully tested, production-ready implementation

---

## Testing Checklist

### ✅ Single-Slide Brand
- [ ] Quarter boxes show correct brand totals
- [ ] Media share shows correct brand percentages
- [ ] Funnel stage shows correct brand percentages
- [ ] BRAND TOTAL appears at bottom
- [ ] MONTHLY TOTAL includes GRP values
- [ ] No CARRIED FORWARD rows

### ✅ Multi-Slide Brand (3 slides)
- [ ] Slides 1-2: NO quarter boxes, NO media share, NO funnel stage, NO BRAND TOTAL
- [ ] Slide 3 (last): ALL indicators present
- [ ] Quarter boxes show brand-level totals (all campaigns from all slides)
- [ ] Media share shows brand-level percentages
- [ ] Funnel stage shows brand-level percentages
- [ ] BRAND TOTAL sums ALL campaigns from all 3 slides
- [ ] BRAND TOTAL has darker gray styling
- [ ] Explanatory note appears below table

### ✅ Post-Processing
- [ ] Campaign cells merge vertically with word wrap
- [ ] MONTHLY TOTAL cells merge horizontally
- [ ] BRAND TOTAL cells merge horizontally
- [ ] Font normalization works correctly
- [ ] No errors in post-processing pipeline

### ✅ Edge Cases
- [ ] Brand with exactly 40 rows (no split)
- [ ] Brand with 41 rows (splits to 2 slides)
- [ ] Brand with 0 TV spend (GRP shows 0 or blank)
- [ ] Very long campaign names wrap correctly

---

## Template Shape Requirements

The template must have the following shapes on data slides:

### Required Shape Names

**Quarter Boxes:**
- `Q1_Box` - Q1 budget display
- `Q2_Box` - Q2 budget display
- `Q3_Box` - Q3 budget display
- `Q4_Box` - Q4 budget display

**Media Share Boxes:**
- `TV_Share_Box` - TV percentage
- `DIG_Share_Box` - Digital percentage
- `OTHER_Share_Box` - Other media percentage

**Funnel Stage Boxes:**
- `AWA_Box` - Awareness percentage
- `CON_Box` - Consideration percentage
- `PUR_Box` - Purchase percentage

**Optional:**
- `Brand_Total_Note` - Note text shape (can be created programmatically if missing)

### Verification Script

Before implementing, run this to check template:

```python
from pptx import Presentation

template_path = "template/Template_V4_FINAL_071025.pptx"
prs = Presentation(template_path)

required_shapes = [
    "Q1_Box", "Q2_Box", "Q3_Box", "Q4_Box",
    "TV_Share_Box", "DIG_Share_Box", "OTHER_Share_Box",
    "AWA_Box", "CON_Box", "PUR_Box"
]

# Check first data slide (usually index 1 or 2)
slide = prs.slides[1]  # Adjust index as needed

print("Checking template shapes...")
for shape_name in required_shapes:
    shape = next((s for s in slide.shapes if getattr(s, "name", "") == shape_name), None)
    if shape:
        print(f"✅ {shape_name}: Found")
    else:
        print(f"❌ {shape_name}: MISSING")
```

**Action if shapes missing:**
1. Update template to add shapes
2. OR create shapes programmatically
3. OR use existing shapes with different names (update config)

---

## Time & Resource Estimate

| Phase | Tasks | Time | Cumulative |
|-------|-------|------|------------|
| **Phase 1** | Quick fixes | 30 min | 30 min |
| **Phase 2** | BRAND TOTAL | 45 min | 1h 15min |
| **Phase 3** | Brand indicators | 2-3h | 3h 15min - 4h 15min |
| **Phase 4** | Modularization | 1h | 4h 15min - 5h 15min |
| **Phase 5** | Testing | 1h | 5h 15min - 6h 15min |

**Total Estimate:** 5-6 hours

**Recommended Approach:**
- Day 1: Phases 1-2 (2 hours) → Test & review
- Day 2: Phase 3 (2-3 hours) → Test & review
- Day 3: Phases 4-5 (2 hours) → Final testing & documentation

---

## Risk Assessment

### Low Risk ✅
- ❌ MONTHLY TOTAL GRP addition
- 🔤 Text wrapping fix
- 🗑️ CARRIED FORWARD removal

### Medium Risk ⚠️
- 🔄 BRAND TOTAL rename and styling
- 🟢 Quarter boxes implementation

### Higher Risk ⚠️⚠️
- 📊 Media share boxes (depends on data columns)
- 🎯 Funnel stage boxes (depends on data columns)
- 🧩 Modularization architecture

### Mitigation Strategies
1. **Incremental testing:** Test after each phase
2. **Feature branch:** Can revert if issues arise
3. **Data validation:** Check for required columns before calculations
4. **Graceful degradation:** If shapes missing, log warning but don't crash
5. **Backup plan:** Keep old GRAND TOTAL logic commented out until verified

---

## Success Criteria

Implementation is successful when:

1. ✅ **MONTHLY TOTAL** includes GRP column with correct campaign sums
2. ✅ **BRAND TOTAL** (renamed) appears only on last slide with special styling
3. ✅ **Quarter boxes** show brand-level totals on last slide only
4. ✅ **Media share boxes** show brand-level percentages on last slide only
5. ✅ **Funnel stage boxes** show brand-level percentages on last slide only
6. ✅ **Campaign text** wraps on full words only (no mid-word breaks)
7. ✅ **CARRIED FORWARD** logic completely removed
8. ✅ **All indicators** are modularized with configurable scope
9. ✅ **Multi-slide brands** display correctly (indicators only on last slide)
10. ✅ **Post-processing** completes without errors
11. ✅ **Visual inspection** confirms all changes look professional

---

## Questions for Approval

Before I start implementation, please confirm:

### 1. Template Shape Names
Do the quarter/media/funnel boxes already exist in the template?
- If YES: What are their shape names?
- If NO: Should I create them programmatically or should template be updated first?

### 2. BRAND TOTAL Styling
Which styling option do you prefer?
- **Option A:** Darker gray background (180, 180, 180)
- **Option B:** Bold text + thicker border
- **Option C:** Different color (light blue/yellow)

**Recommendation:** Option A (professional, subtle difference)

### 3. Missing Data Columns
What should happen if data columns are missing?
- `Mapped Media Type` (for media share)
- `Funnel Stage` (for funnel stage)

**Recommendation:** Log warning and skip that indicator, continue with others

### 4. Implementation Approach
How would you like to proceed?
- **Option A:** Implement all phases at once (5-6 hours), then test
- **Option B:** Implement Phase 1, test, get approval, then continue
- **Option C:** Implement Phases 1-2, test, get approval, then continue

**Recommendation:** Option C (reduces risk, allows for mid-implementation feedback)

---

## Approval Request

Please review this plan and provide:
1. ✅ **Approval to proceed** OR feedback/changes
2. 📋 **Answers to the 4 questions** above
3. 🎯 **Priority:** Should I start immediately or wait for template updates?

Once approved, I'll create the feature branch and begin implementation following this plan exactly.

---

**Prepared by:** Claude
**Date:** 27 October 2025
**Status:** AWAITING YOUR APPROVAL

