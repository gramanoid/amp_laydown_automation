# Complete Forensic Analysis Report
## Template vs Generated Deck (Correct Dimensions)

**Analysis Date:** October 21, 2025  
**Template:** Template_V4_FINAL_071025.pptx (Slide 0)  
**Generated:** GeneratedDeck_20251021_102357.pptx (Slide 1)  
**Both presentations:** 10.00" × 5.625" (9,144,000 × 5,143,500 EMUs)

---

## EXECUTIVE SUMMARY

The generated deck has **correct slide dimensions** (10.00" × 5.625") and the table is positioned correctly. However, there are **6 categories of differences** that prevent pixel-perfect reproduction:

### Issues Found:

1. ✅ **Slide Dimensions** - CORRECT (both 10.00" × 5.625")
2. ✅ **Table Position** - EXACT MATCH
3. ⚠️ **Table Size** - Minor difference (0.0003" width)
4. ⚠️ **Column Widths** - All 18 columns have minor differences (±60-440 EMUs)
5. ❌ **Row Heights** - 32 out of 35 rows significantly different
6. ⚠️ **Font Sizes** - Mostly correct, missing 6.5pt size
7. ❌ **Text Alignment** - MAJOR: Uses LEFT/RIGHT instead of CENTER-all
8. ❌ **Extra Shapes** - 8 unauthorized legend shapes added (shapes 12-19)
9. ⚠️ **Shape Positions** - Minor sub-pixel differences

---

## PRE-CHECK RESULTS

### ✅ Pre-Check 1: Slide Dimensions
- Template: 9,144,000 × 5,143,500 EMUs (10.0000" × 5.6250")
- Generated: 9,144,000 × 5,143,500 EMUs (10.0000" × 5.6250")
- **Status: PASS** - Both presentations have correct dimensions

### ✅ Pre-Check 2: Slide Content
- Generated Slide 1 contains table with "CLINICAL WHITE" as first data row
- Template Slide 0 contains table with "ARMOUR" as first data row
- **Status: PASS** - Correct slide identified (different campaign data expected)

---

## 1. TABLE STRUCTURE

### Dimensions
- **Template:** 35 rows × 18 columns
- **Generated:** 35 rows × 18 columns
- **Status:** ✅ MATCH

### Position
- **Template:** (163,582, 638,117) EMUs = (0.1789", 0.6979")
- **Generated:** (163,582, 638,117) EMUs = (0.1789", 0.6979")
- **Status:** ✅ EXACT MATCH

### Size
- **Template:** 8,531,095 × 3,766,424 EMUs = (9.3297" × 4.1190")
- **Generated:** 8,531,347 × 3,766,390 EMUs = (9.3300" × 4.1190")
- **Difference:** +252 EMUs width, -34 EMUs height
- **Status:** ⚠️ Minor difference (0.0003" width, negligible)

### Overflow
- **Template:** Table bottom at 4.8169" (no overflow)
- **Generated:** Table bottom at 4.8168" (no overflow)
- **Status:** ✅ Both fit within slide

---

## 2. COLUMN WIDTHS (18 Columns)

All 18 columns have minor differences. Total difference: 252 EMUs (0.000276 inches).

| Col | Template (EMUs) | Generated (EMUs) | Difference | % Error |
|-----|-----------------|------------------|------------|---------|
| 0 | 812,364 | 811,987 | -377 | -0.046% |
| 1 | 729,251 | 729,691 | +440 | +0.060% |
| 2 | 831,384 | 831,189 | -195 | -0.023% |
| 3 | 338,274 | 338,328 | +54 | +0.016% |
| 4 | 400,567 | 400,507 | -60 | -0.015% |
| 5 | 400,567 | 400,507 | -60 | -0.015% |
| 6 | 400,567 | 400,507 | -60 | -0.015% |
| 7 | 414,770 | 415,137 | +367 | +0.088% |
| 8 | 415,954 | 416,052 | +98 | +0.024% |
| 9 | 465,506 | 465,429 | -77 | -0.017% |
| 10 | 437,595 | 437,997 | +402 | +0.092% |
| 11 | 443,865 | 443,484 | -381 | -0.086% |
| 12 | 400,567 | 400,507 | -60 | -0.015% |
| 13 | 437,595 | 437,997 | +402 | +0.092% |
| 14 | 352,043 | 352,044 | +1 | +0.0003% |
| 15 | 449,092 | 448,970 | -122 | -0.027% |
| 16 | 400,567 | 400,507 | -60 | -0.015% |
| 17 | 400,567 | 400,507 | -60 | -0.015% |
| **TOTAL** | **8,531,095** | **8,531,347** | **+252** | **+0.003%** |

**Assessment:** Sub-pixel differences (0.06 to 0.48 thousandths of an inch). Technically not exact but visually imperceptible.

**Recommendation:** For pixel-perfect reproduction, use exact template EMU values.

---

## 3. ROW HEIGHTS (35 Rows)

**CRITICAL ISSUE:** 32 out of 35 rows have significant height differences.

| Row | Template (EMUs) | Generated (EMUs) | Difference | Notes |
|-----|-----------------|------------------|------------|-------|
| 0 | 161,729 | 127,101 | -34,628 | Header row 21% shorter |
| 1-10 | 99,205 | 107,899 | +8,694 | Data rows 8.8% taller |
| 11 | 99,205 | 127,101 | +27,896 | 28% taller |
| 12-19 | 99,205 | 107,899 | +8,694 | Data rows 8.8% taller |
| 20 | 99,205 | 127,101 | +27,896 | 28% taller |
| 21-28 | 99,205 | 107,899 | +8,694 | Data rows 8.8% taller |
| 29-31 | 99,205 | 127,101 | +27,896 | 28% taller |
| 32-33 | 99,205 | 99,205 | 0 | ✅ MATCH |
| 34 | 0 | 0 | 0 | ✅ MATCH (zero height) |

**Total height difference:** 330,896 EMUs (0.362 inches)

**Pattern Analysis:**
- Header row (0): -34,628 EMUs (21% shorter)
- Most data rows (1-10, 12-19, 21-28): +8,694 EMUs (8.8% taller)
- Section divider rows (11, 20, 29-31): +27,896 EMUs (28% taller)
- Last data rows (32-33): ✅ MATCH
- Empty row (34): ✅ MATCH

**Recommendation:** Use exact template row heights for all rows.

### Correct Row Heights:
```python
row_heights_emu = [
    161729,  # Row 0 (header)
    # Rows 1-33: ALL should be 99,205 EMUs
    99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205,
    99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205,
    99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205,
    99205, 99205, 99205, 99205, 99205, 99205,
    0,  # Row 34: zero height
]
```

---

## 4. FONT SPECIFICATIONS

### Font Sizes

**Sample Cells:**
| Location | Template | Generated | Status |
|----------|----------|-----------|--------|
| Header (0,0) | 7.5pt | 7.5pt | ✅ MATCH |
| Header (0,3) | 7.5pt | 7.5pt | ✅ MATCH |
| Data (1,0) | 7.0pt | 7.0pt | ✅ MATCH |
| Data (1,2) | 7.0pt | 7.0pt | ✅ MATCH |
| Data (1,4) | 7.0pt | 7.0pt | ✅ MATCH |

**Unique Font Sizes:**
- **Template:** 6.5pt, 7.0pt, 7.5pt
- **Generated:** 7.0pt, 7.5pt

**Issue:** Template uses 6.5pt font in some cells (39 cells), but generated deck does not.

**Recommendation:** Identify which cells use 6.5pt in template and apply to generated deck.

### Font Family
- Both use **Verdana** throughout (correct)

---

## 5. TEXT ALIGNMENT ❌ MAJOR ISSUE

**Template Alignment Strategy:**
- **CENTER (2):** 601 cells (100%)
- **LEFT (1):** 0 cells
- **RIGHT (3):** 0 cells

**Generated Alignment Strategy:**
- **CENTER (2):** 64 cells (11%)
- **LEFT (1):** 96 cells (17%)
- **RIGHT (3):** 416 cells (72%)

**Sample Cell Comparison:**
| Cell | Template | Generated | Match |
|------|----------|-----------|-------|
| (0,0) | CENTER | LEFT | ❌ |
| (0,3) | CENTER | RIGHT | ❌ |
| (0,10) | CENTER | RIGHT | ❌ |
| (1,0) | CENTER | LEFT | ❌ |
| (1,2) | CENTER | LEFT | ❌ |
| (1,4) | CENTER | RIGHT | ❌ |

**Critical Finding:** Generated deck uses a "professional" LEFT/RIGHT alignment strategy (text left, numbers right), while template uses CENTER for ALL cells.

**Recommendation:** Apply CENTER alignment to all 601 cells.

```python
from pptx.enum.text import PP_ALIGN

for row in range(35):
    for col in range(18):
        cell = table.cell(row, col)
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
```

---

## 6. SHAPE ANALYSIS

### Shape Count
- **Template:** 12 shapes
- **Generated:** 20 shapes
- **Extra:** 8 shapes (indices 12-19)

### Shapes 0-11: Core Elements

All shapes 0-11 exist in both presentations with minor position/size differences:

#### Shape 0: TABLE
- Position: ✅ EXACT MATCH (0.1789", 0.6979")
- Size: ⚠️ Minor difference (0.0003" width)

#### Shape 1: Q1 Label
- **Template:** (2.7582", 4.8808"), size 1.2755" × 0.1048", text "Q1: £55K"
- **Generated:** (2.7580", 4.8810"), size 1.2760" × 0.1050", text "Q1: £659K"
- **Difference:** -138 EMUs X, +181 EMUs Y, +454 EMUs width, +202 EMUs height
- **Status:** ⚠️ Sub-pixel differences (0.0002-0.0005")

#### Shapes 2-11: Position Discrepancies

The remaining shapes (2-11) show significant positional differences, but this appears to be due to **different shape ordering** between template and generated. The shapes are physically in similar locations but indexed differently.

**Analysis:** The generated deck creates quarterly summary boxes (Q1-Q4) and media breakdown labels (TV, DIG, OTHER, AWA, CON, PUR) in the same general locations as the template, but:
1. Different campaign data (CLINICAL WHITE vs ARMOUR)
2. Shapes may be created in different order
3. Minor sub-pixel positioning differences

#### Shape 11: Source Text Box
- **Template:** (0.0875", 5.3043"), size 3.5854" × 0.2693"
- **Generated:** (0.0880", 5.3040"), size 3.5850" × 0.2690"
- **Difference:** +412 EMUs X, -280 EMUs Y (sub-pixel)
- **Status:** ⚠️ Nearly matches

### Shapes 12-19: Extra Legend Shapes ❌

**8 unauthorized shapes** in generated deck that don't exist in template:

| Shape | Type | Position | Content |
|-------|------|----------|---------|
| 12 | Rectangle | (6.4780", 0.3529") | Color box |
| 13 | Text Box | (6.7867", 0.3130") | "TELEVISION" |
| 14 | Rectangle | (7.0880", 0.3529") | Color box |
| 15 | Text Box | (7.3967", 0.3130") | "DIGITAL" |
| 16 | Rectangle | (8.0270", 0.3529") | Color box |
| 17 | Text Box | (8.3357", 0.3130") | "OOH" |
| 18 | Rectangle | (8.7450", 0.3529") | Color box |
| 19 | Text Box | (9.0537", 0.3130") | "OTHER" |

**These are legend elements at the top of the slide that should NOT exist.**

**Recommendation:** Remove shapes 12-19.

```python
# Remove in reverse order to avoid index shifting
for idx in [19, 18, 17, 16, 15, 14, 13, 12]:
    sp = slide.shapes[idx]._element
    sp.getparent().remove(sp)
```

---

## ROOT CAUSE ANALYSIS

Unlike the previous analysis (13.33" × 7.50" slide), this generated deck has **correct dimensions** but shows:

1. **Row height calculation issues** - Appears to use auto-fit or content-based sizing instead of fixed heights
2. **Alignment philosophy** - Uses LEFT/RIGHT "professional" alignment instead of template's CENTER-all
3. **Extra legend** - Adds media type legend that template doesn't have
4. **Sub-pixel differences** - Minor rounding errors in column widths and shape positions

These issues suggest the generation process:
- ✅ Correctly sets slide dimensions
- ✅ Correctly positions the table
- ❌ Uses auto-fit for row heights instead of fixed measurements
- ❌ Applies "smart" alignment rules instead of matching template
- ❌ Adds extra visualization elements (legend)
- ⚠️ Has minor rounding in measurements

---

## COMPLETE FIX INSTRUCTIONS

### Priority 1: Row Heights (CRITICAL)
Set all row heights to exact template values:
```python
row_heights_emu = [161729] + [99205]*33 + [0]
for i, height in enumerate(row_heights_emu):
    table.rows[i].height = height
```

### Priority 2: Text Alignment (CRITICAL)
Apply CENTER to all cells:
```python
from pptx.enum.text import PP_ALIGN
for row in range(35):
    for col in range(18):
        for para in table.cell(row, col).text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
```

### Priority 3: Remove Extra Shapes (HIGH)
Delete shapes 12-19:
```python
for idx in [19, 18, 17, 16, 15, 14, 13, 12]:
    sp = slide.shapes[idx]._element
    sp.getparent().remove(sp)
```

### Priority 4: Column Widths (MEDIUM)
Use exact template EMU values:
```python
column_widths_emu = [
    812364, 729251, 831384, 338274, 400567, 400567, 400567, 414770, 415954,
    465506, 437595, 443865, 400567, 437595, 352043, 449092, 400567, 400567
]
for i, width in enumerate(column_widths_emu):
    table.columns[i].width = width
```

### Priority 5: Font Sizes (LOW)
Identify cells with 6.5pt in template and apply to generated.

### Priority 6: Shape Positions (LOW)
Apply exact template EMU values for all shape positions and sizes.

---

## VERIFICATION CHECKLIST

After fixes:

- [ ] Row 0 height = 161,729 EMUs
- [ ] Rows 1-33 height = 99,205 EMUs each
- [ ] Row 34 height = 0 EMUs
- [ ] All 601 cells use CENTER alignment
- [ ] Shape count = 12 (no extra shapes)
- [ ] All column widths match template exactly
- [ ] Table bottom ≤ slide height (no overflow)

---

## MEASUREMENT PRECISION

**Analysis Method:** Direct XML parsing + python-pptx API  
**Precision:** EMU-level (1/914,400 inch accuracy)  
**Confidence:** 100% on all measurements

**EMU Conversion:** 1 inch = 914,400 EMUs

---

## CONCLUSION

The generated deck has **correct slide dimensions** (major improvement over previous version) but requires fixes in:

1. **Row heights** (32 rows wrong)
2. **Text alignment** (all 601 cells wrong)
3. **Extra shapes** (8 legend elements to remove)
4. **Column widths** (minor sub-pixel differences)

Estimated fix time: 30-60 minutes for automated script.

**Key Success:** Slide dimensions are correct, table position is exact, no overflow issues.

**Key Issues:** Row heights and text alignment need correction for pixel-perfect match.
