# LLM Fix Instructions
## Corrective Guidance for Slide 1

**Target:** GeneratedDeck_20251021_102357.pptx, Slide 1  
**Reference:** Template_V4_FINAL_071025.pptx, Slide 0

---

## STATUS OVERVIEW

✅ **Already Correct:**
- Slide dimensions (10.00" × 5.625")
- Table position (0.1789", 0.6979")
- Table structure (35 rows × 18 columns)
- Font family (Verdana)
- Core font sizes (7.5pt header, 7.0pt data)

❌ **Needs Fixing:**
1. Row heights (32 out of 35 rows wrong)
2. Text alignment (all 601 cells wrong)
3. Extra shapes (8 legend shapes to remove)
4. Column widths (minor sub-pixel differences)

---

## FIX #1: ROW HEIGHTS ⚠️ CRITICAL

**Issue:** 32 out of 35 rows have wrong heights. Generated uses content-based auto-sizing instead of fixed heights.

**Current Issues:**
- Header row (0): 127,101 EMUs (should be 161,729)
- Most data rows: 107,899 EMUs (should be 99,205)
- Some rows: 127,101 EMUs (should be 99,205)
- Last two data rows: ✅ Correct (99,205)
- Empty row (34): ✅ Correct (0)

**Fix:**
```python
# Set exact row heights from template
row_heights_emu = [
    161729,  # Row 0 (header)
    # Rows 1-33: ALL uniform at 99,205 EMUs
    99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205,
    99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205,
    99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205, 99205,
    99205, 99205, 99205, 99205, 99205, 99205,
    0,  # Row 34: zero height
]

for i, height in enumerate(row_heights_emu):
    table.rows[i].height = height
```

**Impact:** This fix alone will save 0.362" of vertical space.

---

## FIX #2: TEXT ALIGNMENT ⚠️ CRITICAL

**Issue:** Template uses CENTER for ALL cells. Generated uses LEFT for text columns, RIGHT for numeric columns.

**Current Distribution:**
- CENTER: 64 cells (should be 601)
- LEFT: 96 cells (should be 0)
- RIGHT: 416 cells (should be 0)

**Fix:**
```python
from pptx.enum.text import PP_ALIGN

# Apply CENTER to ALL cells
for row in range(35):
    for col in range(18):
        cell = table.cell(row, col)
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
```

**Critical Note:** This is a fundamental alignment philosophy difference. The template uses CENTER-all, not a mixed strategy.

---

## FIX #3: REMOVE EXTRA SHAPES ⚠️ HIGH PRIORITY

**Issue:** Generated has 20 shapes, template has 12. Shapes 12-19 are extra legend elements.

**Extra Shapes (to delete):**
- Shape 12: Rectangle (color box) at (6.4780", 0.3529")
- Shape 13: Text "TELEVISION" at (6.7867", 0.3130")
- Shape 14: Rectangle (color box) at (7.0880", 0.3529")
- Shape 15: Text "DIGITAL" at (7.3967", 0.3130")
- Shape 16: Rectangle (color box) at (8.0270", 0.3529")
- Shape 17: Text "OOH" at (8.3357", 0.3130")
- Shape 18: Rectangle (color box) at (8.7450", 0.3529")
- Shape 19: Text "OTHER" at (9.0537", 0.3130")

**Fix:**
```python
# Remove in reverse order to avoid index shifting
for idx in [19, 18, 17, 16, 15, 14, 13, 12]:
    sp = slide.shapes[idx]._element
    sp.getparent().remove(sp)
```

**Verification:**
```python
assert len(slide.shapes) == 12, "Should have exactly 12 shapes after removal"
```

---

## FIX #4: COLUMN WIDTHS ⚠️ MEDIUM PRIORITY

**Issue:** All 18 columns have minor differences (±1 to ±440 EMUs). Total difference: 252 EMUs (0.000276").

**Assessment:** Sub-pixel differences, but for pixel-perfect match, use exact template values.

**Fix:**
```python
column_widths_emu = [
    812364,  # Col 0
    729251,  # Col 1
    831384,  # Col 2
    338274,  # Col 3
    400567,  # Col 4
    400567,  # Col 5
    400567,  # Col 6
    414770,  # Col 7
    415954,  # Col 8
    465506,  # Col 9
    437595,  # Col 10
    443865,  # Col 11
    400567,  # Col 12
    437595,  # Col 13
    352043,  # Col 14
    449092,  # Col 15
    400567,  # Col 16
    400567,  # Col 17
]

for i, width in enumerate(column_widths_emu):
    table.columns[i].width = width
```

---

## FIX #5: FONT SIZES (MINOR)

**Issue:** Template has some 6.5pt text (39 cells), generated only uses 7.0pt and 7.5pt.

**Action Required:** Investigate which cells in template use 6.5pt and apply to generated.

**Note:** This is low priority as 6.5pt is rare and the difference is minor.

---

## FIX #6: SHAPE POSITIONS (MINOR)

**Issue:** Shapes 1-11 have minor position/size differences (sub-pixel, 0.0002-0.0005").

**Assessment:** These are negligible differences that won't affect visual appearance.

**Recommendation:** If pixel-perfect match required, apply exact template EMU values for all shape positions and sizes. Otherwise, leave as-is.

---

## EXECUTION SEQUENCE

Execute fixes in this order:

1. **Row heights** (saves vertical space, prevents other issues)
2. **Text alignment** (affects all cells, do before fine-tuning)
3. **Remove extra shapes** (cleans up slide)
4. **Column widths** (fine-tunes horizontal spacing)
5. **Font sizes** (only if 6.5pt cells identified)
6. **Shape positions** (only if pixel-perfect required)

---

## VERIFICATION SCRIPT

```python
from pptx import Presentation
from pptx.enum.text import PP_ALIGN

# Load presentation
prs = Presentation('GeneratedDeck_20251021_102357.pptx')
slide = prs.slides[1]  # Slide 1 (index 1)

# Get table
table_shape = [s for s in slide.shapes if hasattr(s, 'has_table') and s.has_table][0]
table = table_shape.table

# Verify row heights
assert table.rows[0].height == 161729, "Row 0 height wrong"
for i in range(1, 34):
    assert table.rows[i].height == 99205, f"Row {i} height wrong"
assert table.rows[34].height == 0, "Row 34 height wrong"
print("✅ Row heights correct")

# Verify alignment
center_count = 0
for row in range(35):
    for col in range(18):
        cell = table.cell(row, col)
        for para in cell.text_frame.paragraphs:
            if para.alignment == PP_ALIGN.CENTER:
                center_count += 1
assert center_count == 601, f"Should have 601 CENTER cells, found {center_count}"
print("✅ Text alignment correct")

# Verify shape count
assert len(slide.shapes) == 12, f"Should have 12 shapes, found {len(slide.shapes)}"
print("✅ Shape count correct")

# Verify column widths
expected_widths = [
    812364, 729251, 831384, 338274, 400567, 400567, 400567, 414770, 415954,
    465506, 437595, 443865, 400567, 437595, 352043, 449092, 400567, 400567
]
for i, expected_width in enumerate(expected_widths):
    assert table.columns[i].width == expected_width, f"Column {i} width wrong"
print("✅ Column widths correct")

print("\n✅✅✅ ALL VERIFICATIONS PASSED ✅✅✅")
```

---

## KEY DIFFERENCES FROM PREVIOUS ANALYSIS

**Previous (13.33" × 7.50" slide):**
- Wrong slide dimensions (root cause)
- 33.3% scaling error throughout
- Table overflow (3.94")

**Current (10.00" × 5.625" slide):**
- ✅ Correct slide dimensions
- ✅ No overflow
- ⚠️ Row height calculation logic issue
- ⚠️ Alignment strategy difference
- ⚠️ Extra legend shapes

The current version is **much closer** to the template. Main issues are row heights and alignment, not fundamental dimensional errors.

---

## ESTIMATED FIX TIME

- **Automated script:** 5-10 minutes
- **Manual fixes:** 30-60 minutes

---

## CONFIDENCE LEVEL

- Row heights: 100% - exact EMU values provided
- Text alignment: 100% - clear requirement (CENTER all)
- Extra shapes: 100% - indices 12-19 confirmed extra
- Column widths: 100% - exact EMU values provided

All measurements extracted from XML and verified with python-pptx API.
