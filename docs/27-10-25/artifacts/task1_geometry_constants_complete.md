# Task 1: Template V4 Geometry Constants - Complete

**Status:** ✅ COMPLETE (Pre-existing)
**Completed:** Already implemented (found during session)
**Time:** 0.5h (verification only)

---

## Summary

Template V4 geometry constants were **already captured and implemented** in `amp_automation/presentation/template_geometry.py`. During verification, confirmed all constants match template measurements exactly.

---

## Constants Verified

### Column Widths (18 columns)
```python
TEMPLATE_V4_COLUMN_WIDTHS_EMU = [
    812364,  # Column 0 (Campaign) - 0.888 in
    729251,  # Column 1 (Media/Month) - 0.798 in
    831384,  # Column 2 (Metric) - 0.909 in
    338274,  # Column 3 (Total) - 0.370 in
    400567,  # Column 4 (Q1) - 0.438 in
    # ... (14 more columns)
]
```

### Row Heights
```python
TEMPLATE_V4_ROW_HEIGHT_HEADER_EMU = 161729  # 0.177 in (12.7 pt)
TEMPLATE_V4_ROW_HEIGHT_BODY_EMU = 99205     # 0.108 in (7.8 pt)
TEMPLATE_V4_ROW_HEIGHT_TRAILER_EMU = 0      # GRAND TOTAL row
```

### Table Position & Size
```python
TEMPLATE_V4_TABLE_LEFT_EMU = 163582      # 0.18 in from left
TEMPLATE_V4_TABLE_TOP_EMU = 638117       # 0.70 in from top
TEMPLATE_V4_TABLE_WIDTH_EMU = 8531095    # 9.33 in wide
TEMPLATE_V4_TABLE_HEIGHT_EMU = 3766424   # 4.12 in tall
```

---

## Verification Results

✅ **Extracted from template:** All measurements match `Template_V4_FINAL_071025.pptx`
✅ **Module exists:** `amp_automation/presentation/template_geometry.py`
✅ **Properly imported:** Used in `assembly.py` (lines 698-711)
✅ **Applied correctly:** Row heights, column widths, table bounds all reference constants
✅ **Includes helper classes:** `TemplateTableBounds` dataclass for convenience

---

## Usage in Code

The constants are actively used in `assembly.py`:

1. **Column widths** (line 1452, 1660):
   ```python
   column_widths_source = (
       column_widths_config if column_widths_config
       else TEMPLATE_V4_COLUMN_WIDTHS_INCHES
   )
   ```

2. **Row heights** (lines 1421, 1429, 1437, 1444):
   ```python
   header_inches = float(
       row_heights_config.get("header_inches",
                            TEMPLATE_V4_ROW_HEIGHT_HEADER_INCHES)
   )
   ```

3. **Table bounds** (lines 2838-2841):
   ```python
   table_shape.left = Inches(TEMPLATE_V4_TABLE_BOUNDS.left)
   table_shape.top = Inches(TEMPLATE_V4_TABLE_BOUNDS.top)
   table_shape.width = Inches(TEMPLATE_V4_TABLE_BOUNDS.width)
   table_shape.height = Inches(TEMPLATE_V4_TABLE_BOUNDS.height)
   ```

---

## Files Verified

- ✅ `amp_automation/presentation/template_geometry.py` - Constants module
- ✅ `amp_automation/presentation/assembly.py` - Imports and uses constants
- ✅ `template/Template_V4_FINAL_071025.pptx` - Source template

---

## Next Steps

✅ **Task 1 Complete** - Constants captured and verified
⏭️ **Task 2 Next** - Update continuation slide layout to use these constants consistently

---

## Notes

- Constants were implemented previously (not during this session)
- All measurements verified against template using python-pptx
- 18 columns, 35 rows (header + 33 body + 1 GRAND TOTAL)
- Header row: 161729 EMU (0.177 in / 12.7 pt)
- Body rows: 99205 EMU (0.108 in / 7.8 pt)
- Conversion: 1 inch = 914400 EMU = 72 points
