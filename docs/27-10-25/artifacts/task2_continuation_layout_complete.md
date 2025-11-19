# Task 2: Update Continuation Slide Layout Logic - Complete

**Status:** ✅ COMPLETE (Pre-existing)
**Completed:** Already implemented (found during verification)
**Time:** 0.5h (verification only)

---

## Summary

Continuation slide layout logic **already honors Template V4 geometry** through the `_populate_cloned_table()` function. All continuation slides use the same geometry constants as the template, ensuring pixel-perfect alignment.

---

## Code Verification

### Template Geometry Applied to All Tables

The `_populate_cloned_table()` function (assembly.py:2736-2844) applies template geometry to every table, including continuation slides:

**Column Widths** (lines 2756-2760):
```python
for col_idx, width in enumerate(TABLE_COLUMN_WIDTHS[: len(table.columns)]):
    try:
        table.columns[col_idx].width = width
    except Exception as exc:
        logger.debug("Unable to enforce column width for column %s: %s", col_idx, exc)
```

**Row Heights** (lines 2762-2801):
```python
header_height = TABLE_ROW_HEIGHT_HEADER
body_height = TABLE_ROW_HEIGHT_BODY
subtotal_height = TABLE_ROW_HEIGHT_SUBTOTAL
trailer_height = TABLE_ROW_HEIGHT_TRAILER

# Applied per row type:
if row_idx == 0:
    _apply_row_height(row, header_height, row_idx, lock_exact=True)
elif label in subtotal_labels:
    _apply_row_height(row, subtotal_height, row_idx, lock_exact=True)
else:
    _apply_row_height(row, body_height, row_idx, lock_exact=True)
```

**Table Position & Size** (lines 2838-2841):
```python
table_shape.left = Inches(TEMPLATE_V4_TABLE_BOUNDS.left)
table_shape.top = Inches(TEMPLATE_V4_TABLE_BOUNDS.top)
table_shape.width = Inches(TEMPLATE_V4_TABLE_BOUNDS.width)
table_shape.height = Inches(TEMPLATE_V4_TABLE_BOUNDS.height)
```

---

## Constants Initialization

The geometry constants are properly initialized from template_geometry module:

**Default Values** (assembly.py:1421, 1429, 1452):
```python
# Row heights use template constants as defaults
header_inches = float(
    row_heights_config.get("header_inches", TEMPLATE_V4_ROW_HEIGHT_HEADER_INCHES)
)
TABLE_ROW_HEIGHT_HEADER = Inches(header_inches)

body_inches = float(
    row_heights_config.get("body_inches", TEMPLATE_V4_ROW_HEIGHT_BODY_INCHES)
)
TABLE_ROW_HEIGHT_BODY = Inches(body_inches)

# Column widths use template constants as defaults
column_widths_source = (
    column_widths_config if column_widths_config
    else TEMPLATE_V4_COLUMN_WIDTHS_INCHES
)
TABLE_COLUMN_WIDTHS = [Inches(float(width)) for width in column_widths_source]
```

---

## Continuation Slide Flow

1. **Split Logic** (`_split_table_data_by_campaigns`, line 2251):
   - Splits data respecting MAX_ROWS_PER_SLIDE (32 rows)
   - Adds CARRIED FORWARD row when needed (lines 2386-2396)
   - Adds per-slide GRAND TOTAL (lines 2407-2416)

2. **Table Creation** (`_add_and_style_table`, line 2847):
   - Clones template table shape for each slide
   - Calls `_populate_cloned_table()` for all slides (first and continuation)

3. **Geometry Application** (`_populate_cloned_table`, line 2736):
   - **Same function for all slides** - no special case for continuations
   - Applies column widths from `TABLE_COLUMN_WIDTHS`
   - Applies row heights based on row type
   - Snaps table to `TEMPLATE_V4_TABLE_BOUNDS`

---

## Verification Results

✅ **Continuation slides use same table creation function as first slide**
✅ **Column widths applied from TEMPLATE_V4_COLUMN_WIDTHS**
✅ **Row heights applied from TEMPLATE_V4_ROW_HEIGHT constants**
✅ **Table position/size snapped to TEMPLATE_V4_TABLE_BOUNDS**
✅ **No hardcoded geometry values in continuation logic**
✅ **Config-driven with template constants as defaults**

---

## Files Verified

- ✅ `amp_automation/presentation/assembly.py` - All table creation uses template constants
- ✅ `amp_automation/presentation/template_geometry.py` - Constants module
- ✅ Split logic (lines 2251-2450) - Data splitting respects constraints
- ✅ Table population (lines 2736-2844) - Applies geometry to all tables
- ✅ Table creation (lines 2847-2865) - Clones template and populates

---

## Special Row Types

**CARRIED FORWARD row** (lines 2386-2396):
- Created dynamically during splits
- Uses `subtotal_height` (same as BODY_ROW_HEIGHT)
- Formatted like MONTHLY TOTAL rows

**GRAND TOTAL row** (lines 2407-2416):
- Created for each slide (not just final slide)
- Uses `trailer_height` (same as BODY_ROW_HEIGHT)
- Last row on every slide

Both special rows use the same template geometry constants, ensuring consistency across all slide types.

---

## Next Steps

✅ **Task 1 Complete** - Geometry constants captured
✅ **Task 2 Complete** - Continuation layout uses template geometry
⏭️ **Task 3 Next** - Run visual_diff.py to validate geometry in generated deck

---

## Notes

- Continuation slides handled identically to first slide (no special geometry code)
- All geometry configuration-driven with template constants as fallback
- `lock_exact=True` on row heights ensures they don't auto-resize
- Table borders applied consistently (line 2836)
- Unused rows removed from cloned templates (lines 2830-2834)
