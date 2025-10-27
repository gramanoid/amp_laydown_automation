# MONTHLY TOTAL Fix Summary - 24 Oct 2025

## Issues Identified

### Issue #1: Wrong Font
- **Problem**: MONTHLY TOTAL rows displayed Calibri 18pt instead of Verdana 6pt
- **Root Cause**: Font constants set to 10pt, font name not explicitly set during merge
- **Fix**: Updated constants to 6pt, added explicit `FONT_NAME = "Verdana"` setting in merge styling

### Issue #2: Extra Text in Merged Cells
- **Problem**: MONTHLY TOTAL cells included campaign names (e.g., "MONTHLY TOTAL\nTELEVISION")
- **Root Cause**: Generation process stores multi-line text in single cell, merge preserved all lines
- **Fix**: Updated `normalize_label()` to extract only first line before newline character

## Code Changes

### File: `amp_automation/presentation/postprocess/cell_merges.py`

**Change 1 - Font Configuration:**
```python
# OLD
MONTHLY_TOTAL_FONT_SIZE = 10

# NEW
FONT_NAME = "Verdana"
MONTHLY_TOTAL_FONT_SIZE = 6  # Body text
```

**Change 2 - Always Set Font Name:**
```python
# In _apply_cell_styling()
for run in paragraph.runs:
    # Always set font name to Verdana
    run.font.name = FONT_NAME
    if font_size is not None:
        run.font.size = Pt(font_size)
```

**Change 3 - Extract First Line Only:**
```python
def normalize_label(text: str) -> str:
    if not text:
        return ""
    # Take only the first line (before any newline character)
    first_line = text.split('\n')[0].split('\r')[0]
    return first_line.strip().upper()
```

## Verification Results

### Before Fix:
- Column 1 text: `'MONTHLY TOTAL (£ 000)\nTELEVISION'`
- Font: Calibri 18pt
- Merged cells included campaign names

### After Fix:
- Column 1 text: `'MONTHLY TOTAL (£ 000)'` ✅
- Font: Verdana 6pt ✅
- Clean merge without extra text ✅

## Current Deck State

**File**: `output/presentations/run_20251024_172953/deck_test_fonts.pptx`

**Processing Stats:**
- 228 operations completed
- 0 failures
- ~190 MONTHLY TOTAL merges across 76 slides

**Merge Properties:**
- Horizontal merge: Columns 1-3
- Font: Verdana 6pt, bold
- Alignment: Center horizontal, middle vertical
- Text: "MONTHLY TOTAL (£ 000)" only (no campaign names)

## Workflow Applied

1. **Unmerge all cells** - Clean slate
2. **Normalize fonts** - Verdana 6pt body, 7pt header/bottom
3. **Merge MONTHLY TOTAL** - Horizontal merge with correct text and fonts

## Next Steps

MONTHLY TOTAL merges complete and verified. Ready for:
- GRAND TOTAL merges (bottom row, columns 1-3)
- CARRIED FORWARD merges (bottom row, columns 1-3)
- Campaign vertical merges (column 1 only)
