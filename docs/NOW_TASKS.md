# NOW Tasks - Priority Work Items

## 2025-10-27 - Campaign Cell Text Wrapping Issue

**PRIORITY: HIGH**

### Problem
Campaign names in column A are still breaking mid-word despite smart line breaking implementation:
- "FACES-CONDITION" displays as "FACES-CONDITIO\nN" (breaks mid-word)
- Should display as "FACES\nCONDITION" (two lines, clean break)

### What We've Tried
1. ✅ Implemented `_smart_line_break()` function in `cell_merges.py`
   - Replaces dashes with spaces
   - Intelligently splits words: 2 words = 1 per line, 3 words = 2+1, etc.
   - Function CONFIRMED WORKING via debug output

2. ✅ Applied smart breaks during table generation in `assembly.py:668`
   - Text is correctly formatted with `\n` characters
   - Debug confirmed: "FACES-CONDITION" → "FACES\nCONDITION"

3. ✅ Set font to 6pt Verdana, bold, centered
4. ✅ Applied vertical cell merging

### Root Cause
PowerPoint is overriding our `\n` line breaks with its own word wrapping because:
- Cell width is too narrow for campaign text
- PowerPoint's auto word-wrap is breaking words mid-character
- Our explicit line breaks are being ignored

### Potential Solutions to Try Tomorrow

#### Option 1: Increase Campaign Column Width
- Widen column A to accommodate longer campaign names
- File: `assembly.py` - find column width settings
- May need to adjust overall table layout

#### Option 2: Disable Word Wrap + Shrink to Fit
- Set `text_frame.word_wrap = False`
- Set `text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`
- File: `assembly.py` or `cell_merges.py`

#### Option 3: Force Text Box Behavior
- Investigate if we can set fixed text box properties that respect `\n`
- May need to set different text properties during merge

#### Option 4: Reduce Font Size Conditionally
- For long campaign names (>15 chars), use 5pt instead of 6pt
- User previously rejected 5pt globally, but might accept conditional

### Files Involved
- `amp_automation/presentation/assembly.py` (line 668)
- `amp_automation/presentation/postprocess/cell_merges.py` (_smart_line_break)
- Reference screenshot: Shows "FACES-CONDITIO\nN" breaking mid-word

### Current Workaround
None - issue remains unfixed

### Test Case
Generate deck and check slide "RSA - SENSODYNE (25)":
- CLINICAL WHITE ✓ (should be 2 lines)
- DUOFLEX BODYGUARD ✓ (should be 2 lines)
- FACES-CONDITION ✗ (currently breaks as "FACES-CONDITIO\nN")
- FEEL FAMILIAR ✓ (should be 2 lines)

---

## Completed Today (2025-10-27)
- ✅ Fixed timestamp to use local system time (Arabian Standard Time UTC+4)
- ✅ Implemented smart line breaking function
- ✅ Added media channel vertical merging
- ✅ Corrected font sizes: 6pt body, 7pt BRAND TOTAL
- ✅ Removed debug print statements
