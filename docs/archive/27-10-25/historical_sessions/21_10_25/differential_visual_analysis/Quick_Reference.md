# Quick Reference Guide
## Template vs Generated Deck Comparison

**Files:** Template_V4_FINAL_071025.pptx (Slide 0) vs GeneratedDeck_20251021_102357.pptx (Slide 1)

---

## ✅ WHAT'S CORRECT

- Slide dimensions: 10.00" × 5.625" (9,144,000 × 5,143,500 EMUs)
- Table position: (163,582, 638,117) EMUs = (0.1789", 0.6979")
- Table structure: 35 rows × 18 columns
- No overflow (table fits within slide)
- Font family: Verdana
- Primary font sizes: 7.5pt (header), 7.0pt (data)

---

## ❌ WHAT NEEDS FIXING

### 1. Row Heights (32 out of 35 wrong)

**Quick Fix:**
```python
row_heights_emu = [161729] + [99205]*33 + [0]
for i, h in enumerate(row_heights_emu):
    table.rows[i].height = h
```

### 2. Text Alignment (all 601 cells wrong)

**Quick Fix:**
```python
from pptx.enum.text import PP_ALIGN
for r in range(35):
    for c in range(18):
        for p in table.cell(r, c).text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
```

### 3. Extra Shapes (8 legend shapes)

**Quick Fix:**
```python
for idx in [19, 18, 17, 16, 15, 14, 13, 12]:
    slide.shapes[idx]._element.getparent().remove(slide.shapes[idx]._element)
```

### 4. Column Widths (minor differences)

**Quick Fix:**
```python
column_widths_emu = [
    812364, 729251, 831384, 338274, 400567, 400567, 400567, 414770, 415954,
    465506, 437595, 443865, 400567, 437595, 352043, 449092, 400567, 400567
]
for i, w in enumerate(column_widths_emu):
    table.columns[i].width = w
```

---

## EXACT ROW HEIGHTS (35 rows)

| Row | EMUs | Inches |
|-----|------|--------|
| 0 | 161,729 | 0.1768" |
| 1-33 | 99,205 | 0.1085" |
| 34 | 0 | 0.0000" |

**Total:** 3,435,494 EMUs (3.7571")

---

## EXACT COLUMN WIDTHS (18 columns)

| Col | EMUs | Inches |
|-----|------|--------|
| 0 | 812,364 | 0.8885" |
| 1 | 729,251 | 0.7975" |
| 2 | 831,384 | 0.9092" |
| 3 | 338,274 | 0.3699" |
| 4-6,12,16-17 | 400,567 | 0.4381" |
| 7 | 414,770 | 0.4536" |
| 8 | 415,954 | 0.4549" |
| 9 | 465,506 | 0.5091" |
| 10,13 | 437,595 | 0.4786" |
| 11 | 443,865 | 0.4854" |
| 14 | 352,043 | 0.3850" |
| 15 | 449,092 | 0.4911" |

**Total:** 8,531,095 EMUs (9.3297")

---

## ISSUE SUMMARY

| Issue | Severity | Count | Fix Time |
|-------|----------|-------|----------|
| Row heights | CRITICAL | 32 rows | 2 min |
| Text alignment | CRITICAL | 601 cells | 3 min |
| Extra shapes | HIGH | 8 shapes | 1 min |
| Column widths | MEDIUM | 18 columns | 2 min |
| Font sizes | LOW | ~39 cells | 5 min |

**Total Fix Time:** ~15 minutes (automated)

---

## ONE-LINE FIXES

```python
# Fix row heights
[table.rows[i].__setattr__('height', [161729] + [99205]*33 + [0][i]) for i in range(35)]

# Fix alignment (all cells to CENTER)
[p.__setattr__('alignment', PP_ALIGN.CENTER) for r in range(35) for c in range(18) for p in table.cell(r,c).text_frame.paragraphs]

# Remove extra shapes (12-19)
[slide.shapes[i]._element.getparent().remove(slide.shapes[i]._element) for i in [19,18,17,16,15,14,13,12]]

# Fix column widths
[table.columns[i].__setattr__('width', w) for i, w in enumerate([812364,729251,831384,338274,400567,400567,400567,414770,415954,465506,437595,443865,400567,437595,352043,449092,400567,400567])]
```

---

## VERIFICATION

```python
# After fixes
assert table.rows[0].height == 161729
assert all(table.rows[i].height == 99205 for i in range(1, 34))
assert table.rows[34].height == 0
assert sum(1 for r in range(35) for c in range(18) for p in table.cell(r,c).text_frame.paragraphs if p.alignment == PP_ALIGN.CENTER) == 601
assert len(slide.shapes) == 12
print("✅ All fixes verified")
```

---

## EMU CONVERSION

1 inch = 914,400 EMUs

```python
inches_to_emu = lambda inches: int(inches * 914400)
emu_to_inches = lambda emu: emu / 914400
```

---

## KEY INSIGHT

Unlike the previous 13.33" × 7.50" slide, this version has **correct dimensions** but wrong:
1. **Row sizing logic** (uses auto-fit instead of fixed)
2. **Alignment philosophy** (LEFT/RIGHT instead of CENTER-all)
3. **Extra elements** (legend that shouldn't exist)

**These are logic errors, not scaling errors.**
