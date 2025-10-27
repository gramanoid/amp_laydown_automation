# Task 4: Manual PowerPoint Review → Compare Sign-Off

**Status:** ⏸️ AWAITING MANUAL EXECUTION
**Estimated Time:** 0.5h
**Priority:** 🔴 CRITICAL

---

## Purpose

Perform final visual quality sign-off using PowerPoint's built-in **Review → Compare** feature to validate that the generated deck matches Template V4 structural requirements.

**What we're validating:**
- ✅ Table geometry (position, size, alignment)
- ✅ Column widths and row heights
- ✅ Cell formatting (borders, fills, fonts)
- ✅ Slide layout consistency

**What we're NOT validating:**
- ❌ Content differences (campaign names, budget values, percentages)
- ❌ Data accuracy (covered by separate validation)

---

## Prerequisites

**Files needed:**
1. **Generated deck:** `output/presentations/run_20251027_135302/presentations.pptx` (88 slides)
2. **Template:** `template/Template_V4_FINAL_071025.pptx` (reference)

**Software:**
- Microsoft PowerPoint (desktop application)
- Screenshot tool (Windows Snipping Tool or Snip & Sketch)

---

## Step-by-Step Instructions

### Step 1: Open Generated Deck

1. Navigate to: `D:\Drive\projects\work\AMP Laydowns Automation\output\presentations\run_20251027_135302\`
2. Open `presentations.pptx` in PowerPoint
3. Wait for deck to fully load (88 slides)

### Step 2: Launch Compare Feature

1. In PowerPoint ribbon, click **Review** tab
2. Click **Compare** button (far right of ribbon)
3. In file picker dialog, browse to template:
   - Path: `D:\Drive\projects\work\AMP Laydowns Automation\template\`
   - File: `Template_V4_FINAL_071025.pptx`
4. Click **Merge** or **Compare** button

### Step 3: Review Comparison Results

PowerPoint will display a comparison view with:
- **Left pane:** Revision list (all detected differences)
- **Center pane:** Current slide with change markers
- **Right pane (optional):** Side-by-side view

**Focus on these difference categories:**

#### ✅ ACCEPTABLE Differences (Content)
- Text content changes (campaign names, budget values, market names)
- Cell values (numbers, percentages, dates)
- Chart data points
- Summary tile percentages

#### ⚠️ INVESTIGATE Differences (Structure)
- Table position or size changes
- Column width variations
- Row height variations
- Font family changes (should be Calibri 18pt)
- Border style changes
- Cell fill color changes (not data-driven)
- Missing or extra table elements
- Header row formatting
- GRAND TOTAL row formatting

### Step 4: Capture Evidence

For each **structural difference** found:

1. **Screenshot the difference:**
   - Use Snipping Tool to capture the change marker
   - Save to: `docs/27-10-25/artifacts/screenshots/`
   - Filename format: `compare_issue_<number>_<description>.png`
   - Example: `compare_issue_01_column_width_mismatch.png`

2. **Document the difference:**
   - Slide number
   - Difference category (geometry/formatting/structure)
   - Description of issue
   - Severity (CRITICAL/HIGH/MEDIUM/LOW)

### Step 5: Analyze Findings

**For EACH structural difference found, determine:**

1. **Is this a bug or expected behavior?**
   - Check if python-pptx has known limitations
   - Check if template cloning introduced the difference
   - Check if config overrides explain the difference

2. **Does this affect visual quality?**
   - Will stakeholders notice this difference?
   - Does it impact readability or professionalism?
   - Is it consistent across all slides or isolated?

3. **Action required:**
   - CRITICAL: Blocks sign-off, must fix immediately
   - HIGH: Should fix before production use
   - MEDIUM: Can defer to future iteration
   - LOW: Cosmetic only, can ignore

### Step 6: Sign-Off Decision

**Sign-off criteria:**
- ✅ No CRITICAL structural differences found
- ✅ Any HIGH differences have mitigation plan
- ✅ Geometry matches template (validated in Tasks 1-3)
- ✅ Overall visual quality meets stakeholder expectations

**If sign-off granted:**
- Proceed to Task 5 (archive adopt-template-cloning-pipeline OpenSpec change)

**If sign-off blocked:**
- Document blocking issues in findings report
- Create remediation tasks
- Update MASTER_TODOLIST with new HIGH priority items

---

## Documentation Template

Copy this template to `task4_powerpoint_compare_signoff.md` when complete:

```markdown
# Task 4: PowerPoint Review → Compare Sign-Off - [COMPLETE/BLOCKED]

**Status:** [✅ COMPLETE / ❌ BLOCKED]
**Completed:** [Date]
**Time:** [Actual hours]
**Reviewer:** [Your name]

---

## Comparison Summary

**Decks compared:**
- Generated: `presentations.pptx` (88 slides, 565KB)
- Template: `Template_V4_FINAL_071025.pptx`

**Total differences detected:** [Number]
**Structural differences:** [Number]
**Content differences (ignored):** [Number]

---

## Structural Differences Found

### [Issue #1 - Title]

**Severity:** [CRITICAL/HIGH/MEDIUM/LOW]
**Slide(s):** [Slide numbers]
**Category:** [Geometry/Formatting/Structure]

**Description:**
[Detailed description of what's different]

**Screenshot:**
![Issue 1](screenshots/compare_issue_01_description.png)

**Analysis:**
[Why this happened, expected vs actual behavior]

**Action:**
[Required fix or justification for accepting]

---

[Repeat for each structural difference]

---

## Sign-Off Decision

**Decision:** [APPROVED ✅ / BLOCKED ❌]

**Rationale:**
[Explanation of sign-off decision based on findings]

**Blocking issues (if any):**
1. [Issue that must be resolved]
2. [Issue that must be resolved]

**Deferred issues (if any):**
1. [Issue accepted for this iteration, will address later]
2. [Issue accepted for this iteration, will address later]

---

## Recommendations

**Immediate actions:**
- [Action 1]
- [Action 2]

**Future improvements:**
- [Improvement 1]
- [Improvement 2]

---

## Next Steps

[If APPROVED:]
✅ Task 4 Complete - Sign-off granted
⏭️ Task 5 Next - Archive adopt-template-cloning-pipeline OpenSpec change

[If BLOCKED:]
❌ Task 4 Blocked - Remediation required
🔧 Create remediation tasks and update MASTER_TODOLIST
```

---

## Tips for Efficient Review

1. **Use keyboard shortcuts:**
   - `Page Down`: Next slide
   - `Page Up`: Previous slide
   - `Ctrl+Home`: First slide
   - `Ctrl+End`: Last slide

2. **Filter the revision list:**
   - Ignore "Content" changes
   - Focus on "Formatting" and "Structure" changes

3. **Sample slides for review:**
   - Slide 1 (first slide)
   - Slide 10 (mid-deck)
   - Slide 30 (mid-deck)
   - Slide 50 (mid-deck)
   - Slide 88 (last slide)
   - Any slide with continuation (look for CARRIED FORWARD rows)

4. **Known acceptable differences:**
   - All text content (campaign names, values, dates)
   - Summary tile percentages
   - Chart data series
   - Legend labels

---

## Expected Outcome

**Best case (likely):**
- ✅ Only content differences detected
- ✅ Structure matches template exactly
- ✅ Sign-off granted immediately
- ✅ Move to Task 5 (archive OpenSpec change)

**Alternate case:**
- ⚠️ Minor formatting differences detected (font sizes, cell borders)
- ⚠️ Differences documented and analyzed
- ✅ Sign-off granted with deferred improvements
- ✅ Move to Task 5 with notes for future work

**Worst case (unlikely):**
- ❌ Major structural differences detected
- ❌ Geometry misalignment found
- ❌ Sign-off blocked
- 🔧 Remediation tasks created

---

## Reference Documents

**Tasks 1-3 Results:**
- `docs/27-10-25/artifacts/task1_geometry_constants_complete.md` - Geometry verified
- `docs/27-10-25/artifacts/task2_continuation_layout_complete.md` - Layout verified
- `docs/27-10-25/artifacts/task3_visual_diff_complete.md` - Visual diff baseline

**OpenSpec Change:**
- `openspec/changes/adopt-template-cloning-pipeline/proposal.md` - Why this matters
- `openspec/changes/adopt-template-cloning-pipeline/tasks.md` - Task 7/8 (88% complete)

---

## Notes

- This is the FINAL validation before archiving the template cloning OpenSpec change
- Tasks 1-3 already verified geometry at code level - this is visual confirmation
- PowerPoint Compare is more thorough than visual_diff.py for structural validation
- Expected completion time: 30 minutes (most differences will be content-based)
- If no structural issues found, this completes the critical path for visual parity
