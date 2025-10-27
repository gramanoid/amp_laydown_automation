# Campaign Pagination Task - Found Documentation

Generated: 2025-10-27 14:00
Search completed: Campaign pagination/row limit task located

---

## Task Overview

**Task Name:** Campaign Pagination Design / No-Campaign-Splitting Pagination Strategy

**Current Status:** PENDING (in "Next" or "Later" sections across multiple BRAIN_RESET files)

**Priority:** MEDIUM-TERM (deferred until visual parity and post-processing work complete)

---

## Key Findings from Documentation

### 1. Current Row Limit: 32 Rows Per Slide

**Source:** `docs/20_10_25/brain_reset.md:30`
> **Row-limit & splitting update (20 Oct morning):** Raised `MAX_ROWS_PER_SLIDE` to 32, removed forced one-campaign-per-slide split.

**Previous limit:** 17 rows (updated 14 Oct 2025)
**Current limit:** 32 body rows per slide (updated 20 Oct 2025)

**Template Capacity:**
- Maximum 32 body rows per slide
- Plus: CARRIED FORWARD row, slide-level GRAND TOTAL
- Headers, footers, summary tiles, legend on every slide

### 2. Problem Statement

**From:** `docs/22-10-25/22-10-25.md:23`
> Launch a discovery Q&A to design a **no-campaign-splitting pagination strategy** so each campaign stays on a single slide; document options and prepare the follow-on OpenSpec change once current blockers resolve.

**From:** `docs/23-10-25/BRAIN_RESET_231025.md:42`
> Facilitate a Q&A-led discovery to design **campaign pagination that prevents splits across slides**, then raise an OpenSpec change once prerequisites clear.

**Current Behavior:**
- Campaigns can split across multiple continuation slides when they have >32 rows
- Example: A campaign with 62 rows creates 2 slides (32 + 30 rows)
- This impacts visual consistency and campaign readability

**Issue:** "Without successful campaign merges, continuation slides present split Campaign columns, impacting visual parity." (docs/23-10-25/BRAIN_RESET_231025.md:46)

### 3. Task Description

**Comprehensive description from** `docs/27-10-25/BRAIN_RESET_271025.md`:
- [ ] **Campaign pagination design** - Strategy to prevent campaign splits across slides
- [ ] **Create OpenSpec proposal** - Document campaign pagination approach once design complete

**Approach:** Q&A-led discovery to explore options before implementation

### 4. Related Context

**Domain Context** (`openspec/project.md:67`):
> Template V4 geometry: up to 32 body rows per slide (plus carried-forward + slide GRAND TOTAL), summary tiles, footers, legend

**Structural Requirement** (`docs/17_10_25/brain_reset.md:46`):
> "MONTHLY TOTAL (£ 000)" appears after the final media row of each campaign (even across slide splits). A slide-level "GRAND TOTAL" closes every slide.

**Current implementation:**
- Removed forced one-campaign-per-slide split (20 Oct 2025)
- Campaigns can now span multiple slides if they exceed 32 rows
- Continuation slides properly retain headers, CARRIED FORWARD rows, and GRAND TOTAL

### 5. Historical Timeline

- **14 Oct 2025:** `max_rows_per_slide` was 17
- **20 Oct 2025 morning:** Raised `MAX_ROWS_PER_SLIDE` to 32, removed forced one-campaign-per-slide split
- **22 Oct 2025:** Task added to "Next Focus" as item #4 - design no-campaign-splitting pagination strategy
- **23 Oct 2025:** Carried forward to "Longer-Term Follow-Ups"
- **24 Oct 2025:** Listed in "Next" priorities (after immediate work)
- **27 Oct 2025:** Still pending in "Next" section

---

## Options to Consider (For Future Q&A Discovery)

### Option A: Smart Pagination (Prevent Campaign Splits)
**Concept:** Before starting a new campaign, check if it fits on current slide. If not, start it on a fresh slide.

**Pros:**
- Each campaign stays intact on one slide (or starts fresh on continuation)
- Better readability and visual consistency
- Simpler for stakeholders to review individual campaigns

**Cons:**
- May leave empty rows on slides (inefficient space usage)
- Could increase total slide count
- Need to handle campaigns with >32 rows (still need splits for very large campaigns)

**Implementation complexity:** Medium - requires lookahead logic in assembly phase

### Option B: Campaign-Aware Splitting
**Concept:** Allow splits but ensure campaign boundaries are respected for MONTHLY TOTAL rows.

**Pros:**
- Efficient space usage (no empty rows)
- Minimizes slide count
- MONTHLY TOTAL rows always appear at end of campaign section

**Cons:**
- Campaigns still split across slides
- More complex continuation logic
- MONTHLY TOTAL needs special handling on continuation slides

**Implementation complexity:** High - requires sophisticated row accounting

### Option C: Dynamic Row Limit by Campaign Size
**Concept:** Adjust effective row limit based on campaign sizes to minimize splits.

**Pros:**
- Optimizes for both space efficiency and campaign integrity
- Can handle various campaign size distributions

**Cons:**
- Very complex logic
- Unpredictable slide layouts
- Difficult to test all edge cases

**Implementation complexity:** Very High - not recommended

### Option D: Status Quo (Current Behavior)
**Concept:** Keep current implementation - campaigns split naturally at 32-row boundary.

**Pros:**
- Already implemented and working
- Consistent, predictable layout
- Maximizes space efficiency

**Cons:**
- Campaigns split across slides impact readability
- Visual continuity concerns
- Stakeholder feedback may require changes

**Implementation complexity:** None (already done)

---

## Recommended Next Steps

### Phase 1: Requirements Gathering (Q&A Discovery)
1. **Stakeholder input:**
   - Is campaign splitting acceptable for large campaigns (>32 rows)?
   - What's the typical campaign size distribution in production data?
   - Are there campaigns that commonly exceed 32 rows?

2. **Data analysis:**
   - Analyze `BulkPlanData_2025_10_14.xlsx` to identify:
     - Campaign row count distribution
     - % of campaigns with >32 rows
     - Average rows per campaign by market

3. **Business requirements:**
   - Must every campaign start on a fresh slide?
   - Or is it acceptable for small campaigns (<10 rows) to share slides?
   - Priority: space efficiency vs. campaign integrity?

### Phase 2: Design Selection
Based on Q&A discovery, select one option and document:
- Detailed implementation approach
- Edge cases and handling
- Impact on existing code (`amp_automation/presentation/assembly.py`)
- Test scenarios

### Phase 3: OpenSpec Proposal
Create `openspec/changes/implement-campaign-pagination/` with:
- `proposal.md` - Why, what changes, impact
- `tasks.md` - Implementation checklist
- `design.md` - Technical decisions and trade-offs
- `specs/presentation/spec.md` - Delta requirements

### Phase 4: Implementation
After approval, implement and test.

---

## Current Blockers

**Why this task is deferred:**
1. Visual parity work takes priority (business-critical)
2. Post-processing validation incomplete
3. Test suite needs rehydration
4. Need stakeholder input on requirements (Q&A discovery)

**Prerequisites before starting:**
- ✅ Template geometry constants captured (to understand exact constraints)
- ✅ Visual diff baseline established (to validate changes don't break layout)
- ✅ Continuation slide logic stable (foundation for pagination changes)
- ⏭️ Q&A discovery with stakeholders (to define requirements)

---

## Files to Review When Starting This Task

1. **Current implementation:**
   - `amp_automation/presentation/assembly.py` - Slide creation and splitting logic
   - `config/master_config.json` - `max_rows_per_slide` configuration (currently 32)

2. **Test data:**
   - `template/BulkPlanData_2025_10_14.xlsx` - Analyze campaign row distributions

3. **Related specs:**
   - `openspec/changes/archive/2025-10-21-update-table-styling-continuations/` - Continuation slide requirements

4. **Validation tools:**
   - `tools/validate_structure.py` - Structural validation after changes

---

## Summary

**Task:** Design and implement campaign pagination strategy to prevent (or minimize) campaign splits across slides.

**Current State:** 32-row limit allows campaigns to split naturally at boundary. Continuation slides work correctly but campaigns can be fragmented.

**Goal:** Improve campaign readability and visual consistency by implementing smart pagination that respects campaign boundaries.

**Approach:** Q&A-led discovery → Design selection → OpenSpec proposal → Implementation

**Timeline:** NEXT (after current high-priority visual parity and testing work)

**Estimated Effort:**
- Requirements gathering: 1-2 hours
- Design & proposal: 2-3 hours
- Implementation: 4-6 hours
- Testing: 2-3 hours
- **Total: 9-14 hours**
