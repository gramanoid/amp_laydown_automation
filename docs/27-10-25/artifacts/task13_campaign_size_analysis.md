# Task 13: Campaign Size Distribution Analysis - Complete

**Status:** ✅ COMPLETE
**Completed:** 27 Oct 2025
**Time:** 1.5h
**Data Source:** `template/BulkPlanData_2025_10_14.xlsx`

---

## Executive Summary

Campaign size analysis reveals **27.4% of campaigns exceed 32 rows**, requiring splits. Smart pagination would prevent splits for small campaigns but **increases total slide count by 45.7%** (186 → 271 slides) - significantly higher than the original 5-15% estimate.

**Recommendation:** Proceed with caution or explore alternative approaches (e.g., higher row threshold).

---

## Dataset Overview

**Data Source:** BulkPlanData_2025_10_14.xlsx
- **Total rows:** 4,914
- **Total columns:** 114
- **Unique market/brand/year combinations:** 60
- **Total campaigns analyzed:** 186

**Grouping:** Market Sub-Cluster + Plan Brand + Plan Year

---

## Campaign Size Distribution

### Summary Statistics

| Metric | Value |
|--------|-------|
| **Total campaigns** | 186 |
| **Campaigns > 32 rows** | 51 (27.4%) |
| **Campaigns ≤ 32 rows** | 135 (72.6%) |
| **Average campaign size** | 26.4 rows |
| **Median campaign size** | 12 rows (estimated from distribution) |
| **Maximum campaign size** | 232 rows |
| **Minimum campaign size** | 1 row |

### Largest Campaign

**Campaign:** Release Starts Here
- **Market:** GINE (Guinea?)
- **Brand:** Haleon | Panadol
- **Year:** 2025
- **Size:** 232 rows
- **Impact:** Would require 8 continuation slides (232 / 32 = 7.25 → 8 slides)

### Size Buckets

| Size Range | Count | Percentage | Notes |
|------------|-------|------------|-------|
| **1-10 rows** | 53 | 28.5% | Very small campaigns |
| **11-20 rows** | 57 | 30.6% | Small campaigns |
| **21-32 rows** | 25 | 13.4% | Medium campaigns (fit on one slide) |
| **33-50 rows** | 27 | 14.5% | Large campaigns (require 2 slides) |
| **51-100 rows** | 21 | 11.3% | Very large campaigns (3-4 slides) |
| **>100 rows** | 3 | 1.6% | Mega campaigns (4+ slides) |

**Key Insight:** 72.6% of campaigns fit on a single slide (≤32 rows). These are the campaigns that would benefit from smart pagination.

---

## Distribution Details

### Most Common Sizes

| Rows | Campaigns | Percentage |
|------|-----------|------------|
| 12 | 13 | 7.0% |
| 11 | 11 | 5.9% |
| 2 | 8 | 4.3% |
| 24 | 8 | 4.3% |
| 14 | 7 | 3.8% |
| 13 | 7 | 3.8% |
| 1 | 7 | 3.8% |
| 4 | 7 | 3.8% |

### Edge Cases

**Exactly 32 rows:** 1 campaign (0.5%)
- This is the boundary case - fits perfectly on one slide
- Smart pagination should treat as "fits on current slide"

**Very large campaigns (>100 rows):**
- 232 rows: 1 campaign (GINE - Panadol - Release Starts Here)
- 155 rows: 1 campaign
- 148 rows: 1 campaign
- 133 rows: 1 campaign
- 132 rows: 1 campaign
- 113 rows: 1 campaign

**Total: 6 campaigns (3.2%)** would require 4+ slides even with smart pagination

---

## Market-Level Analysis

### Average Campaign Size by Market

| Market | Avg Rows | Campaign Count | Notes |
|--------|----------|----------------|-------|
| **GINE** | 62.6 | 20 | Highest avg - many large campaigns |
| **South Africa** | 25.4 | 55 | Largest volume of campaigns |
| **Turkey** | 23.5 | 30 | Moderate size campaigns |
| **North Africa** | 22.5 | 2 | Small sample size |
| **KSA** | 27.6 | 23 | Above average size |
| **Egypt** | 18.8 | 13 | Below average size |
| **Pakistan** | 17.1 | 20 | Below average size |
| **East Africa** | 15.9 | 15 | Smaller campaigns |
| **French West Africa** | 6.8 | 8 | Smallest campaigns |

### Key Observations

**GINE market is an outlier:**
- Average: 62.6 rows (2.4x overall average of 26.4)
- 20 campaigns total
- Includes the 232-row mega campaign
- Smart pagination will have highest impact here

**South Africa has most campaigns:**
- 55 campaigns (29.6% of all campaigns)
- Average size: 25.4 rows (close to overall average)
- Represents typical campaign distribution

**French West Africa has smallest campaigns:**
- Average: 6.8 rows
- All campaigns likely fit comfortably on shared slides
- Smart pagination less beneficial here

---

## Smart Pagination Impact Estimate

### Slide Count Comparison

| Approach | Total Slides | Notes |
|----------|--------------|-------|
| **Current (sequential fill)** | 186 | Campaigns split at 32-row boundary |
| **Smart pagination** | 271 | Campaigns <32 rows don't split |
| **Increase** | +85 slides | **+45.7%** |

### Analysis of Impact

**Why such a large increase?**

1. **Current approach maximizes density:**
   - Fills slides to 32 rows sequentially
   - Campaigns can split mid-campaign
   - Optimizes for minimum slide count

2. **Smart pagination prioritizes readability:**
   - Prevents splits for campaigns <32 rows
   - Starts fresh slides when campaigns don't fit
   - Leaves unused capacity on slides

3. **Example scenario:**
   - Current: Campaign A (25 rows) + Campaign B (20 rows) = 45 rows → 2 slides with split
   - Smart: Campaign A (25 rows) on Slide 1 (7 rows unused), Campaign B (20 rows) on Slide 2 (12 rows unused) = 2 slides
   - Both use 2 slides, but smart pagination wastes 19 rows of capacity

4. **Cumulative effect:**
   - 135 campaigns ≤32 rows
   - Each time a campaign doesn't fit, a fresh slide starts
   - Wasted capacity accumulates across all decks

### Visual Example

**Current approach (186 slides):**
```
Slide 1: Campaign A (25) + Campaign B (7) = 32 rows ✅ Full
Slide 2: Campaign B (18) + Campaign C (14) = 32 rows ✅ Full
Slide 3: Campaign D (30) + Campaign E (2) = 32 rows ✅ Full
...
```

**Smart pagination (271 slides):**
```
Slide 1: Campaign A (25) = 25 rows ❌ 7 rows wasted
Slide 2: Campaign B (25) = 25 rows ❌ 7 rows wasted
Slide 3: Campaign C (14) = 14 rows ❌ 18 rows wasted
Slide 4: Campaign D (30) = 30 rows ❌ 2 rows wasted
Slide 5: Campaign E (2) = 2 rows ❌ 30 rows wasted!
...
```

---

## Trade-Off Analysis

### Benefits of Smart Pagination

✅ **Improved readability:**
- Campaigns stay together (no mid-campaign splits)
- Easier to find specific campaigns
- Better UX for stakeholders

✅ **Professional appearance:**
- Each campaign is a cohesive unit
- No confusing "CARRIED FORWARD" rows for small campaigns
- Clearer visual hierarchy

### Drawbacks of Smart Pagination

❌ **45.7% more slides:**
- 186 → 271 slides
- Longer decks to navigate
- More file size

❌ **Wasted capacity:**
- ~85 slides worth of unused rows
- Could fit more data with sequential fill

❌ **Higher than projected:**
- Original estimate: 5-15% increase
- Actual estimate: 45.7% increase
- **3x higher than expected!**

---

## Recommendations

### Option 1: Proceed with Smart Pagination (As Designed)

**Pros:**
- Implements original vision
- Maximum readability improvement

**Cons:**
- 45.7% slide increase may not be acceptable
- Significant wasted capacity

**Recommendation:** **Get stakeholder buy-in first** - 45.7% is a large increase

### Option 2: Increase Row Threshold (Alternative Approach)

**Modify design:** Instead of 32 rows, allow 40-45 rows per slide

**Impact:**
- More campaigns fit on one slide
- Reduces split frequency
- Lower slide count increase

**Analysis needed:**
- Re-run with threshold=40 or 45
- Check if template supports more rows visually

**Recommendation:** Worth exploring if 45.7% is too high

### Option 3: Hybrid Approach (Selective Smart Pagination)

**Modify design:** Only apply smart pagination to campaigns >10 rows and <32 rows

**Rationale:**
- Very small campaigns (1-10 rows) can split without readability loss
- Medium campaigns (11-31 rows) benefit most from staying together
- Large campaigns (>32 rows) already split

**Impact:** Would reduce slide increase (needs re-analysis)

**Recommendation:** Good middle ground

### Option 4: Market-Specific Smart Pagination

**Modify design:** Enable smart pagination only for high-value markets (e.g., South Africa, GINE)

**Rationale:**
- Apply where it has most impact
- Allow sequential fill for markets with tiny campaigns

**Impact:** Would reduce slide increase

**Recommendation:** Adds complexity but could be viable

### Option 5: Defer Smart Pagination (Keep Current Approach)

**Rationale:**
- 45.7% increase is too high
- Current approach works
- Focus on other priorities

**Recommendation:** Valid if stakeholders don't see value in readability improvement

---

## Next Steps (Task 14)

Based on this analysis, **Task 14: Design Smart Pagination Algorithm** should consider:

1. **Validate stakeholder acceptance** of 45.7% slide increase
2. **Explore alternative thresholds** (40-45 rows instead of 32)
3. **Consider hybrid approaches** (selective or market-specific pagination)
4. **Document trade-offs clearly** in design.md

**Blocker:** Should not proceed with implementation until slide increase is acceptable to stakeholders.

---

## Data Files

**Analysis script:** `tools/analyze_campaign_sizes.py`
**Data source:** `template/BulkPlanData_2025_10_14.xlsx`

**Key outputs:**
- Total campaigns: 186
- Campaigns >32 rows: 51 (27.4%)
- Average size: 26.4 rows
- Max size: 232 rows (GINE - Panadol - Release Starts Here)
- Smart pagination impact: +85 slides (+45.7%)

---

## Appendix: Full Distribution

<details>
<summary>Click to expand full size distribution</summary>

| Rows | Campaigns | Percentage |
|------|-----------|------------|
| 1 | 7 | 3.8% |
| 2 | 8 | 4.3% |
| 3 | 6 | 3.2% |
| 4 | 7 | 3.8% |
| 5 | 2 | 1.1% |
| 6 | 4 | 2.2% |
| 7 | 3 | 1.6% |
| 8 | 4 | 2.2% |
| 9 | 6 | 3.2% |
| 10 | 6 | 3.2% |
| 11 | 11 | 5.9% |
| 12 | 13 | 7.0% |
| 13 | 7 | 3.8% |
| 14 | 7 | 3.8% |
| 15 | 4 | 2.2% |
| 16 | 3 | 1.6% |
| 17 | 2 | 1.1% |
| 18 | 5 | 2.7% |
| 19 | 1 | 0.5% |
| 20 | 2 | 1.1% |
| 21 | 4 | 2.2% |
| 22 | 3 | 1.6% |
| 24 | 8 | 4.3% |
| 25 | 4 | 2.2% |
| 27 | 1 | 0.5% |
| 28 | 3 | 1.6% |
| 30 | 2 | 1.1% |
| 31 | 1 | 0.5% |
| 32 | 1 | 0.5% |
| 34-232 | 51 | 27.4% |

</details>

---

## Conclusion

Campaign size analysis complete. **Key finding: Smart pagination increases slides by 45.7%**, which is **3x higher than the original 5-15% estimate**. This suggests the design may need adjustment before proceeding with implementation.

**Recommendation:** Pause before Task 14 to validate stakeholder acceptance of 45.7% slide increase, or explore alternative approaches (higher row threshold, hybrid, market-specific).
