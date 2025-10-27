# Campaign Pagination Design

## Context

The current slide generation logic fills slides up to 32 body rows regardless of campaign boundaries. This causes campaigns to split across continuation slides, reducing readability and making it harder for stakeholders to review individual campaigns in isolation.

**Current behavior:**
- Slide capacity: 32 body rows
- Campaigns added sequentially until capacity reached
- When capacity exceeded, create continuation slide
- No awareness of campaign boundaries

**Problem:**
- Campaign with 62 rows creates 2 slides: [32 rows] + [30 rows]
- Campaign split mid-way, MONTHLY TOTAL appears on second slide only
- Stakeholders must review multiple slides to see complete campaign

## Goals / Non-Goals

**Goals:**
- Prevent campaign splits for campaigns <32 rows
- Improve campaign readability and visual consistency
- Maintain continuation slide functionality for large campaigns
- Minimal performance impact (<5% generation time increase)
- Backward compatible (feature toggle)

**Non-Goals:**
- Perfect space efficiency (some slides may have <32 rows)
- Eliminate all campaign splits (campaigns >32 rows must still split)
- Change continuation slide formatting or layout
- Modify MONTHLY TOTAL or GRAND TOTAL logic

## Decisions

### Decision 1: Smart Pagination Algorithm (Option A)

**Chosen:** Before starting a new campaign, check if it fits on current slide. If not, start on fresh slide.

**Algorithm:**
```python
def should_start_campaign_on_fresh_slide(
    remaining_capacity: int,
    campaign_row_count: int,
    min_rows_threshold: int = 5
) -> bool:
    """
    Determine if campaign should start on fresh slide.

    Args:
        remaining_capacity: Rows available on current slide
        campaign_row_count: Total rows in campaign (including media rows + MONTHLY TOTAL)
        min_rows_threshold: Minimum rows worth starting fresh slide (default 5)

    Returns:
        True if campaign should start fresh slide, False otherwise
    """
    # Case 1: Campaign fits on current slide - continue
    if campaign_row_count <= remaining_capacity:
        return False

    # Case 2: Very small remaining capacity - start fresh
    if remaining_capacity < min_rows_threshold:
        return True

    # Case 3: Campaign too large for any slide - start fresh then split
    if campaign_row_count > MAX_ROWS_PER_SLIDE:
        return True

    # Case 4: Campaign doesn't fit current slide - start fresh
    return True
```

**Example scenarios:**

| Remaining Capacity | Campaign Rows | Decision | Rationale |
|-------------------|---------------|----------|-----------|
| 25 | 20 | Continue | Fits with room to spare |
| 10 | 15 | Fresh slide | Doesn't fit, worth starting fresh |
| 3 | 8 | Fresh slide | Too little space remaining |
| 15 | 40 | Fresh slide | Large campaign, will split anyway |
| 32 | 32 | Continue | Exact fit (first campaign on slide) |

**Alternatives considered:**
- **Option B (Campaign-Aware Splitting):** More complex, doesn't solve readability issue
- **Option C (Dynamic Limits):** Too unpredictable, testing nightmare
- **Option D (Status Quo):** Doesn't address stakeholder concerns

### Decision 2: Configuration Toggle

**Chosen:** Feature gated behind `features.smart_campaign_pagination` flag in config.

**Rationale:**
- Allows gradual rollout and validation
- Backward compatible (default OFF)
- Easy to disable if issues found
- No code changes required to toggle

**Configuration:**
```json
{
  "features": {
    "smart_campaign_pagination": false,
    "clone_pipeline_enabled": true
  },
  "presentation": {
    "table": {
      "max_rows_per_slide": 32,
      "min_rows_for_fresh_slide": 5
    }
  }
}
```

### Decision 3: Large Campaign Handling

**Chosen:** Campaigns >32 rows start on fresh slide, then split naturally.

**Rationale:**
- Large campaigns must split (template constraint)
- Starting fresh provides visual cue that "campaign begins here"
- Continuation slides properly formatted with headers and CARRIED FORWARD
- MONTHLY TOTAL appears on final continuation slide

**Example: 62-row campaign**

Before (current):
- Slide N: [Campaign A: 20 rows] + [Campaign B: 12 rows] → 32 rows
- Slide N+1: [Campaign B: 50 more rows] → need another split
- Slide N+2: [Campaign B: 38 more rows, MONTHLY TOTAL]

After (smart pagination):
- Slide N: [Campaign A: 20 rows] → 20 rows (started fresh on prior slide)
- Slide N+1: [Campaign B: 32 rows] → Fresh start, first chunk
- Slide N+2: [Campaign B: 30 rows, MONTHLY TOTAL] → Continuation

### Decision 4: Minimum Rows Threshold

**Chosen:** `min_rows_for_fresh_slide = 5` (configurable)

**Rationale:**
- If <5 rows remain on slide, not worth squeezing in small campaign
- Better to start fresh for visual consistency
- Avoids "orphan" small campaigns at bottom of slides
- Configurable for future tuning

**Edge cases:**
- If remaining capacity is 1-4 rows, skip to fresh slide
- If next campaign is 1-4 rows, allow it (better than wasting full slide)

## Risks / Trade-offs

### Risk 1: Increased Slide Count
**Impact:** Slide count may increase 5-15% depending on campaign size distribution

**Mitigation:**
- Analyzed `BulkPlanData_2025_10_14.xlsx`: Most campaigns <25 rows
- Estimated impact: +5-10% slide count (acceptable for readability gain)
- Configuration toggle allows disabling if slide count becomes issue

### Risk 2: Inconsistent Slide Fullness
**Impact:** Some slides may have 15 rows, others 32 rows (less uniform)

**Mitigation:**
- Stakeholder feedback prioritizes campaign integrity over space efficiency
- Visual consistency within campaigns more important than across slides
- GRAND TOTAL still appears on every slide (consistent footer)

### Risk 3: Performance Impact
**Impact:** Lookahead logic adds computational overhead

**Mitigation:**
- Algorithm is O(1) per campaign (single comparison)
- Total overhead: O(N) where N = number of campaigns (~60-70)
- Expected impact: <1 second additional generation time
- Acceptable for 5-10 minute total generation time

### Risk 4: Backward Compatibility
**Impact:** Existing workflows may depend on current slide splits

**Mitigation:**
- Feature toggle ensures backward compatibility
- Default: OFF until validated
- Existing tests continue to pass with feature disabled
- New tests cover both modes

## Migration Plan

### Phase 1: Implementation & Testing (Week 1)
1. Implement algorithm in `assembly.py`
2. Add configuration toggle
3. Write unit tests
4. Test with fresh deck generation

### Phase 2: Validation (Week 1-2)
1. Generate decks with feature ON and OFF
2. Compare slide counts and layouts
3. Visual inspection with stakeholder
4. Performance benchmarking

### Phase 3: Controlled Rollout (Week 2)
1. Enable for single market/brand test
2. Validate output meets expectations
3. Document any issues or edge cases
4. Adjust `min_rows_threshold` if needed

### Phase 4: Full Rollout (Week 3)
1. Enable by default in config
2. Update documentation
3. Archive comparison artifacts
4. Mark OpenSpec change as complete

### Rollback Plan
If issues found:
1. Set `features.smart_campaign_pagination = false` in config
2. Regenerate decks (reverts to original behavior)
3. No code changes required

## Open Questions

- ✅ **Q:** Which option to choose? **A:** Option A (Smart Pagination)
- ✅ **Q:** Should feature be enabled by default? **A:** No, default OFF for safe rollout
- ⏭️ **Q:** What should `min_rows_threshold` be? **A:** Start with 5, tune based on validation
- ⏭️ **Q:** How to handle campaigns with exactly 32 rows? **A:** Allow on current slide if space available
- ⏭️ **Q:** Should we log when fresh slide triggered? **A:** Yes, at INFO level for observability

## Technical Notes

### Implementation Location
- **Primary:** `amp_automation/presentation/assembly.py` - `build_presentation()` function
- **Config:** `config/master_config.json` - Feature flags
- **Tests:** `tests/test_assembly_split.py` - Continuation and pagination tests

### Key Functions to Modify
1. `build_presentation()` - Main assembly loop, check before adding campaign
2. `create_continuation_slide()` - Ensure works with fresh slide starts
3. `calculate_campaign_row_count()` - NEW helper function
4. `should_start_campaign_on_fresh_slide()` - NEW decision function

### Data Structures
```python
class SlideState:
    rows_used: int
    remaining_capacity: int
    campaigns: List[Campaign]
    has_carried_forward: bool

class Campaign:
    name: str
    rows: List[Row]
    row_count: int  # includes media rows + MONTHLY TOTAL
```

### Logging Strategy
```python
if should_start_fresh:
    logger.info(
        f"Starting campaign '{campaign_name}' on fresh slide "
        f"(rows={campaign_row_count}, remaining={remaining_capacity})"
    )
```

## Success Metrics

Post-implementation, validate:
- ✅ 0 campaigns <32 rows split across slides
- ✅ All campaigns >32 rows start on fresh slide
- ✅ Slide count increase <15%
- ✅ Generation time increase <5%
- ✅ All structural validation tests pass
- ✅ Stakeholder approval on visual consistency
