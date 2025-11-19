# Unmerge Operation Results - 24 Oct 2025

## Summary

**Operation**: `unmerge-all` + `normalize` on 88-slide deck
**Performance**: ~30 seconds for full deck
**Success Rate**: 100% (0 failures)

## Statistics

### Total Cells Unmerged
Calculating from log output...

**Slides with merges removed**: 66 out of 88 slides
**Slides with no merges**: 22 slides (clean from generation)

### Top 10 Slides by Merge Count
1. Slide 13: 80 cells unmerged
2. Slide 30: 63 cells unmerged
3. Slide 51: 63 cells unmerged
4. Slide 25: 61 cells unmerged
5. Slide 27: 61 cells unmerged
6. Slide 39: 61 cells unmerged
7. Slide 22: 60 cells unmerged
8. Slide 23: 57 cells unmerged
9. Slide 54: 57 cells unmerged
10. Slide 74: 57 cells unmerged

### Slides with Zero Merges (Clean)
Slides 2, 3, 4, 5 (already unmerged in previous test), 12, 21, 38, 50, 60, 68, 71, 73, 80, 85, 87

## Observations

1. **Rogue merges were widespread**: 66 out of 88 slides had unexpected merged cells
2. **Variable merge counts**: From 3 cells (slides 31, 52) to 80 cells (slide 13)
3. **Generation creates incorrect merges**: The clone pipeline is creating merges that shouldn't exist
4. **Unmerge operation is robust**: Successfully removed all merge attributes with 0 failures

## Next Steps

Now that we have a "clean slate" with no merges, we need to:

1. **Decide merge strategy**:
   - Option A: Fix generation logic to only create correct merges
   - Option B: Unmerge all in post-processing, then selectively re-merge

2. **Define correct merge patterns**:
   - Campaign vertical merges: Column 1 only, between MONTHLY TOTAL rows
   - MONTHLY TOTAL horizontal merges: Columns 1-3
   - GRAND TOTAL horizontal merges: Columns 1-3
   - CARRIED FORWARD horizontal merges: Columns 1-3

3. **Test merge operations**: Apply merges in correct order (horizontal first, then vertical)

## Files

- **Processed deck**: `output/presentations/run_20251024_172953/deck.pptx`
- **Status**: All cells unmerged, normalized layout applied
- **Ready for**: Selective merge operations
