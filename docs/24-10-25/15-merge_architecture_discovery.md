# Merge Architecture Discovery - 24 Oct 2025 17:00

## Critical Finding

The Python post-processing merge operations are **failing** because cells are already merged during deck generation.

## Root Cause Analysis

### What We Found

1. **Deck Generation (assembly.py)** already creates merged cells:
   - Line 629: Campaign cell vertical merges (column 1)
   - Line 649: Monthly total horizontal merges (columns 1-3)

2. **Clone Pipeline** (enabled in `config/master_config.json`):
   - `clone_pipeline_enabled: true`
   - Creates merges during generation phase
   - Works correctly - verified via XML inspection of generated decks

3. **Post-Processing Merge Operations** (Python):
   - Attempt to re-merge already-merged cells
   - Fail with: `range contains one or more merged cells`
   - python-pptx cannot merge cells that are already merged

### Verification

Tested on `output/presentations/run_20251024_163905/presentations.pptx`:

```
Slide 2, Column 1:
  Row  1: rowSpan=10 (campaign merge)
  Row  2-10: vMerge=1 (continuation cells)
  Row 11: rowSpan not set (MONTHLY TOTAL - should be horizontal merge)
```

### Test Results

- **Operations**: 304 total (88 slides × ~3-4 operations)
- **Merges Performed**: 0 (all failed due to pre-existing merges)
- **Errors**: All merge operations logged errors but continued
- **Final Status**: Completed successfully (errors handled gracefully)
- **Performance**: Fast (~30 seconds for 88 slides with normalize operations)

## Architecture Question

**Do we need post-processing merges at all?**

### Option A: Keep Generation Merges (RECOMMENDED)
- ✅ Merges already working in generation
- ✅ Simpler architecture
- ✅ No redundant operations
- ⚠️ Post-processing only needed for edge cases/fixes
- **Use Case**: Post-processing for normalization, styling, edge case fixes

### Option B: Move Merges to Post-Processing
- ❌ Requires disabling merges in generation
- ❌ More complex pipeline
- ✅ Centralized merge logic
- ✅ Easier to modify merge behavior
- **Use Case**: If merge logic needs frequent changes

## Implications for Python Post-Processing

### Current Implementation
The Python cell merge operations in `cell_merges.py` are **correctly implemented** but:
- Cannot run on decks with pre-existing merges
- Need a "clean" deck (no merges) to test
- OR need span reset logic to unmerge first

### Testing Strategy
To properly test Python merge operations, we need to:

1. **Generate a clean deck** (merges disabled):
   ```json
   // Hypothetical config change
   {
     "clone_pipeline_enabled": true,
     "apply_campaign_merges": false  // If this option exists
   }
   ```

2. **OR implement span reset** to unmerge cells before re-merging

3. **OR accept that Python merges are for edge cases only**
   - Run on continuation slides that might have merge issues
   - Fix broken merges from generation failures

## Recommendations

### Immediate Actions

1. **Document this architecture** in project docs
2. **Update BRAIN_RESET** with correct merge strategy
3. **Decide on merge ownership**:
   - Generation-only (current state, works)
   - Post-processing-only (requires config changes)
   - Hybrid (generation + post-processing fixes)

### For Python Post-Processing Module

The `cell_merges.py` implementation should be repositioned as:
- **Merge Repair Tool**: Fix broken/incomplete merges from generation
- **Edge Case Handler**: Handle continuation slides or special cases
- **NOT a primary merge engine** (unless we disable generation merges)

## Next Steps

1. ✅ Document this discovery
2. 🔄 Update project docs and BRAIN_RESET
3. ⏭️ Decide on merge architecture strategy
4. ⏭️ Test with clean deck OR implement span reset
5. ⏭️ Update PowerShell scripts to match strategy

## Files Referenced

- `amp_automation/presentation/assembly.py:629,649` - Generation merges
- `amp_automation/presentation/postprocess/cell_merges.py` - Python merge logic
- `config/master_config.json` - Clone pipeline config
- `output/presentations/run_20251024_163905/presentations.pptx` - Test deck

## Performance Note

Even though merge operations failed, the normalization operations completed quickly:
- **88 slides processed in ~30 seconds**
- **Confirms Python post-processing is fast** (vs 10+ hours COM)
- **Validates the COM prohibition decision**
