# Tasks for Campaign Pagination Implementation

## Status: Pending
- **Created**: 2025-10-27
- **Priority**: MEDIUM (after visual parity and post-processing validation)
- **Estimated effort**: 9-14 hours

## Phase 1: Analysis & Design (2-3 hours)

- [ ] **1.1 Analyze campaign size distribution**
  - Read `template/BulkPlanData_2025_10_14.xlsx` and compute:
    - Campaign row count distribution (histogram)
    - % of campaigns with >32 rows
    - Average rows per campaign by market
    - Max campaign size in dataset
  - Document findings in `docs/27-10-25/artifacts/campaign_size_analysis.md`

- [ ] **1.2 Design smart pagination algorithm**
  - Define lookahead logic: when to start fresh slide vs. continue
  - Handle edge cases:
    - Campaign with exactly 32 rows
    - Campaign with >32 rows (must split)
    - First campaign on a slide
    - Last campaign on a slide
  - Document algorithm in `design.md`

- [ ] **1.3 Update configuration schema**
  - Add to `config/master_config.json`:
    - `features.smart_campaign_pagination` (boolean, default false)
    - `presentation.table.min_rows_for_fresh_slide` (int, default 5)
  - Document configuration options

## Phase 2: Implementation (4-6 hours)

- [ ] **2.1 Implement campaign lookahead logic**
  - Location: `amp_automation/presentation/assembly.py`
  - Add function: `should_start_campaign_on_fresh_slide(remaining_capacity, campaign_row_count) -> bool`
  - Logic:
    - If campaign_row_count <= remaining_capacity: False (fits on current slide)
    - If campaign_row_count > MAX_ROWS: True (large campaign, start fresh)
    - If campaign_row_count > remaining_capacity: True (doesn't fit, start fresh)
    - Else: False
  - Unit test the function

- [ ] **2.2 Update slide creation logic**
  - Modify slide assembly to check `should_start_campaign_on_fresh_slide()` before adding campaign
  - If True: finalize current slide and start new one
  - Track remaining capacity per slide
  - Ensure CARRIED FORWARD and GRAND TOTAL rows added before finalizing

- [ ] **2.3 Handle >32 row campaigns**
  - Large campaigns still need to split
  - Ensure they start on fresh slide even if split is required
  - Continuation slides properly carry forward headers and formatting
  - Each continuation chunk gets MONTHLY TOTAL at appropriate position

- [ ] **2.4 Wire configuration toggle**
  - Read `features.smart_campaign_pagination` from config
  - Only apply smart pagination if enabled (default: off for backward compatibility)
  - Log when smart pagination triggers fresh slide

## Phase 3: Testing (2-3 hours)

- [ ] **3.1 Update existing tests**
  - Modify `tests/test_assembly_split.py` for new behavior
  - Ensure backward compatibility when feature disabled
  - Add test fixtures for various campaign size scenarios

- [ ] **3.2 Add new test coverage**
  - Test small campaigns (5-10 rows) with smart pagination
  - Test medium campaigns (15-30 rows) with various remaining capacities
  - Test large campaigns (>32 rows) start on fresh slide
  - Test edge case: campaign exactly 32 rows
  - Test configuration toggle (on/off)

- [ ] **3.3 Regenerate test deck**
  - Generate deck with smart pagination enabled
  - Generate deck with smart pagination disabled (baseline)
  - Compare slide counts and campaign distributions
  - Verify no campaigns split mid-way (unless >32 rows)

- [ ] **3.4 Run structural validation**
  - `python tools/validate_structure.py` on both decks
  - Ensure all structural requirements still met
  - Verify CARRIED FORWARD and GRAND TOTAL rows correct

## Phase 4: Documentation & Validation (1-2 hours)

- [ ] **4.1 Update documentation**
  - Update `README.md` with smart pagination feature
  - Update `AGENTS.md` with configuration guidance
  - Document in `docs/27-10-25/artifacts/` with examples

- [ ] **4.2 Visual inspection**
  - Manually review generated deck
  - Check campaign boundaries align with slide boundaries
  - Verify no unexpected splits
  - Compare with stakeholder expectations

- [ ] **4.3 Performance validation**
  - Measure generation time impact (should be negligible)
  - Measure slide count increase (estimate 5-10%)
  - Document trade-offs

- [ ] **4.4 Create before/after comparison**
  - Generate deck without smart pagination (before)
  - Generate deck with smart pagination (after)
  - Document slide count, campaign distribution, visual differences
  - Archive in `docs/27-10-25/artifacts/campaign_pagination_comparison.md`

## Success Criteria

- ✅ Campaigns <32 rows never split across slides
- ✅ Campaigns >32 rows start on fresh slide then split naturally
- ✅ Configuration toggle works (backward compatible)
- ✅ All existing tests pass
- ✅ New tests cover edge cases
- ✅ Structural validation passes
- ✅ Documentation updated
- ✅ Slide count increase <15% (acceptable trade-off)

## Dependencies

- Template geometry constants captured (from complete-oct15-followups 1.1)
- Continuation slide logic stable (already done)
- Test suite rehydrated (in progress)

## Blockers

- None currently (design decision made: Option A)

## Notes

- Feature gated behind configuration flag for safe rollout
- Can be toggled on/off without code changes
- Default: OFF (backward compatible) until validated in production
- Recommend enabling after Phase 4 validation complete
