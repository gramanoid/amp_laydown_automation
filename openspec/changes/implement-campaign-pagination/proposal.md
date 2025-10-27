# Campaign Pagination Implementation

## Why
Current slide generation allows campaigns to split across multiple continuation slides when they exceed 32 rows. This reduces campaign readability and visual consistency, making it harder for stakeholders to review individual campaigns. A no-campaign-splitting pagination strategy would ensure each campaign remains intact on a single slide or starts fresh on a continuation slide.

## What Changes
- Implement smart pagination logic that checks if a campaign fits on the current slide before starting it
- If a campaign doesn't fit (remaining capacity < campaign row count), start it on a fresh slide
- Handle edge case: campaigns with >32 rows must still split but will start on a fresh slide
- Add configuration option to enable/disable smart pagination
- Update continuation slide logic to respect campaign boundaries

## Impact
- Affected specs: `presentation` (slide assembly and pagination)
- Affected code:
  - `amp_automation/presentation/assembly.py` - Core slide creation and splitting logic
  - `config/master_config.json` - Add smart pagination toggle and configuration
  - Tests: `tests/test_assembly_split.py` - Update continuation slide tests
- Affected behavior:
  - Slides may have fewer than 32 rows if next campaign doesn't fit
  - Total slide count may increase slightly (more efficient for readability)
  - Campaign integrity improved (no mid-campaign splits for campaigns <32 rows)

## Design Decision: Option A - Smart Pagination

**Chosen approach:** Before starting a new campaign, check if it fits on current slide. If not, start fresh on next slide.

**Rationale:**
- Best balance of readability vs. space efficiency
- Simple to implement and test
- Predictable behavior for stakeholders
- Handles edge cases gracefully (very large campaigns still split but start fresh)

**Trade-offs accepted:**
- May leave empty rows on some slides (acceptable for improved campaign readability)
- Slight increase in total slide count (5-10% estimated)
- Worth the trade-off for visual consistency and stakeholder usability
