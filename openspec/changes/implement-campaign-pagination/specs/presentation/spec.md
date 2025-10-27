# Presentation Spec - Campaign Pagination

## ADDED Requirements

### Requirement: Smart Campaign Pagination
When `features.smart_campaign_pagination` is enabled, the system SHALL prevent campaign splits across slides for campaigns with fewer than `max_rows_per_slide` rows by starting campaigns on fresh slides when they do not fit on the current slide.

#### Scenario: Small campaign fits on current slide
- **WHEN** a campaign with 20 rows is added to a slide with 25 rows remaining capacity
- **AND** smart campaign pagination is enabled
- **THEN** the campaign is added to the current slide
- **AND** no fresh slide is created

#### Scenario: Campaign doesn't fit on current slide
- **WHEN** a campaign with 25 rows is added to a slide with 15 rows remaining capacity
- **AND** smart campaign pagination is enabled
- **THEN** the current slide is finalized with GRAND TOTAL
- **AND** a fresh slide is started
- **AND** the campaign begins on the fresh slide

#### Scenario: Large campaign exceeds max rows
- **WHEN** a campaign with 50 rows is added
- **AND** smart campaign pagination is enabled
- **THEN** a fresh slide is started for the campaign
- **AND** the first 32 rows are placed on the fresh slide
- **AND** continuation slide(s) are created for remaining rows
- **AND** MONTHLY TOTAL appears on the final continuation slide

#### Scenario: Minimal remaining capacity
- **WHEN** remaining capacity is less than `min_rows_for_fresh_slide` (default 5)
- **AND** a new campaign is starting
- **THEN** a fresh slide is started regardless of campaign size

#### Scenario: Feature disabled (backward compatibility)
- **WHEN** `features.smart_campaign_pagination` is disabled (default)
- **THEN** campaigns are added sequentially until max_rows_per_slide is reached
- **AND** campaigns may split across slides at the 32-row boundary
- **AND** behavior is identical to current implementation

### Requirement: Campaign Boundary Integrity
For campaigns with fewer than `max_rows_per_slide` rows, the system SHALL maintain campaign integrity by ensuring all campaign rows (media rows + MONTHLY TOTAL) appear on a single slide.

#### Scenario: Complete campaign on single slide
- **WHEN** a campaign with 28 rows is processed
- **AND** smart campaign pagination is enabled
- **THEN** all 28 rows (media rows + MONTHLY TOTAL) appear on one slide
- **AND** the campaign is not split

#### Scenario: Campaign boundary respected
- **WHEN** multiple campaigns are added to slides
- **AND** smart campaign pagination is enabled
- **THEN** no campaign with <32 rows is split across slides
- **AND** each campaign starts and ends on the same slide OR starts fresh on continuation

### Requirement: Configuration Control
The system SHALL provide configuration options to control smart campaign pagination behavior without code changes.

#### Scenario: Enable smart pagination
- **WHEN** `features.smart_campaign_pagination` is set to `true` in config
- **THEN** smart pagination logic is applied during slide generation
- **AND** campaigns are checked for fit before adding to current slide

#### Scenario: Disable smart pagination
- **WHEN** `features.smart_campaign_pagination` is set to `false` in config
- **THEN** original pagination logic is used
- **AND** campaigns fill slides sequentially to 32-row capacity

#### Scenario: Configure minimum rows threshold
- **WHEN** `presentation.table.min_rows_for_fresh_slide` is set to 5
- **THEN** fresh slides are started when remaining capacity < 5 rows
- **AND** threshold is applied to fresh slide decisions

## MODIFIED Requirements

### Requirement: Continuation Slide Generation
Slides that split due to row limits SHALL retain headers, legend groups, carried totals, and summary tiles so that downstream slides remain visually consistent with the initial page. **When smart campaign pagination is enabled, large campaigns (>32 rows) SHALL start on fresh slides before splitting into continuations.**

#### Scenario: Split table retains formatting
- **WHEN** a market-brand-year dataset exceeds the maximum table body rows for one slide
- **THEN** the automation creates continuation slides that reproduce headers, column widths, banding, and summary tiles
- **AND** carried total rows appear at the end of each continuation with the correct styling and values

#### Scenario: Large campaign starts fresh then splits
- **WHEN** a campaign with >32 rows is processed
- **AND** smart campaign pagination is enabled
- **THEN** the campaign starts on a fresh slide
- **AND** continuation slides are created for rows beyond the first 32
- **AND** all continuation slides maintain proper formatting

#### Scenario: Campaign split boundary marked clearly
- **WHEN** a large campaign requires multiple slides
- **AND** smart campaign pagination is enabled
- **THEN** the first slide of the campaign has no carried forward from previous campaigns
- **AND** each continuation slide has CARRIED FORWARD row from previous chunk
- **AND** MONTHLY TOTAL appears only on the final continuation slide

## Notes

- Smart campaign pagination is disabled by default for backward compatibility
- Enabling smart pagination may increase total slide count by 5-15%
- Campaigns with exactly 32 rows are treated as "fits on current slide" if capacity available
- The `min_rows_for_fresh_slide` threshold prevents inefficient space usage
- All existing structural requirements (GRAND TOTAL, CARRIED FORWARD, headers) remain unchanged
