## ADDED Requirements
### Requirement: Template V4 Table Styling
Slides SHALL render table cells using Template V4 fonts, alignments, padding, and background fills, including wrapped dual-line media labels and summary rows that reflect config-driven colors.

#### Scenario: Apply template fonts and fills
- **WHEN** the automation builds a campaign table for any market-brand-year combination
- **THEN** each cell uses Calibri (or configured fallback), specified font sizes, alignment, and background fills matching Template V4 contract
- **AND** subtotal and header rows use the configured gray/green fills and carried subtotal styling

### Requirement: Continuation Slide Parity
Slides that split due to row limits SHALL retain headers, legend groups, carried totals, and summary tiles so that downstream slides remain visually consistent with the initial page.

#### Scenario: Split table retains formatting
- **WHEN** a market-brand-year dataset exceeds the maximum table body rows for one slide
- **THEN** the automation creates continuation slides that reproduce headers, column widths, banding, and summary tiles
- **AND** carried total rows appear at the end of each continuation with the correct styling and values
