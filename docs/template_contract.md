# Template V4 Contract (2025-10-14)

## Slide-Level Named Shapes
- `TitlePlaceholder` — Main slide title populated with `{market} - {brand}`.
- `MainDataTable` — Campaign table populated programmatically (max 17 body rows per slide).
- `QuarterBudgetQ1` — Q1 total budget tile (`Jan+Feb+Mar`).
- `QuarterBudgetQ2` — Q2 total budget tile (`Apr+May+Jun`).
- `QuarterBudgetQ3` — Q3 total budget tile (`Jul+Aug+Sep`).
- `QuarterBudgetQ4` — Q4 total budget tile (`Oct+Nov+Dec`).
- `MediaShareTelevision` — Television budget share tile.
- `MediaShareDigital` — Digital budget share tile.
- `MediaShareOther` — Other budget share tile.
- `FunnelShareAwareness` — Budget share for Awareness funnel stage.
- `FunnelShareConsideration` — Budget share for Consideration funnel stage.
- `FunnelSharePurchase` — Budget share for Purchase funnel stage.
- `FooterNotes` — Static copy + appended export date.

### Geometry Reference (inches from top-left)
- Slide master frame `Freeform: Shape 14` — covers entire 10.000" × 5.625" canvas (rounded border).
- Title placeholder — `(left 0.184", top 0.270", width 2.944", height 0.337")`.
- Table placeholder (`Table 14`) — `(0.184", 1.223", 9.124", 2.433")`; automation should align generated table to this rectangle.
- Summary bars & chips (all at `top ≈ 5.136"`):
  - Quarter tiles at `(2.758", 4.881")`, `(4.051", 4.881")`, `(5.392", 4.881")`, `(6.855", 4.881")` with heights `0.105"`.
  - Media share tiles starting at `left 0.179"`, `1.469"`, `2.758"` with widths `≈1.239"`.
  - Funnel share tiles at `left 6.028"`, `7.206"`, `8.385"` with widths `≈1.124"`.
- Footer text frame — `(0.088", 5.304", 3.585", 0.269")`.
- Legend group anchors (layout 0):
  - Television `(6.478", 0.313", width 0.539")`, Digital `(7.088", 0.313", width 0.868")`, OOH `(8.027", 0.313", width 0.648")`, Other `(8.745", 0.313", width 0.767")`.

## Conditional Elements (Disabled by Config)
- `CommentsTitle` — Comments header (currently blanked).
- `CommentsBox` — Comments content placeholder.
- `FunnelChart`, `MediaTypeChart`, `CampaignTypeChart` — Legacy chart shapes (unused in new flow).

## Derived Metrics Summary
- Quarterly tiles derive from monthly budget columns in the data frame (sum of three months, formatted as `£{value:,.0f}K`).
- Media share and funnel share tiles compute percentages against the market/brand/year total cost; values formatted as whole percentages.
- Footer appends `Data as of DD_MM_YY` using the Excel export date.

## Validation Contract
- Shape validator must find all named shapes above before generation proceeds.
- Missing optional elements (comments/charts) are ignored when disabled via configuration.
