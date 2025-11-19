## ADDED Requirements

### Requirement: Clone-based slide assembly
The system SHALL generate each campaign slide by cloning the master template’s `MainDataTable` and associated summary shapes before populating data.

#### Scenario: Populate cloned slide successfully
- **GIVEN** the presentation generator runs with clone mode enabled
- **WHEN** a campaign slide is produced
- **THEN** the table and summary tiles originate from cloned template shapes
- **AND** only text/values differ from the baseline template
- **AND** geometry (position, width, height) matches the master slide within ±1 EMU.

### Requirement: Visual parity validation
The system SHALL execute automated visual diffs across generated slides and fail the build if mean or RMS pixel differences exceed configured thresholds.

#### Scenario: Visual diff regression detected
- **GIVEN** visual diff runners process the latest deck against the baseline template
- **WHEN** any slide’s RMS difference exceeds the threshold (default 0.5)
- **THEN** the build exits non-zero and logs the offending slide index.

### Requirement: Legacy fallback toggle
The system SHALL expose a configuration toggle that reverts to the legacy builder for troubleshooting while the clone pipeline is stabilized.

#### Scenario: Operator disables clone pipeline
- **GIVEN** the operator sets `presentation.clone_pipeline.enabled` to `false`
- **WHEN** the CLI runs
- **THEN** the legacy assembly path executes without invoking cloning helpers.
