## Why
- Hand-built table rendering diverges sharply from the Template V4 baseline (missing row groups, flattened widths, mismatched fills), producing slides that fail client QA.
- Pixel fidelity is non-negotiable; we need to lean on the master template’s geometry instead of recreating it procedurally.

## What Changes
- Introduce a template cloning layer that copies the `MainDataTable`, summary tiles, and legend groups from the master slide and reuses the native shape layout.
- Replace bespoke table construction with a data-population routine that edits cloned shapes in-place (cell text, fills, font overrides only when data requires it).
- Add regression tooling to snapshot cloned slides and assert geometry parity (visual diff + structural checks) across representative campaigns.
- Gate AutoPPTX output behind the new workflow (optionally disabling the legacy adapter once parity is proven).

## Impact
- Affected specs: presentation rendering pipeline, tooling automation.
- Affected code: `amp_automation/presentation/assembly.py`, new cloning helpers under `amp_automation/presentation/`, config toggles (`config/master_config.json`), visual diff tooling & tests.
