## Why
- Python post-processing module was built under the assumption that merge operations needed to be applied as a separate step after deck generation.
- Testing revealed that clone pipeline already creates correct cell merges during generation (assembly.py:629,649).
- Post-processing merge operations fail with "range contains one or more merged cells" because cells are already correctly merged.
- Current architecture is unclear about which phase owns merge responsibility, leading to redundant operations.

## What Changes
- **Clarify merge ownership**: Generation phase (clone pipeline) owns all primary cell merge operations.
- **Reposition Python post-processing**: Focus on table normalization, cell formatting, and edge case repairs (not primary merges).
- **Update COM prohibition guidance**: COM restriction applies to bulk post-processing operations, not generation-time merges.
- **Simplify post-processing CLI**: Remove or document redundant merge operations as edge-case repair tools.
- **Document architecture decision**: Create ADR explaining merge phase separation and rationale.

## Impact
- Affected specs: Post-processing architecture, COM prohibition guidelines, pipeline phase responsibilities.
- Affected code:
  - `amp_automation/presentation/assembly.py` - Already implements merges correctly (no changes needed).
  - `amp_automation/presentation/postprocess/` - CLI and documentation updates to clarify scope.
  - `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` - Clarify COM restriction applies to post-processing.
  - `tools/PostProcessCampaignMerges.ps1` - Deprecate or repurpose for normalization only.
- Affected tests: None (generation merges already working and tested).

## Success Criteria
- ✅ Generation creates correct merges (verified via XML inspection: rowSpan/vMerge attributes).
- ✅ Python post-processing completes normalization in <1 minute for 88-slide deck.
- ✅ Documentation clearly states merge ownership (generation) vs. post-processing scope (normalization).
- ⏭️ End-to-end test: generation → Python normalization → validation passes without errors.
- ⏭️ PowerShell scripts updated to call Python CLI for normalization only (or deprecated).

## Timeline
- **Immediate**: Document architecture clarification (completed: commits d3e2b98, 8320c3f, 3c54b1b).
- **Short-term** (next session): Update PowerShell integration, run end-to-end test.
- **Medium-term**: Deprecate redundant PowerShell COM scripts, expand Python normalization coverage.

## Risks & Mitigation
- **Risk**: Generation merge logic might fail in edge cases (continuation slides, special campaigns).
  - **Mitigation**: Keep Python merge operations as repair tools for edge cases; add regression tests.
- **Risk**: Existing PowerShell workflows depend on post-processing merges.
  - **Mitigation**: Audit PowerShell scripts, update to call Python normalization only, document migration.
- **Risk**: Future template changes might break generation merges.
  - **Mitigation**: Add structural validation tests to catch merge regressions early.

## Notes
- This proposal is retroactive - it documents an architecture insight discovered during Python implementation.
- No breaking changes - generation merges already working correctly.
- Performance validated: Python normalization of 88 slides in ~30 seconds (vs 10+ hours PowerShell COM).
- Related ADR: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md` (COM prohibition for bulk operations).
- Discovery document: `docs/24-10-25/15-merge_architecture_discovery.md` (detailed analysis).
