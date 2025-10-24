# Tasks for Post-Processing Architecture Clarification

## Status: In Progress
- **Created**: 2025-10-24
- **Last Updated**: 2025-10-24 17:10
- **Owner**: Architecture Team

## Completed Tasks

### Phase 1: Discovery & Documentation (Completed 2025-10-24)
- [x] **Implement Python cell merge operations** (commit d3e2b98)
  - Campaign vertical merges
  - Monthly total horizontal merges
  - Summary row horizontal merges
  - Logging improvements and CLI integration

- [x] **Test Python post-processing on 88-slide deck** (2025-10-24 17:00)
  - Result: Normalization successful (~30 seconds)
  - Discovery: Merge operations fail because cells already merged
  - Performance: 60x faster than PowerShell COM baseline

- [x] **Analyze architecture and document findings** (commit 8320c3f)
  - Created `docs/24-10-25/15-merge_architecture_discovery.md`
  - Verified via XML inspection: generation creates correct merges
  - Identified merge ownership: generation (assembly.py:629,649)

- [x] **Update project documentation** (commit 3c54b1b)
  - Updated daily log (docs/24-10-25/24-10-25.md)
  - Updated BRAIN_RESET with corrected strategy
  - Marked completed TODOs, updated NOW/Next/Later sections

- [x] **Create OpenSpec proposal**
  - This document and proposal.md

## Pending Tasks

### Phase 2: Integration & Testing (Next Session)
- [ ] **Update PowerShell scripts to call Python CLI**
  - Modify `tools/PostProcessCampaignMerges.ps1` to use Python for normalization
  - Remove or document merge operations as edge-case repairs
  - Add error handling and logging integration

- [ ] **Run end-to-end pipeline test**
  - Generate fresh deck via CLI
  - Apply Python normalization (no merge operations)
  - Run structural validation
  - Verify merge correctness and table formatting
  - Document results and any issues found

- [ ] **Update COM prohibition ADR**
  - Clarify that COM restriction applies to post-processing bulk operations
  - Note that generation-time merges are acceptable (not bulk operations)
  - Add guidance on when to use COM vs. python-pptx

### Phase 3: Cleanup & Standardization (Future)
- [ ] **Audit and deprecate redundant PowerShell scripts**
  - Identify scripts that perform merge operations
  - Migrate to Python CLI or deprecate
  - Update runbooks and documentation

- [ ] **Expand Python normalization coverage**
  - Add row height normalization (if needed)
  - Add cell margin/padding normalization
  - Add font consistency checks
  - Document normalization operations in CLI help

- [ ] **Add regression tests for merge correctness**
  - Test that generation creates expected merges
  - Test merge behavior on continuation slides
  - Test edge cases (single-row campaigns, etc.)
  - Add to CI/CD pipeline

- [ ] **Create migration guide**
  - Document transition from PowerShell COM to Python
  - Provide examples for common operations
  - Update README and AGENTS.md with new workflow

## Success Metrics
- ✅ Python implementation completed and committed
- ✅ Architecture discovery documented
- ✅ Performance validated (<1 minute for 88 slides)
- ⏭️ End-to-end test passes without errors
- ⏭️ PowerShell scripts updated or deprecated
- ⏭️ Documentation reflects new architecture

## Dependencies
- Python 3.13+ with python-pptx
- Clone pipeline enabled in config/master_config.json
- Fresh deck from generation phase (with merges)

## Blockers
- None currently

## Notes
- This is a retroactive proposal documenting discovered architecture
- No breaking changes - generation already works correctly
- Focus is on clarification and cleanup, not new functionality
