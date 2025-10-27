# Archived Documentation & Logs

This directory contains archived documentation and logs from previous development sessions, organized for historical reference and potential recovery if needed.

**Archive Created:** 27-10-25 (Repository Cleanup - Tier 5)

---

## Directory Structure

### historical_sessions/
Complete session documentation from earlier development phases:
- `14_10_25/` - Initial setup and template analysis (Oct 14)
- `15_10_25/` - Campaign wrapping investigations (Oct 15)
- `17_10_25/` - Cell merge experimentation (Oct 17)
- `20_10_25/` - Validator foundation work (Oct 20)
- `21_10_25/` - Comprehensive visual analysis (Oct 21, 1.8MB)
- `22-10-25/` - Merged cells analysis (Oct 22)
- `23-10-25/` - QA and refinement work (Oct 23)

**Total archived size:** ~2.5MB

Each session directory contains:
- `artifacts/` - Generated presentations, validation reports, diffs
- `logs/` - Execution logs from that session
- `end_day/` - End-of-session documentation
- Other session-specific subdirectories

### legacy_powershell_scripts/
See `tools/archive/README_ARCHIVE.md` for details on deprecated PowerShell scripts.

### analysis_scripts/
See `tools/archive/README_ARCHIVE.md` for details on one-off analysis scripts.

### debug_scripts/
See `tools/archive/README_ARCHIVE.md` for details on debugging scripts.

---

## Active Documentation

Current development documentation is located at:
- `docs/27-10-25/` - Current session (active)
- `docs/24-10-25/` - Recent baseline (kept for reference, 54MB artifacts)

Specifications and plans:
- `openspec/project.md` - Current project specifications
- `AGENTS.md` - Current agent instructions
- `README.md` - Project overview

---

## Log File Reorganization

Production logs have been reorganized from flat timestamp-based structure to date-based organization:

```
logs/production/
├── 2025-10-14/     (12 execution runs)
├── 2025-10-15/     (2 execution runs)
├── 2025-10-17/     (4 execution runs)
├── 2025-10-19/     (2 execution runs)
├── 2025-10-20/     (43 execution runs)
├── 2025-10-21/     (62 execution runs)
├── 2025-10-22/     (13 execution runs)
├── 2025-10-23/     (7 execution runs)
├── 2025-10-24/     (14 execution runs)
└── 2025-10-27/     (34 execution runs, current)
```

**Total:** 196 execution runs organized by date for easier navigation.

---

## If You Need Historical Context

1. **For early design decisions:** Check `14_10_25/` through `17_10_25/` artifacts and end_day docs
2. **For visual analysis:** See `21_10_25/differential_visual_analysis/`
3. **For merge strategy evolution:** Check `22-10-25/merged_cells_analysis/`
4. **For QA baseline:** Review `23-10-25/` for pre-stabilization state
5. **For old logs:** Navigate to `logs/production/2025-10-XX/` by date

---

## Archive Integrity

- ✅ All archived files preserved without modification
- ✅ Metadata (timestamps, file structure) maintained
- ✅ Cross-references documented in this README
- ✅ Active development continues in `docs/27-10-25/` and main branch

**Note:** These archives are read-only historical records. Do not modify without explicit reason.

---

## Cleanup Summary (27-10-25)

**Tier 1-2: Deletions**
- Deleted 2 backup PowerShell files that were duplicates

**Tier 3: Major Archives**
- Moved 7 old session directories (14-23 Oct) to `historical_sessions/`
- Total space freed: ~2.5MB in active docs/

**Tier 4: Script Archives**
- Archived 10 legacy PowerShell scripts
- Archived 8 analysis scripts
- Archived 1 debug script

**Tier 5: Log Reorganization**
- Reorganized 196 production logs from flat to date-based structure
- Improved navigability and maintainability

**Status:** ✅ Archive complete

---

## Maintenance

If you need to restore any archived content:
1. Copy from appropriate archive directory
2. Update documentation and imports as needed
3. Add to commit with clear explanation of restoration reason
4. Document why it was restored

**Remember:** These were archived because they were no longer needed for active development.
