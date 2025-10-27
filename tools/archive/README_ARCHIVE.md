# Archived Tools & Scripts

This directory contains deprecated, legacy, and one-off scripts that are no longer part of the active development workflow.

**Last Archive Update:** 27-10-25 (Tier 6 Tools Reorganization)

---

## Directory Structure

### legacy_powershell_scripts/
Contains deprecated PowerShell scripts (replaced by Python implementations):
- `PostProcessCampaignMerges.ps1` - Legacy cell merging logic (replaced by postprocess/cell_merges.py)
- `PostProcessNormalize.ps1` - Legacy normalization (replaced by postprocess/table_normalizer.py)
- `RebuildCampaignMerges.ps1` - Legacy merge rebuilding
- `ProbeRowHeights.ps1` - Legacy row height inspection
- `SanitizePrimaryColumns.ps1` - Legacy column sanitization
- `VerifyAllowedHorizontalMerges.ps1` - Legacy merge verification
- `FixHorizontalMerges.ps1` - Legacy horizontal merge fixes
- `InspectColumnSpans.ps1` - Legacy column span inspection
- `ReconstructFirstColumns.ps1` - Legacy first column reconstruction

**Reason for archival:** All functionality migrated to Python using python-pptx library for better maintainability, cross-platform compatibility, and integration with main pipeline.

### analysis_scripts/
One-off analysis and diagnostic scripts developed during development:
- `analyze_campaign_sizes.py` - Campaign size distribution analysis
- `analyze_campaign_sizes_threshold.py` - Threshold-based campaign analysis variant
- `inspect_campaign_column.py` - Campaign column structure inspection
- `inspect_columns_ab.py` - Columns A & B structure inspection
- `inspect_fonts.py` - Font usage analysis across presentations
- `inspect_generated_deck.py` - Generated deck structure inspection
- `inspect_monthly_total_rows.py` - Monthly total row analysis
- `dump_template_shapes.py` - Template shape enumeration

**Reason for archival:** These were used during development and debugging but are not part of the standard validation pipeline. Comprehensive validation is now handled by the unified tools/validate/ suite.

### debug_scripts/
Debugging and reproduction scripts:
- `PostProcessCampaignMerges-Repro.ps1` - Campaign merge issue reproduction

**Reason for archival:** Used for debugging specific issues but no longer needed after fixes were implemented.

---

## Active Tools

The active tools have been reorganized into purpose-based directories:

- **tools/validate/** - Data validation and structural validation
  - `validate_all_data.py` - Unified validation orchestration
  - `validate_structure.py` - PowerPoint structural contract validation

- **tools/verify/** - Verification and post-processing checks
  - `verify_deck_fonts.py` - Font verification across decks
  - `verify_monthly_total_fonts.py` - Specialized monthly total font checks
  - `verify_unmerge.py` - Cell unmerge verification

- **tools/** (root level) - Utility scripts
  - `visual_diff.py` - Visual comparison between presentations

---

## If You Need an Archived Script

1. **Find it in this directory** using the structure above
2. **Understand why it was archived** from the reason listed
3. **Check tools/validate/ or tools/verify/** for the modern replacement
4. **Reference the commit history** if you need to understand what the script did

**Never restore from archive without understanding why it was archived.**

---

## Maintenance Notes

- Backup files (`PostProcessCampaignMerges_backup_*.ps1`) have been deleted
- MIGRATION_NOTICE.md referenced these scripts; update if restoring any

**Status:** ✅ Archive complete (27-10-25)
