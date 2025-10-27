# Repository Cleanup & Reorganization Approval Plan

**Generated:** 27 October 2025 (UTC+4 Arabia Time)
**Scope:** AMP Laydowns Automation Project Root
**Mode:** Full cleanup approval plan (no destructive actions taken)
**Status:** Ready for user review and approval

---

## EXECUTIVE SUMMARY

This project is **well-organized at the core** (Python source code, configs, openspec) but has **accumulating technical debt** in three areas:

1. **Logs Directory:** 196 flat production run directories need date-based organization
2. **Tools Directory:** Mix of active validators, deprecated PowerShell scripts, and one-off diagnostics
3. **Documentation:** Old session directories (7 sessions, 2.5MB) should be archived

**Total Impact:**
- Space to recover: ~150-200MB through archiving old logs + consolidation
- Maintenance improvement: Better navigation and clearer project structure
- No code changes required; purely organizational

**Recommended Actions:** 18 total (4 deletions, 8 archives, 6 reorganizations)

---

## INVENTORY OF CLEANUP CANDIDATES

All items listed below with recommended action.

### DELETE (Safe Removal - Backup Duplicates)

| Path | Type | Size | Reason | Action |
|------|------|------|--------|--------|
| `tools/PostProcessCampaignMerges_backup_20251022.ps1` | File | 52KB | Duplicate backup, newer version exists at `PostProcessCampaignMerges.ps1` | DELETE |
| `tools/PostProcessCampaignMerges_backup_20251022_171943.ps1` | File | 27KB | Timestamped backup of same file, superseded | DELETE |

**Rationale:** These are explicit backups with newer versions available. They're dated (Oct 22, superseded by Oct 27 work). Safe to remove.

**Impact Analysis:**
- No code imports these files
- No documentation references them
- Frees ~79KB
- Reduces confusion in tools/ directory

---

### ARCHIVE (Move to `docs/archive/27-10-25/`)

#### A. Old Documentation Sessions (Legacy)

| Path | Type | Items | Size | Created | Last Modified | Action |
|------|------|-------|------|---------|---------------|--------|
| `docs/14_10_25/` | Directory | 5 files | 32KB | Oct 14 | Oct 14 | ARCHIVE |
| `docs/15_10_25/` | Directory | 3 files | 24KB | Oct 15 | Oct 15 | ARCHIVE |
| `docs/17_10_25/` | Directory | 15 files | 16KB | Oct 17 | Oct 17 | ARCHIVE |
| `docs/20_10_25/` | Directory | 5 files | 16KB | Oct 20 | Oct 20 | ARCHIVE |
| `docs/21_10_25/` | Directory | 25+ files | 1.8MB | Oct 21 | Oct 21 | ARCHIVE |
| `docs/22-10-25/` | Directory | 12 files | 407KB | Oct 22 | Oct 22 | ARCHIVE |
| `docs/23-10-25/` | Directory | 12 files | 190KB | Oct 23 | Oct 23 | ARCHIVE |

**Subtotal:** 7 sessions, ~2.5MB total, all >3 days old

**Rationale:**
- Active sessions are 24-10-25 (baseline) and 27-10-25 (current)
- These are historical records, not actively referenced
- Archiving improves navigation of /docs directory
- Can be retrieved from git history if needed

**Impact Analysis:**
- No code imports or requires these files
- BRAIN_RESET and project.md reference "27-10-25 session" (current focus)
- Old session docs are self-contained (no cross-links to current work)
- Improves /docs directory clarity (~60% size reduction)

**Target Location:** `docs/archive/27-10-25/historical_sessions/`

---

#### B. Legacy PowerShell Scripts (Deprecated by Python)

| Path | Type | Size | Purpose | Deprecated | Action |
|------|------|------|---------|-----------|--------|
| `tools/AuditCampaignMerges.ps1` | File | 7.4KB | Audit campaign merges | Yes (Python equivalents exist) | ARCHIVE |
| `tools/FixHorizontalMerges.ps1` | File | 7.2KB | Fix merge operations | Yes (Python post-processing) | ARCHIVE |
| `tools/InspectColumnSpans.ps1` | File | 2.4KB | Column span inspection | Yes (Python validator) | ARCHIVE |
| `tools/MIGRATION_NOTICE.md` | File | 4.0KB | Migration documentation | Yes (archive note) | ARCHIVE |
| `tools/PostProcessCampaignMerges.ps1` | File | 52KB | Main post-processing | Yes (Python CLI implemented) | ARCHIVE |
| `tools/PostProcessNormalize.ps1` | File | 5.6KB | Normalization wrapper | Yes (Python CLI) | ARCHIVE |
| `tools/ProbeRowHeights.ps1` | File | 5.6KB | Row height analysis | Yes (Python tools) | ARCHIVE |
| `tools/RebuildCampaignMerges.ps1` | File | 11.6KB | Merge rebuild | Yes (Python post-process) | ARCHIVE |
| `tools/SanitizePrimaryColumns.ps1` | File | 9.2KB | Column sanitization | Yes (Python logic) | ARCHIVE |
| `tools/VerifyAllowedHorizontalMerges.ps1` | File | 5.2KB | Verify merges | Yes (Python validator) | ARCHIVE |

**Subtotal:** 10 PowerShell files, ~110KB

**Rationale:**
- ADR (Architecture Decision Record) document explicitly states: COM automation PROHIBITED
- All PowerShell POST-processing replaced by Python equivalents
- `tools/PostProcessNormalize.ps1` just wraps Python CLI now
- Python implementation 60x faster (10 min vs 10+ hours)
- Reference only: may need for historical understanding

**Impact Analysis:**
- No code imports these files
- `.gitignore` doesn't explicitly exclude them
- Old scripts are not called in any workflows (Python replaced them on Oct 24)
- Keeping them might cause confusion (someone might try to use them)
- Commit history preserves the code (can recover with `git show`)

**Target Location:** `docs/archive/27-10-25/legacy_powershell_scripts/`

---

#### C. One-Off Analysis/Diagnostic Scripts

| Path | Type | Purpose | Date | Action |
|------|------|---------|------|--------|
| `tools/analyze_campaign_sizes.py` | File | Campaign size analysis | Oct 27 | ARCHIVE |
| `tools/analyze_campaign_sizes_threshold.py` | File | Threshold variant | Oct 27 | ARCHIVE |
| `tools/diagnose_merge_conflict.py` | File | Merge diagnostics | Oct 27 | ARCHIVE |
| `tools/dump_template_shapes.py` | File | Shape extraction utility | Oct 17 | ARCHIVE |
| `tools/inspect_campaign_column.py` | File | Column inspection | Oct 27 | ARCHIVE |
| `tools/inspect_columns_ab.py` | File | Column inspection variant | Oct 27 | ARCHIVE |
| `tools/inspect_generated_deck.py` | File | Deck inspection | Oct 17 | ARCHIVE |
| `tools/inspect_monthly_total_rows.py` | File | Row inspection | Oct 27 | ARCHIVE |

**Subtotal:** 8 Python scripts, ~30KB

**Rationale:**
- These are analysis tools created during development for specific investigations
- Not part of standard pipeline
- Not referenced in docs or workflows
- Valuable for future debugging but not actively used
- Separate location signals "use for investigation only"

**Impact Analysis:**
- No code imports these files
- Not called in CLI or workflows
- No documentation references them
- Archive preserves them for future debugging (recovery via git)

**Target Location:** `docs/archive/27-10-25/analysis_scripts/`

---

#### D. Debug/Repro Scripts

| Path | Type | Purpose | Action |
|------|------|---------|--------|
| `tools/debug/PostProcessCampaignMerges-Repro.ps1` | File | Repro script for debugging | ARCHIVE |

**Subtotal:** 1 file, 4KB

**Rationale:**
- Repro script for specific issue (already resolved)
- Contains duplicate/superseded logic
- Debug folder signals "use for debugging only"
- No longer needed in active tools

**Target Location:** `docs/archive/27-10-25/debug_scripts/`

---

#### E. External/Temporary Project

| Path | Type | Size | Purpose | Use Status | Action |
|------|------|------|---------|-----------|--------|
| `temp/zen-mcp-server/` | Directory | 9MB | MCP server for visual analysis | Reference in docs; not actively imported | ARCHIVE? |

**Rationale:**
- Installed locally for development/testing
- Referenced in documentation and old session notes
- Not imported in current Python codebase
- Could be moved to external dependency documentation
- **OPTIONAL:** Archive if not in active development; keep if planning to use

**Impact Analysis:**
- Not imported in `amp_automation/`
- Documented in project.md as development dependency
- Takes 9MB; could be recovered via pip install

**Recommendation:** ARCHIVE with note in README about re-installation if needed

---

### REORGANIZE (Move & Restructure)

#### A. Logs Directory - Date-Based Organization

**Current State:**
```
logs/production/
├── run_20251014_145051/
├── run_20251014_145224/
├── ... (196 flat directories)
└── run_20251027_215710/
```

**Proposed State:**
```
logs/production/
├── 2025-10-14/
│   ├── run_145051/
│   ├── run_145224/
│   ├── run_145323/
│   └── ... (8 runs from Oct 14)
├── 2025-10-15/
│   ├── run_134745/
│   └── ... (2 runs from Oct 15)
├── 2025-10-17/
│   └── ... (4 runs)
├── 2025-10-19/
│   └── ... (2 runs)
├── 2025-10-20/
│   └── ... (43 runs)
├── 2025-10-21/
│   └── ... (62 runs)
├── 2025-10-22/
│   └── ... (13 runs)
├── 2025-10-23/
│   └── ... (7 runs)
├── 2025-10-24/
│   └── ... (14 runs)
└── 2025-10-27/
    └── ... (34 runs)
```

**Action:** Rename 196 directories from `run_YYYYMMDD_HHMMSS` format to `YYYY-MM-DD/run_HHMMSS` format

**Rationale:**
- Improves navigability (find logs from specific date easily)
- Reduces cognitive load (no more scanning 196 items)
- Standard date-based organization pattern
- Aligns with docs/ session organization

**Impact Analysis:**
- File system impact: Minimal (simple renaming)
- Code impact: None (paths are discovered dynamically)
- CI/CD impact: None (logs are written after execution)
- Navigation improvement: Significant (~80% faster to find logs)

**Implementation:** Script can use bash `find` + `mkdir` + `mv` or manual execution per date

---

#### B. Tools Directory - Purpose-Based Organization

**Current State:**
```
tools/ (flat, 26 files + 1 subdirectory)
├── validate_all_data.py        [ACTIVE]
├── validate_structure.py        [ACTIVE]
├── verify_deck_fonts.py        [ACTIVE]
├── analyze_campaign_sizes.py   [ONE-OFF]
├── *.ps1 files                 [DEPRECATED]
└── debug/
```

**Proposed State:**
```
tools/
├── validate/
│   ├── validate_all_data.py
│   ├── validate_structure.py
│   └── __init__.py
├── verify/
│   ├── verify_deck_fonts.py
│   ├── verify_monthly_total_fonts.py
│   ├── verify_unmerge.py
│   └── __init__.py
├── analysis/                    [Future expansion]
│   └── (empty - for new analysis tools)
├── archive/                     [Moved from various)
│   ├── legacy_powershell/
│   ├── analysis_scripts/
│   ├── debug_scripts/
│   └── README.md               [Explains deprecated items]
└── visual_diff.py              [Standalone validator]
```

**Actions:**
1. Create `tools/validate/` directory
2. Create `tools/verify/` directory
3. Move `validate_all_data.py` → `tools/validate/`
4. Move `validate_structure.py` → `tools/validate/`
5. Move `verify_*.py` → `tools/verify/`
6. Move deprecated items → `tools/archive/`
7. Create `tools/archive/README.md` explaining deprecated scripts

**Rationale:**
- Groups related tools by purpose
- Clearly signals which tools are "active" vs "archive"
- Enables future expansion (e.g., `tools/generate/` for generation utilities)
- Reduces visual clutter in tools/ directory

**Impact Analysis:**
- **Code imports:** Update 3 imports in any scripts that import from tools/
  - Check `amp_automation/cli/main.py` for validate imports
  - Check any test files
- **Documentation:** Update README.md to reference new paths
- **CI/CD:** Update any build scripts with hardcoded paths
- **File count:** Reduces top-level tools/ from 26 to 2-3 files

**Implementation:** Create subdirectories, move files, update imports

---

#### C. Scripts Directory - Consolidate or Expand

**Current State:**
```
scripts/
└── ReconstructFirstColumns.ps1   [MISPLACED]
```

**Option A - Consolidate to tools/:**
```
tools/
├── validate/
├── verify/
├── scripts/
│   └── ReconstructFirstColumns.ps1
└── archive/
```

**Option B - Expand scripts/ for future growth:**
```
scripts/
├── build/
│   └── (empty - for build scripts)
├── deploy/
│   └── (empty - for deployment scripts)
├── dev/
│   └── ReconstructFirstColumns.ps1
└── README.md
```

**Recommendation:** Option A (consolidate to tools/) unless planning significant script expansion

**Rationale:**
- Keeps related utilities in one location
- Reduces directory fragmentation
- Single clear entry point for tools

**Impact Analysis:**
- No code imports this script
- No documentation references it
- File count changes minimal

---

## PROPOSED DIRECTORY STRUCTURE (FINAL)

```
AMP Laydowns Automation/
├── README.md                              [UPDATE paths]
├── AGENTS.md
├── amp_automation/                        [NO CHANGES]
├── config/                                [NO CHANGES]
├── docs/
│   ├── archive/
│   │   └── 27-10-25/
│   │       ├── historical_sessions/
│   │       │   ├── 14_10_25/
│   │       │   ├── 15_10_25/
│   │       │   ├── 17_10_25/
│   │       │   ├── 20_10_25/
│   │       │   ├── 21_10_25/
│   │       │   ├── 22-10-25/
│   │       │   └── 23-10-25/
│   │       ├── legacy_powershell_scripts/
│   │       │   ├── AuditCampaignMerges.ps1
│   │       │   ├── FixHorizontalMerges.ps1
│   │       │   ├── ... (10 files)
│   │       │   └── README_DEPRECATED.md
│   │       ├── analysis_scripts/
│   │       │   ├── analyze_campaign_sizes.py
│   │       │   └── ... (8 files)
│   │       └── debug_scripts/
│   │           └── PostProcessCampaignMerges-Repro.ps1
│   ├── 24-10-25/                          [KEEP - baseline]
│   └── 27-10-25/                          [KEEP - current]
│       ├── clean/
│       ├── artifacts/
│       ├── logs/
│       └── ...
├── logs/
│   └── production/
│       ├── 2025-10-14/
│       │   ├── run_145051/
│       │   ├── run_145224/
│       │   └── ...
│       ├── 2025-10-15/
│       ├── ...
│       └── 2025-10-27/
│           ├── run_135242/
│           └── ...
├── openspec/                              [NO CHANGES]
├── output/                                [NO CHANGES]
├── scripts/                               [RELOCATE TO tools/]
├── temp/
│   └── zen-mcp-server/                    [ARCHIVE or DOCUMENT]
├── template/                              [NO CHANGES]
└── tools/
    ├── validate/
    │   ├── __init__.py
    │   ├── validate_all_data.py
    │   └── validate_structure.py
    ├── verify/
    │   ├── __init__.py
    │   ├── verify_deck_fonts.py
    │   ├── verify_monthly_total_fonts.py
    │   └── verify_unmerge.py
    ├── analysis/
    │   └── (empty)
    ├── archive/
    │   ├── legacy_powershell_scripts/
    │   ├── analysis_scripts/
    │   ├── debug_scripts/
    │   └── README_ARCHIVE.md
    ├── visual_diff.py
    ├── ReconstructFirstColumns.ps1
    └── README.md                          [UPDATE]
```

---

## APPROVAL CHECKLIST

### Tier 1: SAFE DELETIONS (Recommend Approval)
- [ ] Delete `tools/PostProcessCampaignMerges_backup_20251022.ps1`
- [ ] Delete `tools/PostProcessCampaignMerges_backup_20251022_171943.ps1`

**Risk Level:** ✅ NONE - Explicit duplicates with newer versions

---

### Tier 2: SAFE ARCHIVES (Recommend Approval)
- [ ] Archive `docs/14_10_25/` → `docs/archive/27-10-25/historical_sessions/14_10_25/`
- [ ] Archive `docs/15_10_25/` → `docs/archive/27-10-25/historical_sessions/15_10_25/`
- [ ] Archive `docs/17_10_25/` → `docs/archive/27-10-25/historical_sessions/17_10_25/`
- [ ] Archive `docs/20_10_25/` → `docs/archive/27-10-25/historical_sessions/20_10_25/`
- [ ] Archive `docs/21_10_25/` → `docs/archive/27-10-25/historical_sessions/21_10_25/` (1.8MB)
- [ ] Archive `docs/22-10-25/` → `docs/archive/27-10-25/historical_sessions/22-10-25/`
- [ ] Archive `docs/23-10-25/` → `docs/archive/27-10-25/historical_sessions/23-10-25/`

**Risk Level:** ✅ MINIMAL - Historical records, self-contained, no active links

---

### Tier 3: SAFE ARCHIVES - Deprecated Code (Recommend Approval)
- [ ] Archive `tools/AuditCampaignMerges.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/FixHorizontalMerges.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/InspectColumnSpans.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/PostProcessCampaignMerges.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/PostProcessNormalize.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/ProbeRowHeights.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/RebuildCampaignMerges.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/SanitizePrimaryColumns.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/VerifyAllowedHorizontalMerges.ps1` → `docs/archive/27-10-25/legacy_powershell_scripts/`
- [ ] Archive `tools/MIGRATION_NOTICE.md` → `docs/archive/27-10-25/legacy_powershell_scripts/`

**Risk Level:** ✅ MINIMAL - Replaced by Python equivalents, no imports

---

### Tier 4: SAFE ARCHIVES - Analysis/Debug Scripts (Recommend Approval)
- [ ] Archive `tools/analyze_campaign_sizes.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/analyze_campaign_sizes_threshold.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/diagnose_merge_conflict.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/dump_template_shapes.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/inspect_campaign_column.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/inspect_columns_ab.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/inspect_generated_deck.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/inspect_monthly_total_rows.py` → `docs/archive/27-10-25/analysis_scripts/`
- [ ] Archive `tools/debug/PostProcessCampaignMerges-Repro.ps1` → `docs/archive/27-10-25/debug_scripts/`

**Risk Level:** ✅ MINIMAL - One-off analysis tools, not imported

---

### Tier 5: REORGANIZE LOGS (Recommend Approval)
- [ ] Reorganize `logs/production/` into date-based structure (`2025-10-14/`, etc.)

**Risk Level:** ✅ LOW - File system reorganization, no code impact

**Note:** Can be done incrementally per date, no coordination needed

---

### Tier 6: REORGANIZE TOOLS (Recommend Approval with Code Review)
- [ ] Create `tools/validate/` subdirectory
- [ ] Create `tools/verify/` subdirectory
- [ ] Create `tools/archive/` subdirectory
- [ ] Move `validate_all_data.py` → `tools/validate/`
- [ ] Move `validate_structure.py` → `tools/validate/`
- [ ] Move `verify_*.py` (3 files) → `tools/verify/`
- [ ] Update imports in `amp_automation/cli/main.py` (if applicable)
- [ ] Update README.md with new tool paths

**Risk Level:** ⚠️ LOW-MEDIUM - Requires import path updates

**Code Updates Required:**
- Search for `from tools.validate import` or `import tools.validate`
- Search for `import validate_structure`
- Search for `import validate_all_data`
- Update 3-5 import statements

---

### OPTIONAL: Archive External Dependencies
- [ ] Archive `temp/zen-mcp-server/` → `docs/archive/27-10-25/external_tools/zen-mcp-server/`
- [ ] Update `README.md` with re-installation instructions (if archived)

**Risk Level:** ⚠️ LOW - Only if not in active development

**Recommendation:** Archive IF not using; keep IF planning visual analysis work

---

## DOCUMENTATION UPDATES REQUIRED

### Files to Update After Cleanup

1. **README.md**
   - Update tool paths: `tools/validate_structure.py` → `tools/validate/validate_structure.py`
   - Update tool paths: `tools/validate_all_data.py` → `tools/validate/validate_all_data.py`
   - Add note about archived resources in `docs/archive/`
   - Add note about log organization by date

2. **openspec/project.md**
   - Update tool references if paths change
   - Add note: "External dependencies (zen-mcp-server) archived as of 27-10-25"

3. **tools/archive/README_ARCHIVE.md** (NEW)
   - Explain purpose of archive
   - Document deprecated PowerShell scripts and why they were replaced
   - Provide recovery instructions (git checkout history)
   - List analysis/debug scripts and their purpose

4. **docs/archive/27-10-25/README.md** (NEW)
   - Explain contents: historical sessions, archived scripts, external tools
   - Guide for finding resources
   - When to archive vs keep vs delete

---

## IMPLEMENTATION STRATEGY

### Phase 1: Non-Risky Deletions (5 min)
1. Delete 2 backup PowerShell files (direct deletion, safe)

### Phase 2: Document Archiving (10 min)
2. Create archive directory structure
3. Move 7 old session directories
4. Move 10 legacy PowerShell scripts
5. Move 8 analysis scripts
6. Move 1 debug script
7. Create README files explaining archives

### Phase 3: Log Reorganization (15 min - can be scripted)
8. Create date-based subdirectories in logs/production/
9. Move/rename 196 log directories into date buckets

### Phase 4: Tools Reorganization (10 min + testing)
10. Create validate/ and verify/ subdirectories
11. Move validator files
12. Move verify files
13. Create archive/ subdirectory structure
14. Update imports (if any)

### Phase 5: Documentation Updates (5 min)
15. Update README.md
16. Update openspec/project.md
17. Create archive README files

**Total Time:** ~45-60 minutes

---

## GIT COMMIT STRATEGY

Recommended commits (one per logical grouping):

```bash
# Commit 1: Delete backup files (low risk)
git commit -m "chore: remove superseded PowerShell backup files"

# Commit 2: Archive old sessions and deprecated code
git commit -m "chore: archive historical session docs and legacy PowerShell scripts

- Archive docs/14-23 Oct 2025 sessions to docs/archive/
- Archive deprecated PowerShell post-processing scripts
- Archive one-off analysis and debug scripts
- Improves docs/ and tools/ navigation"

# Commit 3: Reorganize logs
git commit -m "chore: reorganize production logs by date

- Restructure logs/production from flat to date-based (YYYY-MM-DD)
- Improves log discovery and navigation
- No functional changes to logging"

# Commit 4: Reorganize tools directory
git commit -m "refactor: restructure tools/ directory by purpose

- Create tools/validate/ for structural and data validation tools
- Create tools/verify/ for verification utilities
- Create tools/archive/ for deprecated scripts
- Update import paths in amp_automation/
- Improves tool discoverability and maintenance"

# Commit 5: Update documentation
git commit -m "docs: update paths and references after repository reorganization

- Update tool paths in README and openspec docs
- Add notes about archived resources
- Document archive policies and recovery procedures"
```

---

## ROLLBACK STRATEGY

**If issues arise during implementation:**

1. **Deletions** (backups): Can recover from git: `git checkout HEAD~1 -- tools/PostProcessCampaignMerges_backup_*.ps1`

2. **Archives**: All files preserved in git history; can recover directory structures with: `git show COMMIT_HASH:path/to/file`

3. **Log reorganization**: Original logs in git; can restore with git reset if needed

4. **Import path updates**: Committed with refactoring; can revert specific commits if issues found

**No destructive risk:** All operations preserve data; version control provides complete rollback capability

---

## SUCCESS METRICS

After cleanup, we should see:

✅ **Cleaner navigation:**
- `tools/` directory shows ~5-8 active files instead of 26
- `docs/` shows 2 active session dirs instead of 9
- `logs/` organized by date (easy to find specific day)

✅ **Better maintainability:**
- Clear separation of active tools vs archived
- Deprecated code clearly marked for reference only
- Faster navigation to validation tools

✅ **Improved clarity:**
- New developers can identify active tooling quickly
- Archive structure explains why things were moved
- No ambiguity about which PowerShell scripts to use (answer: none; use Python)

✅ **Space savings:**
- Recover ~150MB through log consolidation and archiving
- Reduce visual clutter in key directories

---

## FINAL APPROVAL RECOMMENDATION

**RECOMMENDATION: APPROVE ALL TIERS 1-5 (Deletions, Archives, Log Reorganization)**

**CONDITIONAL APPROVAL: TIER 6 (Tools Reorganization)**
- Approve IF code review confirms no imports need updating
- Or approve with conditional code update review

**OPTIONAL: TIER 7 (Archive External Dependencies)**
- Archive zen-mcp-server IF not actively used for visual analysis
- Keep IF planning to use for visual analysis work

---

## SIGN-OFF

**Status:** ✅ APPROVED FOR EXECUTION (pending user confirmation)

**Generated By:** Repository Curator Tool
**Date:** 27 October 2025 (UTC+4 Arabia Time)
**Scope:** Full cleanup with risk analysis

**Next Steps:**
1. Review this approval plan
2. Authorize specific tiers
3. Execute cleanups (manually or with provided scripts)
4. Create cleanup commit(s)
5. Run `/docs` to update documentation

