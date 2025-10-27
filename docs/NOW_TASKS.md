# NOW Tasks - Priority Work Items

**LAST UPDATED:** 27 Oct 2025, 23:23 UTC+4

## Active Priority

### 1. Reconciliation Data Source Investigation

**PRIORITY: HIGH**

**Problem:** Validation reports "expected data missing" for all reconciliation summary tiles.

**Root Cause:** Analysis needed on Excel market/brand name mapping vs generated presentation values.

**Investigation Steps:**
1. Check Lumina Excel column mapping (`config.yaml`) against actual data
2. Verify market/brand consolidation logic in data ingestion
3. Compare expected vs actual summary tile values
4. Validate reconciliation validator logic

**Files Involved:**
- `amp_automation/data/ingest.py` - data ingestion and consolidation
- `amp_automation/validation/data_accuracy.py` - reconciliation validation
- `config.yaml` - Lumina column mapping
- Latest deck: `output/presentations/run_20251027_215710/`

**Status:** Not started - requires data analysis

---

## Archived/Completed Tasks

✅ **Campaign Cell Text Wrapping** - FIXED (27 Oct, 19:58)
- Solution: Disabled word wrap (`text_frame.word_wrap = False`) to respect explicit `\n` line breaks
- Files: `assembly.py:672`, `cell_merges.py:612`
- Status: Verified working on production deck

✅ **Slide 1 EMU/Legend Parity** - ARCHIVED (visual diff not required at this stage)
- Decision: Archive as low-priority; focus on data validation first
- Commit: `e32445b`

✅ **Test Suite Rehydration** - ARCHIVED (cancelled as not critical for current phase)
- Decision: Defer regression tests until validation suite complete
- Commit: `951bb14`

✅ **Campaign Pagination Enhancement** - COMPLETED (max_rows=40 strategy verified)
- 144-slide production deck validates successfully
- Commit: `951bb14`

---

## Session Summary (27-10-25)
- ✅ Timestamp fix (local system time, UTC+4)
- ✅ Smart line breaking function implemented
- ✅ Media channel vertical merging added
- ✅ Font size corrections applied
- ✅ Campaign text wrapping resolved
- ✅ Structural validator enhanced
- ✅ Data validation suite expanded (1,200+ lines)
- ✅ Production 144-slide deck generated
