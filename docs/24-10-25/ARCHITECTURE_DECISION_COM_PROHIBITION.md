# Architecture Decision Record: PowerPoint COM Bulk Operations Prohibition

**Date:** 24 October 2025
**Status:** ✅ ACCEPTED - MANDATORY
**Severity:** 🚨 CRITICAL - DO NOT IGNORE

---

## 🚨 EXECUTIVE SUMMARY - READ THIS FIRST 🚨

**DO NOT use PowerPoint COM automation for bulk table operations.**

This decision is non-negotiable and based on extensive performance analysis showing COM-based bulk operations are **fundamentally unsuitable** for this use case, resulting in:

- **13x slower** performance (minimum)
- **Hour-long execution times** vs minutes with Python
- **Frequent hangs and timeouts**
- **Unreliable execution** with HRESULT errors

## Decision

**ALL bulk table operations MUST be performed using Python libraries (python-pptx, aspose_slides) instead of PowerPoint COM automation.**

PowerPoint COM (via pywin32/comtypes) is **ONLY** permitted for:
- Opening/closing presentation files
- File format conversions and exports
- Features not exposed by python-pptx (e.g., specific animations, macros)
- One-off operations on individual cells/shapes

## Context

### The Problem We Encountered

On 24 October 2025, we discovered that post-processing operations on an 88-slide presentation were taking **hours** and frequently timing out. Investigation revealed PowerPoint COM automation was the bottleneck.

### The Numbers

**Typical slide processing (32 rows × 18 columns):**

| Operation Type | COM Calls | Time per Slide | Full Deck (88 slides) |
|---------------|-----------|----------------|----------------------|
| Table normalization | ~1,280 | 35+ seconds (hanging) | **Hours** (never completed) |
| Cell formatting | ~576 | Included above | N/A |
| Span resets | ~96 | Included above | N/A |
| Merge operations | ~30-100 | 6+ minutes | **Hours** |
| **TOTAL (COM)** | **~2,000+** | **7+ minutes** | **10+ hours** ⚠️ |
| **Python equivalent** | **~10** | **2.66 seconds** | **~4 minutes** ✅ |

**Performance Difference:** Python is **~150x faster** for full post-processing.

## Root Cause Analysis

### Why COM Automation Failed

PowerPoint COM automation operates via **inter-process communication**:

1. **Each property access = separate COM call**
   ```powershell
   # PowerShell (COM) - Each line is a separate IPC call
   $cell = $table.Cell($row, $col)        # COM call 1
   $frame = $cell.Shape.TextFrame         # COM call 2
   $frame.AutoSize = 0                     # COM call 3
   $frame.MarginLeft = 0                   # COM call 4
   # ... 6+ more COM calls per cell
   ```

2. **Cumulative latency becomes catastrophic**
   - Single COM call: ~1-5ms overhead
   - 1,280 calls/slide × 5ms = 6.4 seconds minimum (just overhead!)
   - 88 slides × 6.4s = 563 seconds = **9.4 minutes** (overhead alone)

3. **COM is not thread-safe or batch-optimized**
   - Cannot parallelize operations
   - No way to batch multiple property changes
   - PowerPoint processes one call at a time

### Why Python Works

Python-pptx manipulates the **underlying OOXML directly**:

1. **Direct XML manipulation (no IPC)**
   ```python
   # Python (python-pptx) - Direct XML modification
   text_frame = cell.text_frame
   text_frame.auto_size = MSO_AUTO_SIZE.NONE  # Direct property set
   text_frame.margin_left = 0                  # Another direct set
   # No inter-process calls!
   ```

2. **Batch operations**
   - Multiple properties changed with single object access
   - XML serialized once at save time
   - Scales linearly with data size

3. **Predictable performance**
   - No COM overhead
   - No timeout issues
   - No HRESULT errors

## Evidence

### Test Results (24 Oct 2025)

**Single Slide (Slide 2) - Normalization Only:**

| Metric | PowerShell COM | Python (python-pptx) | Improvement |
|--------|---------------|---------------------|-------------|
| Time | 35+ seconds (hanging) | 2.66 seconds | **13x faster** |
| Success rate | 0% (timeout) | 100% | N/A |
| Memory usage | Growing (leak) | Stable | N/A |

**Full-Deck Projection (88 slides):**

| Operation | PowerShell COM | Python | Time Saved |
|-----------|---------------|---------|-----------|
| Normalize tables | Hours (never completed) | ~4 minutes | **Hours** |
| + Merge operations | N/A (timeout) | ~6-10 minutes | **Hours** |
| **TOTAL** | **Never completed** | **~10 minutes** | **Infinite** |

### Specific Issues Encountered

1. **Stalling on Slide 2** (24 Oct 2025, 15:48):
   - Script hung at slide 2 for 35+ minutes
   - Process still active but no progress
   - Memory usage climbing steadily
   - Had to be killed manually

2. **Plateau Detection Workaround** (24 Oct 2025, 14:47):
   - Added logic to detect when COM refuses to change row heights
   - Saved 3-4 minutes per slide by early-exiting
   - **Still took 6+ minutes per slide**
   - This was treating symptoms, not the root cause

3. **RPC_E_CALL_REJECTED Errors** (24 Oct 2025, 15:54):
   - COM errors during sanitization: "Call was rejected by callee"
   - PowerPoint refusing operations randomly
   - No way to predict or prevent these errors

4. **File Locking Issues**:
   - PowerPoint must be closed before running scripts
   - Failed operations leave PowerPoint processes hanging
   - Corrupts presentation files if not cleaned up properly

## Decision Rationale

### Why This Is Non-Negotiable

1. **Performance**: 13-150x improvement is not optional
2. **Reliability**: Python completes, COM hangs/fails
3. **Maintainability**: Python code is simpler and more debuggable
4. **Scalability**: Performance degrades linearly (Python) vs exponentially (COM)

### Acceptable COM Usage

COM is **permitted** for:

✅ Opening presentations:
```powershell
$app = New-Object -ComObject PowerPoint.Application
$prs = $app.Presentations.Open($path)
```

✅ Exporting to formats not supported by python-pptx:
```powershell
$prs.SaveAs($path, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsJPG)
```

✅ Accessing features not in python-pptx:
```powershell
# Animation timelines, embedded objects, macros, etc.
$slide.TimeLine.MainSequence.AddEffect(...)
```

### Prohibited COM Usage

🚫 Iterating through all cells:
```powershell
# BAD - DO NOT DO THIS
for ($row = 1; $row -le $rowCount; $row++) {
    for ($col = 1; $col -le $colCount; $col++) {
        $cell = $table.Cell($row, $col)  # COM call per cell
        # ... operations on cell
    }
}
```

🚫 Bulk property changes:
```powershell
# BAD - DO NOT DO THIS
foreach ($cell in $cells) {
    $cell.TextFrame.MarginLeft = 0     # COM call per cell
    $cell.TextFrame.MarginRight = 0    # Another COM call
    # ... etc
}
```

🚫 Cell merging in loops:
```powershell
# BAD - DO NOT DO THIS
for ($row = $start; $row -le $end; $row++) {
    $cell1 = $table.Cell($row, 1)
    $cell2 = $table.Cell($row, 2)
    $cell1.Merge($cell2)  # COM call + merge operation
}
```

## Implementation Guidance

### Correct Approach: Python Modules

All bulk operations are now in `amp_automation/presentation/postprocess/`:

```python
# CORRECT - Use Python
from amp_automation.presentation.postprocess import (
    normalize_table_layout,
    apply_blank_cell_formatting,
    merge_campaign_cells,
)

# Load presentation
prs = Presentation(path)

# Process slides
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            normalize_table_layout(table)
            apply_blank_cell_formatting(table)
            merge_campaign_cells(table)

# Save
prs.save(path)
```

### CLI Integration from PowerShell

```powershell
# Call Python CLI from PowerShell
python -m amp_automation.presentation.postprocess.cli `
    --presentation-path "path\to\deck.pptx" `
    --operations normalize,merge-campaign,merge-monthly `
    --verbose
```

## Consequences

### Benefits

✅ **Performance**: 10+ hours → 10 minutes
✅ **Reliability**: No timeouts, hangs, or HRESULT errors
✅ **Maintainability**: Simpler code, easier debugging
✅ **Testability**: Can test without PowerPoint installed
✅ **Portability**: Works on Linux/Mac (no PowerPoint needed)

### Drawbacks

⚠️ **Limited Feature Set**: python-pptx doesn't support all PowerPoint features
⚠️ **Learning Curve**: Team needs to learn python-pptx API
⚠️ **Migration Effort**: Existing PowerShell scripts need updating

### Mitigation

- Use aspose_slides for features python-pptx lacks
- Maintain minimal COM wrapper for file I/O only
- Document python-pptx patterns in codebase
- Gradual migration with compatibility layer

## Lessons Learned

### Key Insights

1. **Technology choice matters at scale**
   - COM was fine for prototyping (small decks)
   - Becomes catastrophic at production scale (88+ slides)

2. **Profile early, profile often**
   - We spent weeks optimizing COM (plateau detection, watchdogs, etc.)
   - Should have profiled and switched to Python immediately

3. **Treat symptoms vs root cause**
   - Plateau detection was treating symptoms
   - Python addressed the root cause

4. **Document architectural decisions**
   - This document exists to prevent repeating this mistake
   - Future developers must understand WHY this decision was made

### Red Flags for Future

If you encounter these symptoms, **stop using COM immediately**:

🚩 Scripts taking hours for operations that should be minutes
🚩 Frequent timeouts or hangs
🚩 HRESULT errors (RPC_E_CALL_REJECTED, etc.)
🚩 Memory usage climbing steadily during execution
🚩 Need for "watchdog" timeouts to prevent infinite hangs
🚩 Operations that "plateau" and stop making progress

## Enforcement

### Code Review Checklist

Reviewers MUST reject PRs that:

- [ ] Add new PowerShell COM loops over table cells
- [ ] Use COM for bulk property changes
- [ ] Iterate through slides/shapes/cells via COM
- [ ] Perform cell merges in loops via COM
- [ ] Do not have documented justification for COM usage

### Required Documentation for COM Usage

Any PR adding COM code MUST include:

1. **Justification**: Why is COM necessary? (Must be feature not in python-pptx)
2. **Performance Analysis**: Show that operation is O(1) or small constant
3. **Alternatives Considered**: Why can't python-pptx/aspose_slides be used?
4. **Fallback Plan**: What happens if COM fails?

### Approved Reviewers

COM usage PRs require approval from:
- Technical lead
- Developer who implemented Python migration

## Clarifications and Scope (Added 24 Oct 2025 17:20)

### This Prohibition Applies to POST-PROCESSING Only

**IMPORTANT**: This COM prohibition specifically targets **bulk post-processing operations** on already-generated decks. It does NOT prohibit:

1. **Generation-Time Merge Operations** (✅ ACCEPTABLE)
   - Cell merges created during deck generation (assembly.py:629, 649)
   - These are NOT bulk operations - they occur as tables are being built
   - Performance is acceptable because:
     - Merges happen once per table during construction
     - No repeated COM calls over existing cells
     - Part of the normal template cloning flow
     - Completes in minutes for full 88-slide deck

2. **File I/O Operations** (✅ ACCEPTABLE)
   - Opening/closing presentations
   - Saving presentations
   - Format conversions
   - These are O(1) operations, not O(n) bulk loops

### Generation vs. Post-Processing: Key Distinction

| Phase | COM Usage | Status | Rationale |
|-------|-----------|--------|-----------|
| **Generation** (assembly.py) | Cell merges during table creation | ✅ **ACCEPTABLE** | One-time operations during construction; not bulk post-processing |
| **Post-Processing** (PowerShell scripts) | Bulk normalization, bulk merges on existing decks | ❌ **PROHIBITED** | Catastrophic performance (10+ hours); use Python instead |

### Architecture Insight (Discovered 24 Oct 2025)

Testing revealed that **post-processing merge operations are redundant** because:
- Clone pipeline already creates correct cell merges during generation
- Attempting to re-merge cells that are already merged fails (expected behavior)
- Post-processing should focus on normalization and edge case repairs

**Recommendation**:
- Keep generation merges (working correctly)
- Use Python post-processing for normalization only
- Reserve merge operations for edge case repairs (broken decks from failed generation)

See: `docs/24-10-25/15-merge_architecture_discovery.md` for detailed analysis.

### When to Use COM vs. Python

**Use Python (python-pptx)** for:
- ✅ Any bulk operation (loops over cells/rows/columns)
- ✅ Table normalization and formatting
- ✅ Cell merge operations (post-processing repairs)
- ✅ Text content changes across multiple cells
- ✅ Property changes on multiple shapes/cells

**Use COM (PowerShell/pywin32)** ONLY for:
- ✅ Opening/closing/saving presentations (file I/O)
- ✅ Features not exposed by python-pptx (specific animations, macros)
- ✅ Format conversions (PPTX → PDF)
- ✅ Single, isolated operations (not in loops)

**NEVER Use COM for**:
- ❌ Loops over table cells
- ❌ Bulk property changes
- ❌ Post-processing normalization
- ❌ Any operation that scales with deck size (O(n))

### Performance Targets

| Operation Type | Target Time (88 slides) | Method |
|----------------|------------------------|--------|
| Generation with merges | <5 minutes | Python + generation-time COM merges ✅ |
| Post-processing normalization | <1 minute | Python (python-pptx) ✅ |
| Structural validation | <30 seconds | Python ✅ |
| **Total Pipeline** | **<7 minutes** | **All Python + controlled COM** ✅ |

## Related Documents

- `docs/24-10-25/logs/16-python_migration_summary.md` - Detailed migration analysis
- `amp_automation/presentation/postprocess/` - Python implementation
- `docs/24-10-25/BRAIN_RESET_241025.md` - Project status

## References

- [python-pptx documentation](https://python-pptx.readthedocs.io/)
- [Performance test results](docs/24-10-25/logs/test_python_cli.ps1)
- [COM bottleneck analysis](docs/24-10-25/logs/16-python_migration_summary.md)

## Update: E2E Validation and PowerShell Deprecation (27 Oct 2025)

### Validation Results

**E2E Post-Processing Test (Task 6):**
- ✅ **88 slides processed in <1 second** (not 1 minute as originally projected)
- ✅ **100% success rate** (0 failures, 0 errors)
- ✅ **Performance validated:** 1,800x faster than COM automation
- ✅ **All operations successful:**
  - 20 CARRIED FORWARD rows deleted
  - ~250 merge operations applied (campaign + monthly + summary)
  - ~9,000+ cells normalized to Verdana fonts
  - ~600 cells cleaned (pound sign removal)

**Key Insight:** Python post-processing is **even faster** than originally estimated. Original projection was 1 minute for 88 slides, actual time is <1 second.

### PowerShell Script Deprecation (Task 7)

All COM-based PowerShell scripts in `tools/` directory have been deprecated as of 27 Oct 2025:

**Deprecated Scripts (7 total):**
1. ✅ `RebuildCampaignMerges.ps1` - Campaign merge repairs
2. ✅ `SanitizePrimaryColumns.ps1` - Column sanitization
3. ✅ `FixHorizontalMerges.ps1` - Horizontal merge repairs
4. ✅ `AuditCampaignMerges.ps1` - Campaign merge auditing
5. ✅ `VerifyAllowedHorizontalMerges.ps1` - Merge verification
6. ✅ `ProbeRowHeights.ps1` - Row height inspection
7. ✅ `InspectColumnSpans.ps1` - Column span inspection

**Migration Path:**
- ✅ Use `PostProcessNormalize.ps1` (PowerShell wrapper calling Python CLI)
- ✅ Use `py -m amp_automation.presentation.postprocess.cli` directly
- ❌ Deprecated scripts only for emergency legacy deck repairs

### Updated Performance Targets (Validated 27 Oct 2025)

| Operation Type | Original Projection | Actual (Validated) | Method |
|----------------|--------------------|--------------------|--------|
| Generation with merges | <5 minutes | ~3 minutes | Python-pptx ✅ |
| Post-processing (normalize) | <1 minute | **<1 second** | Python-pptx ✅ |
| Structural validation | <30 seconds | <1 second | Python CLI ✅ |
| **Total Pipeline** | **<7 minutes** | **<4 minutes** | **All Python** ✅ |

**Performance Improvement:** Python post-processing is **60x faster** than originally projected!

### Updated Decision Matrix (27 Oct 2025)

**When to Use Python (python-pptx) - ALWAYS PREFER THIS:**

| Operation | Example | Performance | Status |
|-----------|---------|-------------|--------|
| Bulk table operations | Normalize 88 slides | <1 second | ✅ MANDATORY |
| Cell merging (post-process) | Merge campaign cells | <1 second | ✅ MANDATORY |
| Font normalization | Verdana 6pt/7pt | <1 second | ✅ MANDATORY |
| Row/column operations | Delete CARRIED FORWARD | <1 second | ✅ MANDATORY |
| Text content changes | Update cell values | Milliseconds | ✅ MANDATORY |

**When COM is Acceptable - LIMITED USE ONLY:**

| Operation | Example | Justification | Performance |
|-----------|---------|---------------|-------------|
| File I/O | Open/save presentations | No python-pptx alternative | O(1) ✅ |
| Format conversion | PPTX → PDF export | PowerPoint-specific formats | O(1) ✅ |
| Slide export | Slide → PNG/JPG | Image rendering | O(1) ✅ |
| Advanced features | Animations, macros | Not in python-pptx API | O(1) ✅ |

**When COM is NEVER Acceptable - STRICT PROHIBITION:**

| Anti-Pattern | Why Prohibited | Alternative |
|-------------|----------------|-------------|
| Loops over cells | 1,000+ COM calls = catastrophic | Python-pptx iterates XML directly |
| Bulk property changes | Hours vs seconds | Python-pptx batch operations |
| Post-processing merges | 10+ hours vs <1 second | Python CLI `--operations postprocess-all` |
| Table normalization | Hangs and timeouts | Python CLI `--operations normalize` |

### Architecture Status (27 Oct 2025)

✅ **COM prohibition fully implemented and validated:**
- All bulk operations migrated to Python ✅
- E2E post-processing test passed (100% success) ✅
- All COM-based PowerShell scripts deprecated ✅
- Python CLI provides complete functionality ✅
- Performance validated: 1,800x faster than COM ✅

✅ **No COM usage in bulk operations:**
- Generation: python-pptx only ✅
- Post-processing: Python CLI only ✅
- Diagnostics: Python CLI verbose mode ✅

⚠️ **Remaining COM usage (acceptable):**
- Visual diff tool: slide export to PNG (Task 3)
- File format conversions (if needed)
- No bulk operations - all O(1) file I/O

### Enforcement Update (27 Oct 2025)

**Deprecation Warnings:**
All COM-based scripts now display prominent warnings:
```
⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR BULK OPERATIONS ⚠️

Performance: 1,800x slower than Python
Replacement: PostProcessNormalize.ps1 or Python CLI
Status: DEPRECATED as of 27 Oct 2025
```

**Code Review Requirements:**
- ❌ REJECT any new COM bulk operations
- ❌ REJECT loops over table cells using COM
- ❌ REJECT bulk property changes using COM
- ✅ REQUIRE Python-pptx for all bulk operations
- ✅ REQUIRE explicit justification for any COM usage

**Migration Timeline:**
- 24 Oct 2025: COM prohibition established
- 27 Oct 2025: Python CLI validated, PowerShell scripts deprecated
- Future: Complete removal of deprecated COM scripts (after grace period)

### Related Documents (Updated 27 Oct 2025)

- `docs/27-10-25/artifacts/task6_postprocessing_e2e_complete.md` - E2E validation results
- `docs/27-10-25/artifacts/task7_powershell_deprecation_complete.md` - Script deprecation
- `tools/PostProcessNormalize.ps1` - Recommended PowerShell wrapper
- `amp_automation/presentation/postprocess/cli.py` - Python CLI implementation

---

**Last Updated:** 27 October 2025 (E2E validation and PowerShell deprecation)
**Previous Update:** 24 October 2025 (Initial COM prohibition)
**Review Date:** Annually or when adding new COM code
**Owner:** Tech Lead / Architecture Team
