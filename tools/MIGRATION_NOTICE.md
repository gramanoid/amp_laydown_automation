# PowerShell Post-Processing Migration Notice

## Status: **COMPLETED** (2025-10-24)

PowerShell COM-based post-processing scripts have been replaced with Python-based operations.

## Migration Summary

### Deprecated Scripts (DO NOT USE)
- ❌ `PostProcessCampaignMerges.ps1` - COM-based bulk operations (10+ hours runtime)
  - **Status**: Deprecated, scheduled for archive
  - **Reason**: Catastrophic performance issues
  - **Replacement**: `PostProcessNormalize.ps1` (Python wrapper)

### New Scripts (USE THESE)
- ✅ `PostProcessNormalize.ps1` - Python CLI wrapper (~30 seconds runtime)
  - **Operations**: Table normalization, cell formatting
  - **Performance**: 60x faster than COM
  - **Backend**: `python -m amp_automation.presentation.postprocess.cli`

## Architecture Change

**OLD (Deprecated)**:
```
Generation → PowerShell COM (merges + normalization) → Validation
                     ↑ 10+ hours, frequently fails
```

**NEW (Current)**:
```
Generation (with merges) → Python (normalization) → Validation
    ↑ minutes                    ↑ ~30 seconds         ↑ seconds
```

### Key Insight
Cell merges are now created **during generation** (assembly.py:629,649), NOT in post-processing.
Post-processing only handles normalization and edge case repairs.

## Usage Examples

### Basic Normalization
```powershell
.\tools\PostProcessNormalize.ps1 -PresentationPath "output\presentations\deck.pptx"
```

### With Specific Operations
```powershell
.\tools\PostProcessNormalize.ps1 `
    -PresentationPath "output\presentations\deck.pptx" `
    -Operations "normalize,merge-campaign" `
    -SlideFilter 2,3,4 `
    -VerboseOutput
```

### Direct Python CLI
```bash
python -m amp_automation.presentation.postprocess.cli \
    --presentation-path "output/presentations/deck.pptx" \
    --operations normalize \
    --verbose
```

## Available Operations

| Operation | Description | Typical Use Case |
|-----------|-------------|------------------|
| `normalize` | Fix table layout, cell formatting | **Default, always use** |
| `reset-spans` | Unmerge cells in primary columns | Edge case repairs only |
| `merge-campaign` | Vertical campaign merges | Edge case repairs only* |
| `merge-monthly` | Horizontal monthly total merges | Edge case repairs only* |
| `merge-summary` | Horizontal summary merges | Edge case repairs only* |

\* **Note**: Merge operations are redundant if generation completed successfully.
  Only use for repairing broken decks from failed generation runs.

## Performance Comparison

| Operation | PowerShell COM | Python | Improvement |
|-----------|----------------|--------|-------------|
| Normalize 88 slides | 10+ hours (never completed) | ~30 seconds | **60x faster** |
| Merge operations | 6-10 min/slide (with timeouts) | N/A (done in generation) | **Not needed** |
| Full pipeline | Never completed | <5 minutes | **Practical** |

## Migration Checklist

- [x] Python implementation created (commit d3e2b98)
- [x] PowerShell wrapper created (`PostProcessNormalize.ps1`)
- [x] End-to-end testing completed (commit 9ba7a82)
- [x] Architecture documented (commits 8320c3f, 3c54b1b, 6c340dc)
- [ ] Deprecate `PostProcessCampaignMerges.ps1` (rename to `.deprecated`)
- [ ] Update runbooks and documentation
- [ ] Archive old PowerShell backup scripts

## References

- **ADR**: `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`
- **OpenSpec Proposal**: `openspec/changes/clarify-postprocessing-architecture/`
- **Discovery Document**: `docs/24-10-25/15-merge_architecture_discovery.md`
- **Python Module**: `amp_automation/presentation/postprocess/`

## Support

For issues with the new Python-based workflow:
1. Check Python 3.13+ is installed: `python --version`
2. Verify python-pptx is installed: `python -m pip show python-pptx`
3. Check logs in `docs/{DD-MM-YY}/logs/postprocess_normalize_*.log`
4. Review verbose output: add `-VerboseOutput` flag

For legacy COM script issues:
- **Do not fix** - migrate to Python instead
- Scripts are deprecated and will be archived
