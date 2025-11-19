"""
Post-processing operations for PowerPoint presentations.

This package provides Python-based bulk operations for table manipulation,
replacing slow COM-based operations from PowerShell scripts.

Architecture:
- COM (via pywin32/comtypes): Reserved for non-bulk operations, exports, and
  specific features not exposed by file-level libraries
- Python libraries (python-pptx): Handle bulk operations like merges, column
  sanitization, and table manipulation

Definitive Post-Processing Workflow:
  1. unmerge_all_cells         - Clean slate: remove all rogue merges
  2. delete_carried_forward_rows - Remove invalid CARRIED FORWARD rows
  3. merge_campaign_cells      - Merge campaign names vertically (column A)
  4. merge_media_cells         - Merge media channels vertically (column B: TELEVISION, DIGITAL, OOH, etc.)
  5. merge_monthly_total_cells - Merge MONTHLY TOTAL horizontally (gray cells, cols 1-3)
  6. merge_summary_cells       - Merge GRAND TOTAL horizontally (cols 1-3)
  7. fix_grand_total_wrapping  - Ensure single-line display in GRAND TOTAL
  8. remove_pound_signs_from_totals - Remove £ symbols from total rows
  9. normalize_table_fonts     - Enforce Verdana 6pt body, 7pt header

Performance: ~40 seconds for 88-slide deck (vs 10+ hours COM automation)
Validation: 684 operations, 0 failures, 100% success rate

Use CLI: python -m amp_automation.presentation.postprocess.cli --operations postprocess-all
"""

from .table_normalizer import (
    normalize_table_layout,
    apply_blank_cell_formatting,
    normalize_table_fonts,
    delete_carried_forward_rows,
    fix_grand_total_wrapping,
    remove_pound_signs_from_totals,
)
from .cell_merges import (
    merge_campaign_cells,
    merge_media_cells,
    merge_percentage_cells,
    merge_monthly_total_cells,
    merge_summary_cells,
)
from .span_operations import (
    reset_primary_column_spans,
    reset_column_group,
)
from .unmerge_operations import (
    unmerge_all_cells,
    unmerge_column,
    unmerge_row,
    unmerge_primary_columns,
)

__all__ = [
    # Table normalization
    "normalize_table_layout",
    "apply_blank_cell_formatting",
    "normalize_table_fonts",
    "delete_carried_forward_rows",
    "fix_grand_total_wrapping",
    "remove_pound_signs_from_totals",
    # Cell merges
    "merge_campaign_cells",
    "merge_media_cells",
    "merge_percentage_cells",
    "merge_monthly_total_cells",
    "merge_summary_cells",
    # Span operations
    "reset_primary_column_spans",
    "reset_column_group",
    # Unmerge operations
    "unmerge_all_cells",
    "unmerge_column",
    "unmerge_row",
    "unmerge_primary_columns",
]
