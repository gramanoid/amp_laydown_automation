"""
Post-processing operations for PowerPoint presentations.

This package provides Python-based bulk operations for table manipulation,
replacing slow COM-based operations from PowerShell scripts.

Architecture:
- COM (via pywin32/comtypes): Reserved for non-bulk operations, exports, and
  specific features not exposed by file-level libraries
- Python libraries (python-pptx): Handle bulk operations like merges, column
  sanitization, and table manipulation
"""

from .table_normalizer import normalize_table_layout, apply_blank_cell_formatting
from .cell_merges import (
    merge_campaign_cells,
    merge_monthly_total_cells,
    merge_summary_cells,
)
from .span_operations import (
    reset_primary_column_spans,
    reset_column_group,
)

__all__ = [
    # Table normalization
    "normalize_table_layout",
    "apply_blank_cell_formatting",
    # Cell merges
    "merge_campaign_cells",
    "merge_monthly_total_cells",
    "merge_summary_cells",
    # Span operations
    "reset_primary_column_spans",
    "reset_column_group",
]
