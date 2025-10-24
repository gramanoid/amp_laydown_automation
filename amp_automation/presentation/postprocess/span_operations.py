"""
Column and row span reset operations.

This module provides Python-based implementations of span reset operations,
replacing slow COM-based PowerShell operations.
"""

import logging

logger = logging.getLogger(__name__)


def reset_primary_column_spans(table, max_cols: int = 3):
    """
    Reset column spans in primary columns (columns 1-3).

    Equivalent to PowerShell function: Reset-PrimaryColumnSpans

    This function unmerges any horizontally merged cells in the first few columns
    (typically columns 1-3) to prepare for campaign/monthly merges.

    Args:
        table: python-pptx table object
        max_cols: Maximum number of columns to process (default: 3)
    """
    logger.debug(f"Resetting column spans for columns 1-{max_cols}")

    # Note: python-pptx doesn't provide direct access to merged cell information
    # or split operations like PowerShell COM does.
    #
    # For now, this is a placeholder. The actual implementation would need to:
    # 1. Detect merged cells (cells where multiple grid positions share the same cell object)
    # 2. Split them back into individual cells
    #
    # This may require using the underlying OOXML or falling back to COM for these operations.
    #
    # Alternative approach: Since we control the deck generation, we could ensure
    # that primary columns are never merged in the first place.

    logger.warning("reset_primary_column_spans: Not yet fully implemented in python-pptx")
    logger.warning("Consider pre-sanitizing columns using PowerShell COM or ensuring clean generation")


def reset_column_group(table, max_cols: int = 3):
    """
    Reset column group formatting.

    Equivalent to PowerShell function: Reset-ColumnGroup

    This function ensures that the first few columns (typically 1-3) are properly
    formatted and not merged.

    Args:
        table: python-pptx table object
        max_cols: Maximum number of columns to process (default: 3)
    """
    logger.debug(f"Resetting column group for columns 1-{max_cols}")

    # Similar limitation as reset_primary_column_spans
    # This function in PowerShell calls Ensure-FirstColumns which verifies that
    # columns 1-max_cols exist and are separate cells.

    logger.warning("reset_column_group: Not yet fully implemented in python-pptx")
    logger.warning("Consider pre-sanitizing columns using PowerShell COM or ensuring clean generation")
