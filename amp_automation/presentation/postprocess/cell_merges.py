"""
Cell merge operations for campaign, monthly, and summary rows.

This module provides Python-based implementations of cell merge operations,
replacing slow COM-based PowerShell operations.
"""

import logging

logger = logging.getLogger(__name__)


def merge_campaign_cells(table):
    """
    Merge campaign cells vertically in column 1.

    Equivalent to PowerShell function: Campaign merge operations

    This function identifies campaign rows (rows between MONTHLY TOTAL rows)
    and merges the cells in column 1 vertically to create a single cell
    spanning multiple rows for each campaign.

    Args:
        table: python-pptx table object
    """
    logger.debug("Merging campaign cells")

    # Note: python-pptx provides table.cell(row, col).merge(other_cell) for merging,
    # but we need to:
    # 1. Identify campaign start/end rows (between MONTHLY TOTAL rows)
    # 2. Merge cells vertically in column 1
    # 3. Apply styling (center alignment, bold, specific font size)
    #
    # This requires reading cell text to identify MONTHLY TOTAL rows.

    logger.warning("merge_campaign_cells: Not yet fully implemented in python-pptx")
    logger.warning("Implementation requires cell text analysis and vertical merging")


def merge_monthly_total_cells(table):
    """
    Merge monthly total cells horizontally (columns 1-3).

    Equivalent to PowerShell function: Monthly total merge operations

    This function identifies MONTHLY TOTAL rows and merges cells horizontally
    across columns 1-3 to create a single cell for the "MONTHLY TOTAL" label.

    Args:
        table: python-pptx table object
    """
    logger.debug("Merging monthly total cells")

    # Note: python-pptx provides table.cell(row, col).merge(other_cell) for merging,
    # but we need to:
    # 1. Identify MONTHLY TOTAL rows (by reading cell text in column 1)
    # 2. Merge cells horizontally across columns 1-3
    # 3. Apply styling (center alignment, bold, specific font size)

    logger.warning("merge_monthly_total_cells: Not yet fully implemented in python-pptx")
    logger.warning("Implementation requires cell text analysis and horizontal merging")


def merge_summary_cells(table):
    """
    Merge summary cells horizontally (columns 1-3).

    Equivalent to PowerShell function: Summary merge operations

    This function identifies GRAND TOTAL and CARRIED FORWARD rows and merges
    cells horizontally across columns 1-3 to create a single cell for the label.

    Args:
        table: python-pptx table object
    """
    logger.debug("Merging summary cells (GRAND TOTAL, CARRIED FORWARD)")

    # Note: python-pptx provides table.cell(row, col).merge(other_cell) for merging,
    # but we need to:
    # 1. Identify GRAND TOTAL and CARRIED FORWARD rows (by reading cell text)
    # 2. Merge cells horizontally across columns 1-3
    # 3. Apply styling (center alignment, bold, specific font size)

    logger.warning("merge_summary_cells: Not yet fully implemented in python-pptx")
    logger.warning("Implementation requires cell text analysis and horizontal merging")


# Helper function for future implementation
def normalize_label(text: str) -> str:
    """
    Normalize label text for comparison.

    Args:
        text: Cell text content

    Returns:
        Normalized text (uppercase, stripped whitespace)
    """
    if not text:
        return ""
    return text.strip().upper()


def is_monthly_total(cell_text: str) -> bool:
    """
    Check if cell text represents a MONTHLY TOTAL row.

    Args:
        cell_text: Text content from cell in column 1

    Returns:
        True if this is a MONTHLY TOTAL row
    """
    normalized = normalize_label(cell_text)
    return "MONTHLY" in normalized and "TOTAL" in normalized


def is_grand_total(cell_text: str) -> bool:
    """
    Check if cell text represents a GRAND TOTAL row.

    Args:
        cell_text: Text content from cell in column 1

    Returns:
        True if this is a GRAND TOTAL row
    """
    normalized = normalize_label(cell_text)
    return "GRAND" in normalized and "TOTAL" in normalized


def is_carried_forward(cell_text: str) -> bool:
    """
    Check if cell text represents a CARRIED FORWARD row.

    Args:
        cell_text: Text content from cell in column 1

    Returns:
        True if this is a CARRIED FORWARD row
    """
    normalized = normalize_label(cell_text)
    return "CARRIED" in normalized and "FORWARD" in normalized
