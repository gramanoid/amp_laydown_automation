"""
Unmerge operations for cleaning up merged cells.

This module provides Python-based implementations to unmerge cells,
allowing a clean slate before applying intentional merges.
"""

import logging
from typing import Set, Tuple

logger = logging.getLogger(__name__)


def unmerge_all_cells(table) -> int:
    """
    Unmerge ALL cells in a table by removing merge attributes.

    This provides a "clean slate" by resetting all cells to individual cells,
    removing any merged cell configurations from the table.

    Args:
        table: python-pptx table object

    Returns:
        int: Number of cells that were unmerged
    """
    logger.debug(f"Unmerging all cells in table: {len(table.rows)} rows × {len(table.columns)} cols")

    unmerged_count = 0

    # First pass: Remove ALL merge attributes from ALL cells
    for row_idx in range(len(table.rows)):
        for col_idx in range(len(table.columns)):
            cell = table.cell(row_idx, col_idx)
            tc = cell._tc

            # Check if cell has any merge attributes
            row_span = tc.get('rowSpan')
            grid_span = tc.get('gridSpan')
            v_merge = tc.get('vMerge')
            h_merge = tc.get('hMerge')

            has_merge = row_span or grid_span or v_merge or h_merge

            if has_merge:
                # Remove all merge attributes
                if 'rowSpan' in tc.attrib:
                    del tc.attrib['rowSpan']
                if 'gridSpan' in tc.attrib:
                    del tc.attrib['gridSpan']
                if 'vMerge' in tc.attrib:
                    del tc.attrib['vMerge']
                if 'hMerge' in tc.attrib:
                    del tc.attrib['hMerge']

                unmerged_count += 1

    logger.info(f"Unmerged {unmerged_count} cells (removed all merge attributes)")
    return unmerged_count


def unmerge_column(table, col_idx: int) -> int:
    """
    Unmerge all cells in a specific column.

    Args:
        table: python-pptx table object
        col_idx: Column index (0-based)

    Returns:
        int: Number of cells unmerged in this column
    """
    logger.debug(f"Unmerging column {col_idx}")

    unmerged_count = 0

    for row_idx in range(len(table.rows)):
        cell = table.cell(row_idx, col_idx)
        tc = cell._tc

        # Remove vertical merge attributes
        if tc.get('rowSpan'):
            tc.attrib.pop('rowSpan', None)
            unmerged_count += 1
        if tc.get('vMerge'):
            tc.attrib.pop('vMerge', None)
            unmerged_count += 1

    logger.debug(f"Unmerged {unmerged_count} cells in column {col_idx}")
    return unmerged_count


def unmerge_row(table, row_idx: int) -> int:
    """
    Unmerge all cells in a specific row.

    Args:
        table: python-pptx table object
        row_idx: Row index (0-based)

    Returns:
        int: Number of cells unmerged in this row
    """
    logger.debug(f"Unmerging row {row_idx}")

    unmerged_count = 0

    for col_idx in range(len(table.columns)):
        cell = table.cell(row_idx, col_idx)
        tc = cell._tc

        # Remove horizontal merge attributes
        if tc.get('gridSpan'):
            tc.attrib.pop('gridSpan', None)
            unmerged_count += 1
        if tc.get('hMerge'):
            tc.attrib.pop('hMerge', None)
            unmerged_count += 1

    logger.debug(f"Unmerged {unmerged_count} cells in row {row_idx}")
    return unmerged_count


def unmerge_primary_columns(table, max_cols: int = 3) -> int:
    """
    Unmerge the first N columns (typically 1-3 for campaign/media/metrics).

    Args:
        table: python-pptx table object
        max_cols: Number of columns to unmerge (default: 3)

    Returns:
        int: Total number of cells unmerged
    """
    logger.debug(f"Unmerging primary columns (1-{max_cols})")

    total_unmerged = 0

    for col_idx in range(min(max_cols, len(table.columns))):
        total_unmerged += unmerge_column(table, col_idx)

    logger.info(f"Unmerged {total_unmerged} cells in primary columns")
    return total_unmerged
