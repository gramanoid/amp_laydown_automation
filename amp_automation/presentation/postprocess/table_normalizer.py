"""
Table normalization and cell formatting operations.

This module provides Python-based implementations of table layout normalization
and blank cell formatting, replacing slow COM-based PowerShell operations.
"""

import logging
from typing import Optional

from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.util import Pt

logger = logging.getLogger(__name__)

# Constants from PowerShell script
ZERO_WIDTH_SPACE = "\u200B"
BLANK_FONT_NAME = "Calibri"
BLANK_FONT_SIZE = Pt(8)


def set_cell_fixed_layout(cell):
    """
    Set fixed layout properties for a table cell.

    Equivalent to PowerShell function: Set-CellFixedLayout

    Sets:
    - AutoSize = none (no auto-sizing)
    - WordWrap = enabled
    - All margins = 0
    - VerticalAnchor = middle

    Args:
        cell: python-pptx table cell object
    """
    try:
        text_frame = cell.text_frame

        # Set auto-size and word wrap
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        text_frame.word_wrap = True

        # Set margins to 0
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0

        # Set vertical anchor to middle
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    except Exception as e:
        logger.debug(f"Error setting fixed layout for cell: {e}")


def normalize_cell_content(text: str) -> str:
    """
    Normalize cell content by stripping whitespace.

    Args:
        text: Cell text content

    Returns:
        Normalized text with trailing whitespace removed
    """
    if not text:
        return ""
    return text.rstrip()


def ensure_blank_cell_formatting(cell):
    """
    Apply formatting to blank cells and cells containing only "-".

    Equivalent to PowerShell function: Ensure-BlankCellFormatting

    Args:
        cell: python-pptx table cell object
    """
    try:
        # First apply fixed layout
        set_cell_fixed_layout(cell)

        text_frame = cell.text_frame

        # Get normalized content
        content = normalize_cell_content(text_frame.text)
        is_dash = content == "-"
        is_blank = len(content) == 0

        # Only process blank cells or cells with "-"
        if not is_dash and not is_blank:
            return

        # Replace blank cells with zero-width space
        if is_blank:
            text_frame.text = ZERO_WIDTH_SPACE

        # Apply formatting
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

        # Format the paragraph
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER

            # Format runs (text segments)
            for run in paragraph.runs:
                run.font.name = BLANK_FONT_NAME
                run.font.size = BLANK_FONT_SIZE
                run.font.bold = False

    except Exception as e:
        logger.debug(f"Error formatting blank cell: {e}")


def normalize_table_layout(table):
    """
    Normalize layout for all cells in a table.

    Equivalent to PowerShell function: Normalize-TableLayout

    This function iterates through all cells and applies fixed layout
    properties to ensure consistent formatting across the table.

    Args:
        table: python-pptx table object
    """
    logger.debug(f"Normalizing table layout: {len(table.rows)} rows × {len(table.columns)} columns")

    cell_count = 0
    error_count = 0

    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            try:
                set_cell_fixed_layout(cell)
                cell_count += 1
            except Exception as e:
                error_count += 1
                logger.debug(f"Error normalizing cell ({row_idx},{col_idx}): {e}")

    logger.debug(f"Normalized {cell_count} cells ({error_count} errors)")


def apply_blank_cell_formatting(table):
    """
    Apply blank cell formatting to all cells in a table.

    Equivalent to PowerShell function: Apply-BlankCellFormatting

    This function iterates through all cells and applies special formatting
    to blank cells and cells containing only "-".

    Args:
        table: python-pptx table object
    """
    logger.debug(f"Applying blank cell formatting: {len(table.rows)} rows × {len(table.columns)} columns")

    cell_count = 0
    blank_count = 0
    error_count = 0

    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            try:
                content = normalize_cell_content(cell.text_frame.text)
                is_blank = len(content) == 0 or content == "-"

                if is_blank:
                    blank_count += 1

                ensure_blank_cell_formatting(cell)
                cell_count += 1

            except Exception as e:
                error_count += 1
                logger.debug(f"Error formatting cell ({row_idx},{col_idx}): {e}")

    logger.debug(f"Formatted {cell_count} cells ({blank_count} blank, {error_count} errors)")
