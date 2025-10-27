"""
Cell merge operations for campaign, monthly, and summary rows.

This module provides Python-based implementations of cell merge operations,
replacing slow COM-based PowerShell operations.
"""

import logging
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Pt
from pptx.dml.color import RGBColor

logger = logging.getLogger(__name__)

# Font configuration (matching template requirements)
FONT_NAME = "Verdana"
CAMPAIGN_FONT_SIZE = 5  # Campaign text (smaller to prevent mid-word breaks in narrow cells)
MEDIA_FONT_SIZE = 6  # Media column text
MONTHLY_TOTAL_FONT_SIZE = 6  # Body text
SUMMARY_FONT_SIZE = 7  # Bottom row


def merge_campaign_cells(table):
    """
    Merge campaign cells vertically in column 1.

    Equivalent to PowerShell function: Campaign merge operations

    This function identifies campaign rows (rows between MONTHLY TOTAL rows)
    and merges the cells in column 1 vertically to create a single cell
    spanning multiple rows for each campaign.

    Args:
        table: python-pptx table object

    Returns:
        int: Number of campaign merges performed
    """
    logger.debug("Merging campaign cells")

    row_count = len(table.rows)
    campaign_start = None
    campaign_name = None  # Store campaign name when we find it
    merges_performed = 0

    # Iterate through rows (skip header row index 0)
    for row_idx in range(1, row_count):
        cell = table.cell(row_idx, 0)  # Column 0 (CAMPAIGN column)
        cell_text = _get_cell_text(cell)
        is_gray = _has_gray_background(cell)

        # CRITICAL FIX: Some WHITE cells have "MONTHLY TOTAL\nCAMPAIGN_NAME" format
        # Extract campaign name from white cells only (preserve gray MONTHLY TOTAL cells)
        if not is_gray:
            actual_campaign_name = _extract_campaign_name(cell_text)
        else:
            # Gray cell - don't extract, use first line for detection
            actual_campaign_name = normalize_label(cell_text)

        # Check if this is a special row using the EXTRACTED campaign name (not raw text)
        # This prevents white cells with "MONTHLY TOTAL\nCAMPAIGN" from being misidentified
        is_monthly = is_monthly_total(actual_campaign_name) if actual_campaign_name else False
        is_grand = is_grand_total(actual_campaign_name) if actual_campaign_name else False

        # Track campaign start (non-empty, non-special rows, non-gray)
        # Campaign cells are WHITE cells with campaign names
        if actual_campaign_name and not is_monthly and not is_grand and not is_gray:
            if campaign_start is None:
                campaign_start = row_idx
                campaign_name = actual_campaign_name  # Save the campaign name immediately
                logger.debug(f"Found campaign start at row {row_idx}: {campaign_name}")

        # Perform merge when we hit a GRAY MONTHLY TOTAL row
        # (White cells with "MONTHLY TOTAL\nCAMPAIGN" are campaign cells, not triggers)
        if is_monthly and is_gray and campaign_start is not None:
            campaign_end = row_idx - 1

            if campaign_end > campaign_start:
                try:
                    # Merge cells vertically in column 0 (CAMPAIGN column)
                    top_cell = table.cell(campaign_start, 0)
                    bottom_cell = table.cell(campaign_end, 0)

                    # Check if already merged
                    if not _cells_are_same(top_cell, bottom_cell):
                        top_cell.merge(bottom_cell)
                        merges_performed += 1
                        logger.debug(f"Merged campaign rows {campaign_start}-{campaign_end}: {campaign_name}")

                    # Apply styling to merged cell and set cleaned campaign name
                    merged_cell = table.cell(campaign_start, 0)
                    _apply_cell_styling(
                        merged_cell,
                        text=campaign_name,  # Set the cleaned campaign name (removes "MONTHLY TOTAL" prefix)
                        font_size=CAMPAIGN_FONT_SIZE,
                        bold=True,
                        center_align=True,
                        vertical_center=True
                    )

                except Exception as e:
                    logger.error(f"Failed to merge campaign rows {campaign_start}-{campaign_end}: {e}")

            campaign_start = None
            campaign_name = None  # Reset campaign name too

    logger.info(f"Campaign merges completed: {merges_performed} merge(s)")
    return merges_performed


def merge_media_cells(table):
    """
    Merge media channel cells vertically in column 1 (MEDIA column).

    This function identifies media channels (TELEVISION, DIGITAL, OOH, OTHER)
    and merges the cells vertically to span all metric rows for that media channel.

    Args:
        table: python-pptx table object

    Returns:
        int: Number of media channel merges performed
    """
    logger.debug("Merging media channel cells")

    row_count = len(table.rows)
    merges_performed = 0

    # Known media channels to look for
    MEDIA_CHANNELS = ["TELEVISION", "DIGITAL", "OOH", "OTHER", "RADIO", "PRINT", "CINEMA"]

    media_start = None
    media_name = None

    # Iterate through rows (skip header row index 0)
    for row_idx in range(1, row_count):
        cell = table.cell(row_idx, 1)  # Column 1 (0-indexed) is MEDIA column
        cell_text = _get_cell_text(cell).strip().upper()

        # Check if this is a media channel name
        is_media_channel = any(channel in cell_text for channel in MEDIA_CHANNELS)

        # Check if this row is empty or has dash (metric row)
        is_empty_or_dash = cell_text in ["", "-"]

        if is_media_channel:
            # If we were tracking a previous media channel, merge it now
            if media_start is not None and media_start < row_idx - 1:
                media_end = row_idx - 1
                try:
                    top_cell = table.cell(media_start, 1)
                    bottom_cell = table.cell(media_end, 1)

                    # Check if already merged
                    if not _cells_are_same(top_cell, bottom_cell):
                        top_cell.merge(bottom_cell)
                        merges_performed += 1
                        logger.debug(f"Merged {media_name} rows {media_start}-{media_end}")

                    # Apply styling to merged cell
                    merged_cell = table.cell(media_start, 1)
                    _apply_cell_styling(
                        merged_cell,
                        text=media_name,
                        font_size=MEDIA_FONT_SIZE,
                        bold=True,
                        center_align=True,
                        vertical_center=True
                    )

                except Exception as e:
                    logger.error(f"Failed to merge {media_name} rows {media_start}-{media_end}: {e}")

            # Start tracking this new media channel
            media_start = row_idx
            media_name = cell_text
            logger.debug(f"Found media channel at row {row_idx}: {media_name}")

        # Check if this row ends the current media channel (MONTHLY TOTAL or campaign name)
        campaign_cell = table.cell(row_idx, 0)  # Column 0 (CAMPAIGN)
        campaign_text = _get_cell_text(campaign_cell).strip().upper()
        is_special_row = (
            "MONTHLY" in campaign_text and "TOTAL" in campaign_text
        ) or (
            "GRAND" in campaign_text or "BRAND" in campaign_text
        )

        # If we hit a special row and were tracking a media channel, merge it
        if is_special_row and media_start is not None:
            media_end = row_idx - 1

            if media_end >= media_start:
                try:
                    top_cell = table.cell(media_start, 1)
                    bottom_cell = table.cell(media_end, 1)

                    # Check if already merged
                    if not _cells_are_same(top_cell, bottom_cell):
                        top_cell.merge(bottom_cell)
                        merges_performed += 1
                        logger.debug(f"Merged {media_name} rows {media_start}-{media_end}")

                    # Apply styling to merged cell
                    merged_cell = table.cell(media_start, 1)
                    _apply_cell_styling(
                        merged_cell,
                        text=media_name,
                        font_size=MEDIA_FONT_SIZE,
                        bold=True,
                        center_align=True,
                        vertical_center=True
                    )

                except Exception as e:
                    logger.error(f"Failed to merge {media_name} rows {media_start}-{media_end}: {e}")

            media_start = None
            media_name = None

    # Handle case where table ends while still tracking a media channel
    if media_start is not None and media_start < row_count - 1:
        media_end = row_count - 1
        try:
            top_cell = table.cell(media_start, 1)
            bottom_cell = table.cell(media_end, 1)

            if not _cells_are_same(top_cell, bottom_cell):
                top_cell.merge(bottom_cell)
                merges_performed += 1
                logger.debug(f"Merged {media_name} rows {media_start}-{media_end}")

            merged_cell = table.cell(media_start, 1)
            _apply_cell_styling(
                merged_cell,
                text=media_name,
                font_size=MEDIA_FONT_SIZE,
                bold=True,
                center_align=True,
                vertical_center=True
            )

        except Exception as e:
            logger.error(f"Failed to merge {media_name} rows {media_start}-{media_end}: {e}")

    logger.info(f"Media channel merges completed: {merges_performed} merge(s)")
    return merges_performed


def merge_monthly_total_cells(table):
    """
    Merge monthly total cells horizontally (columns 1-3).

    Equivalent to PowerShell function: Monthly total merge operations

    This function identifies MONTHLY TOTAL rows and merges cells horizontally
    across columns 1-3 to create a single cell for the "MONTHLY TOTAL" label.

    IMPORTANT: Only merges cells with gray background. White background cells
    with "MONTHLY TOTAL" text are campaign names and should NOT be merged.

    Args:
        table: python-pptx table object

    Returns:
        int: Number of monthly total merges performed
    """
    logger.debug("Merging monthly total cells")

    row_count = len(table.rows)
    col_count = len(table.columns)
    merges_performed = 0

    # Iterate through rows (skip header row index 0)
    for row_idx in range(1, row_count):
        cell = table.cell(row_idx, 0)  # Column 1 (0-indexed)
        cell_text = _get_cell_text(cell)

        # Check both text content AND background color
        if is_monthly_total(cell_text) and _has_gray_background(cell):
            normalized = normalize_label(cell_text)

            try:
                # Merge horizontally across columns 1-3 (0-2 in 0-indexed)
                # Merge from first column to last column in ONE operation
                left_cell = table.cell(row_idx, 0)
                right_cell = table.cell(row_idx, 2)  # Column 3 (0-indexed as 2)

                # Check if already merged
                if not _cells_are_same(left_cell, right_cell):
                    left_cell.merge(right_cell)
                    merges_performed += 1
                    logger.debug(f"Merged MONTHLY TOTAL row {row_idx} across columns 1-3")

                # Apply styling to merged cell
                merged_cell = table.cell(row_idx, 0)
                _apply_cell_styling(
                    merged_cell,
                    text=normalized,
                    font_size=MONTHLY_TOTAL_FONT_SIZE,
                    bold=True,
                    center_align=True,
                    vertical_center=True
                )

            except Exception as e:
                logger.error(f"Failed to merge MONTHLY TOTAL row {row_idx}: {e}")

    logger.info(f"Monthly total merges completed: {merges_performed} merge(s)")
    return merges_performed


def merge_summary_cells(table):
    """
    Merge summary cells horizontally (columns 1-3).

    Equivalent to PowerShell function: Summary merge operations

    This function identifies GRAND TOTAL and CARRIED FORWARD rows and merges
    cells horizontally across columns 1-3 to create a single cell for the label.

    Args:
        table: python-pptx table object

    Returns:
        int: Number of summary merges performed
    """
    logger.debug("Merging summary cells (GRAND TOTAL, CARRIED FORWARD)")

    row_count = len(table.rows)
    col_count = len(table.columns)
    merges_performed = 0

    # Iterate through rows (skip header row index 0)
    for row_idx in range(1, row_count):
        cell = table.cell(row_idx, 0)  # Column 1 (0-indexed)
        cell_text = _get_cell_text(cell)

        # Check both text content AND background color
        if is_grand_total(cell_text) and _has_gray_background(cell):
            normalized = normalize_label(cell_text)

            try:
                # Merge horizontally across columns 1-3 (0-2 in 0-indexed)
                # Merge from first column to last column in ONE operation
                left_cell = table.cell(row_idx, 0)
                right_cell = table.cell(row_idx, 2)  # Column 3 (0-indexed as 2)

                # Check if already merged
                if not _cells_are_same(left_cell, right_cell):
                    left_cell.merge(right_cell)
                    merges_performed += 1
                    logger.debug(f"Merged summary row {row_idx} ({normalized}) across columns 1-3")

                # Apply styling to merged cell
                merged_cell = table.cell(row_idx, 0)
                _apply_cell_styling(
                    merged_cell,
                    text=normalized,
                    font_size=SUMMARY_FONT_SIZE,
                    bold=True,
                    center_align=True,
                    vertical_center=True
                )

                # Apply green background if this is BRAND TOTAL
                if "BRAND" in normalized:
                    try:
                        fill = merged_cell.fill
                        fill.solid()
                        fill.fore_color.rgb = RGBColor(0x30, 0xea, 0x03)  # Green #30ea03
                        logger.debug(f"Applied green background to BRAND TOTAL row {row_idx}")
                    except Exception as bg_error:
                        logger.warning(f"Failed to apply green background to BRAND TOTAL: {bg_error}")

            except Exception as e:
                logger.error(f"Failed to merge summary row {row_idx} ({normalized}): {e}")

    logger.info(f"Summary merges completed: {merges_performed} merge(s)")
    return merges_performed


# Helper function for future implementation
def normalize_label(text: str) -> str:
    """
    Normalize label text for comparison.

    IMPORTANT: Cells may contain multiple lines (e.g., "MONTHLY TOTAL\nTELEVISION").
    We only want the first line for labels.

    Args:
        text: Cell text content

    Returns:
        Normalized text (uppercase, stripped whitespace, first line only)
    """
    if not text:
        return ""
    # Take only the first line (before any newline character)
    first_line = text.split('\n')[0].split('\r')[0]
    return first_line.strip().upper()


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
    Check if cell text represents a GRAND TOTAL or BRAND TOTAL row.

    Args:
        cell_text: Text content from cell in column 1

    Returns:
        True if this is a GRAND TOTAL or BRAND TOTAL row
    """
    normalized = normalize_label(cell_text)
    return ("GRAND" in normalized or "BRAND" in normalized) and "TOTAL" in normalized


# Private helper functions
def _extract_campaign_name(cell_text: str) -> str:
    """
    Extract campaign name from cell text.

    Some cells have format: "MONTHLY TOTAL ( 000)\nCAMPAIGN_NAME"
    We need to extract just the campaign name (the line after MONTHLY TOTAL).

    Args:
        cell_text: Raw cell text content

    Returns:
        Campaign name (normalized, uppercase), or empty string
    """
    if not cell_text:
        return ""

    # Split into lines
    lines = [line.strip() for line in cell_text.split('\n') if line.strip()]

    if not lines:
        return ""

    # If first line is "MONTHLY TOTAL", use second line as campaign name
    first_line_normalized = lines[0].upper()
    if "MONTHLY" in first_line_normalized and "TOTAL" in first_line_normalized:
        if len(lines) > 1:
            # Return second line (the actual campaign name)
            return lines[1].upper()
        else:
            # Just "MONTHLY TOTAL" with no campaign name - not a campaign row
            return ""

    # Otherwise, first line is the campaign name
    return lines[0].upper()


def _has_gray_background(cell) -> bool:
    """
    Check if a cell has a gray background fill.

    Args:
        cell: python-pptx cell object

    Returns:
        True if cell has gray background
    """
    try:
        fill = cell.fill

        # Check if solid fill
        if fill.type == 1:  # MSO_FILL_TYPE.SOLID
            # Get RGB color
            color = fill.fore_color
            if color.type == 1:  # MSO_COLOR_TYPE.RGB
                rgb = color.rgb
                # Gray colors have R=G=B and are not pure white (255,255,255)
                # Typical gray values in templates: around (191,191,191) to (217,217,217)
                r, g, b = rgb[0], rgb[1], rgb[2]

                # Check if it's a gray color (R≈G≈B) and not white
                is_gray = (abs(r - g) <= 10 and abs(g - b) <= 10 and abs(r - b) <= 10)
                is_not_white = r < 250 and g < 250 and b < 250

                return is_gray and is_not_white

        return False
    except Exception as e:
        logger.debug(f"Error checking cell background: {e}")
        return False


def _get_cell_text(cell) -> str:
    """
    Extract text content from a table cell.

    Args:
        cell: python-pptx cell object

    Returns:
        Text content from the cell, or empty string if no text
    """
    try:
        if cell.text_frame and cell.text_frame.text:
            return cell.text_frame.text
        return ""
    except Exception:
        return ""


def _cells_are_same(cell1, cell2) -> bool:
    """
    Check if two cells are already merged (reference the same underlying cell).

    Args:
        cell1: First python-pptx cell object
        cell2: Second python-pptx cell object

    Returns:
        True if cells are already merged
    """
    try:
        # In python-pptx, merged cells share the same text_frame object
        return cell1.text_frame is cell2.text_frame
    except Exception:
        return False


def _apply_cell_styling(cell, text: str = None, font_size: int = None,
                       bold: bool = False, center_align: bool = False,
                       vertical_center: bool = False):
    """
    Apply formatting to a table cell.

    Args:
        cell: python-pptx cell object
        text: Text to set in cell (optional)
        font_size: Font size in points (optional)
        bold: Whether to make text bold
        center_align: Whether to center-align text horizontally
        vertical_center: Whether to center-align text vertically
    """
    try:
        text_frame = cell.text_frame

        # Enable word wrap (wrap on full words only, no mid-word breaks)
        text_frame.word_wrap = True

        # Enable text auto-fit: shrink text if needed to prevent overflow and mid-word breaks
        # This allows PowerPoint to automatically reduce font size to fit text properly
        text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        # Set margins to allow text to wrap properly within cell bounds
        # Minimal margins to maximize available space for text wrapping
        text_frame.margin_left = Pt(1)   # Minimal left margin
        text_frame.margin_right = Pt(1)  # Minimal right margin
        text_frame.margin_top = Pt(1)    # Minimal top margin
        text_frame.margin_bottom = Pt(1) # Minimal bottom margin

        # Set text if provided - this creates new runs with default formatting
        if text is not None:
            text_frame.text = text

        # Apply vertical alignment
        if vertical_center:
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Apply paragraph-level formatting
        if text_frame.paragraphs:
            paragraph = text_frame.paragraphs[0]

            if center_align:
                paragraph.alignment = PP_ALIGN.CENTER

            # CRITICAL: Setting text creates new runs, so we must format them
            # Force font formatting on all runs
            if paragraph.runs:
                for run in paragraph.runs:
                    # ENTRENCH Verdana font - always set it
                    run.font.name = FONT_NAME
                    if font_size is not None:
                        run.font.size = Pt(font_size)
                    if bold:
                        run.font.bold = True
            else:
                # Edge case: no runs exist yet, create one
                run = paragraph.add_run()
                run.text = text_frame.text
                run.font.name = FONT_NAME
                if font_size is not None:
                    run.font.size = Pt(font_size)
                if bold:
                    run.font.bold = True

    except Exception as e:
        logger.error(f"Failed to apply cell styling: {e}")
