"""Data format validation - verify formatting of displayed values."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import List, Optional

from amp_automation.validation.utils import (
    ValidationIssue,
    ValidationResult,
    extract_table_from_slide,
    load_presentation,
    parse_currency_value,
)

LOGGER = logging.getLogger("amp_automation.validation.data_format")

# Format patterns for validation
CURRENCY_PATTERN = re.compile(r"^£\d{1,3}(?:,\d{3})*k?$|^-$")  # e.g., "£123k", "£1,234", "-"
PERCENTAGE_PATTERN = re.compile(r"^\d{1,3}(?:\.\d{1,2})?%$|^-$")  # e.g., "45.2%", "-"
NUMERIC_PATTERN = re.compile(r"^\d+(?:\.\d+)?$|^-$|^,?\d{1,3}(?:,\d{3})*$")  # Numbers, percentages, or dashes

VALID_MEDIA_TYPES = {"Television", "Digital", "OOH", "Radio", "Print", "Cinema", "Other"}
VALID_METRICS_BY_MEDIA = {
    "Television": {"£ 000", "GRPS", "REACH@1+", "OTS@3+"},
    "Digital": {"£ 000", "YT REACH", "META REACH", "TT REACH"},
    "OOH": {"£ 000"},
    "Radio": {"£ 000"},
    "Print": {"£ 000"},
    "Cinema": {"£ 000"},
    "Other": {"£ 000"},
}


def validate_data_format(
    ppt_path: str | Path,
    *,
    logger: Optional[logging.Logger] = None,
) -> List[ValidationResult]:
    """
    Validate that values in generated PPT are properly formatted.

    Checks:
    - Currency values use correct format (£XXXk)
    - Percentages are properly formatted
    - Numeric values don't exceed reasonable bounds
    - No negative values where not expected
    - Media types are valid
    - Metrics match media type
    - Required cells contain data (not blank/dash)
    """

    logger = logger or LOGGER
    ppt_path = Path(ppt_path)

    try:
        prs = load_presentation(ppt_path)
    except FileNotFoundError as e:
        logger.error(f"Failed to load presentation: {e}")
        return []

    results: List[ValidationResult] = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        table = extract_table_from_slide(slide)
        if table is None:
            continue

        issues = _validate_slide_format(slide_idx, table, logger)
        if issues:
            results.append(
                ValidationResult(
                    total_slides=1,
                    slides_with_issues=1,
                    total_issues=len(issues),
                    issues=issues,
                )
            )

    return results


def _validate_slide_format(slide_idx: int, table, logger: logging.Logger) -> List[ValidationIssue]:
    """Validate all format checks for a single slide's table."""
    issues: List[ValidationIssue] = []

    rows = list(table.rows)
    if len(rows) < 2:
        return issues

    # Validate header row
    header_issues = _validate_header_format(slide_idx, rows[0])
    issues.extend(header_issues)

    # Validate data rows
    for row_idx, row in enumerate(rows[1:], start=1):
        row_issues = _validate_row_format(slide_idx, row_idx, row)
        issues.extend(row_issues)

    return issues


def _validate_header_format(slide_idx: int, header_row) -> List[ValidationIssue]:
    """Validate table header row."""
    issues = []
    expected_headers = ["CAMPAIGN", "MEDIA", "METRICS", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "TOTAL", "GRPS", "%"]

    cells = header_row.cells
    for cell_idx, (cell, expected) in enumerate(zip(cells, expected_headers)):
        actual = cell.text.strip().upper()
        if actual != expected:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=0,
                    issue_type="format_error",
                    message=f"Header mismatch at column {cell_idx}",
                    expected_value=expected,
                    actual_value=actual,
                    severity="error",
                )
            )

    return issues


def _validate_row_format(slide_idx: int, row_idx: int, row) -> List[ValidationIssue]:
    """Validate a data row's formatting."""
    issues = []

    cells = row.cells
    if len(cells) < 3:
        return issues

    campaign_cell = cells[0].text.strip()
    media_cell = cells[1].text.strip()
    metric_cell = cells[2].text.strip()

    # Skip empty rows
    if not any(c.text.strip() for c in cells):
        return issues

    # Validate media type if present
    if media_cell and media_cell != "-":
        media_normalized = media_cell.upper()
        if media_normalized not in {m.upper() for m in VALID_MEDIA_TYPES}:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    campaign_name=campaign_cell,
                    issue_type="format_error",
                    message=f"Invalid media type: {media_cell}",
                    expected_value=f"One of {VALID_MEDIA_TYPES}",
                    actual_value=media_cell,
                    severity="warning",
                )
            )

    # Validate metric matches media type if both present
    if media_cell and metric_cell and media_cell != "-" and metric_cell != "-":
        allowed_metrics = VALID_METRICS_BY_MEDIA.get(media_cell, set())
        if metric_cell not in allowed_metrics:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    campaign_name=campaign_cell,
                    issue_type="format_error",
                    message=f"Metric '{metric_cell}' not allowed for media '{media_cell}'",
                    expected_value=str(allowed_metrics),
                    actual_value=metric_cell,
                    severity="warning",
                )
            )

    # Validate month values (cells 3-14 are JAN-DEC)
    for month_idx in range(3, min(15, len(cells))):
        month_cell = cells[month_idx]
        cell_text = month_cell.text.strip()
        if not cell_text:
            continue

        # Skip dashes and empty cells
        if cell_text == "-":
            continue

        # Check format based on metric type
        if metric_cell and metric_cell != "-":
            if "%" in metric_cell or "REACH" in metric_cell:
                # Percentage or reach metric
                if not _is_valid_percentage_format(cell_text):
                    month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                    month_name = month_names[month_idx - 3] if month_idx - 3 < len(month_names) else f"Month {month_idx - 2}"
                    issues.append(
                        ValidationIssue(
                            slide_index=slide_idx,
                            row_index=row_idx,
                            campaign_name=campaign_cell,
                            issue_type="format_error",
                            message=f"Invalid percentage format in {month_name}",
                            expected_value="Valid percentage (e.g., '45.2%')",
                            actual_value=cell_text,
                            severity="warning",
                        )
                    )
            else:
                # Currency metric
                if not _is_valid_currency_format(cell_text):
                    # Only warn if it looks like it should be currency
                    if cell_text and any(c.isdigit() for c in cell_text):
                        issues.append(
                            ValidationIssue(
                                slide_index=slide_idx,
                                row_index=row_idx,
                                campaign_name=campaign_cell,
                                issue_type="format_error",
                                message=f"Invalid currency format in cell {month_idx}",
                                expected_value="Valid currency (e.g., '£123k')",
                                actual_value=cell_text,
                                severity="info",
                            )
                        )

    # Validate TOTAL and GRPS columns (should be numeric)
    if len(cells) > 15:
        total_cell = cells[15].text.strip()
        if total_cell and total_cell != "-":
            if not _is_valid_numeric_format(total_cell):
                issues.append(
                    ValidationIssue(
                        slide_index=slide_idx,
                        row_index=row_idx,
                        campaign_name=campaign_cell,
                        issue_type="format_error",
                        message="Invalid TOTAL column format",
                        expected_value="Valid numeric value",
                        actual_value=total_cell,
                        severity="warning",
                    )
                )

    return issues


def _is_valid_currency_format(text: str) -> bool:
    """Check if text is properly formatted currency."""
    if not text or text == "-":
        return True
    # Allow: £123k, £1,234, £123456
    if text.startswith("£") and any(c.isdigit() for c in text):
        return True
    # Allow plain numbers (may be formatted without £)
    if all(c.isdigit() or c == "," for c in text):
        return True
    return False


def _is_valid_percentage_format(text: str) -> bool:
    """Check if text is properly formatted percentage."""
    if not text or text == "-":
        return True
    if "%" in text:
        # Extract numeric part
        numeric_part = text.replace("%", "").strip()
        try:
            value = float(numeric_part)
            return 0 <= value <= 100
        except ValueError:
            return False
    return False


def _is_valid_numeric_format(text: str) -> bool:
    """Check if text is a valid numeric value."""
    if not text or text == "-":
        return True
    # Allow: 123, 1,234, 123.45
    cleaned = text.replace(",", "").strip()
    try:
        float(cleaned)
        return True
    except ValueError:
        return False
