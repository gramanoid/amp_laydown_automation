"""Data completeness validation - verify all required data is present."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional

from amp_automation.validation.utils import (
    ValidationIssue,
    ValidationResult,
    extract_table_from_slide,
    load_presentation,
)

LOGGER = logging.getLogger("amp_automation.validation.data_completeness")

# Column indices
CAMPAIGN_COL = 0
MEDIA_COL = 1
METRIC_COL = 2
MONTH_START_COL = 3
MONTH_END_COL = 15  # DEC is at index 14, so range is 3:15


def validate_data_completeness(
    ppt_path: str | Path,
    *,
    logger: Optional[logging.Logger] = None,
) -> List[ValidationResult]:
    """
    Validate that all required data is present in generated PPT.

    Checks:
    - Campaign names are populated
    - Media types are populated
    - Metrics are populated
    - At least some month data exists for each row
    - No completely blank campaigns
    - Proper data density (not too sparse)
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

        issues = _validate_slide_completeness(slide_idx, table, logger)
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


def _validate_slide_completeness(slide_idx: int, table, logger: logging.Logger) -> List[ValidationIssue]:
    """Validate data completeness for a single slide."""
    issues: List[ValidationIssue] = []

    rows = list(table.rows)
    if len(rows) < 2:
        return issues

    for row_idx, row in enumerate(rows[1:], start=1):
        row_issues = _validate_row_completeness(slide_idx, row_idx, row)
        issues.extend(row_issues)

    return issues


def _validate_row_completeness(slide_idx: int, row_idx: int, row) -> List[ValidationIssue]:
    """Validate completeness of a single row."""
    issues = []

    cells = row.cells
    if len(cells) < MONTH_START_COL:
        return issues

    # Extract cell contents
    campaign = cells[CAMPAIGN_COL].text.strip() if len(cells) > CAMPAIGN_COL else ""
    media = cells[MEDIA_COL].text.strip() if len(cells) > MEDIA_COL else ""
    metric = cells[METRIC_COL].text.strip() if len(cells) > METRIC_COL else ""

    # Skip special rows that are allowed to be empty
    if _is_special_row(campaign):
        return issues

    # Skip completely empty rows
    if not any(cells[col].text.strip() for col in range(min(3, len(cells)))):
        return issues

    # Check campaign name is populated for data rows
    if media and media != "-":
        if not campaign:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    issue_type="completeness_error",
                    message="Campaign name missing for data row",
                    expected_value="Campaign name required",
                    actual_value="",
                    severity="error",
                )
            )

        # Check metric is populated for media rows
        if not metric or metric == "-":
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    campaign_name=campaign,
                    issue_type="completeness_error",
                    message=f"Metric missing for {media} row",
                    expected_value="Metric name required",
                    actual_value=metric or "(empty)",
                    severity="error",
                )
            )

    # Check at least some month data exists for campaign rows
    if campaign and media and media != "-" and not _is_total_row(campaign):
        month_data = cells[MONTH_START_COL:MONTH_END_COL]
        non_empty_months = sum(1 for cell in month_data if cell.text.strip() and cell.text.strip() != "-")

        if non_empty_months == 0:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    campaign_name=campaign,
                    issue_type="completeness_error",
                    message="No monthly data values found (all empty or dashes)",
                    expected_value="At least one month with data",
                    actual_value="None",
                    severity="warning",
                )
            )
        elif non_empty_months < 3:
            # Warn if suspiciously sparse (less than 3 months)
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    campaign_name=campaign,
                    issue_type="completeness_error",
                    message=f"Limited monthly data: only {non_empty_months} months have values",
                    expected_value="Data in most months (≥3)",
                    actual_value=f"{non_empty_months} months",
                    severity="info",
                )
            )

    # Validate data density in table
    _validate_table_data_density(slide_idx, cells)

    return issues


def _validate_table_data_density(slide_idx: int, cells) -> List[ValidationIssue]:
    """Validate that table has reasonable data density."""
    # This is checked at slide level, not row level
    # For now, return empty - could be enhanced to check overall table fullness
    return []


def _is_special_row(campaign_text: str) -> bool:
    """Check if row is a special row (totals, subtotals) that can be sparse."""
    upper = campaign_text.upper()
    special_indicators = {
        "MONTHLY TOTAL",
        "BRAND TOTAL",
        "SUBTOTAL",
        "TOTAL",
        "CONTINUED FROM",
        "-",
        "",
    }
    for indicator in special_indicators:
        if indicator in upper:
            return True
    return False


def _is_total_row(campaign_text: str) -> bool:
    """Check if row is a total/subtotal row."""
    upper = campaign_text.upper()
    return "TOTAL" in upper or "SUBTOTAL" in upper
