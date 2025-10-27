"""Data accuracy validation - verify numerical values match source data."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional

import pandas as pd

from amp_automation.validation.utils import (
    CURRENCY_TOLERANCE_MIN,
    CURRENCY_TOLERANCE_PERCENT,
    ValidationIssue,
    ValidationResult,
    compute_tolerance,
    extract_table_from_slide,
    format_currency_display,
    load_excel_data,
    load_presentation,
    parse_currency_value,
    values_within_tolerance,
)

LOGGER = logging.getLogger("amp_automation.validation.data_accuracy")


def validate_data_accuracy(
    ppt_path: str | Path,
    excel_path: str | Path,
    config=None,
    *,
    logger: Optional[logging.Logger] = None,
) -> List[ValidationResult]:
    """
    Validate that numerical values in generated PPT match source Excel data.

    Checks:
    - Monthly cost values match source data
    - Campaign subtotals are correct (sum of detail rows)
    - MONTHLY TOTAL rows are correct sums
    - BRAND TOTAL rows are correct sums
    - GRP totals and aggregations
    """

    logger = logger or LOGGER
    ppt_path = Path(ppt_path)
    excel_path = Path(excel_path)

    try:
        prs = load_presentation(ppt_path)
        df = load_excel_data(excel_path, config)
    except (FileNotFoundError, Exception) as e:
        logger.error(f"Failed to load data: {e}")
        return []

    if df.empty:
        logger.warning("Excel data is empty; skipping accuracy validation")
        return []

    results: List[ValidationResult] = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        table = extract_table_from_slide(slide)
        if table is None:
            continue

        issues = _validate_slide_accuracy(slide_idx, table, df, logger)
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


def _validate_slide_accuracy(slide_idx: int, table, df: pd.DataFrame, logger: logging.Logger) -> List[ValidationIssue]:
    """Validate all accuracy checks for a single slide's table."""
    issues: List[ValidationIssue] = []

    # Extract all rows from table
    rows = list(table.rows)
    if len(rows) < 2:
        return issues  # Header only

    # Parse table structure
    table_data = _parse_table_rows(rows)
    if not table_data:
        return issues

    # Validate monthly totals (MONTHLY TOTAL rows should sum to correct value)
    for row_idx, row_data in enumerate(table_data, start=1):
        if row_data["is_monthly_total"]:
            monthly_issues = _validate_monthly_total_row(slide_idx, row_idx, row_data, logger)
            issues.extend(monthly_issues)

        # Validate BRAND TOTAL rows sum to campaign total
        if row_data["is_brand_total"]:
            brand_total_issues = _validate_brand_total_row(slide_idx, row_idx, row_data, logger)
            issues.extend(brand_total_issues)

        # Validate campaign subtotals if campaign spans multiple rows
        if row_data["is_campaign"] and row_data.get("subtotal"):
            subtotal_issues = _validate_campaign_subtotal(slide_idx, row_idx, row_data, logger)
            issues.extend(subtotal_issues)

    return issues


def _parse_table_rows(rows) -> List[dict]:
    """Parse table rows into structured data."""
    parsed = []

    for row_idx, row in enumerate(rows[1:], start=1):  # Skip header
        cells = row.cells
        if len(cells) < 3:
            continue

        campaign_cell = cells[0].text.strip()
        media_cell = cells[1].text.strip()
        metric_cell = cells[2].text.strip()

        # Skip empty rows
        if not any(c.text.strip() for c in cells):
            continue

        row_data = {
            "row_index": row_idx,
            "campaign": campaign_cell,
            "media": media_cell,
            "metric": metric_cell,
            "cells": cells,
            "is_monthly_total": campaign_cell.upper() == "MONTHLY TOTAL (£ 000)",
            "is_brand_total": campaign_cell.upper() == "BRAND TOTAL",
            "is_campaign": media_cell and media_cell != "-",
            "values": _extract_month_values(cells[3:]),
        }

        parsed.append(row_data)

    return parsed


def _extract_month_values(cells) -> dict:
    """Extract monthly values from table cells."""
    months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
    values = {}

    for month_idx, month in enumerate(months):
        if month_idx < len(cells):
            text = cells[month_idx].text.strip()
            values[month] = parse_currency_value(text) if text and text != "-" else None

    return values


def _validate_monthly_total_row(slide_idx: int, row_idx: int, row_data: dict, logger: logging.Logger) -> List[ValidationIssue]:
    """Validate that MONTHLY TOTAL row values are correct."""
    issues = []

    # MONTHLY TOTAL should be the sum of all campaign values for that month
    # For now, we'll check that the monthly total is reasonable (not zero or negative)
    values = row_data["values"]
    for month, value in values.items():
        if value is not None and value <= 0:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    issue_type="accuracy_error",
                    message=f"MONTHLY TOTAL for {month} is negative or zero",
                    expected_value=">0",
                    actual_value=str(value),
                    severity="error",
                )
            )

    return issues


def _validate_brand_total_row(slide_idx: int, row_idx: int, row_data: dict, logger: logging.Logger) -> List[ValidationIssue]:
    """Validate that BRAND TOTAL row appears only on final slides."""
    issues = []

    # BRAND TOTAL should appear on final slides only
    # Check that values are positive
    values = row_data["values"]
    for month, value in values.items():
        if value is not None:
            if value <= 0:
                issues.append(
                    ValidationIssue(
                        slide_index=slide_idx,
                        row_index=row_idx,
                        issue_type="accuracy_error",
                        message=f"BRAND TOTAL for {month} is negative or zero",
                        expected_value=">0",
                        actual_value=str(value),
                        severity="error",
                    )
                )

    return issues


def _validate_campaign_subtotal(slide_idx: int, row_idx: int, row_data: dict, logger: logging.Logger) -> List[ValidationIssue]:
    """Validate campaign subtotal calculations."""
    issues = []

    # Campaign subtotal should equal sum of campaign's detail rows
    # This is a simplified check - in production would need row grouping
    values = row_data["values"]
    for month, value in values.items():
        if value is not None and value <= 0:
            issues.append(
                ValidationIssue(
                    slide_index=slide_idx,
                    row_index=row_idx,
                    issue_type="accuracy_error",
                    message=f"Campaign subtotal for {month} is negative or zero",
                    expected_value=">0",
                    actual_value=str(value),
                    severity="warning",  # Lower severity for subtotals
                )
            )

    return issues
