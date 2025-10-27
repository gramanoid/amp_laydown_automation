"""Shared validation utilities for data validation modules."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import pandas as pd
from pptx import Presentation

LOGGER = logging.getLogger("amp_automation.validation.utils")

NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d+)?")
MONTH_ORDER = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

# Standard tolerances for different validation types
CURRENCY_TOLERANCE_PERCENT = 0.005  # ±0.5%
CURRENCY_TOLERANCE_MIN = 100.0  # ±£100 minimum
PERCENTAGE_TOLERANCE = 0.005  # ±0.5%


@dataclass(slots=True)
class ValidationIssue:
    """Single validation failure."""

    slide_index: int
    campaign_name: Optional[str] = None
    row_index: Optional[int] = None
    issue_type: str = ""  # e.g., "missing_value", "format_error", "calculation_error"
    message: str = ""
    expected_value: Optional[str] = None
    actual_value: Optional[str] = None
    severity: str = "error"  # "error", "warning", "info"

    def __str__(self) -> str:
        parts = [f"Slide {self.slide_index}"]
        if self.campaign_name:
            parts.append(f"({self.campaign_name})")
        if self.row_index is not None:
            parts.append(f"Row {self.row_index}")
        parts.append(f": {self.message}")
        if self.expected_value and self.actual_value:
            parts.append(f" [expected: {self.expected_value}, got: {self.actual_value}]")
        return " ".join(parts)


@dataclass(slots=True)
class ValidationResult:
    """Aggregated validation results."""

    total_slides: int
    slides_with_issues: int
    total_issues: int
    issues: List[ValidationIssue]

    @property
    def passed(self) -> bool:
        return all(issue.severity != "error" for issue in self.issues)

    @property
    def error_count(self) -> int:
        return sum(1 for issue in self.issues if issue.severity == "error")

    @property
    def warning_count(self) -> int:
        return sum(1 for issue in self.issues if issue.severity == "warning")


def load_presentation(ppt_path: str | Path) -> Presentation:
    """Load a PowerPoint presentation safely."""
    ppt_path = Path(ppt_path)
    if not ppt_path.is_file():
        raise FileNotFoundError(f"Presentation not found: {ppt_path}")
    return Presentation(str(ppt_path))


def load_excel_data(excel_path: str | Path, config=None) -> pd.DataFrame:
    """Load Excel data using existing data ingestion pipeline."""
    from amp_automation.data import load_and_prepare_data
    from amp_automation.config.loader import Config

    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    if config is None:
        config = Config.default()

    dataset = load_and_prepare_data(excel_path, config, LOGGER)
    return dataset.frame.copy()


def extract_table_from_slide(slide, table_shape_name: str = "MainDataTable"):
    """Extract main data table from a slide."""
    for shape in slide.shapes:
        if getattr(shape, "name", "") == table_shape_name and hasattr(shape, "table"):
            return shape.table
    return None


def parse_currency_value(text: str) -> Optional[float]:
    """Parse currency text (e.g., '£123k') to float value."""
    if not text:
        return None
    # Remove currency symbols and spaces
    cleaned = text.replace("£", "").replace(",", "").strip()
    # Check for 'k' suffix (thousands)
    if cleaned.endswith("k"):
        cleaned = cleaned[:-1]
        try:
            return float(cleaned) * 1000
        except ValueError:
            return None
    # Try direct parsing
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_percentage_value(text: str) -> Optional[float]:
    """Parse percentage text (e.g., '45.2%') to float value (0-1 range)."""
    if not text:
        return None
    cleaned = text.replace("%", "").strip()
    try:
        numeric = float(cleaned)
        return numeric / 100.0  # Convert to 0-1 range
    except ValueError:
        return None


def parse_numeric_value(text: str) -> Optional[float]:
    """Parse generic numeric value from text."""
    if not text:
        return None
    match = NUMBER_PATTERN.search(text.replace(",", ""))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


def compute_tolerance(value: float, percent: float = CURRENCY_TOLERANCE_PERCENT, minimum: float = CURRENCY_TOLERANCE_MIN) -> float:
    """Compute tolerance range for numeric comparison."""
    if value is None:
        return minimum
    baseline = max(abs(value) * percent, minimum)
    return baseline


def values_within_tolerance(actual: float, expected: float, tolerance: float) -> bool:
    """Check if actual value is within tolerance of expected."""
    if actual is None or expected is None:
        return False
    return abs(actual - expected) <= tolerance


def format_currency_display(value: float, currency_symbol: str = "£") -> str:
    """Format numeric value as currency display."""
    if value >= 1000:
        return f"{currency_symbol}{value / 1000:.1f}k"
    return f"{currency_symbol}{value:.0f}"


def format_percentage_display(value: float, decimal_places: int = 1) -> str:
    """Format numeric value (0-1 range) as percentage display."""
    percent = value * 100
    if decimal_places == 0:
        return f"{percent:.0f}%"
    return f"{percent:.{decimal_places}f}%"


def results_to_dataframe(results: List[ValidationResult]) -> pd.DataFrame:
    """Convert validation results to DataFrame for reporting."""
    rows = []
    for result in results:
        for issue in result.issues:
            rows.append(
                {
                    "slide_index": issue.slide_index,
                    "campaign_name": issue.campaign_name or "",
                    "row_index": issue.row_index or "",
                    "issue_type": issue.issue_type,
                    "severity": issue.severity,
                    "message": issue.message,
                    "expected_value": issue.expected_value or "",
                    "actual_value": issue.actual_value or "",
                }
            )
    return pd.DataFrame(rows)


def write_validation_report(results: List[ValidationResult], output_path: str | Path) -> Path:
    """Write validation results to CSV file."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not results:
        # Write empty report with headers
        output_path.write_text(
            "slide_index,campaign_name,row_index,issue_type,severity,message,expected_value,actual_value\n",
            encoding="utf-8",
        )
        return output_path

    frame = results_to_dataframe(results)
    if frame.empty:
        output_path.write_text(
            "slide_index,campaign_name,row_index,issue_type,severity,message,expected_value,actual_value\n",
            encoding="utf-8",
        )
    else:
        frame.to_csv(output_path, index=False)

    return output_path


def summarize_validation_results(results: List[ValidationResult]) -> dict:
    """Create summary statistics from validation results."""
    total_issues = sum(r.total_issues for r in results)
    total_errors = sum(r.error_count for r in results)
    total_warnings = sum(r.warning_count for r in results)
    slides_with_issues = sum(r.slides_with_issues for r in results)
    total_slides = sum(r.total_slides for r in results)

    return {
        "total_slides": total_slides,
        "slides_with_issues": slides_with_issues,
        "total_issues": total_issues,
        "error_count": total_errors,
        "warning_count": total_warnings,
        "passed": all(r.passed for r in results),
    }
