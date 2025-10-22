"""Tests for data ingestion helpers."""

from __future__ import annotations

import logging

import pandas as pd
import pytest

from amp_automation.data.ingestion import (
    DataSet,
    _clean_brand,
    _extract_country,
    _validate_row_capacity,
    get_month_specific_tv_metrics,
)


def test_extract_country_handles_separator() -> None:
    assert _extract_country("Global | EMEA | Italy", " | ") == "Italy"
    assert _extract_country(None, " | ") is None


def test_clean_brand_strips_hierarchy() -> None:
    assert _clean_brand("Brand | Sub") == "Sub"
    assert _clean_brand(None) == ""


def test_validate_row_capacity_raises_for_small_dataset() -> None:
    frame = pd.DataFrame({"a": [1]})
    with pytest.raises(ValueError):
        _validate_row_capacity(frame, 2, logging.getLogger("test"))


def test_month_specific_tv_metrics_missing_file(tmp_path) -> None:
    missing = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        get_month_specific_tv_metrics(missing, "UK", "Brand", "Campaign", 2024, "Jan")
