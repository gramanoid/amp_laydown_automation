"""Unit tests for chart data preparation helpers."""

from __future__ import annotations

import pandas as pd

from amp_automation.presentation.charts import (
    prepare_campaign_type_chart_data,
    prepare_funnel_chart_data,
    prepare_media_type_chart_data,
)


def build_sample_frame() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Country": "UK",
                "Brand": "Sensodyne",
                "Year": "2025",
                "Media Type": "Television",
                "Total Cost": 100.0,
                "Funnel Stage": "Awareness",
                "Campaign Type": "Always On",
            },
            {
                "Country": "UK",
                "Brand": "Sensodyne",
                "Year": "2025",
                "Media Type": "Digital",
                "Total Cost": 50.0,
                "Funnel Stage": "Consideration",
                "Campaign Type": "Brand",
            },
        ]
    )


def test_prepare_media_type_chart_data_returns_totals() -> None:
    df = build_sample_frame()
    result = prepare_media_type_chart_data(df, "UK", "Sensodyne", "2025")

    assert result == {"Television": 100.0, "Digital": 50.0}


def test_prepare_funnel_chart_data_handles_empty() -> None:
    df = build_sample_frame()
    result = prepare_funnel_chart_data(df, "UK", "Sensodyne", "2025")

    assert result == {"Awareness": 100.0, "Consideration": 50.0}


def test_prepare_campaign_type_chart_data_normalises_strings() -> None:
    df = build_sample_frame()
    result = prepare_campaign_type_chart_data(df, "UK", "Sensodyne", "2025")

    assert result == {"Always On": 100.0, "Brand": 50.0}
