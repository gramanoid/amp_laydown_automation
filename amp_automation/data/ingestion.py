"""Data ingestion and preparation routines."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping, Optional

import pandas as pd
import numpy as np

from amp_automation.config import Config
from amp_automation.data.adapters import (
    InputFormat,
    NormalizedData,
    get_adapter,
    detect_format,
    MONTH_ALIAS_MAP,
    OUTPUT_SCHEMA,
)


@dataclass(slots=True)
class DataSet:
    """Container for prepared data ready for presentation assembly."""

    frame: pd.DataFrame
    source_format: Optional[InputFormat] = None


def _validate_row_capacity(data_frame: pd.DataFrame, min_rows: int, logger: logging.Logger) -> None:
    """Raise if the dataset does not meet the minimum row requirement."""

    if len(data_frame) < min_rows:
        logger.error("Dataset contains %s rows which is below the minimum threshold %s", len(data_frame), min_rows)
        raise ValueError("Insufficient data rows for presentation generation")


def load_and_prepare_data(
    excel_path: str | Path,
    config: Config,
    logger: logging.Logger,
    *,
    format_type: InputFormat = InputFormat.AUTO,
) -> DataSet:
    """Load raw Excel data and return the cleaned dataset ready for slide assembly.

    This function uses the adapter pattern to support multiple input formats:
    - BulkPlanData (Flight sheet format from Lumina)
    - Flowplan_Summaries (aggregated report format)

    Args:
        excel_path: Path to the input Excel file.
        config: Configuration object.
        logger: Logger instance.
        format_type: Explicit format or AUTO for auto-detection.

    Returns:
        DataSet containing the prepared DataFrame.
    """
    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel source not found: {excel_path}")

    # Get the appropriate adapter
    adapter = get_adapter(excel_path, format_type, logger)
    detected_format = detect_format(excel_path) if format_type == InputFormat.AUTO else format_type
    logger.info("Using %s adapter for %s", detected_format.value, excel_path.name)

    # Normalize data through adapter
    df = adapter.normalize()

    # Validate minimum rows
    data_section = config.section("data")
    excel_section = data_section.get("excel", {})
    validation_rules = excel_section.get("validation", {})
    min_rows = int(validation_rules.get("min_rows", 1))
    _validate_row_capacity(df, min_rows, logger)

    # Apply media type mapping from config
    media_section = data_section.get("media_types", {})
    mapping = media_section.get("mapping", {})
    df["Mapped Media Type"] = df["Media Type"].map(lambda m: mapping.get(m, m))

    # Log TV metrics summary
    tv_campaigns_with_metrics = len(df[(df["Media Type"] == "Television") & df["GRP"].notna()])
    logger.info("TV campaigns with metrics: %s", tv_campaigns_with_metrics)
    logger.info("Final dataset prepared with shape %s", df.shape)

    return DataSet(frame=df, source_format=detected_format)


# --- Legacy function for month-specific TV metrics lookup ---
# This function is kept for backward compatibility and TV metric lookups
# during assembly. It only works with BulkPlanData format.

def _extract_country(raw_value: str | float | None, separator: str) -> Optional[str]:
    """Extract the terminal geography token from a hierarchical value."""
    if raw_value is None or pd.isna(raw_value):
        return None
    parts = str(raw_value).split(separator)
    return parts[-1].strip()


def _clean_brand(raw_value: str | float | None) -> str:
    """Normalise a raw brand cell into its display-ready form."""
    if raw_value is None or pd.isna(raw_value):
        return ""
    brand = str(raw_value)
    if " | " in brand:
        return brand.split(" | ")[-1].strip()
    return brand.strip()


def get_month_specific_tv_metrics(
    raw_excel_path: str | Path,
    country: str,
    brand: str,
    campaign: str,
    year: int,
    month: str,
    *,
    logger: Optional[logging.Logger] = None,
) -> dict[str, float]:
    """Aggregate month-specific TV metrics for a campaign.

    Note: This function currently only supports BulkPlanData format.
    For Flowplan format, metrics are already aggregated in the main data.
    """
    logger = logger or logging.getLogger("amp_automation.data")
    raw_excel_path = Path(raw_excel_path)

    if not raw_excel_path.is_file():
        raise FileNotFoundError(raw_excel_path)

    # Check if this is a Flowplan file - return empty metrics
    # (metrics are already in the main dataset for Flowplan)
    try:
        detected = detect_format(raw_excel_path)
        if detected == InputFormat.FLOWPLAN:
            logger.debug("Flowplan format detected - TV metrics from main dataset")
            return {
                "grp_sum": 0,
                "frequency_avg": np.nan,
                "reach1_avg": np.nan,
                "reach3_avg": np.nan,
            }
    except ValueError:
        pass  # Continue with BulkPlan logic

    cache_attr = "_cached_data"
    cached_df = getattr(get_month_specific_tv_metrics, cache_attr, None)

    if cached_df is None or getattr(get_month_specific_tv_metrics, "_cached_path", None) != raw_excel_path:
        df = pd.read_excel(raw_excel_path, sheet_name="Flight", header=0)
        if "Month" not in df.columns or df["Month"].isna().all():
            if "**Flight Start Date" in df.columns:
                df["**Flight Start Date"] = pd.to_datetime(df["**Flight Start Date"])
                df["Month"] = df["**Flight Start Date"].dt.strftime("%b")

        # Apply same transformations as BulkPlanAdapter for consistency
        # 1. Geography normalization
        if "Plan - Geography" in df.columns:
            def normalize_geography(geo_value: str) -> str:
                if pd.isna(geo_value):
                    return geo_value
                geo_str = str(geo_value)
                # Remove duplicate segments
                if geo_str.endswith("Pakistan | Pakistan"):
                    geo_str = geo_str.replace("Pakistan | Pakistan", "Pakistan")
                elif geo_str.endswith("South Africa | South Africa"):
                    geo_str = geo_str.replace("South Africa | South Africa", "South Africa")
                elif geo_str.endswith("Turkey | Turkey"):
                    geo_str = geo_str.replace("Turkey | Turkey", "Turkey")
                # Remove "East Africa" layer
                if "East Africa | Kenya" in geo_str:
                    geo_str = geo_str.replace("East Africa | Kenya", "Kenya")
                elif "East Africa | Mauritius" in geo_str:
                    geo_str = geo_str.replace("East Africa | Mauritius", "SSA")
                elif "East Africa | Nigeria" in geo_str:
                    geo_str = geo_str.replace("East Africa | Nigeria", "Nigeria")
                elif "East Africa | Uganda" in geo_str:
                    geo_str = geo_str.replace("East Africa | Uganda", "SSA")
                # Rename region codes
                if geo_str.endswith("| FWA"):
                    geo_str = geo_str.replace("| FWA", "| FSA")
                elif geo_str.endswith("| GINE"):
                    geo_str = geo_str.replace("| GINE", "| GNE")
                elif geo_str.endswith("| KSA"):
                    geo_str = geo_str.replace("| KSA", "| Saudi Arabia")
                elif geo_str.endswith("| MOR"):
                    geo_str = geo_str.replace("| MOR", "| Maghreb")
                return geo_str
            df["Plan - Geography"] = df["Plan - Geography"].apply(normalize_geography)

        # 2. Panadol brand splitting
        if "Plan - Brand" in df.columns and "**Product Business" in df.columns:
            panadol_mask = df["Plan - Brand"].astype(str).str.contains("Panadol", case=False, na=False)
            if panadol_mask.sum() > 0:
                pain_mask = panadol_mask & df["**Product Business"].astype(str).str.contains("Pain", case=False, na=False)
                cold_mask = panadol_mask & df["**Product Business"].astype(str).str.contains("Cold", case=False, na=False)
                df.loc[pain_mask, "Plan - Brand"] = df.loc[pain_mask, "Plan - Brand"].str.replace(
                    "Panadol", "Panadol Pain", case=False, regex=False
                )
                df.loc[cold_mask, "Plan - Brand"] = df.loc[cold_mask, "Plan - Brand"].str.replace(
                    "Panadol", "Panadol C&F", case=False, regex=False
                )

        setattr(get_month_specific_tv_metrics, cache_attr, df)
        setattr(get_month_specific_tv_metrics, "_cached_path", raw_excel_path)
    else:
        df = cached_df

    separator = " | "
    month_normalised = MONTH_ALIAS_MAP.get(month, month)

    filtered_data = df[
        (df["Plan - Geography"].apply(lambda value: _extract_country(value, separator)) == country)
        & (df["Plan - Brand"].apply(_clean_brand) == brand)
        & (df["**Campaign Name(s)"] == campaign)
        & (df["Plan - Year"] == year)
        & (df["Month"] == month_normalised)
        & (df["Media Type"] == "Television")
    ]

    # Exclude GNE Pan Asian TV campaigns
    gne_mask = filtered_data["Plan - Geography"].astype(str).str.contains("GNE", na=False)
    pan_asian_mask = filtered_data["Flight Comments"].astype(str).str.contains("Pan Asian TV", na=False)
    filtered_data = filtered_data[~(gne_mask & pan_asian_mask)]

    # Exclude Expert campaigns (same filter as in adapter)
    if "Plan Name" in filtered_data.columns:
        filtered_data = filtered_data[~filtered_data["Plan Name"].astype(str).str.contains("expert", case=False, na=False)]

    if filtered_data.empty:
        return {
            "grp_sum": 0,
            "frequency_avg": np.nan,
            "reach1_avg": np.nan,
            "reach3_avg": np.nan,
        }

    available_cols = filtered_data.columns.tolist()
    missing_cols = [col for col in ["National GRP", "Frequency", "Reach 1+", "Reach 3+"] if col not in available_cols]
    if missing_cols:
        logger.warning("Missing TV metric columns in month-specific function: %s", missing_cols)

    grp_sum = filtered_data["National GRP"].dropna().sum() if "National GRP" in available_cols else 0

    freq_values = filtered_data["Frequency"].dropna() if "Frequency" in available_cols else pd.Series(dtype=float)
    frequency_avg = freq_values.mean() if len(freq_values) > 0 else np.nan

    reach1_values = filtered_data["Reach 1+"].dropna() if "Reach 1+" in available_cols else pd.Series(dtype=float)
    reach1_avg = reach1_values.mean() if len(reach1_values) > 0 else np.nan

    reach3_values = filtered_data["Reach 3+"].dropna() if "Reach 3+" in available_cols else pd.Series(dtype=float)
    reach3_avg = reach3_values.mean() if len(reach3_values) > 0 else np.nan

    return {
        "grp_sum": grp_sum,
        "frequency_avg": frequency_avg,
        "reach1_avg": reach1_avg,
        "reach3_avg": reach3_avg,
    }
