"""Data ingestion and preparation routines."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping, Optional

import pandas as pd
import numpy as np

from amp_automation.config import Config


@dataclass(slots=True)
class DataSet:
    """Container for prepared data ready for presentation assembly."""

    frame: pd.DataFrame


MONTH_ALIAS_MAP = {
    "Jan": "Jan",
    "Feb": "Feb",
    "Mar": "Mar",
    "Apr": "Apr",
    "May": "May",
    "Jun": "Jun",
    "Jul": "Jul",
    "Aug": "Aug",
    "Sep": "Sep",
    "Sept": "Sep",
    "Oct": "Oct",
    "Nov": "Nov",
    "Dec": "Dec",
}
MONTH_ALIAS_MAP.update({k.upper(): v for k, v in MONTH_ALIAS_MAP.items()})


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


def _validate_row_capacity(data_frame: pd.DataFrame, min_rows: int, logger: logging.Logger) -> None:
    """Raise if the dataset does not meet the minimum row requirement."""

    if len(data_frame) < min_rows:
        logger.error("Dataset contains %s rows which is below the minimum threshold %s", len(data_frame), min_rows)
        raise ValueError("Insufficient data rows for presentation generation")


def load_and_prepare_data(excel_path: str | Path, config: Config, logger: logging.Logger) -> DataSet:
    """Load raw Excel data and return the cleaned dataset ready for slide assembly."""

    logger.info("Loading raw Lumina data from %s", excel_path)
    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel source not found: {excel_path}")

    data_section = config.section("data")
    excel_section = data_section.get("excel", {})
    validation_rules = excel_section.get("validation", {})
    geography_section = data_section.get("geography", {})
    media_section = data_section.get("media_types", {})

    raw_df = pd.read_excel(excel_path, header=0)
    logger.info("Loaded %s rows from raw data", len(raw_df))

    min_rows = int(validation_rules.get("min_rows", 1))
    _validate_row_capacity(raw_df, min_rows, logger)

    month_column = raw_df.get("Month")
    if month_column is None or month_column.isna().all():
        flight_column_name = excel_section.get("required_columns", {}).get("flight_start_date_name", "**Flight Start Date")
        if flight_column_name in raw_df.columns:
            logger.info("Month column missing or empty. Extracting from %s", flight_column_name)
            raw_df[flight_column_name] = pd.to_datetime(raw_df[flight_column_name])
            raw_df["Month"] = raw_df[flight_column_name].dt.strftime("%b")
        else:
            raise KeyError("Flight start date column not found for month extraction")

    if "Media Type" not in raw_df.columns:
        raise KeyError("Media Type column not found in source data")

    logger.debug("Media type distribution: %s", raw_df["Media Type"].value_counts().to_dict())

    separator = geography_section.get("separator", " | ")
    mapping = media_section.get("mapping", {})
    recognized_media_types = set(media_section.get("recognized", []))

    def should_exclude_row(row: pd.Series) -> bool:
        geography_raw = str(row.get("Plan - Geography", ""))
        media_type = row.get("Media Type")
        flight_comments = str(row.get("Flight Comments", "")).strip()
        if "GNE" in geography_raw and media_type == "Television" and "Pan Asian TV" in flight_comments:
            return True
        return False

    initial_count = len(raw_df)
    raw_df = raw_df[~raw_df.apply(should_exclude_row, axis=1)]
    logger.info("Filtered out %s rows via GNE Pan Asian TV rule", initial_count - len(raw_df))

    group_cols = [
        "Plan - Geography",
        "Plan - Brand",
        "**Campaign Name(s)",
        "Plan - Year",
        "Month",
        "Media Type",
    ]

    processed_data: list[dict[str, object]] = []

    for name, group in raw_df.groupby(group_cols):
        geography_raw, brand_raw, campaign, year, month_raw, media_type = name

        country = _extract_country(geography_raw, separator)
        brand = _clean_brand(brand_raw)
        month = MONTH_ALIAS_MAP.get(month_raw, month_raw)

        if not country or not brand or not campaign:
            continue

        total_cost = group["*Cost to Client"].sum()

        grp_sum = np.nan
        freq_avg = np.nan
        reach1_avg = np.nan
        reach3_avg = np.nan

        if media_type == "Television":
            if "National GRP" in group.columns:
                grp_values = group["National GRP"].dropna()
                if len(grp_values) > 0:
                    grp_sum = grp_values.sum()
            if "Frequency" in group.columns:
                freq_values = group["Frequency"].dropna()
                if len(freq_values) > 0:
                    freq_avg = freq_values.mean()
            if "Reach 1+" in group.columns:
                reach1_values = group["Reach 1+"].dropna()
                if len(reach1_values) > 0:
                    reach1_avg = reach1_values.mean()
            if "Reach 3+" in group.columns:
                reach3_values = group["Reach 3+"].dropna()
                if len(reach3_values) > 0:
                    reach3_avg = reach3_values.mean()

        campaign_type = group["**Campaign Type"].dropna().iloc[0] if not group["**Campaign Type"].dropna().empty else ""
        funnel_stage = group["**Funnel Stage"].dropna().iloc[0] if not group["**Funnel Stage"].dropna().empty else ""

        processed_data.append(
            {
                "Country": country,
                "Brand": brand,
                "Media Type": media_type,
                "Campaign Name": campaign,
                "Campaign Type": campaign_type,
                "Funnel Stage": funnel_stage,
                "Year": year,
                "Month": month,
                "Total Cost": total_cost if pd.notna(total_cost) else 0,
                "GRP": grp_sum,
                "Frequency": freq_avg,
                "Reach 1+": reach1_avg,
                "Reach 3+": reach3_avg,
            }
        )

    agg_df = pd.DataFrame(processed_data)
    if agg_df.empty:
        raise ValueError("No data found after processing")

    logger.info("Created %s month-level aggregated rows", len(agg_df))

    final_group_cols = [
        "Country",
        "Brand",
        "Media Type",
        "Campaign Name",
        "Campaign Type",
        "Funnel Stage",
        "Year",
    ]

    result_rows: list[dict[str, object]] = []
    months = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]

    for name, group in agg_df.groupby(final_group_cols):
        country, brand, media_type, campaign, campaign_type, funnel_stage, year = name

        row = {
            "Country": country,
            "Brand": brand,
            "Media Type": media_type,
            "Campaign Name": campaign,
            "Campaign Type": campaign_type,
            "Funnel Stage": funnel_stage,
            "Year": year,
        }

        for month in months:
            row[month] = 0

        total_cost = 0
        total_grp = 0
        freq_values: list[float] = []
        reach1_values: list[float] = []
        reach3_values: list[float] = []

        for _, month_row in group.iterrows():
            month = month_row["Month"]
            if month in months:
                row[month] = month_row["Total Cost"]
                total_cost += month_row["Total Cost"]

                if pd.notna(month_row["GRP"]):
                    total_grp += month_row["GRP"]
                if pd.notna(month_row["Frequency"]):
                    freq_values.append(month_row["Frequency"])
                if pd.notna(month_row["Reach 1+"]):
                    reach1_values.append(month_row["Reach 1+"])
                if pd.notna(month_row["Reach 3+"]):
                    reach3_values.append(month_row["Reach 3+"])

        row["Total Cost"] = total_cost
        row["GRP"] = total_grp if total_grp > 0 else np.nan
        row["Frequency"] = float(np.mean(freq_values)) if freq_values else np.nan
        row["Reach 1+"] = float(np.mean(reach1_values)) if reach1_values else np.nan
        row["Reach 3+"] = float(np.mean(reach3_values)) if reach3_values else np.nan
        row["Flight Comments"] = ""

        result_rows.append(row)

    df = pd.DataFrame(result_rows)
    logger.info("Final dataset prepared with shape %s", df.shape)

    expected_columns = [
        "Country",
        "Brand",
        "Media Type",
        "Campaign Name",
        "Campaign Type",
        "Funnel Stage",
        "Year",
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
        "Total Cost",
        "GRP",
        "Frequency",
        "Reach 1+",
        "Reach 3+",
        "Flight Comments",
    ]

    for column in expected_columns:
        if column not in df.columns:
            if column in {"GRP", "Frequency", "Reach 1+", "Reach 3+"}:
                df[column] = np.nan
            elif column in months or column == "Total Cost":
                df[column] = 0
            else:
                df[column] = ""

    df = df[expected_columns]
    df["Mapped Media Type"] = df["Media Type"].map(lambda m: mapping.get(m, m))

    tv_campaigns_with_metrics = len(df[(df["Media Type"] == "Television") & df["GRP"].notna()])
    logger.info("TV campaigns with metrics: %s", tv_campaigns_with_metrics)

    return DataSet(frame=df)


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
    """Aggregate month-specific TV metrics for a campaign."""

    logger = logger or logging.getLogger("amp_automation.data")
    raw_excel_path = Path(raw_excel_path)

    if not raw_excel_path.is_file():
        raise FileNotFoundError(raw_excel_path)

    cache_attr = "_cached_data"
    cached_df = getattr(get_month_specific_tv_metrics, cache_attr, None)

    if cached_df is None or getattr(get_month_specific_tv_metrics, "_cached_path", None) != raw_excel_path:
        df = pd.read_excel(raw_excel_path, header=0)
        if "Month" not in df.columns or df["Month"].isna().all():
            if "**Flight Start Date" in df.columns:
                df["**Flight Start Date"] = pd.to_datetime(df["**Flight Start Date"])
                df["Month"] = df["**Flight Start Date"].dt.strftime("%b")
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

    gne_mask = filtered_data["Plan - Geography"].astype(str).str.contains("GNE", na=False)
    pan_asian_mask = filtered_data["Flight Comments"].astype(str).str.contains("Pan Asian TV", na=False)
    filtered_data = filtered_data[~(gne_mask & pan_asian_mask)]

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
