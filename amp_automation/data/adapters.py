"""Input format adapters for different Excel export types.

This module implements the Adapter pattern to normalize different input formats
(BulkPlanData, Flowplan_Summaries) into a common schema for the pipeline.
"""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd


class InputFormat(Enum):
    """Supported input format types."""

    BULK_PLAN = "bulkplan"
    FLOWPLAN = "flowplan"
    AUTO = "auto"


@dataclass(slots=True)
class NormalizedData:
    """Container for normalized data ready for pipeline processing."""

    frame: pd.DataFrame
    source_format: InputFormat
    source_path: Path


# Common month alias mapping
MONTH_ALIAS_MAP = {
    "Jan": "Jan", "Feb": "Feb", "Mar": "Mar", "Apr": "Apr",
    "May": "May", "Jun": "Jun", "Jul": "Jul", "Aug": "Aug",
    "Sep": "Sep", "Sept": "Sep", "Oct": "Oct", "Nov": "Nov", "Dec": "Dec",
}
MONTH_ALIAS_MAP.update({k.upper(): v for k, v in MONTH_ALIAS_MAP.items()})

MONTHS_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Expected output schema columns
OUTPUT_SCHEMA = [
    "Country", "Brand", "Product", "Media Type", "Campaign Name",
    "Campaign Type", "Funnel Stage", "Year",
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    "Total Cost", "GRP", "Frequency", "Reach 1+", "Reach 3+", "Flight Comments",
]


class InputAdapter(ABC):
    """Base class for input format adapters."""

    def __init__(self, excel_path: Path, logger: Optional[logging.Logger] = None):
        self.excel_path = excel_path
        self.logger = logger or logging.getLogger("amp_automation.data.adapters")

    @abstractmethod
    def normalize(self) -> pd.DataFrame:
        """Transform input data into the common output schema.

        Returns:
            DataFrame with columns matching OUTPUT_SCHEMA.
        """
        pass

    @classmethod
    @abstractmethod
    def can_handle(cls, excel_path: Path) -> bool:
        """Check if this adapter can handle the given file.

        Args:
            excel_path: Path to the Excel file.

        Returns:
            True if this adapter can process the file.
        """
        pass

    def _ensure_output_schema(self, df: pd.DataFrame) -> pd.DataFrame:
        """Ensure DataFrame has all expected columns in correct order."""
        for column in OUTPUT_SCHEMA:
            if column not in df.columns:
                if column in {"GRP", "Frequency", "Reach 1+", "Reach 3+"}:
                    df[column] = np.nan
                elif column in MONTHS_ORDER or column == "Total Cost":
                    df[column] = 0
                else:
                    df[column] = ""
        return df[OUTPUT_SCHEMA]


class BulkPlanAdapter(InputAdapter):
    """Adapter for BulkPlanData Excel exports (Flight sheet format)."""

    # Geography normalization rules
    GEOGRAPHY_NORMALIZATIONS = {
        "Pakistan | Pakistan": "Pakistan",
        "South Africa | South Africa": "South Africa",
        "Turkey | Turkey": "Turkey",
        "East Africa | Kenya": "Kenya",
        "East Africa | Mauritius": "SSA",
        "East Africa | Nigeria": "Nigeria",
        "East Africa | Uganda": "SSA",
        "| FWA": "| FSA",
        "| GINE": "| GNE",
        "| KSA": "| Saudi Arabia",
        "| MOR": "| Maghreb",
    }

    # Product rename mapping
    PRODUCT_RENAMES = {
        "Panadol": "Panadol Product",
        "Panadol Cold and Flu": "Panadol RH",
    }

    @classmethod
    def can_handle(cls, excel_path: Path) -> bool:
        """Check if file has 'Flight' sheet characteristic of BulkPlanData."""
        try:
            xlsx = pd.ExcelFile(excel_path)
            return "Flight" in xlsx.sheet_names
        except Exception:
            return False

    def normalize(self) -> pd.DataFrame:
        """Transform BulkPlanData format into common schema."""
        self.logger.info("Loading BulkPlanData from %s", self.excel_path)

        raw_df = pd.read_excel(self.excel_path, sheet_name="Flight", header=0)
        self.logger.info("Loaded %s rows from BulkPlanData", len(raw_df))

        # Extract/create Month column
        raw_df = self._ensure_month_column(raw_df)

        # Apply data cleaning transformations
        raw_df = self._exclude_expert_campaigns(raw_df)
        raw_df = self._normalize_geography(raw_df)
        raw_df = self._split_panadol_brand(raw_df)
        raw_df = self._exclude_gne_pan_asian(raw_df)

        # Aggregate to monthly level
        agg_df = self._aggregate_to_monthly(raw_df)

        # Pivot to final row-per-campaign format
        result_df = self._pivot_to_final_format(agg_df)

        return self._ensure_output_schema(result_df)

    def _ensure_month_column(self, df: pd.DataFrame) -> pd.DataFrame:
        """Ensure Month column exists, extracting from flight date if needed."""
        month_column = df.get("Month")
        if month_column is None or month_column.isna().all():
            flight_col = "**Flight Start Date"
            if flight_col in df.columns:
                self.logger.info("Extracting Month from %s", flight_col)
                df[flight_col] = pd.to_datetime(df[flight_col])
                df["Month"] = df[flight_col].dt.strftime("%b")
            else:
                raise KeyError("Flight start date column not found for month extraction")
        return df

    def _exclude_expert_campaigns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Exclude rows where Plan Name contains 'expert'."""
        if "Plan Name" not in df.columns:
            return df
        initial_count = len(df)
        df = df[~df["Plan Name"].astype(str).str.contains("expert", case=False, na=False)]
        excluded = initial_count - len(df)
        self.logger.info("Excluded %s Expert campaign rows", excluded)
        return df

    def _normalize_geography(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply geography normalization rules."""
        if "Plan - Geography" not in df.columns:
            return df

        def normalize(geo_value):
            if pd.isna(geo_value):
                return geo_value
            geo_str = str(geo_value)
            for pattern, replacement in self.GEOGRAPHY_NORMALIZATIONS.items():
                if pattern in geo_str:
                    if pattern.startswith("|"):
                        if geo_str.endswith(pattern):
                            geo_str = geo_str.replace(pattern, replacement)
                    else:
                        geo_str = geo_str.replace(pattern, replacement)
            return geo_str

        df["Plan - Geography"] = df["Plan - Geography"].apply(normalize)
        self.logger.info("Applied geography normalization")
        return df

    def _split_panadol_brand(self, df: pd.DataFrame) -> pd.DataFrame:
        """Split Panadol brand based on Product Business (Pain vs Cold)."""
        if "Plan - Brand" not in df.columns or "**Product Business" not in df.columns:
            return df

        panadol_mask = df["Plan - Brand"].astype(str).str.contains("Panadol", case=False, na=False)
        if panadol_mask.sum() == 0:
            return df

        pain_mask = panadol_mask & df["**Product Business"].astype(str).str.contains("Pain", case=False, na=False)
        cold_mask = panadol_mask & df["**Product Business"].astype(str).str.contains("Cold", case=False, na=False)

        df.loc[pain_mask, "Plan - Brand"] = df.loc[pain_mask, "Plan - Brand"].str.replace(
            "Panadol", "Panadol Pain", case=False, regex=False
        )
        df.loc[cold_mask, "Plan - Brand"] = df.loc[cold_mask, "Plan - Brand"].str.replace(
            "Panadol", "Panadol C&F", case=False, regex=False
        )

        self.logger.info("Split Panadol: %s Pain, %s C&F", pain_mask.sum(), cold_mask.sum())
        return df

    def _exclude_gne_pan_asian(self, df: pd.DataFrame) -> pd.DataFrame:
        """Exclude GNE Pan Asian TV rows."""
        def should_exclude(row):
            geography = str(row.get("Plan - Geography", ""))
            media_type = row.get("Media Type")
            comments = str(row.get("Flight Comments", "")).strip()
            return "GNE" in geography and media_type == "Television" and "Pan Asian TV" in comments

        initial_count = len(df)
        df = df[~df.apply(should_exclude, axis=1)]
        self.logger.info("Filtered %s GNE Pan Asian TV rows", initial_count - len(df))
        return df

    def _extract_country(self, geo_value) -> Optional[str]:
        """Extract terminal country from hierarchical geography."""
        if geo_value is None or pd.isna(geo_value):
            return None
        parts = str(geo_value).split(" | ")
        return parts[-1].strip()

    def _clean_brand(self, brand_value) -> str:
        """Extract terminal brand from hierarchical value."""
        if brand_value is None or pd.isna(brand_value):
            return ""
        brand = str(brand_value)
        if " | " in brand:
            return brand.split(" | ")[-1].strip()
        return brand.strip()

    def _extract_product(self, product_business) -> str:
        """Extract product name from Product Business hierarchy."""
        if product_business is None or pd.isna(product_business):
            return ""
        parts = str(product_business).split(" | ")
        product = parts[-1].strip() if parts else ""
        return self.PRODUCT_RENAMES.get(product, product)

    def _aggregate_to_monthly(self, df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate raw data to monthly level."""
        group_cols = [
            "Plan - Geography", "Plan - Brand", "**Campaign Name(s)",
            "Plan - Year", "Month", "Media Type", "**Product Business",
        ]

        processed_data = []
        for name, group in df.groupby(group_cols):
            geo_raw, brand_raw, campaign, year, month_raw, media_type, product_business = name

            country = self._extract_country(geo_raw)
            brand = self._clean_brand(brand_raw)
            month = MONTH_ALIAS_MAP.get(month_raw, month_raw)
            product = self._extract_product(product_business)

            if not country or not brand or not campaign:
                continue

            total_cost = group["*Cost to Client"].sum()

            # TV metrics
            grp_sum = group["National GRP"].dropna().sum() if "National GRP" in group.columns else np.nan
            freq_avg = group["Frequency"].dropna().mean() if "Frequency" in group.columns else np.nan
            reach1_avg = group["Reach 1+"].dropna().mean() if "Reach 1+" in group.columns else np.nan
            reach3_avg = group["Reach 3+"].dropna().mean() if "Reach 3+" in group.columns else np.nan

            campaign_type = group["**Campaign Type"].dropna().iloc[0] if not group["**Campaign Type"].dropna().empty else ""
            funnel_stage = group["**Funnel Stage"].dropna().iloc[0] if not group["**Funnel Stage"].dropna().empty else ""

            processed_data.append({
                "Country": country,
                "Brand": brand,
                "Product": product,
                "Media Type": media_type,
                "Campaign Name": campaign,
                "Campaign Type": campaign_type,
                "Funnel Stage": funnel_stage,
                "Year": year,
                "Month": month,
                "Total Cost": total_cost if pd.notna(total_cost) else 0,
                "GRP": grp_sum if grp_sum > 0 else np.nan,
                "Frequency": freq_avg,
                "Reach 1+": reach1_avg,
                "Reach 3+": reach3_avg,
            })

        agg_df = pd.DataFrame(processed_data)
        self.logger.info("Created %s monthly aggregated rows", len(agg_df))
        return agg_df

    def _pivot_to_final_format(self, agg_df: pd.DataFrame) -> pd.DataFrame:
        """Pivot monthly data to final row-per-campaign format."""
        if agg_df.empty:
            raise ValueError("No data found after processing")

        final_group_cols = [
            "Country", "Brand", "Product", "Media Type",
            "Campaign Name", "Campaign Type", "Funnel Stage", "Year",
        ]

        result_rows = []
        for name, group in agg_df.groupby(final_group_cols):
            country, brand, product, media_type, campaign, campaign_type, funnel_stage, year = name

            row = {
                "Country": country, "Brand": brand, "Product": product,
                "Media Type": media_type, "Campaign Name": campaign,
                "Campaign Type": campaign_type, "Funnel Stage": funnel_stage,
                "Year": year,
            }

            # Initialize all months to 0
            for month in MONTHS_ORDER:
                row[month] = 0

            total_cost = 0
            total_grp = 0
            freq_values, reach1_values, reach3_values = [], [], []

            for _, month_row in group.iterrows():
                month = month_row["Month"]
                if month in MONTHS_ORDER:
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

        result_df = pd.DataFrame(result_rows)
        self.logger.info("Final BulkPlan dataset: %s rows", len(result_df))
        return result_df


class FlowplanAdapter(InputAdapter):
    """Adapter for Flowplan_Summaries Excel exports."""

    # Brand normalization mapping - aligns Flowplan names with config expectations
    BRAND_NORMALIZATIONS = {
        "Panadol (Adult Pain)": "Panadol Pain",
        "Panadol (Adult Cold)": "Panadol C&F",
        "Panadol (Child Pain)": "Panadol Child",
        "Pronamel": "Sensodyne Pronamel",
        "CAC": "Cac-1000",  # Combine CAC variants into single brand
        "Corega": "Polident",  # Same product, different regional name
        # Calpol stays separate (South Africa regional branding)
    }

    # Country normalization - combine Gulf countries into GNE region
    COUNTRY_NORMALIZATIONS = {
        "United Arab Emirates": "GNE",
        "Kuwait": "GNE",
        "Qatar": "GNE",
        "Bahrain": "GNE",
        "Oman": "GNE",
        "Iraq": "GNE",
        "Jordan": "GNE",
        "Lebanon": "GNE",
        "Morocco": "Maghreb",  # Regional grouping
        # Saudi Arabia stays separate
    }

    @classmethod
    def can_handle(cls, excel_path: Path) -> bool:
        """Check if file has Flowplan characteristics (Country.1 and [Current] columns)."""
        try:
            xlsx = pd.ExcelFile(excel_path)
            if "Sheet1" not in xlsx.sheet_names:
                return False
            # Read just header to check columns
            df = pd.read_excel(xlsx, sheet_name="Sheet1", nrows=0)
            has_country1 = "Country.1" in df.columns
            has_current_cost = "Cost to Client (GBP) [Current]" in df.columns
            return has_country1 and has_current_cost
        except Exception:
            return False

    def normalize(self) -> pd.DataFrame:
        """Transform Flowplan format into common schema."""
        self.logger.info("Loading Flowplan_Summaries from %s", self.excel_path)

        raw_df = pd.read_excel(self.excel_path, sheet_name="Sheet1", header=0)
        self.logger.info("Loaded %s rows from Flowplan", len(raw_df))

        # Filter out Expert campaigns
        raw_df = self._exclude_expert_campaigns(raw_df)

        # Convert Month datetime to string format
        raw_df = self._convert_month_format(raw_df)

        # Aggregate to monthly level
        agg_df = self._aggregate_to_monthly(raw_df)

        # Pivot to final row-per-campaign format
        result_df = self._pivot_to_final_format(agg_df)

        return self._ensure_output_schema(result_df)

    def _exclude_expert_campaigns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Exclude rows where Expert == 'Yes'."""
        if "Expert" not in df.columns:
            return df
        initial_count = len(df)
        df = df[df["Expert"].astype(str).str.lower() != "yes"]
        excluded = initial_count - len(df)
        self.logger.info("Excluded %s Expert campaign rows", excluded)
        return df

    def _convert_month_format(self, df: pd.DataFrame) -> pd.DataFrame:
        """Convert Month datetime to abbreviated string format."""
        if "Month" in df.columns:
            df["Month"] = pd.to_datetime(df["Month"]).dt.strftime("%b")
        return df

    def _normalize_brand(self, brand: str) -> str:
        """Normalize brand name to match config expectations."""
        return self.BRAND_NORMALIZATIONS.get(brand, brand)

    def _normalize_country(self, country: str) -> str:
        """Normalize country name (e.g., combine Gulf countries into GNE)."""
        return self.COUNTRY_NORMALIZATIONS.get(country, country)

    def _aggregate_to_monthly(self, df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate raw data to monthly level."""
        group_cols = [
            "Country.1",  # Use clean country name
            "Brand",
            "Campaign Name(s)",
            "Year",
            "Month",
            "Media Type",
            "Product",
        ]

        processed_data = []
        for name, group in df.groupby(group_cols):
            country, brand, campaign, year, month, media_type, product = name

            # Normalize country and brand names to match config expectations
            country = self._normalize_country(country)
            brand = self._normalize_brand(brand)
            month = MONTH_ALIAS_MAP.get(month, month)

            if not country or not brand or not campaign:
                continue

            # Cost - use [Current] variant
            total_cost = group["Cost to Client (GBP) [Current]"].sum()

            # TV metrics - use [Current] variants
            grp_col = "National GRP [Current]"
            freq_col = "Frequency [Current]"
            reach1_col = "Reach 1+ [Current]"
            reach3_col = "Reach 3+ [Current]"

            grp_sum = group[grp_col].dropna().sum() if grp_col in group.columns else np.nan
            freq_avg = group[freq_col].dropna().mean() if freq_col in group.columns else np.nan
            reach1_avg = group[reach1_col].dropna().mean() if reach1_col in group.columns else np.nan
            reach3_avg = group[reach3_col].dropna().mean() if reach3_col in group.columns else np.nan

            campaign_type = group["Campaign Type"].dropna().iloc[0] if not group["Campaign Type"].dropna().empty else ""
            funnel_stage = group["Funnel Stage"].dropna().iloc[0] if not group["Funnel Stage"].dropna().empty else ""

            processed_data.append({
                "Country": country,
                "Brand": brand,
                "Product": product,
                "Media Type": media_type,
                "Campaign Name": campaign,
                "Campaign Type": campaign_type,
                "Funnel Stage": funnel_stage,
                "Year": year,
                "Month": month,
                "Total Cost": total_cost if pd.notna(total_cost) else 0,
                "GRP": grp_sum if grp_sum > 0 else np.nan,
                "Frequency": freq_avg,
                "Reach 1+": reach1_avg,
                "Reach 3+": reach3_avg,
            })

        agg_df = pd.DataFrame(processed_data)
        self.logger.info("Created %s monthly aggregated rows", len(agg_df))
        return agg_df

    def _pivot_to_final_format(self, agg_df: pd.DataFrame) -> pd.DataFrame:
        """Pivot monthly data to final row-per-campaign format."""
        if agg_df.empty:
            raise ValueError("No data found after processing")

        final_group_cols = [
            "Country", "Brand", "Product", "Media Type",
            "Campaign Name", "Campaign Type", "Funnel Stage", "Year",
        ]

        result_rows = []
        for name, group in agg_df.groupby(final_group_cols):
            country, brand, product, media_type, campaign, campaign_type, funnel_stage, year = name

            row = {
                "Country": country, "Brand": brand, "Product": product,
                "Media Type": media_type, "Campaign Name": campaign,
                "Campaign Type": campaign_type, "Funnel Stage": funnel_stage,
                "Year": year,
            }

            # Initialize all months to 0
            for month in MONTHS_ORDER:
                row[month] = 0

            total_cost = 0
            total_grp = 0
            freq_values, reach1_values, reach3_values = [], [], []

            for _, month_row in group.iterrows():
                month = month_row["Month"]
                if month in MONTHS_ORDER:
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

        result_df = pd.DataFrame(result_rows)
        self.logger.info("Final Flowplan dataset: %s rows", len(result_df))
        return result_df


# Registry of available adapters (order matters for auto-detection)
ADAPTER_REGISTRY: list[type[InputAdapter]] = [
    FlowplanAdapter,  # Check first (more specific)
    BulkPlanAdapter,  # Fallback
]


def detect_format(excel_path: Path) -> InputFormat:
    """Auto-detect the input format based on file structure.

    Args:
        excel_path: Path to the Excel file.

    Returns:
        Detected InputFormat enum value.

    Raises:
        ValueError: If format cannot be determined.
    """
    for adapter_cls in ADAPTER_REGISTRY:
        if adapter_cls.can_handle(excel_path):
            if adapter_cls == FlowplanAdapter:
                return InputFormat.FLOWPLAN
            elif adapter_cls == BulkPlanAdapter:
                return InputFormat.BULK_PLAN
    raise ValueError(f"Unable to detect format for: {excel_path}")


def get_adapter(
    excel_path: Path,
    format_type: InputFormat = InputFormat.AUTO,
    logger: Optional[logging.Logger] = None,
) -> InputAdapter:
    """Get the appropriate adapter for the given file and format.

    Args:
        excel_path: Path to the Excel file.
        format_type: Explicit format or AUTO for detection.
        logger: Optional logger instance.

    Returns:
        Configured InputAdapter instance.

    Raises:
        ValueError: If no suitable adapter found.
    """
    if format_type == InputFormat.AUTO:
        format_type = detect_format(excel_path)

    adapter_map = {
        InputFormat.BULK_PLAN: BulkPlanAdapter,
        InputFormat.FLOWPLAN: FlowplanAdapter,
    }

    adapter_cls = adapter_map.get(format_type)
    if adapter_cls is None:
        raise ValueError(f"No adapter for format: {format_type}")

    return adapter_cls(excel_path, logger)
