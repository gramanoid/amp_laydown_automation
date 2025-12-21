#!/usr/bin/env python3
"""Diagnose CEJ percentage calculation for specific market/brand."""

import sys
sys.path.insert(0, "/Users/alexgrama/Developer/AMP Laydowns Automation")

import pandas as pd
from amp_automation.data.adapters import FlowplanAdapter

# Load data using same pipeline as presentation
adapter = FlowplanAdapter("/Users/alexgrama/Developer/AMP Laydowns Automation/input/Flowplan_Summaries_MEA_2025_12_17.xlsx")
df = adapter.normalize()

# Apply media type mapping (same as main code)
media_type_mapping = {
    "TV": "Television",
    "Digital": "Digital",
    "OOH": "OOH",
    "Radio": "Other",
    "Cinema": "Other",
    "Print": "Other",
}
df["Mapped Media Type"] = df["Media Type"].map(media_type_mapping).fillna("Other")

print("=" * 80)
print("DIAGNOSING: SAUDI ARABIA - SENSODYNE")
print("=" * 80)

# Filter for Saudi Arabia + Sensodyne (Slide 8)
# Adapter uses "Country" not "Market"
subset = df[(df["Country"] == "Saudi Arabia") & (df["Brand"] == "Sensodyne")]

print(f"\nTotal rows matching Saudi Arabia + Sensodyne: {len(subset)}")
print(f"\nUnique Funnel Stage values: {subset['Funnel Stage'].unique().tolist()}")
print(f"\nUnique Products: {subset['Product'].unique().tolist()}")

# Group by Funnel Stage and sum Total Cost
funnel_group = subset.groupby("Funnel Stage")["Total Cost"].sum()
print(f"\n--- Funnel Stage Totals ---")
for stage, cost in funnel_group.items():
    print(f"  {stage}: £{cost:,.2f}")

total_cost = subset["Total Cost"].sum()
print(f"\n  TOTAL: £{total_cost:,.2f}")

# Calculate percentages
print(f"\n--- CEJ Percentages ---")
for stage in ["Awareness", "Consideration", "Purchase"]:
    value = float(funnel_group.get(stage, 0.0))
    pct = (value / total_cost * 100) if total_cost > 0 else 0
    rounded_pct = round(pct)
    print(f"  {stage}: £{value:,.2f} / £{total_cost:,.2f} = {pct:.2f}% → rounds to {rounded_pct}%")

# Check if there are any variations in funnel stage naming
print(f"\n--- Raw Funnel Stage Value Counts ---")
print(subset["Funnel Stage"].value_counts())

# Check for any data at brand level vs product level
print(f"\n--- Breakdown by Product ---")
product_funnel = subset.groupby(["Product", "Funnel Stage"])["Total Cost"].sum().unstack(fill_value=0)
print(product_funnel)

# Check if maybe the code filters differently
print(f"\n--- Checking for potential data issues ---")
# Are there any null/empty funnel stages?
null_funnel = subset[subset["Funnel Stage"].isna() | (subset["Funnel Stage"] == "")]
if len(null_funnel) > 0:
    print(f"WARNING: {len(null_funnel)} rows with null/empty Funnel Stage")
    print(f"  Total Cost in null rows: £{null_funnel['Total Cost'].sum():,.2f}")
else:
    print("No null/empty Funnel Stage values found")

# Check for duplicates
print(f"\n--- Data Summary ---")
print(f"Columns in dataframe: {df.columns.tolist()}")
