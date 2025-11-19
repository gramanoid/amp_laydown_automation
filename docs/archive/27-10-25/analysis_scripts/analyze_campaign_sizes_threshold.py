#!/usr/bin/env python3
"""
Campaign Size Distribution Analysis with Configurable Threshold

Analyzes BulkPlanData Excel to determine:
- Row count distribution per campaign
- Percentage of campaigns exceeding threshold
- Impact of smart pagination on slide count at different thresholds
"""

import pandas as pd
import sys
from pathlib import Path
from collections import Counter, defaultdict

def analyze_campaign_sizes(excel_path, threshold=32):
    """Analyze campaign size distribution from Excel data."""

    print(f"Reading Excel data from: {excel_path}")
    print(f"Using threshold: {threshold} rows per slide\n")
    df = pd.read_excel(excel_path)

    print(f"Total rows in Excel: {len(df)}")
    print(f"Total columns: {len(df.columns)}\n")

    # Group by market/brand/year to simulate presentation structure
    # Each unique combination becomes a slide deck
    grouping_cols = []

    # Determine grouping columns
    if '**Market Sub-Cluster' in df.columns:
        grouping_cols.append('**Market Sub-Cluster')
    if 'Plan - Brand' in df.columns:
        grouping_cols.append('Plan - Brand')
    if 'Plan - Year' in df.columns:
        grouping_cols.append('Plan - Year')

    print(f"Grouping by: {grouping_cols}\n")

    # Group data
    grouped = df.groupby(grouping_cols)

    print(f"Total unique market/brand/year combinations: {len(grouped)}\n")

    # For each market/brand/year, analyze campaigns
    all_campaign_sizes = []
    deck_campaign_counts = []

    for (market, brand, year), group_df in grouped:
        deck_id = f"{market} - {brand} ({year})"

        # Within this deck, group by campaign
        if '**Campaign Name(s)' in group_df.columns:
            campaign_col = '**Campaign Name(s)'
        elif '**Campaign (String)' in group_df.columns:
            campaign_col = '**Campaign (String)'
        else:
            # No campaign column, treat entire deck as one campaign
            campaign_col = None

        if campaign_col and campaign_col in group_df.columns:
            campaigns = group_df.groupby(campaign_col)

            deck_campaigns = []
            for campaign_name, campaign_df in campaigns:
                campaign_size = len(campaign_df)
                all_campaign_sizes.append(campaign_size)
                deck_campaigns.append({
                    'name': campaign_name,
                    'size': campaign_size
                })

            deck_campaign_counts.append({
                'deck_id': deck_id,
                'campaign_count': len(deck_campaigns),
                'campaigns': deck_campaigns
            })
        else:
            # No campaign grouping, entire deck is one unit
            deck_size = len(group_df)
            all_campaign_sizes.append(deck_size)
            deck_campaign_counts.append({
                'deck_id': deck_id,
                'campaign_count': 1,
                'campaigns': [{'name': 'Entire Deck', 'size': deck_size}]
            })

    # Calculate statistics
    total_campaigns = len(all_campaign_sizes)
    campaigns_over_threshold = sum(1 for size in all_campaign_sizes if size > threshold)
    pct_over_threshold = (campaigns_over_threshold / total_campaigns * 100) if total_campaigns > 0 else 0

    avg_campaign_size = sum(all_campaign_sizes) / total_campaigns if total_campaigns > 0 else 0

    # Print results
    print("=" * 80)
    print(f"CAMPAIGN SIZE DISTRIBUTION ANALYSIS (THRESHOLD: {threshold} ROWS)")
    print("=" * 80)
    print()

    print(f"Total campaigns analyzed: {total_campaigns}")
    print(f"Campaigns exceeding {threshold} rows: {campaigns_over_threshold} ({pct_over_threshold:.1f}%)")
    print(f"Campaigns <={threshold} rows: {total_campaigns - campaigns_over_threshold} ({100 - pct_over_threshold:.1f}%)")
    print(f"Average campaign size: {avg_campaign_size:.1f} rows")
    print()

    # Estimate slide count impact
    print("=" * 80)
    print(f"SMART PAGINATION IMPACT ESTIMATE (THRESHOLD: {threshold} ROWS)")
    print("=" * 80)
    print()

    # Current (sequential fill to threshold): campaigns can split
    current_slide_estimate = 0
    for deck_info in deck_campaign_counts:
        total_rows = sum(c['size'] for c in deck_info['campaigns'])
        # Add 1 row per campaign for MONTHLY TOTAL
        total_rows += deck_info['campaign_count']
        # Slides = ceil(total_rows / threshold)
        slides_needed = (total_rows + threshold - 1) // threshold
        current_slide_estimate += slides_needed

    # Smart pagination: campaigns <threshold rows don't split
    smart_slide_estimate = 0
    for deck_info in deck_campaign_counts:
        slides_for_deck = 0
        current_slide_capacity = threshold

        for campaign in deck_info['campaigns']:
            campaign_rows = campaign['size'] + 1  # +1 for MONTHLY TOTAL

            if campaign_rows <= threshold:
                # Small campaign - check if fits on current slide
                if campaign_rows <= current_slide_capacity:
                    # Fits on current slide
                    current_slide_capacity -= campaign_rows
                else:
                    # Doesn't fit - start fresh slide
                    slides_for_deck += 1
                    current_slide_capacity = threshold - campaign_rows
            else:
                # Large campaign (>threshold rows) - always starts fresh, then splits
                if current_slide_capacity < threshold:
                    slides_for_deck += 1  # Finalize current slide

                # Start campaign on fresh slide and split
                campaign_slides = (campaign_rows + threshold - 1) // threshold
                slides_for_deck += campaign_slides
                current_slide_capacity = threshold - (campaign_rows % threshold)
                if current_slide_capacity == threshold:
                    current_slide_capacity = 0

        # Finalize last slide if has content
        if current_slide_capacity < threshold:
            slides_for_deck += 1

        smart_slide_estimate += slides_for_deck

    slide_increase = smart_slide_estimate - current_slide_estimate
    slide_increase_pct = (slide_increase / current_slide_estimate * 100) if current_slide_estimate > 0 else 0

    print(f"Current approach (sequential fill to {threshold} rows):")
    print(f"  Estimated total slides: {current_slide_estimate}")
    print()
    print(f"Smart pagination approach (prevent splits for campaigns <={threshold} rows):")
    print(f"  Estimated total slides: {smart_slide_estimate}")
    print(f"  Increase: +{slide_increase} slides ({slide_increase_pct:+.1f}%)")
    print()
    print(f"Trade-off: {slide_increase_pct:+.1f}% more slides for better readability")
    print()

    return {
        'threshold': threshold,
        'total_campaigns': total_campaigns,
        'campaigns_over_threshold': campaigns_over_threshold,
        'pct_over_threshold': pct_over_threshold,
        'avg_campaign_size': avg_campaign_size,
        'current_slide_estimate': current_slide_estimate,
        'smart_slide_estimate': smart_slide_estimate,
        'slide_increase': slide_increase,
        'slide_increase_pct': slide_increase_pct
    }

if __name__ == '__main__':
    excel_path = Path('template/BulkPlanData_2025_10_14.xlsx')

    if not excel_path.exists():
        print(f"Error: Excel file not found at {excel_path}")
        sys.exit(1)

    # Analyze both thresholds
    print("\n" + "=" * 80)
    print("COMPARISON: 32-ROW vs 40-ROW THRESHOLD")
    print("=" * 80)
    print("\n")

    results_32 = analyze_campaign_sizes(excel_path, threshold=32)
    print("\n\n")
    results_40 = analyze_campaign_sizes(excel_path, threshold=40)

    # Comparison summary
    print("\n" + "=" * 80)
    print("THRESHOLD COMPARISON SUMMARY")
    print("=" * 80)
    print()
    print(f"{'Metric':<40} {'32 rows':<15} {'40 rows':<15}")
    print("-" * 80)
    print(f"{'Campaigns over threshold':<40} {results_32['campaigns_over_threshold']:<15} {results_40['campaigns_over_threshold']:<15}")
    print(f"{'% over threshold':<40} {results_32['pct_over_threshold']:.1f}%{'':<10} {results_40['pct_over_threshold']:.1f}%{'':<10}")
    print()
    print(f"{'Current approach (sequential) slides:':<40} {results_32['current_slide_estimate']:<15} {results_40['current_slide_estimate']:<15}")
    print(f"{'Smart pagination slides:':<40} {results_32['smart_slide_estimate']:<15} {results_40['smart_slide_estimate']:<15}")
    print(f"{'Slide increase:':<40} +{results_32['slide_increase']:<14} +{results_40['slide_increase']:<14}")
    print(f"{'Slide increase %:':<40} {results_32['slide_increase_pct']:+.1f}%{'':<10} {results_40['slide_increase_pct']:+.1f}%{'':<10}")
    print()
    print(f"Benefit of 40-row threshold: {results_32['slide_increase_pct'] - results_40['slide_increase_pct']:.1f} percentage points less increase")
    print()
