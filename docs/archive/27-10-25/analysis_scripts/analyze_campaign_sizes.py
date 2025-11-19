#!/usr/bin/env python3
"""
Campaign Size Distribution Analysis

Analyzes BulkPlanData Excel to determine:
- Row count distribution per campaign
- Percentage of campaigns exceeding 32-row slide limit
- Average campaign size by market
- Maximum campaign size
- Impact of smart pagination on slide count
"""

import pandas as pd
import sys
from pathlib import Path
from collections import Counter, defaultdict

def analyze_campaign_sizes(excel_path):
    """Analyze campaign size distribution from Excel data."""

    print(f"Reading Excel data from: {excel_path}")
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
    market_campaign_counts = defaultdict(list)
    deck_campaign_counts = []
    max_campaign_info = None
    max_campaign_size = 0

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
                # Count rows (each row is a media line item or monthly breakdown)
                # In presentation: each campaign has multiple media rows + MONTHLY TOTAL
                campaign_size = len(campaign_df)

                all_campaign_sizes.append(campaign_size)
                market_campaign_counts[market].append(campaign_size)
                deck_campaigns.append({
                    'name': campaign_name,
                    'size': campaign_size
                })

                if campaign_size > max_campaign_size:
                    max_campaign_size = campaign_size
                    max_campaign_info = {
                        'market': market,
                        'brand': brand,
                        'year': year,
                        'campaign': campaign_name,
                        'size': campaign_size
                    }

            deck_campaign_counts.append({
                'deck_id': deck_id,
                'campaign_count': len(deck_campaigns),
                'campaigns': deck_campaigns
            })
        else:
            # No campaign grouping, entire deck is one unit
            deck_size = len(group_df)
            all_campaign_sizes.append(deck_size)
            market_campaign_counts[market].append(deck_size)
            deck_campaign_counts.append({
                'deck_id': deck_id,
                'campaign_count': 1,
                'campaigns': [{'name': 'Entire Deck', 'size': deck_size}]
            })

            if deck_size > max_campaign_size:
                max_campaign_size = deck_size
                max_campaign_info = {
                    'market': market,
                    'brand': brand,
                    'year': year,
                    'campaign': 'Entire Deck',
                    'size': deck_size
                }

    # Calculate statistics
    total_campaigns = len(all_campaign_sizes)
    campaigns_over_32 = sum(1 for size in all_campaign_sizes if size > 32)
    pct_over_32 = (campaigns_over_32 / total_campaigns * 100) if total_campaigns > 0 else 0

    avg_campaign_size = sum(all_campaign_sizes) / total_campaigns if total_campaigns > 0 else 0

    # Distribution
    size_distribution = Counter(all_campaign_sizes)

    # Market averages
    market_averages = {
        market: sum(sizes) / len(sizes) if sizes else 0
        for market, sizes in market_campaign_counts.items()
    }

    # Print results
    print("=" * 80)
    print("CAMPAIGN SIZE DISTRIBUTION ANALYSIS")
    print("=" * 80)
    print()

    print(f"Total campaigns analyzed: {total_campaigns}")
    print(f"Campaigns exceeding 32 rows: {campaigns_over_32} ({pct_over_32:.1f}%)")
    print(f"Average campaign size: {avg_campaign_size:.1f} rows")
    print(f"Maximum campaign size: {max_campaign_size} rows")
    print()

    if max_campaign_info:
        print("Largest campaign:")
        print(f"  Market: {max_campaign_info['market']}")
        print(f"  Brand: {max_campaign_info['brand']}")
        print(f"  Year: {max_campaign_info['year']}")
        print(f"  Campaign: {max_campaign_info['campaign']}")
        print(f"  Size: {max_campaign_info['size']} rows")
        print()

    print("Size distribution:")
    for size in sorted(size_distribution.keys()):
        count = size_distribution[size]
        pct = count / total_campaigns * 100
        bar = '#' * int(pct / 2)
        print(f"  {size:3d} rows: {count:4d} campaigns ({pct:5.1f}%) {bar}")
    print()

    print("Average campaign size by market:")
    for market in sorted(market_averages.keys()):
        avg = market_averages[market]
        count = len(market_campaign_counts[market])
        print(f"  {market:30s}: {avg:6.1f} rows (n={count})")
    print()

    # Estimate slide count impact
    print("=" * 80)
    print("SMART PAGINATION IMPACT ESTIMATE")
    print("=" * 80)
    print()

    # Current (sequential fill to 32 rows): campaigns can split
    current_slide_estimate = 0
    for deck_info in deck_campaign_counts:
        total_rows = sum(c['size'] for c in deck_info['campaigns'])
        # Add 1 row per campaign for MONTHLY TOTAL
        total_rows += deck_info['campaign_count']
        # Slides = ceil(total_rows / 32)
        slides_needed = (total_rows + 31) // 32
        current_slide_estimate += slides_needed

    # Smart pagination: campaigns <32 rows don't split
    smart_slide_estimate = 0
    for deck_info in deck_campaign_counts:
        slides_for_deck = 0
        current_slide_capacity = 32

        for campaign in deck_info['campaigns']:
            campaign_rows = campaign['size'] + 1  # +1 for MONTHLY TOTAL

            if campaign_rows <= 32:
                # Small campaign - check if fits on current slide
                if campaign_rows <= current_slide_capacity:
                    # Fits on current slide
                    current_slide_capacity -= campaign_rows
                else:
                    # Doesn't fit - start fresh slide
                    slides_for_deck += 1
                    current_slide_capacity = 32 - campaign_rows
            else:
                # Large campaign (>32 rows) - always starts fresh, then splits
                if current_slide_capacity < 32:
                    slides_for_deck += 1  # Finalize current slide

                # Start campaign on fresh slide and split
                campaign_slides = (campaign_rows + 31) // 32
                slides_for_deck += campaign_slides
                current_slide_capacity = 32 - (campaign_rows % 32)
                if current_slide_capacity == 32:
                    current_slide_capacity = 0

        # Finalize last slide if has content
        if current_slide_capacity < 32:
            slides_for_deck += 1

        smart_slide_estimate += slides_for_deck

    slide_increase = smart_slide_estimate - current_slide_estimate
    slide_increase_pct = (slide_increase / current_slide_estimate * 100) if current_slide_estimate > 0 else 0

    print(f"Current approach (sequential fill):")
    print(f"  Estimated total slides: {current_slide_estimate}")
    print()
    print(f"Smart pagination approach:")
    print(f"  Estimated total slides: {smart_slide_estimate}")
    print(f"  Increase: +{slide_increase} slides ({slide_increase_pct:+.1f}%)")
    print()
    print(f"Trade-off: {slide_increase_pct:+.1f}% more slides for better readability")
    print()

    return {
        'total_campaigns': total_campaigns,
        'campaigns_over_32': campaigns_over_32,
        'pct_over_32': pct_over_32,
        'avg_campaign_size': avg_campaign_size,
        'max_campaign_size': max_campaign_size,
        'max_campaign_info': max_campaign_info,
        'size_distribution': dict(size_distribution),
        'market_averages': dict(market_averages),
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

    results = analyze_campaign_sizes(excel_path)

    print("=" * 80)
    print("Analysis complete!")
    print("=" * 80)
