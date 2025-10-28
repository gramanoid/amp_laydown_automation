"""Regression tests for reconciliation validator (EC-004, EC-005, EC-007)."""

from __future__ import annotations

import pandas as pd
import pytest

from amp_automation.validation.reconciliation import (
    _normalize_market_name,
    _normalize_brand_name,
    _parse_title_tokens,
    MARKET_CODE_MAP,
)


@pytest.mark.regression
def test_ec004_market_normalization_case_insensitive(market_case_variations):
    """Verify market names normalized case-insensitively (EC-004)."""
    # Create test DataFrame
    df = pd.DataFrame({
        "Country": ["SOUTH AFRICA", "Egypt", "Morocco", "KSA"],
        "Brand": ["Fanta", "Sprite", "Coca-Cola", "Fanta"],
        "Year": [2025, 2025, 2025, 2025],
    })

    # Test case variations
    for canonical, variations in market_case_variations.items():
        for variant in variations:
            normalized = _normalize_market_name(df, variant)
            # Should match one of the actual countries in DataFrame
            assert normalized in df["Country"].unique(), \
                f"'{variant}' normalized to '{normalized}', not in DataFrame: {df['Country'].unique()}"


@pytest.mark.regression
def test_ec005_market_code_mapping(market_case_variations):
    """Verify market code mapping works (MOR→MOROCCO, EC-005)."""
    # Create test DataFrame with full names (post-mapping)
    df = pd.DataFrame({
        "Country": ["SOUTH AFRICA", "EGYPT", "MOROCCO", "KSA"],
        "Brand": ["Fanta", "Sprite", "Coca-Cola", "Fanta"],
        "Year": [2025, 2025, 2025, 2025],
    })

    # Test market code translations
    test_cases = [
        ("MOR", "MOROCCO"),  # Code → Display name
        ("EGYPT", "EGYPT"),  # Identity mapping
        ("KSA", "KSA"),      # Identity mapping
    ]

    for input_val, expected_market in test_cases:
        # Check MARKET_CODE_MAP contains the mapping
        assert expected_market in MARKET_CODE_MAP.values(), \
            f"Expected market '{expected_market}' not in MARKET_CODE_MAP"

        normalized = _normalize_market_name(df, input_val)
        assert normalized in df["Country"].unique(), \
            f"Market code '{input_val}' not normalized correctly. Got '{normalized}'"


@pytest.mark.regression
def test_ec004_brand_normalization_within_market(brand_case_variations):
    """Verify brand names normalized case-insensitively within market (EC-004)."""
    # Create test DataFrame with brands per market
    df = pd.DataFrame({
        "Country": ["SOUTH AFRICA", "SOUTH AFRICA", "SOUTH AFRICA", "EGYPT"],
        "Brand": ["Fanta", "Sprite", "Coca-Cola", "Fanta"],
        "Year": [2025, 2025, 2025, 2025],
    })

    # Test brand case variations
    test_cases = [
        ("SOUTH AFRICA", "fanta", "Fanta"),
        ("SOUTH AFRICA", "FANTA", "Fanta"),
        ("SOUTH AFRICA", "sprite", "Sprite"),
        ("SOUTH AFRICA", "SPRITE", "Sprite"),
    ]

    for market, brand_variant, expected_brand in test_cases:
        normalized = _normalize_brand_name(df, market, brand_variant)
        assert normalized == expected_brand, \
            f"Brand '{brand_variant}' in market '{market}' should normalize to '{expected_brand}', got '{normalized}'"


@pytest.mark.regression
def test_ec010_title_parsing_both_pagination_formats(pagination_formats):
    """Verify both pagination formats recognized in title parsing (EC-010)."""
    test_cases = [
        ("SOUTH AFRICA - Fanta (1 of 3)", ("SOUTH AFRICA", "Fanta")),
        ("SOUTH AFRICA - Fanta (1/3)", ("SOUTH AFRICA", "Fanta")),
        ("SOUTH AFRICA - Fanta", ("SOUTH AFRICA", "Fanta")),
        ("EGYPT - Sprite (2 of 5)", ("EGYPT", "Sprite")),
        ("EGYPT - Sprite (2/5)", ("EGYPT", "Sprite")),
    ]

    for title, expected in test_cases:
        market, brand = _parse_title_tokens(title)
        assert (market, brand) == expected, \
            f"Title '{title}' parsed to ({market}, {brand}), expected {expected}"


@pytest.mark.regression
def test_ec010_title_parsing_edge_cases():
    """Verify title parsing handles edge cases (EC-010)."""
    edge_cases = [
        ("Invalid Title", (None, None)),           # No delimiter
        ("", (None, None)),                        # Empty
        ("ONLY-ONE-PART", (None, None)),          # Missing delimiter
        ("MARKET - BRAND - EXTRA", ("MARKET", "BRAND - EXTRA")),  # Split on first " - "
    ]

    for title, expected in edge_cases:
        market, brand = _parse_title_tokens(title)
        assert (market, brand) == expected, \
            f"Title '{title}' parsed to ({market}, {brand}), expected {expected}"


@pytest.mark.unit
def test_market_code_map_completeness():
    """Verify MARKET_CODE_MAP contains expected markets."""
    expected_markets = [
        "MOROCCO", "SOUTH AFRICA", "KSA", "GINE", "EGYPT",
        "TURKEY", "PAKISTAN", "KENYA", "UGANDA", "NIGERIA",
        "MAURITIUS", "FWA"
    ]

    for market in expected_markets:
        assert market in MARKET_CODE_MAP.values(), \
            f"Expected market '{market}' not in MARKET_CODE_MAP"


@pytest.mark.unit
def test_normalization_functions_return_string():
    """Verify normalization functions return strings (type safety)."""
    df = pd.DataFrame({
        "Country": ["SOUTH AFRICA"],
        "Brand": ["Fanta"],
    })

    market = _normalize_market_name(df, "SOUTH AFRICA")
    assert isinstance(market, str), f"normalize_market_name should return str, got {type(market)}"

    brand = _normalize_brand_name(df, "SOUTH AFRICA", "Fanta")
    assert isinstance(brand, str), f"normalize_brand_name should return str, got {type(brand)}"


@pytest.mark.unit
def test_normalization_fallback_returns_original():
    """Verify normalization falls back to original if no match found."""
    df = pd.DataFrame({
        "Country": ["SOUTH AFRICA"],
        "Brand": ["Fanta"],
    })

    # Non-existent market should return original
    market = _normalize_market_name(df, "NONEXISTENT")
    assert market == "NONEXISTENT", "Should return original if no match found"

    # Non-existent brand should return original
    brand = _normalize_brand_name(df, "SOUTH AFRICA", "NONEXISTENT")
    assert brand == "NONEXISTENT", "Should return original if no match found"
