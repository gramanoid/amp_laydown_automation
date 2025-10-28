"""Regression tests for structural validation of generated decks."""

from __future__ import annotations

from pathlib import Path

import pytest

from tools.validate_structure import DEFAULT_CONTRACT_PATH, validate_presentation
from conftest import skipif_no_deck


@pytest.fixture(scope="module")
def generated_deck(latest_deck_path: Path) -> Path:
    """Path to latest production deck (27-10-25)."""
    return latest_deck_path


@pytest.fixture(scope="module")
def contract_path() -> Path:
    return DEFAULT_CONTRACT_PATH


@pytest.fixture(scope="module")
def excel_path() -> Path:
    return Path("template/BulkPlanData_2025_10_14.xlsx")


@pytest.mark.integration
@skipif_no_deck
def test_structural_validator_passes_production_deck(generated_deck, contract_path, excel_path):
    """Verify structural validator passes on latest 27-10-25 production deck."""
    # 27-10-25 fixes: last-slide-only shapes support, BRAND TOTAL recognition
    issues = validate_presentation(generated_deck, contract_path, excel_path)

    # Production deck should pass structural validation
    assert len(issues) == 0, f"Expected no structural issues, found: {[i.message for i in issues]}"


@pytest.mark.regression
@skipif_no_deck
def test_structural_validator_recognizes_brand_total(generated_deck, contract_path, excel_path):
    """Verify validator recognizes BRAND TOTAL (not GRAND TOTAL) on slides."""
    # Commit 6e83fae: Updated contract to use BRAND TOTAL
    issues = validate_presentation(generated_deck, contract_path, excel_path)
    messages = [issue.message for issue in issues]

    # Should not complain about missing GRAND TOTAL (it's BRAND TOTAL now)
    assert not any("GRAND TOTAL" in message for message in messages), \
        "Validator still expects old GRAND TOTAL label"


@pytest.mark.regression
@skipif_no_deck
def test_structural_validator_handles_last_slide_only_shapes(generated_deck, contract_path, excel_path):
    """Verify validator correctly identifies last-slide-only shapes."""
    # Commit 6e83fae: Indicators only validated on final slides
    issues = validate_presentation(generated_deck, contract_path, excel_path)

    # Should not flag missing indicators on continuation slides
    # (they're only required on final slide per market/brand)
    assert not any("QuarterBudget" in message and "continuation" in message.lower()
                   for message in [i.message for i in issues]), \
        "Validator incorrectly flags missing indicators on continuation slides"
