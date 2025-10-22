from pathlib import Path

import pytest

from tools.validate_structure import DEFAULT_CONTRACT_PATH, validate_presentation


@pytest.fixture(scope="module")
def generated_deck() -> Path:
    return Path("output/presentations/run_20251020_124516/GeneratedDeck_Task11_fixed.pptx")


@pytest.fixture(scope="module")
def contract_path() -> Path:
    return DEFAULT_CONTRACT_PATH


@pytest.fixture(scope="module")
def excel_path() -> Path:
    return Path("template/BulkPlanData_2025_10_14.xlsx")


def test_structural_validator_flags_known_gaps(generated_deck, contract_path, excel_path):
    issues = validate_presentation(generated_deck, contract_path, excel_path)
    messages = [issue.message for issue in issues]

    assert issues, "Expected structural issues in legacy deck."
    assert any("Grand total row not found" in message for message in messages), "Expected missing grand total issue."


def test_structural_validator_keeps_header_intact(generated_deck, contract_path, excel_path):
    issues = validate_presentation(generated_deck, contract_path, excel_path)
    messages = [issue.message for issue in issues]
    assert all("Header row mismatch" not in message for message in messages)
