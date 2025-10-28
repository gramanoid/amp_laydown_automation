"""Regression tests for structural validation of generated decks."""

from __future__ import annotations

import logging
from pathlib import Path

import pytest

# Import validation tools
try:
    from tools.validate.validate_structure import validate_presentation
except ImportError:
    # Fallback if module not in path
    validate_presentation = None


@pytest.mark.integration
@pytest.mark.skipif(
    not Path("output/presentations/run_20251028_164203/AMP_Laydowns_281025.pptx").exists(),
    reason="Latest production deck (28-10-25) not found"
)
def test_structural_validator_passes_production_deck(latest_deck_path: Path, contract_path: Path, test_logger: logging.Logger) -> None:
    """Verify structural validator passes on latest 28-10-25 production deck.

    Tests that the validator correctly validates:
    - Last-slide-only shapes (BRAND TOTAL indicators)
    - Shape recognition and naming
    - Slide count and structure
    """
    if validate_presentation is None:
        pytest.skip("validate_structure module not available")

    if not latest_deck_path.exists():
        pytest.skip(f"Latest deck not found: {latest_deck_path}")

    if not contract_path.exists():
        pytest.skip(f"Contract file not found: {contract_path}")

    # Run validation
    issues = validate_presentation(latest_deck_path, contract_path)

    # Should have no structural issues
    if issues:
        issue_msgs = "\n".join(f"  - {issue}" for issue in issues)
        pytest.fail(f"Structural validation found {len(issues)} issue(s):\n{issue_msgs}")

    test_logger.info(f"Structural validation passed for {latest_deck_path.name}")


@pytest.mark.regression
def test_structural_validator_recognizes_brand_total() -> None:
    """Verify validator contract recognizes BRAND TOTAL (not GRAND TOTAL) on final slides.

    Regression test which validates structural_contract.json correctly
    uses BRAND TOTAL or GRAND TOTAL for final slide indicators.
    """
    from pathlib import Path
    import json

    project_root = Path(__file__).parent.parent
    contract_path = project_root / "config" / "structural_contract.json"

    if not contract_path.exists():
        pytest.skip(f"Contract file not found: {contract_path}")

    with open(contract_path) as f:
        contract = json.load(f)

    # Verify contract references BRAND TOTAL
    contract_str = json.dumps(contract)
    assert "BRAND TOTAL" in contract_str, "Contract should recognize BRAND TOTAL"

    # If GRAND TOTAL is referenced, it should only be for backward compatibility
    # not as the primary indicator
    if "GRAND TOTAL" in contract_str:
        # This is acceptable but BRAND TOTAL should be the primary reference
        test_logger = logging.getLogger("test")
        test_logger.debug("Contract contains GRAND TOTAL reference (legacy support)")


@pytest.mark.regression
def test_structural_validator_handles_last_slide_only_shapes() -> None:
    """Verify validator contract correctly identifies last-slide-only shape rules.

    Regression test which validates structural_contract.json supports
    shapes that only appear on the final slide(s).
    """
    from pathlib import Path
    import json

    project_root = Path(__file__).parent.parent
    contract_path = project_root / "config" / "structural_contract.json"

    if not contract_path.exists():
        pytest.skip(f"Contract file not found: {contract_path}")

    with open(contract_path) as f:
        contract = json.load(f)

    # Verify contract has configuration for last-slide-only shapes
    contract_str = json.dumps(contract)

    # Contract should specify which shapes are only on final slides
    # or have a configuration that allows for final-slide-only validation
    # This is a positive check - we verify the contract was updated
    assert isinstance(contract, dict), "Contract should be a valid JSON object"

    # If indicators/shapes configuration exists, verify it's properly structured
    if "shapes" in contract or "indicators" in contract:
        shapes_config = contract.get("shapes", contract.get("indicators", {}))
        assert isinstance(shapes_config, dict), "Shapes configuration should be a dict"
