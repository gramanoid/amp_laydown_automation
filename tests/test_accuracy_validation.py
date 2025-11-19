"""Test accuracy validation on generated decks."""

import pytest
from pathlib import Path

from amp_automation.validation.accuracy_validator import validate_deck_accuracy


@pytest.mark.integration
def test_latest_deck_accuracy():
    """Validate accuracy of the latest generated deck."""
    # Find the latest deck
    output_dir = Path("output/presentations")
    if not output_dir.exists():
        pytest.skip("No output directory found")

    # Get all run directories
    run_dirs = [d for d in output_dir.iterdir() if d.is_dir() and d.name.startswith("run_")]
    if not run_dirs:
        pytest.skip("No generated decks found")

    # Get the latest run directory
    latest_run = max(run_dirs, key=lambda d: d.name)

    # Find the .pptx file
    pptx_files = list(latest_run.glob("*.pptx"))
    if not pptx_files:
        pytest.skip(f"No .pptx file found in {latest_run}")

    pptx_path = pptx_files[0]
    print(f"\nValidating: {pptx_path}")

    # Run validation
    report = validate_deck_accuracy(pptx_path)

    # Print report
    print("\n" + report.summary())

    # Assert no errors
    assert report.passed, f"Validation failed with {report.error_count} errors"


@pytest.mark.integration
def test_production_deck_accuracy():
    """Validate accuracy of specific production deck if provided."""
    # This can be used to validate a specific deck
    production_deck = Path("output/presentations/run_20251029_130725/AMP_Laydowns_291025.pptx")

    if not production_deck.exists():
        pytest.skip(f"Production deck not found: {production_deck}")

    print(f"\nValidating production deck: {production_deck}")

    # Run validation
    report = validate_deck_accuracy(production_deck)

    # Print report
    print("\n" + report.summary())

    # Assert no errors
    assert report.passed, f"Validation failed with {report.error_count} errors"
