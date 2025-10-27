"""Comprehensive data validation for generated AMP decks.

Runs all data validation checks and generates a unified report covering:
- Structural validation (layout, shapes, contracts)
- Data accuracy (numerical values match source)
- Data format (proper formatting of values)
- Data completeness (required data present)
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import List

from amp_automation.config.loader import Config, load_master_config
from amp_automation.validation.data_accuracy import validate_data_accuracy
from amp_automation.validation.data_completeness import validate_data_completeness
from amp_automation.validation.data_format import validate_data_format
from amp_automation.validation.reconciliation import (
    generate_reconciliation_report,
    write_reconciliation_report,
)
from amp_automation.validation.utils import (
    ValidationResult,
    summarize_validation_results,
    write_validation_report,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def run_all_validations(
    ppt_path: Path,
    excel_path: Path | None = None,
    output_dir: Path | None = None,
    config: Config | None = None,
) -> dict:
    """Run all validation checks and return aggregated results."""

    if config is None:
        config = load_master_config()

    if output_dir is None:
        output_dir = PROJECT_ROOT / "output" / "validation"

    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_summary = {
        "timestamp": timestamp,
        "ppt_path": str(ppt_path),
        "excel_path": str(excel_path) if excel_path else None,
        "validations": {},
        "overall_status": "PASS",
    }

    print(f"Running comprehensive data validation on: {ppt_path}")
    print(f"Output directory: {output_dir}\n")

    # 1. Data Accuracy Validation
    print("1. Validating data accuracy...")
    try:
        if excel_path:
            accuracy_results = validate_data_accuracy(ppt_path, excel_path, config)
            accuracy_summary = summarize_validation_results(accuracy_results)

            accuracy_report_path = output_dir / f"validation_accuracy_{timestamp}.csv"
            write_validation_report(
                [r for res in accuracy_results for r in [res]],
                accuracy_report_path
            )

            results_summary["validations"]["data_accuracy"] = {
                "status": "PASS" if accuracy_summary["passed"] else "FAIL",
                "total_issues": accuracy_summary["total_issues"],
                "errors": accuracy_summary["error_count"],
                "warnings": accuracy_summary["warning_count"],
                "report": str(accuracy_report_path),
            }
            print(f"   [OK] Accuracy check: {accuracy_summary['total_issues']} issues found")
        else:
            results_summary["validations"]["data_accuracy"] = {
                "status": "SKIPPED",
                "reason": "No Excel file provided",
            }
            print("   [--] Accuracy check skipped (no Excel file)")
    except Exception as e:
        results_summary["validations"]["data_accuracy"] = {
            "status": "ERROR",
            "error": str(e),
        }
        print(f"   [XX] Accuracy check failed: {e}")
        results_summary["overall_status"] = "FAIL"

    # 2. Data Format Validation
    print("2. Validating data format...")
    try:
        format_results = validate_data_format(ppt_path)
        format_summary = summarize_validation_results(format_results)

        format_report_path = output_dir / f"validation_format_{timestamp}.csv"
        write_validation_report(
            [r for res in format_results for r in [res]],
            format_report_path
        )

        results_summary["validations"]["data_format"] = {
            "status": "PASS" if format_summary["passed"] else "FAIL",
            "total_issues": format_summary["total_issues"],
            "errors": format_summary["error_count"],
            "warnings": format_summary["warning_count"],
            "report": str(format_report_path),
        }
        print(f"   [OK] Format check: {format_summary['total_issues']} issues found")
    except Exception as e:
        import traceback
        results_summary["validations"]["data_format"] = {
            "status": "ERROR",
            "error": str(e),
        }
        print(f"   [XX] Format check failed: {e}")
        print(f"   Traceback:\n{traceback.format_exc()}")
        results_summary["overall_status"] = "FAIL"

    # 3. Data Completeness Validation
    print("3. Validating data completeness...")
    try:
        completeness_results = validate_data_completeness(ppt_path)
        completeness_summary = summarize_validation_results(completeness_results)

        completeness_report_path = output_dir / f"validation_completeness_{timestamp}.csv"
        write_validation_report(
            [r for res in completeness_results for r in [res]],
            completeness_report_path
        )

        results_summary["validations"]["data_completeness"] = {
            "status": "PASS" if completeness_summary["passed"] else "FAIL",
            "total_issues": completeness_summary["total_issues"],
            "errors": completeness_summary["error_count"],
            "warnings": completeness_summary["warning_count"],
            "report": str(completeness_report_path),
        }
        print(f"   [OK] Completeness check: {completeness_summary['total_issues']} issues found")
    except Exception as e:
        results_summary["validations"]["data_completeness"] = {
            "status": "ERROR",
            "error": str(e),
        }
        print(f"   [XX] Completeness check failed: {e}")
        results_summary["overall_status"] = "FAIL"

    # 4. Reconciliation Validation (if Excel provided)
    print("4. Validating reconciliation (summary tiles vs source)...")
    try:
        if excel_path:
            recon_results = generate_reconciliation_report(ppt_path, excel_path, config)
            recon_passed = all(r.passed for r in recon_results)

            recon_report_path = output_dir / f"validation_reconciliation_{timestamp}.csv"
            write_reconciliation_report(recon_results, recon_report_path)

            results_summary["validations"]["reconciliation"] = {
                "status": "PASS" if recon_passed else "FAIL",
                "slides_checked": len(recon_results),
                "report": str(recon_report_path),
            }
            print(f"   [OK] Reconciliation check: {len(recon_results)} slides verified")
        else:
            results_summary["validations"]["reconciliation"] = {
                "status": "SKIPPED",
                "reason": "No Excel file provided",
            }
            print("   [--] Reconciliation check skipped (no Excel file)")
    except Exception as e:
        results_summary["validations"]["reconciliation"] = {
            "status": "ERROR",
            "error": str(e),
        }
        print(f"   [XX] Reconciliation check failed: {e}")
        results_summary["overall_status"] = "FAIL"

    # Print summary
    print("\n" + "=" * 70)
    print("VALIDATION SUMMARY")
    print("=" * 70)

    for check_name, check_result in results_summary["validations"].items():
        status = check_result.get("status", "UNKNOWN")
        status_symbol = "[OK]" if status == "PASS" else "[XX]" if status == "FAIL" else "[--]"
        print(f"{status_symbol} {check_name.replace('_', ' ').title():.<40} {status}")

    print("=" * 70)
    print(f"Overall Status: {results_summary['overall_status']}")
    print(f"Output directory: {output_dir}")

    return results_summary


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run comprehensive data validation on generated AMP decks."
    )
    parser.add_argument("presentation", type=Path, help="Path to the generated PPTX deck.")
    parser.add_argument(
        "--excel",
        type=Path,
        default=None,
        help="Raw Excel file for accuracy/reconciliation validation.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output directory for validation reports (default: output/validation).",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help="Path to configuration file.",
    )

    args = parser.parse_args()

    if not args.presentation.is_file():
        print(f"Error: Presentation not found: {args.presentation}", file=sys.stderr)
        return 1

    try:
        config = load_master_config(args.config)
    except Exception as e:
        print(f"Error loading configuration: {e}", file=sys.stderr)
        return 1

    results = run_all_validations(
        args.presentation,
        excel_path=args.excel,
        output_dir=args.output,
        config=config,
    )

    # Exit with error code if validation failed
    if results["overall_status"] != "PASS":
        return 1

    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
