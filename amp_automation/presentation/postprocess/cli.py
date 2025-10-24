"""
CLI entry point for post-processing operations.

This module provides a command-line interface that PowerShell scripts can call
to perform bulk table operations using python-pptx instead of slow COM automation.

Usage from PowerShell:
    python -m amp_automation.presentation.postprocess.cli `
        --presentation-path "path/to/deck.pptx" `
        --operations normalize,reset-spans,merge-campaign,merge-monthly,merge-summary `
        --slide-filter 2,3,4 `
        --verbose
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Optional

from pptx import Presentation

from . import (
    normalize_table_layout,
    apply_blank_cell_formatting,
    reset_primary_column_spans,
    reset_column_group,
    merge_campaign_cells,
    merge_monthly_total_cells,
    merge_summary_cells,
)

logger = logging.getLogger(__name__)


class PostProcessorCLI:
    """CLI handler for presentation post-processing operations."""

    OPERATIONS = {
        "normalize": "Normalize table layout and cell formatting",
        "reset-spans": "Reset column spans in primary columns",
        "merge-campaign": "Merge campaign cells vertically",
        "merge-monthly": "Merge monthly total cells horizontally",
        "merge-summary": "Merge summary cells (GRAND TOTAL, CARRIED FORWARD)",
    }

    def __init__(self, presentation_path: Path, slide_filter: Optional[List[int]] = None):
        self.presentation_path = presentation_path
        self.slide_filter = slide_filter
        self.prs = None

    def load_presentation(self):
        """Load the presentation file."""
        if not self.presentation_path.exists():
            raise FileNotFoundError(f"Presentation not found: {self.presentation_path}")

        logger.info(f"Loading presentation: {self.presentation_path}")
        self.prs = Presentation(str(self.presentation_path))
        logger.info(f"Loaded {len(self.prs.slides)} slides")

    def save_presentation(self):
        """Save the presentation file."""
        logger.info(f"Saving presentation: {self.presentation_path}")
        self.prs.save(str(self.presentation_path))
        logger.info("Presentation saved successfully")

    def run_operation(self, operation: str, slide_idx: int, table) -> bool:
        """
        Run a single operation on a table.

        Args:
            operation: Operation name (e.g., "normalize", "merge-campaign")
            slide_idx: Slide index (1-based)
            table: python-pptx table object

        Returns:
            True if operation succeeded, False otherwise
        """
        try:
            if operation == "normalize":
                normalize_table_layout(table)
                apply_blank_cell_formatting(table)
            elif operation == "reset-spans":
                reset_primary_column_spans(table, max_cols=3)
                reset_column_group(table, max_cols=3)
            elif operation == "merge-campaign":
                merge_campaign_cells(table)
            elif operation == "merge-monthly":
                merge_monthly_total_cells(table)
            elif operation == "merge-summary":
                merge_summary_cells(table)
            else:
                logger.warning(f"Unknown operation: {operation}")
                return False

            return True
        except Exception as e:
            logger.error(f"Slide {slide_idx} - Operation '{operation}' failed: {e}")
            return False

    def process(self, operations: List[str]) -> int:
        """
        Process all slides with the specified operations.

        Args:
            operations: List of operation names to run

        Returns:
            Exit code (0 = success, 1 = error)
        """
        self.load_presentation()

        total_operations = 0
        failed_operations = 0

        for slide_idx, slide in enumerate(self.prs.slides, start=1):
            # Apply slide filter if specified
            if self.slide_filter and slide_idx not in self.slide_filter:
                continue

            logger.info(f"Processing slide {slide_idx}")

            # Find the main table (largest table by cell count)
            tables = [shape for shape in slide.shapes if shape.has_table]
            if not tables:
                logger.debug(f"Slide {slide_idx} - No tables found")
                continue

            # Select largest table
            main_table = max(tables, key=lambda s: s.table.rows.__len__() * s.table.columns.__len__())
            table = main_table.table

            row_count = len(table.rows)
            col_count = len(table.columns)
            logger.debug(f"Slide {slide_idx} - Processing table: {row_count} rows × {col_count} columns")

            # Run each operation
            for operation in operations:
                total_operations += 1
                logger.debug(f"Slide {slide_idx} - Running: {operation}")

                if not self.run_operation(operation, slide_idx, table):
                    failed_operations += 1

        self.save_presentation()

        logger.info(f"Completed: {total_operations} operations, {failed_operations} failed")

        return 1 if failed_operations > 0 else 0


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Post-process PowerPoint presentations with bulk table operations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Available operations:
{chr(10).join(f"  {op:15} - {desc}" for op, desc in PostProcessorCLI.OPERATIONS.items())}

Examples:
  # Normalize all slides
  python -m amp_automation.presentation.postprocess.cli \\
      --presentation-path deck.pptx --operations normalize

  # Full post-processing on specific slides
  python -m amp_automation.presentation.postprocess.cli \\
      --presentation-path deck.pptx \\
      --operations normalize,reset-spans,merge-campaign,merge-monthly,merge-summary \\
      --slide-filter 2,3,4 --verbose
""",
    )

    parser.add_argument(
        "--presentation-path",
        type=Path,
        required=True,
        help="Path to the PowerPoint presentation (.pptx)",
    )

    parser.add_argument(
        "--operations",
        type=str,
        required=True,
        help=f"Comma-separated list of operations: {', '.join(PostProcessorCLI.OPERATIONS.keys())}",
    )

    parser.add_argument(
        "--slide-filter",
        type=str,
        help="Comma-separated list of slide numbers to process (1-based). If omitted, processes all slides.",
    )

    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%H:%M:%S",
    )

    # Parse operations
    operations = [op.strip() for op in args.operations.split(",")]
    invalid_ops = [op for op in operations if op not in PostProcessorCLI.OPERATIONS]
    if invalid_ops:
        logger.error(f"Invalid operations: {', '.join(invalid_ops)}")
        logger.error(f"Valid operations: {', '.join(PostProcessorCLI.OPERATIONS.keys())}")
        return 1

    # Parse slide filter
    slide_filter = None
    if args.slide_filter:
        try:
            slide_filter = [int(s.strip()) for s in args.slide_filter.split(",")]
        except ValueError:
            logger.error(f"Invalid slide filter: {args.slide_filter}")
            return 1

    # Run processor
    try:
        processor = PostProcessorCLI(args.presentation_path, slide_filter)
        return processor.process(operations)
    except Exception as e:
        logger.exception(f"Fatal error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
