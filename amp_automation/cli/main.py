"""CLI orchestration for AMP Automation."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence

from amp_automation.config import Config, load_master_config
from amp_automation.utils import configure_logger
from amp_automation.presentation.postprocess.cli import PostProcessorCLI

PROJECT_ROOT = Path(__file__).resolve().parents[2]


@dataclass(slots=True)
class ResolvedPaths:
    """Filesystem locations derived from CLI arguments and configuration."""

    template: Path
    excel: Path
    output_dir: Path
    output_file: Path
    log_dir: Path


def build_parser() -> argparse.ArgumentParser:
    """Construct the argument parser used by the CLI entrypoint."""

    parser = argparse.ArgumentParser(
        description="Generate AMP presentations from Excel data using the configured template.",
    )
    parser.add_argument("--config", help="Path to an alternative master_config.json file.")
    parser.add_argument("--excel", help="Path to the input Excel workbook.")
    parser.add_argument(
        "--template",
        help="Path to the PowerPoint template. Defaults to the configured template location.",
    )
    parser.add_argument(
        "--output",
        help="Optional name for the generated presentation file (placed inside the run directory).",
    )
    parser.add_argument(
        "--output-dir",
        help="Override the output directory base defined in the configuration.",
    )
    parser.add_argument(
        "--log-dir",
        help="Override the log directory base defined in the configuration.",
    )
    parser.add_argument(
        "--list-templates",
        action="store_true",
        help="List available .pptx templates from configured directories and exit.",
    )
    parser.add_argument(
        "--reconcile",
        action="store_true",
        help="Run Excel vs PPT reconciliation checks after generation and emit a CSV report.",
    )
    parser.add_argument(
        "--reconciliation-report",
        help="Optional output path (CSV) for reconciliation results. Defaults to the run directory.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Run the CLI, returning a process exit code."""

    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        config = load_master_config(args.config)
    except Exception as exc:  # pragma: no cover - configuration errors are user facing
        parser.error(str(exc))

    template_search_dirs = _collect_template_dirs(config)

    if args.list_templates:
        _print_available_templates(template_search_dirs)
        return 0

    if not args.excel:
        parser.error("--excel is required unless --list-templates is used")

    try:
        paths = _resolve_paths(args, config, template_search_dirs)
    except FileNotFoundError as exc:
        parser.error(str(exc))

    paths.output_dir.mkdir(parents=True, exist_ok=True)
    paths.output_file.parent.mkdir(parents=True, exist_ok=True)

    logging_config = {}
    try:
        logging_config = config.section("logging")
    except (KeyError, TypeError):
        logging_config = {}

    default_log_level = str(logging_config.get("level", "INFO"))
    console_output = logging_config.get("console_output", True)
    file_output = logging_config.get("file_output", True)

    logger = configure_logger(
        paths.log_dir,
        default_level=default_log_level,
        console_enabled=bool(console_output),
        file_enabled=bool(file_output),
    )

    from amp_automation.presentation import assembly as presentation_assembly

    presentation_assembly.logger = logger
    presentation_assembly.configure(config)

    logger.info("Starting presentation build")
    logger.debug("Excel path: %s", paths.excel)
    logger.debug("Template path: %s", paths.template)
    logger.debug("Output file: %s", paths.output_file)

    success = presentation_assembly.build_presentation(
        template_path=str(paths.template),
        excel_path=str(paths.excel),
        output_path=str(paths.output_file),
    )

    if success:
        logger.info("Presentation generated successfully: %s", paths.output_file)

        # AUTOMATIC POST-PROCESSING: Run complete workflow on generated deck
        # This applies cell merging, font normalization, and styling
        logger.info("Running automatic post-processing...")
        try:
            postprocessor = PostProcessorCLI(paths.output_file, slide_filter=None)
            exit_code = postprocessor.process(PostProcessorCLI.POSTPROCESS_ALL_WORKFLOW)

            if exit_code == 0:
                logger.info("Post-processing completed successfully")
            else:
                logger.warning("Post-processing completed with warnings/errors")
        except Exception as e:
            logger.error(f"Post-processing failed: {e}")
            logger.warning("Presentation generated but post-processing incomplete")

        print(paths.output_file)
        _run_reconciliation_if_requested(args, paths, config, logger)
        return 0

    logger.error("Presentation generation failed")
    return 1


def _print_available_templates(search_dirs: Iterable[Path]) -> None:
    """Emit a list of available templates relative to the project root."""

    candidates: list[Path] = []
    for directory in search_dirs:
        if not directory.is_dir():
            continue
        candidates.extend(sorted(directory.glob("*.pptx")))

    if not candidates:
        print("No template files found in configured directories.")
        return

    for path in candidates:
        try:
            print(path.relative_to(PROJECT_ROOT))
        except ValueError:
            print(path)


def _resolve_paths(
    args: argparse.Namespace,
    config: Config,
    template_dirs: Sequence[Path],
) -> ResolvedPaths:
    """Resolve all filesystem paths needed for a single CLI invocation."""

    template_path = _resolve_template(args.template, config, template_dirs)
    excel_path = _resolve_existing_file("Excel", args.excel)

    output_dir, output_file = _resolve_output_locations(args, config)
    log_dir = _resolve_log_directory(args, config, output_dir)

    return ResolvedPaths(
        template=template_path,
        excel=excel_path,
        output_dir=output_dir,
        output_file=output_file,
        log_dir=log_dir,
    )


def _resolve_template(
    template_arg: str | None,
    config: Config,
    template_dirs: Sequence[Path],
) -> Path:
    """Locate the template file using CLI arguments, config, and template directories."""

    if template_arg:
        return _resolve_existing_file("Template", template_arg)

    template_section = config.section("template")
    location = template_section.get("location") or template_section.get("current")
    if location:
        try:
            return _resolve_existing_file("Template", location)
        except FileNotFoundError:
            pass

    for directory in template_dirs:
        candidate = directory / (template_section.get("current") or "")
        if candidate.is_file():
            return candidate.resolve()

    raise FileNotFoundError("Template file not found in configured locations")


def _resolve_output_locations(
    args: argparse.Namespace,
    config: Config,
) -> tuple[Path, Path]:
    """Determine the run directory and presentation output file path."""

    paths_section = config.section("paths")
    output_section = paths_section.get("output", {})

    # Get output filename config from top-level config
    output_config = config.get("output", {})
    filename_config = output_config.get("filename", {})

    base = Path(args.output_dir) if args.output_dir else Path(output_section.get("presentations") or output_section.get("base") or "output")
    if not base.is_absolute():
        base = PROJECT_ROOT / base

    timestamp_format = output_section.get("timestamp_format", "%Y%m%d_%H%M%S")
    folder_pattern = output_section.get("folder_pattern", "run_{timestamp}")
    create_timestamped = output_section.get("create_timestamped_folders", True)

    # Explicitly use local system time (not UTC)
    from datetime import timezone
    timestamp = datetime.now().astimezone().strftime(timestamp_format)
    run_dir = base / folder_pattern.format(timestamp=timestamp) if create_timestamped else base

    # Use filename pattern from config (default: AMP_Laydowns_{timestamp}.pptx)
    filename_pattern = filename_config.get("pattern", "AMP_Laydowns_{timestamp}.pptx")
    filename_timestamp_format = filename_config.get("timestamp_format", "%d%m%y")
    filename_timestamp = datetime.now().astimezone().strftime(filename_timestamp_format)

    if args.output:
        # Check if args.output looks like a file (has .pptx extension) or directory
        candidate_path = Path(args.output)
        if candidate_path.suffix.lower() == ".pptx":
            # It's a filename - use it
            candidate_name = candidate_path.name
            if "{timestamp}" in candidate_name:
                candidate_name = candidate_name.format(timestamp=filename_timestamp)
            output_name = candidate_name
        else:
            # It's a directory or base path - use the filename pattern
            output_name = filename_pattern.format(timestamp=filename_timestamp)
    else:
        output_name = filename_pattern.format(timestamp=filename_timestamp)

    output_path = run_dir / output_name
    if output_path.suffix.lower() != ".pptx":
        output_path = output_path.with_suffix(".pptx")

    return run_dir.resolve(), output_path.resolve()


def _resolve_log_directory(
    args: argparse.Namespace,
    config: Config,
    output_dir: Path,
) -> Path:
    """Resolve the directory where log files should be written."""

    if args.log_dir:
        log_base = Path(args.log_dir)
    else:
        paths_section = config.section("paths")
        logs_section = paths_section.get("logs", {})
        log_base = Path(logs_section.get("production") or logs_section.get("base") or "logs")

    if not log_base.is_absolute():
        log_base = PROJECT_ROOT / log_base

    if log_base == output_dir:
        return log_base

    timestamp_folder = output_dir.name
    return (log_base / timestamp_folder).resolve()


def _resolve_existing_file(label: str, raw_path: str) -> Path:
    """Return the absolute path to an existing file, raising if it is missing."""

    candidate = Path(raw_path).expanduser()
    if not candidate.is_absolute():
        candidate = PROJECT_ROOT / candidate

    if candidate.is_file():
        return candidate.resolve()

    raise FileNotFoundError(f"{label} file not found: {candidate}")


def _collect_template_dirs(config: Config) -> list[Path]:
    """Collect template search directories based on configuration defaults."""

    paths_section = config.section("paths")
    template_dirs: list[Path] = []

    templates_path = paths_section.get("input", {}).get("templates")
    if templates_path:
        template_dirs.append(_to_absolute_path(templates_path))

    template_section = config.section("template")
    template_location = template_section.get("location")
    if template_location:
        template_dirs.append(_to_absolute_path(Path(template_location).parent))

    template_dirs.append(PROJECT_ROOT / "template")
    template_dirs.append(PROJECT_ROOT)
    template_dirs.append(PROJECT_ROOT.parent)

    # Deduplicate while preserving order
    seen: set[Path] = set()
    unique_dirs: list[Path] = []
    for directory in template_dirs:
        directory = directory.resolve()
        if directory not in seen:
            seen.add(directory)
            unique_dirs.append(directory)

    return unique_dirs


def _to_absolute_path(path_like: str | Path) -> Path:
    """Return the project-root-relative absolute version of *path_like*."""

    path = Path(path_like)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path.resolve()


def _run_reconciliation_if_requested(
    args: argparse.Namespace,
    paths: ResolvedPaths,
    config: Config,
    logger,
) -> None:
    if not getattr(args, "reconcile", False):
        return

    from amp_automation.validation import reconciliation as reconciliation_module

    if args.reconciliation_report:
        report_path = Path(args.reconciliation_report)
        if not report_path.is_absolute():
            report_path = paths.output_dir / report_path
    else:
        report_path = paths.output_dir / "reconciliation_summary.csv"

    try:
        results = reconciliation_module.generate_reconciliation_report(
            paths.output_file,
            paths.excel,
            config,
            logger=logger,
        )
        reconciliation_module.write_reconciliation_report(results, report_path)
        if not results:
            logger.info("Reconciliation produced no data-driven slides; report written to %s", report_path)
            return

        failing = [item for item in results if not item.passed]
        if failing:
            logger.warning(
                "Reconciliation detected mismatches on %s slide(s); review %s",
                len(failing),
                report_path,
            )
        else:
            logger.info("Reconciliation passed for all summary tiles; report saved to %s", report_path)
    except Exception as exc:
        logger.error("Reconciliation failed: %s", exc)


if __name__ == "__main__":  # pragma: no cover - manual execution entrypoint
    sys.exit(main())
