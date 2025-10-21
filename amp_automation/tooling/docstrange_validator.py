"""DocStrange-powered validation utilities."""

from __future__ import annotations

import json
import logging
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

logger = logging.getLogger("amp_automation.docstrange")


class DocStrangeNotAvailable(RuntimeError):
    """Raised when the docstrange CLI could not be located."""


@dataclass(slots=True)
class DocStrangeDiffResult:
    """Result container for DocStrange-based comparisons."""

    generated_output: Path
    reference_output: Path
    diff_output: Path


def docstrange_available(command: str | None = None) -> bool:
    """Return ``True`` if the docstrange CLI can be executed."""

    cmd = command or "docstrange"
    return shutil.which(cmd) is not None


def compare_presentations(
    pptx_path: str | Path,
    reference_path: str | Path,
    output_dir: str | Path,
    *,
    command: str | None = None,
    output_format: str = "markdown",
) -> DocStrangeDiffResult:
    """Compare two presentations by extracting content via DocStrange.

    Parameters
    ----------
    pptx_path / reference_path:
        The generated presentation and the reference/template deck.
    output_dir:
        Directory that will contain the extracted artifacts and diff.
    command:
        Optional override for the ``docstrange`` executable name.
    output_format:
        DocStrange output format (``markdown`` by default).
    """

    cmd = command or "docstrange"
    if not docstrange_available(cmd):  # pragma: no cover - environment dependent
        raise DocStrangeNotAvailable(
            "docstrange CLI not found on PATH. Install DocStrange or adjust the command."
        )

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    generated_output = output_dir / f"generated.{_extension_for_format(output_format)}"
    reference_output = output_dir / f"reference.{_extension_for_format(output_format)}"
    diff_output = output_dir / "docstrange_diff.txt"

    _run_docstrange(cmd, pptx_path, generated_output, output_format)
    _run_docstrange(cmd, reference_path, reference_output, output_format)

    _write_diff(generated_output, reference_output, diff_output)

    logger.info("DocStrange diff written to %s", diff_output)
    return DocStrangeDiffResult(generated_output, reference_output, diff_output)


def _run_docstrange(command: str, pptx_path: str | Path, output_path: Path, fmt: str) -> None:
    args = [
        command,
        str(pptx_path),
        "--output",
        fmt,
        "--output-file",
        str(output_path),
    ]

    logger.debug("Executing DocStrange: %s", " ".join(args))
    completed = subprocess.run(args, capture_output=True, text=True)
    if completed.returncode != 0:
        raise RuntimeError(
            f"DocStrange command failed ({completed.returncode}): {completed.stderr.strip()}"
        )


def _extension_for_format(fmt: str) -> str:
    mapping = {
        "markdown": "md",
        "json": "json",
        "html": "html",
        "csv": "csv",
        "txt": "txt",
    }
    return mapping.get(fmt.lower(), fmt.lower())


def _write_diff(generated_output: Path, reference_output: Path, diff_output: Path) -> None:
    left = generated_output.read_text(encoding="utf-8")
    right = reference_output.read_text(encoding="utf-8")

    if generated_output.suffix == ".json" and reference_output.suffix == ".json":
        try:
            left_obj = json.loads(left)
            right_obj = json.loads(right)
        except json.JSONDecodeError:
            pass
        else:
            left = json.dumps(left_obj, indent=2, sort_keys=True)
            right = json.dumps(right_obj, indent=2, sort_keys=True)

    import difflib

    diff_lines = difflib.unified_diff(
        right.splitlines(),
        left.splitlines(),
        fromfile="reference",
        tofile="generated",
        lineterm="",
    )

    diff_output.write_text("\n".join(diff_lines), encoding="utf-8")
