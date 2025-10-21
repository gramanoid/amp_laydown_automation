"""Utilities for comparing generated presentations against a reference deck."""

from __future__ import annotations

import argparse
import json
import math
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageChops, ImageStat
import win32com.client


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@dataclass
class DiffResult:
    slide: Path
    metrics: dict[str, float]


def _export_presentation(source: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 1
    try:
        presentation = app.Presentations.Open(str(source), False, False, False)
        try:
            presentation.Export(str(destination), "PNG")
        finally:
            presentation.Close()
    finally:
        app.Quit()


def _compare_images(baseline: Path, candidate: Path, diff_output: Path | None = None) -> dict[str, float]:
    baseline_image = Image.open(baseline).convert("RGB")
    candidate_image = Image.open(candidate).convert("RGB")

    if baseline_image.size != candidate_image.size:
        raise ValueError("Images must share identical dimensions for comparison")

    diff_image = ImageChops.difference(baseline_image, candidate_image)
    if diff_output is not None:
        diff_output.parent.mkdir(parents=True, exist_ok=True)
        diff_image.save(diff_output)

    stat = ImageStat.Stat(diff_image)
    mean = sum(stat.mean) / len(stat.mean)
    rms = math.sqrt(sum(value**2 for value in stat.rms) / len(stat.rms))
    extrema = max(value for channel in stat.extrema for value in channel)

    return {
        "mean_difference": mean,
        "rms_difference": rms,
        "max_channel_difference": extrema,
    }


def _diff_all_slides(
    template_export_dir: Path,
    generated_export_dir: Path,
    diff_dir: Path,
    max_slides: int | None,
) -> list[DiffResult]:
    template_slides = sorted(template_export_dir.glob("Slide*.PNG"))
    generated_slides = sorted(generated_export_dir.glob("Slide*.PNG"))
    if not template_slides or not generated_slides:
        raise SystemExit("Exported slide images missing for diff comparison")

    if max_slides is not None:
        template_slides = template_slides[:max_slides]
        generated_slides = generated_slides[:max_slides]

    results: list[DiffResult] = []
    for idx, (baseline_slide, candidate_slide) in enumerate(zip(template_slides, generated_slides), start=1):
        diff_path = diff_dir / f"diff_slide{idx:03d}_vs_reference.png"
        metrics = _compare_images(baseline_slide, candidate_slide, diff_path)
        results.append(DiffResult(candidate_slide, metrics))

    return results


def _print_summary(diff_results: Iterable[DiffResult]) -> None:
    for result in diff_results:
        print(f"Slide {result.slide.name} metrics:")
        for key, value in result.metrics.items():
            print(f"  {key}: {value:.4f}")


def _validate_thresholds(diff_results: Iterable[DiffResult], mean_threshold: float, rms_threshold: float) -> None:
    failing = []
    for result in diff_results:
        mean = result.metrics.get("mean_difference", 0.0)
        rms = result.metrics.get("rms_difference", 0.0)
        if mean > mean_threshold or rms > rms_threshold:
            failing.append((result.slide.name, mean, rms))

    if failing:
        summary = "; ".join(f"{name}: mean={mean:.4f}, rms={rms:.4f}" for name, mean, rms in failing)
        raise SystemExit(f"Visual diff thresholds exceeded -> {summary}")


def _write_json_summary(diff_results: Iterable[DiffResult], output_path: Path) -> None:
    payload = [
        {
            "slide": result.slide.name,
            "metrics": result.metrics,
        }
        for result in diff_results
    ]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, indent=2))


def _default_generated_deck() -> Path:
    presentations_root = PROJECT_ROOT / "output" / "presentations"
    generated_runs = sorted(presentations_root.glob("run_*/GeneratedDeck_*.pptx"))
    if generated_runs:
        return generated_runs[-1]

    generated_runs = sorted(presentations_root.glob("run_*/AMP_Presentation_*.pptx"))
    if generated_runs:
        return generated_runs[-1]

    available = sorted(presentations_root.glob("**/*.pptx"))
    message = "No generated presentations found in output/presentations"
    if available:
        message += f" (found other files: {available})"
    raise SystemExit(message)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run a pixel diff between two PowerPoint decks.")
    parser.add_argument(
        "--reference",
        type=Path,
        default=PROJECT_ROOT / "template" / "Template_V4_FINAL_071025.pptx",
        help="Reference/template PPTX path (default: Template_V4_FINAL_071025.pptx).",
    )
    parser.add_argument(
        "--generated",
        type=Path,
        default=None,
        help="Generated PPTX to validate (default: latest run_* deck in output/presentations).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=PROJECT_ROOT / "output" / "visual_diff",
        help="Root directory for exports, diffs, and summary (default: output/visual_diff).",
    )
    parser.add_argument(
        "--max-slides",
        type=int,
        default=None,
        help="Optional limit for number of slides to diff.",
    )
    parser.add_argument(
        "--mean-threshold",
        type=float,
        default=0.5,
        help="Mean pixel difference threshold (default: 0.5).",
    )
    parser.add_argument(
        "--rms-threshold",
        type=float,
        default=0.5,
        help="RMS pixel difference threshold (default: 0.5).",
    )
    parser.add_argument(
        "--keep-exports",
        action="store_true",
        help="Keep existing PNG exports instead of clearing directories before exporting.",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()

    reference_path = args.reference.resolve()
    generated_path = (args.generated or _default_generated_deck()).resolve()
    output_root = args.output_dir.resolve()

    if not reference_path.is_file():
        raise SystemExit(f"Reference deck not found: {reference_path}")
    if not generated_path.is_file():
        raise SystemExit(f"Generated deck not found: {generated_path}")

    reference_export_dir = output_root / "exports" / "reference" / reference_path.stem
    generated_export_dir = output_root / "exports" / "generated" / generated_path.stem
    diff_dir = output_root / "diffs" / f"{generated_path.stem}_vs_{reference_path.stem}"
    summary_path = diff_dir / "diff_summary.json"

    if not args.keep_exports:
        for directory in (reference_export_dir, generated_export_dir, diff_dir):
            if directory.exists():
                shutil.rmtree(directory)

    reference_has_exports = reference_export_dir.exists() and any(reference_export_dir.glob("Slide*.PNG"))
    generated_has_exports = generated_export_dir.exists() and any(generated_export_dir.glob("Slide*.PNG"))

    if reference_has_exports and args.keep_exports:
        print(f"[visual-diff] Reusing existing reference exports at {reference_export_dir}")
    else:
        print(f"[visual-diff] Exporting reference deck -> {reference_export_dir}")
        _export_presentation(reference_path, reference_export_dir)

    if generated_has_exports and args.keep_exports:
        print(f"[visual-diff] Reusing existing generated exports at {generated_export_dir}")
    else:
        print(f"[visual-diff] Exporting generated deck -> {generated_export_dir}")
        _export_presentation(generated_path, generated_export_dir)

    diff_results = _diff_all_slides(
        reference_export_dir,
        generated_export_dir,
        diff_dir,
        max_slides=args.max_slides,
    )

    _print_summary(diff_results)
    _write_json_summary(diff_results, summary_path)
    _validate_thresholds(diff_results, args.mean_threshold, args.rms_threshold)


if __name__ == "__main__":  # pragma: no cover - utility entrypoint
    main()
