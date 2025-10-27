# Visual Fidelity Verification Report

## Overview
- **Generated deck:** `output/presentations/run_20251015_134745/AMP_Presentation_20251015_134745.pptx`
- **Baseline template:** `template/Template_V4_FINAL_071025.pptx`
- **Comparison method:** Automated export to PNG via `tools/visual_diff.py`, followed by per-slide geometry inspection using `python-pptx`.

## Automated Diff Results
- Exported images:
  - Baseline: `output/visual_diff/template/Slide1.PNG`
  - Candidate: `output/visual_diff/generated/Slide2.PNG`
  - Diff: `output/visual_diff/diff_slide2_vs_template.png`
- Quantitative metrics (Slide 2 vs. Template Slide 1):
  - Mean pixel difference: **42.97**
  - RMS pixel difference: **78.14**
  - Maximum channel delta: **255.00**
- Interpretation: Values are significantly above zero, confirming noticeable visual drift relative to the master slide.

## Layout Geometry Findings
- Shared slide elements (Quarter budget tiles, media share tiles, funnel share tiles, footer notes) match the template exactly in position and size.
- `MainDataTable` discrepancies:
  - Left offset: +95 EMUs (~0.010 inches) from template location.
  - Top offset: +134 EMUs (~0.015 inches).
  - Width delta: −1,650,235 EMUs (~−180.54 inches aggregated column width; indicates compressed layout across continuation logic).
  - Height delta: −1,677,939 EMUs (~−183.56 inches), driven by fewer rows and reduced row heights.
- Column width comparison:
  - **Template** widths (inches): `[0.888, 0.798, 0.909, 0.370, 0.438, 0.438, 0.438, 0.454, 0.455, 0.509, 0.479, 0.485, 0.438, 0.479, 0.385, 0.491, 0.438, 0.438]`
  - **Generated** widths (inches): `[0.650, 0.500, 0.350, 0.430, 0.350, 0.400, 0.720, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375]`
  - The generated table reuses the flattened 0.375″ widths for the majority of columns, shrinking the overall footprint below the template’s 9.33″ target width.

## Recommended Next Steps
1. **Align Table Column Widths:**
   - Update `TABLE_COLUMN_WIDTHS` (and any related config overrides) to mirror the template’s per-column measurements.
   - Ensure `TableLayout.position` matches the template geometry (left 0.179″, width 9.33″, etc.).
2. **Regenerate Presentation:**
   - Re-run the CLI to build a fresh deck after geometry adjustments.
   - Re-execute `tools/visual_diff.py` to produce a new diff; expect metric values trending toward zero.
3. **Visual Sanity Check:**
   - Manually open the regenerated PPTX alongside the template and use PowerPoint’s *Review → Compare* feature to confirm alignment.
4. **Automated Monitoring (Optional):**
   - Extend `visual_diff.py` to iterate across multiple slides, recording metrics and thresholds for regression alerts.

## Artifacts Generated
- `output/visual_diff/template/Slide1.PNG`
- `output/visual_diff/generated/Slide2.PNG` through `Slide101.PNG`
- `output/visual_diff/diff_slide2_vs_template.png`
- `tools/visual_diff.py` utility (PowerPoint automation + Pillow comparison)
