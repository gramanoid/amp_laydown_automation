"""Regression tests for font normalization (EC-002)."""

from __future__ import annotations

import pytest
from pptx.util import Pt

from conftest import find_main_table, skipif_no_deck


@pytest.mark.regression
@skipif_no_deck
def test_ec002_font_sizes_consistent_across_production_deck(latest_deck_path):
    """Verify font sizes are consistent across all slides in production deck (EC-002)."""
    from pptx import Presentation

    prs = Presentation(latest_deck_path)

    font_size_expectations = {
        "header": Pt(7),       # 27-10-25 fix
        "body": Pt(6),         # 27-10-25 fix
    }

    failed_cells = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        table = find_main_table(slide)
        if not table:
            continue

        for row_idx, row in enumerate(table.rows):
            # Determine expected font size based on row position
            if row_idx == 0:
                expected_size = font_size_expectations["header"]
                row_type = "header"
            elif row_idx == len(table.rows) - 1:
                # Last row might be GRAND TOTAL or BRAND TOTAL
                expected_size = font_size_expectations["header"]
                row_type = "grand_total"
            else:
                expected_size = font_size_expectations["body"]
                row_type = "body"

            # Check all cells in row
            for col_idx in range(len(table.columns)):
                cell = table.cell(row_idx, col_idx)
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size and run.font.size != expected_size:
                            failed_cells.append({
                                "slide": slide_idx,
                                "row": row_idx,
                                "col": col_idx,
                                "row_type": row_type,
                                "expected": expected_size,
                                "actual": run.font.size,
                                "text": run.text[:30],
                            })

    assert not failed_cells, \
        f"Font size inconsistencies found:\n" + \
        "\n".join(f"Slide {c['slide']}, Row {c['row']}, Col {c['col']} ({c['row_type']}): " \
                  f"Expected {c['expected']}, got {c['actual']} (text: {c['text']})" \
                  for c in failed_cells[:10])  # Show first 10


@pytest.mark.unit
def test_font_size_header_consistency(blank_slide, cell_style_context):
    """Verify header cell styling applies 7pt font (EC-002)."""
    from amp_automation.presentation.tables import style_table_cell

    prs = blank_slide._parent
    table = blank_slide.shapes.add_table(2, 2, left=0, top=0, width=1000000, height=1000000).table

    header_cell = table.cell(0, 0)
    header_cell.text = "Header"

    # Style header cell
    style_table_cell(
        header_cell,
        cell_style_context,
        is_header=True,
        is_subtotal=False,
        background_rgb=None,
    )

    # Verify font size
    for para in header_cell.text_frame.paragraphs:
        for run in para.runs:
            if run.text.strip():  # Only check non-empty runs
                assert run.font.size == Pt(7), \
                    f"Header should be 7pt, got {run.font.size}"


@pytest.mark.unit
def test_font_size_body_consistency(blank_slide, cell_style_context):
    """Verify body cell styling applies 6pt font (EC-002)."""
    from amp_automation.presentation.tables import style_table_cell

    table = blank_slide.shapes.add_table(2, 2, left=0, top=0, width=1000000, height=1000000).table

    body_cell = table.cell(1, 0)
    body_cell.text = "Body Row"

    # Style body cell
    style_table_cell(
        body_cell,
        cell_style_context,
        is_header=False,
        is_subtotal=False,
        background_rgb=None,
    )

    # Verify font size
    for para in body_cell.text_frame.paragraphs:
        for run in para.runs:
            if run.text.strip():
                assert run.font.size == Pt(6), \
                    f"Body should be 6pt, got {run.font.size}"


@pytest.mark.unit
def test_font_family_consistent(blank_slide, cell_style_context):
    """Verify font family is Calibri throughout (EC-002)."""
    from amp_automation.presentation.tables import style_table_cell

    table = blank_slide.shapes.add_table(2, 2, left=0, top=0, width=1000000, height=1000000).table

    for row_idx in range(2):
        for col_idx in range(2):
            cell = table.cell(row_idx, col_idx)
            cell.text = f"Cell {row_idx},{col_idx}"

            style_table_cell(
                cell,
                cell_style_context,
                is_header=row_idx == 0,
                is_subtotal=False,
                background_rgb=None,
            )

    # Verify font family
    for row_idx in range(2):
        for col_idx in range(2):
            cell = table.cell(row_idx, col_idx)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        assert run.font.name == "Calibri", \
                            f"Font should be Calibri, got {run.font.name}"
