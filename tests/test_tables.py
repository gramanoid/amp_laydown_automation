"""Tests for table assembly helpers."""

from __future__ import annotations

import logging

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from amp_automation.presentation.tables import (
    CellStyleContext,
    TableLayout,
    add_and_style_table,
)


def build_context() -> CellStyleContext:
    return CellStyleContext(
        margin_left_right_pt=3.6,
        margin_emu_lr=45720,
        default_font_name="Calibri",
        font_size_header=Pt(10),
        font_size_body=Pt(9),
        color_black=RGBColor(0, 0, 0),
        color_light_gray_text=RGBColor(200, 200, 200),
        color_table_gray=RGBColor(191, 191, 191),
        color_header_green=RGBColor(0, 255, 0),
        color_subtotal_gray=RGBColor(217, 217, 217),
        color_tv=RGBColor(113, 212, 141),
        color_digital=RGBColor(253, 242, 183),
        color_ooh=RGBColor(255, 191, 0),
        color_other=RGBColor(176, 211, 255),
    )


def build_layout() -> TableLayout:
    return TableLayout(
        placeholder_name="",
        shape_name="TestTable",
        position={
            "left": Inches(1),
            "top": Inches(1),
            "width": Inches(4),
            "height": Inches(2),
        },
        row_height_header=Pt(12),
        row_height_body=Pt(10),
        row_height_subtotal=Pt(10),
        column_widths=[Inches(1.5), Inches(1.5)],
        top_override=None,
        height_rule_available=False,
    )


def test_add_and_style_table_populates_cells() -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    table_data = [
        ["Header A", "Header B"],
        ["Row 1", "100"],
    ]
    cell_metadata: dict[tuple[int, int], dict[str, object]] = {}

    result = add_and_style_table(
        slide,
        table_data,
        cell_metadata,
        build_layout(),
        build_context(),
        logging.getLogger("test"),
    )

    assert result is True
    created_table = slide.shapes[0].table
    assert created_table.cell(0, 0).text.upper() == "HEADER A"
    assert "100" in created_table.cell(1, 1).text
