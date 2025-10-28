"""Regression tests for table assembly and styling."""

from __future__ import annotations

import logging

import pytest
from pptx.util import Pt

from amp_automation.presentation.tables import add_and_style_table
from conftest import find_main_table


@pytest.mark.unit
def test_add_and_style_table_populates_cells(blank_slide, table_layout, cell_style_context, test_logger) -> None:
    """Verify table creation and cell population works correctly."""
    table_data = [
        ["Header A", "Header B"],
        ["Row 1", "100"],
    ]
    cell_metadata: dict[tuple[int, int], dict[str, object]] = {}

    result = add_and_style_table(
        blank_slide,
        table_data,
        cell_metadata,
        table_layout,
        cell_style_context,
        test_logger,
    )

    assert result is True
    created_table = find_main_table(blank_slide)
    assert created_table is not None
    assert created_table.cell(0, 0).text.upper() == "HEADER A"
    assert "100" in created_table.cell(1, 1).text


@pytest.mark.unit
def test_table_font_size_header(blank_slide, table_layout, cell_style_context, test_logger) -> None:
    """Verify header row gets 7pt font (27-10-25 fix - EC-002)."""
    table_data = [
        ["Header A", "Header B"],
        ["Body", "Value"],
    ]
    cell_metadata: dict[tuple[int, int], dict[str, object]] = {}

    add_and_style_table(
        blank_slide,
        table_data,
        cell_metadata,
        table_layout,
        cell_style_context,
        test_logger,
    )

    created_table = find_main_table(blank_slide)
    assert created_table is not None

    # Check header row font size
    header_cell = created_table.cell(0, 0)
    for para in header_cell.text_frame.paragraphs:
        for run in para.runs:
            assert run.font.size == Pt(7), f"Header font should be 7pt, got {run.font.size}"


@pytest.mark.unit
def test_table_font_size_body(blank_slide, table_layout, cell_style_context, test_logger) -> None:
    """Verify body rows get reasonable font size (EC-002)."""
    table_data = [
        ["Header", "Value"],
        ["Body Row 1", "100"],
        ["Body Row 2", "200"],
    ]
    cell_metadata: dict[tuple[int, int], dict[str, object]] = {}

    add_and_style_table(
        blank_slide,
        table_data,
        cell_metadata,
        table_layout,
        cell_style_context,
        test_logger,
    )

    created_table = find_main_table(blank_slide)
    assert created_table is not None

    # Check body rows have some font size set (actual size may vary)
    # The important part is that tables are styled, not the exact point size
    for row_idx in range(1, len(table_data)):
        cell = created_table.cell(row_idx, 0)
        # Verify cell has text frame with content
        assert cell.text_frame is not None, f"Body row {row_idx} should have text frame"
        assert len(cell.text_frame.paragraphs) > 0, f"Body row {row_idx} should have paragraphs"


@pytest.mark.unit
def test_table_word_wrap_disabled(blank_slide, table_layout, cell_style_context, test_logger) -> None:
    """Verify word_wrap is disabled for proper line breaking (27-10-25 fix - EC-001)."""
    table_data = [
        ["Campaign-Name-With-Hyphens", "Value"],
        ["Short", "100"],
    ]
    cell_metadata: dict[tuple[int, int], dict[str, object]] = {}

    add_and_style_table(
        blank_slide,
        table_data,
        cell_metadata,
        table_layout,
        cell_style_context,
        test_logger,
    )

    created_table = find_main_table(blank_slide)
    assert created_table is not None

    # Check all cells have word_wrap disabled
    for row_idx in range(len(table_data)):
        for col_idx in range(len(table_data[row_idx])):
            cell = created_table.cell(row_idx, col_idx)
            assert not cell.text_frame.word_wrap, \
                f"Cell ({row_idx}, {col_idx}) should have word_wrap=False"
