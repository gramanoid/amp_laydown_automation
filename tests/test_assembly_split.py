from __future__ import annotations

from copy import copy

from amp_automation.presentation import assembly


def _format_k(value: float) -> str:
    if value <= 0:
        return "-"
    return f"£{int(round(value / 1000))}K"


def _build_sample_table() -> tuple[list[list[str]], dict[tuple[int, int], dict[str, object]]]:
    months = [month.upper() for month in assembly.TABLE_MONTH_ORDER]
    header = ["CAMPAIGN", "MEDIA", "METRICS", *months, "TOTAL", "GRPs", "%"]

    def budget_row(label: str) -> list[str]:
        row = [label, "TELEVISION", "£ 000"]
        for idx in range(len(months)):
            row.append(_format_k(10_000.0) if idx == 0 else "-")
        row.append(_format_k(10_000.0))
        row.append("")
        row.append("")
        return row

    def monthly_total_row(value: float) -> list[str]:
        row = ["MONTHLY TOTAL (£ 000)", "", ""]
        for idx in range(len(months)):
            row.append(_format_k(value) if idx == 0 else "-")
        row.append(_format_k(value))
        row.append("")
        row.append("")
        return row

    grand_total_value = 25_000.0
    grand_total_row = ["GRAND TOTAL", "", ""]
    for idx in range(len(months)):
        grand_total_row.append(_format_k(grand_total_value) if idx == 0 else "-")
    grand_total_row.append(_format_k(grand_total_value))
    grand_total_row.append("")
    grand_total_row.append("100.0%")

    table_data = [
        header,
        budget_row("Campaign A"),
        monthly_total_row(10_000.0),
        budget_row("Campaign B"),
        monthly_total_row(15_000.0),
        grand_total_row,
    ]

    total_col_idx = 3 + len(months)
    cell_metadata: dict[tuple[int, int], dict[str, object]] = {
        (2, 3): {"value": 10_000.0, "has_data": True, "media_type": "Subtotal"},
        (2, total_col_idx): {"value": 10_000.0, "has_data": True, "media_type": "Subtotal"},
        (4, 3): {"value": 15_000.0, "has_data": True, "media_type": "Subtotal"},
        (4, total_col_idx): {"value": 15_000.0, "has_data": True, "media_type": "Subtotal"},
    }

    return table_data, cell_metadata


def test_split_table_data_creates_continuations() -> None:
    table_data, cell_metadata = _build_sample_table()

    original_max = assembly.MAX_ROWS_PER_SLIDE
    original_boundaries = copy(assembly._CAMPAIGN_BOUNDARIES)
    original_show_carried = assembly.SHOW_CARRIED_SUBTOTAL

    try:
        assembly.MAX_ROWS_PER_SLIDE = 2
        assembly._CAMPAIGN_BOUNDARIES = [(1, 2), (3, 4)]
        assembly.SHOW_CARRIED_SUBTOTAL = True

        splits = assembly._split_table_data_by_campaigns(table_data, cell_metadata)

        assert len(splits) == 2
        first_rows, first_meta, first_continuation = splits[0]
        second_rows, second_meta, second_continuation = splits[1]

        assert first_continuation is False
        assert second_continuation is True

        assert first_rows[-2][0] == "CARRIED FORWARD"
        assert first_rows[-1][0] == "GRAND TOTAL"
        grand_total_row_idx = len(first_rows) - 1
        assert (grand_total_row_idx, 3) in {
            key for key in first_meta if key[0] == grand_total_row_idx
        }

        assert second_rows[-1][0] == "GRAND TOTAL"
        second_grand_total_idx = len(second_rows) - 1
        assert (second_grand_total_idx, 3) in {
            key for key in second_meta if key[0] == second_grand_total_idx
        }
    finally:
        assembly.MAX_ROWS_PER_SLIDE = original_max
        assembly._CAMPAIGN_BOUNDARIES = original_boundaries
        assembly.SHOW_CARRIED_SUBTOTAL = original_show_carried
