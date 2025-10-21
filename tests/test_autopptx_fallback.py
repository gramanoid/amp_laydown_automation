import pandas as pd
from pathlib import Path

from amp_automation.presentation import assembly


def _sample_dataset() -> pd.DataFrame:
    """Provide a minimal dataset compatible with _prepare_main_table_data_detailed."""

    months = {
        "Jan": 1200.0,
        "Feb": 0.0,
        "Mar": 0.0,
        "Apr": 0.0,
        "May": 0.0,
        "Jun": 0.0,
        "Jul": 0.0,
        "Aug": 0.0,
        "Sep": 0.0,
        "Oct": 0.0,
        "Nov": 0.0,
        "Dec": 0.0,
    }

    return pd.DataFrame(
        [
            {
                "Country": "Testland",
                "Brand": "BrandX",
                "Media Type": "Digital",
                "Mapped Media Type": "Digital",
                "Campaign Name": "Campaign Alpha",
                "Campaign Type": "Brand",
                "Funnel Stage": "Awareness",
                "Year": 2025,
                **months,
                "Total Cost": sum(months.values()),
                "GRP": 0.0,
                "Frequency": 1.5,
                "Reach 1+": 0.25,
                "Reach 3+": 0.05,
                "Flight Comments": "",
            }
        ]
    )


def test_generate_autopptx_only_handles_year(monkeypatch, tmp_path):
    calls: dict[str, object] = {}

    monkeypatch.setattr(
        assembly.autopptx_adapter,
        "autopptx_available",
        lambda: True,
    )

    def _fake_generate(template_path, payloads, output_path, **kwargs):
        calls["template_path"] = Path(template_path)
        calls["output_path"] = Path(output_path)
        calls["payloads"] = list(payloads)
        return Path(output_path)

    monkeypatch.setattr(
        assembly.autopptx_adapter,
        "generate_presentation",
        _fake_generate,
    )

    df = _sample_dataset()
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "generated.pptx"

    result = assembly._generate_autopptx_only(
        template_path,
        output_path,
        df,
        [("Testland", "BrandX", 2025)],
        excel_path=None,
    )

    assert result is True
    assert "payloads" in calls and calls["payloads"], "AutoPPTX should receive payloads"

    payload = calls["payloads"][0]
    assert payload.tables, "Slide payload must include table data"
    header_row = payload.tables[0][0]
    assert header_row == assembly.TABLE_HEADER_COLUMNS
