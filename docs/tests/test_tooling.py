from __future__ import annotations

from pathlib import Path

import pytest

from amp_automation.presentation import assembly
from amp_automation.tooling import aspose_converter, docstrange_validator


def test_aspose_requires_credentials(tmp_path: Path) -> None:
    pptx_path = tmp_path / "sample.pptx"
    pptx_path.write_bytes(b"fake")

    with pytest.raises(aspose_converter.AsposeConfigurationError):
        aspose_converter.export_with_aspose(
            pptx_path,
            ["pdf"],
            tmp_path,
            client_id=None,
            client_secret=None,
        )


def test_docstrange_command_probe_returns_false() -> None:
    assert docstrange_validator.docstrange_available(command="__nonexistent_cli__") is False


def test_autopptx_pipeline_skips_when_disabled(monkeypatch) -> None:
    monkeypatch.setitem(assembly.AUTOPPTX_CONFIG, "enabled", False)
    assembly._run_autopptx_pipeline("template.pptx", "output.pptx", [])
