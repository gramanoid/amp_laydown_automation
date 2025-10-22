"""Integration tests for the CLI."""

from __future__ import annotations

from amp_automation.cli import main as cli_main


def test_cli_list_templates(capsys) -> None:
    exit_code = cli_main(["--list-templates"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Template" in captured.out
