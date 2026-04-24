from __future__ import annotations

import pytest

from pubify_ppt.cli import build_parser, main


def test_cli_help_includes_planned_commands(capsys: pytest.CaptureFixture[str]) -> None:
    parser = build_parser()

    with pytest.raises(SystemExit) as exc_info:
        parser.parse_args(["--help"])

    assert exc_info.value.code == 0
    output = capsys.readouterr().out
    assert "ppt init <presentation-id>" in output
    assert "ppt <presentation-id> update [--output <path>]" in output


def test_cli_entrypoint_reports_phase_zero_placeholder(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc_info:
        main(["demo", "update", "--output", "copy.pptx"])

    assert exc_info.value.code == 2
    assert "Phase 0 provides the CLI entrypoint" in capsys.readouterr().err
