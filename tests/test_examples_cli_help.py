"""Tests for top-level cisiphyus --help."""

from __future__ import annotations

import sys
from io import StringIO
from pathlib import Path
from unittest.mock import patch

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import cli  # noqa: E402
import help_text  # noqa: E402


def test_wants_full_help_bare_and_flags() -> None:
    assert help_text.wants_full_help(["cisiphyus"])
    assert help_text.wants_full_help(["cisiphyus", "--help"])
    assert help_text.wants_full_help(["cisiphyus", "-h"])
    assert help_text.wants_full_help(["cisiphyus", "help"])


def test_wants_full_help_false_for_subcommands() -> None:
    assert not help_text.wants_full_help(["cisiphyus", "pull", "--help"])
    assert not help_text.wants_full_help(["cisiphyus", "audit", "--help"])
    assert not help_text.wants_full_help(["cisiphyus", "student_metrics_summary"])
    assert not help_text.wants_full_help(["cisiphyus", "pull", "accreditation"])


def test_format_full_help_includes_command_sections() -> None:
    text = help_text.format_full_help(build_retrieval_parser=cli._build_retrieval_parser)
    assert "cisiphyus audit accreditation" in text
    assert "cisiphyus audit metrics" in text
    assert "cisiphyus audit goal-achievement" in text
    assert "cisiphyus qpr" in text
    assert "cisiphyus trend" in text
    assert "cisiphyus trend cross-year" in text
    assert "--grading-period" in text
    assert "--bootstrap" in text


def test_main_bare_prints_help_without_side_effects(monkeypatch) -> None:
    stdout = StringIO()
    monkeypatch.setattr(sys, "argv", ["cisiphyus"])
    monkeypatch.setattr(sys, "stdout", stdout)

    with (
        patch("config.migrate_legacy_layout") as migrate_mock,
        patch("cli._run_retrieval") as retrieval_mock,
        patch("cli.example_cli.maybe_run_example_app", return_value=None) as example_mock,
        pytest.raises(SystemExit) as exc,
    ):
        cli.main()

    assert exc.value.code == 0
    migrate_mock.assert_not_called()
    retrieval_mock.assert_not_called()
    example_mock.assert_not_called()
    assert "cisiphyus qpr" in stdout.getvalue()
