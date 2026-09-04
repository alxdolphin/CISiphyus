"""Tests for cisiphyus trend --html wiring."""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
TRENDS_DIR = REPO_ROOT / "examples" / "trends"
if str(TRENDS_DIR) not in sys.path:
    sys.path.insert(0, str(TRENDS_DIR))

import trends  # noqa: E402


def test_trend_parser_accepts_html_flag() -> None:
    parser = trends._build_trend_parser()
    args = parser.parse_args(["SY25-26", "SY24-25", "--html"])
    assert args.html is True
    assert args.school_years == ["SY25-26", "SY24-25"]


def test_finish_trend_html_report_rerenders_when_compares_already_ran(tmp_path: Path) -> None:
    output_dir = tmp_path / "trends"
    cross_year = output_dir / "cross_year" / "SY24-25_EOY_vs_SY25-26_EOY"
    cross_year.mkdir(parents=True)
    (cross_year / "trend_summary.json").write_text(
        '{"baseline_school_year":"SY24-25","current_school_year":"SY25-26"}',
        encoding="utf-8",
    )
    snapshots_dir = tmp_path / "snapshots"
    snapshots_dir.mkdir()
    args = trends._build_trend_parser().parse_args(["--html"])

    with patch.object(trends, "render_cross_year_report", return_value=0) as render:
        code = trends.finish_trend_html_report(
            args,
            output_dir=output_dir,
            snapshots_dir=snapshots_dir,
            school_years=["SY24-25", "SY25-26"],
            regression_threshold=0.01,
            compares_already_ran=True,
        )

    assert code == 0
    render.assert_called_once()
