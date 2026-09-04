"""Tests for goal achievement rollup visualization helpers."""

from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
TRENDS_DIR = REPO_ROOT / "examples" / "trends"
if str(TRENDS_DIR) not in sys.path:
    sys.path.insert(0, str(TRENDS_DIR))

import visualize  # noqa: E402


def test_aggregate_school_goal_achievement_mismatches_sorts_by_count() -> None:
    rollup = {
        "schools": {
            "Alpha School": {"mismatched": 2, "mismatch_rate_pct": 10.0},
            "Beta School": {"mismatched": 5, "mismatch_rate_pct": 20.0},
        }
    }
    rows = visualize.aggregate_school_goal_achievement_mismatches(rollup)
    assert rows[0]["school"] == "Beta School"
    assert rows[0]["mismatched"] == 5


def test_line_chart_value_label_rounds_rates() -> None:
    assert visualize._line_chart_value_label(53.9) == 54
    assert visualize._line_chart_value_label(50.6) == 51
    assert visualize._line_chart_value_label(47.7) == 48


def test_gar_goal_not_met_pct_combines_progress_and_no_progress() -> None:
    outcomes = {
        "Goal Met": 4,
        "Goal Not Met, With Progress": 2,
        "Goal Not Met, No Progress": 1,
    }
    assert visualize.gar_goal_not_met_pct(outcomes) == 42.9


def test_load_goal_achievement_rollup_reads_json(tmp_path: Path) -> None:
    audit_dir = tmp_path / "audit"
    audit_dir.mkdir()
    payload = {
        "global": {
            "recorded_goal_met_pct": 42.0,
            "expected_goal_met_pct": 44.0,
            "mismatch_rate_pct": 3.5,
            "recorded": {"Goal Met": 4, "Goal Not Met, With Progress": 2, "Goal Not Met, No Progress": 1},
            "expected": {"Goal Met": 5, "Goal Not Met, With Progress": 1, "Goal Not Met, No Progress": 1},
        },
        "schools": {},
    }
    path = audit_dir / "SY24-25_GoalAchievement_ROLLUP.json"
    path.write_text(json.dumps(payload), encoding="utf-8")
    loaded = visualize.load_goal_achievement_rollup(audit_dir, "SY24-25")
    assert loaded is not None
    assert loaded["global"]["recorded_goal_met_pct"] == 42.0
