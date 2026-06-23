"""Tests for the goal achievement audit example."""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_DIR = REPO_ROOT / "examples" / "goal_achievement"
FIXTURES_DIR = EXAMPLE_DIR / "fixtures"
if str(EXAMPLE_DIR) not in sys.path:
    sys.path.insert(0, str(EXAMPLE_DIR))

import goal_achievement  # noqa: E402

_FIXTURE_SPEC = importlib.util.spec_from_file_location(
    "goal_achievement_build_fixture",
    FIXTURES_DIR / "build_fixture.py",
)
assert _FIXTURE_SPEC and _FIXTURE_SPEC.loader
_build_fixture_module = importlib.util.module_from_spec(_FIXTURE_SPEC)
_FIXTURE_SPEC.loader.exec_module(_build_fixture_module)


def _fixtures(tmp_path: Path) -> tuple[Path, Path]:
    metrics = tmp_path / "metrics.xlsx"
    accred = tmp_path / "accreditation.xlsx"
    _build_fixture_module.build_metrics_fixture(metrics)
    _build_fixture_module.build_accreditation_fixture(accred)
    return metrics, accred


def test_cross_audit_flags_count_and_completion_mismatch(tmp_path: Path) -> None:
    metrics, accred = _fixtures(tmp_path)
    drilldown = goal_achievement.load_drilldown_rows(accred)
    metrics_rows = goal_achievement.audit.load_student_metrics_workbook(metrics)
    results = goal_achievement.evaluate_cross_audit(drilldown, metrics_rows)
    types = {row["exception_type"] for row in results["exception_rows"]}
    assert "achievement_count_mismatch" in types
    assert "all_progress_not_marked_complete" in types


def test_inherited_gar_flags_mismatch(tmp_path: Path) -> None:
    metrics, _ = _fixtures(tmp_path)
    metrics_rows = goal_achievement.audit.load_student_metrics_workbook(metrics)
    gar_module = goal_achievement.audit_sources.load_audit_gar()
    results = goal_achievement.run_inherited_gar_audit(
        gar_module,
        metrics_rows,
        school=None,
        metric=None,
    )
    types = {row["exception_type"] for row in results["exception_rows"]}
    assert "goal_achievement_mismatch" in types
    assert results["summary"]["mismatched"] >= 1


def test_inherited_metrics_audit_flags_blank_baseline_target(tmp_path: Path) -> None:
    metrics, _ = _fixtures(tmp_path)
    metrics_rows = goal_achievement.audit.load_student_metrics_workbook(metrics)
    metrics_module = goal_achievement.audit_sources.load_audit_metrics_all()
    results = goal_achievement.run_inherited_metrics_audit(
        metrics_module,
        metrics_rows,
        include_domain_scale=True,
        ramp_days_after_enrollment=metrics_module.DEFAULT_RAMP_DAYS_AFTER_ENROLLMENT,
        quarter_end_grace_days=metrics_module.DEFAULT_QUARTER_END_GRACE_DAYS,
        as_of_date=None,
    )
    assert results["issue_code_counts"].get("both_baseline_and_target_blank", 0) >= 1


def test_cli_writes_outputs(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("GOAL_ACHIEVEMENT_OUTPUT_DIR", str(tmp_path))
    metrics, accred = _fixtures(tmp_path)
    code = goal_achievement.main(
        [
            "--student-metrics-workbook",
            str(metrics),
            "--accreditation-workbook",
            str(accred),
        ]
    )
    assert code == 0
    payload = json.loads((tmp_path / "goal_achievement_summary.json").read_text(encoding="utf-8"))
    assert payload["metrics_row_count"] == 3
    assert payload["drilldown_row_count"] == 2
    assert (tmp_path / "metrics_audit_flags.csv").is_file()
    assert (tmp_path / "gar_exceptions.csv").is_file()
    assert (tmp_path / "drilldown_exceptions.csv").is_file()
    assert payload["exception_totals"]["gar"] >= 1


def test_cli_resolves_workbooks_via_dual_fetch(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("GOAL_ACHIEVEMENT_OUTPUT_DIR", str(tmp_path))
    metrics, accred = _fixtures(tmp_path)
    with patch.object(
        goal_achievement,
        "resolve_inputs",
        return_value=goal_achievement.AuditInputs(
            student_metrics_workbook=metrics,
            accreditation_workbook=accred,
        ),
    ):
        code = goal_achievement.main([])
    assert code == 0
