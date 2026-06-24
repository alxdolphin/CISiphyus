"""Tests for the goal achievement audit example."""

from __future__ import annotations

import importlib.util
import json
import sys
from datetime import date
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_DIR = REPO_ROOT / "examples" / "goal_achievement"
AUDIT_DIR = REPO_ROOT / "examples" / "audit"
FIXTURES_DIR = EXAMPLE_DIR / "fixtures"
for path in (EXAMPLE_DIR, AUDIT_DIR):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

import audit  # noqa: E402
import goal_achievement  # noqa: E402

_FIXTURE_SPEC = importlib.util.spec_from_file_location(
    "goal_achievement_build_fixture",
    FIXTURES_DIR / "build_fixture.py",
)
assert _FIXTURE_SPEC and _FIXTURE_SPEC.loader
_build_fixture_module = importlib.util.module_from_spec(_FIXTURE_SPEC)
_FIXTURE_SPEC.loader.exec_module(_build_fixture_module)


def _fixtures(tmp_path: Path) -> tuple[Path, Path, Path]:
    metrics = tmp_path / "metrics.xlsx"
    accred = tmp_path / "accreditation.xlsx"
    goal_progress = tmp_path / "goal_progress.xlsx"
    _build_fixture_module.build_metrics_fixture(metrics)
    _build_fixture_module.build_accreditation_fixture(accred)
    _build_fixture_module.build_goal_progress_fixture(goal_progress)
    return metrics, accred, goal_progress


def test_cross_audit_flags_count_and_completion_mismatch(tmp_path: Path) -> None:
    metrics, accred, _ = _fixtures(tmp_path)
    drilldown = goal_achievement.load_drilldown_rows(accred)
    metrics_rows = goal_achievement.audit.load_student_metrics_workbook(metrics)
    results = goal_achievement.evaluate_cross_audit(drilldown, metrics_rows)
    types = {row["exception_type"] for row in results["exception_rows"]}
    assert "achievement_count_mismatch" in types
    assert "all_progress_not_marked_complete" in types


def test_gar_flags_mismatch_on_goal_progress_fixture(tmp_path: Path) -> None:
    _, _, goal_progress = _fixtures(tmp_path)
    records = audit.load_goal_progress_gar_records(goal_progress)
    results = audit.run_gar_audit(records, school=None, metric=None)
    types = {row["exception_type"] for row in results["exception_rows"]}
    assert "goal_achievement_mismatch" in types
    assert results["summary"]["mismatched"] >= 1


def test_gar_parity_summarize_counts(tmp_path: Path) -> None:
    _, _, goal_progress = _fixtures(tmp_path)
    records = audit.load_goal_progress_rows(goal_progress)
    context: dict[str, str] = {}
    results = audit.gar_summarize(records, context)
    summary = results["summary"]
    assert summary["total_rows"] == 3
    assert summary.get("expected::Goal Met", 0) >= 1
    assert summary["mismatched"] >= 1


def test_gar_baseline_grace_on_goal_progress_fixture(tmp_path: Path) -> None:
    _, _, goal_progress = _fixtures(tmp_path)
    records = audit.load_goal_progress_rows(goal_progress)
    grace_rows = [
        record
        for record in records
        if audit.gar_baseline_grace_applies(
            record,
            str(record.get("Metric") or ""),
            date.today(),
        )
    ]
    assert len(grace_rows) >= 1


def test_inherited_metrics_audit_flags_blank_baseline_target(tmp_path: Path) -> None:
    metrics, _, _ = _fixtures(tmp_path)
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
    archive_dir = tmp_path / "goal_achievement_audit"
    metrics, accred, goal_progress = _fixtures(tmp_path)
    import goal_achievement_audit as ga_audit

    monkeypatch.setattr(ga_audit, "goal_achievement_audit_archive_dir", lambda: archive_dir)
    code = goal_achievement.main(
        [
            "--goal-progress-workbook",
            str(goal_progress),
            "--student-metrics-workbook",
            str(metrics),
            "--accreditation-workbook",
            str(accred),
            "--school-year",
            "SY25-26",
        ]
    )
    assert code == 0
    payload = json.loads((tmp_path / "goal_achievement_summary.json").read_text(encoding="utf-8"))
    assert payload["gar_row_count"] == 3
    assert payload["drilldown_row_count"] == 2
    assert (tmp_path / "gar_exceptions.csv").is_file()
    assert (tmp_path / "gar_audit_report.md").is_file()
    assert (tmp_path / "drilldown_exceptions.csv").is_file()
    assert payload["exception_totals"]["gar"] >= 1
    assert (archive_dir / "SY25-26_GoalAchievement_AUDIT.csv").is_file()


def test_cli_resolves_workbooks_via_dual_fetch(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("GOAL_ACHIEVEMENT_OUTPUT_DIR", str(tmp_path))
    metrics, accred, goal_progress = _fixtures(tmp_path)
    with patch.object(
        goal_achievement,
        "resolve_inputs",
        return_value=goal_achievement.AuditInputs(
            goal_progress_workbook=goal_progress,
            student_metrics_workbook=metrics,
            accreditation_workbook=accred,
        ),
    ):
        code = goal_achievement.main([])
    assert code == 0


def test_gar_constants_in_audit_module() -> None:
    assert audit.GOAL_PROGRESS_SHEET == "CIS_StudentProgress_Detail"
    assert audit.GAR_BASELINE_GRACE_DAYS == 45
