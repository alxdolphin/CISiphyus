"""Tests for the student metrics audit example."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_DIR = REPO_ROOT / "examples" / "audit"
FIXTURES_DIR = EXAMPLE_DIR / "fixtures"
if str(EXAMPLE_DIR) not in sys.path:
    sys.path.insert(0, str(EXAMPLE_DIR))
if str(FIXTURES_DIR) not in sys.path:
    sys.path.insert(0, str(FIXTURES_DIR))

import audit  # noqa: E402
from build_fixture import build_fixture  # noqa: E402


def _workbook(tmp_path: Path) -> Path:
    path = tmp_path / "student_metrics_summary_minimal.xlsx"
    build_fixture(path)
    return path


def _evaluate(workbook: Path) -> dict:
    return audit.evaluate_workbook(workbook, audit.AuditConfig())


def _record(**kwargs: object) -> dict[str, object]:
    base = {
        "School": "Test School",
        "Goal": "Improve Attendance",
        "Metric": "Attendance Rate (%)",
        "Baseline": "85",
        "Target": "90",
    }
    base.update(kwargs)
    return base


def test_workbook_loads_fixture_rows(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    records = audit.load_student_metrics_workbook(workbook)
    assert len(records) == 18
    assert records[0]["Metric"] == "Attendance Rate (%)"
    assert audit.row_number(records[0]) == 2


def test_fixture_loads_expected_row_count(tmp_path: Path) -> None:
    results = _evaluate(_workbook(tmp_path))
    assert results["summary"]["rows"] == 18


def test_fixture_flags_domain_and_scale_issues(tmp_path: Path) -> None:
    results = _evaluate(_workbook(tmp_path))
    counts = results["issue_code_counts"]
    assert counts.get("target_scale_mismatch", 0) >= 2
    assert counts.get("baseline_scale_mismatch", 0) >= 1
    assert counts.get("baseline_target_direction_mismatch", 0) >= 2
    assert counts.get("undocumented_value_annotation", 0) >= 1
    assert counts.get("goal_metric_domain_mismatch", 0) >= 1
    assert counts.get("manual_review_domain", 0) >= 1


def test_fixture_flags_structural_issues(tmp_path: Path) -> None:
    results = _evaluate(_workbook(tmp_path))
    counts = results["issue_code_counts"]
    assert counts.get("both_baseline_and_target_blank", 0) >= 1
    assert counts.get("both_blank_with_two_progress_reports", 0) >= 1
    assert counts.get("target_without_baseline", 0) >= 1
    assert counts.get("baseline_without_target", 0) >= 2
    assert counts.get("duplicate_composite_key", 0) >= 1
    assert counts.get("student_client_id_mismatch", 0) >= 2
    assert counts.get("case_manager_blank", 0) >= 1


def test_percent_attendance_rejects_days_absent_target() -> None:
    result = audit.analyze_record(
        _record(
            Goal="Improve Attendance",
            Metric="Attendance Rate (%)",
            Baseline="85",
            Target="5 days",
        )
    )
    assert "target_scale_mismatch" in result["issue_codes"]


def test_attendance_metric_rejects_letter_grade_target() -> None:
    result = audit.analyze_record(
        _record(
            Goal="Improve Attendance",
            Metric="Attendance Rate (%)",
            Baseline="85",
            Target="B",
        )
    )
    assert "target_scale_mismatch" in result["issue_codes"]


def test_core_course_grades_accepts_standard_letter_grade() -> None:
    result = audit.analyze_record(
        _record(
            Goal="Improve Academics",
            Metric="Core Course Grades",
            Baseline="C-",
            Target="B",
        )
    )
    assert "baseline_scale_mismatch" not in result["issue_codes"]
    assert "target_scale_mismatch" not in result["issue_codes"]


def test_undocumented_benchmark_annotation() -> None:
    result = audit.analyze_record(
        _record(
            Goal="Improve Academics",
            Metric="Core Course Grades",
            Baseline="C",
            Target="B (Benchmark)",
        )
    )
    assert "undocumented_value_annotation" in result["issue_codes"]


def test_direction_mismatch_higher_is_better() -> None:
    result = audit.analyze_record(
        _record(
            Goal="Improve Attendance",
            Metric="Attendance Rate (%)",
            Baseline="90",
            Target="85",
        )
    )
    assert "baseline_target_direction_mismatch" in result["issue_codes"]


def test_export_writes_outputs(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    results = _evaluate(workbook)
    paths = audit.export_results(
        results,
        tmp_path,
        workbook=str(workbook),
        sheet="Sheet1",
    )
    for name in (
        "audit_flags",
        "audit_summary_json",
        "audit_summary_md",
    ):
        assert Path(paths[name]).is_file(), name
    payload = json.loads(Path(paths["audit_summary_json"]).read_text(encoding="utf-8"))
    assert payload["detail_row_count"] > 0


def test_strict_exits_on_duplicate_keys(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(tmp_path))
    workbook = _workbook(tmp_path)
    results = _evaluate(workbook)
    assert results["composite_key_audit"]["duplicate_count"] > 0
    code = audit.main(["--workbook", str(workbook), "--strict"])
    assert code == 1


def test_default_cisiphyus_root_is_repo_root() -> None:
    assert audit._default_cisiphyus_root() == REPO_ROOT
    assert (audit._default_cisiphyus_root() / "run.py").is_file()


def test_cisiphyus_cmd_matches_slim_cli(monkeypatch) -> None:
    monkeypatch.delenv("AUDIT_CISPHYUS_HEADED", raising=False)
    monkeypatch.delenv("CISPHYUS_HEADED", raising=False)
    cmd = audit._cisiphyus_cmd(REPO_ROOT)
    assert cmd == [
        sys.executable,
        str(REPO_ROOT / "run.py"),
        "pull",
        "student_metrics_summary",
    ]

    monkeypatch.setenv("CISPHYUS_HEADED", "1")
    assert audit._cisiphyus_cmd(REPO_ROOT)[-1] == "--headed"

    assert audit._cisiphyus_cmd(REPO_ROOT, school_year="SY24-25")[-2:] == [
        "--school-year",
        "SY24-25",
    ]


def test_cli_resolves_workbook_via_cisiphyus_fetch(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(tmp_path))
    workbook = _workbook(tmp_path)
    fetched = audit.FetchResult(
        destination=workbook,
        triggered=True,
        succeeded=True,
        reason="forced",
    )
    with patch.object(audit, "fetch_student_metrics_workbook", return_value=fetched):
        with patch.object(audit, "export_results", return_value={}) as export_mock:
            code = audit.main([])
    assert code == 0
    export_mock.assert_called_once()


def test_progress_status_on_track_for_fixture_row(tmp_path: Path) -> None:
    records = audit.load_student_metrics_workbook(_workbook(tmp_path))
    first = records[0]
    assert audit.progress_status_for_record(first) == "off_track"


def test_progress_rollup_counts_eligible_rows(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    rollup = audit.progress_rollup(workbook)
    assert rollup["eligible_rows"] > 0
    assert "on_track" in rollup["global"]
    assert "progress_index" in rollup
