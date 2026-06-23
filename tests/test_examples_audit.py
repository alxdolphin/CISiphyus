"""Tests for the student metrics audit example."""

from __future__ import annotations

import json
import shutil
import sys

import pytest
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
    assert "target_without_baseline" not in counts
    assert "baseline_without_target" not in counts
    assert counts.get("baseline_without_target_non_goal_context", 0) >= 1
    assert "baseline_without_target_no_goal_context" not in counts
    assert counts.get("duplicate_composite_key", 0) >= 1
    assert counts.get("student_client_id_mismatch", 0) >= 2
    assert counts.get("case_manager_blank", 0) >= 1


def test_no_goal_context_is_accepted_not_flagged(tmp_path: Path) -> None:
    results = _evaluate(_workbook(tmp_path))
    accepted_counts = results["accepted_exception_counts"]
    assert accepted_counts.get("baseline_without_target_no_goal_context", 0) == 1
    assert accepted_counts.get("target_without_baseline_no_goal_context", 0) == 1

    accepted_rows = {
        row["student_id"]: row for row in results.get("accepted_detail_rows") or []
    }
    assert (
        accepted_rows["S012"]["issue_codes"]
        == "baseline_without_target_no_goal_context"
    )
    assert (
        accepted_rows["S011"]["issue_codes"]
        == "target_without_baseline_no_goal_context"
    )

    flagged_rows = results.get("detail_rows") or []
    flagged_ids = {row["student_id"] for row in flagged_rows}
    assert "S012" not in flagged_ids
    assert "S011" not in flagged_ids

    s013_rows = [row for row in flagged_rows if row.get("student_id") == "S013"]
    assert len(s013_rows) == 1
    codes = set(s013_rows[0]["issue_codes"].split(";"))
    assert codes == {"baseline_without_target_non_goal_context"}


def test_export_includes_accepted_exceptions_section(tmp_path: Path) -> None:
    workbook = _workbook(tmp_path)
    results = _evaluate(workbook)
    paths = audit.export_results(
        results,
        tmp_path / "out",
        workbook=str(workbook),
        sheet="Sheet1",
    )
    markdown = Path(paths["audit_summary_md"]).read_text(encoding="utf-8")
    assert "## Accepted exceptions" in markdown
    assert "## Partial goal metrics" in markdown
    assert "## Data quality issues" in markdown
    assert "baseline_without_target_no_goal_context" in markdown

    payload = json.loads(Path(paths["audit_summary_json"]).read_text(encoding="utf-8"))
    assert payload["accepted_detail_row_count"] == 2
    assert payload["accepted_exception_counts"]["baseline_without_target_no_goal_context"] == 1
    assert payload["accepted_exception_counts"]["target_without_baseline_no_goal_context"] == 1
    assert "accepted_detail_rows" not in payload
    assert payload["datasets"]["accepted_no_goal_context"]["row_count"] == 2
    assert payload["datasets"]["partial_goal_metrics_flags"]["row_count"] == 1

    flags_csv = Path(paths["audit_flags"]).read_text(encoding="utf-8")
    assert "S012" not in flags_csv
    assert "S013" not in flags_csv

    no_goal_csv = Path(paths["accepted_no_goal_context"]).read_text(encoding="utf-8")
    assert "S012" in no_goal_csv

    partial_csv = Path(paths["partial_goal_metrics_flags"]).read_text(encoding="utf-8")
    assert "S013" in partial_csv
    assert "non_goal_context" in partial_csv


def test_reconcile_passes_on_fixture(tmp_path: Path) -> None:
    results = _evaluate(_workbook(tmp_path))
    reconciliation = results["reconciliation"]
    assert reconciliation["ok"] is True
    assert reconciliation["error_count"] == 0
    assert reconciliation["baseline_without_target_excl_supplemental"] == (
        reconciliation["baseline_flagged_primary"]
        + reconciliation["baseline_no_abc_goals"]
    )


def test_reconcile_passes_on_sy24_25_archive() -> None:
    workbook = (
        Path(__file__).resolve().parents[1]
        / "artifacts"
        / "archives"
        / "SY24-25"
        / "EOY"
        / "metrics.xlsx"
    )
    if not workbook.is_file():
        pytest.skip("SY24-25 archive not available")
    results = _evaluate(workbook)
    reconciliation = results["reconciliation"]
    assert reconciliation["ok"] is True
    assert reconciliation["baseline_without_target_raw"] == 3393
    assert reconciliation["baseline_without_target_excl_supplemental"] == 1745
    assert reconciliation["baseline_supplemental_excluded"] == 1648
    assert reconciliation["baseline_flagged_primary"] == 1105
    assert reconciliation["baseline_no_abc_goals"] == 640


def test_reconcile_passes_for_all_archived_years() -> None:
    archives = Path(__file__).resolve().parents[1] / "artifacts" / "archives"
    if not archives.is_dir():
        pytest.skip("archives not available")
    failures: list[str] = []
    for year_dir in sorted(archives.glob("SY*/EOY")):
        school_year = year_dir.parent.name
        metrics = sorted(year_dir.glob("metrics.xlsx")) + sorted(
            year_dir.glob("*StudentMetrics*.xlsx")
        )
        if not metrics:
            continue
        results = _evaluate(metrics[0])
        reconciliation = results["reconciliation"]
        if not reconciliation.get("ok"):
            failures.append(
                f"{school_year}: {reconciliation.get('errors')}"
            )
    assert not failures, "\n".join(failures)
    archives = Path("/home/adalton/Projects/CIS/tools/CISiphyus/artifacts/archives")
    workbook = archives / "SY24-25" / "EOY" / "metrics.xlsx"
    if not workbook.is_file():
        pytest.skip("SY24-25 archive not available")
    results = _evaluate(workbook)
    distribution = results["baseline_target_distribution"]
    accepted = results["accepted_exception_counts"]
    raw = distribution["baseline_without_target"]
    excluded = audit.accepted_baseline_supplemental_exclusions(accepted)
    excl = audit.baseline_without_target_excl_supplemental(distribution, accepted)
    assert excl == raw - excluded
    assert excl < raw
    assert results["issue_code_counts"]["baseline_without_target_non_goal_context"] <= excl


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
        "partial_goal_metrics_flags",
        "accepted_no_goal_context",
        "accepted_domain_tracking",
        "accepted_supplemental",
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
    with patch.object(audit, "current_school_year", return_value="SY25-26"):
        with patch.object(
            audit,
            "resolve_workbook_for_school_year",
            return_value=workbook,
        ) as resolve_mock:
            with patch.object(audit, "export_results", return_value={}) as export_mock:
                code = audit.main([])
    assert code == 0
    resolve_mock.assert_called_once_with(
        "SY25-26",
        force_fetch=False,
        fetch_destination=None,
    )
    export_mock.assert_called_once()
    assert export_mock.call_args.kwargs["school_year"] == "SY25-26"


def test_year_all_writes_cumulative_summary(tmp_path: Path, monkeypatch) -> None:
    archives = tmp_path / "archives"
    for school_year in ("SY24-25", "SY25-26"):
        metrics = archives / school_year / "EOY" / "metrics.xlsx"
        metrics.parent.mkdir(parents=True)
        shutil.copy2(_workbook(tmp_path), metrics)
    out_root = tmp_path / "audit-out"
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(out_root))
    monkeypatch.setattr(audit, "default_archives_dir", lambda: archives)
    monkeypatch.setattr(
        audit,
        "discover_audit_school_years",
        lambda **kwargs: ["SY24-25", "SY25-26"],
    )

    code = audit.main(["--year", "all"])
    assert code == 0
    summary_md = out_root / "audit_summary.md"
    summary_json = out_root / "audit_summary.json"
    assert summary_md.is_file()
    assert summary_json.is_file()
    assert not (out_root / "SY24-25" / "audit_summary.md").exists()
    text = summary_md.read_text(encoding="utf-8")
    assert "cumulative" in text.lower()
    assert "SY24-25" in text
    assert "SY25-26" in text
    payload = json.loads(summary_json.read_text(encoding="utf-8"))
    assert payload["cumulative"] is True
    assert payload["school_years"] == ["SY24-25", "SY25-26"]


def test_year_single_writes_per_year_output(tmp_path: Path, monkeypatch) -> None:
    archives = tmp_path / "archives"
    metrics = archives / "SY24-25" / "EOY" / "metrics.xlsx"
    metrics.parent.mkdir(parents=True)
    shutil.copy2(_workbook(tmp_path), metrics)
    out_root = tmp_path / "audit-out"
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(out_root))
    monkeypatch.setattr(audit, "default_archives_dir", lambda: archives)

    code = audit.main(["--year", "SY24"])
    assert code == 0
    assert (out_root / "SY24-25" / "audit_summary.md").is_file()
    assert not (out_root / "audit_summary.md").exists()


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
