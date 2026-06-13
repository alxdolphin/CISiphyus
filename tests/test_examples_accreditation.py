"""Fixture-driven tests for the vendored accreditation flagging example."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_DIR = REPO_ROOT / "examples" / "accreditation"
if str(EXAMPLE_DIR) not in sys.path:
    sys.path.insert(0, str(EXAMPLE_DIR))

import accreditation_monitor as monitor  # noqa: E402
from accreditation_cisiphyus_fetch import (  # noqa: E402
    FetchResult,
    _cisiphyus_cmd,
    _default_cisiphyus_root,
)

FIXTURE = EXAMPLE_DIR / "fixtures" / "accreditation_monitor_minimal.xlsx"


def _evaluate(**kwargs):
    config = monitor.MonitorConfig(**kwargs)
    return monitor.evaluate_workbook(FIXTURE, config)


def test_default_cisiphyus_root_is_repo_root() -> None:
    assert _default_cisiphyus_root() == REPO_ROOT
    assert (_default_cisiphyus_root() / "run.py").is_file()


def test_cisiphyus_cmd_matches_slim_cli(monkeypatch) -> None:
    monkeypatch.delenv("ACCREDITATION_CISPHYUS_HEADED", raising=False)
    monkeypatch.delenv("CISPHYUS_HEADED", raising=False)
    cmd = _cisiphyus_cmd(REPO_ROOT)
    assert cmd == [sys.executable, str(REPO_ROOT / "run.py"), "accreditation"]

    monkeypatch.setenv("CISPHYUS_HEADED", "1")
    assert _cisiphyus_cmd(REPO_ROOT)[-1] == "--headed"


def test_fixture_flags_expected_sites() -> None:
    results = _evaluate()
    assert results["counts"]["sites_evaluated"] == 4
    assert results["counts"]["sites_excluded_inactive"] == 1

    sc_schools = {flag.school for flag in results["site_coordination"]}
    assert "Fail Reporting" in sc_schools
    assert "Fail Adjustment" in sc_schools
    assert "Pass SC" not in sc_schools

    tier1 = {flag.school: flag.flag_codes for flag in results["tier1"]}
    assert "tier1_below_period_count" in tier1["Fail Reporting"]
    assert "tier1_max_exceeds_enrollment" in tier1["Fail Adjustment"]

    cm_schools = {flag.school for flag in results["case_management"]}
    assert "Fail CM" in cm_schools


def test_export_writes_outputs(tmp_path: Path) -> None:
    results = _evaluate()
    monitor.export_results(results, tmp_path)
    for name in (
        "all_flags.csv",
        "sc_flags.csv",
        "cm_flags.csv",
        "tier1_flags.csv",
        "monitoring_summary.md",
    ):
        assert (tmp_path / name).is_file(), name
    payload = json.loads((tmp_path / "monitoring_summary.json").read_text(encoding="utf-8"))
    assert payload["counts"]["total_flagged_sites"] > 0


def test_cli_resolves_workbook_via_cisiphyus_fetch(tmp_path: Path) -> None:
    fetched = FetchResult(
        destination=FIXTURE,
        triggered=True,
        succeeded=True,
        reason="forced",
    )
    with patch.object(monitor, "fetch_accreditation_workbook", return_value=fetched):
        with patch.object(monitor, "export_results", return_value={}) as export_mock:
            code = monitor.main(["--output-dir", str(tmp_path)])
    assert code == 0
    export_mock.assert_called_once()
