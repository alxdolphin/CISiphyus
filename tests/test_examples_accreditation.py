"""Fixture-driven tests for the vendored accreditation flagging example."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

REPO_ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_DIR = REPO_ROOT / "examples" / "accreditation"
if str(EXAMPLE_DIR) not in sys.path:
    sys.path.insert(0, str(EXAMPLE_DIR))

import accreditation as monitor  # noqa: E402
from accreditation import (  # noqa: E402
    FetchResult,
    _cisiphyus_cmd,
    _default_cisiphyus_root,
)

SC_HEADERS = [
    "Organization",
    "School",
    "Site ID",
    "School District",
    "School Status",
    "# of Grading Periods",
    "# of Times Reporting to School Support Team Documented",
    "# of Times Reporting to School Leadership Documented",
    "# of Times Reporting to Organization Documented",
    "# of Times Adjustments Made to School Support Plan",
]

SC_ROWS = [
    ["CIS", "Trimester Site", "1", "District A", "Active", 3, 3, 3, 3, 1],
    ["CIS", "Fail Reporting", "2", "District B", "Active", 4, 2, 4, 4, 2],
    ["CIS", "Fail Adjustment", "3", "District C", "Active", 4, 4, 4, 4, 0],
    ["CIS", "Pass SC", "4", "District D", "Active", 4, 4, 4, 4, 1],
    ["CIS", "Inactive Site", "5", "District E", "Inactive", 4, 0, 0, 0, 0],
]

TI_HEADERS = [
    "Organization",
    "School",
    "Site ID",
    "School District",
    "School Status",
    "Total School Enrollment",
    "# of Tier I Supports",
    "Tier I - Max Served (at any one support)",
]

TI_ROWS = [
    ["CIS", "Trimester Site", "1", "District A", "Active", 500, 3, 100],
    ["CIS", "Fail Reporting", "2", "District B", "Active", 800, 2, 50],
    ["CIS", "Fail Adjustment", "3", "District C", "Active", 600, 4, 700],
    ["CIS", "Pass SC", "4", "District D", "Active", 400, 4, 100],
    ["CIS", "Pass Tier1 Only", "6", "District F", "Active", 300, 4, 50],
    ["CIS", "Inactive Site", "5", "District E", "Inactive", 100, 0, 10],
]

CM_HEADERS = [
    "Organization",
    "School",
    "Site ID",
    "School District",
    "School Status",
    "Total Case Managed Students",
    "% of CM Students with Support Plans",
    "% Students with Tier II/III Supports",
    "% of CM Students with at least one Formal Check-In",
]

CM_ROWS = [
    ["CIS", "Fail Reporting", "2", "District B", "Active", 10, 0.9, 1, 1],
    ["CIS", "Fail CM", "7", "District G", "Active", 20, 0.8, 0.7, 0.6],
    ["CIS", "No CM", "8", "District H", "Active", 0, 0, 0, 0],
    ["CIS", "Pass SC", "4", "District D", "Active", 5, 1, 1, 1],
    ["CIS", "Inactive Site", "5", "District E", "Inactive", 15, 0.5, 0.5, 0.5],
]


def build_accreditation_workbook(path: Path) -> Path:
    """Build a synthetic accreditation workbook for tests."""
    workbook = Workbook()
    workbook.remove(workbook.active)

    for sheet_name, headers, rows in (
        ("Accreditation Site Coordination", SC_HEADERS, SC_ROWS),
        ("Accreditation Tier I", TI_HEADERS, TI_ROWS),
        ("Accreditation Case Management", CM_HEADERS, CM_ROWS),
    ):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([sheet_name])
        sheet.append(headers)
        for row in rows:
            sheet.append(row)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


def _evaluate(workbook: Path, **kwargs):
    config = monitor.MonitorConfig(**kwargs)
    return monitor.evaluate_workbook(workbook, config)


def test_default_cisiphyus_root_is_repo_root() -> None:
    assert _default_cisiphyus_root() == REPO_ROOT
    assert (_default_cisiphyus_root() / "run.py").is_file()


def test_cisiphyus_cmd_matches_slim_cli(monkeypatch) -> None:
    monkeypatch.delenv("ACCREDITATION_CISPHYUS_HEADED", raising=False)
    monkeypatch.delenv("CISPHYUS_HEADED", raising=False)
    cmd = _cisiphyus_cmd(REPO_ROOT)
    assert cmd == [
        sys.executable,
        str(REPO_ROOT / "run.py"),
        "pull",
        "accreditation",
    ]

    monkeypatch.setenv("CISPHYUS_HEADED", "1")
    assert _cisiphyus_cmd(REPO_ROOT)[-1] == "--headed"


def test_fixture_flags_expected_sites(tmp_path: Path) -> None:
    workbook = build_accreditation_workbook(tmp_path / "accreditation_monitor_minimal.xlsx")
    results = _evaluate(workbook)
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


def test_export_writes_outputs(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("ACCREDITATION_OUTPUT_DIR", str(tmp_path))
    workbook = build_accreditation_workbook(tmp_path / "accreditation_monitor_minimal.xlsx")
    results = _evaluate(workbook)
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


def test_cli_resolves_workbook_via_cisiphyus_fetch(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("ACCREDITATION_OUTPUT_DIR", str(tmp_path))
    workbook = build_accreditation_workbook(tmp_path / "accreditation_monitor_minimal.xlsx")
    fetched = FetchResult(
        destination=workbook,
        triggered=True,
        succeeded=True,
        reason="forced",
    )
    with patch.object(monitor, "fetch_accreditation_workbook", return_value=fetched):
        with patch.object(monitor, "export_results", return_value={}) as export_mock:
            code = monitor.main([])
    assert code == 0
    export_mock.assert_called_once()
