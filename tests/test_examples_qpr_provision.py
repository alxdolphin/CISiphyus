"""Tests for QPR workbook provisioning application."""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_DIR = REPO_ROOT / "examples" / "qpr"
if str(EXAMPLE_DIR) not in sys.path:
    sys.path.insert(0, str(EXAMPLE_DIR))

import qpr_cisiphyus_prefetch as prefetch  # noqa: E402
import qpr_provision as provision  # noqa: E402
from qpr_cisiphyus_prefetch import (  # noqa: E402
    _cisiphyus_cmd,
    _default_cisiphyus_root,
)
from qpr_tools_loader import (  # noqa: E402
    default_site_staff_list_path,
    default_template_path,
    load_qpr_tools,
    qpr_tools_path,
)

FIXTURE_METRICS = EXAMPLE_DIR / "fixtures" / "student_metrics_summary_minimal.xlsx"
FIXTURE_TEMPLATE = EXAMPLE_DIR / "fixtures" / "qpr_import_template.xlsx"
FIXTURE_DEADLINES = EXAMPLE_DIR / "fixtures" / "reporting_deadlines_minimal.xlsx"


def test_default_cisiphyus_root_is_repo_root() -> None:
    assert _default_cisiphyus_root() == REPO_ROOT
    assert (_default_cisiphyus_root() / "run.py").is_file()


def test_cisiphyus_cmd_matches_slim_cli(monkeypatch) -> None:
    monkeypatch.delenv("QPR_CISPHYUS_HEADED", raising=False)
    monkeypatch.delenv("CISPHYUS_HEADED", raising=False)
    cmd = _cisiphyus_cmd(REPO_ROOT)
    assert cmd == [
        sys.executable,
        str(REPO_ROOT / "run.py"),
        "student_metrics_summary",
    ]

    monkeypatch.setenv("CISPHYUS_HEADED", "1")
    assert _cisiphyus_cmd(REPO_ROOT)[-1] == "--headed"


def test_prefetch_skips_when_fresh(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("QPR_FETCH_STUDENT_METRICS", raising=False)
    out_dir = tmp_path / "in"
    out_dir.mkdir(parents=True)
    monkeypatch.setenv("QPR_LOCAL_INPUTS_DIR", str(out_dir))
    dest = out_dir / prefetch.DEFAULT_STUDENT_METRICS_FILENAME
    dest.write_bytes(b"fresh")
    calls: list[tuple] = []

    def fake_run(*_a: object, **_k: object) -> subprocess.CompletedProcess[str]:
        calls.append((1,))
        return subprocess.CompletedProcess([], 0, "", "")

    result = prefetch.maybe_prefetch_student_metrics(["provision-batch"], run=fake_run)
    assert result is not None
    assert result.triggered is False
    assert result.reason == "fresh"
    assert calls == []


def test_prefetch_refreshes_stale_workbook(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("QPR_FETCH_STUDENT_METRICS", raising=False)
    monkeypatch.setenv("QPR_CISPHYUS_ROOT", str(tmp_path))
    out_dir = tmp_path / "in"
    out_dir.mkdir(parents=True)
    monkeypatch.setenv("QPR_LOCAL_INPUTS_DIR", str(out_dir))
    dest = out_dir / prefetch.DEFAULT_STUDENT_METRICS_FILENAME
    dest.write_bytes(b"old")
    old = time.time() - (30 * 3600)
    import os

    os.utime(dest, (old, old))
    (tmp_path / "run.py").write_text("# stub\n", encoding="utf-8")
    raw_parent = tmp_path / "artifacts" / "latest" / "student_metrics_summary"
    raw_parent.mkdir(parents=True)
    (raw_parent / "raw.xlsx").write_bytes(b"new")

    def fake_run(*_a: object, **_k: object) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess([], 0, "", "")

    result = prefetch.maybe_prefetch_student_metrics(
        ["provision-batch"],
        require_fresh=True,
        run=fake_run,
    )
    assert result is not None
    assert result.triggered is True
    assert result.succeeded is True
    assert dest.read_bytes() == b"new"


def test_qpr_tools_loader_finds_bundled_module() -> None:
    path = qpr_tools_path()
    assert path.is_file()
    assert path == EXAMPLE_DIR / "qpr_tools.py"
    module = load_qpr_tools()
    assert hasattr(module, "provision_all_site_templates")


def test_default_fixture_paths_exist() -> None:
    assert default_template_path().is_file()
    assert default_site_staff_list_path().is_file()


def test_resolve_paths_use_bundled_fixtures() -> None:
    assert provision.resolve_metric_workbook(FIXTURE_METRICS) == FIXTURE_METRICS.resolve()
    assert provision.resolve_template_path(FIXTURE_TEMPLATE) == FIXTURE_TEMPLATE.resolve()


def test_provision_generates_site_workbooks(tmp_path: Path) -> None:
    code = provision.main(
        [
            str(tmp_path),
            "--grading-period",
            "2.0",
            "--metric-workbook",
            str(FIXTURE_METRICS),
            "--template",
            str(FIXTURE_TEMPLATE),
            "--deadlines",
            str(FIXTURE_DEADLINES),
            "--no-site-staff-filter",
            "--json",
        ]
    )
    assert code == 0
    outputs = list(tmp_path.rglob("*_QPR.xlsx"))
    assert len(outputs) >= 2
    names = {path.name for path in outputs}
    assert any("Lincoln" in name for name in names)
    assert any("Washington" in name for name in names)


def test_provision_single_school(tmp_path: Path) -> None:
    code = provision.main(
        [
            str(tmp_path),
            "--grading-period",
            "2.0",
            "--school-name",
            "Lincoln HS",
            "--metric-workbook",
            str(FIXTURE_METRICS),
            "--template",
            str(FIXTURE_TEMPLATE),
            "--deadlines",
            str(FIXTURE_DEADLINES),
            "--no-site-staff-filter",
        ]
    )
    assert code == 0
    assert list(tmp_path.glob("*.xlsx")), "expected a provisioned workbook"


def test_provision_requires_bundled_qpr_tools(monkeypatch, tmp_path: Path) -> None:
    missing = tmp_path / "qpr_tools.py"
    with patch.object(provision, "load_qpr_tools", side_effect=FileNotFoundError("missing bundled qpr_tools")):
        code = provision.main(["/tmp/out", "--grading-period", "2.0", "--no-site-staff-filter"])
    assert code == 1
