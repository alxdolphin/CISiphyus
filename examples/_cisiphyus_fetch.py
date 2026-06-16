"""Shared cisiphyus pull helpers for downstream example apps."""

from __future__ import annotations

import os
import sys
from pathlib import Path

TRUE_VALUES = frozenset({"1", "true", "yes", "on"})


def flag_true(raw: str | None) -> bool:
    return (raw or "").strip().lower() in TRUE_VALUES


def env_first(*names: str) -> str:
    for name in names:
        value = (os.environ.get(name) or "").strip()
        if value:
            return value
    return ""


def default_cisiphyus_root(example_file: Path) -> Path:
    return example_file.resolve().parents[2]


def load_default_school_year(cisiphyus_root: Path) -> str | None:
    try:
        import config

        reports_path = config.REPORTS_PATH
        if not reports_path.is_file():
            reports_path = cisiphyus_root / "config" / "reports.yaml"
        programs = config.load_school_year_programs(reports_path)
        return config.default_school_year(programs)
    except (ImportError, FileNotFoundError, ValueError):
        return None


def resolve_school_year(
    *,
    cisiphyus_root: Path,
    explicit: str | None = None,
    env_names: tuple[str, ...] = ("CISDM_DEFAULT_SCHOOL_YEAR",),
) -> str | None:
    if explicit and explicit.strip():
        return explicit.strip()
    for name in env_names:
        value = (os.environ.get(name) or "").strip()
        if value:
            return value
    return load_default_school_year(cisiphyus_root)


def student_metrics_filename(school_year: str | None) -> str:
    if school_year:
        return f"{school_year}_StudentMetricsSummary.xlsx"
    return "StudentMetricsSummary.xlsx"


def cisiphyus_pull_cmd(
    cisiphyus_root: Path,
    report_id: str,
    *,
    school_year: str | None = None,
    headed_env_names: tuple[str, ...] = ("CISPHYUS_HEADED",),
) -> list[str]:
    run_py = cisiphyus_root / "run.py"
    cmd: list[str] = [sys.executable, str(run_py), "pull", report_id]
    if school_year:
        cmd.extend(["--school-year", school_year])
    if any(flag_true(os.environ.get(name)) for name in headed_env_names):
        cmd.append("--headed")
    return cmd


def cisiphyus_raw_path(
    cisiphyus_root: Path,
    report_id: str,
    *,
    school_year: str | None = None,
    default_school_year: str | None = None,
) -> Path:
    if school_year and default_school_year and school_year != default_school_year:
        return (
            cisiphyus_root
            / "artifacts"
            / "archives"
            / school_year
            / "pulls"
            / report_id
            / "raw.xlsx"
        )
    return cisiphyus_root / "artifacts" / "latest" / report_id / "raw.xlsx"
