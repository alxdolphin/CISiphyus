#!/usr/bin/env python3
"""Fetch accreditation report from CISDM via cisiphyus before monitoring."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

TRUE_VALUES = frozenset({"1", "true", "yes", "on"})
ACCREDITATION_MAX_AGE_HOURS = 24.0
CISIPHYUS_REPORT_ID = "accreditation"
DEFAULT_ACCREDITATION_FILENAME = "Accreditation_Report.xlsx"


@dataclass(frozen=True)
class FetchResult:
    destination: Path
    triggered: bool
    succeeded: bool
    reason: str


def _flag_true(raw: str | None) -> bool:
    return (raw or "").strip().lower() in TRUE_VALUES


def _default_cisiphyus_root() -> Path:
    # WHY: prototypical app lives at examples/accreditation/ inside the cisiphyus repo
    return Path(__file__).resolve().parents[2]


def _default_local_inputs_dir() -> Path:
    return Path(__file__).resolve().parent / "local_inputs"


def preferred_accreditation_workbook(
    *,
    local_inputs_dir: Path | None = None,
) -> Path:
    explicit = (os.environ.get("ACCREDITATION_WORKBOOK") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    base = Path(
        os.environ.get("ACCREDITATION_LOCAL_INPUTS_DIR", "").strip()
        or str(local_inputs_dir or _default_local_inputs_dir())
    ).expanduser().resolve()
    name = (
        os.environ.get("ACCREDITATION_WORKBOOK_FILENAME", "").strip()
        or DEFAULT_ACCREDITATION_FILENAME
    )
    return (base / name).resolve()


def workbook_freshness(workbook_path: Path | None) -> dict[str, Any]:
    max_age = float(
        os.environ.get("ACCREDITATION_MAX_AGE_HOURS", "").strip()
        or ACCREDITATION_MAX_AGE_HOURS
    )
    freshness: dict[str, Any] = {
        "max_age_hours": max_age,
        "is_stale": True,
        "workbook_exists": False,
    }
    if workbook_path is None or not workbook_path.exists():
        return freshness
    stat = workbook_path.stat()
    age_hours = (datetime.now() - datetime.fromtimestamp(stat.st_mtime)).total_seconds() / 3600
    freshness.update(
        {
            "workbook_exists": True,
            "workbook_age_hours": round(age_hours, 3),
            "workbook_mtime_utc": datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(),
            "is_stale": age_hours > max_age,
        }
    )
    return freshness


def _fetch_reason(*, destination: Path, force_fetch: bool) -> tuple[bool, str]:
    if force_fetch:
        return True, "forced"
    freshness = workbook_freshness(destination if destination.exists() else None)
    if not freshness.get("workbook_exists"):
        return True, "missing"
    if freshness.get("is_stale"):
        return True, "stale"
    return False, "fresh"


def _env_first(*names: str) -> str:
    for name in names:
        value = (os.environ.get(name) or "").strip()
        if value:
            return value
    return ""


def _cisiphyus_cmd(cisiphyus_root: Path) -> list[str]:
    # WHY: the slim run.py only accepts <report_id> and --headed;
    # auth comes from the bootstrapped chrome profile (python run.py --bootstrap)
    run_py = cisiphyus_root / "run.py"
    cmd: list[str] = [sys.executable, str(run_py), CISIPHYUS_REPORT_ID]

    headed = _flag_true(
        _env_first("ACCREDITATION_CISPHYUS_HEADED", "CISPHYUS_HEADED")
    )
    if headed:
        cmd.append("--headed")

    return cmd


def _cisiphyus_latest_raw(cisiphyus_root: Path) -> Path:
    return cisiphyus_root / "artifacts" / "latest" / CISIPHYUS_REPORT_ID / "raw.xlsx"


def _run_cisiphyus_export(
    *,
    cisiphyus_root: Path,
    run: Callable[..., Any],
) -> Path:
    run_py = cisiphyus_root / "run.py"
    if not run_py.is_file():
        raise FileNotFoundError(f"cisiphyus run.py not found at {run_py}")

    completed = run(
        _cisiphyus_cmd(cisiphyus_root),
        cwd=str(cisiphyus_root),
        check=False,
        env=os.environ.copy(),
    )
    if getattr(completed, "returncode", 1) != 0:
        raise RuntimeError(
            f"cisiphyus {CISIPHYUS_REPORT_ID} failed with exit code "
            f"{getattr(completed, 'returncode', 'unknown')}"
        )

    raw = _cisiphyus_latest_raw(cisiphyus_root)
    if not raw.is_file():
        raise FileNotFoundError(f"expected cisiphyus output missing: {raw}")
    return raw


def fetch_accreditation_workbook(
    *,
    destination: Path | None = None,
    force_fetch: bool = False,
    require_fresh: bool = False,
    run: Callable[..., Any] = subprocess.run,
) -> FetchResult:
    """Fetch accreditation report from CISDM when missing, stale, or forced."""
    target = (destination or preferred_accreditation_workbook()).resolve()
    env_force = _flag_true(os.environ.get("ACCREDITATION_FETCH_FROM_CISDM"))
    should_fetch, reason = _fetch_reason(destination=target, force_fetch=force_fetch or env_force)

    if not should_fetch:
        print(
            f"[accreditation_cisiphyus_fetch] workbook fresh at {target} "
            f"(max_age_hours={workbook_freshness(target).get('max_age_hours')}), skipping fetch",
            file=sys.stderr,
        )
        return FetchResult(
            destination=target,
            triggered=False,
            succeeded=True,
            reason="fresh",
        )

    cisiphyus_root = Path(
        _env_first("ACCREDITATION_CISPHYUS_ROOT", "CISIPHYUS_ROOT")
        or str(_default_cisiphyus_root())
    ).expanduser().resolve()

    try:
        raw = _run_cisiphyus_export(cisiphyus_root=cisiphyus_root, run=run)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(raw, target)

        if require_fresh:
            freshness = workbook_freshness(target)
            if freshness.get("is_stale"):
                age = freshness.get("workbook_age_hours")
                raise RuntimeError(
                    "stale_accreditation_workbook_after_refresh: "
                    f"path={target} workbook_age_hours={age} "
                    f"max_age_hours={freshness.get('max_age_hours')}"
                )

        print(
            f"[accreditation_cisiphyus_fetch] refreshed workbook ({reason}) -> {target}",
            file=sys.stderr,
        )
        return FetchResult(
            destination=target,
            triggered=True,
            succeeded=True,
            reason=reason,
        )
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        if require_fresh and not target.exists():
            raise
        if require_fresh and target.exists():
            freshness = workbook_freshness(target)
            if freshness.get("is_stale"):
                raise RuntimeError(
                    f"stale_accreditation_workbook_after_failed_refresh: {exc}"
                ) from exc
        print(
            f"[accreditation_cisiphyus_fetch] refresh failed ({reason}): {exc}",
            file=sys.stderr,
        )
        if not target.exists():
            raise
        return FetchResult(
            destination=target,
            triggered=True,
            succeeded=False,
            reason=reason,
        )
