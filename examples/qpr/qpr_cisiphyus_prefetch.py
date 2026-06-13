#!/usr/bin/env python3
"""cisiphyus fetch before qpr pipeline steps; age-gated refresh (see RUN_CHECKLIST)."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Sequence

TRUE_VALUES = frozenset({"1", "true", "yes", "on"})
STUDENT_METRICS_MAX_AGE_HOURS = 24.0
DEFAULT_STUDENT_METRICS_FILENAME = "SY25-26_StudentMetricsSummary.xlsx"


@dataclass(frozen=True)
class PrefetchResult:
    destination: Path | None
    triggered: bool
    succeeded: bool
    reason: str


def _flag_true(raw: str | None) -> bool:
    return (raw or "").strip().lower() in TRUE_VALUES


def _default_cisiphyus_root() -> Path:
    # WHY: prototypical app lives at examples/qpr/ inside the cisiphyus repo
    return Path(__file__).resolve().parents[2]


def _default_local_inputs_dir() -> Path:
    return Path(__file__).resolve().parent / "local_inputs"


def preferred_student_metrics_destination(
    *,
    local_inputs_dir: Path | None = None,
) -> Path:
    explicit = (os.environ.get("QPR_STUDENT_METRICS_WORKBOOK") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    base = Path(
        os.environ.get("QPR_LOCAL_INPUTS_DIR", "").strip()
        or str(local_inputs_dir or _default_local_inputs_dir())
    ).expanduser().resolve()
    name = (
        os.environ.get("QPR_STUDENT_METRICS_FILENAME", "").strip()
        or DEFAULT_STUDENT_METRICS_FILENAME
    )
    return (base / name).resolve()


def metric_workbook_freshness(workbook_path: Path | None) -> dict[str, Any]:
    freshness: dict[str, Any] = {
        "max_age_hours": STUDENT_METRICS_MAX_AGE_HOURS,
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
            "is_stale": age_hours > STUDENT_METRICS_MAX_AGE_HOURS,
        }
    )
    return freshness


_PIPELINE_METRICS_COMMANDS = frozenset(
    {
        "pipeline-preflight",
        "pipeline-dry-run",
        "pipeline-upload-only",
        "pipeline-upload-notify",
    }
)


def argv_needs_student_metrics(argv: Sequence[str]) -> bool:
    for token in argv:
        if not isinstance(token, str):
            continue
        if token in {"automate", "provision-batch", *_PIPELINE_METRICS_COMMANDS}:
            return True
    return False


def argv_requires_fresh_metrics(argv: Sequence[str]) -> bool:
    if not argv_needs_student_metrics(argv):
        return False
    for token in argv:
        if not isinstance(token, str):
            continue
        if token in {
            "pipeline-upload-only",
            "pipeline-upload-notify",
            "automate",
            "provision-batch",
        }:
            return True
    return False


def _fetch_reason(
    *,
    destination: Path,
    force_fetch: bool,
) -> tuple[bool, str]:
    if force_fetch:
        return True, "forced"
    freshness = metric_workbook_freshness(destination if destination.exists() else None)
    if not freshness.get("workbook_exists"):
        return True, "missing"
    if freshness.get("is_stale"):
        return True, "stale"
    return False, "fresh"


def _run_cisiphyus_export(
    *,
    cisiphyus_root: Path,
    run: Callable[..., Any],
) -> Path:
    run_py = cisiphyus_root / "run.py"
    if not run_py.is_file():
        raise FileNotFoundError(f"cisiphyus run.py not found at {run_py}")

    cmd = _cisiphyus_cmd(cisiphyus_root)
    completed = run(
        cmd,
        cwd=str(cisiphyus_root),
        check=False,
        env=os.environ.copy(),
    )
    if getattr(completed, "returncode", 1) != 0:
        raise RuntimeError(
            f"cisiphyus student_metrics_summary failed with exit code "
            f"{getattr(completed, 'returncode', 'unknown')}"
        )

    raw = cisiphyus_root / "artifacts" / "latest" / "student_metrics_summary" / "raw.xlsx"
    if not raw.is_file():
        raise FileNotFoundError(f"expected cisiphyus output missing: {raw}")
    return raw


def _cisiphyus_cmd(cisiphyus_root: Path) -> list[str]:
    # WHY: the slim run.py only accepts <report_id> and --headed;
    # auth comes from the bootstrapped chrome profile (python run.py --bootstrap)
    run_py = cisiphyus_root / "run.py"
    cmd: list[str] = [sys.executable, str(run_py), "student_metrics_summary"]
    headed = _flag_true(
        os.environ.get("QPR_CISPHYUS_HEADED") or os.environ.get("CISPHYUS_HEADED")
    )
    if headed:
        cmd.append("--headed")
    return cmd


def _enforce_fresh_after_refresh(
    destination: Path,
    *,
    require_fresh: bool,
    allow_stale: bool,
) -> None:
    if not require_fresh or allow_stale:
        return
    freshness = metric_workbook_freshness(destination if destination.exists() else None)
    if freshness.get("workbook_exists") and not freshness.get("is_stale"):
        return
    age = freshness.get("workbook_age_hours")
    raise RuntimeError(
        "stale_student_metrics_after_refresh: "
        f"path={destination} exists={freshness.get('workbook_exists')} "
        f"workbook_age_hours={age} max_age_hours={STUDENT_METRICS_MAX_AGE_HOURS}"
    )


def maybe_prefetch_student_metrics(
    qpr_tools_argv: Sequence[str],
    *,
    require_fresh: bool | None = None,
    run: Callable[..., Any] = subprocess.run,
) -> PrefetchResult | None:
    """refresh student metrics via cisiphyus when missing, stale (>24h), or forced."""
    if not argv_needs_student_metrics(qpr_tools_argv):
        return None

    if require_fresh is None:
        require_fresh = argv_requires_fresh_metrics(qpr_tools_argv)

    destination = preferred_student_metrics_destination()
    force_fetch = _flag_true(os.environ.get("QPR_FETCH_STUDENT_METRICS"))
    allow_stale = _flag_true(os.environ.get("QPR_ALLOW_STALE_METRICS"))
    should_fetch, reason = _fetch_reason(destination=destination, force_fetch=force_fetch)

    if not should_fetch:
        print(
            f"[qpr_cisiphyus_prefetch] metrics fresh at {destination} "
            f"(max_age_hours={STUDENT_METRICS_MAX_AGE_HOURS}), skipping fetch",
            file=sys.stderr,
        )
        resolved = destination if destination.exists() else None
        return PrefetchResult(
            destination=resolved,
            triggered=False,
            succeeded=True,
            reason="fresh",
        )

    cisiphyus_root = Path(
        os.environ.get("QPR_CISPHYUS_ROOT", "").strip()
        or os.environ.get("CISIPHYUS_ROOT", "").strip()
        or str(_default_cisiphyus_root())
    ).expanduser().resolve()

    try:
        raw = _run_cisiphyus_export(cisiphyus_root=cisiphyus_root, run=run)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(raw, destination)
        _enforce_fresh_after_refresh(
            destination,
            require_fresh=require_fresh,
            allow_stale=allow_stale,
        )
        print(
            f"[qpr_cisiphyus_prefetch] refreshed metrics ({reason}) -> {destination}",
            file=sys.stderr,
        )
        return PrefetchResult(
            destination=destination,
            triggered=True,
            succeeded=True,
            reason=reason,
        )
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        _enforce_fresh_after_refresh(
            destination,
            require_fresh=require_fresh,
            allow_stale=allow_stale,
        )
        if require_fresh and not allow_stale:
            raise
        print(f"[qpr_cisiphyus_prefetch] refresh failed ({reason}): {exc}", file=sys.stderr)
        resolved = destination if destination.exists() else None
        return PrefetchResult(
            destination=resolved,
            triggered=True,
            succeeded=False,
            reason=reason,
        )


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    try:
        require_fresh = argv_requires_fresh_metrics(argv)
        maybe_prefetch_student_metrics(argv, require_fresh=require_fresh)
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        print(f"[qpr_cisiphyus_prefetch] error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
