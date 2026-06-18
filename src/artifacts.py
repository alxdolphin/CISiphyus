from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


@dataclass(frozen=True)
class RunResult:
    report_id: str
    status: str
    started_at: str
    finished_at: str
    raw_path: str | None
    error: str | None = None
    school_year: str | None = None
    enrollment_program_id: int | None = None


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")


def write_latest_result(
    *,
    report_id: str,
    status: str,
    started_at: str,
    finished_at: str,
    raw_path: Path | None,
    error: str | None,
    latest_dir: Path,
    school_year: str | None = None,
    enrollment_program_id: int | None = None,
) -> RunResult:
    resolved_raw = str(raw_path.resolve()) if raw_path and raw_path.exists() else None
    result = RunResult(
        report_id=report_id,
        status=status,
        started_at=started_at,
        finished_at=finished_at,
        raw_path=resolved_raw,
        error=error,
        school_year=school_year,
        enrollment_program_id=enrollment_program_id,
    )
    latest_dir.mkdir(parents=True, exist_ok=True)
    write_json(latest_dir / "result.json", asdict(result))
    return result


def write_failure_diag(latest_dir: Path, payload: dict[str, Any]) -> Path:
    latest_dir.mkdir(parents=True, exist_ok=True)
    diag_path = latest_dir / "diag.json"
    write_json(diag_path, payload)
    return diag_path
