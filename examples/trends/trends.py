#!/usr/bin/env python3
"""
Longitudinal trend tracker for accreditation and student metrics audit snapshots.

Run:
  cisiphyus trend SY25-26
  cisiphyus trend SY25-26 SY24-25

Discovers reporting periods from QPR output (and existing snapshots), pulls CISDM
exports via cisiphyus when needed, caches them under artifacts/archives/, captures
snapshots, then compares consecutive quarters.

Outputs: artifacts/snapshots/, artifacts/trends/{school-year}/{Q1_vs_Q2}/, ...
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import sys
from collections import Counter
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
for _example in ("accreditation", "audit"):
    _path = REPO_ROOT / "examples" / _example
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

import accreditation  # noqa: E402
import audit  # noqa: E402

PCT_REGRESSION_FIELDS = (
    "pct_formal_checkins",
    "pct_support_plans",
    "pct_tier23_supports",
)
REPORTING_COUNT_FIELDS = (
    "reporting_leadership",
    "reporting_organization",
    "reporting_support_team",
)
PERIOD_DIR_RE = re.compile(r"^Q([1-4])$", re.IGNORECASE)
ACCREDITATION_CANDIDATES = (
    "Accreditation_Report.xlsx",
    "accreditation.xlsx",
)
METRICS_GLOBS = (
    "*StudentMetrics*.xlsx",
    "*student_metrics*.xlsx",
    "*MetricsSummary*.xlsx",
    "metrics.xlsx",
)
# WHY: school years without QPR quarter folders still get a metrics archive slot
FALLBACK_PERIOD = "EOY"

MOVEMENT_FIELDS = (
    "stream",
    "entity_key",
    "entity_label",
    "metric",
    "baseline",
    "current",
    "delta",
    "movement_type",
    "is_regression",
)

GOAL_PROGRESS_STATUS_ORDER = audit.PROGRESS_STATUS_ORDER
PROGRESS_STATUS_RANK = {
    "on_track": 2,
    "off_track": 1,
    "no_progress_data": 0,
    "indeterminate": 0,
}


def default_inputs_dir() -> Path:
    explicit = (os.environ.get("TREND_INPUTS_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return REPO_ROOT / "artifacts" / "archives"


def default_qpr_dir() -> Path:
    explicit = (os.environ.get("QPR_OUTPUT_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return REPO_ROOT / "artifacts" / "qpr"


def normalize_period(name: str) -> str | None:
    stripped = name.strip()
    if stripped.upper() == FALLBACK_PERIOD:
        return FALLBACK_PERIOD
    match = PERIOD_DIR_RE.match(stripped)
    if match:
        return f"Q{match.group(1)}"
    return None


def period_sort_key(label: str) -> tuple[int, str]:
    if label.upper() == FALLBACK_PERIOD:
        return (5, FALLBACK_PERIOD)
    match = PERIOD_DIR_RE.match(label)
    if match:
        return (int(match.group(1)), label.upper())
    return (999, label)


def discover_available_periods(
    school_year: str,
    *,
    qpr_dir: Path,
    snapshots_dir: Path,
    inputs_dir: Path | None = None,
) -> list[str]:
    periods: set[str] = set()
    qpr_root = qpr_dir / school_year
    if qpr_root.is_dir():
        for child in qpr_root.iterdir():
            if child.is_dir():
                label = normalize_period(child.name)
                if label:
                    periods.add(label)
    if inputs_dir is not None:
        archive_root = inputs_dir / school_year
        if archive_root.is_dir():
            for child in archive_root.iterdir():
                if child.is_dir():
                    label = normalize_period(child.name)
                    if label:
                        periods.add(label)
    for entry in load_manifest(snapshots_dir).get("snapshots", []):
        if entry.get("school_year") == school_year:
            label = normalize_period(str(entry.get("period", "")))
            if label:
                periods.add(label)
    return sorted(periods, key=period_sort_key)


def _load_config_programs() -> dict[str, int]:
    src_dir = REPO_ROOT / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))
    import config

    return config.load_school_year_programs(config.REPORTS_PATH)


def current_school_year() -> str | None:
    try:
        programs = _load_config_programs()
    except (FileNotFoundError, ValueError, OSError):
        return default_school_year()
    import config

    return config.default_school_year(programs) or default_school_year()


def school_year_metrics_cached(school_year: str, *, inputs_dir: Path) -> Path | None:
    fallback = inputs_dir / school_year / FALLBACK_PERIOD
    met = _find_workbook(fallback, (), METRICS_GLOBS)
    if met is not None:
        return met

    cisiphyus_pull = (
        inputs_dir
        / school_year
        / "pulls"
        / "student_metrics_summary"
        / "raw.xlsx"
    )
    if cisiphyus_pull.is_file():
        return cisiphyus_pull.resolve()

    year_dir = inputs_dir / school_year
    if year_dir.is_dir():
        for period_dir in sorted(year_dir.iterdir()):
            if not period_dir.is_dir():
                continue
            met = _find_workbook(period_dir, (), METRICS_GLOBS)
            if met is not None:
                return met
    return None


def pull_school_year_backfill(
    school_year: str,
    *,
    inputs_dir: Path,
    force_fetch: bool,
    include_accreditation: bool | None = None,
) -> int:
    cached = school_year_metrics_cached(school_year, inputs_dir=inputs_dir)
    if cached is not None and not force_fetch:
        print(f"skip {school_year}: metrics cached at {cached}", file=sys.stderr)
        return 0

    dest_dir = inputs_dir / school_year / FALLBACK_PERIOD
    dest_dir.mkdir(parents=True, exist_ok=True)
    metrics_dest = dest_dir / "metrics.xlsx"

    try:
        if school_year == current_school_year():
            acc_path, met_path = pull_workbooks(
                school_year=school_year,
                force_fetch=True,
                include_accreditation=include_accreditation,
            )
            if acc_path is not None:
                shutil.copy2(acc_path, dest_dir / "accreditation.xlsx")
            shutil.copy2(met_path, metrics_dest)
        else:
            print(
                f"[trend] pulling student metrics for {school_year} (historical)",
                file=sys.stderr,
            )
            fetch = audit.fetch_student_metrics_workbook(
                force_fetch=True,
                require_fresh=True,
                school_year=school_year,
            )
            if not fetch.succeeded:
                raise FileNotFoundError(
                    f"student metrics workbook unavailable for {school_year!r}"
                )
            shutil.copy2(fetch.destination, metrics_dest)
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        print(f"error: CISDM pull failed for {school_year}: {exc}", file=sys.stderr)
        return 1

    print(f"archived: {school_year}/{FALLBACK_PERIOD} -> {dest_dir}", file=sys.stderr)
    return 0


def resolve_eoy_workbooks(
    school_year: str,
    *,
    inputs_dir: Path,
) -> tuple[Path | None, Path]:
    period_dir = inputs_dir / school_year / FALLBACK_PERIOD
    met = _find_workbook(period_dir, (), METRICS_GLOBS)
    if met is None:
        raise FileNotFoundError(f"no metrics workbook under {period_dir}")
    acc = _find_workbook(period_dir, ACCREDITATION_CANDIDATES, ("*accreditation*.xlsx",))
    return acc, met


def run_eoy_backfill_and_snapshot(
    school_year: str,
    *,
    inputs_dir: Path,
    snapshots_dir: Path,
    force_fetch: bool,
    force: bool,
    include_accreditation: bool | None = None,
) -> tuple[int, str]:
    cached = school_year_metrics_cached(school_year, inputs_dir=inputs_dir)
    slot = snapshot_slot(snapshots_dir, school_year, FALLBACK_PERIOD)
    if cached is None or force_fetch:
        code = pull_school_year_backfill(
            school_year,
            inputs_dir=inputs_dir,
            force_fetch=force_fetch,
            include_accreditation=include_accreditation,
        )
        if code != 0:
            return 1, "failed"
    elif slot.is_dir() and not force:
        print(f"status {school_year}: snapshotted (EOY exists)", file=sys.stderr)
        return 0, "snapshotted"

    if slot.is_dir() and not force:
        print(f"status {school_year}: snapshotted (EOY exists)", file=sys.stderr)
        return 0, "snapshotted"

    try:
        acc_wb, met_wb = resolve_eoy_workbooks(school_year, inputs_dir=inputs_dir)
        capture_snapshot(
            school_year=school_year,
            period=FALLBACK_PERIOD,
            snapshots_dir=snapshots_dir,
            accreditation_workbook=acc_wb,
            metrics_workbook=met_wb,
            metrics_only=acc_wb is None,
            force=force,
        )
    except (FileExistsError, FileNotFoundError, OSError, RuntimeError, ValueError) as exc:
        print(f"error: {school_year}/{FALLBACK_PERIOD}: {exc}", file=sys.stderr)
        return 1, "failed"

    print(f"captured: {school_year}/{FALLBACK_PERIOD}", file=sys.stderr)
    return 0, "snapshotted"


def discover_school_years(
    *,
    qpr_dir: Path,
    snapshots_dir: Path,
    inputs_dir: Path | None = None,
    include_config_years: bool = False,
) -> list[str]:
    years: set[str] = set()
    if qpr_dir.is_dir():
        for child in qpr_dir.iterdir():
            if child.is_dir() and child.name.upper().startswith("SY"):
                years.add(child.name)
    archives_dir = inputs_dir or default_inputs_dir()
    if archives_dir.is_dir():
        for child in archives_dir.iterdir():
            if child.is_dir() and child.name.upper().startswith("SY"):
                years.add(child.name)
    for entry in load_manifest(snapshots_dir).get("snapshots", []):
        school_year = entry.get("school_year")
        if isinstance(school_year, str) and school_year.upper().startswith("SY"):
            years.add(school_year)
    if include_config_years:
        try:
            years.update(_load_config_programs().keys())
        except (FileNotFoundError, ValueError, OSError):
            pass
    import config

    return sorted(years, key=config._school_year_sort_key)


def _find_workbook(period_dir: Path, names: tuple[str, ...], globs: tuple[str, ...]) -> Path | None:
    for name in names:
        path = period_dir / name
        if path.is_file():
            return path.resolve()
    for pattern in globs:
        matches = sorted(period_dir.glob(pattern))
        if matches:
            return matches[0].resolve()
    return None


def resolve_period_workbooks(
    school_year: str,
    period: str,
    *,
    inputs_dir: Path,
) -> tuple[Path | None, Path] | None:
    period_dir = inputs_dir / school_year / period
    if not period_dir.is_dir():
        return None
    acc = _find_workbook(period_dir, ACCREDITATION_CANDIDATES, ("*accreditation*.xlsx",))
    met = _find_workbook(period_dir, (), METRICS_GLOBS)
    if met is None:
        return None
    return acc, met


def archive_workbooks(
    school_year: str,
    period: str,
    acc_wb: Path | None,
    met_wb: Path,
    *,
    inputs_dir: Path,
) -> tuple[Path | None, Path]:
    dest_dir = inputs_dir / school_year / period
    dest_dir.mkdir(parents=True, exist_ok=True)
    met_dest = dest_dir / "metrics.xlsx"
    shutil.copy2(met_wb, met_dest)
    if acc_wb is None:
        return None, met_dest.resolve()
    acc_dest = dest_dir / acc_wb.name
    shutil.copy2(acc_wb, acc_dest)
    return acc_dest.resolve(), met_dest.resolve()


def accreditation_fetch_requested(*, explicit: bool | None = None) -> bool:
    if explicit is not None:
        return explicit
    return accreditation._flag_true(os.environ.get("ACCREDITATION_FETCH_FROM_CISDM"))


def pull_workbooks(
    *,
    school_year: str,
    force_fetch: bool,
    include_accreditation: bool | None = None,
) -> tuple[Path | None, Path]:
    include_accreditation = accreditation_fetch_requested(explicit=include_accreditation)
    if include_accreditation:
        print(
            f"[trend] pulling accreditation and student metrics from CISDM ({school_year})",
            file=sys.stderr,
        )
    else:
        print(
            f"[trend] pulling student metrics from CISDM ({school_year}); "
            "skip accreditation (ACCREDITATION_FETCH_FROM_CISDM=1 or --accreditation)",
            file=sys.stderr,
        )
    return _resolve_workbooks(
        accreditation_workbook=None,
        metrics_workbook=None,
        school_year=school_year,
        force_fetch=force_fetch,
        include_accreditation=include_accreditation,
    )


def workbooks_for_period(
    school_year: str,
    period: str,
    *,
    inputs_dir: Path,
    pulled: tuple[Path | None, Path] | None,
    latest_pull_period: str | None,
    force_fetch: bool,
) -> tuple[Path | None, Path] | None:
    cached = resolve_period_workbooks(school_year, period, inputs_dir=inputs_dir)
    if cached and not force_fetch:
        return cached
    if latest_pull_period is None or period != latest_pull_period or pulled is None:
        return None
    return archive_workbooks(
        school_year,
        period,
        pulled[0],
        pulled[1],
        inputs_dir=inputs_dir,
    )


def run_trend_school_year(
    school_year: str,
    *,
    snapshots_dir: Path,
    output_dir: Path,
    inputs_dir: Path,
    qpr_dir: Path,
    force: bool = False,
    force_fetch: bool = False,
    skip_compare: bool = False,
    regression_threshold: float = 0.01,
    skip_empty_periods: bool = False,
    include_accreditation: bool | None = None,
) -> tuple[int, str]:
    manifest = load_manifest(snapshots_dir)
    snapshotted = {
        str(entry.get("period"))
        for entry in manifest.get("snapshots", [])
        if entry.get("school_year") == school_year
    }
    qpr_periods = discover_available_periods(
        school_year,
        qpr_dir=qpr_dir,
        snapshots_dir=snapshots_dir,
        inputs_dir=inputs_dir,
    )
    if not qpr_periods:
        try:
            programs = _load_config_programs()
        except (FileNotFoundError, ValueError, OSError):
            programs = {}
        if school_year not in programs:
            message = (
                f"skip {school_year}: no reporting periods and not in school_year_programs"
            )
            if skip_empty_periods:
                print(message, file=sys.stderr)
                return 0, "skipped"
            print(
                f"error: no periods found for {school_year!r} under {qpr_dir / school_year}",
                file=sys.stderr,
            )
            return 1, "failed"
        return run_eoy_backfill_and_snapshot(
            school_year,
            inputs_dir=inputs_dir,
            snapshots_dir=snapshots_dir,
            force_fetch=force_fetch,
            force=force,
            include_accreditation=include_accreditation,
        )

    need_capture = [
        period
        for period in qpr_periods
        if period not in snapshotted or force
    ]
    latest_pull_period: str | None = None
    pulled: tuple[Path | None, Path] | None = None
    if need_capture:
        missing_cache = [
            period
            for period in need_capture
            if resolve_period_workbooks(school_year, period, inputs_dir=inputs_dir) is None
        ]
        if missing_cache or force_fetch:
            latest_pull_period = max(missing_cache or need_capture, key=period_sort_key)
            try:
                pulled = pull_workbooks(
                    school_year=school_year,
                    force_fetch=force_fetch or bool(missing_cache),
                    include_accreditation=include_accreditation,
                )
            except (FileNotFoundError, OSError, RuntimeError) as exc:
                print(f"error: CISDM pull failed: {exc}", file=sys.stderr)
                return 1, "failed"

    captured: list[str] = []
    for period in qpr_periods:
        slot = snapshot_slot(snapshots_dir, school_year, period)
        if period not in need_capture:
            if slot.is_dir():
                captured.append(period)
            continue
        if slot.exists() and not force:
            print(f"skip (exists): {school_year}/{period}")
            captured.append(period)
            continue

        workbooks = workbooks_for_period(
            school_year,
            period,
            inputs_dir=inputs_dir,
            pulled=pulled,
            latest_pull_period=latest_pull_period,
            force_fetch=force_fetch,
        )
        if workbooks is None:
            print(
                f"skip {school_year}/{period}: no cached pull for this period "
                f"(run cisiphyus trend at end of {period} to capture it)",
                file=sys.stderr,
            )
            if slot.is_dir():
                captured.append(period)
            continue
        acc_wb, met_wb = workbooks
        try:
            capture_snapshot(
                school_year=school_year,
                period=period,
                snapshots_dir=snapshots_dir,
                accreditation_workbook=acc_wb,
                metrics_workbook=met_wb,
                force_fetch=False,
                force=force,
                metrics_only=acc_wb is None,
            )
        except (FileExistsError, FileNotFoundError, OSError, RuntimeError, ValueError) as exc:
            print(f"error: {school_year}/{period}: {exc}", file=sys.stderr)
            return 1, "failed"
        print(f"captured: {school_year}/{period}")
        captured.append(period)

    if not captured:
        print(f"error: no snapshots captured for {school_year!r}", file=sys.stderr)
        return 1, "failed"

    if skip_compare or len(captured) < 2:
        return 0, "snapshotted"

    ordered = sorted(set(captured), key=period_sort_key)
    exit_code = 0
    compared = False
    for baseline, current in zip(ordered, ordered[1:]):
        compared = True
        pair_dir = output_dir / school_year / f"{baseline}_vs_{current}"
        try:
            payload = compare_snapshots(
                snapshots_dir=snapshots_dir,
                school_year=school_year,
                baseline_period=baseline,
                current_period=current,
                regression_threshold=regression_threshold,
            )
            paths = export_trend_results(payload, pair_dir)
        except FileNotFoundError as exc:
            print(f"error: {exc}", file=sys.stderr)
            exit_code = 1
            continue
        print(f"compare: {school_year} {baseline} -> {current}")
        print(f"  movements: {payload['movement_count']}")
        print(f"  regressions: {payload['regression_count']}")
        print(f"  output: {pair_dir}")
        for key, path in paths.items():
            print(f"    {key}: {path}")
    if compared:
        return exit_code, "compared" if exit_code == 0 else "failed"
    return 0, "snapshotted"


def default_snapshots_dir() -> Path:
    explicit = (os.environ.get("TREND_SNAPSHOTS_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return REPO_ROOT / "artifacts" / "snapshots"


def default_output_dir() -> Path:
    explicit = (os.environ.get("TREND_OUTPUT_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return REPO_ROOT / "artifacts" / "trends"


def default_school_year() -> str | None:
    raw = (os.environ.get("TREND_SCHOOL_YEAR") or "").strip()
    return raw or None


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def snapshot_slot(snapshots_dir: Path, school_year: str, period: str) -> Path:
    return snapshots_dir / school_year / period


def discover_eoy_snapshot_years(snapshots_dir: Path) -> list[str]:
    """School years that have a captured EOY snapshot directory."""
    if not snapshots_dir.is_dir():
        return []
    years: list[str] = []
    for entry in snapshots_dir.iterdir():
        if not entry.is_dir():
            continue
        if (entry / FALLBACK_PERIOD).is_dir():
            years.append(entry.name)
    import config

    return sorted(years, key=config._school_year_sort_key)


def build_school_rollup(detail_rows: list[dict[str, Any]]) -> dict[str, Any]:
    schools: dict[str, dict[str, Any]] = {}
    for row in detail_rows:
        school = (row.get("school") or "").strip() or "(unknown)"
        bucket = schools.setdefault(
            school,
            {"school": school, "flagged_rows": 0, "issue_code_counts": {}},
        )
        bucket["flagged_rows"] += 1
        codes = (row.get("issue_codes") or "").split(";")
        counts: Counter[str] = Counter(bucket["issue_code_counts"])
        for code in codes:
            code = code.strip()
            if code:
                counts[code] += 1
        bucket["issue_code_counts"] = dict(counts)
    return {"schools": schools}


def _resolve_workbooks(
    *,
    accreditation_workbook: Path | None,
    metrics_workbook: Path | None,
    school_year: str | None = None,
    force_fetch: bool,
    include_accreditation: bool = True,
) -> tuple[Path | None, Path]:
    if metrics_workbook is not None:
        met_path = metrics_workbook.resolve()
        if accreditation_workbook is not None:
            return accreditation_workbook.resolve(), met_path
        if not include_accreditation:
            return None, met_path

    acc_path = accreditation_workbook
    met_path = metrics_workbook
    if include_accreditation and acc_path is None:
        destination = accreditation.preferred_accreditation_workbook()
        fetch = accreditation.fetch_accreditation_workbook(
            destination=destination,
            force_fetch=force_fetch,
            require_fresh=True,
        )
        if not fetch.succeeded:
            raise FileNotFoundError(f"accreditation workbook unavailable at {destination}")
        acc_path = fetch.destination
    if met_path is None:
        destination = audit.preferred_student_metrics_workbook(school_year=school_year)
        fetch = audit.fetch_student_metrics_workbook(
            destination=destination,
            force_fetch=force_fetch,
            require_fresh=True,
            school_year=school_year,
        )
        if not fetch.succeeded:
            raise FileNotFoundError(f"student metrics workbook unavailable at {destination}")
        met_path = fetch.destination
    return (
        acc_path.resolve() if acc_path is not None else None,
        met_path.resolve(),
    )


def load_manifest(snapshots_dir: Path) -> dict[str, Any]:
    path = snapshots_dir / "manifest.json"
    if not path.is_file():
        return {"snapshots": []}
    return _load_json(path)


def save_manifest(snapshots_dir: Path, manifest: dict[str, Any]) -> None:
    _write_json(snapshots_dir / "manifest.json", manifest)


def capture_snapshot(
    *,
    school_year: str,
    period: str,
    snapshots_dir: Path,
    accreditation_workbook: Path | None = None,
    metrics_workbook: Path | None = None,
    force_fetch: bool = False,
    force: bool = False,
    metrics_only: bool = False,
    monitor_config: accreditation.MonitorConfig | None = None,
    audit_config: audit.AuditConfig | None = None,
) -> Path:
    slot = snapshot_slot(snapshots_dir, school_year, period)
    if slot.exists() and not force:
        raise FileExistsError(f"snapshot already exists: {slot} (use --force to replace)")

    if metrics_only:
        if metrics_workbook is None:
            raise ValueError("metrics_only snapshot requires metrics_workbook")
        met_wb = metrics_workbook.resolve()
        acc_wb = None
    else:
        acc_wb, met_wb = _resolve_workbooks(
            accreditation_workbook=accreditation_workbook,
            metrics_workbook=metrics_workbook,
            school_year=school_year,
            force_fetch=force_fetch,
        )
    monitor_cfg = monitor_config or accreditation.MonitorConfig()
    audit_cfg = audit_config or audit.AuditConfig(school_year=school_year)

    if acc_wb is not None:
        acc_results = accreditation.evaluate_workbook(acc_wb, monitor_cfg)
        acc_payload = accreditation.results_to_json_payload(acc_results)
    else:
        acc_payload = {"all_flags": [], "metrics_only": True}
    met_results = audit.evaluate_workbook(met_wb, audit_cfg)
    met_payload = audit.results_to_json_payload(
        met_results,
        workbook=str(met_wb),
        sheet=audit_cfg.sheet_name,
    )
    school_rollup = build_school_rollup(met_results.get("detail_rows") or [])
    grading_period_stats = audit.grading_period_fill_stats(met_wb, audit_cfg.sheet_name)
    progress_payload = met_results["progress_rollup"]

    if slot.exists() and force:
        shutil.rmtree(slot)
    acc_dir = slot / "accreditation"
    met_dir = slot / "metrics"
    acc_dir.mkdir(parents=True, exist_ok=True)
    met_dir.mkdir(parents=True, exist_ok=True)

    acc_json = acc_dir / "monitoring_summary.json"
    met_json = met_dir / "audit_summary.json"
    rollup_json = met_dir / "school_rollup.json"
    grading_json = met_dir / "grading_period_stats.json"
    progress_json = met_dir / "progress_rollup.json"
    _write_json(acc_json, acc_payload)
    _write_json(met_json, met_payload)
    _write_json(rollup_json, school_rollup)
    _write_json(grading_json, grading_period_stats)
    _write_json(progress_json, progress_payload)
    if acc_wb is not None:
        accreditation.write_csv(acc_dir / "all_flags.csv", acc_results["all_flags"])

    captured_at = _utc_now_iso()
    source_workbooks: dict[str, str] = {"metrics": str(met_wb)}
    if acc_wb is not None:
        source_workbooks["accreditation"] = str(acc_wb)
    meta = {
        "school_year": school_year,
        "period": period,
        "captured_at": captured_at,
        "metrics_only": metrics_only,
        "source_workbooks": source_workbooks,
        "accreditation_config": asdict(monitor_cfg),
        "audit_config": asdict(audit_cfg),
        "content_hashes": {
            "accreditation/monitoring_summary.json": _file_hash(acc_json),
            "metrics/audit_summary.json": _file_hash(met_json),
            "metrics/school_rollup.json": _file_hash(rollup_json),
            "metrics/grading_period_stats.json": _file_hash(grading_json),
            "metrics/progress_rollup.json": _file_hash(progress_json),
        },
    }
    _write_json(slot / "snapshot.json", meta)

    manifest = load_manifest(snapshots_dir)
    rel = f"{school_year}/{period}"
    entries = [e for e in manifest.get("snapshots", []) if e.get("path") != rel]
    entries.append(
        {
            "school_year": school_year,
            "period": period,
            "captured_at": captured_at,
            "path": rel,
        }
    )
    entries.sort(key=lambda e: (e.get("school_year", ""), e.get("period", "")))
    manifest["snapshots"] = entries
    save_manifest(snapshots_dir, manifest)
    return slot


def _site_index(payload: dict[str, Any]) -> dict[str, dict[str, Any]]:
    index: dict[str, dict[str, Any]] = {}
    for row in payload.get("all_flags") or []:
        site_id = str(row.get("site_id") or row.get("school") or "")
        if site_id:
            index[site_id] = row
    return index


def _movement_row(
    *,
    stream: str,
    entity_key: str,
    entity_label: str,
    metric: str,
    baseline: Any,
    current: Any,
    movement_type: str,
    is_regression: bool,
) -> dict[str, Any]:
    delta: Any = ""
    if isinstance(baseline, (int, float)) and isinstance(current, (int, float)):
        delta = current - baseline
    elif isinstance(baseline, int) and isinstance(current, int):
        delta = current - baseline
    return {
        "stream": stream,
        "entity_key": entity_key,
        "entity_label": entity_label,
        "metric": metric,
        "baseline": baseline,
        "current": current,
        "delta": delta,
        "movement_type": movement_type,
        "is_regression": is_regression,
    }


def _compare_accreditation(
    baseline: dict[str, Any],
    current: dict[str, Any],
    *,
    regression_threshold: float,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    base_sites = _site_index(baseline)
    curr_sites = _site_index(current)
    all_ids = sorted(set(base_sites) | set(curr_sites))

    for site_id in all_ids:
        base = base_sites.get(site_id)
        curr = curr_sites.get(site_id)
        label = (curr or base or {}).get("school") or site_id

        if base is None:
            rows.append(
                _movement_row(
                    stream="accreditation",
                    entity_key=site_id,
                    entity_label=label,
                    metric="site_presence",
                    baseline="",
                    current="present",
                    movement_type="site_appeared",
                    is_regression=False,
                )
            )
            continue
        if curr is None:
            rows.append(
                _movement_row(
                    stream="accreditation",
                    entity_key=site_id,
                    entity_label=label,
                    metric="site_presence",
                    baseline="present",
                    current="",
                    movement_type="site_removed",
                    is_regression=False,
                )
            )
            continue

        base_flags = set(base.get("flag_codes") or [])
        curr_flags = set(curr.get("flag_codes") or [])
        for code in sorted(curr_flags - base_flags):
            rows.append(
                _movement_row(
                    stream="accreditation",
                    entity_key=site_id,
                    entity_label=label,
                    metric="flag_code",
                    baseline="",
                    current=code,
                    movement_type="flag_added",
                    is_regression=True,
                )
            )
        for code in sorted(base_flags - curr_flags):
            rows.append(
                _movement_row(
                    stream="accreditation",
                    entity_key=site_id,
                    entity_label=label,
                    metric="flag_code",
                    baseline=code,
                    current="",
                    movement_type="flag_resolved",
                    is_regression=False,
                )
            )

        base_details = base.get("details") or {}
        curr_details = curr.get("details") or {}
        numeric_keys = sorted(
            set(base_details) | set(curr_details),
            key=str,
        )
        for key in numeric_keys:
            bv = base_details.get(key)
            cv = curr_details.get(key)
            if not isinstance(bv, (int, float)) or not isinstance(cv, (int, float)):
                continue
            if bv == cv:
                continue
            is_regression = False
            if key in PCT_REGRESSION_FIELDS and (bv - cv) >= regression_threshold:
                is_regression = True
            if key in REPORTING_COUNT_FIELDS and cv < bv:
                is_regression = True
            rows.append(
                _movement_row(
                    stream="accreditation",
                    entity_key=site_id,
                    entity_label=label,
                    metric=key,
                    baseline=bv,
                    current=cv,
                    movement_type="metric_delta",
                    is_regression=is_regression,
                )
            )
    return rows


def _compare_metrics(
    baseline_summary: dict[str, Any],
    current_summary: dict[str, Any],
    baseline_rollup: dict[str, Any],
    current_rollup: dict[str, Any],
    *,
    baseline_grading_period_stats: dict[str, int] | None = None,
    current_grading_period_stats: dict[str, int] | None = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    base_global = baseline_summary.get("issue_code_counts") or {}
    curr_global = current_summary.get("issue_code_counts") or {}
    for code in sorted(set(base_global) | set(curr_global)):
        bv = int(base_global.get(code, 0))
        cv = int(curr_global.get(code, 0))
        if bv == cv:
            continue
        rows.append(
            _movement_row(
                stream="metrics",
                entity_key="global",
                entity_label="global",
                metric=code,
                baseline=bv,
                current=cv,
                movement_type="global_issue_delta",
                is_regression=cv > bv,
            )
        )

    base_schools = (baseline_rollup.get("schools") or {})
    curr_schools = (current_rollup.get("schools") or {})
    for school in sorted(set(base_schools) | set(curr_schools)):
        base = base_schools.get(school) or {"flagged_rows": 0, "issue_code_counts": {}}
        curr = curr_schools.get(school) or {"flagged_rows": 0, "issue_code_counts": {}}
        bf = int(base.get("flagged_rows", 0))
        cf = int(curr.get("flagged_rows", 0))
        if school not in base_schools:
            rows.append(
                _movement_row(
                    stream="metrics",
                    entity_key=school,
                    entity_label=school,
                    metric="school_presence",
                    baseline="",
                    current="present",
                    movement_type="site_appeared",
                    is_regression=False,
                )
            )
        elif school not in curr_schools:
            rows.append(
                _movement_row(
                    stream="metrics",
                    entity_key=school,
                    entity_label=school,
                    metric="school_presence",
                    baseline="present",
                    current="",
                    movement_type="site_removed",
                    is_regression=False,
                )
            )
        elif bf != cf:
            rows.append(
                _movement_row(
                    stream="metrics",
                    entity_key=school,
                    entity_label=school,
                    metric="flagged_rows",
                    baseline=bf,
                    current=cf,
                    movement_type="school_flag_count_delta",
                    is_regression=cf > bf,
                )
            )

        base_codes = base.get("issue_code_counts") or {}
        curr_codes = curr.get("issue_code_counts") or {}
        for code in sorted(set(base_codes) | set(curr_codes)):
            bv = int(base_codes.get(code, 0))
            cv = int(curr_codes.get(code, 0))
            if bv == cv:
                continue
            is_regression = cv > bv or (bv == 0 and cv > 0)
            rows.append(
                _movement_row(
                    stream="metrics",
                    entity_key=school,
                    entity_label=school,
                    metric=code,
                    baseline=bv,
                    current=cv,
                    movement_type="issue_code_delta",
                    is_regression=is_regression,
                )
            )

    base_gp = baseline_grading_period_stats or {}
    curr_gp = current_grading_period_stats or {}
    for period_label in sorted(set(base_gp) | set(curr_gp)):
        bv = int(base_gp.get(period_label, 0))
        cv = int(curr_gp.get(period_label, 0))
        if bv == cv:
            continue
        rows.append(
            _movement_row(
                stream="metrics",
                entity_key="global",
                entity_label="global",
                metric=f"grading_period_fill:{period_label}",
                baseline=bv,
                current=cv,
                movement_type="grading_period_fill_delta",
                is_regression=cv < bv,
            )
        )
    return rows


def _progress_counts(rollup: dict[str, Any]) -> dict[str, int]:
    global_counts = rollup.get("global") or {}
    return {
        status: int(global_counts.get(status, 0))
        for status in GOAL_PROGRESS_STATUS_ORDER
    }


def _progress_on_track_pct(rollup: dict[str, Any]) -> float | None:
    counts = _progress_counts(rollup)
    tracked = counts["on_track"] + counts["off_track"]
    if tracked == 0:
        return None
    return round(100 * counts["on_track"] / tracked, 1)


def _compare_goal_progress(
    base_rollup: dict[str, Any],
    curr_rollup: dict[str, Any],
) -> dict[str, Any]:
    movements: list[dict[str, Any]] = []
    base_index = base_rollup.get("progress_index") or {}
    curr_index = curr_rollup.get("progress_index") or {}

    for metric in ("on_track", "off_track"):
        bv = int((base_rollup.get("global") or {}).get(metric, 0))
        cv = int((curr_rollup.get("global") or {}).get(metric, 0))
        if bv == cv:
            continue
        is_regression = (metric == "on_track" and cv < bv) or (metric == "off_track" and cv > bv)
        movements.append(
            _movement_row(
                stream="goal_progress",
                entity_key="global",
                entity_label="global",
                metric=metric,
                baseline=bv,
                current=cv,
                movement_type=f"goal_{metric}_delta",
                is_regression=is_regression,
            )
        )

    row_level = {
        "improved": 0,
        "worsened": 0,
        "unchanged": 0,
        "appeared": 0,
        "dropped": 0,
    }
    all_keys = sorted(set(base_index) | set(curr_index))
    for key in all_keys:
        base_status = base_index.get(key)
        curr_status = curr_index.get(key)
        if base_status is None:
            row_level["appeared"] += 1
            continue
        if curr_status is None:
            row_level["dropped"] += 1
            continue
        base_rank = PROGRESS_STATUS_RANK.get(str(base_status), 0)
        curr_rank = PROGRESS_STATUS_RANK.get(str(curr_status), 0)
        if curr_rank > base_rank:
            row_level["improved"] += 1
            movements.append(
                _movement_row(
                    stream="goal_progress",
                    entity_key=key,
                    entity_label=key.split("|", 2)[1] if "|" in key else key,
                    metric="progress_status",
                    baseline=base_status,
                    current=curr_status,
                    movement_type="progress_improved",
                    is_regression=False,
                )
            )
        elif curr_rank < base_rank:
            row_level["worsened"] += 1
            movements.append(
                _movement_row(
                    stream="goal_progress",
                    entity_key=key,
                    entity_label=key.split("|", 2)[1] if "|" in key else key,
                    metric="progress_status",
                    baseline=base_status,
                    current=curr_status,
                    movement_type="progress_worsened",
                    is_regression=True,
                )
            )
        else:
            row_level["unchanged"] += 1

    regressions = [row for row in movements if row.get("is_regression")]
    school_changes: list[dict[str, Any]] = []
    base_schools = base_rollup.get("schools") or {}
    curr_schools = curr_rollup.get("schools") or {}
    for school in sorted(set(base_schools) | set(curr_schools)):
        base_pct = _school_on_track_pct(base_schools.get(school, {}))
        curr_pct = _school_on_track_pct(curr_schools.get(school, {}))
        if base_pct is None or curr_pct is None:
            continue
        school_changes.append(
            {
                "school": school,
                "baseline_on_track_pct": base_pct,
                "current_on_track_pct": curr_pct,
                "delta_pct": round(curr_pct - base_pct, 1),
            }
        )
    school_changes.sort(key=lambda row: row["delta_pct"])

    return {
        "baseline": {
            "global": _progress_counts(base_rollup),
            "eligible_rows": int(base_rollup.get("eligible_rows", 0)),
            "on_track_pct": _progress_on_track_pct(base_rollup),
        },
        "current": {
            "global": _progress_counts(curr_rollup),
            "eligible_rows": int(curr_rollup.get("eligible_rows", 0)),
            "on_track_pct": _progress_on_track_pct(curr_rollup),
        },
        "movement_count": len(movements),
        "regression_count": len(regressions),
        "row_level": row_level,
        "school_changes": school_changes,
        "movements": movements,
        "regressions": regressions,
    }


def _school_on_track_pct(school_counts: dict[str, Any]) -> float | None:
    on_track = int(school_counts.get("on_track", 0))
    off_track = int(school_counts.get("off_track", 0))
    tracked = on_track + off_track
    if tracked == 0:
        return None
    return round(100 * on_track / tracked, 1)


def _snapshot_compare_warnings(base_slot: Path, curr_slot: Path) -> list[str]:
    warnings: list[str] = []
    base_meta_path = base_slot / "snapshot.json"
    curr_meta_path = curr_slot / "snapshot.json"
    if not base_meta_path.is_file() or not curr_meta_path.is_file():
        return warnings
    base_meta = _load_json(base_meta_path)
    curr_meta = _load_json(curr_meta_path)
    base_hashes = base_meta.get("content_hashes") or {}
    curr_hashes = curr_meta.get("content_hashes") or {}
    for key in sorted(set(base_hashes) & set(curr_hashes)):
        if base_hashes[key] == curr_hashes[key]:
            warnings.append(f"identical content hash for {key}")
    base_wb = (base_meta.get("source_workbooks") or {}).get("metrics")
    curr_wb = (curr_meta.get("source_workbooks") or {}).get("metrics")
    if (
        base_wb
        and curr_wb
        and Path(base_wb).is_file()
        and Path(curr_wb).is_file()
        and _file_hash(Path(base_wb)) == _file_hash(Path(curr_wb))
    ):
        warnings.append("identical metrics workbook sha256")
    return warnings


def load_snapshot_payloads(
    slot: Path,
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, int], dict[str, Any]]:
    acc = _load_json(slot / "accreditation" / "monitoring_summary.json")
    met = _load_json(slot / "metrics" / "audit_summary.json")
    rollup = _load_json(slot / "metrics" / "school_rollup.json")
    grading_path = slot / "metrics" / "grading_period_stats.json"
    grading = _load_json(grading_path) if grading_path.is_file() else {}
    progress_path = slot / "metrics" / "progress_rollup.json"
    progress = _load_json(progress_path) if progress_path.is_file() else {}
    return acc, met, rollup, grading, progress


def _compare_snapshot_slots(
    *,
    base_slot: Path,
    curr_slot: Path,
    baseline_period: str,
    current_period: str,
    school_year: str,
    regression_threshold: float,
    compare_mode: str = "within_year",
    baseline_school_year: str | None = None,
    current_school_year: str | None = None,
) -> dict[str, Any]:
    if not base_slot.is_dir():
        raise FileNotFoundError(f"baseline snapshot not found: {base_slot}")
    if not curr_slot.is_dir():
        raise FileNotFoundError(f"current snapshot not found: {curr_slot}")

    base_acc, base_met, base_rollup, base_gp, base_progress = load_snapshot_payloads(base_slot)
    curr_acc, curr_met, curr_rollup, curr_gp, curr_progress = load_snapshot_payloads(curr_slot)
    warnings = _snapshot_compare_warnings(base_slot, curr_slot)
    for warning in warnings:
        print(f"warning: {school_year} {baseline_period} vs {current_period}: {warning}", file=sys.stderr)

    movements: list[dict[str, Any]] = []
    if not base_acc.get("metrics_only") and not curr_acc.get("metrics_only"):
        movements.extend(
            _compare_accreditation(
                base_acc,
                curr_acc,
                regression_threshold=regression_threshold,
            )
        )
    movements.extend(
        _compare_metrics(
            base_met,
            curr_met,
            base_rollup,
            curr_rollup,
            baseline_grading_period_stats=base_gp,
            current_grading_period_stats=curr_gp,
        )
    )

    goal_progress: dict[str, Any] | None = None
    if base_progress and curr_progress:
        goal_progress = _compare_goal_progress(base_progress, curr_progress)
    elif not base_progress or not curr_progress:
        missing = []
        if not base_progress:
            missing.append("baseline progress_rollup.json")
        if not curr_progress:
            missing.append("current progress_rollup.json")
        print(
            f"warning: {school_year} goal progress skipped ({', '.join(missing)} missing)",
            file=sys.stderr,
        )

    regressions = [row for row in movements if row.get("is_regression")]
    by_type: Counter[str] = Counter(row["movement_type"] for row in movements)
    by_stream: Counter[str] = Counter(row["stream"] for row in movements)

    return {
        "compare_mode": compare_mode,
        "school_year": school_year,
        "baseline_school_year": baseline_school_year or school_year,
        "current_school_year": current_school_year or school_year,
        "baseline_period": baseline_period,
        "current_period": current_period,
        "compared_at": _utc_now_iso(),
        "regression_threshold": regression_threshold,
        "movement_count": len(movements),
        "regression_count": len(regressions),
        "movements_by_type": dict(by_type),
        "movements_by_stream": dict(by_stream),
        "warnings": warnings,
        "movements": movements,
        "regressions": regressions,
        "goal_progress": goal_progress,
    }


def compare_snapshots(
    *,
    snapshots_dir: Path,
    school_year: str,
    baseline_period: str,
    current_period: str,
    regression_threshold: float = 0.01,
) -> dict[str, Any]:
    base_slot = snapshot_slot(snapshots_dir, school_year, baseline_period)
    curr_slot = snapshot_slot(snapshots_dir, school_year, current_period)
    return _compare_snapshot_slots(
        base_slot=base_slot,
        curr_slot=curr_slot,
        baseline_period=baseline_period,
        current_period=current_period,
        school_year=school_year,
        regression_threshold=regression_threshold,
        compare_mode="within_year",
    )


def compare_cross_year_snapshots(
    *,
    snapshots_dir: Path,
    baseline_school_year: str,
    baseline_period: str,
    current_school_year: str,
    current_period: str,
    regression_threshold: float = 0.01,
) -> dict[str, Any]:
    base_slot = snapshot_slot(snapshots_dir, baseline_school_year, baseline_period)
    curr_slot = snapshot_slot(snapshots_dir, current_school_year, current_period)
    label = f"{baseline_school_year}/{baseline_period} vs {current_school_year}/{current_period}"
    return _compare_snapshot_slots(
        base_slot=base_slot,
        curr_slot=curr_slot,
        baseline_period=baseline_period,
        current_period=current_period,
        school_year=label,
        regression_threshold=regression_threshold,
        compare_mode="cross_year",
        baseline_school_year=baseline_school_year,
        current_school_year=current_school_year,
    )


def _audit_config_from_snapshot_meta(meta: dict[str, Any]) -> audit.AuditConfig:
    cfg = meta.get("audit_config") or {}
    excluded = cfg.get("excluded_student_ids")
    if excluded is not None:
        excluded = set(excluded)
    goal_progress_wb = cfg.get("goal_progress_workbook")
    return audit.AuditConfig(
        sheet_name=cfg.get("sheet_name", audit.DEFAULT_SHEET),
        include_ok_rows=bool(cfg.get("include_ok_rows", False)),
        school_year=meta.get("school_year") or cfg.get("school_year"),
        goal_progress_workbook=Path(goal_progress_wb) if goal_progress_wb else None,
        excluded_student_ids=excluded,
    )


def backfill_progress_rollups(
    snapshots_dir: Path,
    *,
    force: bool = False,
) -> int:
    manifest = load_manifest(snapshots_dir)
    updated = 0
    for entry in manifest.get("snapshots", []):
        rel = str(entry.get("path", ""))
        if not rel:
            continue
        slot = snapshots_dir / rel
        progress_path = slot / "metrics" / "progress_rollup.json"
        if progress_path.is_file() and not force:
            continue
        meta_path = slot / "snapshot.json"
        if not meta_path.is_file():
            continue
        meta = _load_json(meta_path)
        metrics_wb = (meta.get("source_workbooks") or {}).get("metrics")
        if not metrics_wb or not Path(metrics_wb).is_file():
            continue
        audit_cfg = _audit_config_from_snapshot_meta(meta)
        met_results = audit.evaluate_workbook(Path(metrics_wb), audit_cfg)
        progress_payload = met_results["progress_rollup"]
        met_dir = slot / "metrics"
        met_dir.mkdir(parents=True, exist_ok=True)
        _write_json(progress_path, progress_payload)
        meta["content_hashes"] = meta.get("content_hashes") or {}
        meta["content_hashes"]["metrics/progress_rollup.json"] = _file_hash(progress_path)
        _write_json(meta_path, meta)
        updated += 1
    return updated


def run_cross_year_compares(
    school_years: list[str],
    *,
    snapshots_dir: Path,
    output_dir: Path,
    regression_threshold: float,
) -> int:
    backfilled = backfill_progress_rollups(snapshots_dir)
    if backfilled:
        print(f"backfilled progress_rollup.json for {backfilled} snapshot(s)")
    ordered = discover_eoy_snapshot_years(snapshots_dir)
    if school_years:
        allowed = set(school_years)
        ordered = [school_year for school_year in ordered if school_year in allowed]
    exit_code = 0
    for index in range(1, len(ordered)):
        prior_sy = ordered[index - 1]
        curr_sy = ordered[index]
        prior_slot = snapshot_slot(snapshots_dir, prior_sy, FALLBACK_PERIOD)
        curr_slot = snapshot_slot(snapshots_dir, curr_sy, FALLBACK_PERIOD)
        if not prior_slot.is_dir() or not curr_slot.is_dir():
            continue
        curr_period = FALLBACK_PERIOD
        pair_name = f"{prior_sy}_{FALLBACK_PERIOD}_vs_{curr_sy}_{curr_period}"
        pair_dir = output_dir / "cross_year" / pair_name
        try:
            payload = compare_cross_year_snapshots(
                snapshots_dir=snapshots_dir,
                baseline_school_year=prior_sy,
                baseline_period=FALLBACK_PERIOD,
                current_school_year=curr_sy,
                current_period=curr_period,
                regression_threshold=regression_threshold,
            )
            export_trend_results(payload, pair_dir)
        except FileNotFoundError as exc:
            print(f"error: cross-year {pair_name}: {exc}", file=sys.stderr)
            exit_code = 1
            continue
        print(f"compare cross-year: {prior_sy}/{FALLBACK_PERIOD} -> {curr_sy}/{curr_period}")
        print(f"  movements: {payload['movement_count']}")
        print(f"  output: {pair_dir}")
    from visualize import load_cross_year_summaries

    if load_cross_year_summaries(output_dir):
        report_code = render_cross_year_report(
            output_dir=output_dir,
            snapshots_dir=snapshots_dir,
        )
        if report_code != 0:
            return report_code
    return exit_code


def run_cross_year_pipeline(
    *,
    snapshots_dir: Path | None = None,
    output_dir: Path | None = None,
    inputs_dir: Path | None = None,
    qpr_dir: Path | None = None,
    goal_achievement_audit_dir: Path | None = None,
    regression_threshold: float = 0.01,
    force_fetch: bool = True,
    force: bool = False,
    include_accreditation: bool | None = None,
) -> int:
    """Pull fresh CISDM exports, refresh EOY snapshots, compare, and render HTML."""
    snapshots_dir = snapshots_dir or default_snapshots_dir()
    output_dir = output_dir or default_output_dir()
    inputs_dir = inputs_dir or default_inputs_dir()
    qpr_dir = qpr_dir or default_qpr_dir()

    school_years = discover_school_years(
        qpr_dir=qpr_dir,
        snapshots_dir=snapshots_dir,
        inputs_dir=inputs_dir,
        include_config_years=True,
    )
    if not school_years:
        print("error: no school years found for cross-year trends", file=sys.stderr)
        return 1

    exit_code = 0
    for school_year in school_years:
        print(f"=== {school_year} ===", file=sys.stderr)
        code, status = run_trend_school_year(
            school_year,
            snapshots_dir=snapshots_dir,
            output_dir=output_dir,
            inputs_dir=inputs_dir,
            qpr_dir=qpr_dir,
            force=force or force_fetch,
            force_fetch=force_fetch,
            skip_compare=True,
            skip_empty_periods=True,
            include_accreditation=include_accreditation,
        )
        print(f"status {school_year}: {status}", file=sys.stderr)
        if code != 0:
            exit_code = code

    cross_code = run_cross_year_compares(
        school_years,
        snapshots_dir=snapshots_dir,
        output_dir=output_dir,
        regression_threshold=regression_threshold,
    )
    if cross_code != 0:
        return cross_code
    if exit_code != 0:
        return exit_code

    from visualize import load_cross_year_summaries

    if load_cross_year_summaries(output_dir):
        return 0
    return render_cross_year_report(
        output_dir=output_dir,
        snapshots_dir=snapshots_dir,
        goal_achievement_audit_dir=goal_achievement_audit_dir,
    )


def render_cross_year_report(
    *,
    output_dir: Path | None = None,
    snapshots_dir: Path | None = None,
    goal_achievement_audit_dir: Path | None = None,
) -> int:
    output_dir = output_dir or default_output_dir()
    snapshots_dir = snapshots_dir or default_snapshots_dir()
    from visualize import build_cross_year_aggregates, load_cross_year_summaries, render_cross_year_html

    if not load_cross_year_summaries(output_dir):
        print(
            "error: no cross-year compare data under "
            f"{output_dir / 'cross_year'} — run cisiphyus trend --all first",
            file=sys.stderr,
        )
        return 1
    aggregates = build_cross_year_aggregates(
        output_dir,
        snapshots_dir=snapshots_dir,
        goal_achievement_audit_dir=goal_achievement_audit_dir,
    )
    report_path = output_dir / "cross_year" / "index.html"
    render_cross_year_html(aggregates, report_path)
    print(f"cross-year report: {report_path}")
    return 0


CROSS_YEAR_REPORT_ALIASES = frozenset({"cross-year", "eoy", "cross-year-report"})


def _write_movements_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=MOVEMENT_FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in MOVEMENT_FIELDS})


def export_trend_results(payload: dict[str, Any], output_dir: Path) -> dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "trend_movements": output_dir / "trend_movements.csv",
        "regression_flags": output_dir / "regression_flags.csv",
        "trend_summary_json": output_dir / "trend_summary.json",
        "trend_summary_md": output_dir / "trend_summary.md",
        "goal_progress_summary_json": output_dir / "goal_progress_summary.json",
        "goal_progress_movements": output_dir / "goal_progress_movements.csv",
    }
    _write_movements_csv(paths["trend_movements"], payload.get("movements") or [])
    _write_movements_csv(paths["regression_flags"], payload.get("regressions") or [])

    goal_progress = payload.get("goal_progress")
    if goal_progress:
        goal_summary = {k: v for k, v in goal_progress.items() if k not in {"movements", "regressions"}}
        _write_json(paths["goal_progress_summary_json"], goal_summary)
        _write_movements_csv(paths["goal_progress_movements"], goal_progress.get("movements") or [])
    else:
        paths.pop("goal_progress_summary_json")
        paths.pop("goal_progress_movements")

    summary = {
        k: v
        for k, v in payload.items()
        if k not in {"movements", "regressions", "goal_progress"}
    }
    if goal_progress:
        summary["goal_progress_summary"] = {
            k: v for k, v in goal_progress.items() if k not in {"movements", "regressions"}
        }
    _write_json(paths["trend_summary_json"], summary)

    lines = [
        "# Longitudinal Trend Summary",
        "",
        f"- **Compare mode:** {payload.get('compare_mode', 'within_year')}",
        f"- **School year:** {payload.get('school_year', '')}",
    ]
    if payload.get("compare_mode") == "cross_year":
        lines.extend(
            [
                f"- **Baseline school year:** {payload.get('baseline_school_year', '')}",
                f"- **Current school year:** {payload.get('current_school_year', '')}",
            ]
        )
    lines.extend(
        [
            f"- **Baseline:** {payload.get('baseline_period', '')}",
            f"- **Current:** {payload.get('current_period', '')}",
            f"- **Compared at:** {payload.get('compared_at', '')}",
            "",
            "## Student Metrics summary totals",
            "",
            f"- Summary totals changed: **{payload.get('movement_count', 0)}**",
            f"- Totals that went up: **{payload.get('regression_count', 0)}**",
            "",
        ]
    )
    if goal_progress:
        current = goal_progress.get("current") or {}
        baseline = goal_progress.get("baseline") or {}
        row_level = goal_progress.get("row_level") or {}
        lines.extend(
            [
                "## Student goal progress",
                "",
                f"- Tracked goals (baseline EOY): **{baseline.get('eligible_rows', 0)}**",
                f"- Tracked goals (current EOY): **{current.get('eligible_rows', 0)}**",
                f"- On track (baseline EOY): **{baseline.get('on_track_pct', 'n/a')}%**",
                f"- On track (current EOY): **{current.get('on_track_pct', 'n/a')}%**",
                f"- Goal progress movements: **{goal_progress.get('movement_count', 0)}**",
                f"- Goals improved: **{row_level.get('improved', 0)}**",
                f"- Goals fell behind: **{row_level.get('worsened', 0)}**",
                "",
            ]
        )
    warnings = payload.get("warnings") or []
    if warnings:
        lines.append("## Warnings")
        lines.append("")
        for warning in warnings:
            lines.append(f"- {warning}")
        lines.append("")
    by_type = payload.get("movements_by_type") or {}
    if by_type:
        lines.append("## Movements by type")
        lines.append("")
        for key in sorted(by_type):
            lines.append(f"- **{key}**: {by_type[key]}")
        lines.append("")
    paths["trend_summary_md"].write_text("\n".join(lines), encoding="utf-8")
    return {key: str(path) for key, path in paths.items()}


def resolve_compare_periods(
    manifest: dict[str, Any],
    *,
    school_year: str,
) -> tuple[str, str]:
    entries = [
        e
        for e in manifest.get("snapshots", [])
        if e.get("school_year") == school_year
    ]
    if len(entries) < 2:
        raise ValueError(f"need at least two snapshots for {school_year!r}")
    ordered = sorted(entries, key=lambda e: period_sort_key(str(e.get("period", ""))))
    return ordered[-2]["period"], ordered[-1]["period"]


def _cmd_trend(args: argparse.Namespace) -> int:
    snapshots_dir = default_snapshots_dir()
    output_dir = default_output_dir()
    inputs_dir = default_inputs_dir()
    qpr_dir = default_qpr_dir()

    school_years = list(args.school_years)
    if args.all:
        school_years.extend(
            discover_school_years(
                qpr_dir=qpr_dir,
                snapshots_dir=snapshots_dir,
                inputs_dir=inputs_dir,
                include_config_years=True,
            )
        )
    if not school_years:
        fallback = default_school_year()
        if fallback:
            school_years = [fallback]
    # preserve order, drop duplicates
    seen: set[str] = set()
    unique_years: list[str] = []
    for year in school_years:
        if year not in seen:
            seen.add(year)
            unique_years.append(year)
    if not unique_years:
        print(
            "error: provide school year(s), e.g. cisiphyus trend SY25-26, or use --all",
            file=sys.stderr,
        )
        return 1

    exit_code = 0
    include_accreditation = _accreditation_cli_flag(args)
    for school_year in unique_years:
        print(f"=== {school_year} ===")
        code, status = run_trend_school_year(
            school_year,
            snapshots_dir=snapshots_dir,
            output_dir=output_dir,
            inputs_dir=inputs_dir,
            qpr_dir=qpr_dir,
            force=args.force,
            force_fetch=args.force_fetch,
            skip_compare=args.skip_compare,
            regression_threshold=args.regression_threshold,
            skip_empty_periods=args.all,
            include_accreditation=include_accreditation,
        )
        print(f"status {school_year}: {status}", file=sys.stderr)
        if code != 0:
            exit_code = code

    if args.all and not args.skip_compare:
        cross_code = run_cross_year_compares(
            unique_years,
            snapshots_dir=snapshots_dir,
            output_dir=output_dir,
            regression_threshold=args.regression_threshold,
        )
        if cross_code != 0:
            exit_code = cross_code
    return exit_code


def _build_trend_parser(*, prog: str | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Capture snapshots for all available reporting periods in a school year "
            "and compare consecutive quarters."
        ),
        prog=prog,
    )
    parser.add_argument(
        "school_years",
        nargs="*",
        help="School year label(s), e.g. SY25-26 SY24-25",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Process every school year from school_year_programs, QPR output, and snapshots.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Replace existing snapshot slots.",
    )
    parser.add_argument(
        "--force-fetch",
        action="store_true",
        help="Always re-pull CISDM exports (default: pull when no cached copy for the period).",
    )
    parser.add_argument(
        "--skip-compare",
        action="store_true",
        help="Capture snapshots only; skip consecutive period comparisons.",
    )
    parser.add_argument(
        "--regression-threshold",
        type=float,
        default=0.01,
        help="Minimum pct drop to flag accreditation metric regression (default 0.01).",
    )
    parser.add_argument(
        "--accreditation",
        action="store_true",
        help="Pull and snapshot accreditation (default: only when ACCREDITATION_FETCH_FROM_CISDM=1).",
    )
    return parser


def _accreditation_cli_flag(args: argparse.Namespace) -> bool | None:
    return True if args.accreditation else None


def _build_cross_year_parser(*, prog: str | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Refresh EOY snapshots from CISDM, run cross-year compares, and write the HTML report."
        ),
        prog=prog,
    )
    parser.add_argument(
        "--no-fetch",
        action="store_true",
        help="Skip CISDM pull and re-compare; only regenerate HTML from existing artifacts.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-capture snapshots even when period directories already exist.",
    )
    parser.add_argument(
        "--regression-threshold",
        type=float,
        default=0.01,
        help="Minimum pct drop to flag accreditation metric regression (default 0.01).",
    )
    parser.add_argument(
        "--snapshots-dir",
        type=Path,
        default=None,
        help="Snapshot root (default: TREND_SNAPSHOTS_DIR or artifacts/snapshots/).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Trend output root (default: TREND_OUTPUT_DIR or artifacts/trends/).",
    )
    parser.add_argument(
        "--goal-achievement-audit-dir",
        type=Path,
        default=None,
        help="Goal Achievement audit CSV directory (default: GOAL_ACHIEVEMENT_AUDIT_DIR or evaluation path).",
    )
    parser.add_argument(
        "--accreditation",
        action="store_true",
        help="Pull and snapshot accreditation (default: only when ACCREDITATION_FETCH_FROM_CISDM=1).",
    )
    return parser


def main_cross_year(argv: list[str] | None = None) -> int:
    args = _build_cross_year_parser(prog="cisiphyus trend cross-year").parse_args(argv)
    if args.no_fetch:
        return render_cross_year_report(
            output_dir=args.output_dir,
            snapshots_dir=args.snapshots_dir,
            goal_achievement_audit_dir=args.goal_achievement_audit_dir,
        )
    return run_cross_year_pipeline(
        output_dir=args.output_dir,
        snapshots_dir=args.snapshots_dir,
        inputs_dir=None,
        qpr_dir=None,
        goal_achievement_audit_dir=args.goal_achievement_audit_dir,
        regression_threshold=args.regression_threshold,
        force_fetch=True,
        force=args.force,
        include_accreditation=_accreditation_cli_flag(args),
    )


def main_trend(argv: list[str] | None = None) -> int:
    args = _build_trend_parser(prog="cisiphyus trend").parse_args(argv)
    return _cmd_trend(args)


if __name__ == "__main__":
    raise SystemExit(main_trend(sys.argv[1:]))
