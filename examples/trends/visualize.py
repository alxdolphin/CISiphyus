"""Cross-year trend visualization: aggregate artifacts and render HTML report."""

from __future__ import annotations

import csv
import json
import os
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Literal

GLOBAL_METRICS = (
    "baseline_without_target",
    "both_baseline_and_target_blank",
    "baseline_target_direction_mismatch",
)

MOVEMENT_TYPE_ORDER = (
    "issue_code_delta",
    "school_flag_count_delta",
    "global_issue_delta",
    "grading_period_fill_delta",
    "site_appeared",
    "site_removed",
)

MOVEMENT_TYPE_LABELS: dict[str, str] = {
    "issue_code_delta": "Flag type count (by school)",
    "school_flag_count_delta": "Flagged rows at a school",
    "global_issue_delta": "Flag type count (network-wide)",
    "grading_period_fill_delta": "Grading period entry counts",
    "site_appeared": "School added to export",
    "site_removed": "School removed from export",
}

GLOBAL_METRIC_LABELS: dict[str, str] = {
    "baseline_without_target": "Baseline without target",
    "both_baseline_and_target_blank": "Baseline and target blank",
    "baseline_target_direction_mismatch": "Direction mismatches",
}

EOY_PERIOD = "EOY"

GOAL_ACHIEVEMENT_EXCEPTION_TYPES = (
    "goal_achievement_mismatch",
    "needs_manual_review",
)

GOAL_ACHIEVEMENT_EXCEPTION_LABELS: dict[str, str] = {
    "goal_achievement_mismatch": "Goal achievement mismatch",
    "needs_manual_review": "Needs manual review",
}

GAR_RECORDED_OUTCOME_LABELS: dict[str, str] = {
    "Goal Met": "Goal Met",
    "Goal Not Met, With Progress": "With progress",
    "Goal Not Met, No Progress": "No progress",
}
GAR_GOAL_NOT_MET_LABELS = tuple(GAR_RECORDED_OUTCOME_LABELS)[1:]


def gar_goal_not_met_pct(outcomes: dict[str, Any] | None) -> float | None:
    if not outcomes:
        return None
    total = sum(
        int(outcomes.get(label, 0)) for label in GAR_RECORDED_OUTCOME_LABELS
    )
    if not total:
        return None
    not_met = sum(int(outcomes.get(label, 0)) for label in GAR_GOAL_NOT_MET_LABELS)
    return round(100 * not_met / total, 1)

# ponytail: fixed set — add colors here if stacked palette grows
_STACKED_DARK_FILLS = frozenset({"#4c8bf5", "#c678dd", "#56b6c2", "#be5046"})
_SEGMENT_LABEL_MIN_HEIGHT_PX = 14


def _school_year_sort_key(label: str) -> tuple[int, int]:
    stripped = label.strip().upper()
    if not stripped.startswith("SY") or "-" not in stripped:
        return (9999, 9999)
    start_text, end_text = stripped[2:].split("-", 1)
    try:
        return (int(start_text), int(end_text))
    except ValueError:
        return (9999, 9999)


def _transition_pair_label(row: dict[str, Any]) -> str:
    baseline = str(row.get("baseline_school_year", ""))
    current = str(row.get("current_school_year", ""))
    if baseline and current:
        return f"{baseline}→{current}"
    return str(row.get("label", current))


def load_cross_year_summaries(output_dir: Path) -> list[dict[str, Any]]:
    cross_year_dir = output_dir / "cross_year"
    if not cross_year_dir.is_dir():
        return []
    summaries: list[dict[str, Any]] = []
    for pair_dir in cross_year_dir.iterdir():
        if not pair_dir.is_dir():
            continue
        summary_path = pair_dir / "trend_summary.json"
        if not summary_path.is_file():
            continue
        payload = json.loads(summary_path.read_text(encoding="utf-8"))
        payload["_pair_dir"] = str(pair_dir)
        payload["_pair_name"] = pair_dir.name
        summaries.append(payload)
    summaries.sort(key=lambda row: _school_year_sort_key(str(row.get("current_school_year", ""))))
    return summaries


def aggregate_global_metrics(pair_dir: Path) -> dict[str, int | None]:
    movements_path = pair_dir / "trend_movements.csv"
    if not movements_path.is_file():
        return {metric: None for metric in GLOBAL_METRICS}
    values: dict[str, int | None] = {metric: None for metric in GLOBAL_METRICS}
    with movements_path.open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            if row.get("entity_key") != "global":
                continue
            metric = row.get("metric", "")
            if metric not in values:
                continue
            try:
                values[metric] = int(row.get("current", ""))
            except ValueError:
                values[metric] = None
    return values


def aggregate_school_regressions(pair_dir: Path, *, top_n: int = 10) -> list[dict[str, Any]]:
    regressions_path = pair_dir / "regression_flags.csv"
    if not regressions_path.is_file():
        return []
    counts: Counter[str] = Counter()
    metric_counts: dict[str, Counter[str]] = defaultdict(Counter)
    with regressions_path.open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            school = row.get("entity_label", "")
            if not school or row.get("entity_key") == "global":
                continue
            counts[school] += 1
            metric_counts[school][row.get("metric", "")] += 1
    rows: list[dict[str, Any]] = []
    for school, count in counts.most_common(top_n):
        top_metric = ""
        if metric_counts[school]:
            top_metric = metric_counts[school].most_common(1)[0][0]
        rows.append({"school": school, "regressions": count, "top_metric": top_metric})
    return rows


def load_goal_progress_summary(pair_dir: Path) -> dict[str, Any] | None:
    summary_path = pair_dir / "goal_progress_summary.json"
    if not summary_path.is_file():
        return None
    return json.loads(summary_path.read_text(encoding="utf-8"))


def aggregate_school_goal_changes(pair_dir: Path, *, top_n: int = 10) -> list[dict[str, Any]]:
    summary = load_goal_progress_summary(pair_dir)
    if not summary:
        return []
    changes = summary.get("school_changes") or []
    rows: list[dict[str, Any]] = []
    for item in sorted(changes, key=lambda row: row.get("delta_pct", 0))[:top_n]:
        rows.append(
            {
                "school": item.get("school", ""),
                "delta_pct": item.get("delta_pct", 0),
                "baseline_on_track_pct": item.get("baseline_on_track_pct"),
                "current_on_track_pct": item.get("current_on_track_pct"),
            }
        )
    return rows


def discover_eoy_snapshot_years(snapshots_dir: Path) -> list[str]:
    if not snapshots_dir.is_dir():
        return []
    years: list[str] = []
    for entry in snapshots_dir.iterdir():
        if not entry.is_dir():
            continue
        if (entry / EOY_PERIOD).is_dir():
            years.append(entry.name)
    return sorted(years, key=_school_year_sort_key)


def load_eoy_global_issue_counts(snapshots_dir: Path, school_year: str) -> dict[str, int | None]:
    path = snapshots_dir / school_year / EOY_PERIOD / "metrics" / "audit_summary.json"
    if not path.is_file():
        return {metric: None for metric in GLOBAL_METRICS}
    payload = json.loads(path.read_text(encoding="utf-8"))
    counts = payload.get("issue_code_counts") or {}
    values: dict[str, int | None] = {}
    for metric in GLOBAL_METRICS:
        raw = counts.get(metric)
        values[metric] = int(raw) if raw is not None else 0
    return values


def _cisiphyus_root() -> Path:
    return Path(os.environ.get("CISIPHYUS_ROOT", str(Path(__file__).resolve().parents[2])))


def load_metrics_audit_payload(
    snapshots_dir: Path | None,
    school_year: str,
) -> dict[str, Any] | None:
    candidates: list[Path] = []
    if snapshots_dir is not None:
        candidates.append(
            snapshots_dir / school_year / EOY_PERIOD / "metrics" / "audit_summary.json"
        )
    candidates.append(
        _cisiphyus_root() / "artifacts" / "audit" / school_year / "audit_summary.json"
    )
    for path in candidates:
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    return None


def _year_metrics_from_audit_payload(payload: dict[str, Any]) -> dict[str, Any]:
    import sys

    audit_dir = Path(__file__).resolve().parents[1] / "audit"
    if str(audit_dir) not in sys.path:
        sys.path.insert(0, str(audit_dir))
    import audit  # noqa: WPS433

    results = {
        "summary": payload.get("summary") or {},
        "baseline_target_distribution": payload.get("baseline_target_distribution") or {},
        "issue_code_counts": payload.get("issue_code_counts") or {},
        "accepted_exception_counts": payload.get("accepted_exception_counts") or {},
        "detail_rows": [{}] * int(payload.get("detail_row_count") or 0),
    }
    return audit.year_audit_metrics(results)


def load_goal_metric_counts_by_year(
    snapshots_dir: Path | None,
    school_years: list[str],
) -> dict[str, dict[str, int]]:
    by_year: dict[str, dict[str, int]] = {}
    for school_year in school_years:
        payload = load_metrics_audit_payload(snapshots_dir, school_year)
        if payload is None:
            continue
        metrics = _year_metrics_from_audit_payload(payload)
        total = int(metrics.get("rows_evaluated", 0) or 0)
        non_goals = int(metrics.get("accepted_supplemental", 0) or 0)
        goals_with_baseline_target = int(
            metrics.get("baseline_and_target_present", 0) or 0
        )
        goal_metrics_total = total - non_goals
        if goal_metrics_total > 0:
            by_year[school_year] = {
                "goals_with_baseline_target": goals_with_baseline_target,
                "goal_metrics_total": goal_metrics_total,
            }
    return by_year


METRICS_AUDIT_PRIMARY_ROWS: tuple[tuple[str, str, bool], ...] = (
    ("Student Metrics", "rows_evaluated", False),
    ("Baseline + target present", "baseline_and_target_present", True),
    ("Non-Compliant (Baseline without target)", "baseline_without_target_excl_supplemental", True),
    ("Flagged rows", "rows_flagged", False),
)

METRICS_AUDIT_DETAIL_ROWS: tuple[tuple[str, str, bool], ...] = (
    ("Excluded Exceptions (Non-Goal Metrics)", "accepted_supplemental", False),
    ("Direction Mismatches", "direction_mismatches", False),
    ("Scale Mismatches", "scale_mismatches", False),
)


def _metrics_audit_row_values(
    row_defs: tuple[tuple[str, str, bool], ...],
    *,
    by_year: dict[str, dict[str, Any]],
    present_years: list[str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for label, metric_key, show_pct in row_defs:
        values: dict[str, str] = {}
        for school_year in present_years:
            metrics = by_year[school_year]
            raw = int(metrics.get(metric_key, 0) or 0)
            rows_evaluated = int(metrics.get("rows_evaluated", 0) or 0)
            if show_pct and rows_evaluated:
                values[school_year] = f"{raw:,} ({round(100 * raw / rows_evaluated)}%)"
            else:
                values[school_year] = f"{raw:,}"
        rows.append({"label": label, "values": values})
    return rows


def build_metrics_audit_cross_year_table(
    snapshots_dir: Path | None,
    school_years: list[str],
) -> dict[str, Any]:
    by_year: dict[str, dict[str, Any]] = {}
    for school_year in school_years:
        payload = load_metrics_audit_payload(snapshots_dir, school_year)
        if payload is None:
            continue
        by_year[school_year] = _year_metrics_from_audit_payload(payload)
    present_years = [year for year in school_years if year in by_year]
    primary_rows = _metrics_audit_row_values(
        METRICS_AUDIT_PRIMARY_ROWS,
        by_year=by_year,
        present_years=present_years,
    )
    detail_rows = _metrics_audit_row_values(
        METRICS_AUDIT_DETAIL_ROWS,
        by_year=by_year,
        present_years=present_years,
    )
    return {
        "school_years": present_years,
        "rows": primary_rows + detail_rows,
        "primary_rows": primary_rows,
        "detail_rows": detail_rows,
    }


def _render_metrics_audit_body_rows(
    rows: list[dict[str, Any]],
    school_years: list[str],
    *,
    latest_year: str,
    label_class: str = "",
) -> str:
    body_rows: list[str] = []
    label_attr = f' class="{label_class}"' if label_class else ""
    for row in rows:
        cells = []
        for school_year in school_years:
            value = (row.get("values") or {}).get(school_year, "n/a")
            cell_class = ' class="latest-col"' if school_year == latest_year else ""
            cells.append(f"<td{cell_class}>{value}</td>")
        body_rows.append(f"<tr><td{label_attr}>{row['label']}</td>{''.join(cells)}</tr>")
    return "".join(body_rows)


def render_metrics_audit_cross_year_table(table: dict[str, Any]) -> str:
    school_years = table.get("school_years") or []
    primary_rows = table.get("primary_rows") or table.get("rows") or []
    detail_rows = table.get("detail_rows") or []
    if not school_years or not primary_rows:
        return ""
    latest_year = school_years[-1]
    header_cells = "".join(
        f'<th class="latest-col">{year}</th>' if year == latest_year else f"<th>{year}</th>"
        for year in school_years
    )
    primary_body = _render_metrics_audit_body_rows(
        primary_rows, school_years, latest_year=latest_year
    )
    detail_body = ""
    if detail_rows:
        detail_body = _render_metrics_audit_body_rows(
            detail_rows,
            school_years,
            latest_year=latest_year,
            label_class="metrics-audit-detail-label",
        )
    colgroup = (
        f'<colgroup><col class="metrics-audit-label-col">'
        f'<col span="{len(school_years)}"></colgroup>'
    )
    return f"""
  <table class="metrics-audit-table">
    {colgroup}
    <thead><tr><th>Metric</th>{header_cells}</tr></thead>
    <tbody>{primary_body}{detail_body}</tbody>
  </table>
"""


def resolve_goal_achievement_audit_dir(explicit: Path | None = None) -> Path | None:
    if explicit is not None and explicit.is_dir():
        return explicit
    env = os.environ.get("GOAL_ACHIEVEMENT_AUDIT_DIR", "").strip()
    if env:
        path = Path(env)
        return path if path.is_dir() else None
    cisiphyus_root = Path(
        os.environ.get("CISIPHYUS_ROOT", str(Path(__file__).resolve().parents[2]))
    )
    generated = cisiphyus_root / "artifacts" / "goal_achievement_audit"
    if generated.is_dir():
        return generated
    candidate = (
        cisiphyus_root.parent.parent
        / "evaluation"
        / "analysis"
        / "Audits"
        / "Goal Achievement"
    )
    return candidate if candidate.is_dir() else generated


def _cisiphyus_goal_achievement_audit_dir() -> Path:
    cisiphyus_root = Path(
        os.environ.get("CISIPHYUS_ROOT", str(Path(__file__).resolve().parents[2]))
    )
    return cisiphyus_root / "artifacts" / "goal_achievement_audit"


def _load_gar_records_for_school_year(school_year: str) -> list[dict[str, Any]] | None:
    import sys

    audit_dir = Path(__file__).resolve().parents[1] / "audit"
    goal_achievement_dir = Path(__file__).resolve().parents[1] / "goal_achievement"
    for path in (audit_dir, goal_achievement_dir):
        if str(path) not in sys.path:
            sys.path.insert(0, str(path))

    import audit  # noqa: WPS433

    workbook = audit.resolve_goal_progress_workbook(school_year)
    if workbook is None:
        return None

    records = audit.load_goal_progress_gar_records(workbook)
    if not records:
        return None

    excluded = audit.resolve_ewgspe_only_student_ids(
        school_year,
        goal_progress_workbook=workbook,
    )
    if excluded:
        records = [
            record
            for record in records
            if (audit.clean_value(str(record.get("Student ID") or "")) or "") not in excluded
        ]
    return records


def materialize_goal_achievement_audit_csv(school_year: str) -> Path | None:
    output_dir = _cisiphyus_goal_achievement_audit_dir()
    output_path = output_dir / f"{school_year}_GoalAchievement_AUDIT.csv"
    rollup_path = output_dir / f"{school_year}_GoalAchievement_ROLLUP.json"
    if output_path.is_file() and rollup_path.is_file():
        return output_path

    records = _load_gar_records_for_school_year(school_year)
    if not records:
        return output_path if output_path.is_file() else None

    import sys

    audit_dir = Path(__file__).resolve().parents[1] / "audit"
    if str(audit_dir) not in sys.path:
        sys.path.insert(0, str(audit_dir))
    import audit  # noqa: WPS433

    context: dict[str, str] = {}
    outcome_rollup = audit.gar_outcome_rollup(records, context)
    output_dir.mkdir(parents=True, exist_ok=True)
    audit.gar_write_outcome_rollup_json(outcome_rollup, rollup_path)
    if not output_path.is_file():
        exception_rows = audit.gar_build_exception_rows(records, context)
        audit.gar_write_csv_report(exception_rows, output_path)
    return output_path


def resolve_goal_achievement_audit_csv(
    audit_dir: Path | None,
    school_year: str,
) -> Path | None:
    candidates: list[Path] = []
    if audit_dir is not None:
        candidates.append(audit_dir / f"{school_year}_GoalAchievement_AUDIT.csv")
    generated_dir = _cisiphyus_goal_achievement_audit_dir()
    if audit_dir is None or audit_dir.resolve() != generated_dir.resolve():
        candidates.append(generated_dir / f"{school_year}_GoalAchievement_AUDIT.csv")
    for path in candidates:
        if path.is_file():
            return path
    return materialize_goal_achievement_audit_csv(school_year)


def resolve_goal_achievement_rollup_json(
    audit_dir: Path | None,
    school_year: str,
) -> Path | None:
    rollup_name = f"{school_year}_GoalAchievement_ROLLUP.json"
    candidates: list[Path] = []
    if audit_dir is not None:
        candidates.append(audit_dir / rollup_name)
    generated_dir = _cisiphyus_goal_achievement_audit_dir()
    candidates.append(generated_dir / rollup_name)
    for path in candidates:
        if path.is_file():
            return path
    materialize_goal_achievement_audit_csv(school_year)
    for path in candidates:
        if path.is_file():
            return path
    return None


def load_goal_achievement_rollup(
    audit_dir: Path | None,
    school_year: str,
) -> dict[str, Any] | None:
    path = resolve_goal_achievement_rollup_json(audit_dir, school_year)
    if path is None:
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def aggregate_school_goal_achievement_mismatches(
    rollup: dict[str, Any] | None,
    *,
    top_n: int = 10,
) -> list[dict[str, Any]]:
    if not rollup:
        return []
    rows = [
        {
            "school": school,
            "mismatched": int((counts or {}).get("mismatched", 0)),
            "mismatch_rate_pct": (counts or {}).get("mismatch_rate_pct"),
            "recorded_goal_met_pct": (counts or {}).get("recorded_goal_met_pct"),
            "expected_goal_met_pct": (counts or {}).get("expected_goal_met_pct"),
        }
        for school, counts in (rollup.get("schools") or {}).items()
        if int((counts or {}).get("mismatched", 0)) > 0
    ]
    rows.sort(key=lambda row: row["mismatched"], reverse=True)
    return rows[:top_n]


def load_goal_achievement_year(
    audit_dir: Path | None,
    school_year: str,
) -> dict[str, Any] | None:
    path = resolve_goal_achievement_audit_csv(audit_dir, school_year)
    if path is None:
        return None
    with path.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
    by_type = Counter(str(row.get("exception_type", "")) for row in rows)
    school_counts: Counter[str] = Counter()
    school_metrics: dict[str, Counter[str]] = defaultdict(Counter)
    for row in rows:
        school = str(row.get("home_school", ""))
        if not school:
            continue
        school_counts[school] += 1
        metric = str(row.get("metric", ""))
        if metric:
            school_metrics[school][metric] += 1
    return {
        "total": len(rows),
        "by_type": {key: int(value) for key, value in by_type.items()},
        "school_counts": dict(school_counts),
        "school_top_metric": {
            school: metrics.most_common(1)[0][0]
            for school, metrics in school_metrics.items()
            if metrics
        },
    }


def aggregate_school_goal_achievement_exceptions(
    audit_dir: Path,
    school_year: str,
    *,
    top_n: int = 10,
) -> list[dict[str, Any]]:
    payload = load_goal_achievement_year(audit_dir, school_year)
    if not payload:
        return []
    rows = [
        {
            "school": school,
            "exceptions": count,
            "top_metric": payload.get("school_top_metric", {}).get(school, ""),
        }
        for school, count in payload.get("school_counts", {}).items()
    ]
    rows.sort(key=lambda row: row["exceptions"], reverse=True)
    return rows[:top_n]


def load_eoy_progress_rollup(snapshots_dir: Path, school_year: str) -> dict[str, Any] | None:
    path = snapshots_dir / school_year / EOY_PERIOD / "metrics" / "progress_rollup.json"
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def build_cross_year_aggregates(
    output_dir: Path,
    *,
    snapshots_dir: Path | None = None,
    goal_achievement_audit_dir: Path | None = None,
) -> dict[str, Any]:
    summaries = load_cross_year_summaries(output_dir)
    transitions: list[dict[str, Any]] = []
    transition_categories: list[str] = []

    for summary in summaries:
        current_sy = str(summary.get("current_school_year", ""))
        transition_categories.append(current_sy)
        by_type = summary.get("movements_by_type") or {}
        transitions.append(
            {
                "label": current_sy,
                "pair_name": summary.get("_pair_name", ""),
                "baseline_school_year": summary.get("baseline_school_year", ""),
                "current_school_year": current_sy,
                "baseline_period": summary.get("baseline_period", ""),
                "current_period": summary.get("current_period", ""),
                "movement_count": summary.get("movement_count", 0),
                "regression_count": summary.get("regression_count", 0),
                "movements_by_type": {
                    movement_type: int(by_type.get(movement_type, 0))
                    for movement_type in MOVEMENT_TYPE_ORDER
                },
            }
        )

    latest = transitions[-1] if transitions else None
    latest_pair_dir = Path(summaries[-1]["_pair_dir"]) if summaries else None
    school_leaderboard: list[dict[str, Any]] = []

    transition_years: set[str] = set()
    for summary in summaries:
        baseline_sy = summary.get("baseline_school_year")
        current_sy = summary.get("current_school_year")
        if baseline_sy:
            transition_years.add(str(baseline_sy))
        if current_sy:
            transition_years.add(str(current_sy))

    eoy_categories: list[str] = []
    global_series: dict[str, list[int | None]] = {metric: [] for metric in GLOBAL_METRICS}
    exception_count_series: list[int | None] = []
    rows_evaluated_series: list[int | None] = []
    schools_with_exceptions_series: list[int | None] = []
    issue_code_series: dict[str, list[int]] = {}
    goal_achievement_audit_dir = resolve_goal_achievement_audit_dir(
        goal_achievement_audit_dir
    )
    goal_on_track_series: list[int | None] = []
    goal_off_track_series: list[int | None] = []
    goal_on_track_pct_series: list[float | None] = []
    gar_recorded_goal_met_pct_series: list[float | None] = []
    gar_expected_goal_met_pct_series: list[float | None] = []
    gar_recorded_goal_not_met_pct_series: list[float | None] = []
    gar_expected_goal_not_met_pct_series: list[float | None] = []
    gar_mismatch_rate_series: list[float | None] = []
    gar_recorded_composition_series: dict[str, list[int]] = {}
    gar_school_mismatch_leaderboard: list[dict[str, Any]] = []
    goal_progress_transitions: list[dict[str, Any]] = []
    missing_goal_achievement_years: list[str] = []

    if snapshots_dir is not None:
        eoy_categories = discover_eoy_snapshot_years(snapshots_dir)
        if transition_years:
            eoy_categories = [year for year in eoy_categories if year in transition_years]
        all_exception_types: set[str] = set()
        year_exception_types: list[dict[str, int]] = []
        for school_year in eoy_categories:
            gar = load_goal_achievement_year(goal_achievement_audit_dir, school_year)
            gar_rollup = load_goal_achievement_rollup(goal_achievement_audit_dir, school_year)
            if gar:
                exception_count_series.append(int(gar["total"]))
                schools_with_exceptions_series.append(len(gar.get("school_counts", {})))
                type_counts = {
                    str(key): int(value) for key, value in gar["by_type"].items()
                }
                year_exception_types.append(type_counts)
                all_exception_types.update(type_counts)
            else:
                exception_count_series.append(None)
                schools_with_exceptions_series.append(None)
                year_exception_types.append({})
                if resolve_goal_achievement_audit_csv(goal_achievement_audit_dir, school_year) is None:
                    missing_goal_achievement_years.append(school_year)
            if gar_rollup:
                global_rollup = gar_rollup.get("global") or {}
                gar_recorded_goal_met_pct_series.append(
                    global_rollup.get("recorded_goal_met_pct")
                )
                gar_expected_goal_met_pct_series.append(
                    global_rollup.get("expected_goal_met_pct")
                )
                gar_recorded_goal_not_met_pct_series.append(
                    global_rollup.get("recorded_goal_not_met_pct")
                    or gar_goal_not_met_pct(global_rollup.get("recorded"))
                )
                gar_expected_goal_not_met_pct_series.append(
                    global_rollup.get("expected_goal_not_met_pct")
                    or gar_goal_not_met_pct(global_rollup.get("expected"))
                )
                gar_mismatch_rate_series.append(global_rollup.get("mismatch_rate_pct"))
                for outcome in GAR_RECORDED_OUTCOME_LABELS:
                    recorded = (global_rollup.get("recorded") or {}).get(outcome, 0)
                    gar_recorded_composition_series.setdefault(outcome, []).append(
                        int(recorded)
                    )
            else:
                gar_recorded_goal_met_pct_series.append(None)
                gar_expected_goal_met_pct_series.append(None)
                gar_recorded_goal_not_met_pct_series.append(None)
                gar_expected_goal_not_met_pct_series.append(None)
                gar_mismatch_rate_series.append(None)
                for outcome in GAR_RECORDED_OUTCOME_LABELS:
                    gar_recorded_composition_series.setdefault(outcome, []).append(0)
            issue_counts = load_eoy_global_issue_counts(snapshots_dir, school_year)
            for metric in GLOBAL_METRICS:
                global_series[metric].append(issue_counts.get(metric))
            rollup = load_eoy_progress_rollup(snapshots_dir, school_year)
            rows_evaluated_series.append(
                int(rollup.get("eligible_rows", 0)) if rollup else None
            )
            global_counts = (rollup or {}).get("global") or {}
            goal_progress_transitions.append(
                {
                    "label": school_year,
                    "on_track": int(global_counts.get("on_track", 0)),
                    "off_track": int(global_counts.get("off_track", 0)),
                    "on_track_pct": _on_track_pct_from_rollup(rollup),
                    "eligible_rows": int((rollup or {}).get("eligible_rows", 0)),
                }
            )
            goal_on_track_series.append(
                int(global_counts.get("on_track", 0)) if rollup else None
            )
            goal_off_track_series.append(
                int(global_counts.get("off_track", 0)) if rollup else None
            )
            goal_on_track_pct_series.append(_on_track_pct_from_rollup(rollup))
        exception_type_order = [
            key
            for key in GOAL_ACHIEVEMENT_EXCEPTION_TYPES
            if key in all_exception_types
        ]
        exception_type_order.extend(
            sorted(all_exception_types - set(exception_type_order))
        )
        for exception_type in exception_type_order:
            issue_code_series[exception_type] = [
                year_counts.get(exception_type, 0)
                for year_counts in year_exception_types
            ]
        if eoy_categories and goal_achievement_audit_dir:
            school_leaderboard = aggregate_school_goal_achievement_exceptions(
                goal_achievement_audit_dir,
                eoy_categories[-1],
            )
            gar_latest_rollup = load_goal_achievement_rollup(
                goal_achievement_audit_dir,
                eoy_categories[-1],
            )
            gar_school_mismatch_leaderboard = aggregate_school_goal_achievement_mismatches(
                gar_latest_rollup,
            )
    else:
        eoy_categories = list(transition_categories)
        for summary in summaries:
            pair_dir = Path(summary["_pair_dir"])
            global_values = aggregate_global_metrics(pair_dir)
            for metric in GLOBAL_METRICS:
                global_series[metric].append(global_values.get(metric))
            goal_summary = load_goal_progress_summary(pair_dir)
            current = (goal_summary or {}).get("current") or {}
            global_counts = current.get("global") or {}
            goal_progress_transitions.append(
                {
                    "label": str(summary.get("current_school_year", "")),
                    "on_track": int(global_counts.get("on_track", 0)),
                    "off_track": int(global_counts.get("off_track", 0)),
                    "on_track_pct": current.get("on_track_pct"),
                    "eligible_rows": int(current.get("eligible_rows", 0)),
                    "row_level": (goal_summary or {}).get("row_level") or {},
                }
            )
            goal_on_track_series.append(
                int(global_counts.get("on_track", 0)) if goal_summary else None
            )
            goal_off_track_series.append(
                int(global_counts.get("off_track", 0)) if goal_summary else None
            )
            goal_on_track_pct_series.append(current.get("on_track_pct") if goal_summary else None)

    if not school_leaderboard and latest_pair_dir:
        school_leaderboard = aggregate_school_regressions(latest_pair_dir)

    latest_goal_summary = (
        load_goal_progress_summary(latest_pair_dir) if latest_pair_dir else None
    )
    goal_school_leaderboard = (
        aggregate_school_goal_changes(latest_pair_dir) if latest_pair_dir else []
    )

    latest_headline: dict[str, Any] | None = None
    if exception_count_series and eoy_categories:
        latest_exceptions = next(
            (value for value in reversed(exception_count_series) if value is not None),
            None,
        )
        prior_exceptions = next(
            (
                value
                for value in reversed(exception_count_series[:-1])
                if value is not None
            ),
            None,
        )
        exception_delta = None
        if latest_exceptions is not None and prior_exceptions is not None:
            exception_delta = latest_exceptions - prior_exceptions
        schools_with_exceptions = 0
        if goal_achievement_audit_dir and eoy_categories:
            gar_latest = load_goal_achievement_year(
                goal_achievement_audit_dir, eoy_categories[-1]
            )
            if gar_latest:
                schools_with_exceptions = len(gar_latest.get("school_counts", {}))
        latest_rows = next(
            (value for value in reversed(rows_evaluated_series) if value is not None),
            None,
        )
        pair_label = ""
        if latest:
            pair_label = (
                f"{latest['baseline_school_year']}/{latest['baseline_period']} → "
                f"{latest['current_school_year']}/{latest['current_period']}"
            )
        latest_headline = {
            "exceptions": latest_exceptions,
            "exception_delta": exception_delta,
            "rows_evaluated": latest_rows,
            "schools_with_exceptions": schools_with_exceptions,
            "school_year": eoy_categories[-1],
            "pair_label": pair_label,
        }

    latest_final_goal_headline: dict[str, Any] | None = None
    if gar_recorded_goal_met_pct_series and eoy_categories:
        latest_recorded = next(
            (
                value
                for value in reversed(gar_recorded_goal_met_pct_series)
                if value is not None
            ),
            None,
        )
        latest_expected = next(
            (
                value
                for value in reversed(gar_expected_goal_met_pct_series)
                if value is not None
            ),
            None,
        )
        latest_mismatch = next(
            (value for value in reversed(gar_mismatch_rate_series) if value is not None),
            None,
        )
        prior_recorded = next(
            (
                value
                for value in reversed(gar_recorded_goal_met_pct_series[:-1])
                if value is not None
            ),
            None,
        )
        yoy_recorded_delta = None
        if isinstance(latest_recorded, (int, float)) and isinstance(
            prior_recorded, (int, float)
        ):
            yoy_recorded_delta = round(latest_recorded - prior_recorded, 1)
        gar_global: dict[str, Any] = {}
        if goal_achievement_audit_dir:
            latest_rollup = load_goal_achievement_rollup(
                goal_achievement_audit_dir, eoy_categories[-1]
            )
            gar_global = (latest_rollup or {}).get("global") or {}
        latest_final_goal_headline = {
            "school_year": eoy_categories[-1],
            "recorded_goal_met_pct": latest_recorded,
            "expected_goal_met_pct": latest_expected,
            "mismatch_rate_pct": latest_mismatch,
            "yoy_recorded_goal_met_delta_pct": yoy_recorded_delta,
            "total_rows": gar_global.get("total_rows"),
            "mismatched": gar_global.get("mismatched"),
            "manual_review": gar_global.get("manual_review"),
            "matched": gar_global.get("matched"),
            "special_status": gar_global.get("special_status"),
            "classifiable_rows": gar_global.get("classifiable_rows"),
        }

    latest_goal_headline: dict[str, Any] | None = None
    if latest_goal_summary:
        current = latest_goal_summary.get("current") or {}
        baseline = latest_goal_summary.get("baseline") or {}
        row_level = latest_goal_summary.get("row_level") or {}
        base_pct = baseline.get("on_track_pct")
        curr_pct = current.get("on_track_pct")
        yoy_delta = None
        if isinstance(base_pct, (int, float)) and isinstance(curr_pct, (int, float)):
            yoy_delta = round(curr_pct - base_pct, 1)
        latest_goal_headline = {
            "on_track_pct": curr_pct,
            "yoy_on_track_delta_pct": yoy_delta,
            "eligible_rows": current.get("eligible_rows", 0),
            "rows_improved": row_level.get("improved", 0),
            "rows_worsened": row_level.get("worsened", 0),
            "pair_label": (latest_headline or {}).get("pair_label", ""),
        }

    goal_metric_counts = load_goal_metric_counts_by_year(snapshots_dir, eoy_categories)
    eoy_exception_summary = build_eoy_exception_summary(
        eoy_categories,
        exception_count_series,
        schools_with_exceptions_series,
        goal_metric_counts=goal_metric_counts,
    )
    metrics_audit_cross_year = build_metrics_audit_cross_year_table(
        snapshots_dir,
        eoy_categories,
    )

    return {
        "categories": transition_categories,
        "transition_categories": transition_categories,
        "eoy_categories": eoy_categories,
        "transitions": transitions,
        "global_series": global_series,
        "exception_count_series": exception_count_series,
        "rows_evaluated_series": rows_evaluated_series,
        "issue_code_series": issue_code_series,
        "school_leaderboard": school_leaderboard,
        "latest_headline": latest_headline,
        "eoy_exception_summary": eoy_exception_summary,
        "metrics_audit_cross_year": metrics_audit_cross_year,
        "goal_progress_transitions": goal_progress_transitions,
        "goal_on_track_series": goal_on_track_series,
        "goal_off_track_series": goal_off_track_series,
        "goal_on_track_pct_series": goal_on_track_pct_series,
        "gar_recorded_goal_met_pct_series": gar_recorded_goal_met_pct_series,
        "gar_expected_goal_met_pct_series": gar_expected_goal_met_pct_series,
        "gar_recorded_goal_not_met_pct_series": gar_recorded_goal_not_met_pct_series,
        "gar_expected_goal_not_met_pct_series": gar_expected_goal_not_met_pct_series,
        "gar_mismatch_rate_series": gar_mismatch_rate_series,
        "gar_recorded_composition_series": gar_recorded_composition_series,
        "gar_school_mismatch_leaderboard": gar_school_mismatch_leaderboard,
        "latest_final_goal_headline": latest_final_goal_headline,
        "goal_school_leaderboard": goal_school_leaderboard,
        "latest_goal_headline": latest_goal_headline,
        "pair_count": len(transitions),
        "eoy_snapshot_count": len(eoy_categories),
        "missing_goal_achievement_years": missing_goal_achievement_years,
        "goal_achievement_audit_dir": str(goal_achievement_audit_dir)
        if goal_achievement_audit_dir
        else None,
    }


def build_eoy_exception_summary(
    eoy_categories: list[str],
    exception_count_series: list[int | None],
    schools_with_exceptions_series: list[int | None],
    *,
    goal_metric_counts: dict[str, dict[str, int]] | None = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    prior_exceptions: int | None = None
    counts_by_year = goal_metric_counts or {}
    for index, school_year in enumerate(eoy_categories):
        exceptions = (
            exception_count_series[index]
            if index < len(exception_count_series)
            else None
        )
        year_counts = counts_by_year.get(school_year, {})
        goals_with_baseline_target = year_counts.get("goals_with_baseline_target")
        goal_metrics_total = year_counts.get("goal_metrics_total")
        schools_with_exceptions = (
            schools_with_exceptions_series[index]
            if index < len(schools_with_exceptions_series)
            else None
        )
        exception_delta = None
        if exceptions is not None and prior_exceptions is not None:
            exception_delta = exceptions - prior_exceptions
        if exceptions is not None:
            prior_exceptions = exceptions
        goals_pct = None
        if (
            isinstance(goals_with_baseline_target, int)
            and isinstance(goal_metrics_total, int)
            and goal_metrics_total
        ):
            goals_pct = round(
                100 * goals_with_baseline_target / goal_metrics_total
            )
        rows.append(
            {
                "school_year": school_year,
                "exceptions": exceptions,
                "exception_delta": exception_delta,
                "goals_with_baseline_target": goals_with_baseline_target,
                "goals_with_baseline_target_pct": goals_pct,
                "goal_metrics_total": goal_metrics_total,
                "schools_with_exceptions": schools_with_exceptions,
            }
        )
    return rows


def _on_track_pct_from_rollup(rollup: dict[str, Any] | None) -> float | None:
    if not rollup:
        return None
    global_counts = rollup.get("global") or {}
    on_track = int(global_counts.get("on_track", 0))
    off_track = int(global_counts.get("off_track", 0))
    tracked = on_track + off_track
    if tracked == 0:
        return None
    return round(100 * on_track / tracked, 1)


def _line_chart_value_label(value: float) -> int:
    # WHY: stored rates are rounded to 0.1pp; point labels should round, not truncate
    return round(value)


def _svg_line_chart(
    *,
    width: int,
    height: int,
    categories: list[str],
    series: list[dict[str, Any]],
    y_label: str,
    show_value_labels: Literal["none", "all", "latest"] = "none",
    value_label_offset: int = 8,
) -> str:
    if not categories or not series:
        return ""
    margin = {"top": 24, "right": 16, "bottom": 64, "left": 56}
    plot_w = width - margin["left"] - margin["right"]
    plot_h = height - margin["top"] - margin["bottom"]
    numeric_values = [
        value
        for item in series
        for value in item["data"]
        if isinstance(value, (int, float))
    ]
    if not numeric_values:
        return ""
    y_min = 0
    y_max = max(numeric_values) * 1.1 or 1
    colors = ["#4c8bf5", "#e06c75", "#98c379", "#d19a66"]

    def x_pos(index: int) -> float:
        if len(categories) == 1:
            return margin["left"] + plot_w / 2
        return margin["left"] + (index / (len(categories) - 1)) * plot_w

    def y_pos(value: float) -> float:
        return margin["top"] + plot_h - ((value - y_min) / (y_max - y_min)) * plot_h

    parts = [
        f'<svg viewBox="0 0 {width} {height}" width="100%" role="img" '
        f'aria-label="{y_label}">',
        f'<text x="{margin["left"]}" y="16" class="chart-title">{y_label}</text>',
    ]
    for tick in range(5):
        value = y_min + (y_max - y_min) * tick / 4
        y = y_pos(value)
        parts.append(
            f'<line x1="{margin["left"]}" y1="{y:.1f}" x2="{width - margin["right"]}" '
            f'y2="{y:.1f}" class="grid-line"/>'
        )
        parts.append(
            f'<text x="{margin["left"] - 8}" y="{y + 4:.1f}" class="axis-label" '
            f'text-anchor="end">{int(value)}</text>'
        )
    for index, label in enumerate(categories):
        x = x_pos(index)
        parts.append(
            f'<text x="{x:.1f}" y="{height - 12}" class="axis-label" '
            f'text-anchor="middle" transform="rotate(-35 {x:.1f} {height - 12})">{label}</text>'
        )
    for series_index, item in enumerate(series):
        color = colors[series_index % len(colors)]
        points: list[str] = []
        for index, value in enumerate(item["data"]):
            if value is None:
                continue
            points.append(f"{x_pos(index):.1f},{y_pos(float(value)):.1f}")
        if len(points) >= 2:
            parts.append(
                f'<polyline fill="none" stroke="{color}" stroke-width="2" '
                f'points="{" ".join(points)}"/>'
            )
        labeled_indices: set[int] = set()
        if show_value_labels == "all":
            labeled_indices = {
                index for index, value in enumerate(item["data"]) if value is not None
            }
        elif show_value_labels == "latest":
            for index in range(len(item["data"]) - 1, -1, -1):
                if item["data"][index] is not None:
                    labeled_indices.add(index)
                    break
        for index, value in enumerate(item["data"]):
            if value is None:
                continue
            cx = x_pos(index)
            cy = y_pos(float(value))
            parts.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="3" fill="{color}"/>')
            if index in labeled_indices:
                parts.append(
                    f'<text x="{cx:.1f}" y="{cy - value_label_offset:.1f}" '
                    f'class="data-label" text-anchor="middle">{_line_chart_value_label(float(value))}</text>'
                )
        parts.append(
            f'<text x="{width - margin["right"]}" y="{20 + series_index * 14}" '
            f'class="legend" fill="{color}">■ {item["name"]}</text>'
        )
    parts.append("</svg>")
    return "\n".join(parts)


def _svg_stacked_bar_chart(
    *,
    width: int,
    height: int,
    categories: list[str],
    series: list[dict[str, Any]],
    y_label: str,
    show_totals: bool = False,
) -> str:
    if not categories or not series:
        return ""
    margin = {"top": 24, "right": 16, "bottom": 64, "left": 56}
    plot_w = width - margin["left"] - margin["right"]
    plot_h = height - margin["top"] - margin["bottom"]
    totals = [0] * len(categories)
    for item in series:
        for index, value in enumerate(item["data"]):
            totals[index] += int(value or 0)
    y_max = max(totals) * 1.1 or 1
    bar_w = plot_w / max(len(categories), 1) * 0.7
    colors = ["#4c8bf5", "#e5c07b", "#98c379", "#c678dd", "#56b6c2", "#be5046"]

    def y_pos(value: float) -> float:
        return margin["top"] + plot_h - (value / y_max) * plot_h

    parts = [
        f'<svg viewBox="0 0 {width} {height}" width="100%" role="img" '
        f'aria-label="{y_label}">',
        f'<text x="{margin["left"]}" y="16" class="chart-title">{y_label}</text>',
    ]
    for index, label in enumerate(categories):
        x_center = margin["left"] + (index + 0.5) * (plot_w / len(categories))
        x_left = x_center - bar_w / 2
        stack_base = 0.0
        for series_index, item in enumerate(series):
            value = float(item["data"][index] or 0)
            if value <= 0:
                continue
            y_top = y_pos(stack_base + value)
            y_bottom = y_pos(stack_base)
            height_px = y_bottom - y_top
            color = colors[series_index % len(colors)]
            parts.append(
                f'<rect x="{x_left:.1f}" y="{y_top:.1f}" width="{bar_w:.1f}" '
                f'height="{height_px:.1f}" fill="{color}"/>'
            )
            if height_px >= _SEGMENT_LABEL_MIN_HEIGHT_PX:
                label_fill = "#fff" if color in _STACKED_DARK_FILLS else "#1f2328"
                parts.append(
                    f'<text x="{x_center:.1f}" y="{y_top + height_px / 2 + 3.5:.1f}" '
                    f'class="data-label" fill="{label_fill}" text-anchor="middle">'
                    f'{int(value)}</text>'
                )
            stack_base += value
        if show_totals and totals[index] > 0:
            parts.append(
                f'<text x="{x_center:.1f}" y="{y_pos(totals[index]) - 4:.1f}" '
                f'class="data-label" text-anchor="middle">{totals[index]}</text>'
            )
        parts.append(
            f'<text x="{x_center:.1f}" y="{height - 12}" class="axis-label" '
            f'text-anchor="middle" transform="rotate(-35 {x_center:.1f} {height - 12})">{label}</text>'
        )
    for series_index, item in enumerate(series):
        color = colors[series_index % len(colors)]
        parts.append(
            f'<text x="{width - margin["right"]}" y="{20 + series_index * 14}" '
            f'class="legend" fill="{color}">■ {item["name"]}</text>'
        )
    parts.append("</svg>")
    return "\n".join(parts)


def _svg_horizontal_bar_chart(
    *,
    width: int,
    height: int,
    rows: list[dict[str, Any]],
    y_label: str,
    value_key: str = "regressions",
) -> str:
    if not rows:
        return ""
    margin = {"top": 24, "right": 16, "bottom": 24, "left": 200}
    plot_w = width - margin["left"] - margin["right"]
    row_h = min(28, (height - margin["top"] - margin["bottom"]) / max(len(rows), 1))
    chart_h = margin["top"] + margin["bottom"] + row_h * len(rows)
    values = [abs(float(row[value_key])) for row in rows]
    x_max = max(values) * 1.1 or 1

    parts = [
        f'<svg viewBox="0 0 {width} {chart_h}" width="100%" role="img" '
        f'aria-label="{y_label}">',
        f'<text x="{margin["left"]}" y="16" class="chart-title">{y_label}</text>',
    ]
    for index, row in enumerate(rows):
        y = margin["top"] + index * row_h
        value = abs(float(row[value_key]))
        bar_w = (value / x_max) * plot_w
        parts.append(
            f'<text x="{margin["left"] - 8}" y="{y + row_h * 0.65:.1f}" class="axis-label" '
            f'text-anchor="end">{row["school"][:28]}</text>'
        )
        parts.append(
            f'<rect x="{margin["left"]}" y="{y + 4:.1f}" width="{bar_w:.1f}" '
            f'height="{row_h - 8:.1f}" fill="#e06c75"/>'
        )
        parts.append(
            f'<text x="{margin["left"] + bar_w + 6:.1f}" y="{y + row_h * 0.65:.1f}" '
            f'class="axis-label">{row[value_key]}</text>'
        )
    parts.append("</svg>")
    return "\n".join(parts)


def render_cross_year_html(aggregates: dict[str, Any], path: Path) -> None:
    eoy_categories = aggregates.get("eoy_categories") or aggregates.get("categories") or []
    exception_count_series = aggregates.get("exception_count_series") or []
    issue_code_series: dict[str, list[int]] = aggregates.get("issue_code_series") or {}
    global_series = aggregates.get("global_series") or {}
    school_leaderboard = aggregates.get("school_leaderboard") or []
    eoy_exception_summary = aggregates.get("eoy_exception_summary") or []
    gar_recorded_goal_met_pct_series = aggregates.get("gar_recorded_goal_met_pct_series") or []
    gar_expected_goal_met_pct_series = aggregates.get("gar_expected_goal_met_pct_series") or []
    gar_recorded_goal_not_met_pct_series = aggregates.get("gar_recorded_goal_not_met_pct_series") or []
    gar_expected_goal_not_met_pct_series = aggregates.get("gar_expected_goal_not_met_pct_series") or []
    gar_mismatch_rate_series = aggregates.get("gar_mismatch_rate_series") or []
    gar_recorded_composition_series = aggregates.get("gar_recorded_composition_series") or {}
    gar_school_mismatch_leaderboard = aggregates.get("gar_school_mismatch_leaderboard") or []
    final_goal_headline = aggregates.get("latest_final_goal_headline") or {}
    transitions = aggregates.get("transitions") or []
    missing_goal_achievement_years = aggregates.get("missing_goal_achievement_years") or []
    goal_achievement_audit_dir = aggregates.get("goal_achievement_audit_dir") or ""

    exception_series = [
        {
            "name": "Goal achievement exceptions",
            "data": list(exception_count_series),
        }
    ]
    composition_series = [
        {
            "name": GOAL_ACHIEVEMENT_EXCEPTION_LABELS.get(
                exception_type, exception_type.replace("_", " ")
            ),
            "data": values,
        }
        for exception_type, values in issue_code_series.items()
    ]
    global_chart_series = [
        {
            "name": GLOBAL_METRIC_LABELS.get(metric, metric),
            "data": global_series.get(metric, []),
        }
        for metric in GLOBAL_METRICS
    ]
    gar_goal_met_series = [
        {"name": "Recorded Goal Met %", "data": gar_recorded_goal_met_pct_series},
        {"name": "Expected Goal Met %", "data": gar_expected_goal_met_pct_series},
    ]
    gar_goal_not_met_series = [
        {"name": "Recorded Goal Not Met %", "data": gar_recorded_goal_not_met_pct_series},
        {"name": "Expected Goal Not Met %", "data": gar_expected_goal_not_met_pct_series},
    ]
    gar_mismatch_series = [
        {"name": "Mismatch rate %", "data": gar_mismatch_rate_series},
    ]
    gar_recorded_composition_chart_series = [
        {
            "name": GAR_RECORDED_OUTCOME_LABELS[outcome],
            "data": gar_recorded_composition_series.get(outcome, []),
        }
        for outcome in GAR_RECORDED_OUTCOME_LABELS
    ]
    gar_goal_met_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=eoy_categories,
        series=gar_goal_met_series,
        y_label="Goal Met % (recorded vs expected) at each EOY",
        show_value_labels="all",
    )
    gar_goal_not_met_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=eoy_categories,
        series=gar_goal_not_met_series,
        y_label="Goal Not Met % (recorded vs expected) at each EOY",
        show_value_labels="all",
    )
    gar_recorded_composition_svg = _svg_stacked_bar_chart(
        width=900,
        height=340,
        categories=eoy_categories,
        series=gar_recorded_composition_chart_series,
        y_label="Recorded final goal achievement outcomes at each EOY",
        show_totals=True,
    )
    gar_mismatch_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=eoy_categories,
        series=gar_mismatch_series,
        y_label="Recorded vs expected mismatch rate at each EOY",
        show_value_labels="all",
    )
    gar_mismatch_school_svg = _svg_horizontal_bar_chart(
        width=900,
        height=360,
        rows=gar_school_mismatch_leaderboard,
        y_label="Schools with the most recorded vs expected mismatches (latest EOY)",
        value_key="mismatched",
    )
    exception_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=eoy_categories,
        series=exception_series,
        y_label="Goal achievement exceptions at each EOY",
        show_value_labels="all",
    )
    composition_svg = _svg_stacked_bar_chart(
        width=900,
        height=340,
        categories=eoy_categories,
        series=composition_series,
        y_label="Goal achievement exceptions by type at each EOY",
        show_totals=True,
    )
    global_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=eoy_categories,
        series=global_chart_series,
        y_label="Network-wide flag totals at each EOY",
        show_value_labels="all",
    )
    school_rows = school_leaderboard
    school_value_key = "exceptions"
    if school_rows and "exceptions" not in school_rows[0]:
        school_value_key = "regressions"
    school_svg = _svg_horizontal_bar_chart(
        width=900,
        height=360,
        rows=school_rows,
        y_label="Schools with the most exceptions (latest EOY)",
        value_key=school_value_key,
    )

    school_table_rows = "".join(
        f"<tr><td>{row['school']}</td><td>{row.get('exceptions', row.get('regressions', 0))}</td>"
        f"<td><code>{row['top_metric']}</code></td></tr>"
        for row in school_leaderboard
    )

    gar_mismatch_school_table_rows = "".join(
        f"<tr><td>{row['school']}</td><td>{row['mismatched']}</td>"
        f"<td>{row.get('recorded_goal_met_pct', 'n/a')}%</td>"
        f"<td>{row.get('expected_goal_met_pct', 'n/a')}%</td>"
        f"<td>{row.get('mismatch_rate_pct', 'n/a')}%</td></tr>"
        for row in gar_school_mismatch_leaderboard
    )

    audit_headline_html = ""
    gar_exception_summary_html = ""
    metrics_audit_table = aggregates.get("metrics_audit_cross_year") or {}
    metrics_audit_html = render_metrics_audit_cross_year_table(metrics_audit_table)
    if metrics_audit_html:
        audit_headline_html = metrics_audit_html
    if eoy_exception_summary:
        latest_year = eoy_categories[-1] if eoy_categories else ""
        summary_table_rows = []
        for row in eoy_exception_summary:
            school_year = row["school_year"]
            row_class = ' class="latest-row"' if school_year == latest_year else ""
            exceptions = row.get("exceptions")
            exceptions_text = (
                f"{exceptions:,}" if isinstance(exceptions, int) else "n/a"
            )
            delta = row.get("exception_delta")
            if isinstance(delta, int):
                delta_class = "danger" if delta > 0 else ""
                delta_text = f"{delta:+d}"
            else:
                delta_class = ""
                delta_text = "n/a"
            rows_evaluated = row.get("goals_with_baseline_target", row.get("rows_evaluated"))
            goals_pct = row.get("goals_with_baseline_target_pct")
            if isinstance(rows_evaluated, int):
                rows_text = (
                    f"{rows_evaluated:,} ({goals_pct}%)"
                    if isinstance(goals_pct, int)
                    else f"{rows_evaluated:,}"
                )
            else:
                rows_text = "n/a"
            schools = row.get("schools_with_exceptions")
            schools_text = schools if isinstance(schools, int) else "n/a"
            summary_table_rows.append(
                f"<tr{row_class}><td>{school_year}</td>"
                f"<td>{exceptions_text}</td>"
                f'<td class="{delta_class}">{delta_text}</td>'
                f"<td>{rows_text}</td>"
                f"<td>{schools_text}</td></tr>"
            )
        gar_exception_summary_html = f"""
  <table>
    <thead><tr><th>School Year</th><th>Goal achievement exceptions</th><th>Change vs prior EOY</th><th>Baseline + target present (% of goal metrics)</th><th>Schools with exceptions</th></tr></thead>
    <tbody>{"".join(summary_table_rows)}</tbody>
  </table>
"""
        if not metrics_audit_html:
            audit_headline_html = gar_exception_summary_html
            gar_exception_summary_html = ""

    final_goal_headline_html = ""
    if final_goal_headline:
        yoy = final_goal_headline.get("yoy_recorded_goal_met_delta_pct")
        yoy_text = f"{yoy:+.1f}pp" if isinstance(yoy, (int, float)) else "n/a"
        recorded = final_goal_headline.get("recorded_goal_met_pct", "n/a")
        expected = final_goal_headline.get("expected_goal_met_pct", "n/a")
        mismatch = final_goal_headline.get("mismatch_rate_pct", "n/a")
        total_rows = final_goal_headline.get("total_rows")
        mismatched = final_goal_headline.get("mismatched")
        manual_review = final_goal_headline.get("manual_review")
        matched = final_goal_headline.get("matched")
        special_status = final_goal_headline.get("special_status")
        classifiable_rows = final_goal_headline.get("classifiable_rows")
        row_counts_caption = ""
        if isinstance(total_rows, int) and isinstance(mismatched, int) and isinstance(manual_review, int):
            exceptions = mismatched + manual_review
            row_counts_caption = (
                f"{total_rows:,} Goal Progress Detail rows · {mismatched:,} noncompliant · "
                f"{manual_review:,} review required · {exceptions:,} total exceptions "
                f"({mismatched:,} + {manual_review:,})"
            )
            if isinstance(classifiable_rows, int):
                row_counts_caption += (
                    f" · mismatch rate = classifiable rows where recorded ≠ expected ÷ "
                    f"{classifiable_rows:,} classifiable rows"
                )
            if isinstance(matched, int) and isinstance(special_status, int) and special_status:
                compliant = matched + special_status
                row_counts_caption += (
                    f" · {matched:,} matched classifiable rows plus {special_status:,} "
                    f"Exited Within Same Grading Period Enrolled agreements = "
                    f"{compliant:,} report-style compliant"
                )
        final_goal_headline_html = f"""
        <section class="stats">
          <div class="stat"><div class="stat-value">{recorded}%</div><div class="stat-label">Recorded Goal Met (latest EOY)</div></div>
          <div class="stat"><div class="stat-value">{expected}%</div><div class="stat-label">Expected Goal Met (latest EOY)</div></div>
          <div class="stat"><div class="stat-value danger">{mismatch}%</div><div class="stat-label">Mismatch rate</div></div>
          <div class="stat"><div class="stat-value">{yoy_text}</div><div class="stat-label">YoY recorded Goal Met change</div></div>
        </section>
        <p class="caption">{final_goal_headline.get("school_year", "")} · {row_counts_caption or "Goal Progress Detail rows with classifiable recorded and expected outcomes"}</p>
        """

    caveats: list[str] = [
        "<strong>Different rosters each year:</strong> each School Year EOY Student "
        "Metrics file has its own student list. Year-over-year counts are not the same "
        "students followed year over year.",
        "<strong>Audit flag:</strong> a Student Metrics row flagged by "
        "<code>cisiphyus audit metrics</code> (direction mismatch, scale mismatch, "
        "partial goal, etc.).",
        "<strong>Final goal achievement:</strong> the recorded Goal Achievement label "
        "on Goal Progress Detail rows (Goal Met, Goal Not Met With Progress, Goal Not Met "
        "No Progress). Expected outcomes use the same GAR logic as "
        "<code>cisiphyus audit goal-achievement</code>.",
        "<strong>Goal achievement exception:</strong> a goal progress row where recorded "
        "Goal Achievement does not match what baseline, target, and achieved value imply, "
        "or needs manual review.",
    ]
    latest_transition = transitions[-1] if transitions else None
    if (
        latest_transition
        and latest_transition.get("baseline_period") != latest_transition.get("current_period")
    ):
        caveats.append(
            "Latest pair compares "
            f"{latest_transition['baseline_school_year']} {latest_transition['baseline_period']} "
            f"to {latest_transition['current_school_year']} {latest_transition['current_period']} "
            "— periods are not aligned."
        )
        caveats.append(
            "<code>grading_period_fill_delta</code> in the latest pair reflects "
            "when Grading Period data was captured, not student outcomes."
        )
    elif transitions:
        caveats.append("All comparisons use aligned EOY-to-EOY Student Metrics snapshots.")
    if missing_goal_achievement_years:
        missing_labels = ", ".join(missing_goal_achievement_years)
        audit_hint = (
            f" under <code>{goal_achievement_audit_dir}</code>"
            if goal_achievement_audit_dir
            else ""
        )
        caveats.append(
            f"<strong>Missing Goal Achievement audit data:</strong> {missing_labels} "
            f"show <code>n/a</code> for exception counts. Run "
            f"<code>cisiphyus audit goal-achievement --school-year &lt;SY&gt;</code>"
            f"{audit_hint} to generate "
            f"<code>{{school_year}}_GoalAchievement_AUDIT.csv</code> and "
            f"<code>{{school_year}}_GoalAchievement_ROLLUP.json</code>."
        )
    caveats.append(
        "Final goal achievement counts use each School Year Goal Progress Detail export; "
        "they do not follow students who changed schools or left the caseload."
    )
    caveats_html = "".join(f"      <li>{item}</li>\n" for item in caveats)

    goal_section_html = ""
    if final_goal_headline or any(value is not None for value in gar_recorded_goal_met_pct_series):
        goal_section_html = f"""
  <h2>Final goal achievement</h2>
  <p class="caption">Recorded Goal Achievement labels on Goal Progress Detail vs GAR-expected outcomes</p>
  {final_goal_headline_html}
  <h3>Goal Met % over time</h3>
  <div class="chart">{gar_goal_met_svg}</div>
  <h3>Goal Not Met % over time</h3>
  <div class="chart">{gar_goal_not_met_svg}</div>
  <h3>Recorded outcomes over time</h3>
  <div class="chart">{gar_recorded_composition_svg}</div>
  <h3>Mismatch rate over time</h3>
  <div class="chart">{gar_mismatch_svg}</div>
  <h3>Schools with the most mismatches</h3>
  <p class="caption">Latest School Year EOY · recorded Goal Achievement ≠ expected outcome</p>
  <div class="chart">{gar_mismatch_school_svg}</div>
  <table>
    <thead><tr><th>School</th><th>Mismatches</th><th>Recorded Goal Met %</th><th>Expected Goal Met %</th><th>Mismatch rate</th></tr></thead>
    <tbody>{gar_mismatch_school_table_rows}</tbody>
  </table>
"""

    payload_json = json.dumps(aggregates, indent=2)
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>School Year trends</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 24px; color: #1f2328; background: #fff; }}
    h1 {{ font-size: 1.5rem; margin-bottom: 8px; }}
    h2 {{ font-size: 1.1rem; margin: 32px 0 8px; }}
    h3 {{ font-size: 1rem; margin: 24px 0 8px; color: #57606a; }}
    .caption {{ color: #57606a; font-size: 0.85rem; margin: 4px 0 16px; }}
    .warning {{ background: #fff8c5; border: 1px solid #d4a72c; padding: 12px 16px; border-radius: 6px; margin: 16px 0; }}
    .glossary {{ background: #f6f8fa; border: 1px solid #d0d7de; padding: 12px 16px; border-radius: 6px; margin: 16px 0; }}
    .stats {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin: 16px 0; }}
    .stat {{ border: 1px solid #d0d7de; border-radius: 6px; padding: 12px; }}
    .stat-value {{ font-size: 1.6rem; font-weight: 600; }}
    .stat-value.danger {{ color: #cf222e; }}
    td.danger {{ color: #cf222e; font-weight: 600; }}
    .stat-label {{ color: #57606a; font-size: 0.85rem; }}
    .chart {{ margin: 8px 0 24px; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 0.9rem; }}
    th, td {{ border: 1px solid #d0d7de; padding: 8px; text-align: left; }}
    th {{ background: #f6f8fa; }}
    tr.latest-row {{ background: #f6f8fa; font-weight: 600; }}
    .chart-title {{ font: 600 13px system-ui; fill: #1f2328; }}
    .axis-label {{ font: 11px system-ui; fill: #57606a; }}
    .legend {{ font: 11px system-ui; text-anchor: end; }}
    .data-label {{ font: 10px system-ui; fill: #1f2328; }}
    .grid-line {{ stroke: #d0d7de; stroke-width: 1; }}
    .metrics-audit-table {{ border-collapse: collapse; width: 100%; margin: 1rem 0; table-layout: fixed; }}
    .metrics-audit-table th, .metrics-audit-table td {{ border: 1px solid #d0d7de; padding: 0.45rem 0.6rem; text-align: right; }}
    .metrics-audit-table th:first-child, .metrics-audit-table td:first-child {{ text-align: left; }}
    .metrics-audit-table tbody td:first-child:not(.metrics-audit-detail-label) {{ font-weight: 600; }}
    .metrics-audit-label-col {{ width: 30%; }}
    .metrics-audit-detail-label {{ padding-left: 1.5rem; color: #57606a; }}
    .metrics-audit-table .latest-col {{ background: #fff8c5; }}
  </style>
</head>
<body>
  <h1>School Year trends</h1>
  <p class="caption">{aggregates.get("pair_count", 0)} EOY-to-EOY comparisons · {aggregates.get("eoy_snapshot_count", 0)} EOY snapshots</p>

  <div class="glossary">
    <strong>Read this first</strong>
    <ul>
      <li>Each School Year has its <strong>own Student Metrics roster</strong>. Students enter, exit, and change schools — the lists are not the same year to year.</li>
      <li><strong>Audit flag</strong> — a Student Metrics row flagged by <code>cisiphyus audit metrics</code>.</li>
      <li><strong>Final goal achievement</strong> — recorded Goal Achievement label on Goal Progress Detail (Goal Met, With Progress, No Progress) compared to GAR-expected outcomes from baseline, target, and achieved value.</li>
      <li><strong>Goal achievement exception</strong> — recorded Goal Achievement does not match expected logic from baseline, target, and achieved value, or needs manual review. Source: Goal Achievement audit CSVs.</li>
    </ul>
  </div>

  <div class="warning">
    <strong>Caveats</strong>
    <ul>
{caveats_html}    </ul>
  </div>

  <h2>Student metrics audit</h2>
  <p class="caption">Cumulative Student Metrics Summary audit at each School Year EOY — same counts as <code>cisiphyus audit metrics --year all</code></p>
  {audit_headline_html}

{goal_section_html}

  <h2>Goal achievement exceptions</h2>
  <p class="caption">Goal progress rows where recorded Goal Achievement does not match expected logic — from Goal Achievement audit exports. Baseline + target % uses goal metrics only (excludes supplemental non-goal rows).</p>
  {gar_exception_summary_html}

  <h3>Total exceptions over time</h3>
  <p class="caption">Count of goal achievement mismatches and manual-review rows at each School Year EOY</p>
  <div class="chart">{exception_svg}</div>

  <h3>Exceptions by type</h3>
  <p class="caption">Mismatch vs manual review at each EOY</p>
  <div class="chart">{composition_svg}</div>

  <h3>Network-wide totals at each EOY</h3>
  <p class="caption">Affiliate-wide counts for selected flag types at each School Year EOY</p>
  <div class="chart">{global_svg}</div>

  <h3>Schools with the most exceptions</h3>
  <p class="caption">Latest School Year EOY · flagged rows per school</p>
  <div class="chart">{school_svg}</div>
  <table>
    <thead><tr><th>School</th><th>Exceptions</th><th>Most common metric</th></tr></thead>
    <tbody>{school_table_rows}</tbody>
  </table>

  <script type="application/json" id="cross-year-data">{payload_json}</script>
</body>
</html>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
