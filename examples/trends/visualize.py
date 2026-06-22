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
    "baseline_target_direction_mismatch": "Target not better than baseline",
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
        values[metric] = int(raw) if raw is not None else None
    return values


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
    candidate = (
        cisiphyus_root.parent.parent
        / "evaluation"
        / "analysis"
        / "Audits"
        / "Goal Achievement"
    )
    return candidate if candidate.is_dir() else None


def load_goal_achievement_year(audit_dir: Path, school_year: str) -> dict[str, Any] | None:
    path = audit_dir / f"{school_year}_GoalAchievement_AUDIT.csv"
    if not path.is_file():
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
    goal_progress_transitions: list[dict[str, Any]] = []
    missing_goal_achievement_years: list[str] = []

    if snapshots_dir is not None:
        eoy_categories = discover_eoy_snapshot_years(snapshots_dir)
        if transition_years:
            eoy_categories = [year for year in eoy_categories if year in transition_years]
        all_exception_types: set[str] = set()
        year_exception_types: list[dict[str, int]] = []
        for school_year in eoy_categories:
            gar = (
                load_goal_achievement_year(goal_achievement_audit_dir, school_year)
                if goal_achievement_audit_dir
                else None
            )
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
                if goal_achievement_audit_dir is not None:
                    missing_goal_achievement_years.append(school_year)
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

    eoy_exception_summary = build_eoy_exception_summary(
        eoy_categories,
        exception_count_series,
        rows_evaluated_series,
        schools_with_exceptions_series,
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
        "goal_progress_transitions": goal_progress_transitions,
        "goal_on_track_series": goal_on_track_series,
        "goal_off_track_series": goal_off_track_series,
        "goal_on_track_pct_series": goal_on_track_pct_series,
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
    rows_evaluated_series: list[int | None],
    schools_with_exceptions_series: list[int | None],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    prior_exceptions: int | None = None
    for index, school_year in enumerate(eoy_categories):
        exceptions = (
            exception_count_series[index]
            if index < len(exception_count_series)
            else None
        )
        rows_evaluated = (
            rows_evaluated_series[index]
            if index < len(rows_evaluated_series)
            else None
        )
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
        rows.append(
            {
                "school_year": school_year,
                "exceptions": exceptions,
                "exception_delta": exception_delta,
                "rows_evaluated": rows_evaluated,
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
                    f'class="data-label" text-anchor="middle">{int(value)}</text>'
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
    goal_headline = aggregates.get("latest_goal_headline") or {}
    goal_school_leaderboard = aggregates.get("goal_school_leaderboard") or []
    goal_on_track_series = aggregates.get("goal_on_track_series") or []
    goal_off_track_series = aggregates.get("goal_off_track_series") or []
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
    goal_chart_series = [
        {"name": "On track", "data": goal_on_track_series},
        {"name": "Off track", "data": goal_off_track_series},
    ]

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
    goal_svg = _svg_horizontal_bar_chart(
        width=900,
        height=360,
        rows=goal_school_leaderboard,
        y_label="Largest on-track rate drops by School (latest comparison)",
        value_key="delta_pct",
    )
    goal_timeline_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=eoy_categories,
        series=goal_chart_series,
        y_label="Student goal rows on track vs off track at each EOY",
        show_value_labels="all",
    )

    school_table_rows = "".join(
        f"<tr><td>{row['school']}</td><td>{row.get('exceptions', row.get('regressions', 0))}</td>"
        f"<td><code>{row['top_metric']}</code></td></tr>"
        for row in school_leaderboard
    )
    goal_school_table_rows = "".join(
        f"<tr><td>{row['school']}</td><td>{row['delta_pct']}</td>"
        f"<td>{row.get('baseline_on_track_pct', '')}% → {row.get('current_on_track_pct', '')}%</td></tr>"
        for row in goal_school_leaderboard
    )

    audit_headline_html = ""
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
            rows_evaluated = row.get("rows_evaluated")
            rows_text = (
                f"{rows_evaluated:,}"
                if isinstance(rows_evaluated, int)
                else "n/a"
            )
            schools = row.get("schools_with_exceptions")
            schools_text = schools if isinstance(schools, int) else "n/a"
            summary_table_rows.append(
                f"<tr{row_class}><td>{school_year}</td>"
                f"<td>{exceptions_text}</td>"
                f'<td class="{delta_class}">{delta_text}</td>'
                f"<td>{rows_text}</td>"
                f"<td>{schools_text}</td></tr>"
            )
        audit_headline_html = f"""
  <table>
    <thead><tr><th>School Year</th><th>Exceptions</th><th>Change vs prior EOY</th><th>Goals with baseline and target</th><th>Schools with exceptions</th></tr></thead>
    <tbody>{"".join(summary_table_rows)}</tbody>
  </table>
"""

    goal_headline_html = ""
    if goal_headline:
        yoy = goal_headline.get("yoy_on_track_delta_pct")
        yoy_text = f"{yoy:+.1f}pp" if isinstance(yoy, (int, float)) else "n/a"
        goals_tracked = int(goal_headline.get("eligible_rows", 0) or 0)
        goal_headline_html = f"""
        <section class="stats">
          <div class="stat"><div class="stat-value">{goal_headline.get("on_track_pct", "n/a")}%</div><div class="stat-label">On track (latest EOY)</div></div>
          <div class="stat"><div class="stat-value">{yoy_text}</div><div class="stat-label">YoY on-track change</div></div>
          <div class="stat"><div class="stat-value">{goal_headline.get("rows_improved", 0)}</div><div class="stat-label">Goals improved</div></div>
          <div class="stat"><div class="stat-value danger">{goal_headline.get("rows_worsened", 0)}</div><div class="stat-label">Goals fell behind</div></div>
        </section>
        <p class="caption">{goal_headline.get("pair_label", "")} · {goals_tracked:,} student goal rows with Baseline and Target · improved/fell behind = same student, school, goal, and metric year over year</p>
        """

    caveats: list[str] = [
        "<strong>Different rosters each year:</strong> each School Year EOY Student "
        "Metrics file has its own student list. Exception counts are not the same "
        "students followed year over year.",
        "<strong>Exception:</strong> a goal progress row where recorded Goal Achievement "
        "does not match what baseline, target, and achieved value imply, or needs manual review.",
        "<strong>Goal progress:</strong> the only section that matches the same "
        "student + school + goal + metric across years (see caveats below).",
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
            f"<strong>Missing Goal Achievement audit CSV:</strong> {missing_labels} "
            f"show <code>n/a</code> for exception counts. Add "
            f"<code>{{school_year}}_GoalAchievement_AUDIT.csv</code>{audit_hint} "
            f"(from the Student Progress Detail workbook audit) to include them."
        )
    caveats.append(
        "Goal progress does not follow students who changed schools or left the caseload."
    )
    caveats_html = "".join(f"      <li>{item}</li>\n" for item in caveats)

    goal_section_html = ""
    if goal_headline or any(value is not None for value in goal_on_track_series):
        goal_section_html = f"""
  <h2>Student goal progress</h2>
  <p class="caption">Matches students who appear in both years with the same school, goal, and metric</p>
  {goal_headline_html}
  <h3>On track vs off track over time</h3>
  <div class="chart">{goal_timeline_svg}</div>
  <h3>Schools where on-track rate dropped most</h3>
  <p class="caption">Latest School Year comparison · negative change = fewer goals on track vs prior EOY</p>
  <div class="chart">{goal_svg}</div>
  <table>
    <thead><tr><th>School</th><th>On-track change</th><th>Prior EOY → latest EOY</th></tr></thead>
    <tbody>{goal_school_table_rows}</tbody>
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
  </style>
</head>
<body>
  <h1>School Year trends</h1>
  <p class="caption">{aggregates.get("pair_count", 0)} EOY-to-EOY comparisons · {aggregates.get("eoy_snapshot_count", 0)} EOY snapshots</p>

  <div class="glossary">
    <strong>Read this first</strong>
    <ul>
      <li>Each School Year has its <strong>own Student Metrics roster</strong>. Students enter, exit, and change schools — the lists are not the same year to year.</li>
      <li><strong>Exception</strong> — a goal progress row where recorded Goal Achievement does not match the expected result from baseline, target, and achieved value, or needs manual review. Source: Goal Achievement audit CSVs.</li>
      <li><strong>On track</strong> (goal section) — Baseline and Target set; latest Grading Period meets Target (Progress Against Goal).</li>
    </ul>
  </div>

  <div class="warning">
    <strong>Caveats</strong>
    <ul>
{caveats_html}    </ul>
  </div>

  <h2>Goal achievement exceptions</h2>
  <p class="caption">Goal progress rows where recorded Goal Achievement does not match expected logic — from Goal Achievement audit exports</p>
  {audit_headline_html}

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

{goal_section_html}
  <script type="application/json" id="cross-year-data">{payload_json}</script>
</body>
</html>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
