"""Cross-year trend visualization: aggregate artifacts and render HTML report."""

from __future__ import annotations

import csv
import json
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

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


def _school_year_sort_key(label: str) -> tuple[int, int]:
    stripped = label.strip().upper()
    if not stripped.startswith("SY") or "-" not in stripped:
        return (9999, 9999)
    start_text, end_text = stripped[2:].split("-", 1)
    try:
        return (int(start_text), int(end_text))
    except ValueError:
        return (9999, 9999)


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


def build_cross_year_aggregates(output_dir: Path) -> dict[str, Any]:
    summaries = load_cross_year_summaries(output_dir)
    transitions: list[dict[str, Any]] = []
    global_series: dict[str, list[int | None]] = {metric: [] for metric in GLOBAL_METRICS}
    categories: list[str] = []

    for summary in summaries:
        pair_dir = Path(summary["_pair_dir"])
        current_sy = str(summary.get("current_school_year", ""))
        categories.append(current_sy)
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
        global_values = aggregate_global_metrics(pair_dir)
        for metric in GLOBAL_METRICS:
            global_series[metric].append(global_values.get(metric))

    latest = transitions[-1] if transitions else None
    latest_pair_dir = Path(summaries[-1]["_pair_dir"]) if summaries else None
    school_leaderboard = (
        aggregate_school_regressions(latest_pair_dir) if latest_pair_dir else []
    )

    regression_counts = [int(row["regression_count"]) for row in transitions]
    mean_regressions = (
        sum(regression_counts) / len(regression_counts) if regression_counts else 0
    )

    latest_headline: dict[str, Any] | None = None
    if latest:
        movements = int(latest["movement_count"])
        regressions = int(latest["regression_count"])
        by_type = latest["movements_by_type"]
        site_churn = int(by_type.get("site_appeared", 0)) - int(by_type.get("site_removed", 0))
        latest_headline = {
            "movements": movements,
            "regressions": regressions,
            "regression_rate_pct": round(100 * regressions / movements, 1) if movements else 0,
            "site_churn": site_churn,
            "pair_label": (
                f"{latest['baseline_school_year']}/{latest['baseline_period']} → "
                f"{latest['current_school_year']}/{latest['current_period']}"
            ),
        }

    return {
        "categories": categories,
        "transitions": transitions,
        "global_series": global_series,
        "school_leaderboard": school_leaderboard,
        "mean_regressions": round(mean_regressions, 1),
        "latest_headline": latest_headline,
        "pair_count": len(transitions),
    }


def _svg_line_chart(
    *,
    width: int,
    height: int,
    categories: list[str],
    series: list[dict[str, Any]],
    y_label: str,
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
        for index, value in enumerate(item["data"]):
            if value is None:
                continue
            parts.append(
                f'<circle cx="{x_pos(index):.1f}" cy="{y_pos(float(value)):.1f}" r="3" '
                f'fill="{color}"/>'
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
            stack_base += value
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
) -> str:
    if not rows:
        return ""
    margin = {"top": 24, "right": 16, "bottom": 24, "left": 200}
    plot_w = width - margin["left"] - margin["right"]
    row_h = min(28, (height - margin["top"] - margin["bottom"]) / max(len(rows), 1))
    chart_h = margin["top"] + margin["bottom"] + row_h * len(rows)
    x_max = max(int(row["regressions"]) for row in rows) * 1.1 or 1

    parts = [
        f'<svg viewBox="0 0 {width} {chart_h}" width="100%" role="img" '
        f'aria-label="{y_label}">',
        f'<text x="{margin["left"]}" y="16" class="chart-title">{y_label}</text>',
    ]
    for index, row in enumerate(rows):
        y = margin["top"] + index * row_h
        value = int(row["regressions"])
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
            f'class="axis-label">{value}</text>'
        )
    parts.append("</svg>")
    return "\n".join(parts)


def render_cross_year_html(aggregates: dict[str, Any], path: Path) -> None:
    categories = aggregates.get("categories") or []
    transitions = aggregates.get("transitions") or []
    global_series = aggregates.get("global_series") or {}
    school_leaderboard = aggregates.get("school_leaderboard") or []
    headline = aggregates.get("latest_headline") or {}
    mean_regressions = aggregates.get("mean_regressions", 0)

    volume_series = [
        {
            "name": "Movements",
            "data": [int(row["movement_count"]) for row in transitions],
        },
        {
            "name": "Regressions",
            "data": [int(row["regression_count"]) for row in transitions],
        },
    ]
    composition_series = [
        {
            "name": movement_type.replace("_", " "),
            "data": [int(row["movements_by_type"].get(movement_type, 0)) for row in transitions],
        }
        for movement_type in MOVEMENT_TYPE_ORDER
    ]
    global_chart_series = [
        {"name": metric, "data": global_series.get(metric, [])} for metric in GLOBAL_METRICS
    ]

    volume_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=categories,
        series=volume_series,
        y_label="Movement and regression counts by school year",
    )
    composition_svg = _svg_stacked_bar_chart(
        width=900,
        height=340,
        categories=categories,
        series=composition_series,
        y_label="Movements by type (stacked)",
    )
    global_svg = _svg_line_chart(
        width=900,
        height=320,
        categories=categories,
        series=global_chart_series,
        y_label="Global issue row counts (current period)",
    )
    school_svg = _svg_horizontal_bar_chart(
        width=900,
        height=360,
        rows=school_leaderboard,
        y_label="Top schools by regression count (latest pair)",
    )

    school_table_rows = "".join(
        f"<tr><td>{row['school']}</td><td>{row['regressions']}</td>"
        f"<td><code>{row['top_metric']}</code></td></tr>"
        for row in school_leaderboard
    )

    headline_html = ""
    if headline:
        headline_html = f"""
        <section class="stats">
          <div class="stat"><div class="stat-value">{headline.get("movements", 0)}</div><div class="stat-label">Movements</div></div>
          <div class="stat"><div class="stat-value danger">{headline.get("regressions", 0)}</div><div class="stat-label">Regressions</div></div>
          <div class="stat"><div class="stat-value">{headline.get("regression_rate_pct", 0)}%</div><div class="stat-label">Regression rate</div></div>
          <div class="stat"><div class="stat-value">{headline.get("site_churn", 0)}</div><div class="stat-label">Site churn</div></div>
        </section>
        <p class="caption">Latest pair: {headline.get("pair_label", "")}</p>
        """

    payload_json = json.dumps(aggregates, indent=2)
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Cross-year trend report</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 24px; color: #1f2328; background: #fff; }}
    h1 {{ font-size: 1.5rem; margin-bottom: 8px; }}
    h2 {{ font-size: 1.1rem; margin: 32px 0 8px; }}
    .caption {{ color: #57606a; font-size: 0.85rem; margin: 4px 0 16px; }}
    .warning {{ background: #fff8c5; border: 1px solid #d4a72c; padding: 12px 16px; border-radius: 6px; margin: 16px 0; }}
    .stats {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin: 16px 0; }}
    .stat {{ border: 1px solid #d0d7de; border-radius: 6px; padding: 12px; }}
    .stat-value {{ font-size: 1.6rem; font-weight: 600; }}
    .stat-value.danger {{ color: #cf222e; }}
    .stat-label {{ color: #57606a; font-size: 0.85rem; }}
    .chart {{ margin: 8px 0 24px; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 0.9rem; }}
    th, td {{ border: 1px solid #d0d7de; padding: 8px; text-align: left; }}
    th {{ background: #f6f8fa; }}
    .chart-title {{ font: 600 13px system-ui; fill: #1f2328; }}
    .axis-label {{ font: 11px system-ui; fill: #57606a; }}
    .legend {{ font: 11px system-ui; text-anchor: end; }}
    .grid-line {{ stroke: #d0d7de; stroke-width: 1; }}
  </style>
</head>
<body>
  <h1>Cross-year trend report</h1>
  <p class="caption">Source: artifacts/trends/cross_year · {aggregates.get("pair_count", 0)} transitions · mean regressions {mean_regressions}</p>

  <div class="warning">
    <strong>Caveats</strong>
    <ul>
      <li>Latest pair compares SY24-25 EOY to SY25-26 Q2 — periods are not aligned.</li>
      <li><code>grading_period_fill_delta</code> in the latest pair reflects snapshot timing, not data quality decline.</li>
    </ul>
  </div>

  {headline_html}

  <h2>Volume timeline</h2>
  <p class="caption">Y-axis: row count · X-axis: current school year at end of each transition</p>
  <div class="chart">{volume_svg}</div>

  <h2>Movement composition</h2>
  <p class="caption">Stacked movement counts by type per transition</p>
  <div class="chart">{composition_svg}</div>

  <h2>Global issue trajectory</h2>
  <p class="caption">Current-period global issue row counts from trend_movements.csv</p>
  <div class="chart">{global_svg}</div>

  <h2>School regression leaderboard</h2>
  <p class="caption">Latest transition only · regressions per school site</p>
  <div class="chart">{school_svg}</div>
  <table>
    <thead><tr><th>School</th><th>Regressions</th><th>Top issue code</th></tr></thead>
    <tbody>{school_table_rows}</tbody>
  </table>

  <script type="application/json" id="cross-year-data">{payload_json}</script>
</body>
</html>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
