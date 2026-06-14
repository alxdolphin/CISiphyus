#!/usr/bin/env python3
"""
Flag site-level accreditation issues from CISDM accreditation report exports.

Run (fetches from CISDM when needed):
  cisiphyus audit accreditation

Pin a local workbook:
  cisiphyus audit accreditation --workbook path/to/Accreditation_Report.xlsx

Outputs land in artifacts/accreditation/ by default.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from itertools import chain
from pathlib import Path
from typing import Any, Callable, Iterable, Literal

from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent


def default_output_dir() -> Path:
    explicit = (os.environ.get("ACCREDITATION_OUTPUT_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return Path(__file__).resolve().parents[2] / "artifacts" / "accreditation"

# --- cisiphyus fetch (live export when --workbook is omitted) ---

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
    return Path(__file__).resolve().parents[2] / "artifacts" / "accreditation"


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
    # WHY: monitor runs via `audit accreditation`; raw export uses `pull accreditation`
    run_py = cisiphyus_root / "run.py"
    cmd: list[str] = [sys.executable, str(run_py), "pull", CISIPHYUS_REPORT_ID]

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
            f"[accreditation] workbook fresh at {target} "
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
            f"[accreditation] refreshed workbook ({reason}) -> {target}",
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
            f"[accreditation] refresh failed ({reason}): {exc}",
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


# --- load workbook ---

REQUIRED_SHEETS = (
    "Accreditation Site Coordination",
    "Accreditation Tier I",
    "Accreditation Case Management",
)


def to_cell_text(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").split())


def is_coordination_header_row(row_values: tuple[Any, ...]) -> bool:
    if not row_values or len(row_values) < 2:
        return False
    first = to_cell_text(row_values[0])
    second = to_cell_text(row_values[1])
    return first == "Organization" and second == "School"


def disambiguate_headers(labels: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    out: list[str] = []
    for label in labels:
        n = counts.get(label, 0) + 1
        counts[label] = n
        out.append(label if n == 1 else f"{label}__{n}")
    return out


def find_column_index(headers: list[str], *substrings: str) -> int | None:
    """Return first column index whose normalized header contains all substrings."""
    return find_column_index_excluding(headers, *substrings)


def find_column_index_excluding(
    headers: list[str],
    *substrings: str,
    exclude: tuple[str, ...] = (),
) -> int | None:
    needles = [s.lower() for s in substrings]
    banned = [s.lower() for s in exclude]
    for index, header in enumerate(headers):
        normalized = normalize_header(header).lower()
        if not normalized:
            continue
        if any(ban in normalized for ban in banned):
            continue
        if all(needle in normalized for needle in needles):
            return index
    return None


def require_column(headers: list[str], sheet_name: str, *substrings: str) -> int:
    index = find_column_index(headers, *substrings)
    if index is not None:
        return index
    available = [normalize_header(h) for h in headers if h]
    raise ValueError(
        f"Worksheet {sheet_name!r} missing column matching {substrings!r}. "
        f"Available headers: {available}"
    )


def parse_number(value: object) -> tuple[float | None, str | None]:
    text = to_cell_text(value)
    if text is None:
        return None, None
    cleaned = text.replace(",", "").replace("%", "").strip()
    try:
        return float(cleaned), None
    except ValueError:
        return None, f"non_numeric:{text!r}"


def parse_int(value: object) -> tuple[int, str | None]:
    number, warning = parse_number(value)
    if number is None:
        return 0, warning
    return int(number), warning


def parse_sheet_rows(workbook_path: Path, sheet_name: str) -> list[dict[str, Any]]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            available = ", ".join(workbook.sheetnames)
            raise ValueError(f"Worksheet not found: {sheet_name}. Available: {available}")

        sheet = workbook[sheet_name]
        row_iter: Iterable[tuple[Any, ...]] = sheet.iter_rows(values_only=True)
        iterator = iter(row_iter)
        row1 = next(iterator, None)
        if row1 is None:
            return []

        row2 = next(iterator, None)
        if row2 is not None and is_coordination_header_row(row2):
            header_cells = row2
            data_iterator: Iterable[tuple[Any, ...]] = iterator
        else:
            header_cells = row1
            data_iterator = chain([row2], iterator) if row2 is not None else iterator

        raw_headers = [normalize_header(value) for value in header_cells]
        label_indices: list[tuple[int, str]] = [
            (index, header) for index, header in enumerate(raw_headers) if header
        ]
        if not label_indices:
            return []

        labels = [pair[1] for pair in label_indices]
        unique_labels = disambiguate_headers(labels)
        columns = [(label_indices[i][0], unique_labels[i]) for i in range(len(label_indices))]

        records: list[dict[str, Any]] = []
        for excel_row, row_values in enumerate(data_iterator, start=1):
            record: dict[str, Any] = {"_excel_row": excel_row}
            has_data = False
            for col_index, key in columns:
                cell = row_values[col_index] if col_index < len(row_values) else None
                text = to_cell_text(cell)
                record[key] = text
                if text is not None:
                    has_data = True
            if has_data:
                records.append(record)
        return records
    finally:
        workbook.close()


def load_accreditation_workbook(workbook_path: Path) -> dict[str, list[dict[str, Any]]]:
    workbook_path = workbook_path.resolve()
    return {
        sheet_name: parse_sheet_rows(workbook_path, sheet_name)
        for sheet_name in REQUIRED_SHEETS
    }


# --- site flags and export ---

ReportingRule = Literal["each_column", "min_column", "any_column"]

REMEDIATION_HINTS = {
    "sc_reporting_incomplete": (
        "Log site coordination entries for reporting requirement: School Support Team, "
        "School Leadership, and Organization (once per quarter)."
    ),
    "sc_no_support_plan_adjustment": (
        "Add a site coordination entry with the support plan adjusted question answered Yes. "
        "Editing the support plan document alone does not count."
    ),
    "tier1_below_period_count": (
        "Enter at least one Tier I support per marking period in CISDM."
    ),
    "tier1_max_exceeds_enrollment": (
        "Tier I max served on a single entry cannot exceed reported school enrollment."
    ),
    "cm_support_plans_incomplete": (
        "All case-managed students need a completed needs assessment and support plan."
    ),
    "cm_tier23_incomplete": (
        "All case-managed students need at least one Tier II or Tier III support entered."
    ),
    "cm_checkins_incomplete": (
        "All case-managed students need at least one formal check-in."
    ),
}

FLAG_DESCRIPTIONS = {
    "sc_reporting_incomplete": "Reporting entries below threshold",
    "sc_no_support_plan_adjustment": "No site coordination support plan adjustment documented",
    "tier1_below_period_count": "Fewer Tier I supports than marking periods",
    "tier1_max_exceeds_enrollment": "Tier I max served exceeds enrollment",
    "cm_support_plans_incomplete": "Support plans below 100%",
    "cm_tier23_incomplete": "Tier II/III supports below 100%",
    "cm_checkins_incomplete": "Formal check-ins below 100%",
}


@dataclass
class MonitorConfig:
    reporting_rule: ReportingRule = "each_column"
    reporting_threshold: int | None = None
    include_ok_sites: bool = False
    active_only: bool = True


@dataclass
class SiteFlag:
    school: str
    site_id: str | None
    school_district: str | None
    categories: list[str] = field(default_factory=list)
    flag_codes: list[str] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)


def _grading_period_threshold(grading_periods: int, config: MonitorConfig) -> int:
    if config.reporting_threshold is not None:
        return config.reporting_threshold
    return 3 if grading_periods == 3 else 4


def _reporting_values_pass(
    values: list[int],
    threshold: int,
    rule: ReportingRule,
) -> bool:
    if not values:
        return False
    if rule == "each_column":
        return all(value >= threshold for value in values)
    if rule == "min_column":
        return min(values) >= threshold
    return any(value >= threshold for value in values)


def _site_meta(row: dict[str, Any], headers: list[str]) -> tuple[str, str | None, str | None]:
    school_col = require_column(headers, "site", "School")
    site_id_col = find_column_index(headers, "Site ID")
    district_col = find_column_index(headers, "School District")
    school = row.get(headers[school_col]) or "Unknown School"
    site_id = row.get(headers[site_id_col]) if site_id_col is not None else None
    district = row.get(headers[district_col]) if district_col is not None else None
    return school, site_id, district


def _headers_from_rows(rows: list[dict[str, Any]]) -> list[str]:
    if not rows:
        return []
    return [key for key in rows[0] if not key.startswith("_")]


def _normalize_status(value: object) -> str:
    return (to_cell_text(value) or "").casefold()


def filter_active_rows(
    rows: list[dict[str, Any]],
    *,
    active_only: bool,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if not active_only or not rows:
        return rows, {"active_only": active_only, "status_filter_applied": False}

    headers = _headers_from_rows(rows)
    status_index = find_column_index(headers, "School Status")
    if status_index is None:
        return rows, {
            "active_only": active_only,
            "status_filter_applied": False,
            "status_filter_note": "School Status column missing; no rows excluded",
        }

    status_col = headers[status_index]
    filtered = [
        row
        for row in rows
        if _normalize_status(row.get(status_col)) == "active"
    ]
    return filtered, {
        "active_only": True,
        "status_filter_applied": True,
        "sites_before_filter": len(rows),
        "sites_after_filter": len(filtered),
        "sites_excluded_inactive": len(rows) - len(filtered),
    }


def evaluate_site_coordination(
    rows: list[dict[str, Any]],
    config: MonitorConfig,
) -> list[SiteFlag]:
    headers = _headers_from_rows(rows)
    if not headers:
        return []

    gp_col = headers[require_column(headers, "Accreditation Site Coordination", "# of Grading Periods")]
    r_support_col = headers[
        require_column(headers, "Accreditation Site Coordination", "Reporting", "School Support Team")
    ]
    r_lead_col = headers[
        require_column(headers, "Accreditation Site Coordination", "Reporting", "School Leadership")
    ]
    r_org_col = headers[
        require_column(headers, "Accreditation Site Coordination", "Reporting", "Organization")
    ]
    adj_col = headers[
        require_column(headers, "Accreditation Site Coordination", "Adjustments", "School Support Plan")
    ]

    flags: list[SiteFlag] = []
    for row in rows:
        school, site_id, district = _site_meta(row, headers)
        grading_periods, gp_warn = parse_int(row.get(gp_col))
        threshold = _grading_period_threshold(grading_periods or 4, config)

        report_support, rs_warn = parse_int(row.get(r_support_col))
        report_leadership, rl_warn = parse_int(row.get(r_lead_col))
        report_org, ro_warn = parse_int(row.get(r_org_col))
        adjustments, adj_warn = parse_int(row.get(adj_col))

        reporting_values = [report_support, report_leadership, report_org]
        parse_warnings = [w for w in (gp_warn, rs_warn, rl_warn, ro_warn, adj_warn) if w]

        site_flags: list[str] = []
        details: dict[str, Any] = {
            "grading_periods": grading_periods,
            "reporting_threshold": threshold,
            "reporting_support_team": report_support,
            "reporting_leadership": report_leadership,
            "reporting_organization": report_org,
            "support_plan_adjustments": adjustments,
        }
        if parse_warnings:
            details["parse_warnings"] = parse_warnings

        if not _reporting_values_pass(reporting_values, threshold, config.reporting_rule):
            site_flags.append("sc_reporting_incomplete")
            details["reporting_rule"] = config.reporting_rule

        if adjustments < 1:
            site_flags.append("sc_no_support_plan_adjustment")

        if site_flags or config.include_ok_sites:
            flags.append(
                SiteFlag(
                    school=school,
                    site_id=site_id,
                    school_district=district,
                    categories=["site_coordination"],
                    flag_codes=site_flags,
                    details=details,
                )
            )
    return flags


def _grading_periods_by_school(sc_rows: list[dict[str, Any]]) -> dict[str, int]:
    headers = _headers_from_rows(sc_rows)
    if not headers:
        return {}
    school_col = headers[require_column(headers, "site", "School")]
    gp_col = headers[require_column(headers, "site", "# of Grading Periods")]
    out: dict[str, int] = {}
    for row in sc_rows:
        school = row.get(school_col)
        if not school:
            continue
        periods, _ = parse_int(row.get(gp_col))
        out[school] = periods or 4
    return out


def evaluate_tier_i(
    rows: list[dict[str, Any]],
    grading_periods_by_school: dict[str, int],
    config: MonitorConfig,
) -> list[SiteFlag]:
    headers = _headers_from_rows(rows)
    if not headers:
        return []

    enroll_col = headers[require_column(headers, "Accreditation Tier I", "Total School Enrollment")]
    tier1_col = headers[require_column(headers, "Accreditation Tier I", "# of Tier I")]
    max_served_col = headers[
        require_column(headers, "Accreditation Tier I", "Tier I", "Max Served")
    ]

    flags: list[SiteFlag] = []
    for row in rows:
        school, site_id, district = _site_meta(row, headers)
        enrollment, enroll_warn = parse_number(row.get(enroll_col))
        tier1_count, tier1_warn = parse_int(row.get(tier1_col))
        max_served, max_warn = parse_number(row.get(max_served_col))
        grading_periods = grading_periods_by_school.get(school, 4)

        parse_warnings = [w for w in (enroll_warn, tier1_warn, max_warn) if w]
        site_flags: list[str] = []
        details: dict[str, Any] = {
            "grading_periods": grading_periods,
            "tier1_support_count": tier1_count,
            "enrollment": enrollment,
            "tier1_max_served": max_served,
        }
        if parse_warnings:
            details["parse_warnings"] = parse_warnings

        if tier1_count < grading_periods:
            site_flags.append("tier1_below_period_count")

        if enrollment is not None and max_served is not None and max_served > enrollment:
            site_flags.append("tier1_max_exceeds_enrollment")

        if site_flags or config.include_ok_sites:
            flags.append(
                SiteFlag(
                    school=school,
                    site_id=site_id,
                    school_district=district,
                    categories=["tier1"],
                    flag_codes=site_flags,
                    details=details,
                )
            )
    return flags


def _pct_from_row(
    row: dict[str, Any],
    headers: list[str],
    *,
    pct_substrings: tuple[str, ...],
    count_substrings: tuple[str, ...],
    count_exclude: tuple[str, ...],
    total_cm: int,
) -> tuple[float | None, str | None]:
    pct_index = find_column_index(headers, *pct_substrings)
    if pct_index is not None:
        return parse_number(row.get(headers[pct_index]))

    count_index = find_column_index_excluding(
        headers,
        *count_substrings,
        exclude=count_exclude,
    )
    if count_index is not None and total_cm > 0:
        count, warn = parse_int(row.get(headers[count_index]))
        return count / total_cm, warn
    return None, "missing_pct_and_count_columns"


def evaluate_case_management(
    rows: list[dict[str, Any]],
    config: MonitorConfig,
) -> list[SiteFlag]:
    headers = _headers_from_rows(rows)
    if not headers:
        return []

    total_col = headers[
        require_column(headers, "Accreditation Case Management", "Total Case", "Managed Students")
    ]
    pct_sp_col = headers[
        require_column(headers, "Accreditation Case Management", "% of CM Students with Support Plans")
    ]
    pct_checkin_col = headers[
        require_column(
            headers,
            "Accreditation Case Management",
            "% of CM Students with at least one Formal Check-In",
        )
    ]

    flags: list[SiteFlag] = []
    for row in rows:
        school, site_id, district = _site_meta(row, headers)
        total_cm, total_warn = parse_int(row.get(total_col))
        if total_cm <= 0:
            continue

        pct_sp, sp_warn = parse_number(row.get(pct_sp_col))
        pct_tier23, t23_warn = _pct_from_row(
            row,
            headers,
            pct_substrings=("% Students with Tier II/III",),
            count_substrings=("# of CM Students with Tier II/III Supports",),
            count_exclude=("and Support Plan",),
            total_cm=total_cm,
        )
        pct_checkin, ci_warn = parse_number(row.get(pct_checkin_col))

        parse_warnings = [w for w in (total_warn, sp_warn, t23_warn, ci_warn) if w]
        site_flags: list[str] = []
        details: dict[str, Any] = {
            "total_case_managed_students": total_cm,
            "pct_support_plans": pct_sp,
            "pct_tier23_supports": pct_tier23,
            "pct_formal_checkins": pct_checkin,
        }
        if parse_warnings:
            details["parse_warnings"] = parse_warnings

        if pct_sp is not None and pct_sp < 1.0:
            site_flags.append("cm_support_plans_incomplete")
        if pct_tier23 is not None and pct_tier23 < 1.0:
            site_flags.append("cm_tier23_incomplete")
        if pct_checkin is not None and pct_checkin < 1.0:
            site_flags.append("cm_checkins_incomplete")

        if site_flags or config.include_ok_sites:
            flags.append(
                SiteFlag(
                    school=school,
                    site_id=site_id,
                    school_district=district,
                    categories=["case_management"],
                    flag_codes=site_flags,
                    details=details,
                )
            )
    return flags


def evaluate_workbook(workbook_path: Path, config: MonitorConfig) -> dict[str, Any]:
    sheets = load_accreditation_workbook(workbook_path)
    sc_rows, sc_filter = filter_active_rows(
        sheets["Accreditation Site Coordination"],
        active_only=config.active_only,
    )
    ti_rows, _ti_filter = filter_active_rows(
        sheets["Accreditation Tier I"],
        active_only=config.active_only,
    )
    cm_rows, _cm_filter = filter_active_rows(
        sheets["Accreditation Case Management"],
        active_only=config.active_only,
    )

    grading_periods = _grading_periods_by_school(sc_rows)
    sc_flags = evaluate_site_coordination(sc_rows, config)
    ti_flags = evaluate_tier_i(ti_rows, grading_periods, config)
    cm_flags = evaluate_case_management(cm_rows, config)

    def flagged_only(items: list[SiteFlag]) -> list[SiteFlag]:
        if config.include_ok_sites:
            return items
        return [item for item in items if item.flag_codes]

    sc_flagged = flagged_only(sc_flags)
    ti_flagged = flagged_only(ti_flags)
    cm_flagged = flagged_only(cm_flags)

    all_by_school: dict[str, SiteFlag] = {}
    for item in sc_flagged + ti_flagged + cm_flagged:
        key = item.site_id or item.school
        if key not in all_by_school:
            all_by_school[key] = SiteFlag(
                school=item.school,
                site_id=item.site_id,
                school_district=item.school_district,
            )
        merged = all_by_school[key]
        for category in item.categories:
            if category not in merged.categories:
                merged.categories.append(category)
        for code in item.flag_codes:
            if code not in merged.flag_codes:
                merged.flag_codes.append(code)
        merged.details.update(item.details)

    all_flags = sorted(all_by_school.values(), key=lambda f: f.school.lower())

    return {
        "workbook": str(workbook_path.resolve()),
        "run_at": datetime.now(timezone.utc).isoformat(),
        "config": asdict(config),
        "site_filter": sc_filter,
        "counts": {
            "sites_evaluated": sc_filter.get("sites_after_filter", len(sc_rows)),
            "sites_excluded_inactive": sc_filter.get("sites_excluded_inactive", 0),
            "site_coordination_flagged": len(sc_flagged),
            "tier1_flagged": len(ti_flagged),
            "case_management_flagged": len(cm_flagged),
            "total_flagged_sites": len(all_flags),
        },
        "site_coordination": sc_flagged,
        "tier1": ti_flagged,
        "case_management": cm_flagged,
        "all_flags": all_flags,
    }


def _flag_to_csv_row(flag: SiteFlag) -> dict[str, str]:
    return {
        "school": flag.school,
        "site_id": flag.site_id or "",
        "school_district": flag.school_district or "",
        "categories": ";".join(flag.categories),
        "flag_codes": ";".join(flag.flag_codes),
        "details": json.dumps(flag.details, sort_keys=True),
    }


def write_csv(path: Path, flags: list[SiteFlag]) -> None:
    fieldnames = ["school", "site_id", "school_district", "categories", "flag_codes", "details"]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for flag in sorted(flags, key=lambda f: f.school.lower()):
            writer.writerow(_flag_to_csv_row(flag))


def _format_flag_codes(codes: list[str]) -> str:
    parts = []
    for code in codes:
        label = FLAG_DESCRIPTIONS.get(code, code)
        parts.append(f"{code} ({label})")
    return "; ".join(parts)


def write_markdown(path: Path, results: dict[str, Any]) -> None:
    config = results["config"]
    counts = results["counts"]
    lines = [
        "# Accreditation Progress Monitoring Summary",
        "",
        f"- **Workbook:** `{results['workbook']}`",
        f"- **Run at:** {results['run_at']}",
        f"- **Reporting rule:** `{config['reporting_rule']}`",
    ]
    if config.get("reporting_threshold") is not None:
        lines.append(f"- **Reporting threshold override:** {config['reporting_threshold']}")
    if config.get("active_only", True):
        site_filter = results.get("site_filter", {})
        excluded = counts.get("sites_excluded_inactive", 0)
        if site_filter.get("status_filter_applied"):
            lines.append(
                f"- **Site scope:** active only ({counts.get('sites_evaluated', '?')} sites; "
                f"{excluded} inactive excluded)"
            )
        else:
            lines.append("- **Site scope:** active only (School Status column missing; no rows excluded)")
    else:
        lines.append("- **Site scope:** all sites (includes inactive)")
    lines.extend(
        [
            "",
            "## Counts",
            "",
            f"- Active sites evaluated: **{counts.get('sites_evaluated', 'n/a')}**",
            f"- Site Coordination flagged: **{counts['site_coordination_flagged']}**",
            f"- Tier I flagged: **{counts['tier1_flagged']}**",
            f"- Case Management flagged: **{counts['case_management_flagged']}**",
            f"- Total unique flagged sites: **{counts['total_flagged_sites']}**",
            "",
        ]
    )

    sections = [
        ("Site Coordination", results["site_coordination"]),
        ("Tier I", results["tier1"]),
        ("Case Management", results["case_management"]),
    ]
    for title, flags in sections:
        lines.append(f"## {title}")
        lines.append("")
        if not flags:
            lines.append("No flagged sites.")
            lines.append("")
            continue
        lines.append("| School | District | Flags |")
        lines.append("| --- | --- | --- |")
        for flag in sorted(flags, key=lambda f: f.school.lower()):
            district = flag.school_district or ""
            lines.append(f"| {flag.school} | {district} | {_format_flag_codes(flag.flag_codes)} |")
        lines.append("")

    lines.append("## Remediation hints")
    lines.append("")
    for code, hint in REMEDIATION_HINTS.items():
        lines.append(f"- **{code}:** {hint}")
    lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")


def _serialize_flags(flags: list[SiteFlag]) -> list[dict[str, Any]]:
    return [asdict(flag) for flag in flags]


def results_to_json_payload(results: dict[str, Any]) -> dict[str, Any]:
    return {
        "workbook": results["workbook"],
        "run_at": results["run_at"],
        "config": results["config"],
        "site_filter": results.get("site_filter", {}),
        "counts": results["counts"],
        "site_coordination": _serialize_flags(results["site_coordination"]),
        "tier1": _serialize_flags(results["tier1"]),
        "case_management": _serialize_flags(results["case_management"]),
        "all_flags": _serialize_flags(results["all_flags"]),
    }


def write_json_summary(path: Path, results: dict[str, Any]) -> None:
    path.write_text(
        json.dumps(results_to_json_payload(results), indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def export_results(results: dict[str, Any], output_dir: Path) -> dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "sc_flags": output_dir / "sc_flags.csv",
        "tier1_flags": output_dir / "tier1_flags.csv",
        "cm_flags": output_dir / "cm_flags.csv",
        "all_flags": output_dir / "all_flags.csv",
        "monitoring_summary_md": output_dir / "monitoring_summary.md",
        "monitoring_summary_json": output_dir / "monitoring_summary.json",
    }
    write_csv(paths["sc_flags"], results["site_coordination"])
    write_csv(paths["tier1_flags"], results["tier1"])
    write_csv(paths["cm_flags"], results["case_management"])
    write_csv(paths["all_flags"], results["all_flags"])
    write_markdown(paths["monitoring_summary_md"], results)
    write_json_summary(paths["monitoring_summary_json"], results)
    return {key: str(path) for key, path in paths.items()}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Flag site-level accreditation issues from CISDM accreditation exports.",
    )
    parser.add_argument(
        "--workbook",
        action="append",
        type=Path,
        help="Path to an accreditation .xlsx export (repeatable). When omitted, resolves via cisiphyus (age-gated).",
    )
    parser.add_argument(
        "--force-fetch",
        action="store_true",
        help="Always re-export from CISDM even when the local workbook is fresh.",
    )
    parser.add_argument(
        "--fetch-destination",
        type=Path,
        help="Where to save the CISDM export (default: artifacts/accreditation/Accreditation_Report.xlsx).",
    )
    parser.add_argument(
        "--include-inactive-sites",
        action="store_true",
        help="Include inactive sites in monitoring (default: active sites only).",
    )
    parser.add_argument(
        "--reporting-rule",
        choices=["each_column", "min_column", "any_column"],
        default="each_column",
        help="How to evaluate SC reporting columns against threshold.",
    )
    parser.add_argument(
        "--reporting-threshold",
        type=int,
        help="Override per-site reporting threshold (default: 4 or 3 for trimester sites).",
    )
    parser.add_argument(
        "--include-ok-sites",
        action="store_true",
        help="Include sites with no flags in CSV outputs.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print JSON summary to stdout.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = MonitorConfig(
        reporting_rule=args.reporting_rule,
        reporting_threshold=args.reporting_threshold,
        include_ok_sites=args.include_ok_sites,
        active_only=not args.include_inactive_sites,
    )

    workbooks: list[Path] = []
    if args.workbook:
        workbooks.extend(args.workbook)
    else:
        destination = args.fetch_destination or preferred_accreditation_workbook()
        try:
            fetch_result = fetch_accreditation_workbook(
                destination=destination,
                force_fetch=args.force_fetch,
                require_fresh=True,
            )
        except (FileNotFoundError, OSError, RuntimeError) as exc:
            print(f"error: CISDM fetch failed: {exc}", file=sys.stderr)
            return 1
        if not fetch_result.succeeded:
            print(
                f"error: CISDM fetch failed and no local workbook at {destination}",
                file=sys.stderr,
            )
            return 1
        workbooks.append(fetch_result.destination)

    if not workbooks:
        print("error: no workbook available", file=sys.stderr)
        return 1

    out_dir = default_output_dir()
    summaries: list[dict[str, Any]] = []
    for workbook in workbooks:
        workbook = workbook.resolve()
        if not workbook.is_file():
            print(f"error: workbook not found: {workbook}", file=sys.stderr)
            return 1

        results = evaluate_workbook(workbook, config)
        export_results(results, out_dir)
        summaries.append(results)

        print(f"Workbook: {workbook}")
        print(f"Output dir: {out_dir}")
        for key, value in results["counts"].items():
            print(f"  {key}: {value}")

    if args.json:
        payloads = [results_to_json_payload(item) for item in summaries]
        print(json.dumps(payloads if len(payloads) > 1 else payloads[0], indent=2, sort_keys=True))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
