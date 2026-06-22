#!/usr/bin/env python3
"""
Student metrics audit for CISDM student metrics summary exports.

Run (fetches from CISDM when needed):
  cisiphyus audit metrics

Pin a local workbook:
  cisiphyus audit metrics --workbook path/to/StudentMetricsSummary.xlsx

Outputs land in artifacts/audit/ by default.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Literal, TypeAlias

from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent


def default_output_dir() -> Path:
    explicit = (os.environ.get("AUDIT_OUTPUT_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return Path(__file__).resolve().parents[2] / "artifacts" / "audit"

# --- cisiphyus fetch (live export when --workbook is omitted) ---

TRUE_VALUES = frozenset({"1", "true", "yes", "on"})
STUDENT_METRICS_MAX_AGE_HOURS = 24.0
CISIPHYUS_REPORT_ID = "student_metrics_summary"
DEFAULT_STUDENT_METRICS_FILENAME = "SY25-26_StudentMetricsSummary.xlsx"


@dataclass(frozen=True)
class FetchResult:
    destination: Path
    triggered: bool
    succeeded: bool
    reason: str


def _flag_true(raw: str | None) -> bool:
    return (raw or "").strip().lower() in TRUE_VALUES


def _default_cisiphyus_root() -> Path:
    # WHY: prototypical app lives at examples/audit/ inside the cisiphyus repo
    return Path(__file__).resolve().parents[2]


def _default_local_inputs_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "artifacts" / "audit"


def preferred_student_metrics_workbook(
    *,
    local_inputs_dir: Path | None = None,
) -> Path:
    explicit = (os.environ.get("AUDIT_STUDENT_METRICS_WORKBOOK") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    base = Path(
        os.environ.get("AUDIT_LOCAL_INPUTS_DIR", "").strip()
        or str(local_inputs_dir or _default_local_inputs_dir())
    ).expanduser().resolve()
    name = (
        os.environ.get("AUDIT_STUDENT_METRICS_FILENAME", "").strip()
        or DEFAULT_STUDENT_METRICS_FILENAME
    )
    return (base / name).resolve()


def workbook_freshness(workbook_path: Path | None) -> dict[str, Any]:
    max_age = float(
        os.environ.get("AUDIT_MAX_AGE_HOURS", "").strip()
        or STUDENT_METRICS_MAX_AGE_HOURS
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


def _cisiphyus_cmd(cisiphyus_root: Path, school_year: str | None = None) -> list[str]:
    run_py = cisiphyus_root / "run.py"
    cmd: list[str] = [sys.executable, str(run_py), "pull", CISIPHYUS_REPORT_ID]
    headed = _flag_true(
        _env_first("AUDIT_CISPHYUS_HEADED", "CISPHYUS_HEADED")
    )
    if headed:
        cmd.append("--headed")
    if school_year:
        cmd.extend(["--school-year", school_year])
    return cmd


def _cisiphyus_raw_path(cisiphyus_root: Path, school_year: str | None = None) -> Path:
    if school_year:
        src_dir = cisiphyus_root / "src"
        if str(src_dir) not in sys.path:
            sys.path.insert(0, str(src_dir))
        import config

        programs = config.load_school_year_programs(config.REPORTS_PATH)
        default_sy = config.default_school_year(programs)
        if default_sy and school_year != default_sy:
            return config.archives_raw_path(school_year, CISIPHYUS_REPORT_ID)
    return cisiphyus_root / "artifacts" / "latest" / CISIPHYUS_REPORT_ID / "raw.xlsx"


def _run_cisiphyus_export(
    *,
    cisiphyus_root: Path,
    school_year: str | None = None,
    run: Callable[..., Any],
) -> Path:
    run_py = cisiphyus_root / "run.py"
    if not run_py.is_file():
        raise FileNotFoundError(f"cisiphyus run.py not found at {run_py}")

    completed = run(
        _cisiphyus_cmd(cisiphyus_root, school_year=school_year),
        cwd=str(cisiphyus_root),
        check=False,
        env=os.environ.copy(),
    )
    if getattr(completed, "returncode", 1) != 0:
        raise RuntimeError(
            f"cisiphyus {CISIPHYUS_REPORT_ID} failed with exit code "
            f"{getattr(completed, 'returncode', 'unknown')}"
        )

    raw = _cisiphyus_raw_path(cisiphyus_root, school_year=school_year)
    if not raw.is_file():
        raise FileNotFoundError(f"expected cisiphyus output missing: {raw}")
    return raw


def fetch_student_metrics_workbook(
    *,
    destination: Path | None = None,
    force_fetch: bool = False,
    require_fresh: bool = False,
    school_year: str | None = None,
    run: Callable[..., Any] = subprocess.run,
) -> FetchResult:
    """Fetch student metrics summary from CISDM when missing, stale, or forced."""
    target = (destination or preferred_student_metrics_workbook()).resolve()
    env_force = _flag_true(os.environ.get("AUDIT_FETCH_FROM_CISDM"))
    should_fetch, reason = _fetch_reason(destination=target, force_fetch=force_fetch or env_force)

    if not should_fetch:
        print(
            f"[audit] workbook fresh at {target} "
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
        _env_first("AUDIT_CISPHYUS_ROOT", "CISIPHYUS_ROOT")
        or str(_default_cisiphyus_root())
    ).expanduser().resolve()

    try:
        raw = _run_cisiphyus_export(
            cisiphyus_root=cisiphyus_root,
            school_year=school_year,
            run=run,
        )
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(raw, target)

        if require_fresh:
            freshness = workbook_freshness(target)
            if freshness.get("is_stale"):
                age = freshness.get("workbook_age_hours")
                raise RuntimeError(
                    "stale_student_metrics_workbook_after_refresh: "
                    f"path={target} workbook_age_hours={age} "
                    f"max_age_hours={freshness.get('max_age_hours')}"
                )

        print(
            f"[audit] refreshed workbook ({reason}) -> {target}",
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
                    f"stale_student_metrics_workbook_after_failed_refresh: {exc}"
                ) from exc
        print(
            f"[audit] refresh failed ({reason}): {exc}",
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

HEADER_ROW = 1
DEFAULT_SHEET = "Sheet1"

REQUIRED_HEADERS: tuple[str, ...] = (
    "Organization",
    "School",
    "Student ID",
    "Client ID",
    "Goal",
    "Metric",
    "Baseline",
    "Target",
    "Enrollment Status",
    "School Year",
)

COMPOSITE_KEY_FIELDS = ("Student ID", "School", "School Year", "Goal", "Metric")

GRADING_PERIOD_HEADERS = (
    "1st Grading Period",
    "2nd Grading Period",
    "3rd Grading Period",
    "4th Grading Period",
    "5th Grading Period",
    "6th Grading Period",
)

RecordValue: TypeAlias = str | None | int
RowRecord: TypeAlias = dict[str, RecordValue]


def to_cell_text(value: object) -> str | None:
    if value is None:
        return None
    return str(value)


def row_number(record: RowRecord) -> int:
    value = record.get("row_number")
    if isinstance(value, int):
        return value
    if isinstance(value, str) and value.isdigit():
        return int(value)
    raise KeyError("record missing integer row_number")


def validate_headers(headers: list[str | None], sheet_name: str) -> None:
    present = {h for h in headers if h}
    missing = [h for h in REQUIRED_HEADERS if h not in present]
    if missing:
        raise ValueError(f"Worksheet {sheet_name} is missing required headers: {', '.join(missing)}")


def load_rows(workbook_path: Path, sheet_name: str = DEFAULT_SHEET) -> list[RowRecord]:
    # WHY: data_only matches displayed cell values, not formula strings
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        workbook.close()
        raise ValueError(f"Worksheet not found: {sheet_name}. Available: {available}")
    sheet = workbook[sheet_name]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    if len(rows) < HEADER_ROW:
        raise ValueError(f"Worksheet {sheet_name} has no header row")
    headers = [to_cell_text(v) for v in rows[HEADER_ROW - 1]]
    validate_headers(headers, sheet_name)
    records: list[RowRecord] = []
    for excel_row, row_values in enumerate(rows[HEADER_ROW:], start=HEADER_ROW + 1):
        record: RowRecord = {"row_number": excel_row}
        for index, header in enumerate(headers):
            if not header:
                continue
            cell = row_values[index] if index < len(row_values) else None
            record[header] = to_cell_text(cell)
        records.append(record)
    return records


def load_student_metrics_workbook(
    workbook_path: Path,
    *,
    sheet_name: str = DEFAULT_SHEET,
) -> list[RowRecord]:
    return load_rows(workbook_path, sheet_name=sheet_name)

# --- validity rules ---

Direction = Literal["higher_is_better", "lower_is_better", "none"]

GOAL_DOMAINS: dict[str, str] = {
    "Improve Attendance": "Attendance",
    "Improve School Behavior": "Behavior",
    "Improve Academics": "Academics",
    "Improve Social-Emotional Learning": "SEL",
    "Improve Career Readiness": "Career Readiness",
    "Improve College Readiness": "College Readiness",
    "Improve High Risk Behavior": "High Risk Behavior",
}

METRIC_DOMAIN_OVERRIDES: dict[str, frozenset[str]] = {
    "Attendance Rate (%)": frozenset({"Attendance"}),
    "Attendance Rate (days absent)": frozenset({"Attendance"}),
    "Tardies": frozenset({"Attendance", "Behavior"}),
    "Suspensions": frozenset({"Behavior"}),
    "Disciplinary Referrals": frozenset({"Behavior"}),
    "Conduct": frozenset({"Behavior"}),
    "Other Behavior Incidents": frozenset({"Behavior"}),
    "Credits Needed/Credit Completion": frozenset({"Academics"}),
    "Reading Level": frozenset({"Academics"}),
    "GPA": frozenset({"Academics"}),
    "Overall SEAD assessment score": frozenset({"SEL"}),
    "Self Control (SEAD)": frozenset({"SEL"}),
    "Social Awareness (SEAD)": frozenset({"SEL"}),
    "Self-Perception (SEAD)": frozenset({"SEL"}),
    "Social Support (SEAD)": frozenset({"SEL"}),
    "Global Engagement (SEAD)": frozenset({"SEL"}),
    "Social Engagement": frozenset({"SEL"}),
    "Behavior Engagement": frozenset({"SEL"}),
    "Emotional Engagement": frozenset({"SEL"}),
    "Global Engagement (Engagement Survey)": frozenset({"SEL"}),
    "Cognitive Engagement": frozenset({"SEL"}),
    "Other (SEL)": frozenset({"SEL"}),
    "Other High Risk Behavior": frozenset({"High Risk Behavior"}),
    "Drank Alcohol (# of times past 30 days)": frozenset({"High Risk Behavior"}),
    "Violent/bullied someone (# of times past 30 days)": frozenset({"High Risk Behavior"}),
    "CTE Completion": frozenset({"Academics"}),
    "Complete career assessment": frozenset({"Career Readiness"}),
    "Accept a position of employment": frozenset({"Career Readiness"}),
    "Interview with one or more potential employers": frozenset({"Career Readiness"}),
    "Create a resume": frozenset({"Career Readiness"}),
    "Other Career Readiness": frozenset({"Career Readiness"}),
    "Apply to one or more colleges/universities": frozenset({"College Readiness"}),
    "Accepted to one or more colleges/universities": frozenset({"College Readiness"}),
    "Other College Readiness": frozenset({"College Readiness"}),
}

GRADE_SCALE: frozenset[str] = frozenset(
    {"F", "D-", "D", "D+", "C-", "C", "C+", "B-", "B", "B+", "A-", "A", "A+"}
)

GRADE_ORDER: tuple[str, ...] = (
    "F",
    "D-",
    "D",
    "D+",
    "C-",
    "C",
    "C+",
    "B-",
    "B",
    "B+",
    "A-",
    "A",
    "A+",
)

GRADE_RANK: dict[str, int] = {grade: index for index, grade in enumerate(GRADE_ORDER)}


READING_LEVEL_TERMS: frozenset[str] = frozenset(
    {
        "below",
        "below benchmark",
        "well below benchmark",
        "well-below benchmark",
        "at benchmark",
        "at or above benchmark",
        "above benchmark",
        "at grade level",
        "below grade level",
        "expected reading level",
        "expected grade reading level",
        "at or above grade level",
    }
)

READING_LEVEL_NORMALIZATION: dict[str, str] = {
    "well-below benchmark": "well below benchmark",
    "below reading level": "below grade level",
    "expected grade reading level": "expected reading level",
    "proficient": "at benchmark",
    "at risk": "below benchmark",
}

ENGAGEMENT_TERMS: frozenset[str] = frozenset(
    {
        "lower engagement",
        "low engagement",
        "moderate engagement",
        "higher engagement",
        "high engagement",
    }
)

BOOLEAN_TERMS: frozenset[str] = frozenset({"yes", "no"})

ANNOTATION_EXEMPT_SCALES: frozenset[str] = frozenset(
    {"reading_level", "engagement_text", "numeric_or_engagement", "boolean"}
)

PAREN_ANNOTATION_RE = re.compile(r"\([^)]+\)")


def clean_value(value: str | None) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def parse_number(value: str | None) -> float | None:
    text = clean_value(value)
    if text is None:
        return None
    cleaned = text.replace(",", "").replace("%", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def normalize_text(value: str | None) -> str | None:
    text = clean_value(value)
    if text is None:
        return None
    return text.lower()


def normalize_grade(value: str | None) -> str | None:
    text = clean_value(value)
    if text is None:
        return None
    return text.split("(")[0].strip().upper()


def normalize_reading_level(value: str | None) -> str | None:
    text = normalize_text(value)
    if text is None:
        return None
    canonical = " ".join(text.replace("-", " ").split())
    if canonical.startswith("level "):
        suffix = canonical.removeprefix("level ").strip()
        if parse_number(suffix) is not None:
            return suffix
    if canonical.endswith(" letters"):
        prefix = canonical.removesuffix(" letters").strip()
        if parse_number(prefix) is not None:
            return prefix
    return READING_LEVEL_NORMALIZATION.get(canonical, canonical)


def goal_domain(goal: str | None) -> str | None:
    return GOAL_DOMAINS.get(clean_value(goal) or "")


def metric_domains(metric: str | None) -> frozenset[str]:
    metric_name = clean_value(metric)
    if metric_name is None:
        return frozenset()
    if metric_name in METRIC_DOMAIN_OVERRIDES:
        return METRIC_DOMAIN_OVERRIDES[metric_name]
    lowered = metric_name.lower()
    if "attendance rate" in lowered:
        return frozenset({"Attendance"})
    if "suspension" in lowered or "referral" in lowered or "incident" in lowered or "conduct" in lowered:
        return frozenset({"Behavior"})
    if "grade" in lowered or "gpa" in lowered or "reading" in lowered or "credit" in lowered or "score" in lowered:
        return frozenset({"Academics"})
    if "sead" in lowered or "engagement" in lowered or "(sel)" in lowered:
        return frozenset({"SEL"})
    if "career" in lowered or "resume" in lowered or "employment" in lowered:
        return frozenset({"Career Readiness"})
    if "college" in lowered or "universit" in lowered:
        return frozenset({"College Readiness"})
    if (
        "tobacco" in lowered
        or "alcohol" in lowered
        or "high risk" in lowered
        or "bullied" in lowered
        or "violent" in lowered
    ):
        return frozenset({"High Risk Behavior"})
    return frozenset()


def metric_scale(metric: str | None) -> str | None:
    metric_name = clean_value(metric)
    if metric_name is None:
        return None
    lowered = metric_name.lower()
    if metric_name == "Attendance Rate (%)":
        return "percent"
    if (
        "days absent" in lowered
        or "suspension" in lowered
        or "referral" in lowered
        or "incident" in lowered
        or "tard" in lowered
        or "conduct" in lowered
    ):
        return "count"
    if "core course grades" in lowered:
        return "grade_or_numeric"
    if metric_name == "GPA" or "score" in lowered:
        return "numeric"
    if metric_name == "Reading Level":
        return "reading_level"
    if "credits needed" in lowered or "credit completion" in lowered:
        return "numeric"
    if metric_name == "Other (SEL)":
        return "numeric_or_engagement"
    if metric_name in {
        "Complete career assessment",
        "Accept a position of employment",
        "Interview with one or more potential employers",
        "Create a resume",
        "Apply to one or more colleges/universities",
        "Accepted to one or more colleges/universities",
        "CTE Completion",
    }:
        return "boolean"
    if metric_name in {"Other Career Readiness", "Other College Readiness"}:
        return "numeric"
    if "emotional engagement" in lowered:
        return "engagement_text"
    if "engagement" in lowered or "sead" in lowered:
        return "numeric_or_engagement"
    if "alcohol" in lowered or "high risk" in lowered or "bullied" in lowered or "violent" in lowered:
        return "count"
    return None


def metric_direction(metric: str | None) -> Direction:
    scale = metric_scale(metric)
    if scale in {"boolean", "engagement_text", "reading_level"}:
        return "none"
    if scale == "numeric_or_engagement":
        return "higher_is_better"
    if scale == "count":
        return "lower_is_better"
    if scale == "percent":
        return "higher_is_better"
    if scale == "numeric":
        return "higher_is_better"
    if scale == "grade_or_numeric":
        return "higher_is_better"
    return "none"


def value_matches_scale(metric: str | None, value: str | None) -> bool:
    text = clean_value(value)
    if text is None:
        return True
    scale = metric_scale(metric)
    if scale is None:
        return False
    lowered = normalize_text(text) or ""
    if scale == "percent":
        return parse_number(text) is not None and "day" not in lowered and normalize_grade(text) not in GRADE_SCALE
    if scale == "count":
        return parse_number(text) is not None
    if scale == "numeric":
        return parse_number(text) is not None
    if scale == "grade_or_numeric":
        return parse_number(text) is not None or normalize_grade(text) in GRADE_SCALE
    if scale == "reading_level":
        normalized = normalize_reading_level(text)
        return normalized in READING_LEVEL_TERMS or parse_number(normalized) is not None
    if scale == "engagement_text":
        return lowered in ENGAGEMENT_TERMS
    if scale == "numeric_or_engagement":
        return parse_number(text) is not None or lowered in ENGAGEMENT_TERMS
    if scale == "boolean":
        return lowered in BOOLEAN_TERMS
    return False


def has_undocumented_annotation(metric: str | None, value: str | None) -> bool:
    text = clean_value(value)
    if text is None:
        return False
    scale = metric_scale(metric)
    if scale is None or scale in ANNOTATION_EXEMPT_SCALES:
        return False
    return bool(PAREN_ANNOTATION_RE.search(text))


def _numeric_direction_ok(
    baseline: float,
    target: float,
    direction: Direction,
) -> bool:
    if direction == "higher_is_better":
        return target > baseline
    if direction == "lower_is_better":
        return target < baseline
    return True


def _grade_direction_ok(
    baseline_grade: str,
    target_grade: str,
    direction: Direction,
) -> bool:
    baseline_rank = GRADE_RANK.get(baseline_grade)
    target_rank = GRADE_RANK.get(target_grade)
    if baseline_rank is None or target_rank is None:
        return True
    if direction == "higher_is_better":
        return target_rank > baseline_rank
    if direction == "lower_is_better":
        return target_rank < baseline_rank
    return True


def baseline_target_direction_mismatch(
    metric: str | None,
    baseline: str | None,
    target: str | None,
) -> bool:
    direction = metric_direction(metric)
    if direction == "none":
        return False
    baseline_text = clean_value(baseline)
    target_text = clean_value(target)
    if baseline_text is None or target_text is None:
        return False

    baseline_num = parse_number(baseline_text)
    target_num = parse_number(target_text)
    if baseline_num is not None and target_num is not None:
        return not _numeric_direction_ok(baseline_num, target_num, direction)

    baseline_grade = normalize_grade(baseline_text)
    target_grade = normalize_grade(target_text)
    if (
        baseline_grade in GRADE_SCALE
        and target_grade in GRADE_SCALE
        and baseline_grade is not None
        and target_grade is not None
    ):
        return not _grade_direction_ok(baseline_grade, target_grade, direction)
    return False


def analyze_record(record: RowRecord) -> dict[str, object]:
    goal = record.get("Goal")
    metric = record.get("Metric")
    baseline = record.get("Baseline")
    target = record.get("Target")
    issue_codes: list[str] = []

    expected_domain = goal_domain(str(goal) if goal is not None else None)
    allowed_domains = metric_domains(str(metric) if metric is not None else None)
    if expected_domain and allowed_domains and expected_domain not in allowed_domains:
        issue_codes.append("goal_metric_domain_mismatch")
    elif expected_domain is None or not allowed_domains:
        issue_codes.append("manual_review_domain")

    metric_str = str(metric) if metric is not None else None
    baseline_str = str(baseline) if baseline is not None else None
    target_str = str(target) if target is not None else None

    if not value_matches_scale(metric_str, baseline_str):
        issue_codes.append("baseline_scale_mismatch")
    if not value_matches_scale(metric_str, target_str):
        issue_codes.append("target_scale_mismatch")

    if has_undocumented_annotation(metric_str, baseline_str):
        issue_codes.append("undocumented_value_annotation")
    if has_undocumented_annotation(metric_str, target_str):
        issue_codes.append("undocumented_value_annotation")

    if baseline_target_direction_mismatch(metric_str, baseline_str, target_str):
        issue_codes.append("baseline_target_direction_mismatch")

    return {
        "school": clean_value(str(record.get("School")) if record.get("School") is not None else None),
        "goal": clean_value(str(goal) if goal is not None else None),
        "metric": clean_value(metric_str),
        "expected_domain": expected_domain,
        "allowed_domains": sorted(allowed_domains),
        "issue_codes": issue_codes,
    }

# --- run checks and write outputs ---

FLAG_DESCRIPTIONS: dict[str, str] = {
    "goal_metric_domain_mismatch": "Goal domain does not match metric",
    "manual_review_domain": "Goal or metric domain needs manual review",
    "baseline_scale_mismatch": "Baseline value does not match metric scale",
    "target_scale_mismatch": "Target value does not match metric scale",
    "baseline_target_direction_mismatch": "Target is not better than baseline for metric family",
    "undocumented_value_annotation": "Value contains undocumented parenthetical annotation",
    "both_baseline_and_target_blank": "Both baseline and target are blank",
    "both_blank_with_two_progress_reports": "Both blank but two or more grading periods filled",
    "target_without_baseline": "Target set without baseline",
    "baseline_without_target": "Baseline set without target",
    "baseline_without_target_non_goal_context": "Baseline without target in non-goal context",
    "baseline_without_target_no_goal_context": (
        "Baseline without target (no goal context — accepted)"
    ),
    "duplicate_composite_key": "Duplicate Student ID+School+Goal+Metric row",
    "student_client_id_mismatch": "Student ID maps to multiple Client IDs",
    "case_manager_blank": "Case Manager is blank",
}

ACCEPTED_ISSUE_CODES = frozenset({"baseline_without_target_no_goal_context"})


@dataclass
class AuditConfig:
    sheet_name: str = DEFAULT_SHEET
    include_ok_rows: bool = False


def _text(record: RowRecord, field: str) -> str | None:
    value = record.get(field)
    return value if isinstance(value, str) else None


def is_blank(value: str | None) -> bool:
    return clean_value(value) is None


def baseline_target_status(record: RowRecord) -> str:
    baseline_blank = is_blank(_text(record, "Baseline"))
    target_blank = is_blank(_text(record, "Target"))
    if baseline_blank and target_blank:
        return "both_blank"
    if baseline_blank and not target_blank:
        return "target_only"
    if not baseline_blank and target_blank:
        return "baseline_only"
    return "both_present"


ProgressStatus = Literal["on_track", "off_track", "no_progress_data", "indeterminate"]
PROGRESS_STATUS_ORDER = ("on_track", "off_track", "no_progress_data", "indeterminate")


def latest_grading_period_value(record: RowRecord) -> str | None:
    for header in reversed(GRADING_PERIOD_HEADERS):
        text = _text(record, header)
        if not is_blank(text):
            return clean_value(text)
    return None


def progress_meets_target(
    metric: str | None,
    progress: str | None,
    target: str | None,
) -> bool | None:
    direction = metric_direction(metric)
    if direction == "none":
        return None
    progress_text = clean_value(progress)
    target_text = clean_value(target)
    if progress_text is None or target_text is None:
        return None

    progress_num = parse_number(progress_text)
    target_num = parse_number(target_text)
    if progress_num is not None and target_num is not None:
        if direction == "higher_is_better":
            return progress_num >= target_num
        if direction == "lower_is_better":
            return progress_num <= target_num
        return None

    progress_grade = normalize_grade(progress_text)
    target_grade = normalize_grade(target_text)
    if (
        progress_grade in GRADE_SCALE
        and target_grade in GRADE_SCALE
        and progress_grade is not None
        and target_grade is not None
    ):
        progress_rank = GRADE_RANK.get(progress_grade)
        target_rank = GRADE_RANK.get(target_grade)
        if progress_rank is None or target_rank is None:
            return None
        if direction == "higher_is_better":
            return progress_rank >= target_rank
        if direction == "lower_is_better":
            return progress_rank <= target_rank
    return None


def progress_status_for_record(record: RowRecord) -> ProgressStatus:
    if baseline_target_status(record) != "both_present":
        return "no_progress_data"
    progress = latest_grading_period_value(record)
    if progress is None:
        return "no_progress_data"
    metric = _text(record, "Metric")
    if not value_matches_scale(metric, progress):
        return "indeterminate"
    meets = progress_meets_target(metric, progress, _text(record, "Target"))
    if meets is None:
        return "indeterminate"
    return "on_track" if meets else "off_track"


def progress_row_key(record: RowRecord) -> str:
    return "|".join(
        clean_value(_text(record, field)) or ""
        for field in ("Student ID", "School", "Goal", "Metric")
    )


def _empty_progress_counts() -> dict[str, int]:
    return {status: 0 for status in PROGRESS_STATUS_ORDER}


def progress_rollup_from_records(records: list[RowRecord]) -> dict[str, Any]:
    global_counts: Counter[str] = Counter()
    schools: dict[str, Counter[str]] = defaultdict(Counter)
    progress_index: dict[str, str] = {}
    eligible_rows = 0

    for record in records:
        if baseline_target_status(record) != "both_present":
            continue
        eligible_rows += 1
        status = progress_status_for_record(record)
        global_counts[status] += 1
        school = clean_value(_text(record, "School")) or "(unknown)"
        schools[school][status] += 1
        progress_index[progress_row_key(record)] = status

    return {
        "global": {status: int(global_counts.get(status, 0)) for status in PROGRESS_STATUS_ORDER},
        "schools": {
            school: {status: int(counts.get(status, 0)) for status in PROGRESS_STATUS_ORDER}
            for school, counts in sorted(schools.items())
        },
        "eligible_rows": eligible_rows,
        "progress_index": progress_index,
    }


def progress_rollup(
    workbook_path: Path,
    sheet_name: str = DEFAULT_SHEET,
) -> dict[str, Any]:
    return progress_rollup_from_records(load_rows(workbook_path, sheet_name))


def grading_period_count(record: RowRecord) -> int:
    return sum(1 for header in GRADING_PERIOD_HEADERS if not is_blank(_text(record, header)))


def grading_period_fill_stats(
    workbook_path: Path,
    sheet_name: str = DEFAULT_SHEET,
) -> dict[str, int]:
    rows = load_rows(workbook_path, sheet_name)
    stats = {header: 0 for header in GRADING_PERIOD_HEADERS}
    for record in rows:
        for header in GRADING_PERIOD_HEADERS:
            if not is_blank(_text(record, header)):
                stats[header] += 1
    return stats


def student_metric_context(record: RowRecord) -> tuple[str, str]:
    return (
        clean_value(_text(record, "Student ID")) or "",
        clean_value(_text(record, "School Year")) or "",
    )


def student_contexts_with_complete_metric(records: list[RowRecord]) -> set[tuple[str, str]]:
    complete: set[tuple[str, str]] = set()
    for record in records:
        context = student_metric_context(record)
        if context[0] and not is_blank(_text(record, "Baseline")) and not is_blank(_text(record, "Target")):
            complete.add(context)
    return complete


class IssueSink:
    def __init__(
        self,
        *,
        issue_codes: Counter[str],
        detail_map: dict[int, list[str]],
    ) -> None:
        self._issue_codes = issue_codes
        self._detail_map = detail_map

    def add(self, row_num: int, code: str) -> None:
        self._issue_codes[code] += 1
        self._detail_map[row_num].append(code)

    def add_many(self, row_num: int, codes: list[str]) -> None:
        for code in codes:
            self.add(row_num, code)


class AcceptedSink:
    def __init__(
        self,
        *,
        accepted_codes: Counter[str],
        detail_map: dict[int, list[str]],
    ) -> None:
        self._accepted_codes = accepted_codes
        self._detail_map = detail_map

    def add(self, row_num: int, code: str) -> None:
        self._accepted_codes[code] += 1
        self._detail_map[row_num].append(code)

    def add_many(self, row_num: int, codes: list[str]) -> None:
        for code in codes:
            self.add(row_num, code)


def audit_composite_keys(records: list[RowRecord]) -> dict[str, Any]:
    seen: dict[tuple[str, str, str, str, str], int] = {}
    duplicates: list[int] = []
    for record in records:
        key = tuple(clean_value(_text(record, f)) or "" for f in COMPOSITE_KEY_FIELDS)
        if key in seen:
            duplicates.append(row_number(record))
        else:
            seen[key] = row_number(record)
    return {
        "unique_keys": len(seen),
        "duplicate_row_numbers": duplicates,
        "duplicate_count": len(duplicates),
    }


def audit_student_client_ids(records: list[RowRecord]) -> dict[str, Any]:
    sid_to_cids: dict[str, set[str]] = defaultdict(set)
    for record in records:
        sid = clean_value(_text(record, "Student ID"))
        cid = clean_value(_text(record, "Client ID"))
        if sid and cid:
            sid_to_cids[sid].add(cid)
    conflicting = {sid for sid, cids in sid_to_cids.items() if len(cids) > 1}
    return {
        "student_ids_with_multiple_client_ids": len(conflicting),
        "conflicting_student_ids": sorted(conflicting),
    }


def audit_baseline_target(records: list[RowRecord]) -> dict[str, int]:
    counts: Counter[str] = Counter()
    for record in records:
        status = baseline_target_status(record)
        if status == "both_blank":
            counts["both_baseline_and_target_blank"] += 1
        elif status == "target_only":
            counts["target_without_baseline"] += 1
        elif status == "baseline_only":
            counts["baseline_without_target"] += 1
        else:
            counts["baseline_and_target_present"] += 1
    return dict(counts)


def _audit_case_manager_blanks(records: list[RowRecord], sink: IssueSink) -> None:
    if not any("Case Manager" in record for record in records):
        return
    for record in records:
        if is_blank(_text(record, "Case Manager")):
            sink.add(row_number(record), "case_manager_blank")


def _mark_duplicate_composite_keys(key_info: dict[str, Any], sink: IssueSink) -> None:
    for row_num in key_info.get("duplicate_row_numbers") or []:
        sink.add(row_num, "duplicate_composite_key")


def _mark_student_client_id_mismatches(
    records: list[RowRecord],
    *,
    conflict_sid_set: set[str],
    sink: IssueSink,
) -> None:
    if not conflict_sid_set:
        return
    for record in records:
        sid = clean_value(_text(record, "Student ID"))
        if sid and sid in conflict_sid_set:
            sink.add(row_number(record), "student_client_id_mismatch")


def _audit_baseline_target_flags(
    records: list[RowRecord],
    *,
    complete_metric_contexts: set[tuple[str, str]],
    sink: IssueSink,
    accepted_sink: AcceptedSink,
) -> None:
    for record in records:
        status = baseline_target_status(record)
        flags: list[str] = []
        accepted_flags: list[str] = []
        if status == "both_blank":
            flags.append("both_baseline_and_target_blank")
            if grading_period_count(record) >= 2:
                flags.append("both_blank_with_two_progress_reports")
        elif status == "target_only":
            flags.append("target_without_baseline")
        elif status == "baseline_only":
            context = student_metric_context(record)
            if context[0] and context in complete_metric_contexts:
                flags.append("baseline_without_target")
                flags.append("baseline_without_target_non_goal_context")
            else:
                accepted_flags.append("baseline_without_target_no_goal_context")
        if flags:
            sink.add_many(row_number(record), flags)
        if accepted_flags:
            accepted_sink.add_many(row_number(record), accepted_flags)


def _audit_domain_scale_checks(records: list[RowRecord], sink: IssueSink) -> None:
    for record in records:
        analysis = analyze_record(record)
        codes = list(analysis.get("issue_codes") or [])
        if codes:
            sink.add_many(row_number(record), codes)


def _build_detail_rows(
    records: list[RowRecord],
    detail_map: dict[int, list[str]],
) -> list[dict[str, Any]]:
    row_index = {row_number(r): r for r in records}
    return [
        _detail_row(row_index[rn], sorted(set(codes)))
        for rn, codes in sorted(detail_map.items())
        if rn in row_index
    ]


def run_audit(records: list[RowRecord]) -> dict[str, Any]:
    summary: Counter[str] = Counter()
    issue_codes: Counter[str] = Counter()
    detail_map: dict[int, list[str]] = defaultdict(list)
    accepted_codes: Counter[str] = Counter()
    accepted_detail_map: dict[int, list[str]] = defaultdict(list)
    sink = IssueSink(issue_codes=issue_codes, detail_map=detail_map)
    accepted_sink = AcceptedSink(
        accepted_codes=accepted_codes,
        detail_map=accepted_detail_map,
    )

    key_info = audit_composite_keys(records)
    client_info = audit_student_client_ids(records)
    complete_metric_contexts = student_contexts_with_complete_metric(records)
    bt = audit_baseline_target(records)

    summary["rows"] = len(records)
    summary["duplicate_composite_keys"] = key_info["duplicate_count"]
    summary["student_client_id_conflicts"] = client_info["student_ids_with_multiple_client_ids"]
    for key, value in bt.items():
        summary[key] = value

    _audit_case_manager_blanks(records, sink)
    _mark_duplicate_composite_keys(key_info, sink)
    _mark_student_client_id_mismatches(
        records,
        conflict_sid_set=set(client_info.get("conflicting_student_ids") or []),
        sink=sink,
    )
    _audit_baseline_target_flags(
        records,
        complete_metric_contexts=complete_metric_contexts,
        sink=sink,
        accepted_sink=accepted_sink,
    )
    _audit_domain_scale_checks(records, sink)

    detail_rows = _build_detail_rows(records, detail_map)
    accepted_detail_rows = _build_detail_rows(records, accepted_detail_map)
    summary["accepted_exception_rows"] = len(accepted_detail_rows)

    return {
        "composite_key_audit": key_info,
        "student_client_audit": client_info,
        "baseline_target_distribution": bt,
        "summary": dict(summary),
        "issue_code_counts": dict(issue_codes),
        "accepted_exception_counts": dict(accepted_codes),
        "detail_rows": detail_rows,
        "accepted_detail_rows": accepted_detail_rows,
    }


def _detail_row(record: RowRecord, codes: list[str]) -> dict[str, Any]:
    return {
        "row_number": record.get("row_number"),
        "student_id": clean_value(_text(record, "Student ID")),
        "school": clean_value(_text(record, "School")),
        "goal": clean_value(_text(record, "Goal")),
        "metric": clean_value(_text(record, "Metric")),
        "baseline": clean_value(_text(record, "Baseline")),
        "target": clean_value(_text(record, "Target")),
        "issue_codes": ";".join(codes),
    }


def write_audit_flags_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    fieldnames = [
        "row_number",
        "student_id",
        "school",
        "goal",
        "metric",
        "baseline",
        "target",
        "issue_codes",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k) for k in fieldnames})


def _format_issue_counts(counts: dict[str, int]) -> list[str]:
    lines = []
    for code in sorted(counts):
        label = FLAG_DESCRIPTIONS.get(code, code)
        lines.append(f"- **{code}** ({label}): {counts[code]}")
    return lines


def write_summary_markdown(path: Path, payload: dict[str, Any]) -> None:
    counts = payload.get("issue_code_counts") or {}
    accepted_counts = payload.get("accepted_exception_counts") or {}
    summary = payload.get("summary") or {}
    accepted_row_count = len(payload.get("accepted_detail_rows") or [])
    lines = [
        "# Student Metrics Audit Summary",
        "",
        f"- **Workbook:** `{payload.get('workbook', '')}`",
        f"- **Sheet:** {payload.get('sheet', DEFAULT_SHEET)}",
        f"- **Run at:** {payload.get('run_at', '')}",
        "",
        "## Counts",
        "",
        f"- Rows evaluated: **{summary.get('rows', 0)}**",
        f"- Rows flagged: **{len(payload.get('detail_rows') or [])}**",
        f"- Accepted exceptions: **{accepted_row_count}**",
        f"- Duplicate composite keys: **{summary.get('duplicate_composite_keys', 0)}**",
        f"- Student/client ID conflicts: **{summary.get('student_client_id_conflicts', 0)}**",
        "",
    ]
    if accepted_counts:
        lines.extend(["## Accepted exceptions", ""])
        lines.extend(_format_issue_counts(accepted_counts))
        lines.append("")
    lines.extend(["## Issue codes", ""])
    if counts:
        lines.extend(_format_issue_counts(counts))
    else:
        lines.append("No issues found.")
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def results_to_json_payload(results: dict[str, Any], *, workbook: str, sheet: str) -> dict[str, Any]:
    return {
        "workbook": workbook,
        "sheet": sheet,
        "run_at": datetime.now(timezone.utc).isoformat(),
        "summary": results.get("summary", {}),
        "issue_code_counts": results.get("issue_code_counts", {}),
        "accepted_exception_counts": results.get("accepted_exception_counts", {}),
        "accepted_detail_rows": results.get("accepted_detail_rows") or [],
        "baseline_target_distribution": results.get("baseline_target_distribution", {}),
        "composite_key_audit": results.get("composite_key_audit", {}),
        "student_client_audit": results.get("student_client_audit", {}),
        "detail_row_count": len(results.get("detail_rows") or []),
        "accepted_detail_row_count": len(results.get("accepted_detail_rows") or []),
    }


def export_results(results: dict[str, Any], output_dir: Path, *, workbook: str, sheet: str) -> dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = results_to_json_payload(results, workbook=workbook, sheet=sheet)
    paths = {
        "audit_flags": output_dir / "audit_flags.csv",
        "audit_summary_json": output_dir / "audit_summary.json",
        "audit_summary_md": output_dir / "audit_summary.md",
    }
    write_audit_flags_csv(paths["audit_flags"], results.get("detail_rows") or [])
    paths["audit_summary_json"].write_text(
        json.dumps(payload, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    write_summary_markdown(
        paths["audit_summary_md"],
        {
            **payload,
            "detail_rows": results.get("detail_rows"),
            "accepted_detail_rows": results.get("accepted_detail_rows"),
        },
    )
    return {key: str(path) for key, path in paths.items()}


def evaluate_workbook(workbook_path: Path, config: AuditConfig) -> dict[str, Any]:
    records = load_rows(workbook_path, sheet_name=config.sheet_name)
    results = run_audit(records)
    results["workbook"] = str(workbook_path.resolve())
    results["sheet"] = config.sheet_name
    results["config"] = asdict(config)
    return results


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Flag invalid metric/scale combinations in student metrics summary exports.",
    )
    parser.add_argument(
        "--workbook",
        action="append",
        type=Path,
        help="Path to student metrics .xlsx (repeatable). When omitted, resolves via cisiphyus.",
    )
    parser.add_argument(
        "--force-fetch",
        action="store_true",
        help="Always re-export from CISDM even when the local workbook is fresh.",
    )
    parser.add_argument(
        "--fetch-destination",
        type=Path,
        help="Where to save the CISDM export (default: artifacts/audit/).",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET,
        help=f"Worksheet name (default {DEFAULT_SHEET}).",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit 1 if duplicate composite keys or Student ID vs Client ID conflicts.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print JSON summary to stdout.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = AuditConfig(sheet_name=args.sheet)

    workbooks: list[Path] = []
    if args.workbook:
        workbooks.extend(args.workbook)
    else:
        destination = args.fetch_destination or preferred_student_metrics_workbook()
        try:
            fetch_result = fetch_student_metrics_workbook(
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
        export_results(results, out_dir, workbook=str(workbook), sheet=config.sheet_name)
        summaries.append(results)
        print(f"Workbook: {workbook}")
        print(f"Output dir: {out_dir}")
        print(f"  rows: {results['summary'].get('rows', 0)}")
        print(f"  flagged_rows: {len(results.get('detail_rows') or [])}")

    if args.strict:
        for results in summaries:
            if results["composite_key_audit"]["duplicate_count"] > 0:
                return 1
            if results["student_client_audit"]["student_ids_with_multiple_client_ids"] > 0:
                return 1

    if args.json:
        payloads = [
            results_to_json_payload(r, workbook=r["workbook"], sheet=r["sheet"]) for r in summaries
        ]
        print(json.dumps(payloads if len(payloads) > 1 else payloads[0], indent=2, sort_keys=True))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())