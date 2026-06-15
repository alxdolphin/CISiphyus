#!/usr/bin/env python3

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import sys
from functools import lru_cache
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils.cell import get_column_letter

QPR_DIR = Path(__file__).resolve().parent
FIXTURES_DIR = QPR_DIR / "fixtures"
LOCAL_INPUTS_DIR = QPR_DIR / "local_inputs"
REPO_ROOT = QPR_DIR.parents[1]
PROJECT_ROOT = REPO_ROOT
DEFAULT_TEMPLATE_PATH = FIXTURES_DIR / "qpr_import_template.xlsx"
DEFAULT_DEADLINES_PATH = FIXTURES_DIR / "reporting_deadlines_minimal.xlsx"
DEFAULT_QPR_LOCAL_INPUTS_DIR = LOCAL_INPUTS_DIR
_EXAMPLES_DIR = QPR_DIR.parent


def default_student_metrics_filename() -> str:
    env_name = (os.environ.get("QPR_STUDENT_METRICS_FILENAME") or "").strip()
    if env_name:
        return env_name
    if str(_EXAMPLES_DIR) not in sys.path:
        sys.path.insert(0, str(_EXAMPLES_DIR))
    import _cisiphyus_fetch as cis_fetch

    year = cis_fetch.resolve_school_year(cisiphyus_root=REPO_ROOT)
    return cis_fetch.student_metrics_filename(year)


DEFAULT_STUDENT_METRICS_FILENAME = default_student_metrics_filename()
DEFAULT_METRIC_WORKBOOK_PATH = DEFAULT_QPR_LOCAL_INPUTS_DIR / DEFAULT_STUDENT_METRICS_FILENAME
DEFAULT_QPR_OUTPUT_DIR = REPO_ROOT / "artifacts" / "qpr"
DEFAULT_SITE_STAFF_LIST_PATH = FIXTURES_DIR / "site_staff_list_minimal.xlsx"
SITE_LIST_SHEET = "Site List"
SITE_LIST_SITE_COORDINATOR_HEADER = "Site Coordinator"
SITE_LIST_END_DATE_HEADER = "End Date"
METRIC_SUMMARY_GRADING_PERIOD_HEADERS = (
    "1st Grading Period",
    "2nd Grading Period",
    "3rd Grading Period",
    "4th Grading Period",
    "5th Grading Period",
    "6th Grading period",
)
METRIC_SUMMARY_QUARTER_STATUS_HEADERS: dict[int, tuple[str, ...]] = {
    1: (
        "1st Grading Period Status",
        "Q1 Status",
        "Q1 Progress Status",
        "1st Grading Period Progress Status",
    ),
    2: (
        "2nd Grading Period Status",
        "Q2 Status",
        "Q2 Progress Status",
        "2nd Grading Period Progress Status",
    ),
    3: (
        "3rd Grading Period Status",
        "Q3 Status",
        "Q3 Progress Status",
        "3rd Grading Period Progress Status",
    ),
    4: (
        "4th Grading Period Status",
        "Q4 Status",
        "Q4 Progress Status",
        "4th Grading Period Progress Status",
    ),
}
MAX_METRIC_COMMENT_LENGTH = 1000
EXCEL_METRIC_CONTEXT_NOTE_HINT = (
    "Metric context is stored as Excel Notes on Progress Against Goal cells "
    "(Review tab > Notes, or the small corner triangle), not in the modern Comments pane."
)
PROGRESS_AGAINST_GOAL_HEADERS = ("Progress Against Goal", "Progress against goal")
NO_GOAL_ASSIGNED_TEXT = "No Goal Assigned for this Metric"
HIDDEN_PROVISION_HEADERS = (
    "School",
    "Grading Period / Assessment Date",
    "Grading Period",
    "School Year",
    "Organization",
)
CORE_COURSE_GRADES_SHEET = "Core Course Grades"
CORE_COURSE_HEADER = "Core Course"
STANDARDIZED_TEST_SCORE_SHEET = "Standardized Test Score"
STAND_TEST_TYPE_HEADER = "Stand. Test Type"
SEL_SHEET = "SEL"
SEL_METRIC_HEADER = "Social Emotional Learning Metric"
CREDITS_CTE_SHEET = "Credits - CTE Needed"
CREDITS_METRIC_HEADER = "Metric "
FIELD_OPTIONS_SEL_METRIC_KEY = "Social Emotional Learning Metric "
FIELD_OPTIONS_CREDITS_KEY = "Credits"
DEADLINES_SHEET = "Sheet4"
FIELD_OPTIONS_SHEET = "Field Options"
ORGANIZATION_CANONICAL = "CIS of Eastern Pennsylvania"
UAT_NO_ENTITY_ID_ROWS = "UAT_NO_ENTITY_ID_ROWS"
NO_STUDENTS_AFTER_ENROLLMENT_FILTER = "NO_STUDENTS_AFTER_ENROLLMENT_FILTER"
INACTIVE_SITE_COORDINATOR_PROVISION = "INACTIVE_SITE_COORDINATOR_PROVISION"
METRIC_TEMPLATE_SHEETS = frozenset(
    {
        "Attendance Rate_%",
        "Attendance_Days Absent",
        "Tardies",
        "Suspensions",
        "Discliplinary Referrals",
        "Other Behavior Incidents",
        "Conduct",
        "Core Course Grades",
        "GPA",
        "Credits - CTE Needed",
        "Reading Level",
        "Standardized Test Score",
        "SEL",
    }
)


def clean_text(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None




def normalize_text(value: object) -> str:
    text = clean_text(value)
    if text is None:
        return ""
    return " ".join(text.lower().split())




def school_match_keys(label: object) -> frozenset[str]:
    # WHY: site list and caseload exports often differ by trailing school / charter wording only
    n = normalize_text(label)
    if not n:
        return frozenset()
    keys: set[str] = {n}
    strip_suffixes = (
        " charter school",
        " elementary school",
        " middle school",
        " high school",
        " intermediate school",
        " junior high school",
    )
    for suffix in strip_suffixes:
        if n.endswith(suffix):
            keys.add(n[: -len(suffix)].strip())
    if n.endswith(" elementary") and not n.endswith(" elementary school"):
        keys.add(f"{n} school")
    if n.endswith(" middle") and not n.endswith(" middle school"):
        keys.add(f"{n} school")
    if n.endswith(" high") and not n.endswith(" high school"):
        keys.add(f"{n} school")
    if " senior high school" in n:
        keys.add(n.replace(" senior high school", " high school"))
    return frozenset(key for key in keys if key)




def site_group_match_keys(label: object) -> frozenset[str]:
    # why: keep site batch outputs one-school-per-file while still handling minor naming variants
    n = normalize_text(label)
    if not n:
        return frozenset()
    keys: set[str] = {n}
    if n.endswith(" charter school"):
        keys.add(n[: -len(" school")].strip())
    if n.endswith((" elementary", " middle", " high", " intermediate", " junior high")):
        keys.add(f"{n} school")
    if n.endswith(" school"):
        base = n[: -len(" school")].strip()
        if base.endswith((" elementary", " middle", " high", " intermediate", " junior high")):
            keys.add(base)
    if " senior high school" in n:
        keys.add(n.replace(" senior high school", " high school"))
    if " senior high" in n:
        keys.add(n.replace(" senior high", " high"))
    return frozenset(key for key in keys if key)




def caseload_school_matches_resolved(caseload_school: str | None, resolved_schools: set[str]) -> bool:
    if not caseload_school or not resolved_schools:
        return False
    if caseload_school in resolved_schools:
        return True
    row_keys = school_match_keys(caseload_school)
    for resolved in resolved_schools:
        if row_keys & school_match_keys(resolved):
            return True
    return False




def student_id_join_key(value: object) -> str | None:
    # why: excel may yield int/float ids; caseload export often uses plain digit strings
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value != value:
            return None
        if value == int(value):
            return str(int(value))
        return clean_text(value)
    text = clean_text(value)
    if text is None:
        return None
    try:
        numeric = float(text)
        if numeric == int(numeric):
            return str(int(numeric))
    except ValueError:
        pass
    return text




def school_labels_equivalent(a: object, b: object) -> bool:
    ka, kb = school_match_keys(a), school_match_keys(b)
    return bool(ka and kb and (ka & kb))




def partition_schools_by_label_equivalence(schools: list[str]) -> list[list[str]]:
    # why: caseload export often varies charter / school suffix spelling for the same site
    if not schools:
        return []
    parent = list(range(len(schools)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        pi, pj = find(i), find(j)
        if pi != pj:
            parent[pi] = pj

    for i in range(len(schools)):
        for j in range(i + 1, len(schools)):
            if school_labels_equivalent(schools[i], schools[j]):
                union(i, j)
    buckets: dict[int, list[str]] = {}
    for i, s in enumerate(schools):
        r = find(i)
        buckets.setdefault(r, []).append(s)
    return list(buckets.values())




def person_name_equivalent(a: object, b: object) -> bool:
    # WHY: caseload often uses last, first; site list uses first last — same person must still match
    ca, cb = clean_text(a), clean_text(b)
    if ca is None or cb is None:
        return False
    if ca == cb:
        return True

    def equivalence_keys(label: str) -> set[str]:
        keys: set[str] = {normalize_text(label)}
        if "," in label:
            left, _, right = label.partition(",")
            last = normalize_text(left)
            first = normalize_text(right)
            if last and first:
                keys.add(f"{first} {last}")
                keys.add(f"{last} {first}")
        return keys

    return bool(equivalence_keys(ca) & equivalence_keys(cb))




def to_iso_date(value: object) -> str | None:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = clean_text(value)
    if text is None:
        return None
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return None




def header_values(worksheet) -> list[str]:
    headers = [cell.value if cell.value is not None else "" for cell in worksheet[1]]
    while headers and headers[-1] == "":
        headers.pop()
    return [str(value) for value in headers]




def nonempty_option_values(worksheet) -> dict[str, list[str]]:
    headers = header_values(worksheet)
    values: dict[str, list[str]] = {}
    for column_index, header in enumerate(headers, start=1):
        options: list[str] = []
        for row_index in range(2, worksheet.max_row + 1):
            raw = worksheet.cell(row=row_index, column=column_index).value
            text = clean_text(raw)
            if text is not None:
                options.append(text)
        values[header] = options
    return values




def formula_coordinates(worksheet, headers: list[str]) -> dict[str, list[str]]:
    formulas_by_header: dict[str, list[str]] = {header: [] for header in headers}
    for column_index, header in enumerate(headers, start=1):
        for row_index in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_index, column=column_index)
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                formulas_by_header[header].append(cell.coordinate)
    return {header: coords for header, coords in formulas_by_header.items() if coords}




def default_values(worksheet, headers: list[str]) -> dict[str, str]:
    defaults: dict[str, str] = {}
    for column_index, header in enumerate(headers, start=1):
        for row_index in range(2, worksheet.max_row + 1):
            text = clean_text(worksheet.cell(row=row_index, column=column_index).value)
            if text is not None:
                defaults[header] = text
                break
    return defaults




def build_template_schema(template_path: Path) -> dict[str, Any]:
    workbook = load_workbook(template_path, data_only=False)
    sheets: dict[str, dict[str, Any]] = {}
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        headers = header_values(worksheet)
        formulas = formula_coordinates(worksheet, headers)
        sheets[sheet_name] = {
            "headers": headers,
            "formula_headers": sorted(formulas),
            "formula_coordinates": formulas,
            "defaults": default_values(worksheet, headers),
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        }
    field_options = nonempty_option_values(workbook[FIELD_OPTIONS_SHEET])
    workbook.close()
    return {
        "sheet_order": list(sheets),
        "sheets": sheets,
        "field_options": field_options,
    }




def _load_audit_module():
    candidates = (
        Path(__file__).resolve().parent / "audit" / "audit_metrics_all.py",
        PROJECT_ROOT / "tools" / "accreditation" / "scripts" / "audit_metrics_all.py",
    )
    for path in candidates:
        if not path.exists():
            continue
        spec = importlib.util.spec_from_file_location("audit_metrics_all", path)
        if spec is None or spec.loader is None:
            continue
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    return None


@lru_cache(maxsize=1)


def _get_audit_module():
    return _load_audit_module()


SCHEDULE_KEY_ALIASES: dict[str, str] = {
    "reading senior high school": "reading",
    "lincoln leadership academy charter school": "lla",
    "octorara high school": "octorara",
    "octorara middle school": "octorara",
    "fleetwood area high school": "fleetwood",
}




def schedule_key_for_school(school: str | None) -> str | None:
    if school is None:
        return None
    normalized = normalize_text(school)
    if normalized in SCHEDULE_KEY_ALIASES:
        return SCHEDULE_KEY_ALIASES[normalized]
    audit_module = _get_audit_module()
    schedule_key = None
    if audit_module is not None and hasattr(audit_module, "_schedule_key_for_school"):
        schedule_key = audit_module._schedule_key_for_school(school)
    if schedule_key is not None:
        return schedule_key
    return clean_text(school)




def grading_period_key(value: object) -> str | None:
    text = clean_text(value)
    if text is None:
        return None
    cleaned = text.replace("Q", "").replace("MP", "").strip()
    try:
        number = int(float(cleaned))
    except ValueError:
        return None
    if number not in (1, 2, 3, 4):
        return None
    return f"{number}.0"




def canonical_grading_period(value: object) -> str:
    period = grading_period_key(value)
    if period is None:
        raise ValueError("Grading period must map to 1.0, 2.0, 3.0, or 4.0 (examples: 2, 2.0, Q2, MP2).")
    return period




def grading_period_cell_value(period: str) -> int:
    return int(float(period))




def assessment_date_cell_value(iso_value: str | None) -> date | None:
    if iso_value is None:
        return None
    return datetime.strptime(iso_value, "%Y-%m-%d").date()


_PERIOD_CODE_BY_CANONICAL: dict[str, str] = {
    "1.0": "MP1",
    "2.0": "MP2",
    "3.0": "MP3",
    "4.0": "EOY",
}




def load_reporting_deadlines(deadlines_path: Path) -> dict[str, Any]:
    workbook = load_workbook(deadlines_path, data_only=True)
    worksheet = workbook[DEADLINES_SHEET]
    schedules: dict[str, dict[str, Any]] = {}
    reference_issues: dict[str, list[str]] = {}
    for row_index in range(2, worksheet.max_row + 1):
        school = clean_text(worksheet.cell(row=row_index, column=1).value)
        if school is None:
            continue
        dates = {
            "1.0": to_iso_date(worksheet.cell(row=row_index, column=2).value),
            "2.0": to_iso_date(worksheet.cell(row=row_index, column=3).value),
            "3.0": to_iso_date(worksheet.cell(row=row_index, column=4).value),
            "4.0": to_iso_date(worksheet.cell(row=row_index, column=5).value),
        }
        issues: list[str] = []
        parsed_dates: dict[str, date] = {}
        previous: date | None = None
        for period, iso_value in dates.items():
            if iso_value is None:
                issues.append(f"period_{int(float(period))}_missing")
                continue
            current = datetime.strptime(iso_value, "%Y-%m-%d").date()
            parsed_dates[period] = current
            if current.year < 2025 or current.year > 2027:
                issues.append(f"period_{int(float(period))}_out_of_range")
            if previous is not None and current <= previous:
                issues.append(f"period_{int(float(period))}_not_after_previous")
            previous = current
        schedules[school] = {
            "assessment_dates": dates,
        }
        if issues:
            reference_issues[school] = sorted(set(issues))
    if "Reporting Dates" in workbook.sheetnames:
        reporting = workbook["Reporting Dates"]
        for row_index in range(3, reporting.max_row + 1):
            school = clean_text(reporting.cell(row=row_index, column=1).value)
            if school is None:
                continue
            due_dates = {
                "1.0": to_iso_date(reporting.cell(row=row_index, column=3).value),
                "2.0": to_iso_date(reporting.cell(row=row_index, column=5).value),
                "4.0": to_iso_date(reporting.cell(row=row_index, column=7).value),
            }
            schedules.setdefault(school, {"assessment_dates": {}})["due_dates"] = due_dates
    district_lookup: dict[str, str] = {}
    for school in schedules:
        key = normalize_text(school)
        if key:
            district_lookup.setdefault(key, school)
    for alias, target in {
        "stony creek elementary school": "Antietam",
        "stony creek elementary": "Antietam",
        "tyson-schoener elementary school": "Reading",
    }.items():
        district_lookup.setdefault(alias, target)
    workbook.close()
    return {
        "schedules": schedules,
        "reference_issues": reference_issues,
        "school_district_lookup": district_lookup,
    }




def _header_lookup(headers: list[str]) -> dict[str, int]:
    return {header: index for index, header in enumerate(headers)}




def _open_sheet_with_headers(
    workbook_path: Path | None,
    *,
    sheet_name: str | None = None,
) -> tuple[Any, Any, dict[str, int]] | None:
    if workbook_path is None or not workbook_path.exists():
        return None
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    if sheet_name is None:
        worksheet = workbook[workbook.sheetnames[0]]
    else:
        if sheet_name not in workbook.sheetnames:
            workbook.close()
            return None
        worksheet = workbook[sheet_name]
    headers = header_values(worksheet)
    return workbook, worksheet, _header_lookup(headers)




def _append_unique_name(
    names: list[str],
    seen: set[str],
    candidate: object,
) -> None:
    staff_name = clean_text(candidate)
    if staff_name is None or staff_name == "-":
        return
    normalized = normalize_text(staff_name)
    if normalized in seen:
        return
    seen.add(normalized)
    names.append(staff_name)




def _iter_template_sheets(template_schema: dict[str, Any], *, workbook=None):
    for sheet_name, sheet_schema in template_schema["sheets"].items():
        if sheet_name == FIELD_OPTIONS_SHEET:
            continue
        if workbook is not None and sheet_name not in workbook.sheetnames:
            continue
        yield sheet_name, sheet_schema




def split_student_display_name(raw: object) -> tuple[str, str]:
    # why: metric summary uses one display name; template needs first/last columns
    text = clean_text(raw)
    if not text:
        return "", ""
    if "," in text:
        left, _, right = text.partition(",")
        last = left.strip()
        first = right.strip()
        return first, last
    parts = text.split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return "", parts[0]
    return " ".join(parts[:-1]), parts[-1]




def load_metric_roster_lookup(metric_workbook_path: Path | None) -> dict[str, dict[str, str]]:
    # why: provision keys on Case Manager in metric work; caseload CreatedBy is broader than site QPR staff
    opened = _open_sheet_with_headers(metric_workbook_path)
    if opened is None:
        return {}
    workbook, worksheet, header_lookup = opened
    if "Student ID" not in header_lookup:
        workbook.close()
        return {}

    def value_for(row: tuple[object, ...], *header_names: str) -> str:
        for header_name in header_names:
            index = header_lookup.get(header_name)
            if index is None or index >= len(row):
                continue
            text = clean_text(row[index])
            if text is not None:
                return text
        return ""

    def raw_cell(row: tuple[object, ...], header_name: str) -> object | None:
        index = header_lookup.get(header_name)
        if index is None or index >= len(row):
            return None
        return row[index]

    lookup: dict[str, dict[str, str]] = {}
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        student_id = value_for(row, "Student ID")
        if not student_id:
            continue
        case_manager = value_for(row, "Case Manager")
        school = value_for(row, "School")
        if not case_manager or not school:
            continue
        first_name, last_name = split_student_display_name(value_for(row, "Student Name"))
        entity_id = value_for(row, "EntityID") or value_for(row, "Client ID")
        enrollment_end_iso = to_iso_date(raw_cell(row, "Enrollment EndDate")) or ""
        entry = {
            "student_id": student_id,
            "entity_id": entity_id,
            "first_name": first_name,
            "last_name": last_name,
            "organization": value_for(row, "Organization", "Organization Name"),
            "school": school,
            "created_by": case_manager,
            "school_year": value_for(row, "School Year"),
            "enrollment_end_iso": enrollment_end_iso,
        }
        if student_id not in lookup:
            lookup[student_id] = entry
        else:
            prior = lookup[student_id]
            if not prior.get("first_name") and first_name:
                prior["first_name"] = first_name
            if not prior.get("last_name") and last_name:
                prior["last_name"] = last_name
            if not prior.get("school_year") and entry.get("school_year"):
                prior["school_year"] = entry["school_year"]
            if not prior.get("entity_id") and entity_id:
                prior["entity_id"] = entity_id
            pe = prior.get("enrollment_end_iso") or ""
            ne = entry.get("enrollment_end_iso") or ""
            if ne and (not pe or ne < pe):
                prior["enrollment_end_iso"] = ne
    workbook.close()
    return lookup




def effective_student_rows_for_uat(
    student_rows: list[dict[str, str]], uat: bool
) -> tuple[list[dict[str, str]], list[str]]:
    # why: uat templates use EntityID or Client ID as displayed student id; rows without it cannot be uploaded safely
    if not uat:
        return student_rows, []
    skipped: list[str] = []
    effective: list[dict[str, str]] = []
    for row in student_rows:
        if clean_text(row.get("entity_id")):
            effective.append(row)
        else:
            sid = clean_text(row.get("student_id"))
            if sid:
                skipped.append(sid)
    return effective, skipped




def _core_course_field_options(template_schema: dict[str, Any]) -> list[str]:
    opts = template_schema.get("field_options", {}).get(CORE_COURSE_HEADER)
    return list(opts) if opts else []




def _stand_test_type_field_options(template_schema: dict[str, Any]) -> list[str]:
    opts = template_schema.get("field_options", {}).get(STAND_TEST_TYPE_HEADER)
    return list(opts) if opts else []




def _colon_metric_suffix_match_key(label: str) -> str:
    # why: CISDM spacing/slashes may differ from Field Options (core course + stand. test type)
    text = clean_text(label) or ""
    text = " ".join(text.split())
    text = re.sub(r"\s*/\s*", "/", text)
    return text.lower()




def template_option_from_colon_metric(
    metric_raw: object,
    required_prefix_lower: str,
    allowed_options: list[str],
) -> str | None:
    text = clean_text(metric_raw)
    if not text:
        return None
    lower = text.lower()
    if not lower.startswith(required_prefix_lower):
        return None
    if ":" not in text:
        return None
    suffix = text.split(":", 1)[1].strip()
    sk = _colon_metric_suffix_match_key(suffix)
    for opt in allowed_options:
        if _colon_metric_suffix_match_key(opt) == sk:
            return opt
    return None




def template_core_course_from_metric(metric_raw: object, allowed_options: list[str]) -> str | None:
    return template_option_from_colon_metric(metric_raw, "core course grades", allowed_options)




def template_stand_test_type_from_metric(metric_raw: object, allowed_options: list[str]) -> str | None:
    return template_option_from_colon_metric(metric_raw, "standardized test score", allowed_options)




def template_option_match_exact_or_suffix(metric_raw: object, allowed_options: list[str]) -> str | None:
    # why: CISDM metric names often match template Field Options exactly; slash spacing may differ
    text = clean_text(metric_raw)
    if not text:
        return None
    for opt in allowed_options:
        if text == opt:
            return opt
    sk_m = _colon_metric_suffix_match_key(text)
    for opt in allowed_options:
        if _colon_metric_suffix_match_key(opt) == sk_m:
            return opt
    return None




def template_sel_metric_from_metric(metric_raw: object, allowed_options: list[str]) -> str | None:
    matched = template_option_match_exact_or_suffix(metric_raw, allowed_options)
    if matched:
        return matched
    text = clean_text(metric_raw)
    if not text:
        return None
    lower = text.lower()
    if "overall sead" in lower and ("assessment" in lower or "score" in lower):
        for opt in allowed_options:
            if opt == "Overall SEAD":
                return opt
    return None




def template_credits_metric_from_metric(metric_raw: object, allowed_options: list[str]) -> str | None:
    return template_option_match_exact_or_suffix(metric_raw, allowed_options)




def _sel_metric_field_options(template_schema: dict[str, Any]) -> list[str]:
    fo = template_schema.get("field_options", {})
    opts = fo.get(FIELD_OPTIONS_SEL_METRIC_KEY) or fo.get(SEL_METRIC_HEADER)
    return list(opts) if opts else []




def _credits_field_options(template_schema: dict[str, Any]) -> list[str]:
    opts = template_schema.get("field_options", {}).get(FIELD_OPTIONS_CREDITS_KEY)
    return list(opts) if opts else []




def _pick_template_field_from_summary_blocks(
    blocks: list[dict[str, Any]] | None,
    allowed_options: list[str],
    extract: Callable[[object, list[str]], str | None],
) -> str | None:
    # why: one metric row per subtype; template row is per student — template Field Options order wins ties
    if not blocks or not allowed_options:
        return None
    matched: list[str] = []
    for block in blocks:
        m = block.get("Metric")
        choice = extract(m, allowed_options)
        if choice:
            matched.append(choice)
    if not matched:
        return None
    unique = list(dict.fromkeys(matched))
    if len(unique) == 1:
        return unique[0]
    chosen = set(unique)
    for opt in allowed_options:
        if opt in chosen:
            return opt
    return unique[0]




def _ordered_unique_metric_options_from_summary_blocks(
    blocks: list[dict[str, Any]] | None,
    allowed_options: list[str],
    extract: Callable[[object, list[str]], str | None],
) -> list[str]:
    if not blocks or not allowed_options:
        return []
    matched: list[str] = []
    for block in blocks:
        choice = extract(block.get("Metric"), allowed_options)
        if choice:
            matched.append(choice)
    if not matched:
        return []
    unique = list(dict.fromkeys(matched))
    chosen = set(unique)
    ordered = [opt for opt in allowed_options if opt in chosen]
    return ordered if ordered else unique




def _expand_metric_option_rows_for_sheet(
    *,
    rows_for_sheet: list[dict[str, str]],
    sheet_name: str,
    summary_index: dict[tuple[str, str], list[dict[str, Any]]] | None,
    allowed_options: list[str],
    extract: Callable[[object, list[str]], str | None],
) -> list[dict[str, Any]]:
    if not summary_index or not allowed_options:
        return [dict(row) for row in rows_for_sheet]
    expanded: list[dict[str, Any]] = []
    for student in rows_for_sheet:
        base = dict(student)
        join_key = student_id_join_key(student.get("student_id"))
        if join_key is None:
            expanded.append(base)
            continue
        blocks = summary_index.get((join_key, sheet_name))
        options = _ordered_unique_metric_options_from_summary_blocks(blocks, allowed_options, extract)
        if not options:
            expanded.append(base)
            continue
        for option in options:
            option_blocks = [
                block
                for block in (blocks or [])
                if extract(block.get("Metric"), allowed_options) == option
            ]
            row = dict(base)
            row["__metric_option"] = option
            row["__metric_has_goal_target"] = any(
                clean_text(block.get("Target")) is not None for block in option_blocks
            )
            row["__metric_blocks"] = option_blocks
            expanded.append(row)
    return expanded




def template_sheet_for_metric(metric: object) -> str | None:
    # WHY: CISDM metric labels must map to the fixed QPR tab names without changing the template grid
    text = clean_text(metric)
    if text is None:
        return None
    lower = text.lower()
    if "attendance rate (%)" in lower:
        return "Attendance Rate_%"
    if "days absent" in lower and "attendance" in lower:
        return "Attendance_Days Absent"
    if lower == "tardies":
        return "Tardies"
    if "suspension" in lower:
        return "Suspensions"
    if "disciplinary referral" in lower:
        return "Discliplinary Referrals"
    if "other behavior incident" in lower:
        return "Other Behavior Incidents"
    if lower == "conduct":
        return "Conduct"
    if lower.startswith("core course grades"):
        return "Core Course Grades"
    if lower == "gpa":
        return "GPA"
    if "credit" in lower and ("needed" in lower or "completion" in lower):
        return "Credits - CTE Needed"
    if lower == "reading level":
        return "Reading Level"
    if lower.startswith("standardized test score"):
        return "Standardized Test Score"
    if _metric_maps_to_sel(lower):
        return "SEL"
    return None




def _metric_maps_to_sel(lower: str) -> bool:
    if "other (sel)" in lower:
        return True
    if "sead" in lower:
        return True
    hints = (
        "social engagement",
        "behavior engagement",
        "emotional engagement",
        "cognitive engagement",
        "global engagement",
        "social support",
        "social awareness",
        "self-perception",
        "self control",
    )
    return any(term in lower for term in hints)




def _format_value_for_comment(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return clean_text(value)




def _metric_summary_row_from_cells(row: tuple[Any, ...], header_lookup: dict[str, int]) -> dict[str, Any]:
    entry: dict[str, Any] = {}
    for header in (
        "School",
        "Goal",
        "Metric",
        "Baseline",
        "Baseline Time Period",
        "Target",
        "Baseline Date",
        "Target Date",
        "Latest Progress",
    ):
        if header not in header_lookup:
            continue
        idx = header_lookup[header]
        if idx >= len(row):
            continue
        raw = row[idx]
        if header in ("Baseline", "Target"):
            entry[header] = raw
        else:
            entry[header] = _format_value_for_comment(raw)
    for gp_header in METRIC_SUMMARY_GRADING_PERIOD_HEADERS:
        if gp_header not in header_lookup:
            continue
        idx = header_lookup[gp_header]
        if idx >= len(row):
            continue
        entry[gp_header] = _format_value_for_comment(row[idx])
    for status_headers in METRIC_SUMMARY_QUARTER_STATUS_HEADERS.values():
        for status_header in status_headers:
            if status_header not in header_lookup:
                continue
            idx = header_lookup[status_header]
            if idx >= len(row):
                continue
            entry[status_header] = _format_value_for_comment(row[idx])
    return entry




def _student_metric_has_goal_target(row: tuple[Any, ...], header_lookup: dict[str, int]) -> bool:
    target_idx = header_lookup.get("Target")
    if target_idx is not None and target_idx < len(row):
        if clean_text(row[target_idx]) is not None:
            return True
    return False




def load_metric_tracking_workbook(
    metric_workbook_path: Path,
) -> tuple[dict[tuple[str, str], dict[str, bool]], dict[tuple[str, str], list[dict[str, Any]]]]:
    # why: one pass over student metric tracking summary for domain filter, target prefill, and comments
    opened = _open_sheet_with_headers(metric_workbook_path)
    if opened is None:
        return {}, {}
    workbook, worksheet, header_lookup = opened
    required_headers = {"Student ID", "Metric"}
    missing_headers = sorted(required_headers - set(header_lookup))
    if missing_headers:
        workbook.close()
        joined = ", ".join(missing_headers)
        raise ValueError(f"Metric tracking workbook is missing required headers: {joined}.")
    domain_index: dict[tuple[str, str], dict[str, bool]] = {}
    summary_index: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        sid_idx = header_lookup["Student ID"]
        raw_sid = row[sid_idx] if sid_idx < len(row) else None
        join_key = student_id_join_key(raw_sid)
        if join_key is None:
            continue
        metric_idx = header_lookup["Metric"]
        metric = row[metric_idx] if metric_idx < len(row) else None
        sheet_name = template_sheet_for_metric(metric)
        if sheet_name is None:
            continue
        key = (join_key, sheet_name)
        bucket = domain_index.setdefault(
            key,
            {"has_goal_target": False},
        )
        if _student_metric_has_goal_target(row, header_lookup):
            bucket["has_goal_target"] = True
        entry = _metric_summary_row_from_cells(row, header_lookup)
        summary_index.setdefault(key, []).append(entry)
    workbook.close()
    return domain_index, summary_index




def _sheet_student_rows(
    *,
    sheet_name: str,
    student_rows: list[dict[str, str]],
    student_metric_domain_index: dict[tuple[str, str], dict[str, bool]] | None,
    filter_enabled: bool,
) -> list[dict[str, str]]:
    if (
        not student_metric_domain_index
        or sheet_name not in METRIC_TEMPLATE_SHEETS
        or not filter_enabled
    ):
        return student_rows
    matched_rows = [
        row
        for row in student_rows
        if (jk := student_id_join_key(row.get("student_id"))) is not None
        and (jk, sheet_name) in student_metric_domain_index
    ]
    return matched_rows




def _build_comment_for_summary_rows(
    rows: list[dict[str, Any]],
    *,
    grading_period: str,
) -> str:
    def split_value_status(raw_text: str | None) -> tuple[str | None, str | None]:
        text = clean_text(raw_text)
        if text is None:
            return None, None
        if "|" not in text:
            return text, None
        left, right = text.split("|", 1)
        return clean_text(left), clean_text(right)

    lines: list[str] = []
    current_quarter: int | None = None
    try:
        period_index = int(float(grading_period)) - 1
        if 0 <= period_index < 4:
            current_quarter = period_index + 1
    except (TypeError, ValueError):
        current_quarter = None

    for block_index, row in enumerate(rows):
        if block_index:
            lines.append("---")
        baseline = _format_value_for_comment(row.get("Baseline"))
        target = _format_value_for_comment(row.get("Target"))
        has_goal_target = target is not None
        lines.append(f"Baseline: {baseline or 'N/A'}")
        if has_goal_target:
            lines.append(f"Target: {target}")

        if current_quarter is not None:
            quarter_range = range(1, max(1, current_quarter))
        else:
            quarter_range = range(1, 5)

        for quarter in quarter_range:
            period_header = METRIC_SUMMARY_GRADING_PERIOD_HEADERS[quarter - 1]
            parsed_value, quarter_status = split_value_status(_format_value_for_comment(row.get(period_header)))

            if quarter_status is None:
                for status_header in METRIC_SUMMARY_QUARTER_STATUS_HEADERS.get(quarter, ()):
                    quarter_status = clean_text(row.get(status_header))
                    if quarter_status is not None:
                        break
            if quarter_status is None:
                quarter_status = parsed_value

            lines.append(f"Q{quarter}: {quarter_status or 'N/A'}")
    return "\n".join(lines)




def attach_metric_summary_comments(
    workbook,
    *,
    template_schema: dict[str, Any],
    student_rows: list[dict[str, Any]],
    summary_index: dict[tuple[str, str], list[dict[str, Any]]],
    grading_period: str,
    sheet_student_rows: dict[str, list[dict[str, Any]]] | None = None,
) -> dict[str, Any]:
    comment_cells = 0
    sheets_touched: set[str] = set()
    for sheet_name, sheet_schema in template_schema["sheets"].items():
        if sheet_name == FIELD_OPTIONS_SHEET:
            continue
        headers = sheet_schema["headers"]
        progress_header = next((h for h in PROGRESS_AGAINST_GOAL_HEADERS if h in headers), None)
        if progress_header is None:
            continue
        column_index = headers.index(progress_header) + 1
        worksheet = workbook[sheet_name]
        row_limit = sheet_schema["max_row"]
        rows_for_sheet = student_rows if sheet_student_rows is None else sheet_student_rows.get(sheet_name, student_rows)
        for row_index, student in enumerate(rows_for_sheet, start=2):
            if row_limit > 1 and row_index > row_limit:
                break
            join_key = student_id_join_key(student.get("student_id"))
            if join_key is None:
                continue
            row_blocks = student.get("__metric_blocks")
            blocks = row_blocks if isinstance(row_blocks, list) and row_blocks else summary_index.get((join_key, sheet_name))
            if not blocks:
                continue
            school = clean_text(student.get("school"))
            filtered = [
                row
                for row in blocks
                if school is None or school_labels_equivalent(row.get("School"), school)
            ]
            chosen = filtered if filtered else blocks
            text = _build_comment_for_summary_rows(chosen, grading_period=grading_period)
            if not text:
                continue
            if len(text) > MAX_METRIC_COMMENT_LENGTH:
                text = text[: MAX_METRIC_COMMENT_LENGTH - 3] + "..."
            cell = worksheet.cell(row=row_index, column=column_index)
            note = Comment(text, "CISDM Wisdom")
            note.width = 420
            note.height = min(360, 96 + text.count("\n") * 18)
            cell.comment = note
            comment_cells += 1
            sheets_touched.add(sheet_name)
    return {
        "comment_cells": comment_cells,
        "sheets_with_comments": sorted(sheets_touched),
        "where_to_see_in_excel": EXCEL_METRIC_CONTEXT_NOTE_HINT,
    }




def apply_progress_column_width(
    workbook,
    *,
    template_schema: dict[str, Any],
    width: float = 30,
) -> None:
    for sheet_name, sheet_schema in _iter_template_sheets(template_schema, workbook=workbook):
        headers = sheet_schema["headers"]
        progress_header = next((h for h in PROGRESS_AGAINST_GOAL_HEADERS if h in headers), None)
        if progress_header is None:
            continue
        column_index = headers.index(progress_header) + 1
        column_letter = get_column_letter(column_index)
        workbook[sheet_name].column_dimensions[column_letter].width = width




def _column_index_for_exact_header(headers: list[str], title: str) -> int | None:
    # why: full header equality after strip only — no substring match (avoids wrong columns)
    want = title.strip()
    for index, raw in enumerate(headers, start=1):
        if str(raw).strip() == want:
            return index
    return None




def apply_hidden_provision_columns(
    workbook,
    *,
    template_schema: dict[str, Any],
    headers_to_hide: tuple[str, ...] = HIDDEN_PROVISION_HEADERS,
) -> None:
    for sheet_name, sheet_schema in _iter_template_sheets(template_schema, workbook=workbook):
        headers = sheet_schema["headers"]
        worksheet = workbook[sheet_name]
        for title in headers_to_hide:
            column_index = _column_index_for_exact_header(headers, title)
            if column_index is None:
                continue
            column_letter = get_column_letter(column_index)
            worksheet.column_dimensions[column_letter].hidden = True




def _case_manager_names_from_roster(roster_lookup: dict[str, dict[str, str]]) -> list[str]:
    coordinators: list[str] = []
    seen: set[str] = set()
    for row in roster_lookup.values():
        _append_unique_name(coordinators, seen, row.get("created_by"))
    coordinators.sort(key=lambda name: normalize_text(name))
    return coordinators




def roster_coordinator_matches_active_site_staff(
    roster_name: object,
    active_site_coordinator_names: list[str],
) -> bool:
    return any(person_name_equivalent(roster_name, active) for active in active_site_coordinator_names)




def _site_list_end_date_means_active_assignment(value: object) -> bool:
    # why: blank end date => current site assignment; any real date => ended
    if value is None:
        return True
    if isinstance(value, (datetime, date)):
        return False
    text = clean_text(value)
    if text is None:
        return True
    if text == "-":
        return True
    return False




def load_active_site_coordinator_display_names(site_staff_list_path: Path) -> list[str]:
    # why: Site List rows with a past End Date are remnant staffing; blank End Date => current coordinator
    if not site_staff_list_path.exists():
        raise ValueError(f"Site staff workbook not found: {site_staff_list_path}")
    workbook = load_workbook(site_staff_list_path, read_only=True, data_only=True)
    if SITE_LIST_SHEET not in workbook.sheetnames:
        workbook.close()
        raise ValueError(f"Sheet {SITE_LIST_SHEET!r} not found in {site_staff_list_path.name}")
    worksheet = workbook[SITE_LIST_SHEET]
    headers = header_values(worksheet)
    try:
        idx_sc = headers.index(SITE_LIST_SITE_COORDINATOR_HEADER)
        idx_ed = headers.index(SITE_LIST_END_DATE_HEADER)
    except ValueError as exc:
        workbook.close()
        raise ValueError(
            f"Missing {SITE_LIST_SITE_COORDINATOR_HEADER!r} or {SITE_LIST_END_DATE_HEADER!r} in {SITE_LIST_SHEET}"
        ) from exc
    names: list[str] = []
    seen: set[str] = set()
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        raw_sc = row[idx_sc] if idx_sc < len(row) else None
        raw_ed = row[idx_ed] if idx_ed < len(row) else None
        if not _site_list_end_date_means_active_assignment(raw_ed):
            continue
        _append_unique_name(names, seen, raw_sc)
    workbook.close()
    return names




def load_active_site_school_names(site_staff_list_path: Path) -> list[str]:
    # why: site list with blank End Date marks current schools for site-level provisioning
    if not site_staff_list_path.exists():
        raise ValueError(f"Site staff workbook not found: {site_staff_list_path}")
    workbook = load_workbook(site_staff_list_path, read_only=True, data_only=True)
    if SITE_LIST_SHEET not in workbook.sheetnames:
        workbook.close()
        raise ValueError(f"Sheet {SITE_LIST_SHEET!r} not found in {site_staff_list_path.name}")
    worksheet = workbook[SITE_LIST_SHEET]
    headers = header_values(worksheet)
    try:
        idx_school = headers.index("School")
        idx_ed = headers.index(SITE_LIST_END_DATE_HEADER)
    except ValueError as exc:
        workbook.close()
        raise ValueError(
            f"Missing 'School' or {SITE_LIST_END_DATE_HEADER!r} in {SITE_LIST_SHEET}"
        ) from exc
    schools: list[str] = []
    seen: set[str] = set()
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        raw_school = row[idx_school] if idx_school < len(row) else None
        raw_ed = row[idx_ed] if idx_ed < len(row) else None
        if not _site_list_end_date_means_active_assignment(raw_ed):
            continue
        school_name = clean_text(raw_school)
        if school_name is None:
            continue
        school_key = normalize_text(school_name)
        if school_key in seen:
            continue
        seen.add(school_key)
        schools.append(school_name)
    workbook.close()
    return schools




def filter_student_rows_by_active_site_coordinators(
    student_rows: list[dict[str, str]],
    active_site_coordinator_names: list[str],
) -> list[dict[str, str]]:
    return [
        row
        for row in student_rows
        if roster_coordinator_matches_active_site_staff(row.get("created_by"), active_site_coordinator_names)
    ]




def _sort_student_rows(rows: list[dict[str, str]]) -> None:
    rows.sort(
        key=lambda item: (
            item.get("school", ""),
            item.get("last_name", ""),
            item.get("first_name", ""),
            item.get("student_id", ""),
        )
    )




def _schools_from_rows(rows: list[dict[str, str]]) -> set[str]:
    return {school for row in rows if (school := clean_text(row.get("school")))}




def _rows_matching_coordinator(
    roster_lookup: dict[str, dict[str, str]],
    coordinator_name: str,
) -> list[dict[str, str]]:
    return [
        row
        for row in roster_lookup.values()
        if person_name_equivalent(row.get("created_by"), coordinator_name)
    ]




def _rows_matching_resolved_schools(
    roster_lookup: dict[str, dict[str, str]],
    resolved_schools: set[str],
) -> list[dict[str, str]]:
    if not resolved_schools:
        return []
    return [
        row
        for row in roster_lookup.values()
        if caseload_school_matches_resolved(row.get("school"), resolved_schools)
    ]




def _rows_matching_creator_text(
    roster_lookup: dict[str, dict[str, str]],
    coordinator_name: str | None,
) -> list[dict[str, str]]:
    exact_name = clean_text(coordinator_name)
    if exact_name is None:
        return []
    return [row for row in roster_lookup.values() if clean_text(row.get("created_by")) == exact_name]




def _group_student_rows_by_site(student_rows: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    groups: list[dict[str, Any]] = []
    for row in student_rows:
        school_name = clean_text(row.get("school"))
        if school_name is None:
            continue
        row_keys = site_group_match_keys(school_name)
        if not row_keys:
            continue
        target_group: dict[str, Any] | None = None
        for group in groups:
            if group["keys"] & row_keys:
                target_group = group
                break
        if target_group is None:
            target_group = {"keys": set(row_keys), "rows": [], "labels": []}
            groups.append(target_group)
        else:
            target_group["keys"].update(row_keys)
        target_group["rows"].append(row)
        target_group["labels"].append(school_name)
    grouped: dict[str, list[dict[str, str]]] = {}
    for group in groups:
        rows = list(group["rows"])
        _sort_student_rows(rows)
        label = school_label(rows)
        grouped[label] = rows
    return grouped




def filter_student_rows_by_active_site_schools(
    student_rows: list[dict[str, str]],
    active_site_schools: list[str],
) -> list[dict[str, str]]:
    if not active_site_schools:
        return []
    resolved = {school for school in active_site_schools if clean_text(school)}
    return [
        row
        for row in student_rows
        if caseload_school_matches_resolved(row.get("school"), resolved)
    ]




def select_student_rows(
    *,
    roster_lookup: dict[str, dict[str, str]],
    coordinator_name: str | None = None,
    school_name: str | None = None,
) -> tuple[list[dict[str, str]], set[str]]:
    resolved_schools: set[str] = set()
    if school_name is not None:
        resolved_schools.add(school_name)

    if coordinator_name is None:
        if school_name is None:
            return [], set()
        student_rows = _rows_matching_resolved_schools(roster_lookup, {school_name})
        _sort_student_rows(student_rows)
        return student_rows, {school_name}

    student_rows = _rows_matching_coordinator(roster_lookup, coordinator_name)
    if student_rows:
        _sort_student_rows(student_rows)
        return student_rows, _schools_from_rows(student_rows)

    student_rows = _rows_matching_resolved_schools(roster_lookup, resolved_schools)
    if student_rows:
        _sort_student_rows(student_rows)
        return student_rows, resolved_schools

    student_rows = _rows_matching_creator_text(roster_lookup, coordinator_name)
    _sort_student_rows(student_rows)
    if student_rows:
        return student_rows, _schools_from_rows(student_rows)
    return [], resolved_schools




def school_year_folder_name(student_rows: list[dict[str, str]], template_schema: dict[str, Any]) -> str:
    school_year = ""
    for row in student_rows:
        school_year = clean_text(row.get("school_year")) or ""
        if school_year:
            break
    if not school_year:
        defaults = template_schema["field_options"].get("School Year", [])
        school_year = defaults[0] if defaults else ""
    digits = [part for part in school_year.replace("SY", "").replace("/", " ").split() if part.isdigit()]
    if len(digits) >= 2:
        return f"SY{digits[0][-2:]}-{digits[1][-2:]}"
    return "SY-Unknown"




def quarter_label(grading_period: str) -> str:
    return f"Q{int(float(grading_period))}"




def filename_component(value: str) -> str:
    cleaned = "".join("-" if character in '\\/:*?"<>|' else character for character in value)
    return " ".join(cleaned.split()).strip() or "Unknown"




def school_label(student_rows: list[dict[str, str]]) -> str:
    # why: filenames need a single site label; caseload often has one dominant school plus a few strays
    schools = sorted({school for row in student_rows if (school := clean_text(row.get("school")))})
    if not schools:
        return "Multi-School"
    components = partition_schools_by_label_equivalence(schools)
    if len(components) == 1:
        return max(components[0], key=len)
    school_to_ci: dict[str, int] = {}
    for ci, comp in enumerate(components):
        for name in comp:
            school_to_ci[name] = ci
    counts: dict[int, int] = {}
    for row in student_rows:
        s = clean_text(row.get("school"))
        if s is None:
            continue
        ci = school_to_ci.get(s)
        if ci is None:
            continue
        counts[ci] = counts.get(ci, 0) + 1
    if not counts:
        return "Multi-School"
    max_count = max(counts.values())
    winners = [ci for ci, n in counts.items() if n == max_count]
    if len(winners) > 1:
        return "Multi-School"
    chosen = components[winners[0]]
    return max(chosen, key=len)




def default_batch_output_root(template_path: Path) -> Path:
    explicit = (os.environ.get("QPR_OUTPUT_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return DEFAULT_QPR_OUTPUT_DIR




def matching_schedule(schedule_data: dict[str, Any], school: str | None) -> tuple[str | None, dict[str, Any] | None]:
    raw = clean_text(school)
    if raw is None:
        return None, None
    schedules = schedule_data.get("schedules", {})
    if raw in schedules:
        return raw, schedules.get(raw)
    key = schedule_key_for_school(raw)
    normalized_key = normalize_text(key) if key else None
    if normalized_key:
        for name in schedules:
            if normalize_text(name) == normalized_key:
                return name, schedules[name]
    district_lookup = schedule_data.get("school_district_lookup", {})
    district = district_lookup.get(normalize_text(raw))
    if district is None:
        compact_key = normalize_text(raw).replace(" school", "").strip()
        if compact_key:
            district = district_lookup.get(compact_key)
    if district is None:
        for candidate in school_match_keys(raw):
            district = district_lookup.get(candidate)
            if district is not None:
                break
    if district is not None:
        normalized_district = normalize_text(district)
        normalized_district = {
            "antietam": "antietem",
        }.get(normalized_district, normalized_district)
        for name in schedules:
            if normalize_text(name) == normalized_district:
                return name, schedules[name]
    return None, None




def filter_student_rows_by_enrollment_end_before_assessment(
    student_rows: list[dict[str, str]],
    schedule_data: dict[str, Any],
    grading_period: str,
) -> list[dict[str, str]]:
    # why: metric roster includes exited students; QPR is for those still enrolled through assessment
    gp = canonical_grading_period(grading_period)
    kept: list[dict[str, str]] = []
    for row in student_rows:
        end_iso = clean_text(row.get("enrollment_end_iso"))
        if not end_iso:
            kept.append(row)
            continue
        try:
            end_d = datetime.strptime(end_iso, "%Y-%m-%d").date()
        except ValueError:
            kept.append(row)
            continue
        _, schedule = matching_schedule(schedule_data, row.get("school"))
        if schedule is None:
            kept.append(row)
            continue
        assess_iso = schedule["assessment_dates"].get(gp)
        if assess_iso is None:
            kept.append(row)
            continue
        try:
            assess_d = datetime.strptime(assess_iso, "%Y-%m-%d").date()
        except ValueError:
            kept.append(row)
            continue
        if end_d < assess_d:
            continue
        kept.append(row)
    return kept




def schedule_status(
    schedule_data: dict[str, Any],
    school: str | None,
) -> tuple[str, str | None, dict[str, Any] | None]:
    schedule_name, schedule = matching_schedule(schedule_data, school)
    if schedule_name is None or schedule is None:
        return "unmatched", None, None
    if schedule_name in schedule_data["reference_issues"]:
        return "suspect", schedule_name, schedule
    return "matched", schedule_name, schedule




def _period_issue_prefix(grading_period: str) -> str:
    return f"period_{int(float(grading_period))}_"




def schedule_period_has_reference_issue(
    schedule_data: dict[str, Any],
    schedule_name: str,
    grading_period: str | None,
) -> bool:
    issues = schedule_data["reference_issues"].get(schedule_name, [])
    if not issues:
        return False
    if grading_period is None:
        return True
    period_prefix = _period_issue_prefix(grading_period)
    return any(code.startswith(period_prefix) for code in issues)




def populate_known_fields(
    workbook,
    *,
    template_schema: dict[str, Any],
    schedule_data: dict[str, Any],
    student_rows: list[dict[str, str]],
    grading_period: str,
    student_metric_domain_index: dict[tuple[str, str], dict[str, bool]] | None = None,
    uat: bool = False,
    summary_index: dict[tuple[str, str], list[dict[str, Any]]] | None = None,
) -> dict[str, Any]:
    count = 0
    matched_schools: set[str] = set()
    suspect_schools: set[str] = set()
    unmatched_schools: set[str] = set()
    tabs: dict[str, dict[str, int]] = {}
    sheet_student_rows: dict[str, list[dict[str, Any]]] = {}
    allowed_core = _core_course_field_options(template_schema)
    allowed_stand_test = _stand_test_type_field_options(template_schema)
    allowed_sel_metric = _sel_metric_field_options(template_schema)
    allowed_credits = _credits_field_options(template_schema)
    metric_filter_enabled = False
    if student_metric_domain_index:
        metric_student_ids = {jk for jk, _ in student_metric_domain_index}
        metric_filter_enabled = any(
            (jk := student_id_join_key(row.get("student_id"))) is not None and jk in metric_student_ids
            for row in student_rows
        )
    for sheet_name, sheet_schema in _iter_template_sheets(template_schema):
        worksheet = workbook[sheet_name]
        header_lookup = {header: index for index, header in enumerate(sheet_schema["headers"], start=1)}
        row_limit = sheet_schema["max_row"]
        rows_for_sheet = _sheet_student_rows(
            sheet_name=sheet_name,
            student_rows=student_rows,
            student_metric_domain_index=student_metric_domain_index,
            filter_enabled=metric_filter_enabled,
        )
        rows_for_sheet_expanded: list[dict[str, Any]] = [dict(row) for row in rows_for_sheet]
        if sheet_name == CORE_COURSE_GRADES_SHEET and CORE_COURSE_HEADER in header_lookup:
            rows_for_sheet_expanded = _expand_metric_option_rows_for_sheet(
                rows_for_sheet=rows_for_sheet,
                sheet_name=sheet_name,
                summary_index=summary_index,
                allowed_options=allowed_core,
                extract=template_core_course_from_metric,
            )
        elif sheet_name == STANDARDIZED_TEST_SCORE_SHEET and STAND_TEST_TYPE_HEADER in header_lookup:
            rows_for_sheet_expanded = _expand_metric_option_rows_for_sheet(
                rows_for_sheet=rows_for_sheet,
                sheet_name=sheet_name,
                summary_index=summary_index,
                allowed_options=allowed_stand_test,
                extract=template_stand_test_type_from_metric,
            )
        elif sheet_name == SEL_SHEET and SEL_METRIC_HEADER in header_lookup:
            rows_for_sheet_expanded = _expand_metric_option_rows_for_sheet(
                rows_for_sheet=rows_for_sheet,
                sheet_name=sheet_name,
                summary_index=summary_index,
                allowed_options=allowed_sel_metric,
                extract=template_sel_metric_from_metric,
            )
        elif sheet_name == CREDITS_CTE_SHEET and CREDITS_METRIC_HEADER in header_lookup:
            rows_for_sheet_expanded = _expand_metric_option_rows_for_sheet(
                rows_for_sheet=rows_for_sheet,
                sheet_name=sheet_name,
                summary_index=summary_index,
                allowed_options=allowed_credits,
                extract=template_credits_metric_from_metric,
            )
        sheet_student_rows[sheet_name] = rows_for_sheet_expanded
        progress_header = next((name for name in PROGRESS_AGAINST_GOAL_HEADERS if name in header_lookup), None)
        rows_populated = 0
        for row_index, student in enumerate(rows_for_sheet_expanded, start=2):
            if row_limit > 1 and row_index > row_limit:
                break
            school_name = student.get("school")
            display_student_id = clean_text(student.get("entity_id")) if uat else student.get("student_id")
            last_name_out = "Test" if uat else student.get("last_name")
            known_values = {
                "School": school_name,
                "Student ID": display_student_id,
                "First Name": student.get("first_name"),
                "Last Name": last_name_out,
                "Organization": ORGANIZATION_CANONICAL,
                "School Year": student.get("school_year") or template_schema["field_options"].get("School Year", [""])[0],
                "Grading Period": grading_period_cell_value(grading_period),
            }
            status, schedule_name, schedule = schedule_status(schedule_data, school_name)
            if status == "unmatched" or schedule_name is None or schedule is None:
                if school_name:
                    unmatched_schools.add(school_name)
            elif status == "suspect":
                if schedule_period_has_reference_issue(schedule_data, schedule_name, grading_period):
                    if school_name:
                        suspect_schools.add(school_name)
                else:
                    if school_name:
                        matched_schools.add(school_name)
                    known_values["Grading Period / Assessment Date"] = assessment_date_cell_value(
                        schedule["assessment_dates"].get(grading_period),
                    )
            else:
                if school_name:
                    matched_schools.add(school_name)
                known_values["Grading Period / Assessment Date"] = assessment_date_cell_value(
                    schedule["assessment_dates"].get(grading_period),
                )
            metric_option = clean_text(student.get("__metric_option"))
            if metric_option:
                if sheet_name == CORE_COURSE_GRADES_SHEET and CORE_COURSE_HEADER in header_lookup:
                    known_values[CORE_COURSE_HEADER] = metric_option
                elif sheet_name == STANDARDIZED_TEST_SCORE_SHEET and STAND_TEST_TYPE_HEADER in header_lookup:
                    known_values[STAND_TEST_TYPE_HEADER] = metric_option
                elif sheet_name == SEL_SHEET and SEL_METRIC_HEADER in header_lookup:
                    known_values[SEL_METRIC_HEADER] = metric_option
                elif sheet_name == CREDITS_CTE_SHEET and CREDITS_METRIC_HEADER in header_lookup:
                    known_values[CREDITS_METRIC_HEADER] = metric_option
            for header_name, value in known_values.items():
                if value is None or header_name not in header_lookup:
                    continue
                worksheet.cell(row=row_index, column=header_lookup[header_name]).value = value
            if progress_header is not None and (jk := student_id_join_key(student.get("student_id"))) is not None:
                row_level_goal = student.get("__metric_has_goal_target")
                has_goal_target = (
                    bool(row_level_goal)
                    if row_level_goal is not None
                    else bool(
                        student_metric_domain_index
                        and student_metric_domain_index.get((jk, sheet_name), {}).get("has_goal_target")
                    )
                )
                if not has_goal_target:
                    progress_column = header_lookup[progress_header]
                    progress_cell = worksheet.cell(row=row_index, column=progress_column)
                    if clean_text(progress_cell.value) is None:
                        progress_cell.value = NO_GOAL_ASSIGNED_TEXT
            rows_populated = row_index - 1
        count = max(count, len(student_rows))
        tabs[sheet_name] = {"rows_populated": rows_populated}
    result: dict[str, Any] = {
        "student_count": count,
        "schedule_prefill": {
            "matched_schools": sorted(matched_schools),
            "suspect_schools": sorted(suspect_schools),
            "unmatched_schools": sorted(unmatched_schools),
        },
        "tabs": tabs,
        "sheet_student_rows": sheet_student_rows,
    }
    return result




def write_provisioned_template(
    *,
    output_path: Path,
    template_path: Path,
    template_schema: dict[str, Any],
    schedule_data: dict[str, Any],
    student_rows: list[dict[str, str]],
    grading_period: str,
    coordinator_name: str | None,
    school_name: str | None,
    resolved_schools: set[str],
    metric_workbook_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    uat: bool = False,
) -> dict[str, Any]:
    student_rows = filter_student_rows_by_enrollment_end_before_assessment(
        student_rows, schedule_data, grading_period
    )
    if not student_rows:
        raise ValueError(NO_STUDENTS_AFTER_ENROLLMENT_FILTER)
    effective_rows, uat_skipped_ids = effective_student_rows_for_uat(student_rows, uat)
    if uat and not effective_rows:
        raise ValueError(UAT_NO_ENTITY_ID_ROWS)
    workbook = load_workbook(template_path, data_only=False)
    student_metric_domain_index: dict[tuple[str, str], dict[str, bool]] | None = None
    summary_index: dict[tuple[str, str], list[dict[str, Any]]] = {}
    if metric_workbook_path is not None:
        if not metric_workbook_path.exists():
            raise ValueError(f"Metric tracking workbook not found: {metric_workbook_path}")
        student_metric_domain_index, summary_index = load_metric_tracking_workbook(metric_workbook_path)
    populate_result = populate_known_fields(
        workbook,
        template_schema=template_schema,
        schedule_data=schedule_data,
        student_rows=effective_rows,
        grading_period=grading_period,
        student_metric_domain_index=student_metric_domain_index,
        uat=uat,
        summary_index=summary_index,
    )
    metric_summary_comments: dict[str, Any] | None = None
    if metric_workbook_path is not None:
        if summary_index:
            metric_summary_comments = attach_metric_summary_comments(
                workbook,
                template_schema=template_schema,
                student_rows=effective_rows,
                summary_index=summary_index,
                grading_period=grading_period,
                sheet_student_rows=populate_result.get("sheet_student_rows"),
            )
        else:
            metric_summary_comments = {
                "comment_cells": 0,
                "sheets_with_comments": [],
                "note": "no_matched_metric_rows",
                "where_to_see_in_excel": EXCEL_METRIC_CONTEXT_NOTE_HINT,
            }
    apply_progress_column_width(
        workbook,
        template_schema=template_schema,
        width=30,
    )
    apply_hidden_provision_columns(
        workbook,
        template_schema=template_schema,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    result: dict[str, Any] = {
        "output_path": str(output_path.resolve()),
        "student_count": populate_result["student_count"],
        "grading_period": grading_period,
        "coordinator_name": coordinator_name,
        "school_name": school_name,
        "resolved_schools": sorted(resolved_schools),
        "schedule_prefill": populate_result["schedule_prefill"],
        "tabs": populate_result["tabs"],
    }
    if metric_summary_comments is not None:
        result["metric_summary_comments"] = metric_summary_comments
    if uat:
        result["uat"] = True
        result["uat_skipped_missing_entity_id"] = sorted(set(uat_skipped_ids))
    return result




def _load_provision_inputs(
    template_path: Path,
    deadlines_path: Path,
    roster_source_path: Path | None,
) -> tuple[dict[str, Any], dict[str, Any], dict[str, dict[str, str]]]:
    template_schema = build_template_schema(template_path)
    schedule_data = load_reporting_deadlines(deadlines_path)
    roster_lookup = load_metric_roster_lookup(roster_source_path)
    return template_schema, schedule_data, roster_lookup




def provision_site_template(
    *,
    output_path: Path,
    grading_period: str,
    coordinator_name: str | None = None,
    school_name: str | None = None,
    template_path: Path = DEFAULT_TEMPLATE_PATH,
    deadlines_path: Path = DEFAULT_DEADLINES_PATH,
    roster_source_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    metric_workbook_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    uat: bool = False,
    site_staff_list_path: Path | None = DEFAULT_SITE_STAFF_LIST_PATH,
    use_site_staff_filter: bool = True,
) -> dict[str, Any]:
    if coordinator_name is None and school_name is None:
        raise ValueError("Provide coordinator_name or school_name.")
    grading_period = canonical_grading_period(grading_period)
    template_schema, schedule_data, roster_lookup = _load_provision_inputs(
        template_path,
        deadlines_path,
        roster_source_path,
    )
    active_site: list[str] | None = None
    if use_site_staff_filter:
        staff_path = site_staff_list_path or DEFAULT_SITE_STAFF_LIST_PATH
        active_site = load_active_site_coordinator_display_names(staff_path)
        if coordinator_name is not None and not roster_coordinator_matches_active_site_staff(
            coordinator_name, active_site
        ):
            raise ValueError(INACTIVE_SITE_COORDINATOR_PROVISION)
    student_rows, resolved_schools = select_student_rows(
        roster_lookup=roster_lookup,
        coordinator_name=coordinator_name,
        school_name=school_name,
    )
    if use_site_staff_filter and active_site is not None and school_name is not None:
        student_rows = filter_student_rows_by_active_site_coordinators(student_rows, active_site)
    return write_provisioned_template(
        output_path=output_path,
        template_path=template_path,
        template_schema=template_schema,
        schedule_data=schedule_data,
        student_rows=student_rows,
        grading_period=grading_period,
        coordinator_name=coordinator_name,
        school_name=school_name,
        resolved_schools=resolved_schools,
        metric_workbook_path=metric_workbook_path,
        uat=uat,
    )




def provision_all_site_templates(
    *,
    output_directory: Path | None = None,
    grading_period: str,
    template_path: Path = DEFAULT_TEMPLATE_PATH,
    deadlines_path: Path = DEFAULT_DEADLINES_PATH,
    roster_source_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    metric_workbook_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    uat: bool = False,
    site_staff_list_path: Path | None = DEFAULT_SITE_STAFF_LIST_PATH,
    use_site_staff_filter: bool = True,
) -> dict[str, Any]:
    grading_period = canonical_grading_period(grading_period)
    template_schema, schedule_data, roster_lookup = _load_provision_inputs(
        template_path,
        deadlines_path,
        roster_source_path,
    )
    active_site_coordinators: list[str] | None = None
    active_site_schools: list[str] | None = None
    if use_site_staff_filter:
        staff_path = site_staff_list_path or DEFAULT_SITE_STAFF_LIST_PATH
        active_site_coordinators = load_active_site_coordinator_display_names(staff_path)
        active_site_schools = load_active_site_school_names(staff_path)
    all_case_managers = _case_manager_names_from_roster(roster_lookup)
    all_student_rows = list(roster_lookup.values())
    roster_student_count = len(all_student_rows)
    if use_site_staff_filter and active_site_schools is not None:
        all_student_rows = filter_student_rows_by_active_site_schools(
            all_student_rows, active_site_schools
        )
    site_rows = _group_student_rows_by_site(all_student_rows)
    excluded_by_site_staff = (
        [
            name
            for name in all_case_managers
            if not roster_coordinator_matches_active_site_staff(name, active_site_coordinators)
        ]
        if active_site_coordinators is not None
        else []
    )
    results: list[dict[str, Any]] = []
    skipped: list[dict[str, str]] = []
    quarter = quarter_label(grading_period)
    school_year_folder = school_year_folder_name(list(roster_lookup.values()), template_schema)
    if output_directory is None:
        output_directory = default_batch_output_root(template_path)
    batch_output_directory = output_directory / school_year_folder / quarter
    for site_name, student_rows in sorted(site_rows.items(), key=lambda item: normalize_text(item[0])):
        resolved_schools = _schools_from_rows(student_rows)
        if not student_rows:
            skipped.append({"site_name": site_name, "reason": "no_students"})
            continue
        filename = f"{quarter}_{filename_component(site_name)}_QPR.xlsx"
        try:
            result = write_provisioned_template(
                output_path=batch_output_directory / filename,
                template_path=template_path,
                template_schema=template_schema,
                schedule_data=schedule_data,
                student_rows=student_rows,
                grading_period=grading_period,
                coordinator_name=None,
                school_name=None,
                resolved_schools=resolved_schools,
                metric_workbook_path=metric_workbook_path,
                uat=uat,
            )
            result["site_name"] = site_name
        except ValueError as exc:
            msg = str(exc)
            if uat and msg == UAT_NO_ENTITY_ID_ROWS:
                skipped.append({"site_name": site_name, "reason": "uat_no_students_with_entity_id"})
                continue
            if msg == NO_STUDENTS_AFTER_ENROLLMENT_FILTER:
                skipped.append({"site_name": site_name, "reason": "no_students_after_enrollment_filter"})
                continue
            raise
        results.append(result)
    site_count = len(site_rows)
    warning: str | None = None
    if use_site_staff_filter and roster_student_count and not all_student_rows:
        warning = (
            "site staff filter removed all roster rows; pass a matching --site-staff-list "
            "or use --no-site-staff-filter to provision every site in the metrics roster"
        )
    return {
        "grading_period": grading_period,
        "output_directory": str(batch_output_directory.resolve()),
        "sites_processed": site_count,
        "sites_generated": len(results),
        "sites_skipped": len(skipped),
        "coordinators_processed": site_count,
        "coordinators_generated": len(results),
        "coordinators_skipped": len(skipped),
        "total_students": sum(result["student_count"] for result in results),
        "roster_student_count": roster_student_count,
        "results": results,
        "skipped": skipped,
        "use_site_staff_filter": use_site_staff_filter,
        "roster_coordinators_excluded_by_site_staff": excluded_by_site_staff,
        "excel_metric_context_note_hint": EXCEL_METRIC_CONTEXT_NOTE_HINT,
        "warning": warning,
    }




def build_provision_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Provision progress monitoring import workbooks for each unique Case Manager in the Student Metrics summary."
        ),
    )
    parser.add_argument("output_directory", nargs="?", type=Path, help="Directory where tailored workbooks should be written.")
    parser.add_argument("--grading-period", required=True, help="Canonical grading period, e.g. 2.0.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_PATH, help="Template workbook path.")
    parser.add_argument("--deadlines", type=Path, default=DEFAULT_DEADLINES_PATH, help="Reporting deadlines workbook path.")
    parser.add_argument(
        "--metric-workbook",
        type=Path,
        default=DEFAULT_METRIC_WORKBOOK_PATH,
        help=(
            "Student Metrics summary (.xlsx): roster (Case Manager, School, Student ID, names) plus when not combined "
            "with --no-metric-workbook: tab membership, Target-based Progress prefill, read-only Excel Notes on "
            "Progress Against Goal (Review > Notes in Excel; not the Comments pane). "
            f"Default: {DEFAULT_METRIC_WORKBOOK_PATH.name}."
        ),
    )
    parser.add_argument(
        "--no-metric-workbook",
        action="store_true",
        help=(
            "Skip metric tracking features: no per-tab domain filter, no non-goal Progress prefill, no metric Notes. "
            "Roster rows are still loaded from --metric-workbook."
        ),
    )
    parser.add_argument(
        "--uat",
        action="store_true",
        help=(
            "UAT/training templates: write Student ID from EntityID or Client ID and Last Name as Test; "
            "rows without either id are omitted. Roster still comes from --metric-workbook."
        ),
    )
    parser.add_argument(
        "--site-staff-list",
        type=Path,
        help=(
            f"Site/staff workbook containing sheet {SITE_LIST_SHEET!r}: restrict provision to schools "
            f"with a blank {SITE_LIST_END_DATE_HEADER!r}. Omit to provision every site in the metrics roster."
        ),
    )
    parser.add_argument(
        "--no-site-staff-filter",
        action="store_true",
        help=(
            "Do not restrict provision to Site List coordinators with blank End Date (use every Case Manager from metrics)."
        ),
    )
    return parser




def _maybe_prefetch_student_metrics_for_cli(argv: list[str]):
    try:
        import qpr as prefetch
    except ModuleNotFoundError:
        script_dir = Path(__file__).resolve().parent
        if script_dir.name == "scripts":
            qpr_home = script_dir.parent
            if str(qpr_home) not in sys.path:
                sys.path.insert(0, str(qpr_home))
        import qpr as prefetch
    return prefetch.maybe_prefetch_student_metrics(argv)






def main(argv: list[str] | None = None) -> None:
    if argv is None:
        argv = sys.argv[1:]
    try:
        _maybe_prefetch_student_metrics_for_cli(argv)
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        raise RuntimeError(f"cisiphyus_prefetch_failed: {exc}") from exc
    if not argv or argv[0] not in ("provision", "provision-batch"):
        raise SystemExit(
            "usage: qpr_tools.py provision-batch <output_directory> --grading-period <period> [options]"
        )
    args = build_provision_parser().parse_args(argv[1:])
    roster_source_path = args.metric_workbook
    metric_workbook_path = None if args.no_metric_workbook else args.metric_workbook
    use_site_staff_filter = args.site_staff_list is not None and not args.no_site_staff_filter
    result = provision_all_site_templates(
        output_directory=args.output_directory,
        grading_period=args.grading_period,
        template_path=args.template,
        deadlines_path=args.deadlines,
        roster_source_path=roster_source_path,
        metric_workbook_path=metric_workbook_path,
        uat=args.uat,
        site_staff_list_path=args.site_staff_list,
        use_site_staff_filter=use_site_staff_filter,
    )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
