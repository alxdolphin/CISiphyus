#!/usr/bin/env python3

from __future__ import annotations

import argparse
import base64
import csv
import importlib.util
import json
import os
import re
import sys
from collections import defaultdict
from urllib.parse import urlparse
from email.message import EmailMessage
from functools import lru_cache
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Callable, Literal

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils.cell import get_column_letter

_GOOGLE_CREDENTIALS_CLASS = None
_GOOGLE_BUILD_FN = None

# WHY: vendored for CISiphyus examples; paths resolve under examples/qpr/
QPR_DIR = Path(__file__).resolve().parent
FIXTURES_DIR = QPR_DIR / "fixtures"
LOCAL_INPUTS_DIR = QPR_DIR / "local_inputs"
REPO_ROOT = QPR_DIR.parents[1]

PROJECT_ROOT = REPO_ROOT
DEFAULT_TEMPLATE_PATH = FIXTURES_DIR / "qpr_import_template.xlsx"
DEFAULT_DEADLINES_PATH = FIXTURES_DIR / "reporting_deadlines_minimal.xlsx"
DEFAULT_CASELOAD_PATH = LOCAL_INPUTS_DIR / "SY25-26_CaseloadExport.xlsx"
DEFAULT_QPR_DIR = QPR_DIR
DEFAULT_WRITTEN_QPR_OUTPUT_DIR = QPR_DIR / "written"
DEFAULT_QPR_LOCAL_INPUTS_DIR = LOCAL_INPUTS_DIR
DEFAULT_STUDENT_METRICS_FILENAME = "SY25-26_StudentMetricsSummary.xlsx"
DEFAULT_METRIC_WORKBOOK_PATH = DEFAULT_QPR_LOCAL_INPUTS_DIR / DEFAULT_STUDENT_METRICS_FILENAME
LEGACY_METRIC_WORKBOOK_PATH = DEFAULT_METRIC_WORKBOOK_PATH
DEFAULT_SITE_STAFF_LIST_PATH = FIXTURES_DIR / "site_staff_list_minimal.xlsx"
DEFAULT_WRITTEN_QPR_DRAFT_DOCX_PATH = QPR_DIR / "written" / "v5" / "ISS QPR 2026-27 [draft V2].docx"
# v2 template table indices (written/v5/ISS QPR 2026-27 [draft V2].docx)
QPR_TABLE_HEADER = 0
QPR_TABLE_NARRATIVE = 1
QPR_TABLE_SITE_COORDINATION = 2
QPR_TABLE_SCHOOL_GOALS = 3
QPR_TABLE_CASE_OUTCOMES = 4
QPR_TABLE_CASELOAD = 5
QPR_TABLE_REFERRALS = 6
QPR_TABLE_BASIC_NEEDS = 7
QPR_TABLE_PARTNERS = 8
QPR_TABLE_TIER_I = 9
QPR_TABLE_TIER_II = 10
QPR_TABLE_TIER_III = 11
QPR_TABLE_STORY_WHOLE_SCHOOL = 12
QPR_TABLE_STORY_GROUP = 13
QPR_TABLE_STORY_INDIVIDUAL = 14
QPR_TABLE_MIN_COUNT = 12
SITE_LIST_SHEET = "Site List"
SITE_LIST_SITE_COORDINATOR_HEADER = "Site Coordinator"
SITE_LIST_CS_COORDINATOR_HEADER = "CS Coordinator"
SITE_LIST_END_DATE_HEADER = "End Date"
SITE_LIST_MANAGER_HEADER = "Manager"
SITE_LIST_DATA_POC_HEADER = "Data POC"
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
CASELOAD_BUILDUP_TEMPLATE_SCHOOL_YEARS = (
    "2024-2025",
    "2025-2026",
    "2026-2027",
)
MAX_METRIC_COMMENT_LENGTH = 1000
# why: openpyxl writes legacy Notes (vml); Excel 365's Comments pane is threaded-only — users need Review > Notes
EXCEL_METRIC_CONTEXT_NOTE_HINT = (
    "Metric context is stored as Excel Notes on Progress Against Goal cells "
    "(Review tab > Notes, or the small corner triangle), not in the modern Comments pane."
)
PROGRESS_AGAINST_GOAL_HEADERS = ("Progress Against Goal", "Progress against goal")
NO_GOAL_ASSIGNED_TEXT = "No Goal Assigned for this Metric"
# why: hide only roster/schedule metadata columns; keep student id visible for coordinators
HIDDEN_PROVISION_HEADERS = (
    "School",
    "Grading Period / Assessment Date",
    "Grading Period",
    "School Year",
    "Organization",
)
ATTENDANCE_SHEET = "Attendance Rate_%"
CORE_COURSE_GRADES_SHEET = "Core Course Grades"
CORE_COURSE_HEADER = "Core Course"
STANDARDIZED_TEST_SCORE_SHEET = "Standardized Test Score"
STAND_TEST_TYPE_HEADER = "Stand. Test Type"
SEL_SHEET = "SEL"
SEL_METRIC_HEADER = "Social Emotional Learning Metric"
CREDITS_CTE_SHEET = "Credits - CTE Needed"
# why: data-validation options are under Field Options column "Credits"; row-1 header on tab is "Metric "
CREDITS_METRIC_HEADER = "Metric "
FIELD_OPTIONS_SEL_METRIC_KEY = "Social Emotional Learning Metric "
FIELD_OPTIONS_CREDITS_KEY = "Credits"
DEADLINES_SHEET = "Sheet4"
FIELD_OPTIONS_SHEET = "Field Options"
ORGANIZATION_CANONICAL = "CIS of Eastern Pennsylvania"
DEFAULT_CISIPHYUS_DIR = PROJECT_ROOT / "tools" / "cisiphyus"
DEFAULT_CISIPHYUS_LATEST_DIR = DEFAULT_CISIPHYUS_DIR / "artifacts" / "latest"
RECOVERED_REPORT_TIER_I = "tier_i_supports"
RECOVERED_REPORT_TIER_II_III = "tier_ii_iii_supports"
RECOVERED_REPORT_BASIC_NEEDS_STUDENT = "basic_needs_student_summary"
RECOVERED_REPORT_BASIC_NEEDS_SCHOOL = "basic_needs_school_summary"
# keep backward-compatible alias for existing code paths
RECOVERED_REPORT_BASIC_NEEDS = RECOVERED_REPORT_BASIC_NEEDS_STUDENT
RECOVERED_REPORT_SCHOOL_NEEDS = "school_needs_summary_export"
RECOVERED_REPORT_SCHOOL_DEMOGRAPHICS = "school_needs_assessment_school_demographics_export"
RECOVERED_REPORT_SCHOOL_IMPROVEMENT = "school_needs_assessment_improvement_plan_community_data_export"
RECOVERED_REPORT_STUDENT_SUPPORT_DETAIL = "student_support_detail"
RECOVERED_REPORT_SCHOOL_GOALS_PROGRESS = "school_goals_progress_export"
RECOVERED_REPORT_PARENT_GUARDIAN_CONSENT = "parent_guardian_consent"
RECOVERED_REPORT_SUPPORT_SUMMARY_BY_STUDENT = "support_summary_by_student"
RECOVERED_REPORT_GOAL_TRACKING_STUDENT_GOALS = "goal_tracking_student_goals"
# WHY: template typo must match sheet header; skip copy so submission cannot wipe formulas
TEMPLATE_HEADERS_PRESERVE_FORMULAS = frozenset()
REPORT_FIELDS = (
    "sheet",
    "row",
    "column",
    "code",
    "severity",
    "original_value",
    "repaired_value",
    "message",
)
STRUCTURAL_ERROR = "error"
UAT_NO_ENTITY_ID_ROWS = "UAT_NO_ENTITY_ID_ROWS"
# why: provision must not list students who left before the grading period assessment window
NO_STUDENTS_AFTER_ENROLLMENT_FILTER = "NO_STUDENTS_AFTER_ENROLLMENT_FILTER"
# why: batch targets current site coordinators; single-coordinator provision must not bypass staffing
INACTIVE_SITE_COORDINATOR_PROVISION = "INACTIVE_SITE_COORDINATOR_PROVISION"
GOOGLE_OAUTH_SCOPE_DRIVE = "https://www.googleapis.com/auth/drive"
GOOGLE_OAUTH_SCOPE_GMAIL_SEND = "https://www.googleapis.com/auth/gmail.send"
GOOGLE_OAUTH_SCOPE_ADMIN_DIRECTORY_USER_READONLY = (
    "https://www.googleapis.com/auth/admin.directory.user.readonly"
)
DEFAULT_QPR_DELEGATED_USER_EMAIL = "DataTeam@ciseasternpa.org"
DEFAULT_QPR_SERVICE_ACCOUNT_KEY_PATH = DEFAULT_QPR_DIR / ".service-account-key.json"
GOOGLE_OAUTH_SCOPE_SCRIPT_PROJECTS = "https://www.googleapis.com/auth/script.projects"
AUTOMATION_LEDGER_FILE_NAME = ".qpr_automation_ledger.json"
AUTOMATION_LEDGER_MIME_TYPE = "application/json"
AUTOMATION_COMPLETED_STATUSES = frozenset({"completed", "uploaded_not_notified"})
AUTOMATION_IN_PROGRESS_STATUSES = frozenset({"generated", "uploaded_not_notified"})
AUTOMATION_DEFAULT_PERIODS = ("4.0", "3.0", "2.0", "1.0")
SITE_COORDINATOR_EMAIL_HEADER = SITE_LIST_SITE_COORDINATOR_HEADER
SITE_DATA_POC_HEADER = "Data POC"
EMAIL_PATTERN = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
SYNTHETIC_COORDINATOR_EMAIL_DOMAIN = "ciseasternpa.org"
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
BRIDGE_SOURCE_NON_GOAL = "non-goal"
BRIDGE_SOURCE_PROGRESS = "progress"
BRIDGE_SHEET_EXCLUDE = frozenset({FIELD_OPTIONS_SHEET})
BRIDGE_HEADER_ALIASES: dict[str, str] = {
    "grading period": "Grading Period",
    "progress agynst goal": "Progress Against Goal",
}
REFERRAL_COLLECTION_HEADERS = (
    "site_name",
    "school",
    "grading_period",
    "total_referrals",
    "consents_received",
    "non_responsive_families",
    "denied_consent",
    "rejected_by_site_coordinator",
    "rejected_by_school",
    "referrals_from_last_year",
)
REFERRAL_COLLECTION_INT_FIELDS = (
    "total_referrals",
    "consents_received",
    "non_responsive_families",
    "denied_consent",
    "rejected_by_site_coordinator",
    "rejected_by_school",
    "referrals_from_last_year",
)
QPR_EMAIL_DRAFT_REQUIRED_HEADERS = (
    "recipient_email",
    "owner_tags",
    "sites",
    "coordinator_names",
)
QPR_LOCATION_REQUIRED_HEADERS = ("site_name", "qpr_url")
QPR_LEDGER_EMAIL_SUBJECT_TEMPLATE = "QPR for {site_name}"
QPR_LEDGER_EMAIL_BODY_TEMPLATE = "\n".join(
    (
        "Hi {recipient_name},",
        "",
        "The QPR for {site_name} is located here:",
        "{qpr_url}",
        "",
        "Please complete this QPR.",
        "",
        "Thank you,",
        "CISEPA Data Team",
    )
)
PACK_DENOMINATOR_POLICY = {
    "tier_participants": "tier_i_uses_export_aggregate_max_selected_or_served; tier_ii_iii_use_unique_student_id_from_student_support_detail_when_available_else_fallback_max_proxy",
    "case_managed_goal_denominator": "unique_students_with_domain_target_in_student_metrics_summary",
    "case_managed_progress_numerator": "unique_students_with_non_empty_period_progress_for_domain_metric",
    "planned_vs_delivered_coverage": "coverage_equals_delivered_over_planned_with_100_percent_when_planned_is_zero_and_delivered_is_positive",
}
QPR_BASIC_NEED_ROWS = (
    "Meals/Weekend Food Bags",
    "Snacks",
    "Hygiene Products",
    "School Supplies",
    "Uniforms/Clothing",
    "Transportation/Vouchers",
    "CIS on the Go",
    "Housing Referral",
    "Dental or Vision Referral/ Resources",
    "Professional MH Referrals",
    "Holiday Gifts",
    "Other:",
)
QPR_REFERRAL_LABEL_TO_FIELD = {
    "# of referrals made to cis this quarter": "total_referrals",
    "# of referrals from last year (q1 only)": "referrals_from_last_year",
    "# of families consenting to services": "consents_received",
    "# of families that did not respond": "non_responsive_families",
    "# of families declining services": "denied_consent",
    "# referrals rejected by cis": "rejected_by_site_coordinator",
    "# referrals rejected by school": "rejected_by_school",
}
PAPER_PREPOP_REQUIRED_FIELDS: dict[str, tuple[str, ...]] = {
    "tier_i_supports": ("support_name",),
    "tier_ii_iii_supports": ("support_name", "tier"),
    "basic_needs": ("need_category",),
    "school_needs_summary_export": ("school",),
    "school_needs_assessment_school_demographics_export": ("school",),
    "school_needs_assessment_improvement_plan_community_data_export": ("school",),
    "school_goals_progress_export": ("school", "goal_metric"),
    "parent_guardian_consent": ("school", "student_id"),
    "goal_tracking_student_goals": ("school", "student_id", "metric"),
}
PAPER_PREPOP_FIELD_ALIASES: dict[str, dict[str, tuple[str, ...]]] = {
    "tier_i_supports": {
        "school": ("school", "school name"),
        "grading_period": ("grading period", "quarter", "reporting period"),
        "support_name": ("program", "support", "support type", "support provided"),
        "frequency": ("frequency", "frequency of support", "dosage"),
        "participants_count": (
            "# of participants",
            "participants",
            "participant count",
            "# participants",
            "student count",
        ),
        "outcome": ("outcome", "progress/outcomes", "progress outcomes", "notes"),
        "planned_support": (
            "planned support",
            "school support plan supports",
            "student support plan supports",
        ),
    },
    "tier_ii_iii_supports": {
        "school": ("school", "school name"),
        "grading_period": ("grading period", "quarter", "reporting period"),
        "tier": ("tier", "support tier"),
        "goal_domain": ("goal domain", "goal area", "goal category", "metric domain"),
        "support_name": ("program", "support", "support type", "support provided"),
        "frequency": ("frequency", "frequency of support", "dosage"),
        "student_id": ("student id", "client id", "entity id"),
        "participants_count": (
            "# of participants",
            "participants",
            "participant count",
            "# participants",
            "student count",
        ),
        "outcome": ("outcome", "progress/outcomes", "progress outcomes", "notes"),
        "planned_support": (
            "planned support",
            "school support plan supports",
            "student support plan supports",
        ),
    },
    "basic_needs": {
        "school": ("school", "school name"),
        "grading_period": ("grading period", "quarter", "reporting period"),
        "need_category": ("basic need", "need category", "type", "type of resource provided", "item provided"),
        "recipient_id": ("student id", "client id", "family id", "household id"),
        "count": ("count", "# served", "quantity", "number of sessions", "number of units"),
        "details": (
            "description",
            "brief description",
            "referral destination",
            "donation source",
            "notes",
            "activity description",
        ),
    },
    "school_needs_summary_export": {
        "school": ("school",),
        "need": ("needs", "need", "priority"),
        "description": ("description", "summary"),
    },
    "school_needs_assessment_school_demographics_export": {
        "school": ("school",),
        "total enrollment": ("total enrollment",),
        "average daily attendance (%)": ("average daily attendance (%)",),
        "% of students chronically absent": ("% of students chronically absent",),
    },
    "school_needs_assessment_improvement_plan_community_data_export": {
        "school": ("school",),
        "key priorities": ("key priorities",),
        "summary of community data": ("summary of community data",),
    },
    "school_goals_progress_export": {
        "school": ("school", "current school"),
        "grading_period": ("grading period", "quarter", "reporting period"),
        "goal_metric": ("goal metric", "metric", "goal"),
        "data_source": ("data source", "source"),
        "baseline": ("baseline", "baseline value"),
        "target": ("target", "target value"),
        "q1": ("q1", "quarter 1", "first review"),
        "q2": ("q2", "quarter 2", "second review"),
        "q3": ("q3", "quarter 3", "third review"),
        "eoy": ("eoy", "quarter 4", "q4", "fourth review"),
        "progress_narrative": ("progress narrative", "story behind the data", "progress", "latest progress"),
        "action_steps": ("action steps",),
    },
    "parent_guardian_consent": {
        "school": ("most recent enrollment", "school", "current school"),
        "student_id": ("student id", "client id", "system id"),
        "most_recent_consent": ("most recent consent?",),
        "enrollment_status": ("enrollment status",),
    },
    "goal_tracking_student_goals": {
        "school": ("school", "current school"),
        "student_id": ("student id", "client id", "system id"),
        "goal": ("goal",),
        "metric": ("metric",),
        "baseline": ("baseline",),
        "target": ("target",),
        "latest_progress": ("latest progress",),
        "goal_achievement": ("goal achievement",),
        "progress_notes": ("progress notes",),
        "goal_narrative": ("goal narrative",),
    },
}


def _bridge_header_key(value: object) -> str:
    text = clean_text(value)
    if text is None:
        return ""
    return normalize_text(text)


def _bridge_effective_source_headers(
    *,
    source_kind: str,
    sheet_name: str,
    headers: list[str],
) -> list[str]:
    effective = list(headers)
    if source_kind == BRIDGE_SOURCE_PROGRESS and sheet_name == CREDITS_CTE_SHEET and effective:
        if clean_text(effective[0]) == "School Year":
            effective[0] = "School"
    return effective


def _bridge_target_to_source_columns(
    *,
    source_kind: str,
    sheet_name: str,
    source_headers: list[str],
    target_headers: list[str],
) -> tuple[dict[str, int], list[str], list[dict[str, str]]]:
    effective_headers = _bridge_effective_source_headers(
        source_kind=source_kind,
        sheet_name=sheet_name,
        headers=source_headers,
    )
    source_key_to_index: dict[str, int] = {}
    for index, header in enumerate(effective_headers):
        key = _bridge_header_key(header)
        if key and key not in source_key_to_index:
            source_key_to_index[key] = index
    mappings: dict[str, int] = {}
    alias_hits: list[dict[str, str]] = []
    for target_header in target_headers:
        target_key = _bridge_header_key(target_header)
        source_index = source_key_to_index.get(target_key)
        if source_index is None:
            alias_target = BRIDGE_HEADER_ALIASES.get(target_key)
            if alias_target is not None:
                source_index = source_key_to_index.get(_bridge_header_key(alias_target))
                if source_index is not None:
                    alias_hits.append(
                        {
                            "source_kind": source_kind,
                            "sheet": sheet_name,
                            "source_header": source_headers[source_index],
                            "target_header": target_header,
                        }
                    )
        if source_index is not None:
            mappings[target_header] = source_index
    mapped_indexes = set(mappings.values())
    unmapped_source_columns = [
        source_headers[index]
        for index in range(len(source_headers))
        if index not in mapped_indexes and clean_text(source_headers[index]) is not None
    ]
    return mappings, unmapped_source_columns, alias_hits


def _bridge_row_dedupe_key(sheet_name: str, row_values: list[object], target_headers: list[str]) -> tuple[str, ...] | None:
    header_index = {header: idx for idx, header in enumerate(target_headers)}
    student_idx = header_index.get("Student ID")
    if student_idx is None or student_idx >= len(row_values):
        return None
    student_key = student_id_join_key(row_values[student_idx])
    if student_key is None:
        return None
    key: list[str] = [sheet_name, student_key]
    for header in (
        CORE_COURSE_HEADER,
        STAND_TEST_TYPE_HEADER,
        SEL_METRIC_HEADER,
        CREDITS_METRIC_HEADER,
        "Grading Period",
    ):
        idx = header_index.get(header)
        if idx is None or idx >= len(row_values):
            continue
        value = clean_text(row_values[idx]) or ""
        key.append(f"{header}:{value}")
    return tuple(key)


def _bridge_row_score(row_values: list[object]) -> int:
    return sum(1 for value in row_values if clean_text(value) is not None)


def _bridge_template_student_lookup(
    template_workbook,
    template_schema: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    lookup: dict[str, dict[str, Any]] = {}
    for sheet_name in template_schema["sheet_order"]:
        if sheet_name in BRIDGE_SHEET_EXCLUDE:
            continue
        ws = template_workbook[sheet_name]
        headers = template_schema["sheets"][sheet_name]["headers"]
        header_index = {header: idx for idx, header in enumerate(headers)}
        student_idx = header_index.get("Student ID")
        first_idx = header_index.get("First Name")
        last_idx = header_index.get("Last Name")
        if student_idx is None:
            lookup[sheet_name] = {"ids": set(), "name_to_ids": {}}
            continue
        ids: set[str] = set()
        name_to_ids: dict[str, set[str]] = {}
        id_to_profile: dict[str, dict[str, str]] = {}
        school_idx = header_index.get("School")
        school_year_idx = header_index.get("School Year")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if student_idx >= len(row):
                continue
            sid = student_id_join_key(row[student_idx])
            if sid is None:
                continue
            ids.add(sid)
            first_name = clean_text(row[first_idx]) if first_idx is not None and first_idx < len(row) else None
            last_name = clean_text(row[last_idx]) if last_idx is not None and last_idx < len(row) else None
            school_name = clean_text(row[school_idx]) if school_idx is not None and school_idx < len(row) else None
            school_year = clean_text(row[school_year_idx]) if school_year_idx is not None and school_year_idx < len(row) else None
            id_to_profile[sid] = {
                "first_name": first_name or "",
                "last_name": last_name or "",
                "school": school_name or "",
                "school_year": school_year or "",
            }
            if first_name is None or last_name is None:
                continue
            key = f"{normalize_text(first_name)}|{normalize_text(last_name)}"
            name_to_ids.setdefault(key, set()).add(sid)
        lookup[sheet_name] = {"ids": ids, "name_to_ids": name_to_ids, "id_to_profile": id_to_profile}
    return lookup


def cell_int_matches(value: object, expected: int) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return int(value) == expected
    text = clean_text(value)
    if text is None:
        return False
    try:
        return int(float(text)) == expected
    except ValueError:
        return False


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


def grading_period_cell_is_preferred(value: object, period: str) -> bool:
    if isinstance(value, bool):
        return False
    expected = grading_period_cell_value(period)
    if isinstance(value, (int, float)):
        return int(value) == expected
    return False


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


def period_code(value: str) -> str:
    return _PERIOD_CODE_BY_CANONICAL.get(value, "MP?")


def grading_period_from_assessment_date(
    schedule: dict[str, Any],
    raw_value: object,
) -> str | None:
    iso_value = to_iso_date(raw_value)
    if iso_value is None:
        return None
    for grading_period, assessment_date in schedule["assessment_dates"].items():
        if assessment_date == iso_value:
            return grading_period
    return None


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


def _extract_emails_from_text(raw_value: object) -> list[str]:
    text = clean_text(raw_value)
    if text is None:
        return []
    seen: set[str] = set()
    ordered: list[str] = []
    for match in EMAIL_PATTERN.findall(text):
        lowered = match.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        ordered.append(lowered)
    return ordered


def _letters_only(value: str) -> str:
    return "".join(ch for ch in value.lower() if ch.isalpha())


def synthesize_org_email_from_name(display_name: str) -> str | None:
    raw_name = clean_text(display_name)
    if raw_name is None:
        return None
    first_initial = ""
    last_name = ""
    if "," in raw_name:
        left, _, right = raw_name.partition(",")
        last_name = _letters_only(left)
        right_tokens = [token for token in right.split() if token]
        if right_tokens:
            first_initial = _letters_only(right_tokens[0])[:1]
    else:
        tokens = [token for token in raw_name.split() if token]
        if len(tokens) >= 2:
            first_initial = _letters_only(tokens[0])[:1]
            last_name = _letters_only(tokens[-1])
    if not first_initial or not last_name:
        return None
    local = f"{last_name}{first_initial}"
    if not local.isalpha():
        return None
    return f"{local}@{SYNTHETIC_COORDINATOR_EMAIL_DOMAIN}"


def _escape_admin_directory_query_literal(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def build_admin_directory_users_list_query(display_name: str) -> str:
    tokens = [token for token in clean_text(display_name).split() if token] if clean_text(display_name) else []
    if not tokens:
        return "isSuspended=false"
    if len(tokens) == 1:
        token = _escape_admin_directory_query_literal(tokens[0])
        return f"name:'{token}' isSuspended=false"
    given = _escape_admin_directory_query_literal(tokens[0])
    family = _escape_admin_directory_query_literal(" ".join(tokens[1:]))
    return f"givenName:'{given}' familyName:'{family}' isSuspended=false"


def _env_flag_enabled(name: str) -> bool:
    return (os.getenv(name) or "").strip().lower() in {"1", "true", "yes", "on"}


def _directory_resolve_primary_emails(*, directory_service: Any, display_name: str) -> list[str]:
    query = build_admin_directory_users_list_query(display_name)
    response = (
        directory_service.users()
        .list(
            customer="my_customer",
            query=query,
            maxResults=10,
            projection="basic",
            orderBy="familyName",
        )
        .execute()
    )
    users = response.get("users") if isinstance(response, dict) else []
    found: list[str] = []
    for item in users if isinstance(users, list) else []:
        if not isinstance(item, dict):
            continue
        email = clean_text(item.get("primaryEmail"))
        if email is None:
            continue
        found.append(email.lower())
    return list(dict.fromkeys(found))


def _qpr_auth_mode() -> str:
    return (os.getenv("QPR_AUTH_MODE") or "oauth").strip().lower()


def _qpr_use_delegated_directory_auth() -> bool:
    return _qpr_auth_mode() == "delegated"


def _delegated_user_email() -> str:
    return clean_text(os.getenv("QPR_DELEGATED_USER_EMAIL")) or DEFAULT_QPR_DELEGATED_USER_EMAIL


def _service_account_key_path() -> Path:
    explicit = clean_text(os.getenv("QPR_SERVICE_ACCOUNT_KEY"))
    if explicit:
        return Path(explicit).expanduser().resolve()
    return DEFAULT_QPR_SERVICE_ACCOUNT_KEY_PATH


def _delegated_credentials(*, subject: str, scopes: list[str]) -> Any:
    key_path = _service_account_key_path()
    if not key_path.exists():
        raise RuntimeError(
            f"Service account key not found: {key_path}. "
            "Set QPR_SERVICE_ACCOUNT_KEY or place the key at the default path."
        )
    service_account_module = importlib.import_module("google.oauth2.service_account")
    creds = service_account_module.Credentials.from_service_account_file(
        str(key_path),
        scopes=scopes,
    )
    return creds.with_subject(subject)


def _qpr_credentials(*, purpose: Literal["directory", "gmail", "drive", "apps_script"]) -> Any:
    if purpose == "directory" and _qpr_use_delegated_directory_auth():
        return _delegated_credentials(
            subject=_delegated_user_email(),
            scopes=[GOOGLE_OAUTH_SCOPE_ADMIN_DIRECTORY_USER_READONLY],
        )
    required_scopes = _qpr_oauth_scopes(
        include_drive=purpose in {"drive", "gmail"},
        include_gmail_send=purpose == "gmail",
        include_admin_directory_readonly=purpose == "directory",
        include_apps_script_projects=purpose == "apps_script",
    )
    return _oauth_credentials_from_env(required_scopes=required_scopes)


def _directory_email_resolver() -> Callable[[str], list[str]]:
    credentials = _qpr_credentials(purpose="directory")
    directory_service = build("admin", "directory_v1", credentials=credentials)
    cache: dict[str, list[str]] = {}

    def resolver(display_name: str) -> list[str]:
        normalized = normalize_text(display_name)
        if normalized is None:
            return []
        if normalized in cache:
            return cache[normalized]
        resolved = _directory_resolve_primary_emails(
            directory_service=directory_service,
            display_name=display_name,
        )
        cache[normalized] = resolved
        return resolved

    return resolver


def _qpr_oauth_scopes(
    *,
    include_drive: bool = True,
    include_gmail_send: bool = True,
    include_admin_directory_readonly: bool = False,
    include_apps_script_projects: bool = False,
) -> list[str]:
    scopes: list[str] = []
    if include_drive:
        scopes.append(GOOGLE_OAUTH_SCOPE_DRIVE)
    if include_gmail_send:
        scopes.append(GOOGLE_OAUTH_SCOPE_GMAIL_SEND)
    if include_admin_directory_readonly:
        scopes.append(GOOGLE_OAUTH_SCOPE_ADMIN_DIRECTORY_USER_READONLY)
    if include_apps_script_projects:
        scopes.append(GOOGLE_OAUTH_SCOPE_SCRIPT_PROJECTS)
    return scopes


def _google_clients() -> tuple[Any, Any]:
    global _GOOGLE_CREDENTIALS_CLASS, _GOOGLE_BUILD_FN
    if _GOOGLE_CREDENTIALS_CLASS is not None and _GOOGLE_BUILD_FN is not None:
        return _GOOGLE_CREDENTIALS_CLASS, _GOOGLE_BUILD_FN
    credentials_module = importlib.import_module("google.oauth2.credentials")
    discovery_module = importlib.import_module("googleapiclient.discovery")
    _GOOGLE_CREDENTIALS_CLASS = credentials_module.Credentials
    _GOOGLE_BUILD_FN = discovery_module.build
    return _GOOGLE_CREDENTIALS_CLASS, _GOOGLE_BUILD_FN


def build(*args: Any, **kwargs: Any) -> Any:
    _, build_fn = _google_clients()
    return build_fn(*args, **kwargs)


def _oauth_credentials_from_env(*, required_scopes: list[str] | None = None) -> Any:
    credentials_class, _ = _google_clients()
    token_path = DEFAULT_QPR_DIR / ".oauth-refresh-token.json"
    if not token_path.exists():
        raise RuntimeError(f"OAuth refresh token file not found: {token_path}")
    payload = json.loads(token_path.read_text(encoding="utf-8"))
    refresh_token = clean_text(payload.get("refresh_token"))
    client_id = clean_text(payload.get("client_id"))
    client_secret = clean_text(payload.get("client_secret"))
    token_uri = clean_text(payload.get("token_uri")) or "https://oauth2.googleapis.com/token"
    payload_scopes = payload.get("scopes")
    scopes: list[str] | None = payload_scopes if isinstance(payload_scopes, list) else None
    if required_scopes:
        required = list(dict.fromkeys(required_scopes))
        if scopes is not None:
            missing = [scope for scope in required if scope not in scopes]
            if missing:
                hints: list[str] = []
                if GOOGLE_OAUTH_SCOPE_SCRIPT_PROJECTS in missing:
                    hints.append("--with-apps-script-projects")
                if GOOGLE_OAUTH_SCOPE_ADMIN_DIRECTORY_USER_READONLY in missing:
                    hints.append("--with-admin-directory-readonly")
                hint_text = f" with {' '.join(hints)}" if hints else ""
                raise RuntimeError(
                    "OAuth token scopes are missing required permissions for this command. "
                    f"Missing: {missing}. Re-run oauth bootstrap{hint_text}."
                )
        scopes = required
    if not refresh_token or not client_id or not client_secret:
        raise RuntimeError("OAuth refresh token file missing required credentials payload.")
    # why: refresh-token flows can reject explicit scope negotiation (`invalid_scope`)
    # even when the token is valid; rely on granted token scopes instead.
    return credentials_class(
        token=None,
        refresh_token=refresh_token,
        token_uri=token_uri,
        client_id=client_id,
        client_secret=client_secret,
    )


def _resolve_local_metric_workbook(
    *,
    local_inputs_dir: Path | None = None,
) -> tuple[Path | None, list[Path]]:
    explicit = clean_text(os.getenv("QPR_STUDENT_METRICS_WORKBOOK"))
    if explicit:
        candidate = Path(explicit).expanduser().resolve()
        return (candidate if candidate.exists() else None), [candidate]
    filename = clean_text(os.getenv("QPR_STUDENT_METRICS_FILENAME")) or "SY25-26_StudentMetricsSummary.xlsx"
    root = (
        local_inputs_dir
        or Path(clean_text(os.getenv("QPR_LOCAL_INPUTS_DIR")) or str(DEFAULT_QPR_LOCAL_INPUTS_DIR))
    )
    root = Path(root).expanduser().resolve()
    candidates: list[Path] = []
    if root.exists():
        for path in sorted(root.glob("*.xlsx")):
            name = path.name
            if name.startswith("~$"):
                continue
            if "studentmetricssummary" not in normalize_text(name).replace(" ", ""):
                continue
            candidates.append(path.resolve())
    preferred = (root / filename).resolve()
    if preferred.exists() and preferred not in candidates:
        candidates.insert(0, preferred)
    chosen = candidates[0] if candidates else (preferred if preferred.exists() else None)
    return chosen, candidates


def resolve_student_metrics_workbook(
    *,
    local_inputs_dir: Path | None = None,
) -> tuple[Path | None, list[Path], dict[str, Any]]:
    from qpr_cisiphyus_prefetch import metric_workbook_freshness, preferred_student_metrics_destination

    workbook_path, candidates = _resolve_local_metric_workbook(local_inputs_dir=local_inputs_dir)
    freshness = metric_workbook_freshness(workbook_path)
    freshness["destination_path"] = str(
        preferred_student_metrics_destination(local_inputs_dir=local_inputs_dir)
    )
    return workbook_path, candidates, freshness


def _pipeline_metric_workbook_path(cli_path: Path | None) -> Path:
    if cli_path is not None:
        return cli_path.expanduser().resolve()
    workbook_path, _, _ = resolve_student_metrics_workbook()
    if workbook_path is not None:
        return workbook_path
    from qpr_cisiphyus_prefetch import preferred_student_metrics_destination

    return preferred_student_metrics_destination()


def _load_email_map(email_map_path: Path | None) -> dict[str, str]:
    if email_map_path is None or not email_map_path.exists():
        return {}
    payload = json.loads(email_map_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        return {}
    out: dict[str, str] = {}
    for key, value in payload.items():
        k = clean_text(key)
        v = clean_text(value)
        if k is None or v is None:
            continue
        out[normalize_text(k)] = v.lower()
    return out


def _load_synthetic_override_map(override_map_path: Path | None) -> dict[str, str]:
    if override_map_path is None or not override_map_path.exists():
        return {}
    payload = json.loads(override_map_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        return {}
    out: dict[str, str] = {}
    for key, value in payload.items():
        key_text = clean_text(key)
        value_text = clean_text(value)
        if key_text is None or value_text is None or "|" not in key_text:
            continue
        base_raw, _, name_raw = key_text.partition("|")
        base = clean_text(base_raw)
        normalized_name = normalize_text(name_raw)
        if base is None or not normalized_name:
            continue
        out[f"{base.lower()}|{normalized_name}"] = value_text.lower()
    return out


def resolve_synthetic_emails(
    names: list[str],
    overrides: dict[str, str],
) -> tuple[dict[str, str], list[dict[str, str]]]:
    grouped_by_generated: dict[str, dict[str, str]] = {}
    unresolved: list[dict[str, str]] = []
    for name in names:
        coordinator_name = clean_text(name)
        if coordinator_name is None:
            unresolved.append(
                {
                    "coordinator_name": "",
                    "generated_email": "",
                    "reason": "missing_name",
                    "collision_group": "",
                }
            )
            continue
        generated_email = synthesize_org_email_from_name(coordinator_name)
        if generated_email is None:
            unresolved.append(
                {
                    "coordinator_name": coordinator_name,
                    "generated_email": "",
                    "reason": "invalid_pattern_input",
                    "collision_group": "",
                }
            )
            continue
        grouped_by_generated.setdefault(generated_email, {})
        grouped_by_generated[generated_email].setdefault(normalize_text(coordinator_name), coordinator_name)
    resolved: dict[str, str] = {}
    for generated_email, grouped_names in grouped_by_generated.items():
        collision = len(grouped_names) > 1
        for _, coordinator_name in grouped_names.items():
            override_key = f"{generated_email.lower()}|{normalize_text(coordinator_name)}"
            override_email = overrides.get(override_key)
            if override_email is not None:
                override_candidates = _extract_emails_from_text(override_email)
                if not override_candidates:
                    unresolved.append(
                        {
                            "coordinator_name": coordinator_name,
                            "generated_email": generated_email,
                            "reason": "invalid_override_email",
                            "collision_group": generated_email if collision else "",
                        }
                    )
                    continue
                resolved[coordinator_name] = override_candidates[0].lower()
                continue
            if collision:
                unresolved.append(
                    {
                        "coordinator_name": coordinator_name,
                        "generated_email": generated_email,
                        "reason": "collision_needs_override",
                        "collision_group": generated_email,
                    }
                )
                continue
            resolved[coordinator_name] = generated_email
    return resolved, unresolved


def load_site_coordinator_recipient_emails(
    site_staff_list_path: Path,
    *,
    email_map_path: Path | None = None,
    synthetic_override_map_json_path: Path | None = None,
    directory_resolver: Callable[[str], list[str]] | None = None,
    directory_resolution_warnings: list[dict[str, Any]] | None = None,
    recipient_resolution_metadata: list[dict[str, Any]] | None = None,
) -> dict[str, list[str]]:
    workbook = load_workbook(site_staff_list_path, read_only=True, data_only=True)
    if SITE_LIST_SHEET not in workbook.sheetnames:
        workbook.close()
        return {}
    ws = workbook[SITE_LIST_SHEET]
    headers = header_values(ws)
    lookup = _header_lookup(headers)
    school_idx = lookup.get("School")
    coordinator_idx = lookup.get(SITE_COORDINATOR_EMAIL_HEADER)
    end_idx = lookup.get(SITE_LIST_END_DATE_HEADER)
    poc_idx = lookup.get(SITE_DATA_POC_HEADER)
    email_map = _load_email_map(email_map_path)
    synthetic_overrides = _load_synthetic_override_map(synthetic_override_map_json_path)
    recipients: dict[str, list[str]] = {}
    pending_synthetic: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        school = clean_text(row[school_idx]) if school_idx is not None and school_idx < len(row) else None
        if school is None:
            continue
        end_value = row[end_idx] if end_idx is not None and end_idx < len(row) else None
        if not _site_list_end_date_means_active_assignment(end_value):
            continue
        coordinator_raw = row[coordinator_idx] if coordinator_idx is not None and coordinator_idx < len(row) else None
        poc_raw = row[poc_idx] if poc_idx is not None and poc_idx < len(row) else None
        data_poc_emails = _extract_emails_from_text(poc_raw)
        found: list[str] = []
        queued_for_synthetic = False
        resolution_source: str | None = None
        reason: str | None = None
        generated_email: str | None = None
        collision_group: str | None = None
        coordinator_emails = _extract_emails_from_text(coordinator_raw)
        if coordinator_emails:
            found.extend(coordinator_emails)
            resolution_source = "inline"
        else:
            coordinator_name = clean_text(coordinator_raw)
            if coordinator_name is not None:
                mapped = email_map.get(normalize_text(coordinator_name))
                if mapped:
                    found.extend(_extract_emails_from_text(mapped))
                    resolution_source = "map"
                elif directory_resolver is not None:
                    resolved = [addr.lower() for addr in directory_resolver(coordinator_name) if clean_text(addr)]
                    resolved = list(dict.fromkeys(resolved))
                    if len(resolved) == 1:
                        found.extend(resolved)
                        resolution_source = "directory"
                    else:
                        reason = "directory_not_found" if not resolved else "directory_ambiguous"
            else:
                reason = "missing_name"
        found.extend(data_poc_emails)
        if found and resolution_source is None:
            resolution_source = "data_poc"
        deduped = list(dict.fromkeys(found))
        coordinator_name = clean_text(coordinator_raw)
        if not deduped and coordinator_name is not None and not coordinator_emails:
            pending_synthetic.append(
                {
                    "school": school,
                    "coordinator_name": coordinator_name,
                    "reason": reason,
                }
            )
            queued_for_synthetic = True
        if deduped:
            recipients.setdefault(school, [])
            recipients[school] = list(dict.fromkeys(recipients[school] + deduped))
        if recipient_resolution_metadata is not None and not queued_for_synthetic:
            recipient_resolution_metadata.append(
                {
                    "site_name": school,
                    "coordinator_name": clean_text(coordinator_raw),
                    "generated_email": generated_email,
                    "resolution_source": resolution_source or ("unresolved" if not deduped else "data_poc"),
                    "reason": reason if not deduped else None,
                    "collision_group": collision_group,
                }
            )
    if pending_synthetic:
        names = [item["coordinator_name"] for item in pending_synthetic]
        synthetic_resolved, synthetic_unresolved = resolve_synthetic_emails(names, synthetic_overrides)
        unresolved_by_name: dict[str, dict[str, str]] = {
            item["coordinator_name"]: item for item in synthetic_unresolved if item.get("coordinator_name")
        }
        for pending in pending_synthetic:
            school = pending["school"]
            coordinator_name = pending["coordinator_name"]
            resolved_email = synthetic_resolved.get(coordinator_name)
            unresolved_row = unresolved_by_name.get(coordinator_name)
            if resolved_email is not None:
                recipients.setdefault(school, [])
                if resolved_email not in recipients[school]:
                    recipients[school].append(resolved_email)
                if recipient_resolution_metadata is not None:
                    source = "override" if resolved_email != synthesize_org_email_from_name(coordinator_name) else "synthetic"
                    recipient_resolution_metadata.append(
                        {
                            "site_name": school,
                            "coordinator_name": coordinator_name,
                            "generated_email": resolved_email,
                            "resolution_source": source,
                            "reason": None,
                            "collision_group": None,
                        }
                    )
                continue
            reason = pending.get("reason")
            collision_group = None
            generated_email = synthesize_org_email_from_name(coordinator_name)
            if unresolved_row is not None:
                reason = unresolved_row.get("reason") or reason
                generated_email = unresolved_row.get("generated_email") or generated_email
                collision_group = unresolved_row.get("collision_group")
            if reason is None:
                reason = "unresolved"
            if reason in {"directory_not_found", "directory_ambiguous"} and directory_resolution_warnings is not None:
                directory_resolution_warnings.append(
                    {"school": school, "coordinator": coordinator_name, "reason": reason}
                )
            if recipient_resolution_metadata is not None:
                recipient_resolution_metadata.append(
                    {
                        "site_name": school,
                        "coordinator_name": coordinator_name,
                        "generated_email": generated_email,
                        "resolution_source": "unresolved",
                        "reason": reason,
                        "collision_group": collision_group,
                    }
                )
    workbook.close()
    return recipients


def _load_active_site_contacts(site_staff_list_path: Path) -> dict[str, dict[str, list[str]]]:
    workbook = load_workbook(site_staff_list_path, read_only=True, data_only=True)
    if SITE_LIST_SHEET not in workbook.sheetnames:
        workbook.close()
        return {}
    worksheet = workbook[SITE_LIST_SHEET]
    headers = header_values(worksheet)
    lookup = _header_lookup(headers)
    school_idx = lookup.get("School")
    coordinator_idx = lookup.get(SITE_LIST_SITE_COORDINATOR_HEADER)
    poc_idx = lookup.get(SITE_DATA_POC_HEADER)
    end_idx = lookup.get(SITE_LIST_END_DATE_HEADER)
    rows_by_school: dict[str, dict[str, set[str]]] = {}
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        school = clean_text(row[school_idx]) if school_idx is not None and school_idx < len(row) else None
        if school is None:
            continue
        end_value = row[end_idx] if end_idx is not None and end_idx < len(row) else None
        if not _site_list_end_date_means_active_assignment(end_value):
            continue
        bucket = rows_by_school.setdefault(school, {"site_coordinators": set(), "data_pocs": set()})
        coordinator = clean_text(row[coordinator_idx]) if coordinator_idx is not None and coordinator_idx < len(row) else None
        if coordinator is not None:
            bucket["site_coordinators"].add(coordinator)
        data_poc = clean_text(row[poc_idx]) if poc_idx is not None and poc_idx < len(row) else None
        if data_poc is not None:
            bucket["data_pocs"].add(data_poc)
    workbook.close()
    out: dict[str, dict[str, list[str]]] = {}
    for school, values in rows_by_school.items():
        out[school] = {
            "site_coordinators": sorted(values["site_coordinators"], key=lambda value: normalize_text(value)),
            "data_pocs": sorted(values["data_pocs"], key=lambda value: normalize_text(value)),
        }
    return out


def _write_csv_rows(path: Path, *, fieldnames: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def pipeline_recipient_audit(
    *,
    as_of: date | None = None,
    site_staff_list_path: Path = DEFAULT_SITE_STAFF_LIST_PATH,
    email_map_path: Path | None = None,
    synthetic_override_map_json_path: Path | None = None,
    output_directory: Path | None = None,
    directory_email_lookup: bool = False,
    directory_resolver: Callable[[str], list[str]] | None = None,
) -> dict[str, Any]:
    current_date = as_of or date.today()
    active_sites = load_active_site_school_names(site_staff_list_path)
    active_contacts = _load_active_site_contacts(site_staff_list_path)
    directory_resolution_warnings: list[dict[str, Any]] = []
    recipient_resolution_metadata: list[dict[str, Any]] = []
    resolver = directory_resolver
    if resolver is None and directory_email_lookup:
        resolver = _directory_email_resolver()
    recipients_by_site = load_site_coordinator_recipient_emails(
        site_staff_list_path,
        email_map_path=email_map_path,
        synthetic_override_map_json_path=synthetic_override_map_json_path,
        directory_resolver=resolver,
        directory_resolution_warnings=directory_resolution_warnings if directory_email_lookup else None,
        recipient_resolution_metadata=recipient_resolution_metadata,
    )
    sorted_sites = sorted(active_sites, key=lambda value: normalize_text(value))
    recipient_rows: list[dict[str, Any]] = []
    for school in sorted_sites:
        contacts = active_contacts.get(school, {})
        recipients = recipients_by_site.get(school, [])
        recipient_rows.append(
            {
                "site_name": school,
                "site_coordinators": "; ".join(contacts.get("site_coordinators", [])),
                "data_pocs": "; ".join(contacts.get("data_pocs", [])),
                "recipient_emails": "; ".join(recipients),
                "recipient_count": len(recipients),
            }
        )
    gaps: list[dict[str, str]] = []
    seen_gap_keys: set[tuple[str, str, str]] = set()
    specific_gap_sites: set[str] = set()
    for row in recipient_resolution_metadata:
        if clean_text(row.get("resolution_source")) != "unresolved":
            continue
        school = clean_text(row.get("site_name")) or ""
        coordinator = clean_text(row.get("coordinator_name")) or ""
        reason = clean_text(row.get("reason")) or "unresolved"
        key = (school, coordinator, reason)
        if key in seen_gap_keys:
            continue
        seen_gap_keys.add(key)
        gaps.append({"site_name": school, "coordinator": coordinator, "reason": reason})
        if school:
            specific_gap_sites.add(school)
    for warning in directory_resolution_warnings:
        school = clean_text(warning.get("school")) or ""
        coordinator = clean_text(warning.get("coordinator")) or ""
        reason = clean_text(warning.get("reason")) or "missing"
        key = (school, coordinator, reason)
        if key in seen_gap_keys:
            continue
        seen_gap_keys.add(key)
        gaps.append({"site_name": school, "coordinator": coordinator, "reason": reason})
        if school:
            specific_gap_sites.add(school)
    for school in sorted_sites:
        if recipients_by_site.get(school):
            continue
        if school in specific_gap_sites:
            continue
        contacts = active_contacts.get(school, {})
        coordinator = "; ".join(contacts.get("site_coordinators", []))
        key = (school, coordinator, "missing")
        if key in seen_gap_keys:
            continue
        seen_gap_keys.add(key)
        gaps.append({"site_name": school, "coordinator": coordinator, "reason": "missing"})
    gaps.sort(key=lambda row: (normalize_text(row.get("site_name")), normalize_text(row.get("coordinator"))))
    output_root = (output_directory or DEFAULT_WRITTEN_QPR_OUTPUT_DIR).resolve()
    stamp = current_date.isoformat()
    recipients_path = output_root / f"qpr_recipients_{stamp}.csv"
    gaps_path = output_root / f"qpr_recipient_gaps_{stamp}.csv"
    summary_path = output_root / f"qpr_recipients_summary_{stamp}.json"
    ledger_path = output_root / f"qpr_recipients_ledger_{stamp}.json"
    _write_csv_rows(
        recipients_path,
        fieldnames=["site_name", "site_coordinators", "data_pocs", "recipient_emails", "recipient_count"],
        rows=recipient_rows,
    )
    _write_csv_rows(
        gaps_path,
        fieldnames=["site_name", "coordinator", "reason"],
        rows=gaps,
    )
    unique_emails = sorted(
        {
            email
            for recipients in recipients_by_site.values()
            for email in recipients
            if clean_text(email) is not None
        }
    )
    ledger_timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    ledger_rows: list[dict[str, Any]] = []
    for row in recipient_resolution_metadata:
        collision_group = clean_text(row.get("collision_group"))
        reason = clean_text(row.get("reason"))
        resolution_source = clean_text(row.get("resolution_source")) or "unresolved"
        ledger_rows.append(
            {
                "site_name": clean_text(row.get("site_name")),
                "coordinator_name": clean_text(row.get("coordinator_name")),
                "generated_email": clean_text(row.get("generated_email")),
                "resolution_source": resolution_source,
                "collision_group": collision_group,
                "canonical_email": None,
                "match_status": "unchecked",
                "reason": reason,
                "notes": None,
                "timestamp": ledger_timestamp,
            }
        )
    total_rows = len(ledger_rows)
    unresolved_rows = sum(1 for row in ledger_rows if row.get("resolution_source") == "unresolved")
    collision_rows = sum(1 for row in ledger_rows if clean_text(row.get("collision_group")) is not None)
    checked_rows = sum(1 for row in ledger_rows if clean_text(row.get("match_status")) not in (None, "unchecked"))
    unchecked_rows = sum(1 for row in ledger_rows if clean_text(row.get("match_status")) == "unchecked")
    match_rows = sum(1 for row in ledger_rows if clean_text(row.get("match_status")) == "match")
    mismatch_rows = sum(1 for row in ledger_rows if clean_text(row.get("match_status")) == "mismatch")
    ledger_summary = {
        "total_rows": total_rows,
        "resolved_rows": total_rows - unresolved_rows,
        "unresolved_rows": unresolved_rows,
        "collision_rows": collision_rows,
        "checked_rows": checked_rows,
        "unchecked_rows": unchecked_rows,
        "match_rows": match_rows,
        "mismatch_rows": mismatch_rows,
    }
    ledger_payload = {"summary": ledger_summary, "rows": ledger_rows}
    ledger_path.parent.mkdir(parents=True, exist_ok=True)
    ledger_path.write_text(json.dumps(ledger_payload, indent=2), encoding="utf-8")
    summary = {
        "status": "ok",
        "as_of_date": current_date.isoformat(),
        "directory_email_lookup": directory_email_lookup,
        "active_sites": len(sorted_sites),
        "sites_with_recipients": sum(1 for school in sorted_sites if recipients_by_site.get(school)),
        "sites_missing_recipients": sum(1 for school in sorted_sites if not recipients_by_site.get(school)),
        "unique_recipient_emails": len(unique_emails),
        "directory_resolution_warnings": directory_resolution_warnings,
        "ledger_summary": ledger_summary,
        "artifacts": {
            "recipients_csv": str(recipients_path),
            "gaps_csv": str(gaps_path),
            "summary_json": str(summary_path),
            "ledger_json": str(ledger_path),
        },
    }
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


def _split_semicolon_values(raw_value: object) -> list[str]:
    text = clean_text(raw_value)
    if text is None:
        return []
    return [part for part in [clean_text(chunk) for chunk in text.split(";")] if part is not None]


def _is_http_url(raw_value: object) -> bool:
    value = clean_text(raw_value)
    if value is None:
        return False
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _render_qpr_ledger_email(*, recipient_name: str, site_name: str, qpr_url: str) -> dict[str, str]:
    return {
        "email_subject": QPR_LEDGER_EMAIL_SUBJECT_TEMPLATE.format(site_name=site_name),
        "email_body": QPR_LEDGER_EMAIL_BODY_TEMPLATE.format(
            recipient_name=recipient_name,
            site_name=site_name,
            qpr_url=qpr_url,
        ),
    }


def _load_qpr_url_map(location_csv_path: Path | None) -> dict[str, str]:
    if location_csv_path is None:
        return {}
    if not location_csv_path.exists():
        raise ValueError(f"QPR location csv not found: {location_csv_path}")
    with location_csv_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        actual_headers = reader.fieldnames or []
        missing_headers = sorted(set(QPR_LOCATION_REQUIRED_HEADERS) - set(actual_headers))
        if missing_headers:
            raise ValueError(f"QPR location file missing headers: {', '.join(missing_headers)}")
        out: dict[str, str] = {}
        for row in reader:
            site_name = clean_text(row.get("site_name"))
            qpr_url = clean_text(row.get("qpr_url"))
            if site_name is None or qpr_url is None:
                continue
            out[normalize_text(site_name)] = qpr_url
    return out


def _load_existing_missing_recipients(path: Path | None) -> list[dict[str, str]]:
    if path is None:
        return []
    if not path.exists():
        raise ValueError(f"Existing missing recipient csv not found: {path}")
    with path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        rows: list[dict[str, str]] = []
        for row in reader:
            rows.append(
                {
                    "site_name": clean_text(row.get("site_name")) or "",
                    "site_coordinators": clean_text(row.get("site_coordinators")) or "",
                    "owner_tag": clean_text(row.get("owner_tag")) or "",
                    "reason": clean_text(row.get("reason")) or "missing_recipient_email",
                }
            )
    return rows


def pipeline_qpr_email_ledger(
    *,
    as_of: date | None = None,
    recipient_draft_csv_path: Path,
    output_directory: Path | None = None,
    qpr_location_csv_path: Path | None = None,
    existing_missing_sites_csv_path: Path | None = None,
) -> dict[str, Any]:
    if not recipient_draft_csv_path.exists():
        raise ValueError(f"Recipient draft csv not found: {recipient_draft_csv_path}")
    current_date = as_of or date.today()
    qpr_url_by_site = _load_qpr_url_map(qpr_location_csv_path)
    send_rows: list[dict[str, Any]] = []
    missing_qpr_url_rows: list[dict[str, str]] = []
    missing_recipient_rows: list[dict[str, str]] = _load_existing_missing_recipients(existing_missing_sites_csv_path)
    seen_send_keys: set[tuple[str, str]] = set()
    duplicate_send_rows: list[dict[str, str]] = []
    source_rows = 0
    expanded_site_rows = 0
    invalid_url_rows = 0
    with recipient_draft_csv_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        actual_headers = reader.fieldnames or []
        missing_headers = sorted(set(QPR_EMAIL_DRAFT_REQUIRED_HEADERS) - set(actual_headers))
        if missing_headers:
            raise ValueError(f"Recipient draft file missing headers: {', '.join(missing_headers)}")
        for row in reader:
            source_rows += 1
            recipient_email = clean_text(row.get("recipient_email"))
            owner_tag = clean_text(row.get("owner_tags")) or ""
            sites = _split_semicolon_values(row.get("sites"))
            coordinators = _split_semicolon_values(row.get("coordinator_names"))
            coordinator_names = "; ".join(coordinators)
            recipient_name = coordinators[0] if coordinators else "Site Coordinator"
            for site_name in sites:
                expanded_site_rows += 1
                normalized_site = normalize_text(site_name)
                qpr_url = clean_text(row.get("qpr_url")) or qpr_url_by_site.get(normalized_site)
                if recipient_email is None or EMAIL_PATTERN.fullmatch(recipient_email) is None:
                    missing_recipient_rows.append(
                        {
                            "site_name": site_name,
                            "site_coordinators": coordinator_names,
                            "owner_tag": owner_tag,
                            "reason": "missing_recipient_email",
                        }
                    )
                    continue
                if clean_text(site_name) is None:
                    continue
                send_key = (recipient_email.lower(), normalized_site)
                if send_key in seen_send_keys:
                    duplicate_send_rows.append(
                        {
                            "site_name": site_name,
                            "recipient_email": recipient_email,
                            "reason": "duplicate_site_recipient",
                        }
                    )
                    continue
                seen_send_keys.add(send_key)
                if qpr_url is None:
                    missing_qpr_url_rows.append(
                        {
                            "site_name": site_name,
                            "site_coordinators": coordinator_names,
                            "owner_tag": owner_tag,
                            "reason": "missing_qpr_url",
                        }
                    )
                    continue
                if not _is_http_url(qpr_url):
                    invalid_url_rows += 1
                    missing_qpr_url_rows.append(
                        {
                            "site_name": site_name,
                            "site_coordinators": coordinator_names,
                            "owner_tag": owner_tag,
                            "reason": "invalid_qpr_url",
                        }
                    )
                    continue
                rendered = _render_qpr_ledger_email(
                    recipient_name=recipient_name,
                    site_name=site_name,
                    qpr_url=qpr_url,
                )
                send_rows.append(
                    {
                        "recipient_email": recipient_email,
                        "recipient_name": recipient_name,
                        "site_name": site_name,
                        "qpr_url": qpr_url,
                        "email_subject": rendered["email_subject"],
                        "email_body": rendered["email_body"],
                    }
                )
    output_root = (output_directory or DEFAULT_WRITTEN_QPR_OUTPUT_DIR).resolve()
    stamp = current_date.isoformat()
    send_rows_path = output_root / f"qpr_site_email_ledger_{stamp}.csv"
    missing_qpr_url_path = output_root / f"qpr_site_email_missing_qpr_url_{stamp}.csv"
    missing_recipients_path = output_root / f"qpr_site_email_missing_recipients_{stamp}.csv"
    duplicates_path = output_root / f"qpr_site_email_duplicates_{stamp}.csv"
    checklist_path = output_root / f"qpr_site_email_checklist_{stamp}.md"
    summary_path = output_root / f"qpr_site_email_qa_{stamp}.json"
    _write_csv_rows(
        send_rows_path,
        fieldnames=["recipient_email", "recipient_name", "site_name", "qpr_url", "email_subject", "email_body"],
        rows=send_rows,
    )
    _write_csv_rows(
        missing_qpr_url_path,
        fieldnames=["site_name", "site_coordinators", "owner_tag", "reason"],
        rows=missing_qpr_url_rows,
    )
    _write_csv_rows(
        missing_recipients_path,
        fieldnames=["site_name", "site_coordinators", "owner_tag", "reason"],
        rows=missing_recipient_rows,
    )
    _write_csv_rows(
        duplicates_path,
        fieldnames=["site_name", "recipient_email", "reason"],
        rows=duplicate_send_rows,
    )
    owner_tags_present = sorted(
        {
            clean_text(row.get("owner_tag"))
            for row in [*missing_qpr_url_rows, *missing_recipient_rows]
            if clean_text(row.get("owner_tag")) is not None
        }
    )
    qa_checklist = [
        {"check": "recipient_email_present", "pass": len(missing_recipient_rows) == 0},
        {"check": "site_name_present", "pass": all(clean_text(row.get("site_name")) is not None for row in send_rows)},
        {"check": "qpr_url_present_and_http", "pass": len(missing_qpr_url_rows) == 0},
        {"check": "no_duplicate_site_recipient", "pass": len(duplicate_send_rows) == 0},
    ]
    checklist_lines = [
        "# qpr site email pre-send checklist",
        "",
        f"- [ ] confirm send rows generated: `{len(send_rows)}`",
        f"- [ ] confirm missing recipient rows reviewed: `{len(missing_recipient_rows)}`",
        f"- [ ] confirm missing qpr url rows reviewed: `{len(missing_qpr_url_rows)}`",
        f"- [ ] confirm duplicate rows reviewed: `{len(duplicate_send_rows)}`",
        "- [ ] spot-check at least one rendered email body for correct site and link",
    ]
    checklist_path.parent.mkdir(parents=True, exist_ok=True)
    checklist_path.write_text("\n".join(checklist_lines) + "\n", encoding="utf-8")
    summary = {
        "status": "ok",
        "as_of_date": current_date.isoformat(),
        "source_rows": source_rows,
        "expanded_site_rows": expanded_site_rows,
        "sendable_rows": len(send_rows),
        "missing_recipient_rows": len(missing_recipient_rows),
        "missing_qpr_url_rows": len(missing_qpr_url_rows),
        "invalid_qpr_url_rows": invalid_url_rows,
        "duplicate_site_recipient_rows": len(duplicate_send_rows),
        "owner_tags_present": owner_tags_present,
        "subject_template": QPR_LEDGER_EMAIL_SUBJECT_TEMPLATE,
        "body_template": QPR_LEDGER_EMAIL_BODY_TEMPLATE,
        "qa_checklist": qa_checklist,
        "artifacts": {
            "email_ledger_csv": str(send_rows_path),
            "missing_qpr_url_csv": str(missing_qpr_url_path),
            "missing_recipients_csv": str(missing_recipients_path),
            "duplicates_csv": str(duplicates_path),
            "checklist_md": str(checklist_path),
            "summary_json": str(summary_path),
        },
    }
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


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


def _cell_value_by_header(
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    header_name: str,
) -> object:
    column_index = header_lookup.get(header_name)
    if column_index is None:
        return None
    return worksheet.cell(row=row_index, column=column_index).value


def _cell_text_by_header(
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    header_name: str,
) -> str | None:
    return clean_text(
        _cell_value_by_header(
            worksheet,
            row_index=row_index,
            header_lookup=header_lookup,
            header_name=header_name,
        )
    )


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


def load_caseload_lookup(caseload_path: Path | None) -> dict[str, dict[str, str]]:
    opened = _open_sheet_with_headers(caseload_path)
    if opened is None:
        return {}
    workbook, worksheet, header_lookup = opened
    lookup: dict[str, dict[str, str]] = {}

    def value_for(row: tuple[object, ...], *header_names: str) -> str:
        for header_name in header_names:
            index = header_lookup.get(header_name)
            if index is None or index >= len(row):
                continue
            text = clean_text(row[index])
            if text is not None:
                return text
        return ""

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        student_id = clean_text(row[header_lookup["Student ID"]]) if "Student ID" in header_lookup else None
        if student_id is None:
            continue
        entry = {
            "student_id": student_id,
            "entity_id": value_for(row, "EntityID"),
            "first_name": value_for(row, "First Name", "Client.FirstName"),
            "last_name": value_for(row, "Last Name"),
            "organization": value_for(row, "Organization", "Organization Name"),
            "school": value_for(row, "School"),
            "created_by": value_for(row, "Case Manager Assigned", "CreatedBy"),
            "school_year": value_for(row, "School Year"),
        }
        lookup[student_id] = entry
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


def pick_core_course_from_summary_blocks(
    blocks: list[dict[str, Any]] | None,
    allowed_options: list[str],
) -> str | None:
    return _pick_template_field_from_summary_blocks(
        blocks, allowed_options, template_core_course_from_metric
    )


def pick_stand_test_type_from_summary_blocks(
    blocks: list[dict[str, Any]] | None,
    allowed_options: list[str],
) -> str | None:
    return _pick_template_field_from_summary_blocks(
        blocks, allowed_options, template_stand_test_type_from_metric
    )


def pick_sel_metric_from_summary_blocks(
    blocks: list[dict[str, Any]] | None,
    allowed_options: list[str],
) -> str | None:
    return _pick_template_field_from_summary_blocks(
        blocks, allowed_options, template_sel_metric_from_metric
    )


def pick_credits_metric_from_summary_blocks(
    blocks: list[dict[str, Any]] | None,
    allowed_options: list[str],
) -> str | None:
    return _pick_template_field_from_summary_blocks(
        blocks, allowed_options, template_credits_metric_from_metric
    )


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


def normalize_school_year_label(raw: object) -> str | None:
    text = clean_text(raw)
    if text is None:
        return None
    normalized = text.upper().replace("SY", "").replace("/", "-").replace(" ", "")
    match = re.search(r"(\d{2,4})-(\d{2,4})", normalized)
    if match is None:
        return None
    left = match.group(1)
    right = match.group(2)
    if len(left) == 2:
        left = f"20{left}"
    if len(right) == 2:
        right = f"20{right}"
    try:
        start = int(left)
        end = int(right)
    except ValueError:
        return None
    if start < 2000 or end < 2000:
        return None
    if end < start:
        return None
    return f"{start:04d}-{end:04d}"


def _site_staff_year_marker_enabled(value: object) -> bool:
    text = normalize_text(clean_text(value))
    return text in {"y", "yes", "true", "1"}


def load_site_staff_caseload_rows(
    *,
    site_staff_list_path: Path,
    school_name: str | None,
) -> list[dict[str, Any]]:
    if school_name is None or not site_staff_list_path.exists():
        return []
    workbook = load_workbook(site_staff_list_path, read_only=True, data_only=True)
    try:
        if SITE_LIST_SHEET not in workbook.sheetnames:
            return []
        worksheet = workbook[SITE_LIST_SHEET]
        headers = header_values(worksheet)
        lookup = _header_lookup(headers)
        school_idx = lookup.get("School")
        coordinator_idx = lookup.get(SITE_LIST_SITE_COORDINATOR_HEADER)
        end_date_idx = lookup.get(SITE_LIST_END_DATE_HEADER)
        if school_idx is None or coordinator_idx is None:
            return []
        year_marker_cols: dict[str, int] = {}
        for header, col_idx in lookup.items():
            if re.fullmatch(r"\d{2}-\d{2}", header.strip()):
                year_marker_cols[header.strip()] = col_idx
        accepted_names = {normalize_text(school_name)}
        if normalize_text(school_name) == "reading senior high school":
            accepted_names.add("reading high school")
        rows: list[dict[str, Any]] = []
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if row is None:
                continue
            school_value = row[school_idx] if school_idx < len(row) else None
            school_text = clean_text(school_value)
            if school_text is None or normalize_text(school_text) not in accepted_names:
                continue
            coordinator_value = row[coordinator_idx] if coordinator_idx < len(row) else None
            coordinator_text = clean_text(coordinator_value)
            if coordinator_text is None:
                continue
            end_raw = row[end_date_idx] if end_date_idx is not None and end_date_idx < len(row) else None
            markers: dict[str, bool] = {}
            for marker, col_idx in year_marker_cols.items():
                value = row[col_idx] if col_idx < len(row) else None
                markers[marker] = _site_staff_year_marker_enabled(value)
            rows.append(
                {
                    "school": school_text,
                    "site_coordinator": coordinator_text,
                    "active_assignment": _site_list_end_date_means_active_assignment(end_raw),
                    "year_markers": markers,
                }
            )
        return rows
    finally:
        workbook.close()


def _school_year_to_site_staff_marker(school_year: str) -> str:
    left, right = school_year.split("-")
    return f"{left[-2:]}-{right[-2:]}"


def count_site_coordinators_for_school_year(
    *,
    site_staff_rows: list[dict[str, Any]],
    school_year: str,
) -> tuple[int, str]:
    marker = _school_year_to_site_staff_marker(school_year)
    marker_counted: set[str] = set()
    for row in site_staff_rows:
        coordinator = normalize_text(row.get("site_coordinator"))
        if coordinator is None:
            continue
        markers = row.get("year_markers") or {}
        if markers.get(marker):
            marker_counted.add(coordinator)
    if marker_counted:
        return len(marker_counted), f"site_staff_marker:{marker}"
    active_counted: set[str] = set()
    for row in site_staff_rows:
        if not row.get("active_assignment"):
            continue
        coordinator = normalize_text(row.get("site_coordinator"))
        if coordinator is not None:
            active_counted.add(coordinator)
    return len(active_counted), "active_assignment_fallback"


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


def discover_coordinators(
    *,
    roster_lookup: dict[str, dict[str, str]],
    active_site_coordinator_names: list[str] | None = None,
) -> list[str]:
    # why: one workbook per Case Manager from metrics; optional restrict to Site List (active assignments only)
    names = _case_manager_names_from_roster(roster_lookup)
    if not active_site_coordinator_names:
        return names
    return [
        name
        for name in names
        if roster_coordinator_matches_active_site_staff(name, active_site_coordinator_names)
    ]


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


def filename_safe_label(value: str) -> str:
    pieces: list[str] = []
    current: list[str] = []
    for character in value:
        if character.isalnum():
            current.append(character)
            continue
        if current:
            pieces.append("".join(current))
            current = []
    if current:
        pieces.append("".join(current))
    return "_".join(pieces) or "coordinator"


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


def parse_int(value: object, *, default: int = 0) -> int:
    if isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value:
            return default
        return int(value)
    text = clean_text(value)
    if text is None:
        return default
    try:
        return int(float(text.replace(",", "")))
    except ValueError:
        return default


def _normalized_headers(headers: list[str]) -> dict[str, int]:
    lookup: dict[str, int] = {}
    for index, header in enumerate(headers):
        key = normalize_text(header)
        if key and key not in lookup:
            lookup[key] = index
    return lookup


def _resolve_alias_index(
    normalized_lookup: dict[str, int],
    aliases: tuple[str, ...],
) -> int | None:
    for alias in aliases:
        found = normalized_lookup.get(normalize_text(alias))
        if found is not None:
            return found
    return None


def _row_is_empty(row: tuple[object, ...]) -> bool:
    return all(clean_text(value) is None for value in row)


def _normalize_basic_need_category(raw: object) -> str:
    text = clean_text(raw)
    if text is None:
        return "Other:"
    normalized = normalize_text(text)
    aliases = {
        "meal": "Meals/Weekend Food Bags",
        "food bag": "Meals/Weekend Food Bags",
        "weekend": "Meals/Weekend Food Bags",
        "snack": "Snacks",
        "hygiene": "Hygiene Products",
        "school supplies": "School Supplies",
        "uniform": "Uniforms/Clothing",
        "clothing": "Uniforms/Clothing",
        "transportation": "Transportation/Vouchers",
        "voucher": "Transportation/Vouchers",
        "cis on the go": "CIS on the Go",
        "housing": "Housing Referral",
        "dental": "Dental or Vision Referral/ Resources",
        "vision": "Dental or Vision Referral/ Resources",
        "professional mh": "Professional MH Referrals",
        "mental health": "Professional MH Referrals",
        "holiday": "Holiday Gifts",
    }
    for needle, category in aliases.items():
        if needle in normalized:
            return category
    for category in QPR_BASIC_NEED_ROWS:
        if normalize_text(category.rstrip(":")) == normalized.rstrip(":"):
            return category
    return "Other:"


def _canonical_support_tier(raw: object) -> str | None:
    text = normalize_text(raw)
    if not text:
        return None
    if "iii" in text or "3" in text:
        return "Tier III"
    if "ii" in text or "2" in text:
        return "Tier II"
    if "i" in text or "1" in text:
        return "Tier I"
    return None


def _canonical_goal_domain(raw: object) -> str:
    text = normalize_text(raw)
    if not text:
        return "Other"
    if "attend" in text:
        return "Attendance"
    if "behav" in text or "discipline" in text:
        return "Behavior"
    if "academ" in text or "grade" in text or "credit" in text:
        return "Academics"
    if "sel" in text or "social" in text or "emotional" in text:
        return "SEL"
    return "Other"


def _supports_from_text(raw: object) -> list[str]:
    text = clean_text(raw)
    if text is None:
        return []
    parts = [segment.strip() for segment in re.split(r"[;,|/]+", text)]
    cleaned = [part for part in parts if part]
    return cleaned


def _matches_school_filter(row_school: object, school_name: str | None) -> bool:
    if school_name is None:
        return True
    row_text = clean_text(row_school)
    return row_text is not None and school_labels_equivalent(row_text, school_name)


def _matches_grading_period(value: object, grading_period: str) -> bool:
    text = clean_text(value)
    if text is None:
        return True
    key = grading_period_key(text)
    if key is not None:
        return key == grading_period
    compact = normalize_text(text).replace(" ", "")
    if compact in {"q1", "quarter1"}:
        return grading_period == "1.0"
    if compact in {"q2", "quarter2"}:
        return grading_period == "2.0"
    if compact in {"q3", "quarter3"}:
        return grading_period == "3.0"
    if compact in {"q4", "quarter4", "eoy"}:
        return grading_period == "4.0"
    return True


def load_contract_export_rows(
    *,
    export_name: str,
    workbook_path: Path,
    grading_period: str,
    school_name: str | None = None,
) -> list[dict[str, Any]]:
    aliases = PAPER_PREPOP_FIELD_ALIASES.get(export_name)
    required = PAPER_PREPOP_REQUIRED_FIELDS.get(export_name)
    if aliases is None or required is None:
        raise ValueError(f"Unsupported paper report export contract: {export_name}")
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[workbook.sheetnames[0]]
        headers = header_values(worksheet)
        normalized_lookup = _normalized_headers(headers)
        canonical_to_index: dict[str, int] = {}
        missing_required: list[str] = []
        for canonical, alias_values in aliases.items():
            index = _resolve_alias_index(normalized_lookup, alias_values)
            if index is None:
                if canonical in required:
                    missing_required.append(canonical)
                continue
            canonical_to_index[canonical] = index
        if missing_required:
            missing = ", ".join(sorted(missing_required))
            raise ValueError(
                f"{export_name} workbook missing required columns ({missing}) in {workbook_path.name}."
            )
        rows: list[dict[str, Any]] = []
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if _row_is_empty(row):
                continue
            payload: dict[str, Any] = {}
            for field, index in canonical_to_index.items():
                payload[field] = row[index] if index < len(row) else None
            if not _matches_grading_period(payload.get("grading_period"), grading_period):
                continue
            if not _matches_school_filter(payload.get("school"), school_name):
                continue
            rows.append(payload)
        return rows
    finally:
        workbook.close()


def latest_recovered_raw_path(*, report_id: str, cisiphyus_root: Path = DEFAULT_CISIPHYUS_DIR) -> Path:
    latest_result = cisiphyus_root / "artifacts" / "latest" / report_id / "result.json"
    if not latest_result.exists():
        raise ValueError(f"Missing latest recovered artifact metadata for report_id={report_id!r}.")
    payload = json.loads(latest_result.read_text(encoding="utf-8"))
    raw_path_text = clean_text(payload.get("raw_path"))
    if raw_path_text is None:
        raise ValueError(f"Recovered artifact metadata missing raw_path for report_id={report_id!r}.")
    raw_path = Path(raw_path_text)
    if not raw_path.exists():
        raise ValueError(f"Recovered artifact raw workbook not found at {raw_path}.")
    return raw_path


def _detect_header_row(
    worksheet,
    *,
    max_scan_rows: int = 120,
    minimum_nonempty_cells: int = 4,
) -> tuple[int, list[str]]:
    for row_index, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True),
        start=1,
    ):
        values = [clean_text(value) for value in row]
        nonempty = [value for value in values if value is not None]
        if len(nonempty) < minimum_nonempty_cells:
            continue
        headers = [value if value is not None else "" for value in values]
        return row_index, headers
    raise ValueError(f"Could not detect header row on sheet {worksheet.title!r}.")


def _worksheet_by_alias(workbook, aliases: tuple[str, ...]):
    for sheet_name in workbook.sheetnames:
        normalized_sheet = normalize_text(sheet_name).replace("_", " ")
        for alias in aliases:
            normalized_alias = normalize_text(alias).replace("_", " ")
            if normalized_sheet == normalized_alias or normalized_alias in normalized_sheet:
                return workbook[sheet_name]
    joined = ", ".join(aliases)
    raise ValueError(f"Workbook does not contain required sheet alias(es): {joined}.")


def _row_payload(headers: list[str], row: tuple[object, ...]) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for index, header in enumerate(headers):
        header_text = clean_text(header)
        if header_text is None:
            continue
        payload[header_text] = row[index] if index < len(row) else None
    return payload


def _value_from_row_aliases(row_payload: dict[str, Any], aliases: tuple[str, ...]) -> object | None:
    normalized_row = {normalize_text(key): value for key, value in row_payload.items()}
    for alias in aliases:
        normalized_alias = normalize_text(alias)
        if normalized_alias in normalized_row:
            return normalized_row[normalized_alias]
    return None


def _school_year_contains_grading_period(raw_school_year: object, grading_period: str) -> bool:
    school_year = clean_text(raw_school_year)
    if school_year is None:
        return True
    # why: recovered support reports are usually pre-filtered to one SY; keep rows for target quarter pass-through
    return True


def _rows_from_recovered_sheet(
    *,
    workbook_path: Path,
    sheet_aliases: tuple[str, ...],
    minimum_nonempty_cells: int = 4,
) -> tuple[list[dict[str, Any]], str]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        worksheet = _worksheet_by_alias(workbook, sheet_aliases)
        header_row_index, headers = _detect_header_row(
            worksheet,
            minimum_nonempty_cells=minimum_nonempty_cells,
        )
        rows: list[dict[str, Any]] = []
        for row in worksheet.iter_rows(min_row=header_row_index + 1, values_only=True):
            if _row_is_empty(row):
                continue
            rows.append(_row_payload(headers, row))
        return rows, worksheet.title
    finally:
        workbook.close()


def load_recovered_tier_i_rows(
    *,
    workbook_path: Path,
    grading_period: str,
    school_name: str | None = None,
) -> list[dict[str, Any]]:
    rows, _sheet_name = _rows_from_recovered_sheet(
        workbook_path=workbook_path,
        sheet_aliases=("Tier I Support Detail",),
    )
    normalized: list[dict[str, Any]] = []
    for row in rows:
        school = _value_from_row_aliases(row, ("School",))
        if not _matches_school_filter(school, school_name):
            continue
        if not _school_year_contains_grading_period(_value_from_row_aliases(row, ("School Year",)), grading_period):
            continue
        support_name = clean_text(_value_from_row_aliases(row, ("Support Name", "Activity")))
        if support_name is None:
            continue
        planned_raw = _value_from_row_aliases(
            row,
            ("Planned as a Tier I Support in Support Plan? (Any Provider or Activity)",),
        )
        planned_support = support_name if normalize_text(planned_raw) == "yes" else ""
        support_note = clean_text(_value_from_row_aliases(row, ("Support Note",)))
        activity_text = clean_text(_value_from_row_aliases(row, ("Activity",)))
        support_category = clean_text(_value_from_row_aliases(row, ("Support Category",)))
        normalized.append(
            {
                "school": clean_text(school),
                "grading_period": grading_period,
                "support_category": support_category,
                "support_name": support_name,
                "activity": activity_text,
                "frequency": clean_text(_value_from_row_aliases(row, ("When was this support provided?",))),
                "sessions_count": parse_int(_value_from_row_aliases(row, ("Number of Sessions",)), default=0),
                "participants_count": parse_int(_value_from_row_aliases(row, ("# of Students Served",)), default=0),
                "parents_count": parse_int(
                    _value_from_row_aliases(row, ("# of Parents / Guardians Served",)),
                    default=0,
                ),
                "participants_selected_count": parse_int(
                    _value_from_row_aliases(row, ("Number of Students Selected in CISDM", "# of Students Served")),
                    default=0,
                ),
                "provider_type_1": clean_text(_value_from_row_aliases(row, ("Provider Type 1",))),
                "provider_name_1": clean_text(_value_from_row_aliases(row, ("Provider Name 1",))),
                "provider_type_2": clean_text(_value_from_row_aliases(row, ("Provider Type 2",))),
                "provider_name_2": clean_text(_value_from_row_aliases(row, ("Provider Name 2",))),
                "outcome": support_note or activity_text or "",
                "planned_support": planned_support,
            }
        )
    return normalized


def load_recovered_tier_ii_iii_rows(
    *,
    workbook_path: Path,
    grading_period: str,
    school_name: str | None = None,
) -> list[dict[str, Any]]:
    rows, _sheet_name = _rows_from_recovered_sheet(
        workbook_path=workbook_path,
        sheet_aliases=("CIS_byProvider_Tier2and3Support",),
    )
    normalized: list[dict[str, Any]] = []
    for row in rows:
        school = _value_from_row_aliases(row, ("School Where Support Was Provided",))
        if not _matches_school_filter(school, school_name):
            continue
        support_name = clean_text(_value_from_row_aliases(row, ("Student Support Name",)))
        activity = clean_text(_value_from_row_aliases(row, ("Activity",)))
        category = clean_text(_value_from_row_aliases(row, ("Student Support Category",)))
        provider_name = clean_text(_value_from_row_aliases(row, ("Provider Name",)))
        base_name = support_name or activity
        if base_name is None:
            continue
        display_name = f"{base_name} - {activity}" if support_name and activity and support_name != activity else base_name
        tier_ii_supports = parse_int(_value_from_row_aliases(row, ("# of Tier II Supports",)), default=0)
        tier_ii_hours = parse_int(_value_from_row_aliases(row, ("# of Tier II Hours",)), default=0)
        tier_iii_supports = parse_int(_value_from_row_aliases(row, ("# of Tier III Supports",)), default=0)
        tier_iii_hours = parse_int(_value_from_row_aliases(row, ("# of Tier III Hours",)), default=0)
        if tier_ii_supports > 0 or tier_ii_hours > 0:
            normalized.append(
                {
                    "school": clean_text(school),
                    "grading_period": grading_period,
                    "tier": "Tier II",
                    "goal_domain": category,
                    "support_name": display_name,
                    "frequency": "",
                    "participants_count": tier_ii_supports,
                    "outcome": activity or "",
                    "planned_support": "",
                }
            )
        if tier_iii_supports > 0 or tier_iii_hours > 0:
            normalized.append(
                {
                    "school": clean_text(school),
                    "grading_period": grading_period,
                    "tier": "Tier III",
                    "goal_domain": category,
                    "support_name": display_name,
                    "frequency": "",
                    "participants_count": tier_iii_supports,
                    "outcome": activity or "",
                    "planned_support": "",
                }
            )
    return normalized


def load_recovered_student_support_detail_rows(
    *,
    workbook_path: Path,
    grading_period: str,
    school_name: str | None = None,
) -> list[dict[str, Any]]:
    rows, _sheet_name = _rows_from_recovered_sheet(
        workbook_path=workbook_path,
        sheet_aliases=("CIS Student Support Detail",),
    )
    normalized: list[dict[str, Any]] = []
    for row in rows:
        school = _value_from_row_aliases(
            row,
            ("School Where Support Was Provided", "School Where \nSupport Was Provided", "Home School"),
        )
        if not _matches_school_filter(school, school_name):
            continue
        support_name = clean_text(_value_from_row_aliases(row, ("Student Support Name", "Activity")))
        if support_name is None:
            continue
        tier = _canonical_support_tier(_value_from_row_aliases(row, ("Tier",)))
        if tier not in {"Tier I", "Tier II", "Tier III"}:
            continue
        student_id = student_id_join_key(
            _value_from_row_aliases(row, ("Student ID", "Client ID", "Entity ID", "Student System ID"))
        )
        normalized.append(
            {
                "school": clean_text(school),
                "grading_period": grading_period,
                "tier": tier,
                "support_name": support_name,
                "student_id": student_id,
            }
        )
    return normalized


def deduped_participants_by_tier_support(
    rows: list[dict[str, Any]],
) -> tuple[dict[str, dict[str, int]], dict[str, int]]:
    buckets: dict[str, dict[str, set[str]]] = {"Tier I": {}, "Tier II": {}, "Tier III": {}}
    missing_student_id_rows = 0
    for row in rows:
        tier = _canonical_support_tier(row.get("tier"))
        support = clean_text(row.get("support_name"))
        if tier not in buckets or support is None:
            continue
        student_id = student_id_join_key(row.get("student_id"))
        if student_id is None:
            missing_student_id_rows += 1
            continue
        support_bucket = buckets[tier].setdefault(support, set())
        support_bucket.add(student_id)
    deduped = {
        tier: {support: len(student_ids) for support, student_ids in support_buckets.items()}
        for tier, support_buckets in buckets.items()
    }
    diagnostics = {
        "rows_total": len(rows),
        "rows_with_student_id": len(rows) - missing_student_id_rows,
        "rows_missing_student_id": missing_student_id_rows,
    }
    return deduped, diagnostics


def apply_deduped_participants_to_tier_summaries(
    *,
    tier_ii_summary: list[dict[str, str]],
    tier_iii_summary: list[dict[str, str]],
    deduped_counts: dict[str, dict[str, int]],
) -> dict[str, int]:
    # why: tier i is schoolwide; export # served is authoritative, not student_support_detail headcount

    def apply(summary_rows: list[dict[str, str]], count_map: dict[str, int]) -> int:
        normalized_counts = {normalize_text(name): count for name, count in count_map.items()}
        updated = 0
        for row in summary_rows:
            program = clean_text(row.get("program"))
            if program is None:
                continue
            count = normalized_counts.get(normalize_text(program))
            if count is None:
                continue
            row["participants"] = str(count)
            updated += 1
        return updated

    return {
        "tier_ii_supports_updated": apply(tier_ii_summary, deduped_counts.get("Tier II", {})),
        "tier_iii_supports_updated": apply(tier_iii_summary, deduped_counts.get("Tier III", {})),
    }


def summarize_school_context_rows(
    *,
    school_needs_rows: list[dict[str, Any]],
    school_demographics_rows: list[dict[str, Any]],
    school_improvement_rows: list[dict[str, Any]],
) -> list[str]:
    bullets: list[str] = []
    if school_needs_rows:
        top_need = school_needs_rows[0]
        need = clean_text(top_need.get("need")) or clean_text(top_need.get("needs"))
        description = clean_text(top_need.get("description"))
        if need or description:
            bullets.append(f"Top identified school need: {need or 'Not specified'} - {description or 'No description provided.'}")
    if school_demographics_rows:
        demo = school_demographics_rows[0]
        enrollment = clean_text(demo.get("total enrollment"))
        attendance = clean_text(demo.get("average daily attendance (%)"))
        chron_absent = clean_text(demo.get("% of students chronically absent"))
        fragments = []
        if enrollment:
            fragments.append(f"Enrollment: {enrollment}")
        if attendance:
            fragments.append(f"Average daily attendance: {attendance}%")
        if chron_absent:
            fragments.append(f"Chronically absent: {chron_absent}%")
        if fragments:
            bullets.append("School demographics snapshot - " + "; ".join(fragments) + ".")
    if school_improvement_rows:
        improvement = school_improvement_rows[0]
        priorities = clean_text(improvement.get("key priorities"))
        community = clean_text(improvement.get("summary of community data"))
        if priorities:
            bullets.append(f"School improvement priorities: {priorities[:350]}")
        if community:
            bullets.append(f"Community context summary: {community[:350]}")
    return bullets


def _summarize_one_school_goal_row(
    row: dict[str, Any],
    *,
    grading_period: str,
) -> dict[str, str]:
    quarter_field_map = {"1.0": "q1", "2.0": "q2", "3.0": "q3", "4.0": "eoy"}
    quarter_field = quarter_field_map.get(grading_period, "q1")
    return {
        "goal_metric": clean_text(row.get("goal_metric")) or "",
        "data_source": clean_text(row.get("data_source")) or "",
        "baseline": clean_text(row.get("baseline")) or "",
        "target": clean_text(row.get("target")) or "",
        "q1": clean_text(row.get("q1")) or "",
        "q2": clean_text(row.get("q2")) or "",
        "q3": clean_text(row.get("q3")) or "",
        "eoy": clean_text(row.get("eoy")) or "",
        "quarter_value": clean_text(row.get(quarter_field)) or "",
        "progress_narrative": clean_text(row.get("progress_narrative")) or "",
        "action_steps": clean_text(row.get("action_steps")) or "",
    }


def summarize_school_goals_rows(
    rows: list[dict[str, Any]],
    *,
    grading_period: str,
) -> list[dict[str, str]]:
    if not rows:
        return []
    return [_summarize_one_school_goal_row(row, grading_period=grading_period) for row in rows]


def _domain_from_metric(metric: object) -> str:
    sheet = template_sheet_for_metric(metric)
    if sheet in {"Attendance Rate_%", "Attendance_Days Absent", "Tardies"}:
        return "Attendance"
    if sheet in {"Suspensions", "Discliplinary Referrals", "Other Behavior Incidents", "Conduct"}:
        return "Behavior"
    if sheet in {"Core Course Grades", "GPA", "Credits - CTE Needed", "Reading Level", "Standardized Test Score"}:
        return "Academics"
    if sheet == "SEL":
        return "SEL"
    return "Other"


def load_student_metrics_domain_summary(
    *,
    workbook_path: Path,
    school_name: str | None,
) -> dict[str, dict[str, str]]:
    opened = _open_sheet_with_headers(workbook_path)
    if opened is None:
        return {}
    workbook, worksheet, header_lookup = opened
    try:
        required_headers = {"School", "Student ID", "Metric", "Target"}
        if sorted(required_headers - set(header_lookup)):
            return {}
        school_idx = header_lookup["School"]
        student_idx = header_lookup["Student ID"]
        metric_idx = header_lookup["Metric"]
        target_idx = header_lookup["Target"]
        qidx = {
            1: header_lookup.get(METRIC_SUMMARY_GRADING_PERIOD_HEADERS[0]),
            2: header_lookup.get(METRIC_SUMMARY_GRADING_PERIOD_HEADERS[1]),
            3: header_lookup.get(METRIC_SUMMARY_GRADING_PERIOD_HEADERS[2]),
            4: header_lookup.get(METRIC_SUMMARY_GRADING_PERIOD_HEADERS[3]),
        }
        domains = ("Attendance", "Behavior", "Academics", "SEL", "Other")
        goal_students: dict[str, set[str]] = {d: set() for d in domains}
        progress_students: dict[str, dict[int, set[str]]] = {
            d: {1: set(), 2: set(), 3: set(), 4: set()} for d in domains
        }
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if _row_is_empty(row):
                continue
            school_value = row[school_idx] if school_idx < len(row) else None
            school_text = clean_text(school_value)
            if school_name is not None and (
                school_text is None or normalize_text(school_text) != normalize_text(school_name)
            ):
                continue
            student_raw = row[student_idx] if student_idx < len(row) else None
            student_key = student_id_join_key(student_raw)
            if student_key is None:
                continue
            target_raw = row[target_idx] if target_idx < len(row) else None
            if clean_text(target_raw) is None:
                continue
            metric_raw = row[metric_idx] if metric_idx < len(row) else None
            domain = _domain_from_metric(metric_raw)
            goal_students[domain].add(student_key)
            for quarter, idx in qidx.items():
                if idx is None or idx >= len(row):
                    continue
                if clean_text(row[idx]) is not None:
                    progress_students[domain][quarter].add(student_key)
        out: dict[str, dict[str, str]] = {}
        for domain in domains:
            goal_count = len(goal_students[domain])
            q_progress = {q: len(progress_students[domain][q]) for q in (1, 2, 3, 4)}

            def pct(progress: int) -> str:
                if goal_count <= 0:
                    return ""
                return f"{round((progress / goal_count) * 100):d}%"

            out[domain] = {
                "eoy2526_percent": pct(q_progress[4]),
                "q1_goal": str(goal_count) if goal_count else "",
                "q1_prog": str(q_progress[1]) if q_progress[1] else "",
                "q1_percent": pct(q_progress[1]),
                "q2_goal": str(goal_count) if goal_count else "",
                "q2_prog": str(q_progress[2]) if q_progress[2] else "",
                "q2_percent": pct(q_progress[2]),
                "q3_goal": str(goal_count) if goal_count else "",
                "q3_prog": str(q_progress[3]) if q_progress[3] else "",
                "q3_percent": pct(q_progress[3]),
                "eoy_goal": str(goal_count) if goal_count else "",
                "eoy_prog": str(q_progress[4]) if q_progress[4] else "",
                "eoy_percent": pct(q_progress[4]),
            }
        return out
    finally:
        workbook.close()


def load_student_metrics_caseload_rows(
    *,
    workbook_path: Path,
    school_name: str | None,
) -> list[dict[str, str]]:
    opened = _open_sheet_with_headers(workbook_path)
    if opened is None:
        return []
    workbook, worksheet, header_lookup = opened
    try:
        required_headers = {"School", "Student ID", "School Year"}
        if sorted(required_headers - set(header_lookup)):
            return []
        school_idx = header_lookup["School"]
        student_idx = header_lookup["Student ID"]
        school_year_idx = header_lookup["School Year"]
        begin_idx = header_lookup.get("Enrollment BeginDate")
        end_idx = header_lookup.get("Enrollment EndDate")
        status_idx = header_lookup.get("Enrollment Status")
        out: list[dict[str, str]] = []
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if _row_is_empty(row):
                continue
            school_text = clean_text(row[school_idx] if school_idx < len(row) else None)
            if school_name is not None and (
                school_text is None or normalize_text(school_text) != normalize_text(school_name)
            ):
                continue
            student_key = student_id_join_key(row[student_idx] if student_idx < len(row) else None)
            if student_key is None:
                continue
            school_year = normalize_school_year_label(row[school_year_idx] if school_year_idx < len(row) else None)
            if school_year is None:
                continue
            out.append(
                {
                    "student_id": student_key,
                    "school": school_text or "",
                    "school_year": school_year,
                    "enrollment_begin_iso": to_iso_date(row[begin_idx]) if begin_idx is not None and begin_idx < len(row) else "",
                    "enrollment_end_iso": to_iso_date(row[end_idx]) if end_idx is not None and end_idx < len(row) else "",
                    "enrollment_status": clean_text(row[status_idx]) if status_idx is not None and status_idx < len(row) else "",
                }
            )
        return out
    finally:
        workbook.close()


def _percent_to_goal_string(*, progress_count: int, goal_count: int) -> str:
    if goal_count <= 0:
        return ""
    return f"{round((progress_count / goal_count) * 100):d}%"


def build_caseload_buildup_summary(
    *,
    site_staff_rows: list[dict[str, Any]],
    student_metric_rows: list[dict[str, str]],
    schedule_data: dict[str, Any],
    school_name: str | None,
) -> tuple[list[dict[str, str]], dict[str, Any]]:
    diagnostics: dict[str, Any] = {
        "school_name": school_name or "",
        "schedule_name": "",
        "rows": [],
    }
    schedule_name, schedule = matching_schedule(schedule_data, school_name)
    diagnostics["schedule_name"] = schedule_name or ""
    students_by_year: dict[str, dict[str, dict[str, str]]] = defaultdict(dict)
    for row in student_metric_rows:
        school_year = normalize_school_year_label(row.get("school_year"))
        if school_year is None:
            continue
        student_id = clean_text(row.get("student_id"))
        if student_id is None:
            continue
        bucket = students_by_year[school_year].setdefault(
            student_id,
            {
                "enrollment_begin_iso": "",
                "enrollment_end_iso": "",
                "enrollment_status": "",
            },
        )
        begin_iso = clean_text(row.get("enrollment_begin_iso")) or ""
        end_iso = clean_text(row.get("enrollment_end_iso")) or ""
        if begin_iso and (not bucket["enrollment_begin_iso"] or begin_iso < bucket["enrollment_begin_iso"]):
            bucket["enrollment_begin_iso"] = begin_iso
        if end_iso and (not bucket["enrollment_end_iso"] or end_iso > bucket["enrollment_end_iso"]):
            bucket["enrollment_end_iso"] = end_iso
        status = clean_text(row.get("enrollment_status")) or ""
        if status and not bucket["enrollment_status"]:
            bucket["enrollment_status"] = status
    rows: list[dict[str, str]] = []
    for school_year in CASELOAD_BUILDUP_TEMPLATE_SCHOOL_YEARS:
        goal_coordinators, goal_source = count_site_coordinators_for_school_year(
            site_staff_rows=site_staff_rows,
            school_year=school_year,
        )
        goal_count = goal_coordinators * 60
        goal_text = str(goal_count) if goal_count > 0 else ""
        row_out: dict[str, str] = {"school_year": school_year, "goal": goal_text}
        row_diag: dict[str, Any] = {
            "school_year": school_year,
            "goal_source": goal_source,
            "goal_site_coordinator_count": goal_coordinators,
            "blank_reasons": [],
        }
        students = students_by_year.get(school_year, {})
        for quarter, field_base in ((1, "q1"), (2, "q2"), (3, "q3"), (4, "eoy")):
            period = f"{quarter}.0"
            assess_iso = (schedule or {}).get("assessment_dates", {}).get(period)
            if assess_iso is None:
                row_out[field_base] = ""
                row_out[f"{field_base}_percent"] = ""
                row_diag["blank_reasons"].append(f"{field_base}:assessment_date_missing")
                continue
            try:
                assess_date = datetime.strptime(assess_iso, "%Y-%m-%d").date()
            except ValueError:
                row_out[field_base] = ""
                row_out[f"{field_base}_percent"] = ""
                row_diag["blank_reasons"].append(f"{field_base}:assessment_date_invalid")
                continue
            active_count = 0
            for student in students.values():
                status_key = normalize_text(student.get("enrollment_status")) or ""
                if status_key in {"inactive", "withdrawn"}:
                    continue
                begin_iso = clean_text(student.get("enrollment_begin_iso"))
                end_iso = clean_text(student.get("enrollment_end_iso"))
                begin_ok = True
                end_ok = True
                if begin_iso:
                    try:
                        begin_ok = datetime.strptime(begin_iso, "%Y-%m-%d").date() <= assess_date
                    except ValueError:
                        begin_ok = True
                if end_iso:
                    try:
                        end_ok = datetime.strptime(end_iso, "%Y-%m-%d").date() >= assess_date
                    except ValueError:
                        end_ok = True
                if begin_ok and end_ok:
                    active_count += 1
            row_out[field_base] = str(active_count) if active_count > 0 else ""
            row_out[f"{field_base}_percent"] = _percent_to_goal_string(
                progress_count=active_count,
                goal_count=goal_count,
            )
            if active_count == 0:
                row_diag["blank_reasons"].append(f"{field_base}:no_active_students")
            if goal_count <= 0:
                row_diag["blank_reasons"].append(f"{field_base}:goal_unavailable")
        diagnostics["rows"].append(row_diag)
        rows.append(row_out)
    return rows, diagnostics


def derive_caseload_buildup_from_sources(
    *,
    site_staff_list_path: Path,
    student_metrics_path: Path | None,
    deadlines_path: Path,
    school_name: str | None,
) -> tuple[list[dict[str, str]], dict[str, Any]]:
    if (
        school_name is None
        or student_metrics_path is None
        or not student_metrics_path.exists()
        or not site_staff_list_path.exists()
        or not deadlines_path.exists()
    ):
        return [], {"reason": "missing_required_sources"}
    schedule_data = load_reporting_deadlines(deadlines_path)
    site_staff_rows = load_site_staff_caseload_rows(
        site_staff_list_path=site_staff_list_path,
        school_name=school_name,
    )
    student_metric_rows = load_student_metrics_caseload_rows(
        workbook_path=student_metrics_path,
        school_name=school_name,
    )
    rows, diagnostics = build_caseload_buildup_summary(
        site_staff_rows=site_staff_rows,
        student_metric_rows=student_metric_rows,
        schedule_data=schedule_data,
        school_name=school_name,
    )
    diagnostics["source_counts"] = {
        "site_staff_rows": len(site_staff_rows),
        "student_metric_rows": len(student_metric_rows),
    }
    return rows, diagnostics


def build_qpr_draft_data_pack(
    *,
    grading_period: str,
    tier_i_rows: list[dict[str, Any]],
    tier_ii_iii_rows: list[dict[str, Any]],
    basic_needs_rows: list[dict[str, Any]],
    school_context_bullets: list[str],
    referral_rows: list[dict[str, Any]] | None,
    student_support_detail_rows: list[dict[str, Any]] | None = None,
    school_goals_rows: list[dict[str, Any]] | None = None,
    student_metrics_summary: dict[str, dict[str, str]] | None = None,
    parent_guardian_consent_rows: list[dict[str, Any]] | None = None,
    support_summary_by_student_rows: list[dict[str, Any]] | None = None,
    goal_tracking_student_goal_rows: list[dict[str, Any]] | None = None,
    caseload_buildup_summary: list[dict[str, str]] | None = None,
    caseload_buildup_diagnostics: dict[str, Any] | None = None,
    source_artifacts: dict[str, str] | None = None,
) -> dict[str, Any]:
    source_artifacts = source_artifacts or {}
    tier_i_summary = summarize_tier_i_rows(tier_i_rows)
    tier_ii_iii_summary = summarize_tier_ii_iii_rows(tier_ii_iii_rows)
    planned_vs_delivered = summarize_planned_vs_delivered(
        tier_i_rows=tier_i_rows,
        tier_ii_iii_rows=tier_ii_iii_rows,
    )
    planned_vs_delivered_top = summarize_planned_vs_delivered_top_rows(planned_vs_delivered)
    basic_needs_summary = summarize_basic_needs_rows(basic_needs_rows)
    school_goals_summaries = summarize_school_goals_rows(school_goals_rows or [], grading_period=grading_period)
    school_goals_summary = school_goals_summaries[0] if school_goals_summaries else {}
    parent_guardian_consent_summary = summarize_parent_guardian_consent_rows(parent_guardian_consent_rows or [])
    support_summary_by_student_summary = summarize_support_summary_by_student_rows(
        support_summary_by_student_rows or []
    )
    caseload_buildup_summary = caseload_buildup_summary or []
    caseload_buildup_diagnostics = caseload_buildup_diagnostics or {}
    goal_tracking_summary = summarize_goal_tracking_student_goal_rows(goal_tracking_student_goal_rows or [])
    if not school_goals_summary and goal_tracking_summary.get("top_metric"):
        school_goals_summary = {
            "goal_metric": clean_text(goal_tracking_summary.get("top_metric")) or "",
            "data_source": "goal_tracking_student_goals",
            "baseline": "",
            "target": "",
            "q1": "",
            "q2": "",
            "q3": "",
            "eoy": "",
            "quarter_value": "",
            "progress_narrative": clean_text(goal_tracking_summary.get("sample_progress_note")) or "",
            "action_steps": "",
        }
    dedupe_method = {
        "tier_i_participants": "tier_i_export_aggregate",
        "tier_ii_participants": "fallback_max_proxy",
        "tier_iii_participants": "fallback_max_proxy",
    }
    dedupe_diagnostics = {
        "rows_total": 0,
        "rows_with_student_id": 0,
        "rows_missing_student_id": 0,
    }
    denominator_policy = dict(PACK_DENOMINATOR_POLICY)
    if student_support_detail_rows:
        deduped_counts, dedupe_diagnostics = deduped_participants_by_tier_support(student_support_detail_rows)
        dedupe_updates = apply_deduped_participants_to_tier_summaries(
            tier_ii_summary=tier_ii_iii_summary["Tier II"],
            tier_iii_summary=tier_ii_iii_summary["Tier III"],
            deduped_counts=deduped_counts,
        )
        if dedupe_updates["tier_ii_supports_updated"] > 0:
            dedupe_method["tier_ii_participants"] = "unique_student_id"
        if dedupe_updates["tier_iii_supports_updated"] > 0:
            dedupe_method["tier_iii_participants"] = "unique_student_id"
    # why: referral summary section is intentionally coordinator-entered in written qpr.
    referral_summary_by_quarter: dict[int, dict[str, int]] = {}
    referral_summary = {field: 0 for field in REFERRAL_COLLECTION_INT_FIELDS}
    manual_required = []
    manual_required.append(
        {
            "field_group": "referral_summary",
            "reason": "Referral summary section is intentionally manual for site coordinator entry.",
            "source": "coordinator_entry",
        }
    )
    manual_required.append(
        {
            "field_group": "tier_progress_outcomes_narrative",
            "reason": "Recovered artifacts provide quantitative support activity but not complete qualitative outcome narrative.",
            "source": "coordinator_review",
        }
    )
    if not student_support_detail_rows:
        manual_required.append(
            {
                "field_group": "tier_participant_counts_dedupe",
                "reason": "Student Support Detail rows were not provided; Tier II/III participant counts use non-deduped fallback logic.",
                "source": "student_support_detail_export",
            }
        )
    elif dedupe_diagnostics["rows_with_student_id"] == 0:
        manual_required.append(
            {
                "field_group": "tier_participant_counts_dedupe",
                "reason": "Student Support Detail rows lacked student identifiers after filtering; Tier II/III participant counts use fallback logic.",
                "source": "student_support_detail_export",
            }
        )
    if not school_goals_summary:
        manual_required.append(
            {
                "field_group": "school_goals_progress",
                "reason": "School Goals and Progress export was not available or had no rows for this school/period.",
                "source": "school_goals_progress_export",
            }
        )
    if not student_metrics_summary:
        manual_required.append(
            {
                "field_group": "case_managed_student_progress_outcomes",
                "reason": "Student Metrics Summary export was not available or had no rows for this school.",
                "source": "student_metrics_summary",
            }
        )
    caseload_autofilled = any(
        clean_text(row.get(field))
        for row in caseload_buildup_summary
        for field in ("goal", "q1", "q2", "q3", "eoy")
    )
    if not caseload_autofilled:
        manual_required.append(
            {
                "field_group": "caseload_buildup",
                "reason": "Case-Load Build-Up auto-prefill could not derive values from site staff + student metrics for this run.",
                "source": "site_staff_list_and_student_metrics_summary",
            }
        )
    if not parent_guardian_consent_rows:
        manual_required.append(
            {
                "field_group": "parent_guardian_consent_summary",
                "reason": "Parent Guardian Consent export was not available for this school.",
                "source": "parent_guardian_consent",
            }
        )
    if not support_summary_by_student_rows:
        manual_required.append(
            {
                "field_group": "support_summary_by_student_validation",
                "reason": "Support Summary by Student export was not available for validation.",
                "source": "support_summary_by_student",
            }
        )
    if not goal_tracking_student_goal_rows:
        manual_required.append(
            {
                "field_group": "goal_tracking_student_goals_summary",
                "reason": "Goal Tracking - Student Goals export was not available for supplemental goal context.",
                "source": "goal_tracking_student_goals",
            }
        )
    return {
        "grading_period": grading_period,
        "source_artifacts": source_artifacts,
        "tier_i_summary": tier_i_summary,
        "tier_ii_summary": tier_ii_iii_summary["Tier II"],
        "tier_iii_summary": tier_ii_iii_summary["Tier III"],
        "planned_vs_delivered": planned_vs_delivered,
        "planned_vs_delivered_top": planned_vs_delivered_top,
        "basic_needs_summary": basic_needs_summary,
        "school_goals_summary": school_goals_summary,
        "school_goals_summaries": school_goals_summaries,
        "case_managed_student_progress_summary": student_metrics_summary or {},
        "school_context_bullets": school_context_bullets,
        "referral_summary": referral_summary,
        "referral_summary_by_quarter": {
            str(quarter): values for quarter, values in referral_summary_by_quarter.items()
        },
        "parent_guardian_consent_summary": parent_guardian_consent_summary,
        "support_summary_by_student_summary": support_summary_by_student_summary,
        "goal_tracking_student_goals_summary": goal_tracking_summary,
        "caseload_buildup_summary": caseload_buildup_summary,
        "caseload_buildup_diagnostics": caseload_buildup_diagnostics,
        "dedupe_method": dedupe_method,
        "dedupe_diagnostics": dedupe_diagnostics,
        "denominator_policy": denominator_policy,
        "manual_required": manual_required,
    }


def _join_unique_text_values(values: list[str], *, limit: int = 5) -> str:
    seen: set[str] = set()
    ordered: list[str] = []
    for raw in values:
        text = clean_text(raw)
        if text is None:
            continue
        key = normalize_text(text)
        if key in seen:
            continue
        seen.add(key)
        ordered.append(text)
        if len(ordered) >= limit:
            break
    return "; ".join(ordered)


def _provider_delivery_label(
    *,
    provider_type_1: str | None,
    provider_name_1: str | None,
    provider_type_2: str | None = None,
    provider_name_2: str | None = None,
) -> str:
    partner_tokens = {"partner", "community partner", "external partner"}
    types = [clean_text(provider_type_1), clean_text(provider_type_2)]
    names = [clean_text(provider_name_1), clean_text(provider_name_2)]
    for provider_type, provider_name in zip(types, names, strict=False):
        if provider_type is None and provider_name is None:
            continue
        normalized_type = normalize_text(provider_type or "")
        if any(token in normalized_type for token in partner_tokens):
            return "Partner-provided"
        if provider_name and normalize_text(provider_name) not in {"", "cis", "cis staff"}:
            if provider_type and "partner" in normalized_type:
                return "Partner-provided"
    return "CIS-provided"


def _format_tier_frequency_label(
    *,
    when_provided: str | None,
    sessions_count: int,
    deliveries: int,
) -> str:
    parts: list[str] = []
    if sessions_count > 0:
        session_word = "session" if sessions_count == 1 else "sessions"
        parts.append(f"{sessions_count} {session_word}")
    if when_provided:
        parts.append(when_provided)
    if deliveries > 1 and sessions_count <= 0:
        parts.append(f"{deliveries} entries")
    return " · ".join(parts)


def _format_participants_label(*, students: int, parents: int = 0) -> str:
    if students <= 0 and parents <= 0:
        return ""
    if parents > 0:
        return f"students: {students}; parents: {parents}"
    return str(students)


def summarize_tier_i_rows(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    grouped: dict[str, dict[str, Any]] = {}
    for row in rows:
        support = clean_text(row.get("support_name"))
        if support is None:
            continue
        bucket = grouped.setdefault(
            support,
            {
                "activities": [],
                "participants_sum": 0,
                "participants_max": 0,
                "participants_selected_max": 0,
                "parents_max": 0,
                "sessions_sum": 0,
                "frequencies": defaultdict(int),
                "outcomes": [],
                "deliveries": 0,
                "provider_labels": [],
            },
        )
        bucket["deliveries"] += 1
        activity = clean_text(row.get("activity"))
        if activity:
            bucket["activities"].append(activity)
        participant_count = parse_int(row.get("participants_count"), default=0)
        bucket["participants_sum"] += participant_count
        bucket["participants_max"] = max(bucket["participants_max"], participant_count)
        selected_count = parse_int(row.get("participants_selected_count"), default=0)
        if selected_count > 0:
            bucket["participants_selected_max"] = max(bucket["participants_selected_max"], selected_count)
        parents_count = parse_int(row.get("parents_count"), default=0)
        bucket["parents_max"] = max(bucket["parents_max"], parents_count)
        bucket["sessions_sum"] += parse_int(row.get("sessions_count"), default=0)
        frequency = clean_text(row.get("frequency"))
        if frequency:
            bucket["frequencies"][frequency] += 1
        outcome = clean_text(row.get("outcome"))
        if outcome:
            bucket["outcomes"].append(outcome)
        bucket["provider_labels"].append(
            _provider_delivery_label(
                provider_type_1=clean_text(row.get("provider_type_1")),
                provider_name_1=clean_text(row.get("provider_name_1")),
                provider_type_2=clean_text(row.get("provider_type_2")),
                provider_name_2=clean_text(row.get("provider_name_2")),
            )
        )
    summary: list[dict[str, str]] = []
    for support, bucket in grouped.items():
        frequencies: dict[str, int] = bucket["frequencies"]
        when_provided = max(frequencies.items(), key=lambda item: item[1])[0] if frequencies else ""
        frequency = _format_tier_frequency_label(
            when_provided=when_provided,
            sessions_count=bucket["sessions_sum"],
            deliveries=bucket["deliveries"],
        )
        provider_label = bucket["provider_labels"][0] if bucket["provider_labels"] else "CIS-provided"
        if frequency and provider_label:
            frequency = f"{frequency} ({provider_label})"
        elif provider_label:
            frequency = provider_label
        activities = _join_unique_text_values(bucket["activities"])
        program = f"{support}: {activities}" if activities and normalize_text(activities) != normalize_text(support) else support
        progress = _join_unique_text_values(bucket["outcomes"], limit=3)
        students = (
            bucket["participants_selected_max"]
            or bucket["participants_max"]
            or bucket["participants_sum"]
        )
        participants = _format_participants_label(students=students, parents=bucket["parents_max"])
        summary.append(
            {
                "program": program,
                "frequency": frequency,
                "participants": participants,
                "progress": progress,
            }
        )
    return sorted(summary, key=lambda row: normalize_text(row["program"]))


def summarize_tier_ii_iii_rows(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, str]]]:
    grouped: dict[str, dict[str, dict[str, Any]]] = {"Tier II": {}, "Tier III": {}}
    for row in rows:
        tier = _canonical_support_tier(row.get("tier"))
        if tier not in {"Tier II", "Tier III"}:
            continue
        support = clean_text(row.get("support_name"))
        if support is None:
            continue
        category = clean_text(row.get("goal_domain")) or support
        bucket = grouped[tier].setdefault(
            category,
            {
                "support_titles": [],
                "participants": set(),
                "participant_count_total": 0,
                "participant_count_max": 0,
                "frequencies": defaultdict(int),
                "outcomes": [],
                "provider_labels": [],
            },
        )
        bucket["support_titles"].append(support)
        student_key = student_id_join_key(row.get("student_id"))
        if student_key is not None:
            bucket["participants"].add(student_key)
        participant_count = parse_int(row.get("participants_count"), default=0)
        bucket["participant_count_total"] += participant_count
        bucket["participant_count_max"] = max(bucket["participant_count_max"], participant_count)
        frequency = clean_text(row.get("frequency"))
        if frequency:
            bucket["frequencies"][frequency] += 1
        outcome = clean_text(row.get("outcome"))
        if outcome:
            bucket["outcomes"].append(outcome)
        bucket["provider_labels"].append(
            _provider_delivery_label(
                provider_type_1=clean_text(row.get("provider_type_1")),
                provider_name_1=clean_text(row.get("provider_name_1")),
            )
        )
    out: dict[str, list[dict[str, str]]] = {"Tier II": [], "Tier III": []}
    for tier, category_map in grouped.items():
        for category, bucket in category_map.items():
            frequencies: dict[str, int] = bucket["frequencies"]
            top_frequency = max(frequencies.items(), key=lambda item: item[1])[0] if frequencies else ""
            participants = len(bucket["participants"]) or bucket["participant_count_max"] or bucket["participant_count_total"]
            titles = _join_unique_text_values(bucket["support_titles"])
            program = titles if titles else category
            progress = _join_unique_text_values(bucket["outcomes"], limit=3)
            provider_label = bucket["provider_labels"][0] if bucket["provider_labels"] else ""
            frequency = top_frequency
            if provider_label:
                frequency = f"{frequency} ({provider_label})" if frequency else provider_label
            out[tier].append(
                {
                    "program": program,
                    "frequency": frequency,
                    "participants": str(participants),
                    "progress": progress.strip(),
                }
            )
        out[tier] = sorted(out[tier], key=lambda row: normalize_text(row["program"]))
    return out


def summarize_planned_vs_delivered(
    *,
    tier_i_rows: list[dict[str, Any]],
    tier_ii_iii_rows: list[dict[str, Any]],
) -> list[dict[str, str]]:
    planned_counts: dict[str, int] = defaultdict(int)
    delivered_counts: dict[str, int] = defaultdict(int)
    for row in [*tier_i_rows, *tier_ii_iii_rows]:
        support_name = clean_text(row.get("support_name"))
        if support_name is not None:
            delivered_counts[support_name] += 1
        planned_raw = row.get("planned_support")
        for planned_name in _supports_from_text(planned_raw):
            planned_counts[planned_name] += 1
    support_names = sorted(set(planned_counts) | set(delivered_counts), key=normalize_text)
    summary: list[dict[str, str]] = []
    for support in support_names:
        planned = planned_counts.get(support, 0)
        delivered = delivered_counts.get(support, 0)
        coverage = "0%"
        if planned > 0:
            coverage = f"{round((delivered / planned) * 100)}%"
        elif delivered > 0:
            coverage = "100%"
        summary.append(
            {
                "support": support,
                "planned": str(planned),
                "delivered": str(delivered),
                "coverage": coverage,
            }
        )
    return summary


def _coverage_percent_value(coverage: str) -> int:
    text = clean_text(coverage) or ""
    if text.endswith("%"):
        text = text[:-1]
    return parse_int(text, default=0)


def summarize_planned_vs_delivered_top_rows(
    rows: list[dict[str, str]],
    *,
    limit: int = 5,
) -> list[dict[str, str]]:
    if limit <= 0:
        return []
    ranked = sorted(
        rows,
        key=lambda row: (
            -_coverage_percent_value(row.get("coverage", "")),
            -parse_int(row.get("delivered"), default=0),
            normalize_text(row.get("support")),
        ),
    )
    return ranked[:limit]


def format_planned_vs_delivered_for_docx(
    rows: list[dict[str, str]],
    *,
    limit: int = 3,
) -> list[str]:
    top_rows = summarize_planned_vs_delivered_top_rows(rows, limit=limit)
    if not top_rows:
        return []
    lines = ["- planned vs delivered (top coverage):"]
    for row in top_rows:
        support = clean_text(row.get("support")) or "n/a"
        delivered = clean_text(row.get("delivered")) or "0"
        planned = clean_text(row.get("planned")) or "0"
        coverage = clean_text(row.get("coverage")) or "0%"
        lines.append(f"- {support}: {delivered} delivered / {planned} planned ({coverage})")
    return lines


def summarize_basic_needs_rows(rows: list[dict[str, Any]]) -> dict[str, dict[str, str]]:
    buckets: dict[str, dict[str, Any]] = {
        category: {"recipients": set(), "count_total": 0, "details": []} for category in QPR_BASIC_NEED_ROWS
    }
    for row in rows:
        category = _normalize_basic_need_category(row.get("need_category"))
        bucket = buckets.setdefault(category, {"recipients": set(), "count_total": 0, "details": []})
        recipient = student_id_join_key(row.get("recipient_id"))
        if recipient is not None:
            bucket["recipients"].add(recipient)
        bucket["count_total"] += parse_int(row.get("count"), default=0)
        detail = clean_text(row.get("details"))
        if detail:
            bucket["details"].append(detail)
    output: dict[str, dict[str, str]] = {}
    for category, bucket in buckets.items():
        unique_count = len(bucket["recipients"])
        if unique_count == 0 and bucket["count_total"] > 0:
            unique_count = bucket["count_total"]
        detail_text = ""
        if bucket["details"]:
            detail_text = bucket["details"][0]
        output[category] = {"count": str(unique_count) if unique_count else "", "description": detail_text}
    return output


def write_referral_collection_template(path: Path) -> dict[str, str]:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(REFERRAL_COLLECTION_HEADERS))
        writer.writeheader()
        writer.writerow(
            {
                "site_name": "",
                "school": "",
                "grading_period": "1.0",
                "total_referrals": "",
                "consents_received": "",
                "non_responsive_families": "",
                "denied_consent": "",
                "rejected_by_site_coordinator": "",
                "rejected_by_school": "",
                "referrals_from_last_year": "",
            }
        )
    return {"status": "ok", "template_path": str(path.resolve())}


def load_referral_collection_rows_for_school(
    *,
    referrals_path: Path,
    school_name: str | None = None,
    through_grading_period: str | None = None,
) -> list[dict[str, Any]]:
    through = canonical_grading_period(through_grading_period) if through_grading_period else None
    through_num = int(float(through)) if through else 4
    if not referrals_path.exists():
        raise ValueError(f"Referral collection file not found: {referrals_path}")
    with referrals_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        actual_headers = reader.fieldnames or []
        missing_headers = sorted(set(REFERRAL_COLLECTION_HEADERS) - set(actual_headers))
        if missing_headers:
            raise ValueError(f"Referral collection file missing headers: {', '.join(missing_headers)}")
        rows: list[dict[str, Any]] = []
        for row in reader:
            if not _matches_school_filter(row.get("school"), school_name):
                continue
            row_period = grading_period_key(row.get("grading_period"))
            if row_period is None:
                continue
            if int(float(row_period)) > through_num:
                continue
            payload = dict(row)
            for field in REFERRAL_COLLECTION_INT_FIELDS:
                raw = payload.get(field)
                if clean_text(raw) is None:
                    payload[field] = 0
                    continue
                parsed = parse_int(raw, default=-1)
                if parsed < 0:
                    raise ValueError(
                        f"Referral field {field!r} must be a non-negative integer for school "
                        f"{clean_text(payload.get('school')) or '<unknown>'}."
                    )
                payload[field] = parsed
            if row_period == "1.0" and clean_text(row.get("referrals_from_last_year")) is None:
                raise ValueError(
                    "Q1 referral rows must include referrals_from_last_year (0 is allowed when none)."
                )
            rows.append(payload)
    return rows


def load_referral_collection_rows(
    *,
    referrals_path: Path,
    grading_period: str,
    school_name: str | None = None,
) -> list[dict[str, Any]]:
    if not referrals_path.exists():
        raise ValueError(f"Referral collection file not found: {referrals_path}")
    with referrals_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        actual_headers = reader.fieldnames or []
        missing_headers = sorted(set(REFERRAL_COLLECTION_HEADERS) - set(actual_headers))
        if missing_headers:
            raise ValueError(f"Referral collection file missing headers: {', '.join(missing_headers)}")
        rows: list[dict[str, Any]] = []
        for row in reader:
            if not _matches_grading_period(row.get("grading_period"), grading_period):
                continue
            if not _matches_school_filter(row.get("school"), school_name):
                continue
            if grading_period == "1.0" and clean_text(row.get("referrals_from_last_year")) is None:
                raise ValueError(
                    "Q1 referral rows must include referrals_from_last_year (0 is allowed when none)."
                )
            payload = dict(row)
            for field in REFERRAL_COLLECTION_INT_FIELDS:
                raw = payload.get(field)
                if clean_text(raw) is None:
                    payload[field] = 0
                    continue
                parsed = parse_int(raw, default=-1)
                if parsed < 0:
                    raise ValueError(
                        f"Referral field {field!r} must be a non-negative integer for school "
                        f"{clean_text(payload.get('school')) or '<unknown>'}."
                    )
                payload[field] = parsed
            rows.append(payload)
    return rows


def summarize_referral_rows_by_quarter(rows: list[dict[str, Any]]) -> dict[int, dict[str, int]]:
    by_quarter: dict[int, dict[str, int]] = {
        quarter: {field: 0 for field in REFERRAL_COLLECTION_INT_FIELDS} for quarter in (1, 2, 3, 4)
    }
    for row in rows:
        row_period = grading_period_key(row.get("grading_period"))
        if row_period is None:
            continue
        quarter = int(float(row_period))
        if quarter not in by_quarter:
            continue
        for field in REFERRAL_COLLECTION_INT_FIELDS:
            by_quarter[quarter][field] += parse_int(row.get(field), default=0)
    return by_quarter


def summarize_referral_rows(
    rows: list[dict[str, Any]],
    *,
    grading_period: str,
) -> dict[str, int]:
    totals = {field: 0 for field in REFERRAL_COLLECTION_INT_FIELDS}
    for row in rows:
        if not _matches_grading_period(row.get("grading_period"), grading_period):
            continue
        for field in REFERRAL_COLLECTION_INT_FIELDS:
            totals[field] += parse_int(row.get(field), default=0)
    if grading_period != "1.0":
        totals["referrals_from_last_year"] = 0
    return totals


def _is_truthy_yes(value: object) -> bool:
    text = clean_text(value)
    if text is None:
        return False
    return normalize_text(text) in {"yes", "y", "true", "1", "met", "achieved", "completed"}


def summarize_parent_guardian_consent_rows(rows: list[dict[str, Any]]) -> dict[str, int]:
    students_all: set[str] = set()
    students_recent_consent: set[str] = set()
    students_recent_consent_enrolled: set[str] = set()
    for row in rows:
        student_key = student_id_join_key(row.get("student_id"))
        if student_key is None:
            continue
        students_all.add(student_key)
        if _is_truthy_yes(row.get("most_recent_consent")):
            students_recent_consent.add(student_key)
            enrollment_status = normalize_text(clean_text(row.get("enrollment_status")) or "")
            if enrollment_status in {"enrolled", "active"}:
                students_recent_consent_enrolled.add(student_key)
    return {
        "students_total": len(students_all),
        "students_recent_consent": len(students_recent_consent),
        "students_recent_consent_enrolled": len(students_recent_consent_enrolled),
    }


def load_recovered_support_summary_by_student_rows(
    *,
    workbook_path: Path,
    school_name: str | None = None,
) -> list[dict[str, Any]]:
    rows, _sheet_name = _rows_from_recovered_sheet(
        workbook_path=workbook_path,
        sheet_aliases=("Support Summary by Student",),
        minimum_nonempty_cells=8,
    )
    normalized: list[dict[str, Any]] = []
    for row in rows:
        school = _value_from_row_aliases(row, ("Current School", "School"))
        if not _matches_school_filter(school, school_name):
            continue
        student_key = student_id_join_key(_value_from_row_aliases(row, ("Student ID", "Client ID", "Student System ID")))
        if student_key is None:
            continue
        normalized.append(
            {
                "student_id": student_key,
                "enrollment_status": clean_text(_value_from_row_aliases(row, ("Enrollment Status",))),
                "tier_i_supports": parse_int(_value_from_row_aliases(row, ("# of Tier I Supports",)), default=0),
                "tier_ii_supports": parse_int(_value_from_row_aliases(row, ("# of Tier II Supports",)), default=0),
                "tier_iii_supports": parse_int(_value_from_row_aliases(row, ("# of Tier III Supports",)), default=0),
            }
        )
    return normalized


def summarize_support_summary_by_student_rows(rows: list[dict[str, Any]]) -> dict[str, int]:
    enrolled: set[str] = set()
    tier_i_students: set[str] = set()
    tier_ii_students: set[str] = set()
    tier_iii_students: set[str] = set()
    all_students: set[str] = set()
    for row in rows:
        student_key = student_id_join_key(row.get("student_id"))
        if student_key is None:
            continue
        all_students.add(student_key)
        enrollment_status = normalize_text(clean_text(row.get("enrollment_status")) or "")
        if enrollment_status in {"enrolled", "active"}:
            enrolled.add(student_key)
        if parse_int(row.get("tier_i_supports"), default=0) > 0:
            tier_i_students.add(student_key)
        if parse_int(row.get("tier_ii_supports"), default=0) > 0:
            tier_ii_students.add(student_key)
        if parse_int(row.get("tier_iii_supports"), default=0) > 0:
            tier_iii_students.add(student_key)
    return {
        "students_total": len(all_students),
        "students_enrolled": len(enrolled),
        "students_with_tier_i_supports": len(tier_i_students),
        "students_with_tier_ii_supports": len(tier_ii_students),
        "students_with_tier_iii_supports": len(tier_iii_students),
    }


def summarize_goal_tracking_student_goal_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    students_with_goals: set[str] = set()
    metric_counts: defaultdict[str, int] = defaultdict(int)
    goals_total = 0
    goals_achieved = 0
    sample_note = ""
    for row in rows:
        student_key = student_id_join_key(row.get("student_id"))
        metric = clean_text(row.get("metric"))
        if student_key is None or metric is None:
            continue
        goals_total += 1
        students_with_goals.add(student_key)
        metric_counts[metric] += 1
        if _is_truthy_yes(row.get("goal_achievement")):
            goals_achieved += 1
        if not sample_note:
            sample_note = (
                clean_text(row.get("progress_notes"))
                or clean_text(row.get("goal_narrative"))
                or clean_text(row.get("latest_progress"))
                or ""
            )
    top_metric = ""
    if metric_counts:
        top_metric = max(metric_counts.items(), key=lambda item: item[1])[0]
    return {
        "students_with_goals": len(students_with_goals),
        "goals_total": goals_total,
        "goals_achieved": goals_achieved,
        "top_metric": top_metric,
        "sample_progress_note": sample_note[:350] if sample_note else "",
    }


def _render_tier_table_rows(rows: list[dict[str, str]]) -> list[str]:
    if not rows:
        return ["| | | | |"]
    return [
        f"| {row['program']} | {row['frequency']} | {row['participants']} | {row['progress']} |" for row in rows
    ]


def _replace_table_rows_after_header(
    markdown_lines: list[str],
    *,
    header_prefix: str,
    new_rows: list[str],
) -> None:
    for index, line in enumerate(markdown_lines):
        if not line.startswith(header_prefix):
            continue
        start = index + 1
        end = start
        while end < len(markdown_lines) and markdown_lines[end].startswith("|"):
            end += 1
        markdown_lines[start:end] = new_rows
        return
    raise ValueError(f"Could not find table header {header_prefix!r} in report template.")


def _replace_basic_needs_table_rows(
    markdown_lines: list[str],
    basic_needs_summary: dict[str, dict[str, str]],
) -> None:
    for index, line in enumerate(markdown_lines):
        if not line.startswith("| ") or "|" not in line:
            continue
        cells = [cell.strip() for cell in line.strip().split("|")[1:-1]]
        if len(cells) < 3:
            continue
        category = cells[0]
        if category not in basic_needs_summary:
            continue
        payload = basic_needs_summary[category]
        cells[1] = payload["count"]
        cells[2] = payload["description"]
        markdown_lines[index] = f"| {' | '.join(cells)} |"


def _ensure_q1_referral_row(markdown_lines: list[str]) -> None:
    label = "**# of referrals from last year (Q1 only)**"
    for line in markdown_lines:
        if label in line:
            return
    for index, line in enumerate(markdown_lines):
        if "**# of referrals made to CIS this quarter**" in line:
            markdown_lines.insert(index + 1, f"| {label} | | | | | |")
            return


def _replace_referral_rows(
    markdown_lines: list[str],
    referral_summary: dict[str, int],
    *,
    grading_period: str,
) -> None:
    if grading_period == "1.0":
        _ensure_q1_referral_row(markdown_lines)
    quarter_col = int(float(grading_period))
    for index, line in enumerate(markdown_lines):
        if not line.startswith("| **"):
            continue
        cells = [cell.strip() for cell in line.strip().split("|")[1:-1]]
        if len(cells) < 6:
            continue
        label = normalize_text(cells[0].replace("**", ""))
        field = QPR_REFERRAL_LABEL_TO_FIELD.get(label)
        if field is None:
            continue
        quarter_values = [parse_int(value, default=0) for value in cells[1:5]]
        quarter_values[quarter_col - 1] = referral_summary.get(field, 0)
        if field == "referrals_from_last_year" and grading_period != "1.0":
            quarter_values[quarter_col - 1] = 0
        cells[1:5] = [str(value) if value > 0 else "" for value in quarter_values]
        total = sum(quarter_values)
        cells[5] = str(total) if total > 0 else ""
        markdown_lines[index] = f"| {' | '.join(cells)} |"


def _insert_planned_vs_delivered_section(markdown_lines: list[str], rows: list[dict[str, str]]) -> None:
    if not rows:
        return
    block = [
        "",
        "**Planned vs Delivered Supports**",
        "",
        "| **Support** | **Planned Entries** | **Delivered Entries** | **Coverage** |",
        "| --- | --- | --- | --- |",
    ]
    for row in rows:
        block.append(f"| {row['support']} | {row['planned']} | {row['delivered']} | {row['coverage']} |")
    for index, line in enumerate(markdown_lines):
        if line.strip() == "**Basic Needs Provided**":
            markdown_lines[index:index] = block
            return


def _insert_school_context_section(markdown_lines: list[str], bullets: list[str]) -> None:
    if not bullets:
        return
    block = [
        "",
        "**School Context Highlights (Auto-Drafted)**",
        "",
    ]
    block.extend([f"- {bullet}" for bullet in bullets])
    for index, line in enumerate(markdown_lines):
        if line.strip() == "**Progress on Planned Supports**":
            markdown_lines[index:index] = block
            return
    markdown_lines.extend(block)


def prepopulate_paper_report_markdown(
    *,
    draft_path: Path,
    output_path: Path,
    grading_period: str,
    tier_i_summary: list[dict[str, str]],
    tier_ii_summary: list[dict[str, str]],
    tier_iii_summary: list[dict[str, str]],
    basic_needs_summary: dict[str, dict[str, str]],
    planned_vs_delivered_summary: list[dict[str, str]],
    referral_summary: dict[str, int],
    school_context_bullets: list[str] | None = None,
) -> dict[str, Any]:
    lines = draft_path.read_text(encoding="utf-8").splitlines()
    _replace_table_rows_after_header(
        lines,
        header_prefix="| **Program (s)/ Supports Provided** _What Planned Support did you coordinate this grading period? Please provide a brief description_ |",
        new_rows=_render_tier_table_rows(tier_i_summary),
    )
    _replace_table_rows_after_header(
        lines,
        header_prefix="| **Program (s)/ Supports Provided** _What Planned Support did you coordinate this grading period to address the school-wide goal? Please provide a brief description_ |",
        new_rows=_render_tier_table_rows(tier_ii_summary),
    )
    # why: tier iii has identical header text to tier ii, so replace second occurrence manually
    second_header_seen = 0
    for index, line in enumerate(lines):
        if line.startswith(
            "| **Program (s)/ Supports Provided** _What Planned Support did you coordinate this grading period to address the school-wide goal? Please provide a brief description_ |"
        ):
            second_header_seen += 1
            if second_header_seen != 2:
                continue
            start = index + 1
            end = start
            while end < len(lines) and lines[end].startswith("|"):
                end += 1
            lines[start:end] = _render_tier_table_rows(tier_iii_summary)
            break
    _insert_planned_vs_delivered_section(lines, planned_vs_delivered_summary)
    _insert_school_context_section(lines, school_context_bullets or [])
    _replace_basic_needs_table_rows(lines, basic_needs_summary)
    _replace_referral_rows(lines, referral_summary, grading_period=grading_period)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {"output_path": str(output_path.resolve())}


def _load_site_staff_context(*, site_staff_list_path: Path, school_name: str | None) -> dict[str, str]:
    context = {
        "school": school_name or "",
        "site_coordinator": "",
        "manager": "",
        "data_poc": "",
        "support_team_members": "",
    }
    if school_name is None:
        return context
    alias_school_names = {
        "reading senior high school": "reading high school",
    }
    normalized_target = normalize_text(school_name)
    accepted_names = {normalized_target}
    alias_target = alias_school_names.get(normalized_target)
    if alias_target is not None:
        accepted_names.add(alias_target)
    workbook = load_workbook(site_staff_list_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[SITE_LIST_SHEET]
        headers = header_values(worksheet)
        lookup = _header_lookup(headers)
        school_column = lookup.get("School")
        if school_column is None:
            return context
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            school_value = row[school_column] if school_column < len(row) else None
            school_text = clean_text(school_value)
            if school_text is None or normalize_text(school_text) not in accepted_names:
                continue
            for key, header in (
                ("site_coordinator", SITE_LIST_SITE_COORDINATOR_HEADER),
                ("manager", "Manager"),
                ("data_poc", "Data POC"),
            ):
                col = lookup.get(header)
                if col is None or col >= len(row):
                    continue
                value = clean_text(row[col])
                if value:
                    context[key] = value
            # why: these columns hold cross-functional school support contributors
            support_fields = []
            for header in ("Funding/Deliverables", "SSS/SW", "AC", "PFL"):
                col = lookup.get(header)
                if col is None or col >= len(row):
                    continue
                value = clean_text(row[col])
                if value:
                    support_fields.append(f"{header}: {value}")
            context["support_team_members"] = "; ".join(support_fields)
            return context
    finally:
        workbook.close()
    return context


def _report_quarter_index(grading_period: str) -> int:
    try:
        raw = int(float(grading_period))
    except (TypeError, ValueError):
        return 4
    return max(1, min(raw, 4))


_QUARTER_PROGRESS_PATTERNS: dict[int, tuple[str, ...]] = {
    1: (r"\b1st quarter\b", r"\bfirst quarter\b", r"\bq1\b", r"\bquarter 1\b"),
    2: (r"\b2nd quarter\b", r"\bsecond quarter\b", r"\bq2\b", r"\bquarter 2\b"),
    3: (r"\b3rd quarter\b", r"\bthird quarter\b", r"\bq3\b", r"\bquarter 3\b"),
    4: (r"\b4th quarter\b", r"\bfourth quarter\b", r"\bq4\b", r"\beoy\b", r"\bend of year\b"),
}


def _progress_text_chunks(text: str) -> list[str]:
    chunks = [chunk.strip() for chunk in re.split(r"\n+|; ", text) if chunk.strip()]
    sentences: list[str] = []
    for chunk in chunks:
        for sentence in re.split(r"(?<=[.!?])\s+", chunk):
            cleaned = sentence.strip()
            if cleaned:
                sentences.append(cleaned)
    return sentences or chunks


def _docx_tier_progress_excerpt(progress: str, *, report_quarter: int, max_chars: int = 500) -> str:
    # WHY: full SYSDM note stacks blow docx table layout; keep report-quarter lines only.
    text = clean_text(progress) or ""
    if not text:
        return ""
    chunks = _progress_text_chunks(text)
    patterns = _QUARTER_PROGRESS_PATTERNS.get(report_quarter, ())
    selected = [
        chunk
        for chunk in chunks
        if any(re.search(pattern, chunk.lower()) for pattern in patterns)
    ]
    if not selected:
        selected = chunks[-1:] if chunks else [text]
    excerpt = " ".join(selected)
    if len(excerpt) > max_chars:
        excerpt = excerpt[: max_chars - 1].rstrip() + "…"
    return excerpt


def _set_table_cell_text(cell: object, value: str) -> None:
    # WHY: cell.text replaces all paragraphs and can disturb table styling in Word.
    text = value if value is not None else ""
    paragraphs = cell.paragraphs  # type: ignore[attr-defined]
    if not paragraphs:
        cell.add_paragraph(text)  # type: ignore[attr-defined]
        return
    paragraphs[0].text = text
    for paragraph in paragraphs[1:]:
        element = paragraph._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)


def _quarter_visible_value(*, value: str, quarter: int, report_quarter: int) -> str:
    if quarter <= report_quarter:
        return value
    return ""


def _write_qpr_docx(
    *,
    output_path: Path,
    grading_period: str,
    draft_pack: dict[str, Any],
    site_staff_context: dict[str, str],
    template_docx_path: Path | None = None,
) -> dict[str, str]:
    try:
        from docx import Document
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "python-docx is required for docx output. Install it or run within an environment that provides docx."
        ) from exc

    if template_docx_path is not None and template_docx_path.exists():
        document = Document(template_docx_path)
    else:
        document = Document()
    if len(document.tables) < QPR_TABLE_MIN_COUNT:
        raise RuntimeError(
            f"Template docx does not contain expected QPR tables (need >= {QPR_TABLE_MIN_COUNT})."
        )

    def set_cell(table_idx: int, row_idx: int, col_idx: int, value: str) -> None:
        if table_idx >= len(document.tables):
            return
        table = document.tables[table_idx]
        if row_idx >= len(table.rows) or col_idx >= len(table.columns):
            return
        _set_table_cell_text(table.cell(row_idx, col_idx), value)

    # WHY: site coordination + activities remain manual; clear scaffold cells only.

    report_quarter = _report_quarter_index(grading_period)

    submitted_to = clean_text(site_staff_context.get("manager")) or clean_text(site_staff_context.get("data_poc")) or ""
    set_cell(QPR_TABLE_HEADER, 0, 1, datetime.now().strftime("%m/%d/%Y"))
    set_cell(QPR_TABLE_HEADER, 0, 3, clean_text(site_staff_context.get("school")) or "")
    set_cell(QPR_TABLE_HEADER, 1, 1, clean_text(site_staff_context.get("site_coordinator")) or "")
    set_cell(QPR_TABLE_HEADER, 1, 3, submitted_to)
    set_cell(QPR_TABLE_HEADER, 2, 1, clean_text(site_staff_context.get("support_team_members")) or "")

    school_goals_rows = draft_pack.get("school_goals_summaries") or []
    if not school_goals_rows:
        legacy_goal = draft_pack.get("school_goals_summary", {})
        if legacy_goal:
            school_goals_rows = [legacy_goal]

    if QPR_TABLE_SITE_COORDINATION < len(document.tables):
        site_coord_table = document.tables[QPR_TABLE_SITE_COORDINATION]
        for row_i in range(1, len(site_coord_table.rows)):
            for col_i in range(len(site_coord_table.columns)):
                _set_table_cell_text(site_coord_table.cell(row_i, col_i), "")

    goals_table = document.tables[QPR_TABLE_SCHOOL_GOALS]
    first_goal_row = 1
    for row_i in range(first_goal_row, len(goals_table.rows)):
        for col_i in range(len(goals_table.rows[row_i].cells)):
            _set_table_cell_text(goals_table.cell(row_i, col_i), "")
    limit = max(0, len(goals_table.rows) - first_goal_row)
    for offset, school_goals in enumerate(school_goals_rows[:limit]):
        row_i = first_goal_row + offset
        _set_table_cell_text(goals_table.cell(row_i, 0), school_goals.get("goal_metric", ""))
        _set_table_cell_text(goals_table.cell(row_i, 1), school_goals.get("data_source", ""))
        _set_table_cell_text(goals_table.cell(row_i, 2), school_goals.get("baseline", ""))
        _set_table_cell_text(goals_table.cell(row_i, 3), school_goals.get("target", ""))
        _set_table_cell_text(
            goals_table.cell(row_i, 4),
            _quarter_visible_value(value=school_goals.get("q1", ""), quarter=1, report_quarter=report_quarter),
        )
        _set_table_cell_text(
            goals_table.cell(row_i, 5),
            _quarter_visible_value(value=school_goals.get("q2", ""), quarter=2, report_quarter=report_quarter),
        )
        _set_table_cell_text(
            goals_table.cell(row_i, 6),
            _quarter_visible_value(value=school_goals.get("q3", ""), quarter=3, report_quarter=report_quarter),
        )
        _set_table_cell_text(
            goals_table.cell(row_i, 7),
            _quarter_visible_value(value=school_goals.get("eoy", ""), quarter=4, report_quarter=report_quarter),
        )
    if school_goals_rows:
        primary_goal = school_goals_rows[0]
        set_cell(QPR_TABLE_NARRATIVE, 1, 0, primary_goal.get("progress_narrative", ""))
        set_cell(QPR_TABLE_NARRATIVE, 1, 1, primary_goal.get("action_steps", ""))

    if len(document.tables) > QPR_TABLE_CASE_OUTCOMES:
        progress = draft_pack.get("case_managed_student_progress_summary", {})
        row_map = {"Attendance": 2, "Behavior": 3, "Academics": 4, "SEL": 5, "Other": 6}
        for domain, row_i in row_map.items():
            payload = progress.get(domain, {})
            set_cell(QPR_TABLE_CASE_OUTCOMES, row_i, 1, payload.get("eoy2526_percent", ""))
            set_cell(QPR_TABLE_CASE_OUTCOMES, row_i, 2, payload.get("q1_goal", ""))
            set_cell(QPR_TABLE_CASE_OUTCOMES, row_i, 3, payload.get("q1_prog", ""))
            set_cell(QPR_TABLE_CASE_OUTCOMES, row_i, 4, payload.get("q1_percent", ""))
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                5,
                _quarter_visible_value(value=payload.get("q2_goal", ""), quarter=2, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                6,
                _quarter_visible_value(value=payload.get("q2_prog", ""), quarter=2, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                7,
                _quarter_visible_value(value=payload.get("q2_percent", ""), quarter=2, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                8,
                _quarter_visible_value(value=payload.get("q3_goal", ""), quarter=3, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                9,
                _quarter_visible_value(value=payload.get("q3_prog", ""), quarter=3, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                10,
                _quarter_visible_value(value=payload.get("q3_percent", ""), quarter=3, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                11,
                _quarter_visible_value(value=payload.get("eoy_goal", ""), quarter=4, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                12,
                _quarter_visible_value(value=payload.get("eoy_prog", ""), quarter=4, report_quarter=report_quarter),
            )
            set_cell(
                QPR_TABLE_CASE_OUTCOMES,
                row_i,
                13,
                _quarter_visible_value(value=payload.get("eoy_percent", ""), quarter=4, report_quarter=report_quarter),
            )

    def fill_support_table(table_idx: int, rows: list[dict[str, str]]) -> None:
        table = document.tables[table_idx]
        needed_rows = len(rows) + 2
        while len(table.rows) < needed_rows:
            table.add_row()
        for row_i in range(2, len(table.rows)):
            for col_i in range(4):
                _set_table_cell_text(table.cell(row_i, col_i), "")
        for row_i, item in enumerate(rows, start=2):
            _set_table_cell_text(table.cell(row_i, 0), item.get("program", ""))
            _set_table_cell_text(table.cell(row_i, 1), item.get("frequency", ""))
            _set_table_cell_text(table.cell(row_i, 2), item.get("participants", ""))
            cisdm_text = clean_text(item.get("progress")) or ""
            progress_excerpt = _docx_tier_progress_excerpt(
                cisdm_text,
                report_quarter=report_quarter,
            )
            _set_table_cell_text(table.cell(row_i, 3), progress_excerpt)

    fill_support_table(QPR_TABLE_TIER_I, draft_pack.get("tier_i_summary", []))
    fill_support_table(QPR_TABLE_TIER_II, draft_pack.get("tier_ii_summary", []))
    fill_support_table(QPR_TABLE_TIER_III, draft_pack.get("tier_iii_summary", []))

    basic_table = document.tables[QPR_TABLE_BASIC_NEEDS]
    basic_needs = draft_pack.get("basic_needs_summary", {})
    row_by_label: dict[str, int] = {}
    for row_i in range(1, len(basic_table.rows)):
        label = normalize_text(basic_table.cell(row_i, 0).text)
        if label:
            row_by_label[label] = row_i
    for category in QPR_BASIC_NEED_ROWS:
        payload = basic_needs.get(category, {"count": "", "description": ""})
        row_i = row_by_label.get(normalize_text(category))
        if row_i is None:
            continue
        _set_table_cell_text(basic_table.cell(row_i, 1), payload.get("count", ""))
        _set_table_cell_text(basic_table.cell(row_i, 2), payload.get("description", ""))

    referral_table = document.tables[QPR_TABLE_REFERRALS]
    for row_i in range(2, len(referral_table.rows)):
        for c in range(1, 6):
            _set_table_cell_text(referral_table.cell(row_i, c), "")

    if len(document.tables) > QPR_TABLE_CASELOAD:
        caseload_table = document.tables[QPR_TABLE_CASELOAD]
        by_year = {
            normalize_school_year_label(item.get("school_year")): item
            for item in (draft_pack.get("caseload_buildup_summary") or [])
            if normalize_school_year_label(item.get("school_year")) is not None
        }
        for row_i in range(1, len(caseload_table.rows)):
            school_year = normalize_school_year_label(caseload_table.cell(row_i, 0).text)
            if school_year is None:
                continue
            payload = by_year.get(school_year)
            if payload is None:
                continue
            values = (
                payload.get("goal", ""),
                payload.get("q1", ""),
                payload.get("q1_percent", ""),
                payload.get("q2", ""),
                payload.get("q2_percent", ""),
                payload.get("q3", ""),
                payload.get("q3_percent", ""),
                payload.get("eoy", ""),
                payload.get("eoy_percent", ""),
            )
            for col_i, value in enumerate(values, start=1):
                if col_i >= len(caseload_table.columns):
                    break
                _set_table_cell_text(caseload_table.cell(row_i, col_i), value)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return {"output_path": str(output_path.resolve())}


def coordinator_last_name(coordinator_name: str) -> str:
    if "," in coordinator_name:
        return coordinator_name.split(",", 1)[0].strip() or "Coordinator"
    parts = coordinator_name.split()
    return parts[-1] if parts else "Coordinator"


def filename_component(value: str) -> str:
    cleaned = "".join("-" if character in '\\/:*?"<>|' else character for character in value)
    return " ".join(cleaned.split()).strip() or "Unknown"


def _slugify_filename(value: str | None, *, fallback: str) -> str:
    text = clean_text(value) or fallback
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or fallback


def build_written_qpr_generation_paths(
    *,
    output_dir: Path,
    school_name: str | None,
    grading_period: str,
) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    school_slug = _slugify_filename(school_name, fallback="multi-site")
    gp_slug = canonical_grading_period(grading_period).replace(".", "p")
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    stem_prefix = f"{school_slug}--gp{gp_slug}"
    existing = sorted(output_dir.glob(f"*--qpr--{stem_prefix}--g*.docx"))
    generation = len(existing) + 1
    stem = f"{timestamp}--qpr--{stem_prefix}--g{generation:03d}"
    output_docx = output_dir / f"{stem}.docx"
    draft_pack_json = output_dir / f"{stem}-pack.json"
    return output_docx, draft_pack_json


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
    return template_path.resolve().parent


def issue(
    issues: list[dict[str, Any]],
    *,
    sheet: str,
    row: int | None,
    column: str | None,
    code: str,
    severity: str,
    original_value: object = None,
    repaired_value: object = None,
    message: str,
) -> None:
    issues.append(
        {
            "sheet": sheet,
            "row": row,
            "column": column,
            "code": code,
            "severity": severity,
            "original_value": "" if original_value is None else str(original_value),
            "repaired_value": "" if repaired_value is None else str(repaired_value),
            "message": message,
        }
    )


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


def latest_ended_grading_period(schedule_data: dict[str, Any], *, as_of: date) -> str | None:
    latest: str | None = None
    for period in AUTOMATION_DEFAULT_PERIODS:
        ended = False
        for schedule in schedule_data.get("schedules", {}).values():
            assessment_iso = (schedule.get("assessment_dates") or {}).get(period)
            if assessment_iso is None:
                continue
            assessment_date = assessment_date_cell_value(assessment_iso)
            if assessment_date is not None and assessment_date <= as_of:
                ended = True
                break
        if ended:
            latest = period
            break
    return latest


def _period_used_for_deadline(
    schedule: dict[str, Any],
    *,
    target_period: str,
) -> str:
    if target_period == "3.0":
        period_three = (schedule.get("assessment_dates") or {}).get("3.0")
        if period_three is None:
            return "2.0"
    return target_period


def resolve_eligible_site_period_pairs(
    *,
    site_rows: dict[str, list[dict[str, str]]],
    schedule_data: dict[str, Any],
    target_grading_period: str,
    as_of: date,
) -> dict[str, Any]:
    target = canonical_grading_period(target_grading_period)
    pairs: list[dict[str, str]] = []
    skipped: list[dict[str, str]] = []
    for site_name, rows in sorted(site_rows.items(), key=lambda item: normalize_text(item[0])):
        school_name = clean_text(rows[0].get("school")) if rows else None
        schedule_name, schedule = matching_schedule(schedule_data, school_name)
        if schedule is None:
            skipped.append({"site_name": site_name, "grading_period": target, "reason": "schedule_unmatched"})
            continue
        period_for_deadline = _period_used_for_deadline(schedule, target_period=target)
        assessment_iso = (schedule.get("assessment_dates") or {}).get(period_for_deadline)
        assessment = assessment_date_cell_value(assessment_iso) if assessment_iso else None
        if assessment is not None and as_of < assessment:
            skipped.append({"site_name": site_name, "grading_period": target, "reason": "target_period_not_closed"})
            continue
        due_iso = (schedule.get("due_dates") or {}).get(period_for_deadline)
        due_date = assessment_date_cell_value(due_iso) if due_iso else None
        if due_date is not None and as_of > due_date:
            skipped.append({"site_name": site_name, "grading_period": target, "reason": "target_period_past_due"})
            continue
        if schedule_period_has_reference_issue(schedule_data, schedule_name, period_for_deadline):
            skipped.append({"site_name": site_name, "grading_period": target, "reason": "schedule_reference_issue"})
            continue
        pairs.append({"site_name": site_name, "grading_period": target})
    return {"pairs": pairs, "skipped": skipped}


def _automation_record_key(*, grading_period: str, site_name: str) -> str:
    return f"{canonical_grading_period(grading_period)}|{normalize_text(site_name)}"


def select_actionable_automation_workset(
    *,
    site_rows: dict[str, list[dict[str, str]]],
    schedule_data: dict[str, Any],
    as_of: date,
    ledger_completed: dict[str, Any],
    force: bool,
) -> dict[str, Any]:
    best = {
        "target_grading_period": None,
        "eligible_pairs": [],
        "schedule_skipped": [],
        "ledger_skipped": [],
    }
    for period in AUTOMATION_DEFAULT_PERIODS:
        resolved = resolve_eligible_site_period_pairs(
            site_rows=site_rows,
            schedule_data=schedule_data,
            target_grading_period=period,
            as_of=as_of,
        )
        eligible: list[dict[str, str]] = []
        ledger_skipped: list[dict[str, str]] = []
        for pair in resolved["pairs"]:
            key = _automation_record_key(
                grading_period=pair["grading_period"],
                site_name=pair["site_name"],
            )
            status = (ledger_completed.get(key) or {}).get("status")
            if not force and status in AUTOMATION_COMPLETED_STATUSES:
                ledger_skipped.append(
                    {
                        "site_name": pair["site_name"],
                        "grading_period": pair["grading_period"],
                        "reason": "already_in_ledger",
                    }
                )
                continue
            eligible.append(pair)
        if eligible:
            return {
                "target_grading_period": period,
                "eligible_pairs": eligible,
                "schedule_skipped": resolved["skipped"],
                "ledger_skipped": ledger_skipped,
            }
        best = {
            "target_grading_period": period,
            "eligible_pairs": [],
            "schedule_skipped": resolved["skipped"],
            "ledger_skipped": ledger_skipped,
        }
    return best


def set_repaired_value(
    issues: list[dict[str, Any]],
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    column_name: str,
    code: str,
    message: str,
    value: object,
) -> None:
    column_index = header_lookup[column_name]
    cell = worksheet.cell(row=row_index, column=column_index)
    original = cell.value
    cell.value = value
    issue(
        issues,
        sheet=worksheet.title,
        row=row_index,
        column=column_name,
        code=code,
        severity="info",
        original_value=original,
        repaired_value=value,
        message=message,
    )


def workbook_has_structure_errors(
    workbook,
    template_schema: dict[str, Any],
    issues: list[dict[str, Any]],
) -> bool:
    structural_valid = True
    actual_sheets = set(workbook.sheetnames)
    expected_sheets = set(template_schema["sheet_order"])
    for sheet_name in template_schema["sheet_order"]:
        if sheet_name not in actual_sheets:
            structural_valid = False
            issue(
                issues,
                sheet=sheet_name,
                row=None,
                column=None,
                code="missing_required_sheet",
                severity=STRUCTURAL_ERROR,
                message=f"Worksheet {sheet_name} is missing from the submission.",
            )
            continue
        worksheet = workbook[sheet_name]
        expected_sheet = template_schema["sheets"][sheet_name]
        expected_headers = template_schema["sheets"][sheet_name]["headers"]
        actual_headers = header_values(worksheet)
        row_mismatch = (
            worksheet.max_row != expected_sheet["max_row"]
            if expected_sheet["max_row"] > 1
            else worksheet.max_row < 1
        )
        column_mismatch = worksheet.max_column != expected_sheet["max_column"]
        if column_mismatch and actual_headers == expected_headers:
            column_mismatch = False
        if row_mismatch or column_mismatch:
            structural_valid = False
            issue(
                issues,
                sheet=sheet_name,
                row=None,
                column=None,
                code="worksheet_dimensions_mismatch",
                severity=STRUCTURAL_ERROR,
                original_value=f"{worksheet.max_row}x{worksheet.max_column}",
                repaired_value=f"{expected_sheet['max_row']}x{expected_sheet['max_column']}",
                message="Worksheet dimensions do not match the reference template.",
            )
        if actual_headers != expected_headers:
            structural_valid = False
            issue(
                issues,
                sheet=sheet_name,
                row=1,
                column=None,
                code="header_mismatch",
                severity=STRUCTURAL_ERROR,
                original_value=" | ".join(actual_headers),
                repaired_value=" | ".join(expected_headers),
                message="Header row does not match the reference template.",
            )
        for header, coordinates in template_schema["sheets"][sheet_name]["formula_coordinates"].items():
            if header in TEMPLATE_HEADERS_PRESERVE_FORMULAS:
                continue
            column_index = expected_headers.index(header) + 1
            for coordinate in coordinates:
                cell = worksheet[coordinate]
                value = cell.value
                if not (isinstance(value, str) and value.startswith("=")):
                    structural_valid = False
                    issue(
                        issues,
                        sheet=sheet_name,
                        row=cell.row,
                        column=expected_headers[column_index - 1],
                        code="formula_overwritten",
                        severity=STRUCTURAL_ERROR,
                        original_value=value,
                        message=f"Formula expected at {coordinate} based on the reference template.",
                    )
    for extra_sheet in sorted(actual_sheets - expected_sheets):
        issue(
            issues,
            sheet=extra_sheet,
            row=None,
            column=None,
            code="unexpected_sheet",
            severity="warning",
            message=f"Worksheet {extra_sheet} is not part of the reference template.",
        )
    return structural_valid


def copy_submission_values_into_template(
    submission_workbook,
    repaired_workbook,
    *,
    template_schema: dict[str, Any],
) -> None:
    for sheet_name, sheet_schema in _iter_template_sheets(template_schema):
        source_ws = submission_workbook[sheet_name]
        target_ws = repaired_workbook[sheet_name]
        formula_headers = set(sheet_schema["formula_headers"])
        row_limit = sheet_schema["max_row"]
        max_row = source_ws.max_row if row_limit <= 1 else row_limit
        for column_index, header in enumerate(sheet_schema["headers"], start=1):
            if header in formula_headers or header in TEMPLATE_HEADERS_PRESERVE_FORMULAS:
                continue
            for row_index in range(2, max_row + 1):
                target_ws.cell(row=row_index, column=column_index).value = source_ws.cell(
                    row=row_index,
                    column=column_index,
                ).value


def populated_rows_by_sheet(
    submission_workbook,
    template_workbook,
    *,
    template_schema: dict[str, Any],
) -> dict[str, set[int]]:
    populated: dict[str, set[int]] = {}
    for sheet_name, sheet_schema in _iter_template_sheets(template_schema):
        source_ws = submission_workbook[sheet_name]
        template_ws = template_workbook[sheet_name]
        formula_headers = set(sheet_schema["formula_headers"])
        row_limit = sheet_schema["max_row"]
        max_row = max(source_ws.max_row, template_ws.max_row) if row_limit <= 1 else row_limit
        row_numbers: set[int] = set()
        for column_index, header in enumerate(sheet_schema["headers"], start=1):
            if header in formula_headers or header in TEMPLATE_HEADERS_PRESERVE_FORMULAS:
                continue
            for row_index in range(2, max_row + 1):
                source_value = source_ws.cell(row=row_index, column=column_index).value
                template_value = template_ws.cell(row=row_index, column=column_index).value
                if source_value != template_value:
                    row_numbers.add(row_index)
        populated[sheet_name] = row_numbers
    return populated


def validate_row_identity(
    issues: list[dict[str, Any]],
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    caseload_lookup: dict[str, dict[str, str]],
) -> None:
    if "Student ID" not in header_lookup:
        return
    student_id = _cell_text_by_header(
        worksheet,
        row_index=row_index,
        header_lookup=header_lookup,
        header_name="Student ID",
    )
    if student_id is None:
        return
    caseload_row = caseload_lookup.get(student_id)
    if caseload_row is not None:
        return
    first_name = _cell_text_by_header(
        worksheet,
        row_index=row_index,
        header_lookup=header_lookup,
        header_name="First Name",
    )
    last_name = _cell_text_by_header(
        worksheet,
        row_index=row_index,
        header_lookup=header_lookup,
        header_name="Last Name",
    )
    if first_name is None and last_name is None:
        return
    issue(
        issues,
        sheet=worksheet.title,
        row=row_index,
        column="Student ID",
        code="student_identity_lookup_missing",
        severity="warning",
        original_value=student_id,
        message="Student identity fields were not repaired because no authoritative lookup source was provided.",
    )


def repair_shared_defaults(
    issues: list[dict[str, Any]],
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    template_schema: dict[str, Any],
    caseload_lookup: dict[str, dict[str, str]],
) -> None:
    defaults = template_schema["field_options"]
    student_id = None
    caseload_row: dict[str, str] | None = None
    if "Student ID" in header_lookup:
        student_id = _cell_text_by_header(
            worksheet,
            row_index=row_index,
            header_lookup=header_lookup,
            header_name="Student ID",
        )
        if student_id is not None:
            caseload_row = caseload_lookup.get(student_id)

    if caseload_row is not None:
        caseload_fields = (
            ("School", "school", "filled_school_from_caseload", "corrected_school_from_caseload"),
            ("First Name", "first_name", "filled_first_name_from_caseload", "corrected_first_name_from_caseload"),
            ("Last Name", "last_name", "filled_last_name_from_caseload", "corrected_last_name_from_caseload"),
        )
        for header_name, caseload_key, fill_code, correct_code in caseload_fields:
            if header_name not in header_lookup:
                continue
            caseload_value = caseload_row.get(caseload_key) or None
            if caseload_value is None:
                continue
            current = _cell_text_by_header(
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                header_name=header_name,
            )
            if current is None:
                set_repaired_value(
                    issues,
                    worksheet,
                    row_index=row_index,
                    header_lookup=header_lookup,
                    column_name=header_name,
                    code=fill_code,
                    message=f"Filled {header_name} from the caseload export for Student ID {student_id}.",
                    value=caseload_value,
                )
            elif current != caseload_value:
                set_repaired_value(
                    issues,
                    worksheet,
                    row_index=row_index,
                    header_lookup=header_lookup,
                    column_name=header_name,
                    code=correct_code,
                    message=f"Corrected {header_name} to the caseload export value for Student ID {student_id}.",
                    value=caseload_value,
                )

    if "Organization" in header_lookup:
        current = worksheet.cell(row=row_index, column=header_lookup["Organization"]).value
        current_text = clean_text(current)
        if current_text is None:
            set_repaired_value(
                issues,
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                column_name="Organization",
                code="set_organization_canonical",
                message="Filled Organization with the agency standard value.",
                value=ORGANIZATION_CANONICAL,
            )
        elif current_text != ORGANIZATION_CANONICAL:
            set_repaired_value(
                issues,
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                column_name="Organization",
                code="corrected_organization_canonical",
                message="Corrected Organization to the agency standard value.",
                value=ORGANIZATION_CANONICAL,
            )
    if "School Year" in header_lookup:
        current = worksheet.cell(row=row_index, column=header_lookup["School Year"]).value
        if clean_text(current) is None:
            school_years = defaults.get("School Year", [])
            if school_years:
                set_repaired_value(
                    issues,
                    worksheet,
                    row_index=row_index,
                    header_lookup=header_lookup,
                    column_name="School Year",
                    code="defaulted_school_year",
                    message="Filled School Year from the reference template options.",
                    value=school_years[0],
                )


def repair_grading_period(
    issues: list[dict[str, Any]],
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    schedule_data: dict[str, Any],
) -> str | None:
    if "Grading Period" not in header_lookup:
        return None
    cell = worksheet.cell(row=row_index, column=header_lookup["Grading Period"])
    canonical = grading_period_key(cell.value)
    if canonical is None and "Grading Period / Assessment Date" in header_lookup and "School" in header_lookup:
        school = _cell_text_by_header(
            worksheet,
            row_index=row_index,
            header_lookup=header_lookup,
            header_name="School",
        )
        status, schedule_name, schedule = schedule_status(schedule_data, school)
        if status == "matched" and schedule_name is not None and schedule is not None:
            assessment_cell = worksheet.cell(
                row=row_index,
                column=header_lookup["Grading Period / Assessment Date"],
            )
            canonical = grading_period_from_assessment_date(schedule, assessment_cell.value)
            if canonical is not None:
                set_repaired_value(
                    issues,
                    worksheet,
                    row_index=row_index,
                    header_lookup=header_lookup,
                    column_name="Grading Period",
                    code="derived_grading_period",
                    message=f"Derived Grading Period from {schedule_name} assessment date.",
                    value=grading_period_cell_value(canonical),
                )
                return canonical
    if canonical is None:
        return None
    if not grading_period_cell_is_preferred(cell.value, canonical):
        set_repaired_value(
            issues,
            worksheet,
            row_index=row_index,
            header_lookup=header_lookup,
            column_name="Grading Period",
            code="normalized_grading_period",
            message="Normalized Grading Period to the canonical template option.",
            value=grading_period_cell_value(canonical),
        )
    return canonical


def repair_schedule_fields(
    issues: list[dict[str, Any]],
    worksheet,
    *,
    row_index: int,
    header_lookup: dict[str, int],
    schedule_data: dict[str, Any],
    grading_period: str | None,
) -> None:
    if grading_period is None:
        return
    school = _cell_text_by_header(
        worksheet,
        row_index=row_index,
        header_lookup=header_lookup,
        header_name="School",
    )
    status, schedule_name, schedule = schedule_status(schedule_data, school)
    if status == "unmatched" or schedule_name is None or schedule is None:
        return
    if status == "suspect" and schedule_period_has_reference_issue(
        schedule_data,
        schedule_name,
        grading_period,
    ):
        issue(
            issues,
            sheet=worksheet.title,
            row=row_index,
            column="School",
            code="reference_schedule_suspect",
            severity="warning",
            original_value=school,
            message="Schedule-derived fields were not repaired because the reference deadline row looks unreliable.",
        )
        return
    if "Grading Period / Assessment Date" in header_lookup:
        assessment_date = schedule["assessment_dates"].get(grading_period)
        if assessment_date is not None:
            assessment_cell = worksheet.cell(
                row=row_index,
                column=header_lookup["Grading Period / Assessment Date"],
            )
            current_iso = to_iso_date(assessment_cell.value)
            is_date_typed = isinstance(assessment_cell.value, (date, datetime)) and not isinstance(
                assessment_cell.value,
                bool,
            )
            if current_iso != assessment_date or not is_date_typed:
                empty = clean_text(assessment_cell.value) is None
                set_repaired_value(
                    issues,
                    worksheet,
                    row_index=row_index,
                    header_lookup=header_lookup,
                    column_name="Grading Period / Assessment Date",
                    code=(
                        "derived_assessment_date"
                        if empty
                        else (
                            "corrected_assessment_date"
                            if current_iso != assessment_date
                            else "normalized_assessment_date_type"
                        )
                    ),
                    message=(
                        f"Filled assessment date from {schedule_name} {period_code(grading_period)}."
                        if empty
                        else (
                            f"Aligned assessment date to {schedule_name} {period_code(grading_period)} reporting deadlines."
                            if current_iso != assessment_date
                            else "Normalized assessment date to a date-typed cell for import compatibility."
                        )
                    ),
                    value=assessment_date_cell_value(assessment_date),
                )
def repair_workbook(
    workbook,
    *,
    template_schema: dict[str, Any],
    schedule_data: dict[str, Any],
    issues: list[dict[str, Any]],
    populated_rows: dict[str, set[int]],
    caseload_lookup: dict[str, dict[str, str]],
) -> None:
    for sheet_name, sheet_schema in _iter_template_sheets(template_schema, workbook=workbook):
        worksheet = workbook[sheet_name]
        header_lookup = {header: index for index, header in enumerate(sheet_schema["headers"], start=1)}
        for row_index in sorted(populated_rows.get(sheet_name, set())):
            repair_shared_defaults(
                issues,
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                template_schema=template_schema,
                caseload_lookup=caseload_lookup,
            )
            validate_row_identity(
                issues,
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                caseload_lookup=caseload_lookup,
            )
            grading_period = repair_grading_period(
                issues,
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                schedule_data=schedule_data,
            )
            repair_schedule_fields(
                issues,
                worksheet,
                row_index=row_index,
                header_lookup=header_lookup,
                schedule_data=schedule_data,
                grading_period=grading_period,
            )


def write_issue_csv(path: Path, issues: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_FIELDS)
        writer.writeheader()
        for row in issues:
            writer.writerow({field: row.get(field, "") for field in REPORT_FIELDS})


def write_issue_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)


def build_summary(issues: list[dict[str, Any]]) -> dict[str, int]:
    summary = {"error": 0, "warning": 0, "info": 0}
    for row in issues:
        severity = row["severity"]
        summary[severity] = summary.get(severity, 0) + 1
    return summary


def validate_workbook(
    submission_path: Path,
    *,
    repaired_path: Path | None = None,
    report_json_path: Path | None = None,
    report_csv_path: Path | None = None,
    template_path: Path = DEFAULT_TEMPLATE_PATH,
    deadlines_path: Path = DEFAULT_DEADLINES_PATH,
    caseload_path: Path | None = DEFAULT_CASELOAD_PATH,
) -> dict[str, Any]:
    if repaired_path is None:
        repaired_path = submission_path.with_name(f"{submission_path.stem}.repaired{submission_path.suffix}")
    template_schema = build_template_schema(template_path)
    schedule_data = load_reporting_deadlines(deadlines_path)
    caseload_lookup = load_caseload_lookup(caseload_path)
    submission_workbook = load_workbook(submission_path, data_only=False)
    issues: list[dict[str, Any]] = []
    structural_valid = workbook_has_structure_errors(submission_workbook, template_schema, issues)
    if structural_valid:
        repaired_workbook = load_workbook(template_path, data_only=False)
        populated_rows = populated_rows_by_sheet(
            submission_workbook,
            repaired_workbook,
            template_schema=template_schema,
        )
        copy_submission_values_into_template(
            submission_workbook,
            repaired_workbook,
            template_schema=template_schema,
        )
        repair_workbook(
            repaired_workbook,
            template_schema=template_schema,
            schedule_data=schedule_data,
            issues=issues,
            populated_rows=populated_rows,
            caseload_lookup=caseload_lookup,
        )
        repaired_path.parent.mkdir(parents=True, exist_ok=True)
        repaired_workbook.save(repaired_path)
        repaired_workbook.close()
    submission_workbook.close()
    payload = {
        "submission_path": str(submission_path.resolve()),
        "repaired_path": str(repaired_path.resolve()) if structural_valid else None,
        "template_path": str(template_path.resolve()),
        "deadlines_path": str(deadlines_path.resolve()),
        "caseload_path": str(caseload_path.resolve()) if caseload_path is not None and caseload_path.exists() else None,
        "structural_valid": structural_valid,
        "issue_count": len(issues),
        "repair_count": sum(1 for row in issues if row["severity"] == "info"),
        "summary": build_summary(issues),
        "issues": issues,
        "reference_issues": schedule_data["reference_issues"],
    }
    if report_json_path is not None:
        write_issue_json(report_json_path, payload)
    if report_csv_path is not None:
        write_issue_csv(report_csv_path, issues)
    return payload


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
        "results": results,
        "skipped": skipped,
        "use_site_staff_filter": use_site_staff_filter,
        "roster_coordinators_excluded_by_site_staff": excluded_by_site_staff,
        "excel_metric_context_note_hint": EXCEL_METRIC_CONTEXT_NOTE_HINT,
    }


def bridge_legacy_imports_to_unified_qpr(
    *,
    non_goal_path: Path,
    progress_path: Path,
    template_path: Path,
    output_path: Path | None = None,
    write_output: bool = False,
    report_json_path: Path | None = None,
) -> dict[str, Any]:
    template_schema = build_template_schema(template_path)
    target_workbook = load_workbook(template_path, data_only=False)
    template_student_lookup = _bridge_template_student_lookup(target_workbook, template_schema)
    for sheet_name in template_schema["sheet_order"]:
        if sheet_name in BRIDGE_SHEET_EXCLUDE:
            continue
        target_ws = target_workbook[sheet_name]
        sheet_headers = template_schema["sheets"][sheet_name]["headers"]
        max_columns = len(sheet_headers)
        formula_cells: set[str] = set()
        for coordinates in template_schema["sheets"][sheet_name]["formula_coordinates"].values():
            formula_cells.update(coordinates)
        for row_index in range(2, target_ws.max_row + 1):
            for column_index in range(1, max_columns + 1):
                cell = target_ws.cell(row=row_index, column=column_index)
                if cell.coordinate in formula_cells:
                    continue
                cell.value = None
    row_cursor: dict[str, int] = {
        sheet_name: 2
        for sheet_name in template_schema["sheet_order"]
        if sheet_name not in BRIDGE_SHEET_EXCLUDE
    }
    row_signatures: dict[str, set[tuple[object, ...]]] = {
        sheet_name: set()
        for sheet_name in template_schema["sheet_order"]
        if sheet_name not in BRIDGE_SHEET_EXCLUDE
    }
    row_best_by_key: dict[str, dict[tuple[str, ...], dict[str, int]]] = {
        sheet_name: {}
        for sheet_name in template_schema["sheet_order"]
        if sheet_name not in BRIDGE_SHEET_EXCLUDE
    }
    per_sheet: dict[str, dict[str, int]] = {
        sheet_name: {
            "rows_read": 0,
            "rows_written": 0,
            "rows_skipped_empty": 0,
            "rows_skipped_duplicate": 0,
            "rows_skipped_missing_student_id": 0,
            "rows_filled_from_template_defaults": 0,
            "rows_replaced_by_richer_match": 0,
            "rows_student_id_corrected_by_name": 0,
            "rows_student_name_corrected_by_id": 0,
            "rows_unresolved_student_id_mismatch": 0,
        }
        for sheet_name in template_schema["sheet_order"]
        if sheet_name not in BRIDGE_SHEET_EXCLUDE
    }
    errors: list[dict[str, str]] = []
    alias_hits: list[dict[str, str]] = []
    unmapped_source_columns: list[dict[str, str]] = []
    sources = (
        (BRIDGE_SOURCE_NON_GOAL, non_goal_path),
        (BRIDGE_SOURCE_PROGRESS, progress_path),
    )
    for source_kind, source_path in sources:
        workbook = load_workbook(source_path, read_only=True, data_only=True)
        try:
            for sheet_name in template_schema["sheet_order"]:
                if sheet_name in BRIDGE_SHEET_EXCLUDE:
                    continue
                if sheet_name not in workbook.sheetnames:
                    errors.append(
                        {
                            "source_kind": source_kind,
                            "sheet": sheet_name,
                            "code": "MISSING_SOURCE_SHEET",
                            "message": f"Missing required sheet {sheet_name!r} in {source_path.name}.",
                        }
                    )
                    continue
                source_ws = workbook[sheet_name]
                target_ws = target_workbook[sheet_name]
                source_headers = header_values(source_ws)
                target_headers = template_schema["sheets"][sheet_name]["headers"]
                template_defaults = template_schema["sheets"][sheet_name]["defaults"]
                header_mapping, unmapped_cols, mapping_alias_hits = _bridge_target_to_source_columns(
                    source_kind=source_kind,
                    sheet_name=sheet_name,
                    source_headers=source_headers,
                    target_headers=target_headers,
                )
                alias_hits.extend(mapping_alias_hits)
                for column_name in unmapped_cols:
                    if column_name == "Attendnace Rate (Temporary)":
                        continue
                    unmapped_source_columns.append(
                        {
                            "source_kind": source_kind,
                            "sheet": sheet_name,
                            "column": column_name,
                        }
                    )
                missing_targets = [header for header in target_headers if header not in header_mapping]
                missing_without_defaults = [
                    header for header in missing_targets if clean_text(template_defaults.get(header)) is None
                ]
                if missing_without_defaults:
                    errors.append(
                        {
                            "source_kind": source_kind,
                            "sheet": sheet_name,
                            "code": "MISSING_REQUIRED_COLUMNS",
                            "message": (
                                f"Source {source_path.name} does not contain required columns: "
                                f"{', '.join(missing_without_defaults)}."
                            ),
                        }
                    )
                    continue
                student_id_source_index = header_mapping.get("Student ID")
                student_id_target_index = target_headers.index("Student ID")
                first_name_target_index = target_headers.index("First Name") if "First Name" in target_headers else None
                last_name_target_index = target_headers.index("Last Name") if "Last Name" in target_headers else None
                sheet_lookup = template_student_lookup.get(
                    sheet_name,
                    {"ids": set(), "name_to_ids": {}, "id_to_profile": {}},
                )
                known_ids: set[str] = sheet_lookup["ids"]
                name_to_ids: dict[str, set[str]] = sheet_lookup["name_to_ids"]
                id_to_profile: dict[str, dict[str, str]] = sheet_lookup["id_to_profile"]
                school_target_index = target_headers.index("School") if "School" in target_headers else None
                school_year_target_index = target_headers.index("School Year") if "School Year" in target_headers else None
                for row in source_ws.iter_rows(min_row=2, values_only=True):
                    per_sheet[sheet_name]["rows_read"] += 1
                    row_values: list[object] = []
                    for target_header in target_headers:
                        source_index = header_mapping.get(target_header)
                        if source_index is not None and source_index < len(row):
                            value = row[source_index]
                        else:
                            value = template_defaults.get(target_header)
                            if clean_text(value) is not None:
                                per_sheet[sheet_name]["rows_filled_from_template_defaults"] += 1
                        row_values.append(value)
                    if all(clean_text(value) is None for value in row_values):
                        per_sheet[sheet_name]["rows_skipped_empty"] += 1
                        continue
                    student_id_value = (
                        row[student_id_source_index]
                        if student_id_source_index is not None and student_id_source_index < len(row)
                        else None
                    )
                    student_id_key = student_id_join_key(student_id_value)
                    if student_id_key is not None and student_id_key in id_to_profile:
                        profile = id_to_profile[student_id_key]
                        if first_name_target_index is not None and profile.get("first_name"):
                            current_first = clean_text(row_values[first_name_target_index])
                            if current_first != profile["first_name"]:
                                row_values[first_name_target_index] = profile["first_name"]
                                per_sheet[sheet_name]["rows_student_name_corrected_by_id"] += 1
                        if last_name_target_index is not None and profile.get("last_name"):
                            current_last = clean_text(row_values[last_name_target_index])
                            if current_last != profile["last_name"]:
                                row_values[last_name_target_index] = profile["last_name"]
                                per_sheet[sheet_name]["rows_student_name_corrected_by_id"] += 1
                    if student_id_key is None or (known_ids and student_id_key not in known_ids):
                        fallback_sid: str | None = None
                        if first_name_target_index is not None and last_name_target_index is not None:
                            first_name = clean_text(row_values[first_name_target_index])
                            last_name = clean_text(row_values[last_name_target_index])
                            if first_name is not None and last_name is not None:
                                name_key = f"{normalize_text(first_name)}|{normalize_text(last_name)}"
                                candidates = set(name_to_ids.get(name_key, set()))
                                if len(candidates) > 1:
                                    school_name = (
                                        clean_text(row_values[school_target_index])
                                        if school_target_index is not None
                                        else None
                                    )
                                    school_year = (
                                        clean_text(row_values[school_year_target_index])
                                        if school_year_target_index is not None
                                        else None
                                    )
                                    filtered: set[str] = set()
                                    for candidate_sid in candidates:
                                        candidate = id_to_profile.get(candidate_sid, {})
                                        school_ok = True
                                        year_ok = True
                                        if school_name and candidate.get("school"):
                                            school_ok = school_labels_equivalent(school_name, candidate.get("school"))
                                        if school_year and candidate.get("school_year"):
                                            year_ok = normalize_text(school_year) == normalize_text(candidate.get("school_year"))
                                        if school_ok and year_ok:
                                            filtered.add(candidate_sid)
                                    if filtered:
                                        candidates = filtered
                                if len(candidates) == 1:
                                    fallback_sid = next(iter(candidates))
                        if fallback_sid is not None:
                            row_values[student_id_target_index] = fallback_sid
                            student_id_key = fallback_sid
                            per_sheet[sheet_name]["rows_student_id_corrected_by_name"] += 1
                        elif student_id_key is not None:
                            per_sheet[sheet_name]["rows_unresolved_student_id_mismatch"] += 1
                    if student_id_key is None:
                        per_sheet[sheet_name]["rows_skipped_missing_student_id"] += 1
                        continue
                    row_values[student_id_target_index] = student_id_key
                    signature = tuple(row_values)
                    if signature in row_signatures[sheet_name]:
                        per_sheet[sheet_name]["rows_skipped_duplicate"] += 1
                        continue
                    dedupe_key = _bridge_row_dedupe_key(sheet_name, row_values, target_headers)
                    row_score = _bridge_row_score(row_values)
                    if dedupe_key is not None:
                        prior = row_best_by_key[sheet_name].get(dedupe_key)
                        if prior is not None:
                            if row_score > prior["score"]:
                                target_row = prior["row_index"]
                                for column_index, value in enumerate(row_values, start=1):
                                    target_ws.cell(row=target_row, column=column_index).value = value
                                row_best_by_key[sheet_name][dedupe_key]["score"] = row_score
                                row_signatures[sheet_name].add(signature)
                                per_sheet[sheet_name]["rows_replaced_by_richer_match"] += 1
                            else:
                                per_sheet[sheet_name]["rows_skipped_duplicate"] += 1
                            continue
                    next_row = row_cursor[sheet_name]
                    for column_index, value in enumerate(row_values, start=1):
                        target_ws.cell(row=next_row, column=column_index).value = value
                    row_signatures[sheet_name].add(signature)
                    if dedupe_key is not None:
                        row_best_by_key[sheet_name][dedupe_key] = {
                            "row_index": next_row,
                            "score": row_score,
                        }
                    row_cursor[sheet_name] += 1
                    per_sheet[sheet_name]["rows_written"] += 1
        finally:
            workbook.close()
    if write_output and output_path is None:
        raise ValueError("--output is required in write mode.")
    if write_output and not errors:
        assert output_path is not None
        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_workbook.save(output_path)
    target_workbook.close()
    payload = {
        "mode": "write" if write_output else "dry-run",
        "write_output": write_output,
        "non_goal_path": str(non_goal_path.resolve()),
        "progress_path": str(progress_path.resolve()),
        "template_path": str(template_path.resolve()),
        "output_path": str(output_path.resolve()) if write_output and output_path is not None and not errors else None,
        "structural_valid": len(errors) == 0,
        "error_count": len(errors),
        "errors": errors,
        "per_sheet": per_sheet,
        "unmapped_source_columns": unmapped_source_columns,
        "header_aliases_applied": alias_hits,
    }
    if report_json_path is not None:
        write_issue_json(report_json_path, payload)
    return payload


def _prepare_site_rows_for_automation(
    *,
    roster_lookup: dict[str, dict[str, str]],
    site_staff_list_path: Path,
    use_site_staff_filter: bool = True,
) -> dict[str, Any]:
    all_student_rows = list(roster_lookup.values())
    active_site_coordinators: list[str] | None = None
    active_site_schools: list[str] | None = None
    if use_site_staff_filter:
        active_site_coordinators = load_active_site_coordinator_display_names(site_staff_list_path)
        active_site_schools = load_active_site_school_names(site_staff_list_path)
        all_student_rows = filter_student_rows_by_active_site_schools(
            all_student_rows,
            active_site_schools,
        )
    site_rows = _group_student_rows_by_site(all_student_rows)
    all_case_managers = _case_manager_names_from_roster(roster_lookup)
    excluded = (
        [
            name
            for name in all_case_managers
            if not roster_coordinator_matches_active_site_staff(name, active_site_coordinators)
        ]
        if active_site_coordinators is not None
        else []
    )
    return {
        "site_rows": site_rows,
        "all_student_rows": all_student_rows,
        "excluded_by_site_staff": excluded,
    }


def _load_drive_json_file(
    drive_service: Any,
    *,
    parent_folder_id: str,
    filename: str = AUTOMATION_LEDGER_FILE_NAME,
) -> tuple[dict[str, Any], str | None]:
    query = (
        f"name = '{filename}' and '{parent_folder_id}' in parents and trashed = false "
        f"and mimeType = '{AUTOMATION_LEDGER_MIME_TYPE}'"
    )
    listing = drive_service.files().list(q=query, spaces="drive", fields="files(id,name)", pageSize=1).execute()
    files = listing.get("files") or []
    if not files:
        return {"runs_by_date": {}, "completed": {}}, None
    file_id = files[0]["id"]
    content = drive_service.files().get_media(fileId=file_id).execute()
    if isinstance(content, bytes):
        payload = json.loads(content.decode("utf-8") or "{}")
    elif isinstance(content, str):
        payload = json.loads(content or "{}")
    else:
        payload = content if isinstance(content, dict) else {}
    payload.setdefault("runs_by_date", {})
    payload.setdefault("completed", {})
    return payload, file_id


def _write_drive_json_file(
    drive_service: Any,
    *,
    parent_folder_id: str,
    payload: dict[str, Any],
    file_id: str | None,
    filename: str = AUTOMATION_LEDGER_FILE_NAME,
) -> str:
    try:
        media_module = importlib.import_module("googleapiclient.http")
        media_upload_cls = getattr(media_module, "MediaInMemoryUpload")
    except ImportError as exc:
        raise RuntimeError(
            "google-api-python-client is required for Drive ledger writes."
        ) from exc
    body = json.dumps(payload, indent=2, sort_keys=True).encode("utf-8")
    media = media_upload_cls(
        body,
        mimetype=AUTOMATION_LEDGER_MIME_TYPE,
        resumable=False,
    )
    if file_id:
        updated = drive_service.files().update(
            fileId=file_id,
            body={"mimeType": AUTOMATION_LEDGER_MIME_TYPE},
            media_body=media,
            fields="id",
        ).execute()
        return str(updated.get("id") or file_id)
    created = drive_service.files().create(
        body={"name": filename, "mimeType": AUTOMATION_LEDGER_MIME_TYPE, "parents": [parent_folder_id]},
        media_body=media,
        fields="id",
    ).execute()
    created_id = clean_text(created.get("id"))
    if created_id is None:
        raise RuntimeError("Drive ledger create did not return id.")
    return created_id


def _ensure_drive_folder(
    drive_service: Any,
    *,
    parent_id: str,
    folder_name: str,
) -> str:
    query = (
        f"name = '{folder_name}' and '{parent_id}' in parents and trashed = false "
        "and mimeType = 'application/vnd.google-apps.folder'"
    )
    listing = drive_service.files().list(q=query, spaces="drive", fields="files(id,name)", pageSize=1).execute()
    files = listing.get("files") or []
    if files:
        return str(files[0]["id"])
    created = drive_service.files().create(
        body={
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        },
        fields="id",
    ).execute()
    return str(created["id"])


def _upload_file_to_drive(
    drive_service: Any,
    *,
    parent_folder_id: str,
    local_path: Path,
    convert_to_google_sheet: bool = False,
) -> dict[str, Any]:
    metadata: dict[str, Any] = {"name": local_path.name, "parents": [parent_folder_id]}
    if convert_to_google_sheet:
        metadata["mimeType"] = "application/vnd.google-apps.spreadsheet"
    created = drive_service.files().create(
        body=metadata,
        media_body=str(local_path),
        fields="id,name,mimeType,webViewLink",
    ).execute()
    return {
        "id": created.get("id"),
        "name": created.get("name"),
        "mime_type": created.get("mimeType"),
        "web_view_link": created.get("webViewLink"),
    }


def _send_notification_email(
    gmail_service: Any,
    *,
    sender: str,
    recipient: str,
    site_name: str,
    grading_period: str,
    drive_file: dict[str, Any],
) -> str:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = f"QPR {quarter_label(grading_period)} ready for {site_name}"
    message.set_content(
        "\n".join(
            [
                f"A QPR workbook was submitted for {site_name}.",
                f"Grading period: {grading_period}",
                f"Open file: {drive_file.get('web_view_link') or drive_file.get('id')}",
                "",
                "Thank you,",
                "CISEPA Data Team",
            ]
        )
    )
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    sent = gmail_service.users().messages().send(userId="me", body={"raw": raw}).execute()
    message_id = clean_text(sent.get("id"))
    if message_id is None:
        raise RuntimeError("Gmail send did not return id.")
    return message_id


def _apps_script_attach_module():
    path = Path(__file__).resolve().parent / "apps_script_attach.py"
    spec = importlib.util.spec_from_file_location("qpr_apps_script_attach", path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load apps script attach module at {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_deadline_automation(
    *,
    as_of: date | None = None,
    output_directory: Path | None = None,
    template_path: Path = DEFAULT_TEMPLATE_PATH,
    deadlines_path: Path = DEFAULT_DEADLINES_PATH,
    roster_source_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    metric_workbook_path: Path | None = DEFAULT_METRIC_WORKBOOK_PATH,
    site_staff_list_path: Path = DEFAULT_SITE_STAFF_LIST_PATH,
    email_map_path: Path | None = None,
    synthetic_override_map_json_path: Path | None = None,
    drive_root_folder_id: str | None = None,
    max_sites_per_run: int | None = None,
    dry_run: bool = False,
    no_upload: bool = False,
    no_notify: bool = False,
    force: bool = False,
    allow_create_missing_script_id: bool = False,
    directory_email_lookup: bool = False,
) -> dict[str, Any]:
    current_date = as_of or date.today()
    template_schema, schedule_data, roster_lookup = _load_provision_inputs(
        template_path,
        deadlines_path,
        roster_source_path,
    )
    prepared = _prepare_site_rows_for_automation(
        roster_lookup=roster_lookup,
        site_staff_list_path=site_staff_list_path,
        use_site_staff_filter=True,
    )
    site_rows: dict[str, list[dict[str, str]]] = prepared["site_rows"]
    directory_resolution_warnings: list[dict[str, Any]] = []
    directory_resolver = _directory_email_resolver() if directory_email_lookup else None
    recipients_by_site = load_site_coordinator_recipient_emails(
        site_staff_list_path,
        email_map_path=email_map_path,
        synthetic_override_map_json_path=synthetic_override_map_json_path,
        directory_resolver=directory_resolver,
        directory_resolution_warnings=directory_resolution_warnings if directory_email_lookup else None,
    )
    drive_service = None
    ledger_payload: dict[str, Any] = {"runs_by_date": {}, "completed": {}}
    ledger_file_id: str | None = None
    ledger_write = bool(drive_root_folder_id and not dry_run)
    if ledger_write:
        required_scopes = _qpr_oauth_scopes(include_drive=True, include_gmail_send=False)
        creds = _oauth_credentials_from_env(required_scopes=required_scopes)
        drive_service = build("drive", "v3", credentials=creds)
        ledger_payload, ledger_file_id = _load_drive_json_file(
            drive_service,
            parent_folder_id=drive_root_folder_id,
        )
    ledger_completed = ledger_payload.setdefault("completed", {})
    workset = select_actionable_automation_workset(
        site_rows=site_rows,
        schedule_data=schedule_data,
        as_of=current_date,
        ledger_completed=ledger_completed,
        force=force,
    )
    eligible_pairs = list(workset.get("eligible_pairs", []))
    skipped_by_site_cap: list[dict[str, Any]] = []
    pairs_to_process = eligible_pairs
    if not dry_run and max_sites_per_run is not None and max_sites_per_run >= 0:
        pairs_to_process = eligible_pairs[:max_sites_per_run]
        skipped_by_site_cap = [
            {
                "site_name": pair["site_name"],
                "grading_period": pair["grading_period"],
                "reason": "max_sites_per_run_cap",
            }
            for pair in eligible_pairs[max_sites_per_run:]
        ]
    submit_menu_enabled = (
        (os.getenv("QPR_ATTACH_SUBMIT_MENU") or "").strip().lower() in {"1", "true", "yes", "on"}
        and not no_upload
        and not dry_run
    )
    script_service = None
    appscript_module = None
    results: list[dict[str, Any]] = []
    sites_generated = 0
    sites_uploaded = 0
    sites_notified = 0
    notifications_sent = 0
    sites_script_attached = 0
    for pair in pairs_to_process:
        site_name = pair["site_name"]
        grading_period = pair["grading_period"]
        key = _automation_record_key(grading_period=grading_period, site_name=site_name)
        quarter = quarter_label(grading_period)
        output_root = output_directory or default_batch_output_root(template_path)
        output_path = Path(output_root) / quarter / f"{filename_safe_label(site_name)}_{quarter}_QPR.xlsx"
        row: dict[str, Any] = {
            "site_name": site_name,
            "grading_period": grading_period,
            "status": "pending",
            "recipients": recipients_by_site.get(site_name, []),
            "output_path": str(output_path),
        }
        try:
            provisioned = write_provisioned_template(
                output_path=output_path,
                template_path=template_path,
                template_schema=template_schema,
                schedule_data=schedule_data,
                student_rows=site_rows.get(site_name, []),
                grading_period=grading_period,
                coordinator_name=None,
                school_name=site_name,
                resolved_schools=_schools_from_rows(site_rows.get(site_name, [])),
                metric_workbook_path=metric_workbook_path,
            )
            row.update(provisioned)
            sites_generated += 1
        except ValueError as exc:
            if str(exc) == NO_STUDENTS_AFTER_ENROLLMENT_FILTER:
                row["status"] = "skipped_no_students_after_enrollment_filter"
                results.append(row)
                continue
            raise
        if dry_run:
            row["status"] = "dry_run_preview"
            results.append(row)
            continue
        ledger_entry: dict[str, Any] = {
            "grading_period": grading_period,
            "site_name": site_name,
            "output_path": row["output_path"],
            "generated": True,
            "uploaded": False,
            "notified": False,
            "notifications": [],
            "status": "generated",
        }
        drive_file: dict[str, Any] | None = None
        if not no_upload:
            if drive_service is None or drive_root_folder_id is None:
                raise RuntimeError("drive_root_folder_id is required when uploads are enabled.")
            period_number = int(float(grading_period))
            parent_override = clean_text(os.getenv(f"QPR_DRIVE_Q{period_number}_FOLDER_ID"))
            if parent_override:
                parent_id = parent_override
            else:
                parent_id = _ensure_drive_folder(
                    drive_service,
                    parent_id=drive_root_folder_id,
                    folder_name=quarter,
                )
            site_folder = _ensure_drive_folder(
                drive_service,
                parent_id=parent_id,
                folder_name=filename_safe_label(site_name),
            )
            drive_file = _upload_file_to_drive(
                drive_service,
                parent_folder_id=site_folder,
                local_path=Path(row["output_path"]),
                convert_to_google_sheet=submit_menu_enabled,
            )
            ledger_entry["drive_file"] = drive_file
            ledger_entry["uploaded"] = True
            sites_uploaded += 1
            if submit_menu_enabled and drive_file.get("mime_type") == "application/vnd.google-apps.spreadsheet":
                if appscript_module is None:
                    appscript_module = _apps_script_attach_module()
                if script_service is None:
                    required_scopes = _qpr_oauth_scopes(
                        include_drive=True,
                        include_gmail_send=False,
                        include_apps_script_projects=True,
                    )
                    creds = _oauth_credentials_from_env(required_scopes=required_scopes)
                    script_service = build("script", "v1", credentials=creds)
                prior_script = appscript_module.parse_ledger_script_metadata(ledger_entry.get("apps_script"))
                attached = appscript_module.attach_submit_menu_script(
                    script_service,
                    spreadsheet_id=str(drive_file["id"]),
                    source_dir=Path(__file__).resolve().parent / "apps_script",
                    script_title=f"QPR Submit Menu - {site_name}",
                    existing_script_id=prior_script.get("script_id"),
                    require_existing_script_id=not allow_create_missing_script_id,
                )
                attached = appscript_module.update_timestamped_result(
                    attached,
                    updated_at=datetime.utcnow().isoformat(timespec="seconds"),
                )
                ledger_entry["apps_script"] = appscript_module.to_ledger_script_metadata(attached)
                row["apps_script"] = attached
                sites_script_attached += 1
        if no_notify:
            ledger_entry["status"] = "completed" if not no_upload else "generated"
            row["status"] = "ok"
        else:
            if drive_file is None:
                raise RuntimeError("notification requires uploaded file metadata.")
            sender = clean_text(os.getenv("QPR_MAIL_SENDER"))
            if sender is None:
                raise RuntimeError("QPR_MAIL_SENDER is required when notifications are enabled.")
            required_scopes = _qpr_oauth_scopes(include_drive=True, include_gmail_send=True)
            creds = _oauth_credentials_from_env(required_scopes=required_scopes)
            gmail_service = build("gmail", "v1", credentials=creds)
            notifications: list[dict[str, str]] = []
            for recipient in row["recipients"]:
                try:
                    message_id = _send_notification_email(
                        gmail_service,
                        sender=sender,
                        recipient=recipient,
                        site_name=site_name,
                        grading_period=grading_period,
                        drive_file=drive_file,
                    )
                except Exception as exc:  # pragma: no cover - exercised with mocks
                    ledger_entry["status"] = "uploaded_not_notified"
                    ledger_entry["notifications"] = notifications
                    ledger_entry["notified"] = False
                    ledger_completed[key] = ledger_entry
                    if ledger_write and drive_service is not None and drive_root_folder_id is not None:
                        ledger_file_id = _write_drive_json_file(
                            drive_service,
                            parent_folder_id=drive_root_folder_id,
                            payload=ledger_payload,
                            file_id=ledger_file_id,
                        )
                    raise RuntimeError(f"notification_failed:{site_name}:{exc}") from exc
                notifications.append({"recipient": recipient, "message_id": message_id})
                notifications_sent += 1
            ledger_entry["notifications"] = notifications
            ledger_entry["notified"] = True
            ledger_entry["status"] = "completed"
            row["status"] = "ok"
            sites_notified += 1
        ledger_completed[key] = ledger_entry
        if ledger_write and drive_service is not None and drive_root_folder_id is not None:
            ledger_file_id = _write_drive_json_file(
                drive_service,
                parent_folder_id=drive_root_folder_id,
                payload=ledger_payload,
                file_id=ledger_file_id,
            )
        results.append(row)
    if ledger_write and drive_service is not None and drive_root_folder_id is not None:
        _write_drive_json_file(
            drive_service,
            parent_folder_id=drive_root_folder_id,
            payload=ledger_payload,
            file_id=ledger_file_id,
        )
    return {
        "status": "dry_run" if dry_run else "ok",
        "as_of_date": current_date.isoformat(),
        "target_grading_period": workset.get("target_grading_period"),
        "eligible_pairs": len(eligible_pairs),
        "eligible_pairs_processed": len(pairs_to_process),
        "schedule_skipped": workset.get("schedule_skipped", []),
        "ledger_skipped": workset.get("ledger_skipped", []),
        "sites_skipped_duplicate": len(workset.get("ledger_skipped", [])),
        "skipped_by_site_cap": skipped_by_site_cap,
        "sites_generated": sites_generated,
        "sites_uploaded": sites_uploaded,
        "sites_notified": sites_notified,
        "notifications_sent": notifications_sent,
        "sites_script_attached": sites_script_attached,
        "submit_menu_enabled": submit_menu_enabled,
        "ledger_write": ledger_write,
        "results": results,
        "excluded_by_site_staff": prepared.get("excluded_by_site_staff", []),
        "directory_email_lookup": directory_email_lookup,
        "directory_resolution_warnings": directory_resolution_warnings,
    }


def run_deadline_notify_retry(
    *,
    as_of: date | None = None,
    site_staff_list_path: Path = DEFAULT_SITE_STAFF_LIST_PATH,
    email_map_path: Path | None = None,
    synthetic_override_map_json_path: Path | None = None,
    drive_root_folder_id: str,
    directory_email_lookup: bool = False,
) -> dict[str, Any]:
    current_date = as_of or date.today()
    required_scopes = _qpr_oauth_scopes(include_drive=True, include_gmail_send=True)
    creds = _oauth_credentials_from_env(required_scopes=required_scopes)
    drive_service = build("drive", "v3", credentials=creds)
    gmail_service = build("gmail", "v1", credentials=creds)
    ledger_payload, ledger_file_id = _load_drive_json_file(
        drive_service,
        parent_folder_id=drive_root_folder_id,
    )
    completed = ledger_payload.setdefault("completed", {})
    directory_resolution_warnings: list[dict[str, Any]] = []
    directory_resolver = _directory_email_resolver() if directory_email_lookup else None
    recipients_by_site = load_site_coordinator_recipient_emails(
        site_staff_list_path,
        email_map_path=email_map_path,
        synthetic_override_map_json_path=synthetic_override_map_json_path,
        directory_resolver=directory_resolver,
        directory_resolution_warnings=directory_resolution_warnings if directory_email_lookup else None,
    )
    sender = clean_text(os.getenv("QPR_MAIL_SENDER"))
    if sender is None:
        raise RuntimeError("QPR_MAIL_SENDER is required for notify retry.")
    sites_processed = 0
    sites_completed = 0
    sites_still_pending = 0
    notifications_sent = 0
    results: list[dict[str, Any]] = []
    for key in sorted(completed):
        entry = completed.get(key) or {}
        if entry.get("status") != "uploaded_not_notified":
            continue
        sites_processed += 1
        site_name = clean_text(entry.get("site_name")) or ""
        grading_period = clean_text(entry.get("grading_period")) or "1.0"
        prior = entry.get("notifications") if isinstance(entry.get("notifications"), list) else []
        prior_recipient_set = {clean_text(item.get("recipient")) for item in prior if isinstance(item, dict)}
        prior_recipient_set.discard(None)
        all_recipients = recipients_by_site.get(site_name, [])
        remaining = [recipient for recipient in all_recipients if recipient not in prior_recipient_set]
        newly_added: list[dict[str, str]] = []
        failed: list[str] = []
        for recipient in remaining:
            try:
                message_id = _send_notification_email(
                    gmail_service,
                    sender=sender,
                    recipient=recipient,
                    site_name=site_name,
                    grading_period=grading_period,
                    drive_file=entry.get("drive_file") or {},
                )
            except Exception:  # pragma: no cover
                failed.append(recipient)
                continue
            newly_added.append({"recipient": recipient, "message_id": message_id})
            notifications_sent += 1
        entry["notifications"] = list(prior) + newly_added
        if failed:
            entry["status"] = "uploaded_not_notified"
            entry["notified"] = False
            sites_still_pending += 1
        else:
            entry["status"] = "completed"
            entry["notified"] = True
            sites_completed += 1
        completed[key] = entry
        results.append(
            {
                "site_name": site_name,
                "grading_period": grading_period,
                "newly_notified": len(newly_added),
                "remaining_recipients": failed,
            }
        )
    if sites_processed:
        _write_drive_json_file(
            drive_service,
            parent_folder_id=drive_root_folder_id,
            payload=ledger_payload,
            file_id=ledger_file_id,
        )
    return {
        "status": "ok",
        "as_of_date": current_date.isoformat(),
        "sites_processed": sites_processed,
        "sites_completed": sites_completed,
        "sites_still_pending": sites_still_pending,
        "notifications_sent": notifications_sent,
        "results": results,
        "directory_email_lookup": directory_email_lookup,
        "directory_resolution_warnings": directory_resolution_warnings,
    }


def run_apps_script_redeploy(
    *,
    drive_root_folder_id: str,
    source_dir: Path,
    require_existing_script_id: bool = False,
    max_sites_per_run: int | None = None,
) -> dict[str, Any]:
    required_scopes = _qpr_oauth_scopes(
        include_drive=True,
        include_gmail_send=False,
        include_apps_script_projects=True,
    )
    creds = _oauth_credentials_from_env(required_scopes=required_scopes)
    drive_service = build("drive", "v3", credentials=creds)
    script_service = build("script", "v1", credentials=creds)
    ledger_payload, ledger_file_id = _load_drive_json_file(
        drive_service,
        parent_folder_id=drive_root_folder_id,
    )
    completed = ledger_payload.setdefault("completed", {})
    attach_mod = _apps_script_attach_module()
    keys = [key for key, entry in completed.items() if isinstance(entry, dict) and (entry.get("drive_file") or {}).get("id")]
    if max_sites_per_run is not None and max_sites_per_run >= 0:
        keys = keys[:max_sites_per_run]
    sites_updated = 0
    sites_failed = 0
    results: list[dict[str, Any]] = []
    for key in keys:
        entry = completed[key]
        spreadsheet_id = clean_text((entry.get("drive_file") or {}).get("id"))
        if spreadsheet_id is None:
            continue
        script_meta = attach_mod.parse_ledger_script_metadata(entry.get("apps_script"))
        try:
            attached = attach_mod.attach_submit_menu_script(
                script_service,
                spreadsheet_id=spreadsheet_id,
                source_dir=source_dir,
                script_title=f"QPR Submit Menu - {clean_text(entry.get('site_name')) or 'Site'}",
                existing_script_id=script_meta.get("script_id"),
                require_existing_script_id=require_existing_script_id,
            )
            attached = attach_mod.update_timestamped_result(
                attached,
                updated_at=datetime.utcnow().isoformat(timespec="seconds"),
            )
            entry["apps_script"] = attach_mod.to_ledger_script_metadata(attached)
            completed[key] = entry
            results.append({"key": key, "status": "ok", "apps_script": attached})
            sites_updated += 1
        except Exception as exc:  # pragma: no cover
            results.append({"key": key, "status": "error", "error": str(exc)})
            sites_failed += 1
    if keys:
        _write_drive_json_file(
            drive_service,
            parent_folder_id=drive_root_folder_id,
            payload=ledger_payload,
            file_id=ledger_file_id,
        )
    return {
        "status": "ok",
        "sites_processed": len(keys),
        "sites_updated": sites_updated,
        "sites_failed": sites_failed,
        "results": results,
    }


def pipeline_preflight(
    *,
    local_inputs_dir: Path | None = None,
    metric_refresh_triggered: bool = False,
    metric_refresh_succeeded: bool = False,
) -> dict[str, Any]:
    workbook_path, candidates, freshness = resolve_student_metrics_workbook(
        local_inputs_dir=local_inputs_dir
    )
    if workbook_path:
        status = "stale" if freshness.get("is_stale") else "ready"
    else:
        status = "blocked"
    refresh_status = "skipped"
    if metric_refresh_triggered:
        refresh_status = "ok" if metric_refresh_succeeded else "failed"
    return {
        "status": status,
        "metric_workbook_path": str(workbook_path) if workbook_path else freshness.get("destination_path"),
        "metric_workbook_age_hours": freshness.get("workbook_age_hours"),
        "metric_refresh_triggered": metric_refresh_triggered,
        "metric_refresh_status": refresh_status,
        "local_input": {
            "directory": str((local_inputs_dir or DEFAULT_QPR_LOCAL_INPUTS_DIR).resolve()),
            "workbook_path": str(workbook_path) if workbook_path else None,
            "candidate_workbooks": [str(path) for path in candidates],
            "freshness": freshness,
        },
        "missing_required_files": [] if workbook_path else ["Student Metrics workbook"],
    }


def pipeline_dry_run(**kwargs: Any) -> dict[str, Any]:
    return run_deadline_automation(dry_run=True, no_upload=True, no_notify=True, **kwargs)


def pipeline_upload_only(**kwargs: Any) -> dict[str, Any]:
    return run_deadline_automation(no_upload=False, no_notify=True, **kwargs)


def pipeline_upload_notify(**kwargs: Any) -> dict[str, Any]:
    return run_deadline_automation(no_upload=False, no_notify=False, **kwargs)


def pipeline_notify_retry(**kwargs: Any) -> dict[str, Any]:
    return run_deadline_notify_retry(**kwargs)


def paper_report_prepopulate(
    *,
    draft_path: Path,
    output_path: Path,
    grading_period: str,
    tier_i_path: Path | None,
    tier_ii_iii_path: Path | None,
    basic_needs_path: Path | None,
    referrals_path: Path | None,
    school_name: str | None = None,
    draft_pack_json_path: Path | None = None,
    use_recovered_artifacts: bool = False,
    school_needs_summary_path: Path | None = None,
    school_demographics_path: Path | None = None,
    school_improvement_path: Path | None = None,
    school_goals_progress_path: Path | None = None,
    student_metrics_path: Path | None = None,
    student_support_detail_path: Path | None = None,
    parent_guardian_consent_path: Path | None = None,
    support_summary_by_student_path: Path | None = None,
    goal_tracking_student_goals_path: Path | None = None,
    deadlines_path: Path = DEFAULT_DEADLINES_PATH,
    cisiphyus_root: Path = DEFAULT_CISIPHYUS_DIR,
) -> dict[str, Any]:
    grading_period = canonical_grading_period(grading_period)
    source_artifacts: dict[str, str] = {}
    if use_recovered_artifacts:
        tier_i_artifact = tier_i_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_TIER_I, cisiphyus_root=cisiphyus_root
        )
        tier_ii_iii_artifact = tier_ii_iii_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_TIER_II_III, cisiphyus_root=cisiphyus_root
        )
        basic_needs_artifact = basic_needs_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_BASIC_NEEDS_STUDENT, cisiphyus_root=cisiphyus_root
        )
        try:
            basic_needs_school_summary_artifact = latest_recovered_raw_path(
                report_id=RECOVERED_REPORT_BASIC_NEEDS_SCHOOL,
                cisiphyus_root=cisiphyus_root,
            )
        except ValueError:
            basic_needs_school_summary_artifact = None
        school_needs_artifact = school_needs_summary_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_SCHOOL_NEEDS, cisiphyus_root=cisiphyus_root
        )
        school_demographics_artifact = school_demographics_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_SCHOOL_DEMOGRAPHICS, cisiphyus_root=cisiphyus_root
        )
        school_improvement_artifact = school_improvement_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_SCHOOL_IMPROVEMENT, cisiphyus_root=cisiphyus_root
        )
        if school_goals_progress_path is not None:
            school_goals_artifact = school_goals_progress_path
        else:
            try:
                school_goals_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_SCHOOL_GOALS_PROGRESS,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                school_goals_artifact = None
        if student_metrics_path is not None:
            student_metrics_artifact = student_metrics_path
        else:
            try:
                student_metrics_artifact = latest_recovered_raw_path(
                    report_id="student_metrics_summary",
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                student_metrics_artifact = None
        student_support_detail_artifact = student_support_detail_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_STUDENT_SUPPORT_DETAIL, cisiphyus_root=cisiphyus_root
        )
        if parent_guardian_consent_path is not None:
            parent_guardian_consent_artifact = parent_guardian_consent_path
        else:
            try:
                parent_guardian_consent_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_PARENT_GUARDIAN_CONSENT,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                parent_guardian_consent_artifact = None
        if support_summary_by_student_path is not None:
            support_summary_by_student_artifact = support_summary_by_student_path
        else:
            try:
                support_summary_by_student_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_SUPPORT_SUMMARY_BY_STUDENT,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                support_summary_by_student_artifact = None
        if goal_tracking_student_goals_path is not None:
            goal_tracking_student_goals_artifact = goal_tracking_student_goals_path
        else:
            try:
                goal_tracking_student_goals_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_GOAL_TRACKING_STUDENT_GOALS,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                goal_tracking_student_goals_artifact = None
        source_artifacts.update(
            {
                "tier_i_supports": str(tier_i_artifact.resolve()),
                "tier_ii_iii_supports": str(tier_ii_iii_artifact.resolve()),
                "basic_needs": str(basic_needs_artifact.resolve()),
                "basic_needs_student_summary": str(basic_needs_artifact.resolve()),
                "school_needs_summary_export": str(school_needs_artifact.resolve()),
                "school_needs_assessment_school_demographics_export": str(school_demographics_artifact.resolve()),
                "school_needs_assessment_improvement_plan_community_data_export": str(
                    school_improvement_artifact.resolve()
                ),
                "student_support_detail": str(student_support_detail_artifact.resolve()),
            }
        )
        if school_goals_artifact is not None:
            source_artifacts["school_goals_progress_export"] = str(school_goals_artifact.resolve())
        if student_metrics_artifact is not None:
            source_artifacts["student_metrics_summary"] = str(student_metrics_artifact.resolve())
        if parent_guardian_consent_artifact is not None:
            source_artifacts["parent_guardian_consent"] = str(parent_guardian_consent_artifact.resolve())
        if support_summary_by_student_artifact is not None:
            source_artifacts["support_summary_by_student"] = str(support_summary_by_student_artifact.resolve())
        if goal_tracking_student_goals_artifact is not None:
            source_artifacts["goal_tracking_student_goals"] = str(goal_tracking_student_goals_artifact.resolve())
        if basic_needs_school_summary_artifact is not None:
            source_artifacts["basic_needs_school_summary"] = str(basic_needs_school_summary_artifact.resolve())
        tier_i_rows = load_recovered_tier_i_rows(
            workbook_path=tier_i_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        tier_ii_iii_rows = load_recovered_tier_ii_iii_rows(
            workbook_path=tier_ii_iii_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        basic_needs_rows = load_contract_export_rows(
            export_name="basic_needs",
            workbook_path=basic_needs_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_needs_rows = load_contract_export_rows(
            export_name="school_needs_summary_export",
            workbook_path=school_needs_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_demographics_rows = load_contract_export_rows(
            export_name="school_needs_assessment_school_demographics_export",
            workbook_path=school_demographics_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_improvement_rows = load_contract_export_rows(
            export_name="school_needs_assessment_improvement_plan_community_data_export",
            workbook_path=school_improvement_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_context_bullets = summarize_school_context_rows(
            school_needs_rows=school_needs_rows,
            school_demographics_rows=school_demographics_rows,
            school_improvement_rows=school_improvement_rows,
        )
        school_goals_rows = (
            load_contract_export_rows(
                export_name="school_goals_progress_export",
                workbook_path=school_goals_artifact,
                grading_period=grading_period,
                school_name=school_name,
            )
            if school_goals_artifact is not None
            else []
        )
        student_metrics_summary = (
            load_student_metrics_domain_summary(
                workbook_path=student_metrics_artifact,
                school_name=school_name,
            )
            if student_metrics_artifact is not None
            else {}
        )
        student_support_detail_rows = load_recovered_student_support_detail_rows(
            workbook_path=student_support_detail_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        parent_guardian_consent_rows = (
            load_contract_export_rows(
                export_name="parent_guardian_consent",
                workbook_path=parent_guardian_consent_artifact,
                grading_period=grading_period,
                school_name=school_name,
            )
            if parent_guardian_consent_artifact is not None
            else []
        )
        support_summary_by_student_rows = (
            load_recovered_support_summary_by_student_rows(
                workbook_path=support_summary_by_student_artifact,
                school_name=school_name,
            )
            if support_summary_by_student_artifact is not None
            else []
        )
        goal_tracking_student_goal_rows = (
            load_contract_export_rows(
                export_name="goal_tracking_student_goals",
                workbook_path=goal_tracking_student_goals_artifact,
                grading_period=grading_period,
                school_name=school_name,
            )
            if goal_tracking_student_goals_artifact is not None
            else []
        )
    else:
        if tier_i_path is None or tier_ii_iii_path is None or basic_needs_path is None:
            raise ValueError("tier_i_path, tier_ii_iii_path, and basic_needs_path are required.")
        tier_i_rows = load_contract_export_rows(
            export_name="tier_i_supports",
            workbook_path=tier_i_path,
            grading_period=grading_period,
            school_name=school_name,
        )
        tier_ii_iii_rows = load_contract_export_rows(
            export_name="tier_ii_iii_supports",
            workbook_path=tier_ii_iii_path,
            grading_period=grading_period,
            school_name=school_name,
        )
        basic_needs_rows = load_contract_export_rows(
            export_name="basic_needs",
            workbook_path=basic_needs_path,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_context_bullets = []
        school_goals_rows = (
            load_contract_export_rows(
                export_name="school_goals_progress_export",
                workbook_path=school_goals_progress_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if school_goals_progress_path is not None
            else []
        )
        student_metrics_summary = (
            load_student_metrics_domain_summary(
                workbook_path=student_metrics_path,
                school_name=school_name,
            )
            if student_metrics_path is not None
            else {}
        )
        student_support_detail_rows = (
            load_recovered_student_support_detail_rows(
                workbook_path=student_support_detail_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if student_support_detail_path is not None
            else None
        )
        parent_guardian_consent_rows = (
            load_contract_export_rows(
                export_name="parent_guardian_consent",
                workbook_path=parent_guardian_consent_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if parent_guardian_consent_path is not None
            else []
        )
        support_summary_by_student_rows = (
            load_recovered_support_summary_by_student_rows(
                workbook_path=support_summary_by_student_path,
                school_name=school_name,
            )
            if support_summary_by_student_path is not None
            else []
        )
        goal_tracking_student_goal_rows = (
            load_contract_export_rows(
                export_name="goal_tracking_student_goals",
                workbook_path=goal_tracking_student_goals_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if goal_tracking_student_goals_path is not None
            else []
        )
        source_artifacts.update(
            {
                "tier_i_supports": str(tier_i_path.resolve()),
                "tier_ii_iii_supports": str(tier_ii_iii_path.resolve()),
                "basic_needs": str(basic_needs_path.resolve()),
            }
        )
        if student_support_detail_path is not None:
            source_artifacts["student_support_detail"] = str(student_support_detail_path.resolve())
        if school_goals_progress_path is not None:
            source_artifacts["school_goals_progress_export"] = str(school_goals_progress_path.resolve())
        if student_metrics_path is not None:
            source_artifacts["student_metrics_summary"] = str(student_metrics_path.resolve())
        if parent_guardian_consent_path is not None:
            source_artifacts["parent_guardian_consent"] = str(parent_guardian_consent_path.resolve())
        if support_summary_by_student_path is not None:
            source_artifacts["support_summary_by_student"] = str(support_summary_by_student_path.resolve())
        if goal_tracking_student_goals_path is not None:
            source_artifacts["goal_tracking_student_goals"] = str(goal_tracking_student_goals_path.resolve())

    # why: written referral summary is intentionally manual; do not prefill.
    referral_rows = None

    draft_pack = build_qpr_draft_data_pack(
        grading_period=grading_period,
        tier_i_rows=tier_i_rows,
        tier_ii_iii_rows=tier_ii_iii_rows,
        basic_needs_rows=basic_needs_rows,
        school_context_bullets=school_context_bullets,
        referral_rows=referral_rows,
        student_support_detail_rows=student_support_detail_rows,
        school_goals_rows=school_goals_rows,
        student_metrics_summary=student_metrics_summary,
        parent_guardian_consent_rows=parent_guardian_consent_rows,
        support_summary_by_student_rows=support_summary_by_student_rows,
        goal_tracking_student_goal_rows=goal_tracking_student_goal_rows,
        source_artifacts=source_artifacts,
    )
    if draft_pack_json_path is not None:
        draft_pack_json_path.parent.mkdir(parents=True, exist_ok=True)
        draft_pack_json_path.write_text(json.dumps(draft_pack, indent=2), encoding="utf-8")

    write_result = prepopulate_paper_report_markdown(
        draft_path=draft_path,
        output_path=output_path,
        grading_period=grading_period,
        tier_i_summary=draft_pack["tier_i_summary"],
        tier_ii_summary=draft_pack["tier_ii_summary"],
        tier_iii_summary=draft_pack["tier_iii_summary"],
        basic_needs_summary=draft_pack["basic_needs_summary"],
        planned_vs_delivered_summary=draft_pack["planned_vs_delivered"],
        referral_summary=draft_pack["referral_summary"],
        school_context_bullets=draft_pack["school_context_bullets"],
    )
    return {
        "status": "ok",
        "grading_period": grading_period,
        "school_name": school_name,
        "source_counts": {
            "tier_i_rows": len(tier_i_rows),
            "tier_ii_iii_rows": len(tier_ii_iii_rows),
            "basic_needs_rows": len(basic_needs_rows),
            "referral_rows": len(referral_rows or []),
        },
        "summaries": {
            "tier_i": len(draft_pack["tier_i_summary"]),
            "tier_ii": len(draft_pack["tier_ii_summary"]),
            "tier_iii": len(draft_pack["tier_iii_summary"]),
            "planned_vs_delivered": len(draft_pack["planned_vs_delivered"]),
            "basic_needs_non_empty": sum(1 for item in draft_pack["basic_needs_summary"].values() if item["count"]),
            "school_context_bullets": len(draft_pack["school_context_bullets"]),
            "manual_required": len(draft_pack["manual_required"]),
        },
        "source_artifacts": source_artifacts,
        "manual_required": draft_pack["manual_required"],
        "draft_pack_json_path": str(draft_pack_json_path.resolve()) if draft_pack_json_path is not None else None,
        "output_path": write_result["output_path"],
    }


def paper_report_prepopulate_docx(
    *,
    output_docx_path: Path,
    grading_period: str,
    tier_i_path: Path | None,
    tier_ii_iii_path: Path | None,
    basic_needs_path: Path | None,
    referrals_path: Path | None,
    school_name: str | None = None,
    draft_pack_json_path: Path | None = None,
    use_recovered_artifacts: bool = False,
    school_needs_summary_path: Path | None = None,
    school_demographics_path: Path | None = None,
    school_improvement_path: Path | None = None,
    school_goals_progress_path: Path | None = None,
    student_metrics_path: Path | None = None,
    student_support_detail_path: Path | None = None,
    parent_guardian_consent_path: Path | None = None,
    support_summary_by_student_path: Path | None = None,
    goal_tracking_student_goals_path: Path | None = None,
    deadlines_path: Path = DEFAULT_DEADLINES_PATH,
    cisiphyus_root: Path = DEFAULT_CISIPHYUS_DIR,
    site_staff_list_path: Path = DEFAULT_SITE_STAFF_LIST_PATH,
    template_docx_path: Path = DEFAULT_WRITTEN_QPR_DRAFT_DOCX_PATH,
) -> dict[str, Any]:
    grading_period = canonical_grading_period(grading_period)
    source_artifacts: dict[str, str] = {}
    if use_recovered_artifacts:
        tier_i_artifact = tier_i_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_TIER_I, cisiphyus_root=cisiphyus_root
        )
        tier_ii_iii_artifact = tier_ii_iii_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_TIER_II_III, cisiphyus_root=cisiphyus_root
        )
        basic_needs_artifact = basic_needs_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_BASIC_NEEDS_STUDENT, cisiphyus_root=cisiphyus_root
        )
        try:
            basic_needs_school_summary_artifact = latest_recovered_raw_path(
                report_id=RECOVERED_REPORT_BASIC_NEEDS_SCHOOL,
                cisiphyus_root=cisiphyus_root,
            )
        except ValueError:
            basic_needs_school_summary_artifact = None
        school_needs_artifact = school_needs_summary_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_SCHOOL_NEEDS, cisiphyus_root=cisiphyus_root
        )
        school_demographics_artifact = school_demographics_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_SCHOOL_DEMOGRAPHICS, cisiphyus_root=cisiphyus_root
        )
        school_improvement_artifact = school_improvement_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_SCHOOL_IMPROVEMENT, cisiphyus_root=cisiphyus_root
        )
        if school_goals_progress_path is not None:
            school_goals_artifact = school_goals_progress_path
        else:
            try:
                school_goals_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_SCHOOL_GOALS_PROGRESS,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                school_goals_artifact = None
        if student_metrics_path is not None:
            student_metrics_artifact = student_metrics_path
        else:
            try:
                student_metrics_artifact = latest_recovered_raw_path(
                    report_id="student_metrics_summary",
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                student_metrics_artifact = None
        student_support_detail_artifact = student_support_detail_path or latest_recovered_raw_path(
            report_id=RECOVERED_REPORT_STUDENT_SUPPORT_DETAIL, cisiphyus_root=cisiphyus_root
        )
        if parent_guardian_consent_path is not None:
            parent_guardian_consent_artifact = parent_guardian_consent_path
        else:
            try:
                parent_guardian_consent_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_PARENT_GUARDIAN_CONSENT,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                parent_guardian_consent_artifact = None
        if support_summary_by_student_path is not None:
            support_summary_by_student_artifact = support_summary_by_student_path
        else:
            try:
                support_summary_by_student_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_SUPPORT_SUMMARY_BY_STUDENT,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                support_summary_by_student_artifact = None
        if goal_tracking_student_goals_path is not None:
            goal_tracking_student_goals_artifact = goal_tracking_student_goals_path
        else:
            try:
                goal_tracking_student_goals_artifact = latest_recovered_raw_path(
                    report_id=RECOVERED_REPORT_GOAL_TRACKING_STUDENT_GOALS,
                    cisiphyus_root=cisiphyus_root,
                )
            except ValueError:
                goal_tracking_student_goals_artifact = None
        source_artifacts.update(
            {
                "tier_i_supports": str(tier_i_artifact.resolve()),
                "tier_ii_iii_supports": str(tier_ii_iii_artifact.resolve()),
                "basic_needs": str(basic_needs_artifact.resolve()),
                "basic_needs_student_summary": str(basic_needs_artifact.resolve()),
                "school_needs_summary_export": str(school_needs_artifact.resolve()),
                "school_needs_assessment_school_demographics_export": str(school_demographics_artifact.resolve()),
                "school_needs_assessment_improvement_plan_community_data_export": str(
                    school_improvement_artifact.resolve()
                ),
                "student_support_detail": str(student_support_detail_artifact.resolve()),
            }
        )
        if school_goals_artifact is not None:
            source_artifacts["school_goals_progress_export"] = str(school_goals_artifact.resolve())
        if student_metrics_artifact is not None:
            source_artifacts["student_metrics_summary"] = str(student_metrics_artifact.resolve())
        if parent_guardian_consent_artifact is not None:
            source_artifacts["parent_guardian_consent"] = str(parent_guardian_consent_artifact.resolve())
        if support_summary_by_student_artifact is not None:
            source_artifacts["support_summary_by_student"] = str(support_summary_by_student_artifact.resolve())
        if goal_tracking_student_goals_artifact is not None:
            source_artifacts["goal_tracking_student_goals"] = str(goal_tracking_student_goals_artifact.resolve())
        if basic_needs_school_summary_artifact is not None:
            source_artifacts["basic_needs_school_summary"] = str(basic_needs_school_summary_artifact.resolve())
        tier_i_rows = load_recovered_tier_i_rows(
            workbook_path=tier_i_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        tier_ii_iii_rows = load_recovered_tier_ii_iii_rows(
            workbook_path=tier_ii_iii_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        basic_needs_rows = load_contract_export_rows(
            export_name="basic_needs",
            workbook_path=basic_needs_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_needs_rows = load_contract_export_rows(
            export_name="school_needs_summary_export",
            workbook_path=school_needs_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_demographics_rows = load_contract_export_rows(
            export_name="school_needs_assessment_school_demographics_export",
            workbook_path=school_demographics_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_improvement_rows = load_contract_export_rows(
            export_name="school_needs_assessment_improvement_plan_community_data_export",
            workbook_path=school_improvement_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_context_bullets = summarize_school_context_rows(
            school_needs_rows=school_needs_rows,
            school_demographics_rows=school_demographics_rows,
            school_improvement_rows=school_improvement_rows,
        )
        school_goals_rows = (
            load_contract_export_rows(
                export_name="school_goals_progress_export",
                workbook_path=school_goals_artifact,
                grading_period=grading_period,
                school_name=school_name,
            )
            if school_goals_artifact is not None
            else []
        )
        student_metrics_summary = (
            load_student_metrics_domain_summary(
                workbook_path=student_metrics_artifact,
                school_name=school_name,
            )
            if student_metrics_artifact is not None
            else {}
        )
        student_support_detail_rows = load_recovered_student_support_detail_rows(
            workbook_path=student_support_detail_artifact,
            grading_period=grading_period,
            school_name=school_name,
        )
        parent_guardian_consent_rows = (
            load_contract_export_rows(
                export_name="parent_guardian_consent",
                workbook_path=parent_guardian_consent_artifact,
                grading_period=grading_period,
                school_name=school_name,
            )
            if parent_guardian_consent_artifact is not None
            else []
        )
        support_summary_by_student_rows = (
            load_recovered_support_summary_by_student_rows(
                workbook_path=support_summary_by_student_artifact,
                school_name=school_name,
            )
            if support_summary_by_student_artifact is not None
            else []
        )
        goal_tracking_student_goal_rows = (
            load_contract_export_rows(
                export_name="goal_tracking_student_goals",
                workbook_path=goal_tracking_student_goals_artifact,
                grading_period=grading_period,
                school_name=school_name,
            )
            if goal_tracking_student_goals_artifact is not None
            else []
        )
    else:
        if tier_i_path is None or tier_ii_iii_path is None or basic_needs_path is None:
            raise ValueError("tier_i_path, tier_ii_iii_path, and basic_needs_path are required.")
        tier_i_rows = load_contract_export_rows(
            export_name="tier_i_supports",
            workbook_path=tier_i_path,
            grading_period=grading_period,
            school_name=school_name,
        )
        tier_ii_iii_rows = load_contract_export_rows(
            export_name="tier_ii_iii_supports",
            workbook_path=tier_ii_iii_path,
            grading_period=grading_period,
            school_name=school_name,
        )
        basic_needs_rows = load_contract_export_rows(
            export_name="basic_needs",
            workbook_path=basic_needs_path,
            grading_period=grading_period,
            school_name=school_name,
        )
        school_context_bullets = []
        school_goals_rows = (
            load_contract_export_rows(
                export_name="school_goals_progress_export",
                workbook_path=school_goals_progress_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if school_goals_progress_path is not None
            else []
        )
        student_metrics_summary = (
            load_student_metrics_domain_summary(
                workbook_path=student_metrics_path,
                school_name=school_name,
            )
            if student_metrics_path is not None
            else {}
        )
        student_support_detail_rows = (
            load_recovered_student_support_detail_rows(
                workbook_path=student_support_detail_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if student_support_detail_path is not None
            else None
        )
        parent_guardian_consent_rows = (
            load_contract_export_rows(
                export_name="parent_guardian_consent",
                workbook_path=parent_guardian_consent_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if parent_guardian_consent_path is not None
            else []
        )
        support_summary_by_student_rows = (
            load_recovered_support_summary_by_student_rows(
                workbook_path=support_summary_by_student_path,
                school_name=school_name,
            )
            if support_summary_by_student_path is not None
            else []
        )
        goal_tracking_student_goal_rows = (
            load_contract_export_rows(
                export_name="goal_tracking_student_goals",
                workbook_path=goal_tracking_student_goals_path,
                grading_period=grading_period,
                school_name=school_name,
            )
            if goal_tracking_student_goals_path is not None
            else []
        )
        source_artifacts.update(
            {
                "tier_i_supports": str(tier_i_path.resolve()),
                "tier_ii_iii_supports": str(tier_ii_iii_path.resolve()),
                "basic_needs": str(basic_needs_path.resolve()),
            }
        )
        if student_support_detail_path is not None:
            source_artifacts["student_support_detail"] = str(student_support_detail_path.resolve())
        if school_goals_progress_path is not None:
            source_artifacts["school_goals_progress_export"] = str(school_goals_progress_path.resolve())
        if student_metrics_path is not None:
            source_artifacts["student_metrics_summary"] = str(student_metrics_path.resolve())
        if parent_guardian_consent_path is not None:
            source_artifacts["parent_guardian_consent"] = str(parent_guardian_consent_path.resolve())
        if support_summary_by_student_path is not None:
            source_artifacts["support_summary_by_student"] = str(support_summary_by_student_path.resolve())
        if goal_tracking_student_goals_path is not None:
            source_artifacts["goal_tracking_student_goals"] = str(goal_tracking_student_goals_path.resolve())
    # why: written referral summary is intentionally manual; do not prefill.
    referral_rows = None
    caseload_buildup_summary, caseload_buildup_diagnostics = derive_caseload_buildup_from_sources(
        site_staff_list_path=site_staff_list_path,
        student_metrics_path=student_metrics_artifact if use_recovered_artifacts else student_metrics_path,
        deadlines_path=deadlines_path,
        school_name=school_name,
    )
    draft_pack = build_qpr_draft_data_pack(
        grading_period=grading_period,
        tier_i_rows=tier_i_rows,
        tier_ii_iii_rows=tier_ii_iii_rows,
        basic_needs_rows=basic_needs_rows,
        school_context_bullets=school_context_bullets,
        referral_rows=referral_rows,
        student_support_detail_rows=student_support_detail_rows,
        school_goals_rows=school_goals_rows,
        student_metrics_summary=student_metrics_summary,
        parent_guardian_consent_rows=parent_guardian_consent_rows,
        support_summary_by_student_rows=support_summary_by_student_rows,
        goal_tracking_student_goal_rows=goal_tracking_student_goal_rows,
        caseload_buildup_summary=caseload_buildup_summary,
        caseload_buildup_diagnostics=caseload_buildup_diagnostics,
        source_artifacts=source_artifacts,
    )
    if draft_pack_json_path is not None:
        draft_pack_json_path.parent.mkdir(parents=True, exist_ok=True)
        draft_pack_json_path.write_text(json.dumps(draft_pack, indent=2), encoding="utf-8")
    site_staff_context = _load_site_staff_context(
        site_staff_list_path=site_staff_list_path,
        school_name=school_name,
    )
    docx_result = _write_qpr_docx(
        output_path=output_docx_path,
        grading_period=canonical_grading_period(grading_period),
        draft_pack=draft_pack,
        site_staff_context=site_staff_context,
        template_docx_path=template_docx_path,
    )
    return {
        "status": "ok",
        "grading_period": grading_period,
        "school_name": school_name,
        "source_counts": {
            "tier_i_rows": len(tier_i_rows),
            "tier_ii_iii_rows": len(tier_ii_iii_rows),
            "basic_needs_rows": len(basic_needs_rows),
            "referral_rows": len(referral_rows or []),
        },
        "summaries": {
            "tier_i": len(draft_pack["tier_i_summary"]),
            "tier_ii": len(draft_pack["tier_ii_summary"]),
            "tier_iii": len(draft_pack["tier_iii_summary"]),
            "planned_vs_delivered": len(draft_pack["planned_vs_delivered"]),
            "basic_needs_non_empty": sum(1 for item in draft_pack["basic_needs_summary"].values() if item["count"]),
            "school_context_bullets": len(draft_pack["school_context_bullets"]),
            "manual_required": len(draft_pack["manual_required"]),
        },
        "source_artifacts": source_artifacts,
        "manual_required": draft_pack["manual_required"],
        "draft_pack_json_path": str(draft_pack_json_path.resolve()) if draft_pack_json_path is not None else None,
        "output_docx_path": docx_result["output_path"],
    }


def build_paper_prepopulate_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Prepopulate QPR paper-report markdown sections from CISDM exports and referral collection CSV.",
    )
    parser.add_argument("--draft-md", required=True, type=Path, help="Input markdown template for the paper report.")
    parser.add_argument("--output-md", required=True, type=Path, help="Output markdown path for populated report.")
    parser.add_argument("--grading-period", required=True, help="Canonical grading period (for example 1.0, 2.0).")
    parser.add_argument("--tier-i", required=False, type=Path, help="Tier I supports export workbook path.")
    parser.add_argument("--tier-ii-iii", required=False, type=Path, help="Tier II/III supports export workbook path.")
    parser.add_argument("--basic-needs", required=False, type=Path, help="Basic needs export workbook path.")
    parser.add_argument(
        "--referrals",
        required=False,
        type=Path,
        help="Quarterly referral collection CSV path (outside CISDM).",
    )
    parser.add_argument(
        "--use-recovered-artifacts",
        action="store_true",
        help="Load latest cisiphyus recovered artifacts instead of direct export files.",
    )
    parser.add_argument("--cisiphyus-root", type=Path, default=DEFAULT_CISIPHYUS_DIR)
    parser.add_argument("--school-needs-summary", type=Path, default=None)
    parser.add_argument("--school-demographics", type=Path, default=None)
    parser.add_argument("--school-improvement", type=Path, default=None)
    parser.add_argument("--school-goals-progress", type=Path, default=None)
    parser.add_argument("--student-metrics", type=Path, default=None)
    parser.add_argument("--student-support-detail", type=Path, default=None)
    parser.add_argument("--parent-guardian-consent", type=Path, default=None)
    parser.add_argument("--support-summary-by-student", type=Path, default=None)
    parser.add_argument("--goal-tracking-student-goals", type=Path, default=None)
    parser.add_argument("--deadlines", type=Path, default=DEFAULT_DEADLINES_PATH)
    parser.add_argument(
        "--draft-pack-json",
        type=Path,
        default=None,
        help="Optional JSON output path for generated draft data pack.",
    )
    parser.add_argument("--school-name", default=None, help="Optional school filter.")
    return parser


def build_paper_prepopulate_docx_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate formatted QPR written report docx from CISDM exports/recovered artifacts.",
    )
    parser.add_argument(
        "--output-docx",
        required=False,
        type=Path,
        default=None,
        help="Output docx path. If omitted, a sortable generated filename is used in --output-dir.",
    )
    parser.add_argument(
        "--output-dir",
        required=False,
        type=Path,
        default=DEFAULT_WRITTEN_QPR_OUTPUT_DIR,
        help="Directory used for generated filenames when --output-docx is omitted.",
    )
    parser.add_argument("--grading-period", required=True, help="Canonical grading period (for example 1.0, 2.0).")
    parser.add_argument("--tier-i", required=False, type=Path, help="Tier I supports workbook path.")
    parser.add_argument("--tier-ii-iii", required=False, type=Path, help="Tier II/III supports workbook path.")
    parser.add_argument("--basic-needs", required=False, type=Path, help="Basic needs workbook path.")
    parser.add_argument("--referrals", required=False, type=Path, help="Referral collection CSV path.")
    parser.add_argument("--use-recovered-artifacts", action="store_true")
    parser.add_argument("--cisiphyus-root", type=Path, default=DEFAULT_CISIPHYUS_DIR)
    parser.add_argument("--school-needs-summary", type=Path, default=None)
    parser.add_argument("--school-demographics", type=Path, default=None)
    parser.add_argument("--school-improvement", type=Path, default=None)
    parser.add_argument("--school-goals-progress", type=Path, default=None)
    parser.add_argument("--student-metrics", type=Path, default=None)
    parser.add_argument("--student-support-detail", type=Path, default=None)
    parser.add_argument("--parent-guardian-consent", type=Path, default=None)
    parser.add_argument("--support-summary-by-student", type=Path, default=None)
    parser.add_argument("--goal-tracking-student-goals", type=Path, default=None)
    parser.add_argument("--deadlines", type=Path, default=DEFAULT_DEADLINES_PATH)
    parser.add_argument("--draft-pack-json", type=Path, default=None)
    parser.add_argument("--school-name", default=None, help="Optional school filter.")
    parser.add_argument("--site-staff-list", type=Path, default=DEFAULT_SITE_STAFF_LIST_PATH)
    parser.add_argument("--template-docx", type=Path, default=DEFAULT_WRITTEN_QPR_DRAFT_DOCX_PATH)
    return parser


def build_referral_template_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create CSV template for quarterly referral collection.")
    parser.add_argument("--output", required=True, type=Path, help="Referral collection template output CSV path.")
    return parser


def build_bridge_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Bridge legacy QPR imports into a unified QPR workbook template.",
    )
    parser.add_argument("--non-goal", required=True, type=Path, dest="non_goal", help="Legacy non-goal workbook.")
    parser.add_argument("--progress", required=True, type=Path, help="Legacy progress monitoring workbook.")
    parser.add_argument("--template", required=True, type=Path, help="Unified QPR template workbook.")
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument("--dry-run", action="store_true", help="Validate and map without writing output workbook.")
    mode_group.add_argument("--write", action="store_true", help="Write bridged workbook output.")
    parser.add_argument("--output", type=Path, default=None, help="Output workbook path for --write mode.")
    parser.add_argument(
        "--report-json",
        type=Path,
        default=None,
        dest="report_json",
        help="Optional JSON report path for dry-run/write details.",
    )
    return parser


def build_validate_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Validate and safely repair a CISEPA progress monitoring import workbook.",
    )
    parser.add_argument("submission", type=Path, help="Submitted workbook path.")
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE_PATH,
        help="Reference template workbook path.",
    )
    parser.add_argument(
        "--deadlines",
        type=Path,
        default=DEFAULT_DEADLINES_PATH,
        help="Reporting deadlines workbook path.",
    )
    parser.add_argument(
        "--caseload",
        type=Path,
        default=DEFAULT_CASELOAD_PATH,
        help="Caseload export workbook path (lookup source when present).",
    )
    parser.add_argument(
        "--repaired",
        type=Path,
        default=None,
        help="Output path for repaired workbook (default: <stem>.repaired<suffix> next to submission).",
    )
    parser.add_argument(
        "--report-json",
        type=Path,
        default=None,
        dest="report_json",
        help="Write full validation payload JSON.",
    )
    parser.add_argument(
        "--report-csv",
        type=Path,
        default=None,
        dest="report_csv",
        help="Write issues CSV.",
    )
    return parser


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
        default=DEFAULT_SITE_STAFF_LIST_PATH,
        help=(
            f"Site/staff workbook containing sheet {SITE_LIST_SHEET!r}: Site Coordinators with a blank "
            f"{SITE_LIST_END_DATE_HEADER!r} are treated as current. Default: {DEFAULT_SITE_STAFF_LIST_PATH.name}."
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


def build_pipeline_parser(command_name: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=f"{command_name} automation command.")
    parser.add_argument("--as-of-date", default=None, help="ISO date override (YYYY-MM-DD).")
    parser.add_argument("--output-directory", type=Path, default=None, help="Output root for generated QPR files.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_PATH)
    parser.add_argument("--deadlines", type=Path, default=DEFAULT_DEADLINES_PATH)
    parser.add_argument(
        "--metric-workbook",
        type=Path,
        default=None,
        help=(
            "Student Metrics summary path. When omitted, uses tools/qpr/local_inputs "
            f"(default filename {DEFAULT_STUDENT_METRICS_FILENAME}) with cisiphyus refresh when stale."
        ),
    )
    parser.add_argument("--site-staff-list", type=Path, default=DEFAULT_SITE_STAFF_LIST_PATH)
    parser.add_argument("--email-map-json", type=Path, default=None)
    parser.add_argument("--synthetic-override-map-json", type=Path, default=None)
    parser.add_argument("--drive-root-folder-id", default=clean_text(os.getenv("QPR_DRIVE_ROOT_FOLDER_ID")))
    parser.add_argument("--max-sites-per-run", type=int, default=None)
    parser.add_argument("--no-upload", action="store_true")
    parser.add_argument("--no-notify", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--allow-create-missing-script-id", action="store_true")
    parser.add_argument(
        "--directory-email-lookup",
        action="store_true",
        default=_env_flag_enabled("QPR_DIRECTORY_EMAIL_LOOKUP"),
    )
    return parser


def build_notify_retry_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Retry pending QPR notifications from automation ledger.")
    parser.add_argument("--as-of-date", default=None, help="ISO date override (YYYY-MM-DD).")
    parser.add_argument("--site-staff-list", type=Path, default=DEFAULT_SITE_STAFF_LIST_PATH)
    parser.add_argument("--email-map-json", type=Path, default=None)
    parser.add_argument("--synthetic-override-map-json", type=Path, default=None)
    parser.add_argument("--drive-root-folder-id", default=clean_text(os.getenv("QPR_DRIVE_ROOT_FOLDER_ID")))
    parser.add_argument(
        "--directory-email-lookup",
        action="store_true",
        default=_env_flag_enabled("QPR_DIRECTORY_EMAIL_LOOKUP"),
    )
    parser.add_argument("--max-sites-per-run", type=int, default=None)
    return parser


def build_apps_script_redeploy_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Redeploy submit menu Apps Script to uploaded QPR spreadsheets.")
    parser.add_argument("--drive-root-folder-id", default=clean_text(os.getenv("QPR_DRIVE_ROOT_FOLDER_ID")))
    parser.add_argument(
        "--source-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "apps_script",
        help="Apps Script project source directory.",
    )
    parser.add_argument("--require-existing-script-id", action="store_true")
    parser.add_argument("--allow-create-missing-script-id", action="store_true")
    parser.add_argument("--max-sites-per-run", type=int, default=None)
    return parser


def build_oauth_bootstrap_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Bootstrap/update QPR OAuth refresh token payload.")
    parser.add_argument("--with-admin-directory-readonly", action="store_true")
    parser.add_argument("--with-apps-script-projects", action="store_true")
    return parser


def build_recipient_audit_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Resolve and export active-site QPR coordinator recipients.")
    parser.add_argument("--as-of-date", default=None, help="ISO date override (YYYY-MM-DD).")
    parser.add_argument("--site-staff-list", type=Path, default=DEFAULT_SITE_STAFF_LIST_PATH)
    parser.add_argument("--email-map-json", type=Path, default=None)
    parser.add_argument("--synthetic-override-map-json", type=Path, default=None)
    parser.add_argument("--output-directory", type=Path, default=DEFAULT_WRITTEN_QPR_OUTPUT_DIR)
    parser.add_argument(
        "--directory-email-lookup",
        action="store_true",
        default=_env_flag_enabled("QPR_DIRECTORY_EMAIL_LOOKUP"),
    )
    return parser


def build_qpr_email_ledger_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build one-email-per-site QPR notification ledger with template and QA outputs."
    )
    parser.add_argument("--as-of-date", default=None, help="ISO date override (YYYY-MM-DD).")
    parser.add_argument("--recipient-draft-csv", type=Path, required=True)
    parser.add_argument("--qpr-location-csv", type=Path, default=None)
    parser.add_argument("--existing-missing-sites-csv", type=Path, default=None)
    parser.add_argument("--output-directory", type=Path, default=DEFAULT_WRITTEN_QPR_OUTPUT_DIR)
    return parser


def _parse_as_of_date(value: str | None) -> date | None:
    if clean_text(value) is None:
        return None
    return datetime.strptime(value, "%Y-%m-%d").date()


def oauth_bootstrap(
    *,
    include_admin_directory_readonly: bool = False,
    include_apps_script_projects: bool = False,
) -> dict[str, Any]:
    token_path = DEFAULT_QPR_DIR / ".oauth-refresh-token.json"
    payload: dict[str, Any] = {}
    if token_path.exists():
        payload = json.loads(token_path.read_text(encoding="utf-8"))
    payload["scopes"] = _qpr_oauth_scopes(
        include_admin_directory_readonly=include_admin_directory_readonly,
        include_apps_script_projects=include_apps_script_projects,
    )
    token_path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return {"status": "ok", "token_path": str(token_path), "scopes": payload["scopes"]}


def _maybe_prefetch_student_metrics_for_cli(argv: list[str]):
    try:
        import qpr_cisiphyus_prefetch as prefetch
    except ModuleNotFoundError:
        script_dir = Path(__file__).resolve().parent
        if script_dir.name == "scripts":
            qpr_home = script_dir.parent
            if str(qpr_home) not in sys.path:
                sys.path.insert(0, str(qpr_home))
        import qpr_cisiphyus_prefetch as prefetch
    return prefetch.maybe_prefetch_student_metrics(argv)


def main(argv: list[str] | None = None) -> None:
    if argv is None:
        argv = sys.argv[1:]
    prefetch_result = None
    try:
        prefetch_result = _maybe_prefetch_student_metrics_for_cli(argv)
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        raise RuntimeError(f"cisiphyus_prefetch_failed: {exc}") from exc
    if argv and argv[0] == "paper-report-referral-template":
        args = build_referral_template_parser().parse_args(argv[1:])
        result = write_referral_collection_template(args.output)
    elif argv and argv[0] == "paper-report-prepopulate-docx":
        args = build_paper_prepopulate_docx_parser().parse_args(argv[1:])
        output_docx_path = args.output_docx
        draft_pack_json_path = args.draft_pack_json
        if output_docx_path is None:
            auto_docx, auto_pack = build_written_qpr_generation_paths(
                output_dir=args.output_dir,
                school_name=args.school_name,
                grading_period=args.grading_period,
            )
            output_docx_path = auto_docx
            if draft_pack_json_path is None:
                draft_pack_json_path = auto_pack
        result = paper_report_prepopulate_docx(
            output_docx_path=output_docx_path,
            grading_period=args.grading_period,
            tier_i_path=args.tier_i,
            tier_ii_iii_path=args.tier_ii_iii,
            basic_needs_path=args.basic_needs,
            referrals_path=args.referrals,
            school_name=args.school_name,
            draft_pack_json_path=draft_pack_json_path,
            use_recovered_artifacts=args.use_recovered_artifacts,
            school_needs_summary_path=args.school_needs_summary,
            school_demographics_path=args.school_demographics,
            school_improvement_path=args.school_improvement,
            school_goals_progress_path=args.school_goals_progress,
            student_metrics_path=args.student_metrics,
            student_support_detail_path=args.student_support_detail,
            parent_guardian_consent_path=args.parent_guardian_consent,
            support_summary_by_student_path=args.support_summary_by_student,
            goal_tracking_student_goals_path=args.goal_tracking_student_goals,
            deadlines_path=args.deadlines,
            cisiphyus_root=args.cisiphyus_root,
            site_staff_list_path=args.site_staff_list,
            template_docx_path=args.template_docx,
        )
    elif argv and argv[0] == "paper-report-prepopulate":
        args = build_paper_prepopulate_parser().parse_args(argv[1:])
        result = paper_report_prepopulate(
            draft_path=args.draft_md,
            output_path=args.output_md,
            grading_period=args.grading_period,
            tier_i_path=args.tier_i,
            tier_ii_iii_path=args.tier_ii_iii,
            basic_needs_path=args.basic_needs,
            referrals_path=args.referrals,
            school_name=args.school_name,
            draft_pack_json_path=args.draft_pack_json,
            use_recovered_artifacts=args.use_recovered_artifacts,
            school_needs_summary_path=args.school_needs_summary,
            school_demographics_path=args.school_demographics,
            school_improvement_path=args.school_improvement,
            school_goals_progress_path=args.school_goals_progress,
            student_metrics_path=args.student_metrics,
            student_support_detail_path=args.student_support_detail,
            parent_guardian_consent_path=args.parent_guardian_consent,
            support_summary_by_student_path=args.support_summary_by_student,
            goal_tracking_student_goals_path=args.goal_tracking_student_goals,
            deadlines_path=args.deadlines,
            cisiphyus_root=args.cisiphyus_root,
        )
    elif argv and argv[0] == "bridge-imports":
        args = build_bridge_parser().parse_args(argv[1:])
        result = bridge_legacy_imports_to_unified_qpr(
            non_goal_path=args.non_goal,
            progress_path=args.progress,
            template_path=args.template,
            output_path=args.output,
            write_output=args.write,
            report_json_path=args.report_json,
        )
    elif argv and argv[0] in ("provision", "provision-batch"):
        args = build_provision_parser().parse_args(argv[1:])
        roster_source_path = args.metric_workbook
        metric_workbook_path = None if args.no_metric_workbook else args.metric_workbook
        result = provision_all_site_templates(
            output_directory=args.output_directory,
            grading_period=args.grading_period,
            template_path=args.template,
            deadlines_path=args.deadlines,
            roster_source_path=roster_source_path,
            metric_workbook_path=metric_workbook_path,
            uat=args.uat,
            site_staff_list_path=args.site_staff_list,
            use_site_staff_filter=not args.no_site_staff_filter,
        )
    elif argv and argv[0] == "pipeline-preflight":
        build_pipeline_parser("pipeline-preflight").parse_args(argv[1:])
        result = pipeline_preflight(
            metric_refresh_triggered=bool(prefetch_result and prefetch_result.triggered),
            metric_refresh_succeeded=bool(prefetch_result and prefetch_result.succeeded),
        )
    elif argv and argv[0] == "pipeline-dry-run":
        args = build_pipeline_parser("pipeline-dry-run").parse_args(argv[1:])
        result = pipeline_dry_run(
            as_of=_parse_as_of_date(args.as_of_date),
            output_directory=args.output_directory,
            template_path=args.template,
            deadlines_path=args.deadlines,
            metric_workbook_path=_pipeline_metric_workbook_path(args.metric_workbook),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            drive_root_folder_id=args.drive_root_folder_id,
            max_sites_per_run=args.max_sites_per_run,
            force=args.force,
            allow_create_missing_script_id=args.allow_create_missing_script_id,
            directory_email_lookup=args.directory_email_lookup,
        )
    elif argv and argv[0] == "pipeline-upload-only":
        args = build_pipeline_parser("pipeline-upload-only").parse_args(argv[1:])
        result = pipeline_upload_only(
            as_of=_parse_as_of_date(args.as_of_date),
            output_directory=args.output_directory,
            template_path=args.template,
            deadlines_path=args.deadlines,
            metric_workbook_path=_pipeline_metric_workbook_path(args.metric_workbook),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            drive_root_folder_id=args.drive_root_folder_id,
            max_sites_per_run=args.max_sites_per_run,
            force=args.force,
            allow_create_missing_script_id=args.allow_create_missing_script_id,
            directory_email_lookup=args.directory_email_lookup,
        )
    elif argv and argv[0] == "pipeline-upload-notify":
        args = build_pipeline_parser("pipeline-upload-notify").parse_args(argv[1:])
        result = pipeline_upload_notify(
            as_of=_parse_as_of_date(args.as_of_date),
            output_directory=args.output_directory,
            template_path=args.template,
            deadlines_path=args.deadlines,
            metric_workbook_path=_pipeline_metric_workbook_path(args.metric_workbook),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            drive_root_folder_id=args.drive_root_folder_id,
            max_sites_per_run=args.max_sites_per_run,
            force=args.force,
            allow_create_missing_script_id=args.allow_create_missing_script_id,
            directory_email_lookup=args.directory_email_lookup,
        )
    elif argv and argv[0] == "pipeline-notify-retry":
        args = build_notify_retry_parser().parse_args(argv[1:])
        result = pipeline_notify_retry(
            as_of=_parse_as_of_date(args.as_of_date),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            drive_root_folder_id=args.drive_root_folder_id,
            directory_email_lookup=args.directory_email_lookup,
        )
    elif argv and argv[0] == "pipeline-recipient-audit":
        args = build_recipient_audit_parser().parse_args(argv[1:])
        result = pipeline_recipient_audit(
            as_of=_parse_as_of_date(args.as_of_date),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            output_directory=args.output_directory,
            directory_email_lookup=args.directory_email_lookup,
        )
    elif argv and argv[0] == "pipeline-qpr-email-ledger":
        args = build_qpr_email_ledger_parser().parse_args(argv[1:])
        result = pipeline_qpr_email_ledger(
            as_of=_parse_as_of_date(args.as_of_date),
            recipient_draft_csv_path=args.recipient_draft_csv,
            output_directory=args.output_directory,
            qpr_location_csv_path=args.qpr_location_csv,
            existing_missing_sites_csv_path=args.existing_missing_sites_csv,
        )
    elif argv and argv[0] == "export-coordinator-email-map":
        args = build_recipient_audit_parser().parse_args(argv[1:])
        result = pipeline_recipient_audit(
            as_of=_parse_as_of_date(args.as_of_date),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            output_directory=args.output_directory,
            directory_email_lookup=args.directory_email_lookup,
        )
    elif argv and argv[0] == "apps-script-redeploy":
        args = build_apps_script_redeploy_parser().parse_args(argv[1:])
        result = run_apps_script_redeploy(
            drive_root_folder_id=args.drive_root_folder_id,
            source_dir=args.source_dir,
            require_existing_script_id=(
                args.require_existing_script_id and not args.allow_create_missing_script_id
            ),
            max_sites_per_run=args.max_sites_per_run,
        )
    elif argv and argv[0] == "oauth-bootstrap":
        args = build_oauth_bootstrap_parser().parse_args(argv[1:])
        result = oauth_bootstrap(
            include_admin_directory_readonly=args.with_admin_directory_readonly,
            include_apps_script_projects=args.with_apps_script_projects,
        )
    elif argv and argv[0] == "automate":
        args = build_pipeline_parser("automate").parse_args(argv[1:])
        result = run_deadline_automation(
            as_of=_parse_as_of_date(args.as_of_date),
            output_directory=args.output_directory,
            template_path=args.template,
            deadlines_path=args.deadlines,
            metric_workbook_path=_pipeline_metric_workbook_path(args.metric_workbook),
            site_staff_list_path=args.site_staff_list,
            email_map_path=args.email_map_json,
            synthetic_override_map_json_path=args.synthetic_override_map_json,
            drive_root_folder_id=args.drive_root_folder_id,
            max_sites_per_run=args.max_sites_per_run,
            dry_run=args.dry_run,
            no_upload=args.no_upload,
            no_notify=args.no_notify,
            force=args.force,
            allow_create_missing_script_id=args.allow_create_missing_script_id,
            directory_email_lookup=args.directory_email_lookup,
        )
    else:
        args = build_validate_parser().parse_args(argv)
        result = validate_workbook(
            args.submission,
            repaired_path=args.repaired,
            report_json_path=args.report_json,
            report_csv_path=args.report_csv,
            template_path=args.template,
            deadlines_path=args.deadlines,
            caseload_path=args.caseload,
        )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
