#!/usr/bin/env python3
"""Provision QPR import workbooks using bundled qpr_tools provision logic."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from qpr_cisiphyus_prefetch import maybe_prefetch_student_metrics, preferred_student_metrics_destination
from qpr_tools_loader import (
    REPO_ROOT,
    default_deadlines_path,
    default_site_staff_list_path,
    default_template_path,
    load_qpr_tools,
)


def resolve_metric_workbook(explicit: Path | None) -> Path:
    if explicit is not None:
        return explicit.expanduser().resolve()
    preferred = preferred_student_metrics_destination()
    result = maybe_prefetch_student_metrics(["provision-batch"], require_fresh=True)
    if result is not None and result.succeeded and result.destination is not None:
        return result.destination
    raise FileNotFoundError(
        "No student metrics workbook available after cisiphyus retrieval. "
        f"Expected refresh into {preferred}, or pass --metric-workbook to pin a local export."
    )


def resolve_template_path(explicit: Path | None) -> Path:
    if explicit is not None:
        return explicit.expanduser().resolve()
    path = default_template_path()
    if path.is_file():
        return path
    raise FileNotFoundError(
        f"No QPR template found at {path}. Provide --template with a local import workbook."
    )


def resolve_site_staff_filter(
    *,
    explicit: Path | None,
    no_site_staff_filter: bool,
) -> tuple[Path | None, bool]:
    if no_site_staff_filter:
        return None, False
    if explicit is not None:
        return explicit.expanduser().resolve(), True
    bundled = default_site_staff_list_path()
    if bundled.is_file():
        return bundled, True
    return None, False


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Provision progress-monitoring import workbooks (.xlsx) for each site in the metrics roster."
        ),
    )
    parser.add_argument(
        "output_directory",
        nargs="?",
        type=Path,
        help="Directory where tailored workbooks are written (default: ./examples/qpr/output).",
    )
    parser.add_argument("--grading-period", required=True, help="Canonical grading period, e.g. 2.0.")
    parser.add_argument(
        "--school-name",
        help="Provision a single school/site instead of all sites in the metrics roster.",
    )
    parser.add_argument(
        "--coordinator-name",
        help="Provision a single site coordinator instead of all sites.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        help="QPR import template workbook (default: examples/qpr/fixtures/qpr_import_template.xlsx).",
    )
    parser.add_argument(
        "--deadlines",
        type=Path,
        help="Reporting deadlines workbook (default: examples/qpr/fixtures/reporting_deadlines_minimal.xlsx).",
    )
    parser.add_argument(
        "--metric-workbook",
        type=Path,
        help="Student metrics summary (roster + metric prefill). When omitted, resolves via cisiphyus (age-gated).",
    )
    parser.add_argument(
        "--no-metric-workbook",
        action="store_true",
        help="Skip metric prefill/notes; roster still comes from --metric-workbook.",
    )
    parser.add_argument(
        "--no-site-staff-filter",
        action="store_true",
        help="Do not restrict provision to active Site List coordinators.",
    )
    parser.add_argument(
        "--site-staff-list",
        type=Path,
        help="Site/staff workbook for active-site filtering (default: bundled fixture when present).",
    )
    parser.add_argument(
        "--uat",
        action="store_true",
        help="Write UAT/training templates (EntityID/Client ID + Test last names).",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print the provision result JSON to stdout.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        qpr = load_qpr_tools()
        metric_workbook = resolve_metric_workbook(args.metric_workbook)
        template_path = resolve_template_path(args.template)
        deadlines_path = (args.deadlines or default_deadlines_path()).expanduser().resolve()
        output_directory = (args.output_directory or (REPO_ROOT / "examples" / "qpr" / "output")).resolve()
        metric_prefill = None if args.no_metric_workbook else metric_workbook
        site_staff_list, use_site_staff_filter = resolve_site_staff_filter(
            explicit=args.site_staff_list,
            no_site_staff_filter=args.no_site_staff_filter,
        )

        if args.school_name or args.coordinator_name:
            if args.school_name and args.coordinator_name:
                print("error: provide only one of --school-name or --coordinator-name", file=sys.stderr)
                return 1
            output_path = output_directory / f"Q{int(float(args.grading_period))}_provisioned_QPR.xlsx"
            result = qpr.provision_site_template(
                output_path=output_path,
                grading_period=args.grading_period,
                coordinator_name=args.coordinator_name,
                school_name=args.school_name,
                template_path=template_path,
                deadlines_path=deadlines_path,
                roster_source_path=metric_workbook,
                metric_workbook_path=metric_prefill,
                uat=args.uat,
                site_staff_list_path=site_staff_list,
                use_site_staff_filter=use_site_staff_filter,
            )
        else:
            result = qpr.provision_all_site_templates(
                output_directory=output_directory,
                grading_period=args.grading_period,
                template_path=template_path,
                deadlines_path=deadlines_path,
                roster_source_path=metric_workbook,
                metric_workbook_path=metric_prefill,
                uat=args.uat,
                site_staff_list_path=site_staff_list,
                use_site_staff_filter=use_site_staff_filter,
            )
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True))
    elif "results" in result:
        print(f"Output directory: {result['output_directory']}")
        print(f"Sites generated: {result['sites_generated']} / {result['sites_processed']}")
        print(f"Total students: {result['total_students']}")
        for item in result.get("results", []):
            print(f"  {item['output_path']} ({item['student_count']} students)")
        for skip in result.get("skipped", []):
            print(f"  skipped {skip['site_name']}: {skip['reason']}")
    else:
        print(f"Output: {result['output_path']} ({result['student_count']} students)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
