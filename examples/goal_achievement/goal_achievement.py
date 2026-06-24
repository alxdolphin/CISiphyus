"""Shim: goal-achievement audit implementation lives in examples/audit/goal_achievement_audit.py."""

from __future__ import annotations

import sys
from pathlib import Path

_audit_dir = Path(__file__).resolve().parents[1] / "audit"
_example_dir = Path(__file__).resolve().parent
for path in (_audit_dir, _example_dir):
    inserted = str(path)
    if inserted not in sys.path:
        sys.path.insert(0, inserted)

import audit  # noqa: F401
import audit_sources  # noqa: F401
import goal_achievement_audit as _ga

AuditInputs = _ga.AuditInputs
build_parser = _ga.build_parser
evaluate_cross_audit = _ga.evaluate_cross_audit
load_drilldown_rows = _ga.load_drilldown_rows
run_inherited_gar_audit = _ga.run_inherited_gar_audit
run_inherited_metrics_audit = _ga.run_inherited_metrics_audit
resolve_inputs = _ga.resolve_inputs


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    school_year = (
        audit.normalize_school_year(args.school_year)
        if args.school_year
        else None
    )
    if (
        args.goal_progress_workbook
        and args.student_metrics_workbook
        and args.accreditation_workbook
    ):
        inputs = AuditInputs(
            goal_progress_workbook=args.goal_progress_workbook.expanduser().resolve(),
            student_metrics_workbook=args.student_metrics_workbook.expanduser().resolve(),
            accreditation_workbook=args.accreditation_workbook.expanduser().resolve(),
        )
    else:
        inputs = resolve_inputs(school_year=school_year, force_fetch=args.force_fetch)

    payload = _ga.run_goal_achievement_audit(
        inputs,
        school_year=school_year,
        school=args.school,
        metric=args.metric,
    )
    _ga.export_goal_achievement_results(payload, _ga.default_output_dir())
    return 0


__all__ = [
    "AuditInputs",
    "audit",
    "audit_sources",
    "build_parser",
    "evaluate_cross_audit",
    "load_drilldown_rows",
    "main",
    "resolve_inputs",
    "run_inherited_gar_audit",
    "run_inherited_metrics_audit",
]
