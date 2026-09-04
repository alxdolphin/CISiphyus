#!/usr/bin/env python3
"""
Provision QPR import workbooks using bundled qpr_tools provision logic.

Run:
  cisiphyus qpr --grading-period 2.0

Single site:
  cisiphyus qpr --grading-period 2.0 --school-name "Lincoln HS"
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from types import ModuleType
from typing import Any, Callable, Sequence

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[1]
FIXTURES_DIR = SCRIPT_DIR / "fixtures"

def qpr_tools_path() -> Path:
    return SCRIPT_DIR / "qpr_tools.py"


def load_qpr_tools() -> ModuleType:
    path = qpr_tools_path()
    if not path.is_file():
        raise FileNotFoundError(
            f"qpr_tools.py not found at {path}. "
            "The bundled QPR tools module is missing from examples/qpr/."
        )
    module_name = "qpr_tools_bundled"
    if module_name in sys.modules:
        return sys.modules[module_name]
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load qpr_tools from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def default_template_path() -> Path:
    return FIXTURES_DIR / "qpr_import_template.xlsx"


def default_deadlines_path() -> Path:
    return FIXTURES_DIR / "reporting_deadlines_minimal.xlsx"


def default_site_staff_list_path() -> Path:
    return FIXTURES_DIR / "site_staff_list_minimal.xlsx"

# --- cisiphyus prefetch (live export when --metric-workbook is omitted) ---

TRUE_VALUES = frozenset({"1", "true", "yes", "on"})
STUDENT_METRICS_MAX_AGE_HOURS = 24.0
FALLBACK_STUDENT_METRICS_FILENAME = "StudentMetricsSummary.xlsx"


@dataclass(frozen=True)
class PrefetchResult:
    destination: Path | None
    triggered: bool
    succeeded: bool
    reason: str


def _flag_true(raw: str | None) -> bool:
    return (raw or "").strip().lower() in TRUE_VALUES


def _default_cisiphyus_root() -> Path:
    # WHY: prototypical app lives at examples/qpr/ inside the cisiphyus repo
    return Path(__file__).resolve().parents[2]


def current_school_year() -> str | None:
    src_dir = _default_cisiphyus_root() / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))
    try:
        import config

        return config.default_school_year(config.load_school_year_programs())
    except (FileNotFoundError, ValueError, OSError):
        return None


def default_student_metrics_filename() -> str:
    school_year = current_school_year()
    if school_year:
        return f"{school_year}_StudentMetricsSummary.xlsx"
    return FALLBACK_STUDENT_METRICS_FILENAME


def _default_local_inputs_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "artifacts" / "qpr"


def default_output_dir() -> Path:
    explicit = (os.environ.get("QPR_OUTPUT_DIR") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    return Path(__file__).resolve().parents[2] / "artifacts" / "qpr"


def preferred_student_metrics_destination(
    *,
    local_inputs_dir: Path | None = None,
) -> Path:
    explicit = (os.environ.get("QPR_STUDENT_METRICS_WORKBOOK") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    base = Path(
        os.environ.get("QPR_LOCAL_INPUTS_DIR", "").strip()
        or str(local_inputs_dir or _default_local_inputs_dir())
    ).expanduser().resolve()
    name = (
        os.environ.get("QPR_STUDENT_METRICS_FILENAME", "").strip()
        or default_student_metrics_filename()
    )
    return (base / name).resolve()


def metric_workbook_freshness(workbook_path: Path | None) -> dict[str, Any]:
    freshness: dict[str, Any] = {
        "max_age_hours": STUDENT_METRICS_MAX_AGE_HOURS,
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
            "is_stale": age_hours > STUDENT_METRICS_MAX_AGE_HOURS,
        }
    )
    return freshness


_PIPELINE_METRICS_COMMANDS = frozenset(
    {
        "pipeline-preflight",
        "pipeline-dry-run",
        "pipeline-upload-only",
        "pipeline-upload-notify",
    }
)


def argv_needs_student_metrics(argv: Sequence[str]) -> bool:
    for token in argv:
        if not isinstance(token, str):
            continue
        if token in {"automate", "provision-batch", *_PIPELINE_METRICS_COMMANDS}:
            return True
    return False


def argv_requires_fresh_metrics(argv: Sequence[str]) -> bool:
    if not argv_needs_student_metrics(argv):
        return False
    for token in argv:
        if not isinstance(token, str):
            continue
        if token in {
            "pipeline-upload-only",
            "pipeline-upload-notify",
            "automate",
            "provision-batch",
        }:
            return True
    return False


def _fetch_reason(
    *,
    destination: Path,
    force_fetch: bool,
) -> tuple[bool, str]:
    if force_fetch:
        return True, "forced"
    freshness = metric_workbook_freshness(destination if destination.exists() else None)
    if not freshness.get("workbook_exists"):
        return True, "missing"
    if freshness.get("is_stale"):
        return True, "stale"
    return False, "fresh"


def _run_cisiphyus_export(
    *,
    cisiphyus_root: Path,
    run: Callable[..., Any],
) -> Path:
    run_py = cisiphyus_root / "run.py"
    if not run_py.is_file():
        raise FileNotFoundError(f"cisiphyus run.py not found at {run_py}")

    cmd = _cisiphyus_cmd(cisiphyus_root)
    completed = run(
        cmd,
        cwd=str(cisiphyus_root),
        check=False,
        env=os.environ.copy(),
    )
    if getattr(completed, "returncode", 1) != 0:
        raise RuntimeError(
            f"cisiphyus student_metrics_summary failed with exit code "
            f"{getattr(completed, 'returncode', 'unknown')}"
        )

    raw = cisiphyus_root / "artifacts" / "latest" / "student_metrics_summary" / "raw.xlsx"
    if not raw.is_file():
        raise FileNotFoundError(f"expected cisiphyus output missing: {raw}")
    return raw


def _cisiphyus_cmd(cisiphyus_root: Path) -> list[str]:
    # WHY: auth comes from the bootstrapped chrome profile (cisiphyus --bootstrap)
    run_py = cisiphyus_root / "run.py"
    cmd: list[str] = [sys.executable, str(run_py), "student_metrics_summary"]
    headed = _flag_true(
        os.environ.get("QPR_CISPHYUS_HEADED") or os.environ.get("CISPHYUS_HEADED")
    )
    if headed:
        cmd.append("--headed")
    return cmd


def _enforce_fresh_after_refresh(
    destination: Path,
    *,
    require_fresh: bool,
    allow_stale: bool,
) -> None:
    if not require_fresh or allow_stale:
        return
    freshness = metric_workbook_freshness(destination if destination.exists() else None)
    if freshness.get("workbook_exists") and not freshness.get("is_stale"):
        return
    age = freshness.get("workbook_age_hours")
    raise RuntimeError(
        "stale_student_metrics_after_refresh: "
        f"path={destination} exists={freshness.get('workbook_exists')} "
        f"workbook_age_hours={age} max_age_hours={STUDENT_METRICS_MAX_AGE_HOURS}"
    )


def maybe_prefetch_student_metrics(
    qpr_tools_argv: Sequence[str],
    *,
    require_fresh: bool | None = None,
    run: Callable[..., Any] = subprocess.run,
) -> PrefetchResult | None:
    """refresh student metrics via cisiphyus when missing, stale (>24h), or forced."""
    if not argv_needs_student_metrics(qpr_tools_argv):
        return None

    if require_fresh is None:
        require_fresh = argv_requires_fresh_metrics(qpr_tools_argv)

    destination = preferred_student_metrics_destination()
    force_fetch = _flag_true(os.environ.get("QPR_FETCH_STUDENT_METRICS"))
    allow_stale = _flag_true(os.environ.get("QPR_ALLOW_STALE_METRICS"))
    should_fetch, reason = _fetch_reason(destination=destination, force_fetch=force_fetch)

    if not should_fetch:
        print(
            f"[qpr] metrics fresh at {destination} "
            f"(max_age_hours={STUDENT_METRICS_MAX_AGE_HOURS}), skipping fetch",
            file=sys.stderr,
        )
        resolved = destination if destination.exists() else None
        return PrefetchResult(
            destination=resolved,
            triggered=False,
            succeeded=True,
            reason="fresh",
        )

    cisiphyus_root = Path(
        os.environ.get("QPR_CISPHYUS_ROOT", "").strip()
        or os.environ.get("CISIPHYUS_ROOT", "").strip()
        or str(_default_cisiphyus_root())
    ).expanduser().resolve()

    try:
        raw = _run_cisiphyus_export(cisiphyus_root=cisiphyus_root, run=run)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(raw, destination)
        _enforce_fresh_after_refresh(
            destination,
            require_fresh=require_fresh,
            allow_stale=allow_stale,
        )
        print(
            f"[qpr] refreshed metrics ({reason}) -> {destination}",
            file=sys.stderr,
        )
        return PrefetchResult(
            destination=destination,
            triggered=True,
            succeeded=True,
            reason=reason,
        )
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        _enforce_fresh_after_refresh(
            destination,
            require_fresh=require_fresh,
            allow_stale=allow_stale,
        )
        print(f"[qpr] refresh failed ({reason}): {exc}", file=sys.stderr)
        resolved = destination if destination.exists() else None
        return PrefetchResult(
            destination=resolved,
            triggered=True,
            succeeded=False,
            reason=reason,
        )


# --- provision CLI ---

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
    # WHY: bundled minimal fixture is for tests only; live metrics use different school names
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
        help="Directory where tailored workbooks are written (default: artifacts/qpr/).",
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
        help=(
            "Site/staff workbook (Site List sheet): restrict provision to active schools. "
            "Omit to provision every site in the metrics roster."
        ),
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
        output_directory = (args.output_directory or default_output_dir()).resolve()
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
        if result.get("warning"):
            print(f"warning: {result['warning']}", file=sys.stderr)
        for item in result.get("results", []):
            print(f"  {item['output_path']} ({item['student_count']} students)")
        for skip in result.get("skipped", []):
            print(f"  skipped {skip['site_name']}: {skip['reason']}")
    else:
        print(f"Output: {result['output_path']} ({result['student_count']} students)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
