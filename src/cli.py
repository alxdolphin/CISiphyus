#!/usr/bin/env python3
# cisiphyus: chrome-profile cisdm export retriever (report fetch + bootstrap)
# usage:
#   cisiphyus student_metrics_summary
#   cisiphyus pull student_metrics_summary --school-year SY24-25
#   cisiphyus --bootstrap
#   cisiphyus audit accreditation
#   cisiphyus audit metrics
#   cisiphyus trend SY25-26
#   cisiphyus trend cross-year
#   cisiphyus qpr --grading-period 2.0 ...
# required: config/reports.yaml, config/export_urls.env
# outputs: artifacts/latest/<report_id>/raw.xlsx

from __future__ import annotations

import argparse
import sys

import config
import example_cli


def __getattr__(name: str):
    # WHY: keep example dispatch import-light; tests introspect bootstrap via cli
    if name in {"dismiss_blocking_dialogs", "verify_cisdm_auth", "classify_auth_probe"}:
        import bootstrap

        return getattr(bootstrap, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def print_summary(result) -> None:
    if result.status != "success":
        latest_dir = config.latest_report_dir(result.report_id)
        result_path = latest_dir / "result.json"
        diag_path = latest_dir / "diag.json"
        print(f"FAILED: {result.error}", file=sys.stderr)
        print(f"See {result_path}", file=sys.stderr)
        if diag_path.exists():
            print(f"Diagnostics: {diag_path}", file=sys.stderr)
        sys.exit(1)


def _build_retrieval_parser(*, prog: str | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Retrieve CISDM exports using a Chrome work profile.",
        prog=prog,
    )
    parser.add_argument("report_id", nargs="?", default="student_metrics_summary")
    parser.add_argument(
        "--school-year",
        default=None,
        metavar="SYxx-yy",
        help="School year for year-scoped direct exports (e.g. SY24-25).",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Show the browser window during retrieval (default: headless).",
    )
    parser.add_argument(
        "--bootstrap",
        action="store_true",
        help="Open headed Chrome for manual login, verify CISDM auth, and save bootstrap marker.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Replace an existing bootstrap profile under config/chrome-user-data/.",
    )
    parser.add_argument(
        "--verify-report",
        default=None,
        metavar="REPORT_ID",
        help="Optional: probe a report URL during bootstrap (e.g. accreditation).",
    )
    return parser


def _run_retrieval(argv: list[str], *, prog: str | None = None) -> None:
    from playwright.sync_api import sync_playwright

    import bootstrap
    import report

    args = _build_retrieval_parser(prog=prog).parse_args(argv)
    chrome_user_data_dir = config.default_chrome_user_data_dir()
    chrome_profile_directory = config.DEFAULT_CHROME_PROFILE_DIRECTORY
    verify_report = (args.verify_report or "").strip() or None

    if args.bootstrap:
        result = bootstrap.bootstrap_profile(
            chrome_user_data_dir,
            chrome_profile_directory,
            force=args.force,
            verify_report=verify_report,
            playwright_factory=sync_playwright,
        )
    else:
        result = report.run_report(
            report_id=args.report_id,
            headed=args.headed,
            chrome_user_data_dir=chrome_user_data_dir,
            chrome_profile_directory=chrome_profile_directory,
            school_year=(args.school_year or "").strip() or None,
        )

    print_summary(result)


def _maybe_run_pull(argv: list[str]) -> bool:
    # WHY: monitor runs via `audit accreditation`; raw export uses `pull <id>`
    if len(argv) < 3 or argv[1] != "pull":
        return False
    _run_retrieval(argv[2:], prog="cisiphyus pull")
    return True


def main() -> None:
    config.migrate_legacy_layout()

    if _maybe_run_pull(sys.argv):
        return

    example_code = example_cli.maybe_run_example_app(sys.argv)
    if example_code is not None:
        raise SystemExit(example_code)

    _run_retrieval(sys.argv[1:])


if __name__ == "__main__":
    main()
