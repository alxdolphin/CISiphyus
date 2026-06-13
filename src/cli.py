#!/usr/bin/env python3
# cisiphyus: chrome-profile cisdm export retriever (report fetch + bootstrap)
# usage:
#   python run.py student_metrics_summary
#   python run.py --bootstrap
#   cisiphyus accreditation --output-dir ...
#   cisiphyus qpr <output_dir> --grading-period 2.0 ...
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


def main() -> None:
    config.migrate_legacy_layout()

    example_code = example_cli.maybe_run_example_app(sys.argv)
    if example_code is not None:
        raise SystemExit(example_code)

    from playwright.sync_api import sync_playwright

    import bootstrap
    import report

    parser = argparse.ArgumentParser(
        description="Retrieve CISDM exports using a Chrome work profile."
    )
    parser.add_argument("report_id", nargs="?", default="student_metrics_summary")
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

    args = parser.parse_args()
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
        )

    print_summary(result)


if __name__ == "__main__":
    main()
