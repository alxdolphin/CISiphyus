"""Dispatch installed cisiphyus CLI invocations to examples/ entrypoints."""

from __future__ import annotations

import sys
from pathlib import Path

AUDIT_USAGE = """usage: cisiphyus audit <target> [options]

targets:
  accreditation   Site-level accreditation compliance monitoring
  metrics         Student metrics summary row-level audit
"""

TREND_USAGE = """usage: cisiphyus trend <school-year> [school-year ...] [options]

Discover available reporting periods, capture snapshots, and compare consecutive quarters.

examples:
  cisiphyus trend SY25-26
  cisiphyus trend SY25-26 SY24-25
  cisiphyus trend --all
"""


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _run_example_module(*, example_dir: Path, module_name: str, argv: list[str]) -> int:
    if not example_dir.is_dir():
        print(
            f"Error: Example directory not found at {example_dir}\n"
            "This command is only available when running from a repository checkout "
            "(e.g., pip install -e .).",
            file=sys.stderr,
        )
        return 1
    if str(example_dir) not in sys.path:
        sys.path.insert(0, str(example_dir))
    module = __import__(module_name)
    return int(module.main(argv))


def run_accreditation_monitor(argv: list[str]) -> int:
    return _run_example_module(
        example_dir=repo_root() / "examples" / "accreditation",
        module_name="accreditation",
        argv=argv,
    )


def run_metrics_audit(argv: list[str]) -> int:
    return _run_example_module(
        example_dir=repo_root() / "examples" / "audit",
        module_name="audit",
        argv=argv,
    )


def run_audit_app(argv: list[str]) -> int:
    if not argv or argv[0] in ("--help", "-h"):
        print(AUDIT_USAGE, file=sys.stderr)
        return 0 if argv and argv[0] in ("--help", "-h") else 1

    target = argv[0]
    rest = argv[1:]
    if target == "accreditation":
        return run_accreditation_monitor(rest)
    if target == "metrics":
        return run_metrics_audit(rest)

    print(f"Error: unknown audit target {target!r}. Use 'accreditation' or 'metrics'.", file=sys.stderr)
    print(AUDIT_USAGE, file=sys.stderr)
    return 1


def run_qpr_provision(argv: list[str]) -> int:
    return _run_example_module(
        example_dir=repo_root() / "examples" / "qpr",
        module_name="qpr",
        argv=argv,
    )


def run_trend(argv: list[str]) -> int:
    if not argv or argv[0] in ("--help", "-h"):
        print(TREND_USAGE, file=sys.stderr)
        return 0 if argv and argv[0] in ("--help", "-h") else 1
    example_dir = repo_root() / "examples" / "trends"
    if not example_dir.is_dir():
        print(
            f"Error: Example directory not found at {example_dir}\n"
            "This command is only available when running from a repository checkout "
            "(e.g., pip install -e .).",
            file=sys.stderr,
        )
        return 1
    if str(example_dir) not in sys.path:
        sys.path.insert(0, str(example_dir))
    trends = __import__("trends")
    return int(trends.main_trend(argv))


def maybe_run_example_app(argv: list[str]) -> int | None:
    if len(argv) < 2:
        return None
    command = argv[1]
    rest = argv[2:]

    if command == "qpr":
        return run_qpr_provision(rest)

    if command == "audit":
        return run_audit_app(rest)

    if command == "trend":
        return run_trend(rest)

    return None
