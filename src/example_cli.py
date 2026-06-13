"""Dispatch installed cisiphyus CLI invocations to examples/ entrypoints."""

from __future__ import annotations

import sys
from pathlib import Path

ACCREDITATION_MONITOR_FLAGS = frozenset(
    {
        "--workbook",
        "--force-fetch",
        "--fetch-destination",
        "--include-inactive-sites",
        "--output-dir",
        "--reporting-rule",
        "--reporting-threshold",
        "--include-ok-sites",
        "--json",
    }
)


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _run_example_module(*, example_dir: Path, module_name: str, argv: list[str]) -> int:
    if str(example_dir) not in sys.path:
        sys.path.insert(0, str(example_dir))
    module = __import__(module_name)
    return int(module.main(argv))


def run_accreditation_monitor(argv: list[str]) -> int:
    return _run_example_module(
        example_dir=repo_root() / "examples" / "accreditation",
        module_name="accreditation_monitor",
        argv=argv,
    )


def run_qpr_provision(argv: list[str]) -> int:
    return _run_example_module(
        example_dir=repo_root() / "examples" / "qpr",
        module_name="qpr_provision",
        argv=argv,
    )


def maybe_run_example_app(argv: list[str]) -> int | None:
    if len(argv) < 2:
        return None
    command = argv[1]
    rest = argv[2:]

    if command == "qpr":
        return run_qpr_provision(rest)

    if command == "accreditation" and _has_accreditation_monitor_flag(rest):
        return run_accreditation_monitor(rest)

    return None


def _has_accreditation_monitor_flag(argv: list[str]) -> bool:
    for token in argv:
        flag = token.split("=", 1)[0]
        if flag in ("--help", "-h"):
            return True
        if flag in ACCREDITATION_MONITOR_FLAGS:
            return True
    return False
