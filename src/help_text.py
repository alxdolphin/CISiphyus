"""Assemble full cisiphyus CLI help from existing argparse parsers."""

from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import example_cli

_SECTION_RULE = "=" * 72


def wants_full_help(argv: list[str]) -> bool:
    if len(argv) <= 1:
        return True
    return len(argv) == 2 and argv[1] in ("--help", "-h", "help")


def _section(title: str, body: str) -> str:
    return f"\n{_SECTION_RULE}\n{title}\n{_SECTION_RULE}\n\n{body.strip()}\n"


def _example_dir(name: str) -> Path:
    return example_cli.repo_root() / "examples" / name


def _import_example_builder(example_subdir: str, module_name: str, builder: str):
    example_dir = _example_dir(example_subdir)
    if not example_dir.is_dir():
        return None
    if str(example_dir) not in sys.path:
        sys.path.insert(0, str(example_dir))
    module = __import__(module_name)
    return getattr(module, builder)


def _parser_help(builder: Callable[..., Any], *, prog: str) -> str:
    parser = builder()
    parser.prog = prog
    return parser.format_help()


def format_full_help(*, build_retrieval_parser: Callable[..., Any]) -> str:
    parts = [
        "CISiphyus: retrieve CISDM exports and run downstream tools.",
        "Requires config/reports.yaml and local/export_urls.env (see README).",
    ]

    parts.append(
        _section(
            "DEFAULT RETRIEVAL",
            build_retrieval_parser(prog="cisiphyus").format_help(),
        )
    )
    parts.append(
        _section(
            "PULL",
            build_retrieval_parser(prog="cisiphyus pull").format_help(),
        )
    )
    parts.append(_section("AUDIT TARGETS", example_cli.AUDIT_USAGE.strip()))

    missing_examples: list[str] = []

    accred = _import_example_builder("accreditation", "accreditation", "build_parser")
    if accred is None:
        missing_examples.append("accreditation")
    else:
        parts.append(_section("AUDIT ACCREDITATION", _parser_help(accred, prog="cisiphyus audit accreditation")))

    audit = _import_example_builder("audit", "audit", "build_parser")
    if audit is None:
        missing_examples.append("audit")
    else:
        parts.append(_section("AUDIT METRICS", _parser_help(audit, prog="cisiphyus audit metrics")))

    qpr = _import_example_builder("qpr", "qpr", "build_parser")
    if qpr is None:
        missing_examples.append("qpr")
    else:
        parts.append(_section("QPR", _parser_help(qpr, prog="cisiphyus qpr")))

    trend = _import_example_builder("trends", "trends", "_build_trend_parser")
    cross_year = _import_example_builder("trends", "trends", "_build_cross_year_parser")
    if trend is None or cross_year is None:
        missing_examples.append("trend")
    else:
        parts.append(_section("TREND", trend(prog="cisiphyus trend").format_help()))
        parts.append(
            _section("TREND CROSS-YEAR", cross_year(prog="cisiphyus trend cross-year").format_help())
        )

    if missing_examples:
        parts.append(
            _section(
                "EXAMPLES UNAVAILABLE",
                "Example commands require a repository checkout (pip install -e .).\n"
                f"Missing: {', '.join(sorted(set(missing_examples)))}",
            )
        )

    return "\n".join(parts).strip() + "\n"
