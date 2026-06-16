from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import artifacts
import config
import url_refresh
import validate


def _resolve_report_urls(
    report_id: str,
    profile: dict,
    *,
    school_year: str | None = None,
    school_year_programs: dict[str, int] | None = None,
) -> tuple[str, str, str, str | None, int | None]:
    retrieval_strategy = str(profile.get("retrieval_strategy", "direct_url")).strip().lower()
    if retrieval_strategy not in ("direct_url", "ui_export"):
        raise RuntimeError(
            f"Report {report_id} retrieval_strategy must be direct_url or ui_export, "
            f"got {retrieval_strategy!r}"
        )

    if retrieval_strategy == "ui_export":
        ui_entry_url_env = profile.get("ui_entry_url_env")
        if not isinstance(ui_entry_url_env, str) or not ui_entry_url_env.strip():
            raise RuntimeError(
                f"Report {report_id} retrieval_strategy ui_export requires ui_entry_url_env in reports.yaml"
            )
        entry_key = ui_entry_url_env.strip()
        entry_url = os.environ.get(entry_key, "")
        if not entry_url:
            raise RuntimeError(f"Missing ReportViewer entry URL env var: {entry_key}")
        return retrieval_strategy, entry_url, entry_key, None, None

    url_env = profile.get("url_env")
    if not url_env or not isinstance(url_env, str) or not str(url_env).strip():
        raise RuntimeError(f"Report {report_id} missing url_env")
    url_env = str(url_env).strip()
    export_url_value = os.environ.get(url_env, "")
    if not export_url_value:
        raise RuntimeError(f"Missing export URL env var: {url_env}")

    resolved_school_year: str | None = None
    enrollment_program_id: int | None = None
    year_scope = profile.get("year_scope")
    if year_scope is not None:
        if not isinstance(year_scope, dict):
            raise RuntimeError(f"Report {report_id} year_scope must be a mapping")
        param = year_scope.get("param")
        if not isinstance(param, str) or not param.strip():
            raise RuntimeError(f"Report {report_id} year_scope.param must be a non-empty string")
        programs = school_year_programs or {}
        if not programs:
            raise RuntimeError(
                f"Report {report_id} has year_scope but school_year_programs is empty in reports.yaml"
            )
        resolved_school_year = school_year or config.default_school_year(programs)
        if not resolved_school_year:
            raise RuntimeError(
                f"Report {report_id} requires a school year; pass --school-year or set "
                "CISDM_DEFAULT_SCHOOL_YEAR"
            )
        enrollment_program_id = config.resolve_program_id(programs, resolved_school_year)
        export_url_value = url_refresh.resolve_year_scoped_url(
            export_url_value,
            param=param.strip(),
            program_id=enrollment_program_id,
        )

    return (
        retrieval_strategy,
        export_url_value,
        url_env,
        resolved_school_year,
        enrollment_program_id,
    )


def _resolve_context_url(report_id: str, profile: dict) -> str | None:
    """Optional module/form page visited before export to establish the CISDM
    app context (issues the Context_CaseWorthy session cookie)."""
    context_url_env = profile.get("context_url_env")
    if context_url_env is None:
        return None
    if not isinstance(context_url_env, str) or not context_url_env.strip():
        raise RuntimeError(f"Report {report_id} context_url_env must be a non-empty string")
    entry_key = context_url_env.strip()
    context_url = os.environ.get(entry_key, "")
    if not context_url:
        raise RuntimeError(f"Missing context URL env var: {entry_key}")
    return context_url


def _resolve_output_paths(
    report_id: str,
    *,
    school_year: str | None,
    default_school_year: str | None,
) -> tuple[Path, Path]:
    if school_year and default_school_year and school_year != default_school_year:
        output_dir = config.archives_pull_dir(school_year, report_id)
        raw_path = config.archives_raw_path(school_year, report_id)
        return output_dir, raw_path
    return config.latest_report_dir(report_id), config.latest_raw_path(report_id)


def _failure_diag(
    *,
    error: str,
    fetch_diag: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "error": error,
        "auth_source": "chrome_user_data_profile",
    }
    if fetch_diag:
        payload.update(fetch_diag)
    return payload


def run_report(
    report_id: str,
    headed: bool,
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
    school_year: str | None = None,
) -> artifacts.RunResult:
    import fetch

    started_at = artifacts.utc_now_iso()
    output_dir = config.latest_report_dir(report_id)
    raw_path = config.latest_raw_path(report_id)
    resolved_school_year: str | None = None
    enrollment_program_id: int | None = None

    try:
        config.load_env_file(config.ENV_PATH)
        reports = config.load_reports(config.REPORTS_PATH)
        school_year_programs = config.load_school_year_programs(config.REPORTS_PATH)
        default_school_year = config.default_school_year(school_year_programs)

        if report_id not in reports:
            available = sorted(reports.keys())
            raise RuntimeError(f"Unknown report_id: {report_id}. Available: {available}")

        profile = reports[report_id]
        (
            retrieval_strategy,
            target_url,
            _env_key,
            resolved_school_year,
            enrollment_program_id,
        ) = _resolve_report_urls(
            report_id,
            profile,
            school_year=school_year,
            school_year_programs=school_year_programs,
        )
        output_dir, raw_path = _resolve_output_paths(
            report_id,
            school_year=resolved_school_year,
            default_school_year=default_school_year,
        )
        context_url = _resolve_context_url(report_id, profile)
        min_size_bytes = int(profile.get("min_size_bytes", 1))
        required_columns = profile.get("required_columns", [])

        if not chrome_user_data_dir.exists():
            chrome_user_data_dir.mkdir(parents=True, exist_ok=True)

        fetch.fetch_excel_download(
            strategy=retrieval_strategy,
            target_url=target_url,
            context_url=context_url,
            output_path=raw_path,
            chrome_user_data_dir=chrome_user_data_dir,
            chrome_profile_directory=chrome_profile_directory,
            headed=headed,
        )

        if not raw_path.exists() or raw_path.stat().st_size == 0:
            raise RuntimeError("Download completed but raw.xlsx is missing or empty")

        validation_ok, validation_reason, _metadata = validate.validate_workbook(
            raw_path,
            required_columns=required_columns,
            min_size_bytes=min_size_bytes,
        )
        if not validation_ok:
            raise RuntimeError(f"Validation failed: {validation_reason}")

        return artifacts.write_latest_result(
            report_id=report_id,
            status="success",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=raw_path,
            error=None,
            latest_dir=output_dir,
            school_year=resolved_school_year,
            enrollment_program_id=enrollment_program_id,
        )
    except fetch.FetchFailedError as exc:
        error = f"{type(exc).__name__}: {exc}"
        artifacts.write_failure_diag(
            output_dir,
            _failure_diag(
                error=error,
                fetch_diag=exc.diag,
            ),
        )
        return artifacts.write_latest_result(
            report_id=report_id,
            status="failed",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=raw_path if raw_path.exists() else None,
            error=error,
            latest_dir=output_dir,
            school_year=resolved_school_year,
            enrollment_program_id=enrollment_program_id,
        )
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        artifacts.write_failure_diag(
            output_dir,
            _failure_diag(error=error),
        )
        return artifacts.write_latest_result(
            report_id=report_id,
            status="failed",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=raw_path if raw_path.exists() else None,
            error=error,
            latest_dir=output_dir,
            school_year=resolved_school_year,
            enrollment_program_id=enrollment_program_id,
        )
