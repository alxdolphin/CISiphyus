from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import artifacts
import config
import fetch
import validate


def _resolve_report_urls(
    report_id: str,
    profile: dict,
) -> tuple[str, str, str]:
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
        return retrieval_strategy, entry_url, entry_key

    url_env = profile.get("url_env")
    if not url_env or not isinstance(url_env, str) or not str(url_env).strip():
        raise RuntimeError(f"Report {report_id} missing url_env")
    url_env = str(url_env).strip()
    export_url = os.environ.get(url_env, "")
    if not export_url:
        raise RuntimeError(f"Missing export URL env var: {url_env}")
    return retrieval_strategy, export_url, url_env


def _failure_diag(
    *,
    error: str,
    cookies_file: Path | None,
    storage_state_file: Path | None,
    fetch_diag: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "error": error,
        "cookies_file": str(cookies_file) if cookies_file else None,
        "storage_state_file": str(storage_state_file) if storage_state_file else None,
    }
    if fetch_diag:
        payload.update(fetch_diag)
    return payload


def run_report(
    report_id: str,
    headed: bool,
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
    cookies_file: Path | None,
    storage_state_file: Path | None = None,
) -> artifacts.RunResult:
    started_at = artifacts.utc_now_iso()
    latest_dir = config.latest_report_dir(report_id)
    raw_path = config.latest_raw_path(report_id)

    try:
        config.load_env_file(config.ENV_PATH)
        reports = config.load_reports(config.REPORTS_PATH)

        if report_id not in reports:
            available = sorted(reports.keys())
            raise RuntimeError(f"Unknown report_id: {report_id}. Available: {available}")

        profile = reports[report_id]
        retrieval_strategy, target_url, _env_key = _resolve_report_urls(report_id, profile)
        min_size_bytes = int(profile.get("min_size_bytes", 1))
        required_columns = profile.get("required_columns", [])

        if not chrome_user_data_dir.exists():
            chrome_user_data_dir.mkdir(parents=True, exist_ok=True)

        fetch.fetch_excel_download(
            strategy=retrieval_strategy,
            target_url=target_url,
            output_path=raw_path,
            chrome_user_data_dir=chrome_user_data_dir,
            chrome_profile_directory=chrome_profile_directory,
            cookies_file=cookies_file,
            storage_state_file=storage_state_file,
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
            latest_dir=latest_dir,
        )
    except fetch.FetchFailedError as exc:
        error = f"{type(exc).__name__}: {exc}"
        artifacts.write_failure_diag(
            latest_dir,
            _failure_diag(
                error=error,
                cookies_file=cookies_file,
                storage_state_file=storage_state_file,
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
            latest_dir=latest_dir,
        )
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        artifacts.write_failure_diag(
            latest_dir,
            _failure_diag(
                error=error,
                cookies_file=cookies_file,
                storage_state_file=storage_state_file,
            ),
        )
        return artifacts.write_latest_result(
            report_id=report_id,
            status="failed",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=raw_path if raw_path.exists() else None,
            error=error,
            latest_dir=latest_dir,
        )
