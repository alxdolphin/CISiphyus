from __future__ import annotations

import json
import os
import shutil
import time
from pathlib import Path
from typing import Any, Callable

from playwright.sync_api import sync_playwright

import artifacts
import auth_cookies
import browser
import config

def _parse_settle_seconds(raw: str | None, default: int = 2) -> int:
    """Parse the settle-seconds override defensively: a malformed env var must
    never prevent the CLI from importing this module."""
    if raw is None:
        return default
    try:
        value = int(raw.strip())
    except (TypeError, ValueError):
        return default
    return value if value >= 0 else default


BOOTSTRAP_SETTLE_SECONDS = _parse_settle_seconds(
    os.environ.get("CISIPHYUS_BOOTSTRAP_SETTLE_SECONDS")
)


def classify_auth_probe(url: str, body_preview: str, *, title: str = "") -> tuple[bool, str]:
    lower = body_preview.lower()
    url_lower = url.lower()
    title_lower = title.lower()

    if "403 forbidden" in lower or "you do not have permission" in lower:
        return False, "html_forbidden_response"

    login_signals = (
        "sign in",
        "log in",
        "username",
        "forgot password",
        "forgot your password",
        "enter your password",
        "enter your login",
    )
    login_page = any(signal in lower for signal in login_signals)

    if "cis_prod.caseworthy" in url_lower and not login_page:
        return True, "auth_verified"

    if "caseworthy" in url_lower and "login" not in url_lower and not login_page:
        return True, "auth_verified"

    if login_page:
        return False, "html_login_or_auth_response"

    if "caseworthy" in lower:
        return True, "auth_verified"

    if not body_preview.strip() and title_lower.startswith("cis_"):
        if "cis_prod.caseworthy" in url_lower or (
            "caseworthy" in url_lower and "login" not in url_lower
        ):
            return True, "auth_verified_title_url"

    return False, "auth_state_unclear"


def dismiss_blocking_dialogs(page: Any) -> None:
    for _ in range(2):
        try:
            page.keyboard.press("Escape")
            time.sleep(0.25)
        except Exception:
            break

    for label in ("Never", "Not now", "No thanks", "Cancel"):
        try:
            button = page.get_by_role("button", name=label)
            if button.count() > 0:
                button.first.click(timeout=1500)
                time.sleep(0.25)
        except Exception:
            continue


def capture_page_body_preview(page: Any) -> tuple[str, str | None]:
    try:
        return page.locator("body").inner_text(timeout=10_000), None
    except Exception as exc:
        body_error = f"{type(exc).__name__}: {exc}"

    try:
        body_preview = page.evaluate(
            "() => (document.body && document.body.innerText) ? document.body.innerText : ''"
        )
        if isinstance(body_preview, str) and body_preview.strip():
            return body_preview, body_error
    except Exception as eval_exc:
        body_error = f"{body_error}; evaluate failed: {type(eval_exc).__name__}: {eval_exc}"

    return "", body_error


def probe_page_auth(page: Any) -> tuple[bool, str, dict[str, Any]]:
    metadata: dict[str, Any] = {}
    try:
        metadata["verify_url"] = page.url
        metadata["verify_title"] = page.title()
    except Exception as exc:
        metadata["verify_error"] = f"{type(exc).__name__}: {exc}"
        return False, "auth_verify_probe_failed", metadata

    body_preview, body_error = capture_page_body_preview(page)
    if body_preview:
        metadata["body_preview"] = config.redact_sensitive_text(body_preview)
    elif body_error:
        metadata["body_capture_error"] = body_error

    ok, reason = classify_auth_probe(
        metadata["verify_url"],
        body_preview,
        title=metadata.get("verify_title", ""),
    )
    return ok, reason, metadata


def verify_cisdm_auth(
    page: Any,
    *,
    base_url: str,
    reload: bool = False,
) -> tuple[bool, str, dict[str, Any]]:
    dismiss_blocking_dialogs(page)
    ok, reason, metadata = probe_page_auth(page)
    if ok or reason in {"html_forbidden_response", "html_login_or_auth_response"}:
        return ok, reason, metadata
    if not reload:
        return ok, reason, metadata

    try:
        page.goto(base_url, wait_until="domcontentloaded", timeout=config.NAVIGATION_TIMEOUT_MS)
        time.sleep(BOOTSTRAP_SETTLE_SECONDS)
        dismiss_blocking_dialogs(page)
        return probe_page_auth(page)
    except Exception as exc:
        return False, "auth_verify_navigation_failed", {"verify_error": f"{type(exc).__name__}: {exc}"}


def flush_auth_state(page: Any, *, base_url: str) -> None:
    # networkidle never fires on CISDM's long-polling SPA; domcontentloaded
    # plus the settle sleep is enough since this only flushes auth state.
    page.goto(base_url, wait_until="domcontentloaded", timeout=config.NAVIGATION_TIMEOUT_MS)
    time.sleep(BOOTSTRAP_SETTLE_SECONDS)
    dismiss_blocking_dialogs(page)


def _resolve_verify_report_url(report_id: str) -> str:
    config.load_env_file(config.ENV_PATH)
    reports = config.load_reports(config.REPORTS_PATH)
    if report_id not in reports:
        available = sorted(reports.keys())
        raise RuntimeError(
            f"Unknown verify-report id: {report_id}. Available: {available}"
        )

    profile = reports[report_id]
    retrieval_strategy = str(profile.get("retrieval_strategy", "direct_url")).strip().lower()
    if retrieval_strategy == "ui_export":
        entry_key = profile.get("ui_entry_url_env")
        if not isinstance(entry_key, str) or not entry_key.strip():
            raise RuntimeError(
                f"Report {report_id} retrieval_strategy ui_export requires ui_entry_url_env"
            )
        entry_url = os.environ.get(entry_key.strip(), "")
        if not entry_url:
            raise RuntimeError(f"Missing ReportViewer entry URL env var: {entry_key.strip()}")
        return entry_url

    url_env = profile.get("url_env")
    if not url_env or not isinstance(url_env, str) or not str(url_env).strip():
        raise RuntimeError(f"Report {report_id} missing url_env")
    export_url = os.environ.get(str(url_env).strip(), "")
    if not export_url:
        raise RuntimeError(f"Missing export URL env var: {url_env}")
    return export_url


def verify_report_probe(page: Any, report_id: str) -> tuple[bool, str, dict[str, Any]]:
    target_url = _resolve_verify_report_url(report_id)
    page.goto(target_url, wait_until="domcontentloaded", timeout=config.NAVIGATION_TIMEOUT_MS)
    time.sleep(BOOTSTRAP_SETTLE_SECONDS)
    dismiss_blocking_dialogs(page)
    ok, reason, metadata = probe_page_auth(page)
    metadata["verify_report_id"] = report_id
    metadata["verify_report_url"] = config.redact_url(target_url)
    return ok, reason, metadata


def headless_smoke_check(
    p: Any,
    *,
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
) -> dict[str, Any]:
    context = browser.launch_persistent_context(
        p,
        chrome_user_data_dir=chrome_user_data_dir,
        chrome_profile_directory=chrome_profile_directory,
        headed=False,
        accept_downloads=False,
    )
    try:
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(config.BASE_URL, wait_until="domcontentloaded", timeout=config.NAVIGATION_TIMEOUT_MS)
        time.sleep(BOOTSTRAP_SETTLE_SECONDS)
        dismiss_blocking_dialogs(page)

        auth_ok, auth_reason, probe_metadata = probe_page_auth(page)
        if not auth_ok:
            raise RuntimeError(
                f"Headless smoke check failed ({auth_reason}). "
                "Persisted profile did not survive a cold headless relaunch."
            )

        # Context_CaseWorthy is a session cookie issued only when entering an
        # app module; it never exists in a cold relaunch at the base URL.
        # The auth probe above is the real pass/fail signal, so only require
        # the auth cookie here and record the full inventory for the audit trail.
        auth_cookies.assert_required_cookies_in_context(
            context,
            cookie_urls=[config.BASE_URL, page.url],
            required_names=("Auth_CaseWorthy",),
        )
        return {
            "headless_smoke_check": "passed",
            "headless_auth_verify_reason": auth_reason,
            "headless_cookies": auth_cookies.context_cookie_inventory(
                context,
                cookie_urls=[config.BASE_URL, page.url],
            ),
            **probe_metadata,
        }
    finally:
        context.close()


def bootstrap_profile(
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
    *,
    force: bool = False,
    verify_report: str | None = None,
    playwright_factory: Callable[[], Any] | None = None,
) -> artifacts.RunResult:
    started_at = artifacts.utc_now_iso()
    marker_path = config.auth_bootstrap_marker_path(chrome_user_data_dir)
    profile_dir = chrome_user_data_dir / chrome_profile_directory
    extra_diag: dict[str, Any] = {}

    try:
        browser.reject_system_chrome_dir(chrome_user_data_dir)

        if config.bootstrap_profile_exists(chrome_user_data_dir):
            if not force:
                raise RuntimeError(
                    f"Bootstrap profile already exists at {chrome_user_data_dir}. "
                    "Delete it manually or re-run with --force."
                )
            shutil.rmtree(chrome_user_data_dir)

        chrome_user_data_dir.mkdir(parents=True, exist_ok=True)

        pw_factory = playwright_factory or sync_playwright
        with pw_factory() as p:
            browser.warn_headed_linux_gui_environment("bootstrap")
            context = browser.launch_persistent_context(
                p,
                chrome_user_data_dir=chrome_user_data_dir,
                chrome_profile_directory=chrome_profile_directory,
                headed=True,
                accept_downloads=False,
            )
            try:
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(config.BASE_URL, wait_until="domcontentloaded", timeout=60_000)
                print(
                    "Bootstrap: log in and wait until the CISDM dashboard fully renders\n"
                    "(not just the post-SSO redirect or the login screen).\n"
                    "Once the dashboard is visible, press Enter here."
                )
                input()
                auth_verified, verify_reason, verify_metadata = verify_cisdm_auth(
                    page,
                    base_url=config.BASE_URL,
                    reload=True,
                )
                if not auth_verified:
                    raise RuntimeError(
                        f"CISDM auth verification failed ({verify_reason}). "
                        "Complete SSO/MFA in the browser window and try again with --force."
                    )

                # Capture cookies at their richest state: CISDM drops
                # Context_CaseWorthy when navigating back to the base URL,
                # so the check must run before any further navigation.
                context_cookie_metadata = auth_cookies.assert_required_cookies_in_context(
                    context,
                    cookie_urls=[config.BASE_URL, page.url],
                )

                exported_cookies = auth_cookies.export_caseworthy_cookies(
                    context,
                    cookie_urls=[config.BASE_URL, page.url],
                )
                config.COOKIES_EXPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
                config.COOKIES_EXPORT_PATH.write_text(
                    json.dumps({"cookies": exported_cookies}, indent=2),
                    encoding="utf-8",
                )
                config.COOKIES_EXPORT_PATH.chmod(0o600)
                extra_diag["cookies_export_path"] = str(config.COOKIES_EXPORT_PATH)
                extra_diag["cookies_export_count"] = len(exported_cookies)

                report_probe_metadata: dict[str, Any] = {}
                if verify_report:
                    report_ok, report_reason, report_probe_metadata = verify_report_probe(
                        page, verify_report
                    )
                    if not report_ok:
                        raise RuntimeError(
                            f"ReportViewer probe failed for {verify_report!r} ({report_reason}). "
                            "Confirm export URLs in config/export_urls.env."
                        )

                flush_auth_state(page, base_url=config.BASE_URL)
                extra_diag["post_flush_cookies"] = auth_cookies.context_cookie_inventory(
                    context,
                    cookie_urls=[config.BASE_URL, page.url],
                )
            finally:
                context.close()

            disk_cookie_metadata = auth_cookies.assert_required_cookies_on_disk(profile_dir)
            smoke_metadata = headless_smoke_check(
                p,
                chrome_user_data_dir=chrome_user_data_dir,
                chrome_profile_directory=chrome_profile_directory,
            )

            marker_payload = {
                "bootstrapped_at": artifacts.utc_now_iso(),
                "base_url": config.BASE_URL,
                "chrome_user_data_dir": str(chrome_user_data_dir),
                "chrome_profile_directory": chrome_profile_directory,
                "auth_verify_reason": verify_reason,
                "disk_verified_at": artifacts.utc_now_iso(),
            }
            marker_payload.update(verify_metadata)
            marker_payload.update(report_probe_metadata)
            marker_payload.update(context_cookie_metadata)
            marker_payload.update(extra_diag)
            marker_payload.update(disk_cookie_metadata)
            marker_payload.update(smoke_metadata)
            marker_path.parent.mkdir(parents=True, exist_ok=True)
            marker_path.write_text(
                json.dumps(marker_payload, indent=2, sort_keys=True),
                encoding="utf-8",
            )

        return artifacts.write_latest_result(
            report_id="bootstrap",
            status="success",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=None,
            error=None,
            latest_dir=config.latest_report_dir("bootstrap"),
        )
    except Exception as exc:
        latest_dir = config.latest_report_dir("bootstrap")
        diag_payload: dict[str, Any] = {
            "error": f"{type(exc).__name__}: {exc}",
            "auth_source": "chrome_user_data_profile",
            **extra_diag,
        }
        if isinstance(exc, auth_cookies.RequiredCookiesError) and exc.diagnostics:
            diag_payload.update(exc.diagnostics)
        artifacts.write_failure_diag(latest_dir, diag_payload)
        return artifacts.write_latest_result(
            report_id="bootstrap",
            status="failed",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=None,
            error=f"{type(exc).__name__}: {exc}",
            latest_dir=latest_dir,
        )
