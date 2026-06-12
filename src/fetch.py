from __future__ import annotations

import json
import re
import time
from pathlib import Path
from typing import Any, Callable, Literal

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

import auth_cookies
import browser
import config
import url_refresh
from bootstrap import capture_page_body_preview, classify_auth_probe

RetrievalStrategy = Literal["direct_url", "ui_export"]

# Silent SSO re-auth happens asynchronously a few seconds after page load.
AUTH_COOKIE_WAIT_SECONDS = 45
CONTEXT_COOKIE_WAIT_SECONDS = 30


class FetchFailedError(RuntimeError):
    def __init__(self, message: str, diag: dict[str, Any]) -> None:
        super().__init__(message)
        self.diag = diag


def classify_no_download_reason(body_preview: str) -> tuple[str, str | None]:
    lower = body_preview.lower()
    if "rsexecutionnotfound" in lower:
        return (
            "report_viewer_session_expired",
            "Report viewer execution or session in the URL expired. "
            "Refresh the ReportViewer entry URL in config/export_urls.env.",
        )
    if "report session" in lower and ("expired" in lower or "cannot be found" in lower):
        return (
            "report_viewer_session_expired",
            "Report viewer session in the URL expired. Refresh the entry URL in config/export_urls.env.",
        )
    if "execution" in lower and "expired" in lower and "report" in lower:
        return (
            "report_viewer_session_expired",
            "Report execution id in the URL expired. Refresh the entry URL in config/export_urls.env.",
        )
    return "no_download_after_export_navigation", None


def _session_diag() -> dict[str, Any]:
    return {
        "auth_source": "chrome_user_data_profile",
    }


def _inject_exported_cookies(context: Any, diag: dict[str, Any]) -> None:
    """Seed the session with cookies captured during bootstrap so the export
    starts fully authenticated, matching the proven cookie-overlay flow."""
    cookies_path = config.COOKIES_EXPORT_PATH
    if not cookies_path.is_file():
        return

    diag["cookies_file"] = str(cookies_path)
    try:
        raw = json.loads(cookies_path.read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            cookies = raw.get("cookies", [])
        else:
            cookies = raw
        if not isinstance(cookies, list):
            raise ValueError(
                f"expected a list of cookies, got {type(cookies).__name__}"
            )
        valid = [
            cookie
            for cookie in cookies
            if isinstance(cookie, dict)
            and cookie.get("name")
            and cookie.get("value") is not None
            and cookie.get("domain")
        ]
        if valid:
            context.add_cookies(valid)
    except (OSError, json.JSONDecodeError, PlaywrightError, TypeError, ValueError) as exc:
        diag["cookies_inject_error"] = f"{type(exc).__name__}: {exc}"
        raise FetchFailedError(
            f"cookie_injection_failed: exported cookies file is unreadable or "
            f"malformed ({type(exc).__name__}). Re-run bootstrap to regenerate "
            f"{cookies_path}. Command: {config.BOOTSTRAP_COMMAND}",
            diag,
        ) from exc
    diag["cookies_injected_count"] = len(valid)


def _warmup_page(page: Any) -> None:
    page.goto(config.BASE_URL, wait_until="domcontentloaded", timeout=60_000)


def _handle_no_download_timeout(
    page: Any,
    *,
    strategy: RetrievalStrategy,
    warmup_url: str,
) -> tuple[str, str | None]:
    body_preview, body_error = capture_page_body_preview(page)
    if body_error and not body_preview:
        body_preview = f"<body capture failed: {body_error}>"

    post_url = page.url
    post_title = page.title()
    _auth_ok, auth_reason = classify_auth_probe(post_url, body_preview, title=post_title)

    if auth_reason == "html_forbidden_response":
        if strategy == "direct_url" and "cis_prod.caseworthy" in warmup_url.lower():
            return (
                "export_forbidden_after_authenticated_warmup",
                "Warmup appears authenticated but export URL returned forbidden. "
                "Refresh the export URL in config/export_urls.env.",
            )
        return (
            "html_forbidden_response",
            f"Automation profile appears unauthorized. Run bootstrap first: {config.BOOTSTRAP_COMMAND}",
        )

    if auth_reason == "html_login_or_auth_response":
        return (
            "html_login_or_auth_response",
            f"Automation profile appears unauthenticated. Run bootstrap first: {config.BOOTSTRAP_COMMAND}",
        )

    reason, recommendation = classify_no_download_reason(body_preview)
    return reason, recommendation


def _ui_export_download(page: Any, output_path: Path) -> None:
    toolbar_frame = page
    for frame in page.frames:
        try:
            if (
                frame.locator("a.ExportLink").count() > 0
                or frame.locator("a:has(span.glyphui-save)").count() > 0
            ):
                toolbar_frame = frame
                break
        except PlaywrightError:
            continue

    total_pages = toolbar_frame.locator("[id$='_ctl09_ctl00_TotalPages']").first
    deadline = time.monotonic() + 180.0
    while time.monotonic() < deadline:
        try:
            raw = total_pages.inner_text(timeout=5_000).strip()
            match = re.search(r"\d+", raw)
            if match and int(match.group(0)) >= 1:
                break
        except PlaywrightError:
            pass
        time.sleep(0.5)

    time.sleep(1)

    save_export = toolbar_frame.locator("a.ExportLink").first
    if save_export.count() == 0:
        save_export = toolbar_frame.locator("a:has(span.glyphui-save)").first
    if save_export.count() == 0:
        save_export = toolbar_frame.get_by_role("button", name="Export drop down menu").first

    excel_link = toolbar_frame.locator("a.ActiveLink[title='Excel']").first
    if excel_link.count() == 0:
        excel_link = toolbar_frame.locator("a[title='Excel']").first

    save_export.hover(timeout=20_000)
    try:
        excel_link.wait_for(state="visible", timeout=8_000)
    except PlaywrightTimeoutError:
        save_export.click(timeout=20_000)
        excel_link.wait_for(state="visible", timeout=15_000)
    excel_link.click(timeout=20_000)


def _direct_url_download(page: Any, export_url: str) -> None:
    try:
        # Send the current page (module/warmup) as referer: a bare navigation
        # has none, and CISDM may reject exports without an in-app origin.
        page.goto(
            export_url,
            wait_until="domcontentloaded",
            timeout=max(config.NAVIGATION_TIMEOUT_MS, config.DOWNLOAD_TIMEOUT_MS),
            referer=page.url,
        )
    except PlaywrightError as exc:
        message = str(exc)
        if "Download is starting" not in message and "net::ERR_ABORTED" not in message:
            raise


def _resolve_direct_export_url(page: Any, target_url: str) -> tuple[str, dict[str, Any]]:
    try:
        candidates = url_refresh.discover_export_url_candidates(page)
        return url_refresh.choose_refreshed_export_url(target_url, candidates)
    except Exception as exc:
        return target_url, {
            "export_url_refreshed": False,
            "export_url_refresh_error": f"{type(exc).__name__}: {exc}",
        }


def _record_context_cookies(diag: dict[str, Any], context: Any) -> None:
    # Browser may already be unusable when a fetch fails; never let the
    # cookie snapshot mask the original error.
    try:
        diag["fetch_context_cookies"] = auth_cookies.context_cookie_inventory(
            context,
            cookie_urls=[config.BASE_URL],
        )
    except Exception as exc:
        diag["fetch_context_cookies_error"] = f"{type(exc).__name__}: {exc}"


def _establish_app_context(page: Any, context: Any, context_url: str) -> bool:
    """Visit the module/form page so CISDM issues Context_CaseWorthy.

    The export endpoints return 403 in a fresh session because the context
    session cookie only exists after entering an app module."""
    page.goto(
        context_url,
        wait_until="domcontentloaded",
        timeout=config.NAVIGATION_TIMEOUT_MS,
    )
    return auth_cookies.wait_for_cookie_in_context(
        context,
        "Context_CaseWorthy",
        timeout_seconds=CONTEXT_COOKIE_WAIT_SECONDS,
        cookie_urls=[config.BASE_URL, page.url],
    )


def fetch_excel_download(
    *,
    strategy: RetrievalStrategy,
    target_url: str,
    output_path: Path,
    context_url: str | None = None,
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
    headed: bool,
    playwright_factory: Callable[[], Any] | None = None,
) -> dict[str, Any]:
    browser.reject_system_chrome_dir(chrome_user_data_dir)
    chrome_user_data_dir.mkdir(parents=True, exist_ok=True)

    diag = _session_diag()
    pw_factory = playwright_factory or sync_playwright
    download_timeout = (
        config.REPORT_VIEWER_DOWNLOAD_TIMEOUT_MS
        if strategy == "ui_export"
        else config.DOWNLOAD_TIMEOUT_MS
    )

    with pw_factory() as p:
        context = browser.launch_persistent_context(
            p,
            chrome_user_data_dir=chrome_user_data_dir,
            chrome_profile_directory=chrome_profile_directory,
            headed=headed,
            accept_downloads=True,
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            _inject_exported_cookies(context, diag)
            _warmup_page(page)
            diag["warmup_url"] = page.url

            auth_ready = auth_cookies.wait_for_cookie_in_context(
                context,
                "Auth_CaseWorthy",
                timeout_seconds=AUTH_COOKIE_WAIT_SECONDS,
                cookie_urls=[config.BASE_URL],
            )
            diag["warmup_auth_cookie_present"] = auth_ready
            if not auth_ready:
                _record_context_cookies(diag, context)
                raise FetchFailedError(
                    "warmup_unauthenticated: Auth_CaseWorthy never appeared within "
                    f"{AUTH_COOKIE_WAIT_SECONDS}s of warmup. Run bootstrap first: "
                    f"{config.BOOTSTRAP_COMMAND}",
                    diag,
                )

            if context_url:
                context_ready = _establish_app_context(page, context, context_url)
                diag["context_url_visited"] = config.redact_url(context_url)
                diag["context_cookie_present"] = context_ready

            export_url_to_use = target_url
            if strategy == "direct_url":
                export_url_to_use, refresh_meta = _resolve_direct_export_url(page, target_url)
                diag.update(refresh_meta)
            elif strategy == "ui_export":
                page.goto(
                    target_url,
                    wait_until="domcontentloaded",
                    timeout=config.NAVIGATION_TIMEOUT_MS,
                )

            try:
                with page.expect_download(timeout=download_timeout) as download_info:
                    if strategy == "ui_export":
                        _ui_export_download(page, output_path)
                    else:
                        _direct_url_download(page, export_url_to_use)

                download = download_info.value
                output_path.parent.mkdir(parents=True, exist_ok=True)
                download.save_as(output_path)
                diag["download_suggested_filename"] = download.suggested_filename
                return diag
            except PlaywrightTimeoutError:
                diag["post_navigation_url"] = page.url
                _record_context_cookies(diag, context)
                reason, hint = _handle_no_download_timeout(
                    page, strategy=strategy, warmup_url=diag["warmup_url"]
                )
                message = f"Download failed: {reason}"
                if hint:
                    message = f"{message}. {hint}"
                raise FetchFailedError(message, diag) from None
            except PlaywrightError as exc:
                diag["post_navigation_url"] = page.url
                _record_context_cookies(diag, context)
                if strategy == "ui_export":
                    raise FetchFailedError(f"report_viewer_excel_export_failed: {exc}", diag) from exc
                raise FetchFailedError(str(exc), diag) from exc
        finally:
            context.close()
