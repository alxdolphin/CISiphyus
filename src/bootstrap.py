from __future__ import annotations

import json
import shutil
import time
from pathlib import Path
from typing import Any, Callable

from playwright.sync_api import sync_playwright

import artifacts
import browser
import config

BOOTSTRAP_SETTLE_SECONDS = 2


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


def bootstrap_profile(
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
    *,
    force: bool = False,
    playwright_factory: Callable[[], Any] | None = None,
) -> artifacts.RunResult:
    started_at = artifacts.utc_now_iso()
    marker_path = config.auth_bootstrap_marker_path(chrome_user_data_dir)

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
                    "Bootstrap: log in until CISDM loads (not the login screen).\n"
                    "DOnce logged in, press Enter here."
                )
                input()
                auth_verified, verify_reason, verify_metadata = verify_cisdm_auth(
                    page,
                    base_url=config.BASE_URL,
                    reload=False,
                )
                if not auth_verified:
                    raise RuntimeError(
                        f"CISDM auth verification failed ({verify_reason}). "
                        "Complete SSO/MFA in the browser window and try again with --force."
                    )

                time.sleep(BOOTSTRAP_SETTLE_SECONDS)
                marker_payload = {
                    "bootstrapped_at": artifacts.utc_now_iso(),
                    "base_url": config.BASE_URL,
                    "final_url": page.url,
                    "final_title": page.title(),
                    "chrome_user_data_dir": str(chrome_user_data_dir),
                    "chrome_profile_directory": chrome_profile_directory,
                    "auth_verify_reason": verify_reason,
                }
                marker_payload.update(verify_metadata)
                marker_path.parent.mkdir(parents=True, exist_ok=True)
                marker_path.write_text(
                    json.dumps(marker_payload, indent=2, sort_keys=True),
                    encoding="utf-8",
                )
            finally:
                context.close()

        return artifacts.RunResult(
            report_id="bootstrap",
            status="success",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=None,
            error=None,
        )
    except Exception as exc:
        return artifacts.RunResult(
            report_id="bootstrap",
            status="failed",
            started_at=started_at,
            finished_at=artifacts.utc_now_iso(),
            raw_path=None,
            error=f"{type(exc).__name__}: {exc}",
        )
