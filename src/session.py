from __future__ import annotations

import getpass
import json
import os
import platform
from pathlib import Path
from typing import Any

import config

DEFAULT_STORAGE_STATE_PATH = config.ROOT / "session" / "storage_state.json"


def wsl_default_cookies_candidates() -> list[Path]:
    if platform.system() != "Linux":
        return []

    users_root = Path("/mnt/c/Users")
    if not users_root.exists():
        return []

    usernames: list[str] = []
    for env_name in ("WINDOWS_USERNAME", "WIN_USERNAME", "USER"):
        value = os.environ.get(env_name, "").strip()
        if value:
            usernames.append(value)
    try:
        usernames.append(getpass.getuser())
    except Exception:
        pass

    unique_usernames = list(dict.fromkeys(usernames))
    return list(
        dict.fromkeys(
            users_root / username / "Projects" / "CISDM_cookies.json"
            for username in unique_usernames
        )
    )


def default_cookies_candidates() -> list[Path]:
    root = config.ROOT
    candidates: list[Path] = [root.parent / "CISDM_cookies.json"]
    if len(root.parents) >= 2:
        candidates.append(root.parents[1] / "CISDM_cookies.json")
    if len(root.parents) >= 3:
        candidates.append(root.parents[2] / "CISDM_cookies.json")
    candidates.extend(wsl_default_cookies_candidates())
    return list(dict.fromkeys(candidates))


def default_cookies_file() -> Path | None:
    for candidate in default_cookies_candidates():
        if candidate.exists():
            return candidate
    return None


def default_storage_state_file() -> Path | None:
    if DEFAULT_STORAGE_STATE_PATH.exists():
        return DEFAULT_STORAGE_STATE_PATH
    return None


def normalize_same_site(raw_value: str | None) -> str | None:
    if not raw_value:
        return None
    normalized = raw_value.strip().lower()
    mapping = {
        "no_restriction": "None",
        "none": "None",
        "lax": "Lax",
        "strict": "Strict",
    }
    return mapping.get(normalized)


def load_playwright_cookies(cookies_file: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    raw = json.loads(cookies_file.read_text(encoding="utf-8"))
    source_cookies = raw if isinstance(raw, list) else raw.get("cookies", [])

    if not isinstance(source_cookies, list):
        raise ValueError("Cookies file must contain a list or a {cookies: [...]} object.")

    converted: list[dict[str, Any]] = []
    skipped = 0
    for cookie in source_cookies:
        if not isinstance(cookie, dict):
            skipped += 1
            continue

        name = cookie.get("name")
        value = cookie.get("value")
        domain = cookie.get("domain")
        path = cookie.get("path") or "/"
        if not name or value is None or not domain:
            skipped += 1
            continue

        converted_cookie: dict[str, Any] = {
            "name": str(name),
            "value": str(value),
            "domain": str(domain),
            "path": str(path),
            "secure": bool(cookie.get("secure", False)),
            "httpOnly": bool(cookie.get("httpOnly", False)),
        }

        same_site = normalize_same_site(cookie.get("sameSite"))
        if same_site:
            converted_cookie["sameSite"] = same_site

        expiration_date = cookie.get("expirationDate")
        if expiration_date is not None:
            try:
                converted_cookie["expires"] = float(expiration_date)
            except (TypeError, ValueError):
                pass

        converted.append(converted_cookie)

    metadata = {
        "cookies_file": str(cookies_file),
        "cookies_loaded_count": len(converted),
        "cookies_skipped_count": skipped,
    }
    return converted, metadata


def load_storage_state_payload(
    storage_state_file: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    raw = json.loads(storage_state_file.read_text(encoding="utf-8"))
    cookies = raw.get("cookies", []) if isinstance(raw, dict) else []
    if not isinstance(cookies, list):
        raise ValueError("Storage state file must contain {cookies: [...]} structure.")

    valid: list[dict[str, Any]] = []
    skipped = 0
    for cookie in cookies:
        if not isinstance(cookie, dict):
            skipped += 1
            continue
        if not cookie.get("name") or cookie.get("value") is None or not cookie.get("domain"):
            skipped += 1
            continue
        valid.append(cookie)

    origins = raw.get("origins", []) if isinstance(raw, dict) else []
    if not isinstance(origins, list):
        raise ValueError("Storage state file must contain {origins: [...]} structure.")

    local_storage_origins: list[dict[str, Any]] = []
    origins_skipped = 0
    local_storage_items_loaded = 0
    local_storage_items_skipped = 0

    for origin_entry in origins:
        if not isinstance(origin_entry, dict):
            origins_skipped += 1
            continue

        origin = origin_entry.get("origin")
        raw_local_storage = origin_entry.get("localStorage", [])
        if not origin or not isinstance(raw_local_storage, list):
            origins_skipped += 1
            continue

        local_storage_items: list[dict[str, str]] = []
        for item in raw_local_storage:
            if not isinstance(item, dict):
                local_storage_items_skipped += 1
                continue
            name = item.get("name")
            value = item.get("value")
            if not name or value is None:
                local_storage_items_skipped += 1
                continue
            local_storage_items.append({"name": str(name), "value": str(value)})

        local_storage_items_loaded += len(local_storage_items)
        local_storage_origins.append(
            {"origin": str(origin), "local_storage": local_storage_items}
        )

    metadata = {
        "storage_state_file": str(storage_state_file),
        "storage_state_cookies_loaded_count": len(valid),
        "storage_state_cookies_skipped_count": skipped,
        "storage_state_origins_loaded_count": len(local_storage_origins),
        "storage_state_origins_skipped_count": origins_skipped,
        "storage_state_local_storage_items_loaded_count": local_storage_items_loaded,
        "storage_state_local_storage_items_skipped_count": local_storage_items_skipped,
    }
    return valid, local_storage_origins, metadata


def restore_local_storage(
    page: Any,
    local_storage_origins: list[dict[str, Any]],
) -> dict[str, Any]:
    restored_origins = 0
    restored_items = 0
    failed_origins: list[dict[str, str]] = []

    for local_storage_origin in local_storage_origins:
        origin = local_storage_origin.get("origin")
        items = local_storage_origin.get("local_storage", [])
        if not origin or not items:
            continue

        try:
            page.goto(str(origin), wait_until="domcontentloaded", timeout=60_000)
            page.evaluate(
                """entries => {
                    for (const entry of entries) {
                        window.localStorage.setItem(entry.name, entry.value);
                    }
                }""",
                items,
            )
            restored_origins += 1
            restored_items += len(items)
        except Exception as exc:
            failed_origins.append(
                {"origin": str(origin), "error": f"{type(exc).__name__}: {exc}"[:240]}
            )

    return {
        "storage_state_local_storage_origins_restored_count": restored_origins,
        "storage_state_local_storage_items_restored_count": restored_items,
        "storage_state_local_storage_restore_failures_count": len(failed_origins),
        "storage_state_local_storage_restore_failures": failed_origins,
    }


def apply_session_to_context(
    context: Any,
    page: Any,
    storage_state_file: Path | None,
    cookies_file: Path | None,
) -> None:
    state_origins: list[dict[str, Any]] = []

    if storage_state_file:
        if not storage_state_file.exists():
            raise FileNotFoundError(f"Storage state file not found: {storage_state_file}")
        state_cookies, state_origins, _metadata = load_storage_state_payload(storage_state_file)
        if state_cookies:
            context.add_cookies(state_cookies)

    if cookies_file:
        if not cookies_file.exists():
            raise FileNotFoundError(f"Cookies file not found: {cookies_file}")
        cookies_to_add, _metadata = load_playwright_cookies(cookies_file)
        if cookies_to_add:
            context.add_cookies(cookies_to_add)

    if storage_state_file and storage_state_file.exists() and state_origins:
        restore_local_storage(page=page, local_storage_origins=state_origins)
