from __future__ import annotations

import sqlite3
import time
from pathlib import Path
from typing import Any, Sequence

REQUIRED_COOKIE_NAMES = ("Auth_CaseWorthy", "Context_CaseWorthy")
CASEWORTHY_DOMAIN_SUFFIX = "caseworthy.net"


class RequiredCookiesError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        missing: tuple[str, ...] = (),
        diagnostics: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.missing = missing
        self.diagnostics = diagnostics or {}


def _domain_matches_caseworthy(domain: str) -> bool:
    normalized = domain.lower().strip().lstrip(".")
    return normalized == CASEWORTHY_DOMAIN_SUFFIX or normalized.endswith(
        f".{CASEWORTHY_DOMAIN_SUFFIX}"
    )


def _filter_caseworthy_cookies(
    cookies: list[dict[str, Any]],
    required_names: tuple[str, ...] = REQUIRED_COOKIE_NAMES,
) -> list[dict[str, Any]]:
    matched: list[dict[str, Any]] = []
    for cookie in cookies:
        if not isinstance(cookie, dict):
            continue
        name = str(cookie.get("name", "")).strip()
        domain = str(cookie.get("domain", "")).strip()
        if name in required_names and _domain_matches_caseworthy(domain):
            matched.append(cookie)
    return matched


def cookie_inventory(cookies: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Redacted view of every cookie (name/domain/expires only, never values)."""
    inventory: list[dict[str, Any]] = []
    for cookie in cookies:
        if not isinstance(cookie, dict):
            continue
        entry: dict[str, Any] = {
            "name": str(cookie.get("name", "")),
            "domain": str(cookie.get("domain", "")),
        }
        expires = cookie.get("expires")
        if expires is not None:
            try:
                entry["expires"] = float(expires)
            except (TypeError, ValueError):
                pass
        inventory.append(entry)
    return inventory


def cookie_metadata(
    cookies: list[dict[str, Any]],
    required_names: tuple[str, ...] = REQUIRED_COOKIE_NAMES,
) -> dict[str, Any]:
    entries: list[dict[str, Any]] = []
    for cookie in _filter_caseworthy_cookies(cookies, required_names):
        entry: dict[str, Any] = {
            "name": str(cookie.get("name", "")),
            "domain": str(cookie.get("domain", "")),
        }
        expires = cookie.get("expires")
        if expires is not None:
            try:
                entry["expires"] = float(expires)
            except (TypeError, ValueError):
                pass
        entries.append(entry)

    present_names = {entry["name"] for entry in entries}
    return {
        "required_cookies_present": sorted(present_names),
        "required_cookies_missing": sorted(
            name for name in required_names if name not in present_names
        ),
        "cookie_domains": sorted({entry["domain"] for entry in entries}),
        "required_cookie_count": len(entries),
    }


def _collect_context_cookies(
    context: Any,
    cookie_urls: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    merged: dict[tuple[str, str], dict[str, Any]] = {}

    def _merge(batch: list[dict[str, Any]]) -> None:
        for cookie in batch:
            if not isinstance(cookie, dict):
                continue
            key = (str(cookie.get("name", "")), str(cookie.get("domain", "")))
            merged[key] = cookie

    _merge(context.cookies())
    for url in cookie_urls or ():
        url = url.strip()
        if not url:
            continue
        try:
            _merge(context.cookies([url]))
        except Exception:
            continue

    return list(merged.values())


def export_caseworthy_cookies(
    context: Any,
    cookie_urls: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Full cookie dicts (values included) for every CaseWorthy-domain cookie.
    Output is Playwright add_cookies-compatible; treat it as a secret."""
    return [
        cookie
        for cookie in _collect_context_cookies(context, cookie_urls)
        if isinstance(cookie, dict)
        and _domain_matches_caseworthy(str(cookie.get("domain", "")))
    ]


def context_cookie_inventory(
    context: Any,
    cookie_urls: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    return cookie_inventory(_collect_context_cookies(context, cookie_urls))


def wait_for_cookie_in_context(
    context: Any,
    cookie_name: str,
    *,
    timeout_seconds: float,
    poll_interval_seconds: float = 1.0,
    cookie_urls: Sequence[str] | None = None,
) -> bool:
    """Poll until the named cookie appears; CISDM issues auth/context cookies
    asynchronously after page load, so a single check races against the SPA."""
    deadline = time.monotonic() + timeout_seconds
    while True:
        names = {
            str(cookie.get("name", ""))
            for cookie in _collect_context_cookies(context, cookie_urls)
        }
        if cookie_name in names:
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(poll_interval_seconds)


def assert_required_cookies_in_context(
    context: Any,
    *,
    cookie_urls: Sequence[str] | None = None,
    required_names: tuple[str, ...] = REQUIRED_COOKIE_NAMES,
) -> dict[str, Any]:
    collected = _collect_context_cookies(context, cookie_urls)
    metadata = cookie_metadata(collected, required_names)
    missing = tuple(metadata["required_cookies_missing"])
    if missing:
        raise RequiredCookiesError(
            "Required CISDM cookies missing from browser context: "
            + ", ".join(missing),
            missing=missing,
            diagnostics={
                **metadata,
                "all_context_cookies": cookie_inventory(collected),
            },
        )
    return metadata


def _read_cookie_entries_safe(
    cookies_db: Path,
) -> tuple[list[dict[str, str]], str | None]:
    """Read cookie name/domain rows, treating any sqlite failure (locked,
    corrupt, unsupported schema) as an unusable-but-diagnosable profile."""
    if not cookies_db.is_file():
        return [], None

    try:
        connection = sqlite3.connect(f"file:{cookies_db}?mode=ro", uri=True)
        try:
            rows = connection.execute("SELECT name, host_key FROM cookies").fetchall()
        finally:
            connection.close()
    except sqlite3.Error as exc:
        return [], f"{type(exc).__name__}: {exc}"

    return [{"name": str(row[0]), "domain": str(row[1])} for row in rows], None


def read_cookie_entries_from_sqlite(cookies_db: Path) -> list[dict[str, str]]:
    entries, _error = _read_cookie_entries_safe(cookies_db)
    return entries


def read_cookie_names_from_sqlite(cookies_db: Path) -> set[str]:
    return {entry["name"] for entry in read_cookie_entries_from_sqlite(cookies_db)}


def assert_required_cookies_on_disk(profile_dir: Path) -> dict[str, Any]:
    cookies_db = profile_dir / "Cookies"
    if not cookies_db.is_file() or cookies_db.stat().st_size == 0:
        raise RequiredCookiesError(
            f"Chrome cookie database missing or empty: {cookies_db}",
            missing=REQUIRED_COOKIE_NAMES,
        )

    entries_on_disk, db_error = _read_cookie_entries_safe(cookies_db)
    names_on_disk = {entry["name"] for entry in entries_on_disk}
    missing = tuple(name for name in REQUIRED_COOKIE_NAMES if name not in names_on_disk)
    if missing:
        diagnostics: dict[str, Any] = {
            "cookies_db_path": str(cookies_db),
            "all_disk_cookies": entries_on_disk,
        }
        if db_error:
            diagnostics["cookies_db_error"] = db_error
        message = (
            f"Chrome cookie database unreadable ({db_error}): {cookies_db}"
            if db_error
            else "Required CISDM cookies missing from profile Cookies database: "
            + ", ".join(missing)
        )
        raise RequiredCookiesError(message, missing=missing, diagnostics=diagnostics)

    return {
        "cookies_db_path": str(cookies_db),
        "cookies_db_size_bytes": cookies_db.stat().st_size,
        "required_cookies_present_on_disk": sorted(
            name for name in REQUIRED_COOKIE_NAMES if name in names_on_disk
        ),
        "required_cookies_missing_on_disk": list(missing),
    }


def disk_has_required_cookies(profile_dir: Path) -> bool:
    cookies_db = profile_dir / "Cookies"
    if not cookies_db.is_file() or cookies_db.stat().st_size == 0:
        return False
    names_on_disk = read_cookie_names_from_sqlite(cookies_db)
    return all(name in names_on_disk for name in REQUIRED_COOKIE_NAMES)
