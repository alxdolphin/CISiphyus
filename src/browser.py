from __future__ import annotations

import os
import platform
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Error as PlaywrightError

import config


def _running_under_wsl() -> bool:
    try:
        with open("/proc/version", encoding="utf-8") as handle:
            return "microsoft" in handle.read().lower()
    except OSError:
        return False


def warn_headed_linux_gui_environment(use_case: str) -> None:
    if platform.system() != "Linux":
        return
    display = os.environ.get("DISPLAY")
    wayland = os.environ.get("WAYLAND_DISPLAY")
    if not display and not wayland:
        print(
            f"Warning ({use_case}): DISPLAY and WAYLAND_DISPLAY are unset — "
            "Playwright Chromium likely has no screen to draw on. "
            "Use a WSL2 terminal with WSLg (Windows 11) or set DISPLAY (often `:0` under WSLg).",
            file=sys.stderr,
        )
        return
    if _running_under_wsl():
        print(
            f"Note ({use_case}): WSL2 detected. Uses Playwright bundled Chromium via WSLg.",
            file=sys.stderr,
        )


def chromium_profile_launch_args(chrome_profile_directory: str, *, headed: bool) -> list[str]:
    args: list[str] = [f"--profile-directory={chrome_profile_directory}"]
    if headed and platform.system() == "Linux":
        args.extend(["--window-position=80,80", "--window-size=1280,900"])
    return args


def reject_system_chrome_dir(chrome_user_data_dir: Path) -> None:
    if config.is_system_chrome_user_data_dir(chrome_user_data_dir):
        raise RuntimeError(
            "Chrome user data dir points to the system default profile root. "
            f"Use the dedicated automation profile at {config.default_chrome_user_data_dir()}."
        )


def launch_persistent_context(
    p: Any,
    *,
    chrome_user_data_dir: Path,
    chrome_profile_directory: str,
    headed: bool,
    accept_downloads: bool,
) -> Any:
    launch_kwargs: dict[str, Any] = {
        "user_data_dir": str(chrome_user_data_dir),
        "headless": not headed,
        "accept_downloads": accept_downloads,
        "args": chromium_profile_launch_args(chrome_profile_directory, headed=headed),
    }
    try:
        return p.chromium.launch_persistent_context(**launch_kwargs)
    except PlaywrightError as exc:
        raise RuntimeError(f"browser_launch_failed: {exc}") from exc
