from __future__ import annotations

import json
import os
import platform
import re
import shutil
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import yaml

import auth_cookies


def find_project_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / "pyproject.toml").is_file():
            return parent
    return here.parents[2]


ROOT = find_project_root()
CONFIG_DIR = ROOT / "config"
REPORTS_PATH = CONFIG_DIR / "reports.yaml"
ENV_PATH = CONFIG_DIR / "export_urls.env"
ARTIFACTS = ROOT / "artifacts"
LATEST_DIR = ARTIFACTS / "latest"
AUTH_BOOTSTRAP_MARKER_NAME = ".cisiphyus_auth_bootstrap.json"
COOKIES_EXPORT_PATH = CONFIG_DIR / "CISDM_cookies.json"
DEFAULT_CHROME_PROFILE_DIRECTORY = "Profile 1"
BOOTSTRAP_COMMAND = "python run.py --bootstrap"

BASE_URL = "https://cw.caseworthy.net/cis_prod.caseworthy"
DOWNLOAD_TIMEOUT_MS = 300_000
REPORT_VIEWER_DOWNLOAD_TIMEOUT_MS = 300_000
NAVIGATION_TIMEOUT_MS = 120_000


def migrate_legacy_layout() -> None:
    legacy_chrome = ROOT / "src" / "config" / "chrome-user-data"
    new_chrome = CONFIG_DIR / "chrome-user-data"
    if legacy_chrome.is_dir() and not new_chrome.exists():
        new_chrome.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(legacy_chrome), str(new_chrome))


def load_env_file(path: Path) -> None:
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if (value.startswith("'") and value.endswith("'")) or (
            value.startswith('"') and value.endswith('"')
        ):
            value = value[1:-1]
        os.environ.setdefault(key, value)


def load_reports(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Missing reports.yaml: {path}")

    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}

    reports = data.get("reports")
    if not isinstance(reports, dict):
        raise ValueError("reports.yaml must contain top-level key: reports")
    return reports


def redact_url(url: str) -> str:
    split = urlsplit(url)
    safe_pairs: list[tuple[str, str]] = []
    for key, value in parse_qsl(split.query, keep_blank_values=True):
        if key.lower() in {"formid", "exportrowscount"}:
            safe_pairs.append((key, value))
        else:
            safe_pairs.append((key, "<redacted>"))
    return urlunsplit((split.scheme, split.netloc, split.path, urlencode(safe_pairs), ""))


def redact_sensitive_text(text: str) -> str:
    redacted = text
    patterns = [
        r"LS=[^\"'&<>\s]+",
        r"CurrentUserID(?:%3D|=)?\d+",
        r"UsersID(?:%3D|=)?\d+",
        r"@CurrentUserID(?:%3D|=)?\d+",
        r"%2540CurrentUserID%3D\d+",
        r"%2540UsersID%3D\d+",
    ]
    for pattern in patterns:
        redacted = re.sub(pattern, "<redacted>", redacted, flags=re.IGNORECASE)
    return redacted[:1200]


def system_chrome_user_data_dir() -> Path | None:
    system_name = platform.system()
    if system_name == "Windows":
        local_app_data = os.environ.get("LOCALAPPDATA")
        if not local_app_data:
            return None
        return Path(local_app_data) / "Google" / "Chrome" / "User Data"
    if system_name == "Linux":
        return Path.home() / ".config" / "google-chrome"
    if system_name == "Darwin":
        return Path.home() / "Library" / "Application Support" / "Google" / "Chrome"
    return None


def default_chrome_user_data_dir() -> Path:
    return CONFIG_DIR / "chrome-user-data"


def is_system_chrome_user_data_dir(candidate: Path) -> bool:
    system_dir = system_chrome_user_data_dir()
    if system_dir is None:
        return False
    return os.path.normcase(str(candidate.resolve(strict=False))) == os.path.normcase(
        str(system_dir.resolve(strict=False))
    )


def auth_bootstrap_marker_path(chrome_user_data_dir: Path) -> Path:
    return chrome_user_data_dir / AUTH_BOOTSTRAP_MARKER_NAME


def bootstrap_profile_exists(chrome_user_data_dir: Path) -> bool:
    marker_path = auth_bootstrap_marker_path(chrome_user_data_dir)
    if not marker_path.is_file():
        return False

    try:
        marker = json.loads(marker_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False

    if not isinstance(marker, dict) or not marker.get("disk_verified_at"):
        return False

    profile_dir = chrome_user_data_dir / DEFAULT_CHROME_PROFILE_DIRECTORY
    return auth_cookies.disk_has_required_cookies(profile_dir)


def latest_report_dir(report_id: str) -> Path:
    return LATEST_DIR / report_id


def latest_raw_path(report_id: str) -> Path:
    return latest_report_dir(report_id) / "raw.xlsx"
