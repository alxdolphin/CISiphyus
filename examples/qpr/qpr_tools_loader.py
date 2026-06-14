"""Load bundled qpr_tools from the QPR example directory."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

QPR_DIR = Path(__file__).resolve().parent
REPO_ROOT = QPR_DIR.parents[1]
FIXTURES_DIR = QPR_DIR / "fixtures"


def qpr_tools_path() -> Path:
    return QPR_DIR / "qpr_tools.py"


def load_qpr_tools() -> ModuleType:
    path = qpr_tools_path()
    if not path.is_file():
        raise FileNotFoundError(
            f"qpr_tools.py not found at {path}. "
            "The bundled QPR tools module is missing from examples/qpr/."
        )
    module_name = "qpr_tools_bundled"
    if module_name in sys.modules:
        return sys.modules[module_name]
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load qpr_tools from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def default_template_path() -> Path:
    return FIXTURES_DIR / "qpr_import_template.xlsx"


def default_deadlines_path() -> Path:
    return FIXTURES_DIR / "reporting_deadlines_minimal.xlsx"


def default_site_staff_list_path() -> Path:
    return FIXTURES_DIR / "site_staff_list_minimal.xlsx"
