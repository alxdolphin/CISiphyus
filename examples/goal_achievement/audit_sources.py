"""Load metrics audit logic from the parent CIS repo."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType

TOOLS_DIR = Path(__file__).resolve().parents[3]


def _load_module(name: str, path: Path) -> ModuleType:
    if not path.is_file():
        raise FileNotFoundError(f"audit source not found: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load {name} from {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_audit_metrics_all() -> ModuleType:
    return _load_module(
        "cis_audit_metrics_all",
        TOOLS_DIR / "accreditation/scripts/audit_metrics_all.py",
    )
