"""Load production qpr_tools from a CIS monorepo checkout."""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path
from types import ModuleType

REPO_ROOT = Path(__file__).resolve().parents[2]


def default_cis_monorepo_root() -> Path:
    explicit = (os.environ.get("CIS_MONOREPO_ROOT") or "").strip()
    if explicit:
        return Path(explicit).expanduser().resolve()
    # WHY: CISiphyus is usually nested at <CIS>/tools/CISiphyus
    return REPO_ROOT.parent.parent


def qpr_tools_path(cis_root: Path | None = None) -> Path:
    root = cis_root or default_cis_monorepo_root()
    return root / "tools" / "qpr" / "qpr_tools.py"


def load_qpr_tools(*, cis_root: Path | None = None) -> ModuleType:
    path = qpr_tools_path(cis_root)
    if not path.is_file():
        raise FileNotFoundError(
            f"qpr_tools.py not found at {path}. "
            "Set CIS_MONOREPO_ROOT to a CIS checkout containing tools/qpr/qpr_tools.py."
        )
    module_name = "qpr_tools_production"
    if module_name in sys.modules:
        return sys.modules[module_name]
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load qpr_tools from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def default_template_path(cis_root: Path | None = None) -> Path:
    root = cis_root or default_cis_monorepo_root()
    return root / "evaluation" / "reports" / "QPR" / "[TEMPLATE] QPR Import.xlsx"


def default_deadlines_path(cis_root: Path | None = None) -> Path:
    root = cis_root or default_cis_monorepo_root()
    local = root / "tools" / "qpr" / "local_inputs" / "SY25-26_ReportingDeadlines.xlsx"
    if local.is_file():
        return local
    return root / "evaluation" / "reports" / "SY25-26_ReportingDeadlines.xlsx"
