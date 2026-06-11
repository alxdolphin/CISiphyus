#!/usr/bin/env python3
"""Run cisiphyus from the repository root without a prior install."""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from cli import main

if __name__ == "__main__":
    main()
