from __future__ import annotations

import os
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
SIBLING_PUBIFY_DATA = ROOT.parent / "pubify-data" / "src"
SIBLING_PUBIFY_MPL = ROOT.parent / "pubify-mpl" / "src"
MPLCONFIGDIR = ROOT / ".pytest_cache" / "matplotlib"
MPLCONFIGDIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", str(MPLCONFIGDIR))

for path in (SRC, SIBLING_PUBIFY_DATA, SIBLING_PUBIFY_MPL):
    if path.exists() and str(path) not in sys.path:
        sys.path.insert(0, str(path))
