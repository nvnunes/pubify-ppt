from __future__ import annotations

from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
SIBLING_PUBIFY_DATA = ROOT.parent / "pubify-data" / "src"

for path in (SRC, SIBLING_PUBIFY_DATA):
    if path.exists() and str(path) not in sys.path:
        sys.path.insert(0, str(path))
