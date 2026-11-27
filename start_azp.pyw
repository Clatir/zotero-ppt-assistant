"""Punkt startowy okienka AZP bez okna konsoli w tle (D16)."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from azp.main import main

if __name__ == "__main__":
    main()
