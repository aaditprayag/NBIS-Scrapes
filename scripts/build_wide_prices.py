"""Write one wide CSV per price series.

Usage:
  python scripts/build_wide_prices.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.wide_prices import build_all


def main() -> None:
    for path in build_all():
        print(path)


if __name__ == "__main__":
    main()
