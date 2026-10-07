"""Find data centers within 60 km of each Nebius campus.

Usage:
  python scripts/build_neighbors.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.neighbors import build


def main() -> None:
    detail, summary = build()
    print(detail)
    print(summary)


if __name__ == "__main__":
    main()
