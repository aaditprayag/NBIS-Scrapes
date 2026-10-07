"""Pull Wayback on-demand NVIDIA prices for CoreWeave, Crusoe, and Lambda.

Usage:
  python scripts/fetch_peer_prices.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.peer_history import collect, write_history


def main() -> None:
    rows = collect()
    paths = write_history(rows)
    companies = sorted({row["company"] for row in rows})
    print(f"rows {len(rows)} companies {companies}")
    for company in companies:
        skus = sorted({row["sku"] for row in rows if row["company"] == company})
        days = sorted({row["captured_on"] for row in rows if row["company"] == company})
        print(f"  {company}: {len(skus)} GPUs, {days[0] if days else '-'} to {days[-1] if days else '-'} ({len(days)} dates)")
        print("   ", ", ".join(skus))
    for path in paths:
        print(path)


if __name__ == "__main__":
    main()
