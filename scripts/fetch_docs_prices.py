"""Fetch Nebius's own GPU rate card and write data/out/docs_gpu_prices.csv.

Usage:
  python scripts/fetch_docs_prices.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.docs_prices import fetch_official, write_outputs


def main() -> None:
    prices, events = fetch_official()
    write_outputs(prices, events)
    print(f"prices {len(prices)}")
    print(f"changelog events {len(events)}")
    interesting = [
        row for row in prices
        if not row["preemptible"] and any(token in row["item"] for token in ("H100", "H200", "B200", "B300"))
    ]
    for row in interesting:
        print(
            f"{row['captured_on']}  {row['vintage']:7} {row['vintage_date'] or '-':10}  "
            f"{row['usd_per_gpu_hour']:6}  {row['platform_id'] or row['item']}"
        )


if __name__ == "__main__":
    main()
