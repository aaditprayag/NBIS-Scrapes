"""Download unique Wayback copies of nebius.com/prices and write a price history.

Usage:
  python scripts/fetch_wayback_prices.py
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.price_history import fetch_all, on_demand_changes


def write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    stamp_file = ROOT / "data" / "fixtures" / "wayback_price_timestamps.txt"
    stamps = [line.strip() for line in stamp_file.read_text(encoding="utf-8").splitlines() if line.strip()]
    rows = fetch_all(stamps=stamps)
    changes = on_demand_changes(rows)
    out = ROOT / "data" / "out"
    write_csv(out / "gpu_price_history.csv", rows)
    write_csv(out / "gpu_price_changes.csv", changes)
    print(f"observations {len(rows)}")
    print(f"on-demand changes {len(changes)}")
    for row in changes:
        prev = "start" if row["previous_usd"] is None else row["previous_usd"]
        print(f"  {row['captured_on']}  {row['sku']:<12} {prev} -> {row['usd_per_gpu_hour']}")


if __name__ == "__main__":
    main()
