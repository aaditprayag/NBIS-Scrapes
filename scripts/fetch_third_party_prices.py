"""Add today's third-party Nebius GPU prices and check them against the list.

Usage:
  python scripts/fetch_third_party_prices.py
  python scripts/fetch_third_party_prices.py --import observations.csv ...
  python scripts/fetch_third_party_prices.py --offline

The default run adds the latest gpuhunt catalog and any SkyPilot commits after
the last one on file. --import adds rows from observation CSVs in the shared
research layout. --offline only re-checks the file against the list.
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.third_party_prices import (
    GPUHUNT,
    SKYPILOT,
    fetch_gpuhunt_latest,
    fetch_skypilot,
    from_observation_file,
    load,
    merge,
    write_outputs,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Third-party Nebius GPU prices.")
    parser.add_argument("--import", dest="imports", nargs="*", default=[])
    parser.add_argument("--offline", action="store_true")
    args = parser.parse_args()

    rows = load()
    new = []
    for name in args.imports:
        new.extend(from_observation_file(Path(name)))
    if not args.offline and not args.imports:
        gpuhunt_days = {row["date"] for row in rows if row["source"] == GPUHUNT}
        latest = fetch_gpuhunt_latest()
        new.extend(row for row in latest if row["date"] not in gpuhunt_days)
        skypilot_days = [row["date"] for row in rows if row["source"] == SKYPILOT]
        new.extend(fetch_skypilot(max(skypilot_days) if skypilot_days else ""))
    rows = merge(rows, new)
    for path in write_outputs(rows):
        print(path)
    checked = load()
    print(f"rows {len(checked)}")
    print(f"dates {len({row['date'] for row in checked})}")
    for verdict, count in Counter(row["matches_nebius_list"] for row in checked).most_common():
        print(f"  {verdict}: {count}")


if __name__ == "__main__":
    main()
