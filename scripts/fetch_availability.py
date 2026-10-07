"""Capture today's GPU availability, or rebuild the history.

Usage:
  python scripts/fetch_availability.py
  python scripts/fetch_availability.py --history --chart "B200 Availability (3Fourteen chart, digitized).csv"
  python scripts/fetch_availability.py --farm-history --akash-history
  python scripts/fetch_availability.py --import research_obs.csv
  python scripts/fetch_availability.py --wide

With no flag it pulls RunPod, Vast, Shadeform, the newest gpuhunt RunPod
catalog, the last two days from 500.farm, the last three from Akash, and
io.net now. It adds them to data/out/gpu_availability.csv and to the full copy
in data/cache/, writes a snapshot in data/snapshots/, and rebuilds the wide
files.

--history re-reads one gpuhunt catalog per day from the list in data/fixtures/,
about a thousand paced requests. --farm-history re-pulls 500.farm from Oct 2021
in fourteen queries and rewrites data/out/gpu_prices_500farm.csv and the full
tape in data/cache/. --akash-history re-pulls Akash's daily breakdown from 2024.
None of these is part of the daily run.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.availability import (
    GPUHUNT_VERSIONS,
    append_rows,
    build_wide,
    collect_live,
    fetch_500farm_history,
    fetch_akash_history,
    gpuhunt_history,
    import_observations,
    parse_digitized_chart,
    read_versions,
    write_snapshot,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="GPU availability capture and history.")
    parser.add_argument("--history", action="store_true", help="re-read one gpuhunt RunPod catalog per day")
    parser.add_argument("--chart", help="digitized 3Fourteen / Silicon Data B200 chart CSV")
    parser.add_argument("--farm-history", action="store_true", help="re-pull 500.farm's Vast counts and price tape")
    parser.add_argument("--akash-history", action="store_true", help="re-pull Akash's daily GPU breakdown")
    parser.add_argument("--import", dest="imports", action="append", default=[], help="research CSV in the long-file columns")
    parser.add_argument("--wide", action="store_true", help="only rebuild the wide files")
    args = parser.parse_args()

    rows = []
    if args.chart:
        rows += parse_digitized_chart(Path(args.chart).read_text(encoding="utf-8"), source_url=Path(args.chart).name)
    for name in args.imports:
        rows += import_observations(Path(name).read_text(encoding="utf-8"))
    if args.history:
        rows += gpuhunt_history(read_versions(GPUHUNT_VERSIONS.read_text(encoding="utf-8")))
    if args.farm_history:
        rows += fetch_500farm_history()
    if args.akash_history:
        rows += fetch_akash_history()
    backfill = args.history or args.chart or args.imports or args.wide or args.farm_history or args.akash_history
    if not backfill:
        snapshot = collect_live()
        print(write_snapshot(snapshot))
        for name, source in snapshot["sources"].items():
            print(f"{name:15} {source.get('http_status')}  {source.get('error') or 'ok'}")
        rows = snapshot["rows"]
    if rows:
        path, added = append_rows(rows)
        print(f"{path}  rows read {len(rows)}  new keys {added}")
    for path in build_wide():
        print(path)


if __name__ == "__main__":
    main()
