"""Write the console billing catalogue and the region first-seen table.

Usage:
  python scripts/build_console_catalog.py
  python scripts/build_console_catalog.py saved_pages.json
  python scripts/build_console_catalog.py export.csv --date 2026-10-06

With no file, the latest data/out/console_catalog_*.csv is re-read and only
the region table is rebuilt. A .json file is the listFullView pages saved
from a logged-in browser. See docs/console-catalog.md.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.console_catalog import OUT, parse_full_view, read_export, write_outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Nebius console billing catalogue.")
    parser.add_argument("path", nargs="?")
    parser.add_argument("--date", default="")
    args = parser.parse_args()

    if args.path:
        path = Path(args.path)
        day = args.date or datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc).strftime("%Y-%m-%d")
        if path.suffix == ".json":
            rows = parse_full_view(json.loads(path.read_text(encoding="utf-8")))
        else:
            rows = read_export(path)
    else:
        path = sorted(OUT.glob("console_catalog_*.csv"))[-1]
        day = path.stem.rsplit("_", 1)[-1]
        rows = read_export(path)
    catalog, regions = write_outputs(rows, day)
    print(catalog)
    print(regions)
    print(f"skus {len(rows)}")
    print(f"regions {len({row['region'] for row in rows if row['region']})}")


if __name__ == "__main__":
    main()
