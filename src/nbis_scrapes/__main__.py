"""Command line.

  python -m nbis_scrapes panel
  python -m nbis_scrapes map
  python -m nbis_scrapes demand
  python -m nbis_scrapes availability
  python -m nbis_scrapes all
"""

from __future__ import annotations

import argparse
import json

from nbis_scrapes import availability
from nbis_scrapes.demand import collect_live, write_snapshot
from nbis_scrapes.mapbuild import write_map
from nbis_scrapes.panel import write_outputs


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build the Nebius power tables, the campus map, a price snapshot, or today's GPU availability."
    )
    parser.add_argument("command", choices=("panel", "map", "demand", "availability", "all"))
    args = parser.parse_args()
    if args.command in ("panel", "all"):
        print(write_outputs())
    if args.command in ("map", "all"):
        print(write_map())
    if args.command in ("demand", "all"):
        snapshot = collect_live()
        path = write_snapshot(snapshot)
        brief = {
            name: {
                "http_status": source.get("http_status"),
                "error": source.get("error"),
                "parsed_keys": sorted((source.get("parsed") or {}).keys()) if isinstance(source.get("parsed"), dict) else (
                    len(source["parsed"]) if isinstance(source.get("parsed"), list) else None
                ),
            }
            for name, source in snapshot["sources"].items()
        }
        print(path)
        print(json.dumps(brief, indent=2))
    if args.command == "availability":
        snapshot = availability.collect_live()
        print(availability.write_snapshot(snapshot))
        print(availability.append_rows(snapshot["rows"])[0])
        for path in availability.build_wide():
            print(path)


if __name__ == "__main__":
    main()
