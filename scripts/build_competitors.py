"""Write competitor capacity within 200 km of each Nebius campus.

Usage:
  python scripts/build_competitors.py
  python scripts/build_competitors.py research/out_*.jsonl ...

With no files, data/competitor_projects.jsonl is read as it is. Research
slices given on the command line are joined first, duplicates dropped, and
the result replaces data/competitor_projects.jsonl.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.competitors import build, load_projects, merge_projects, read_slice, write_projects


def main() -> None:
    slices = [Path(name) for name in sys.argv[1:]]
    if slices:
        ordered = sorted(path for path in slices if path.suffix == ".json") + sorted(path for path in slices if path.suffix == ".jsonl")
        projects, dropped = merge_projects([read_slice(path) for path in ordered])
        print(write_projects(projects))
        print(f"dropped {len(dropped)}")
    projects = load_projects()
    print(f"projects {len(projects)}")
    print(f"phases {sum(len(project.get('phases') or []) for project in projects)}")
    for path in build(projects):
        print(path)


if __name__ == "__main__":
    main()
