"""Fetch Token Factory list prices into CSVs.

Usage:
  python scripts/fetch_token_factory_prices.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.token_factory import fetch_history, price_changes, write_outputs


def main() -> None:
    rows = fetch_history()
    write_outputs(rows)
    changes = price_changes(rows)
    print(f"rows {len(rows)}")
    print(f"changes {len(changes)}")
    for row in changes:
        print(
            f"{row['captured_on']}  {row['model_id']}  "
            f"in {row['previous_input_usd_per_million']} -> {row['input_usd_per_million']}  "
            f"out {row['previous_output_usd_per_million']} -> {row['output_usd_per_million']}"
        )


if __name__ == "__main__":
    main()
