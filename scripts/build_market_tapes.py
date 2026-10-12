"""Build the peer GPU-price tables and the wholesale power-cost table.

Usage:
  python scripts/build_market_tapes.py
"""

from __future__ import annotations

import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nbis_scrapes.market_tapes import (
    CACHE,
    parse_coreweave,
    parse_crusoe,
    parse_lambda,
    write_power_tape,
    write_price_tape,
)

PAGES = {
    "crusoe": "https://www.crusoe.ai/cloud/pricing",
    "lambda": "https://lambda.ai/pricing",
    "coreweave": "https://www.coreweave.com/pricing",
}
PARSERS = {"crusoe": parse_crusoe, "lambda": parse_lambda, "coreweave": parse_coreweave}
EMBER = "https://files.ember-energy.org/public-downloads/price/outputs/european_wholesale_electricity_price_data_monthly.csv"
EIA = {
    "ice_electric-2024final.xlsx": "https://www.eia.gov/electricity/wholesale/xls/archive/ice_electric-2024final.xlsx",
    "ice_electric-2025final.xlsx": "https://www.eia.gov/electricity/wholesale/xls/archive/ice_electric-2025final.xlsx",
    "ice_electric-2026.xlsx": "https://www.eia.gov/electricity/wholesale/xls/ice_electric-2026.xlsx",
}
UA = {"User-Agent": "NBIS-Scrapes/0.1 (personal research)"}


def _get(url: str) -> bytes:
    request = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def main() -> None:
    CACHE.mkdir(parents=True, exist_ok=True)
    captured = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    peers = []
    for name, url in PAGES.items():
        path = CACHE / f"{name}_pricing.html"
        if not path.exists():
            path.write_bytes(_get(url))
        html = path.read_text(encoding="utf-8", errors="replace")
        peers.extend(PARSERS[name](html))
        print(name, [(row["sku"], row["usd_per_gpu_hour"]) for row in PARSERS[name](html)])
    print(write_price_tape(peers, captured))

    ember = CACHE / "ember_monthly.csv"
    if not ember.exists():
        ember.write_bytes(_get(EMBER))
    eia_paths = []
    for name, url in EIA.items():
        path = CACHE / name
        if not path.exists():
            path.write_bytes(_get(url))
        eia_paths.append(path)
    print(write_power_tape(ember, eia_paths))


if __name__ == "__main__":
    main()
