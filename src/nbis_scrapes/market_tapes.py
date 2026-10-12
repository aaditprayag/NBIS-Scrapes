"""Published on-demand GPU-hour prices, and the wholesale electricity cost of one GPU-hour.

The price tables read the published GPU-hour cards for Nebius, CoreWeave,
Crusoe, and Lambda. CoreWeave quotes a node, so the per-GPU figure divides
that node price by the GPU count printed next to it.

The power table is the cost of electricity under the sites that can bill
inside a year. Finland, France, the UK, Spain, and Estonia are Ember's
monthly load-weighted day-ahead prices, averaged to a quarter. PJM West is
the EIA reprint of ICE peak prices, averaged across the days in the quarter.
Peak PJM is not the same product as an all-hours Nordic day-ahead price.
"""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from nbis_scrapes.wide_prices import OUT, WIDE, _clean, _write_wide, fiscal_quarter, quarter_end

ROOT = Path(__file__).resolve().parents[2]
CACHE = ROOT / "data" / "cache"

# H100 SXM board power, and a facility overhead. Stated in the row name.
H100_KW = 0.7
PUE = {"Finland": 1.15, "France": 1.25, "United Kingdom": 1.25, "Spain": 1.25, "Estonia": 1.2, "PJM West peak": 1.3}

COREWEAVE_GPUS = {
    "NVIDIA HGX H100": ("H100", 8),
    "NVIDIA HGX H200": ("H200", 8),
    "NVIDIA HGX B200": ("B200", 8),
    "NVIDIA HGX B300": ("B300", 8),
    "NVIDIA GB200 NVL72": ("GB200", 4),
    "NVIDIA GH200": ("GH200", 1),
}


def _lines(html: str) -> list[str]:
    plain = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.IGNORECASE)
    plain = re.sub(r"<[^>]+>", "\n", plain)
    return [re.sub(r"\s+", " ", line).strip() for line in plain.splitlines() if line.strip()]


def parse_crusoe(html: str) -> list[dict]:
    lines = _lines(html)
    found = []
    seen = set()
    for index, line in enumerate(lines):
        match = re.fullmatch(r"NVIDIA (H100|H200|B200|L40S|A100)", line)
        if not match:
            continue
        window = " ".join(lines[index:index + 6])
        price = re.search(r"\$([0-9]+(?:\.[0-9]+)?)/GPU-hr", window)
        if not price:
            continue
        sku = match.group(1)
        if sku in seen:
            continue
        seen.add(sku)
        found.append({"company": "Crusoe", "sku": sku, "usd_per_gpu_hour": float(price.group(1)), "basis": "published per GPU-hour"})
    return found


def parse_coreweave(html: str) -> list[dict]:
    lines = _lines(html)
    found = []
    seen = set()
    for index, line in enumerate(lines):
        if line != "On-Demand Price:":
            continue
        ahead = " ".join(lines[index:index + 3])
        price = re.search(r"\$([0-9]+(?:\.[0-9]+)?)", ahead)
        if not price:
            continue
        sku = None
        gpus = None
        for prior in reversed(lines[max(0, index - 8):index]):
            spec = COREWEAVE_GPUS.get(prior)
            if spec:
                sku, gpus = spec
                break
        if sku is None or sku in seen:
            continue
        node = float(price.group(1))
        seen.add(sku)
        found.append({
            "company": "CoreWeave",
            "sku": sku,
            "usd_per_gpu_hour": round(node / gpus, 4),
            "basis": f"on-demand node price {node} divided by {gpus} GPUs",
        })
    return found


def parse_lambda(html: str) -> list[dict]:
    lines = _lines(html)
    start = next((i for i, line in enumerate(lines) if line == "Instances pricing"), None)
    if start is None:
        return []
    section = lines[start:start + 80]
    found = []
    seen = set()
    pending = None
    for line in section:
        if line.startswith("NVIDIA B200"):
            pending = "B200"
        elif line.startswith("NVIDIA H100 SXM"):
            pending = "H100"
        elif line.startswith("NVIDIA H100 PCIe"):
            pending = "H100 PCIe"
        elif line.startswith("NVIDIA GH200"):
            pending = "GH200"
        price = re.fullmatch(r"\$([0-9]+(?:\.[0-9]+)?)", line)
        if price and pending and pending not in seen:
            seen.add(pending)
            found.append({
                "company": "Lambda",
                "sku": pending,
                "usd_per_gpu_hour": float(price.group(1)),
                "basis": "first on-demand price in the instances section",
            })
            pending = None
    return found


def nebius_history(path: Path | None = None) -> list[dict]:
    path = path or (WIDE / "gpu_page_on_demand.csv")
    if not path.exists():
        return []
    with path.open(encoding="utf-8") as handle:
        rows = list(csv.reader(handle))
    if len(rows) < 3:
        return []
    dates = rows[1][1:]
    found = []
    for row in rows[2:]:
        sku = row[0]
        for day, value in zip(dates, row[1:]):
            if not value:
                continue
            found.append({
                "company": "Nebius",
                "sku": sku,
                "captured_on": day,
                "usd_per_gpu_hour": float(value),
                "basis": "on-demand column of nebius.com/prices",
            })
    return found


def _quarter_key(day: datetime) -> str:
    quarter = (day.month - 1) // 3 + 1
    return f"{day.year}-Q{quarter}"


def ember_quarterly(path: Path, countries: list[str]) -> dict[str, dict[str, float]]:
    buckets: dict[str, dict[str, list[float]]] = defaultdict(lambda: defaultdict(list))
    with path.open(encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            if row["Country"] not in countries:
                continue
            day = datetime.strptime(row["Date"], "%Y-%m-%d")
            if day < datetime(2024, 1, 1):
                continue
            buckets[row["Country"]][_quarter_key(day)].append(float(row["Price (EUR/MWhe)"]))
    out = {}
    for country, quarters in buckets.items():
        out[country] = {}
        for key, values in quarters.items():
            if len(values) < 3:
                continue
            out[country][key] = sum(values) / len(values)
    return out


def eia_pjm_quarterly(paths: list[Path]) -> dict[str, float]:
    buckets: dict[str, list[float]] = defaultdict(list)
    for path in paths:
        book = __import__("openpyxl").load_workbook(path, read_only=True, data_only=True)
        sheet = book[book.sheetnames[0]]
        header = [cell.value for cell in next(sheet.iter_rows(max_row=1))]
        hub_i = header.index("Price hub")
        date_i = header.index("Delivery start date")
        price_i = next(i for i, name in enumerate(header) if name and "Wtd avg" in str(name))
        for row in sheet.iter_rows(min_row=2, values_only=True):
            hub = row[hub_i]
            if not hub or "PJM" not in str(hub):
                continue
            day = row[date_i]
            price = row[price_i]
            if not isinstance(day, datetime) or not isinstance(price, (int, float)):
                continue
            if day < datetime(2024, 1, 1):
                continue
            buckets[_quarter_key(day)].append(float(price))
        book.close()
    return {key: sum(values) / len(values) for key, values in buckets.items() if values}


def quarter_end_from_key(key: str) -> str:
    year, quarter = key.split("-Q")
    return quarter_end(f"Q{quarter} {year}")


def write_price_tape(peer_rows: list[dict], captured_on: str) -> list[Path]:
    nebius = nebius_history()
    peers = []
    for row in peer_rows:
        peers.append({**row, "captured_on": captured_on})
    all_rows = nebius + peers
    long_path = OUT / "neocloud_prices.csv"
    long_path.parent.mkdir(parents=True, exist_ok=True)
    with long_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["company", "sku", "captured_on", "usd_per_gpu_hour", "basis"])
        writer.writeheader()
        writer.writerows(all_rows)

    by_company: dict[str, dict[str, dict[str, str]]] = defaultdict(lambda: defaultdict(dict))
    for row in all_rows:
        by_company[row["company"]][row["sku"]][row["captured_on"]] = _clean(row["usd_per_gpu_hour"])
    written = [long_path]
    for company, grid in by_company.items():
        columns = sorted({day for values in grid.values() for day in values})
        path = WIDE / f"neocloud_{company.lower()}.csv"
        _write_wide(path, "USD per GPU-hour", columns, grid)
        written.append(path)
    return written


def write_power_tape(ember_path: Path, eia_paths: list[Path]) -> Path:
    countries = ["Finland", "France", "United Kingdom", "Spain", "Estonia"]
    quarterly = ember_quarterly(ember_path, countries)
    pjm = eia_pjm_quarterly(eia_paths)
    grid: dict[str, dict[str, str]] = {}
    for country, values in quarterly.items():
        name = f"{country} day-ahead, EUR/MWh"
        grid[name] = {quarter_end_from_key(key): _clean(price) for key, price in values.items()}
        pue = PUE[country]
        cost_name = f"{country} implied H100 power cost, EUR/GPU-hour"
        grid[cost_name] = {
            day: _clean(float(price) * H100_KW / 1000 * pue)
            for day, price in grid[name].items()
        }
    grid["PJM West peak, USD/MWh"] = {quarter_end_from_key(key): _clean(price) for key, price in pjm.items()}
    grid["PJM West implied H100 power cost, USD/GPU-hour"] = {
        day: _clean(float(price) * H100_KW / 1000 * PUE["PJM West peak"])
        for day, price in grid["PJM West peak, USD/MWh"].items()
    }
    columns = sorted({day for values in grid.values() for day in values})
    path = WIDE / "power_cost.csv"
    _write_wide(path, "Wholesale price, then implied GPU-hour power cost", columns, grid)
    return path


def fiscal_labels(columns: list[str]) -> list[str]:
    return [fiscal_quarter(day) for day in columns]
