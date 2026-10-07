"""Turn the long price histories into wide CSVs.

Each file is one series. The series names run down the rows. Each date is a
column. The row above the dates is Nebius's fiscal quarter for that date.
Nebius reports on a calendar year, so April-June 2025 is Q2 '25.
"""

from __future__ import annotations

import csv
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "out"
WIDE = OUT / "wide"

SKU_ORDER = ["H100", "H200", "B200", "B300", "RTX PRO 6000", "L40S", "L40S-Intel", "L40S-AMD", "GB200", "GB300"]


def fiscal_quarter(iso_date: str) -> str:
    year, month, _day = (int(part) for part in iso_date.split("-"))
    quarter = (month - 1) // 3 + 1
    return f"Q{quarter} '{year % 100:02d}"


def _day_before(iso_date: str) -> str:
    current = datetime.strptime(iso_date, "%Y-%m-%d")
    return (current - timedelta(days=1)).strftime("%Y-%m-%d")


def _clean(value) -> str:
    if value is None or value == "":
        return ""
    number = float(value)
    text = f"{number:.4f}".rstrip("0").rstrip(".")
    return text


def _write_wide(path: Path, row_header: str, columns: list[str], grid: dict[str, dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Fiscal quarter", *[fiscal_quarter(day) for day in columns]])
        writer.writerow([row_header, *columns])
        for name in grid:
            writer.writerow([name, *[grid[name].get(day, "") for day in columns]])


def _sku_from_item(item: str, platform_id: str, vintage: str = "", vintage_date: str = "") -> str:
    text = f"{item} {platform_id}".upper().replace("®", "")
    component = item.endswith(". GPU") and vintage == "before" and vintage_date == "2025-10-01"
    if "H200" in text:
        base = "H200"
    elif "H100" in text:
        base = "H100"
    elif "B300" in text or "GB300" in text:
        base = "B300" if "B300" in text and "GB300" not in text else "GB300"
    elif "B200" in text or "GB200" in text:
        base = "B200" if "GB200" not in text else "GB200"
    elif "RTX" in text:
        base = "RTX PRO 6000"
    elif "L40" in text:
        if "AMD" in text:
            base = "L40S-AMD"
        elif "INTEL" in text:
            base = "L40S-Intel"
        else:
            base = "L40S"
    else:
        base = item
    if component and "component" not in base:
        return f"{base} GPU component"
    if "STANDALONE" in text:
        return f"Standalone {base}"
    if platform_id.endswith("sxm-a"):
        return f"{base} me-west1"
    return base


def docs_column(vintage: str, vintage_date: str, captured_on: str) -> str | None:
    if vintage == "before" and vintage_date:
        return _day_before(vintage_date)
    if vintage == "from" and vintage_date:
        return vintage_date
    if vintage == "undated" and captured_on:
        return captured_on
    return None


def pivot_docs(rows: list[dict], preemptible: bool, standalone: bool) -> tuple[list[str], dict[str, dict[str, str]]]:
    table: dict[str, dict[str, str]] = {}
    for row in rows:
        is_preempt = str(row["preemptible"]).lower() == "true"
        if is_preempt != preemptible:
            continue
        is_standalone = "standalone" in row["item"].lower()
        if is_standalone != standalone:
            continue
        if "GPU hour" not in row.get("item", "") and "NVLink" not in row["item"] and "L40" not in row["item"] and "RTX" not in row["item"]:
            if "GPU" not in row["item"]:
                continue
        column = docs_column(row["vintage"], row["vintage_date"], row["captured_on"])
        if column is None:
            continue
        name = _sku_from_item(row["item"], row["platform_id"], row["vintage"], row["vintage_date"])
        table.setdefault(name, {})
        table[name][column] = _clean(row["usd_per_gpu_hour"])
    columns = sorted({day for values in table.values() for day in values})
    ordered = _order_names(list(table))
    return columns, {name: table[name] for name in ordered}


def pivot_page(rows: list[dict], kind: str) -> tuple[list[str], dict[str, dict[str, str]]]:
    table: dict[str, dict[str, str]] = {}
    for row in rows:
        if row["kind"] != kind or str(row["gpus"]) not in {"1", "1.0"}:
            continue
        if row["usd_per_gpu_hour"] in ("", None):
            continue
        table.setdefault(row["sku"], {})
        table[row["sku"]][row["captured_on"]] = _clean(row["usd_per_gpu_hour"])
    columns = sorted({day for values in table.values() for day in values})
    ordered = _order_names(list(table))
    return columns, {name: table[name] for name in ordered}


def pivot_tokens(rows: list[dict], field: str) -> tuple[list[str], dict[str, dict[str, str]]]:
    table: dict[str, dict[str, str]] = {}
    for row in rows:
        value = row.get(field) or ""
        if value == "":
            continue
        table.setdefault(row["model_id"], {})
        table[row["model_id"]][row["captured_on"]] = _clean(value)
    columns = sorted({day for values in table.values() for day in values})
    return columns, {name: table[name] for name in sorted(table)}


def _order_names(names: list[str]) -> list[str]:
    rank = {name: index for index, name in enumerate(SKU_ORDER)}

    def key(name: str) -> tuple:
        base = name.replace(" GPU component", "").replace("Standalone ", "").replace(" me-west1", "")
        return (rank.get(base, 100), name)

    return sorted(names, key=key)


def build_all(out_dir: Path | None = None) -> list[Path]:
    wide = out_dir or WIDE
    docs = list(csv.DictReader((OUT / "docs_gpu_prices.csv").open(encoding="utf-8")))
    page = list(csv.DictReader((OUT / "gpu_price_history.csv").open(encoding="utf-8")))
    tokens = list(csv.DictReader((OUT / "token_factory_prices.csv").open(encoding="utf-8")))
    written = []

    specs = [
        ("gpu_rate_card_on_demand.csv", "USD per GPU-hour", *pivot_docs(docs, False, False)),
        ("gpu_rate_card_preemptible.csv", "USD per GPU-hour", *pivot_docs(docs, True, False)),
        ("gpu_rate_card_standalone.csv", "USD per GPU-hour", *pivot_docs(docs, False, True)),
        ("gpu_page_on_demand.csv", "USD per GPU-hour", *pivot_page(page, "on_demand")),
        ("gpu_page_preemptible.csv", "USD per GPU-hour", *pivot_page(page, "preemptible")),
        ("gpu_page_commitment.csv", "USD per GPU-hour", *pivot_page(page, "commitment")),
        ("gpu_page_reserve.csv", "USD per GPU-hour", *pivot_page(page, "reserve")),
        ("gpu_page_spot_floor.csv", "USD per GPU-hour", *pivot_page(page, "spot_floor")),
        ("gpu_page_scheduled_next.csv", "USD per GPU-hour", *pivot_page(page, "scheduled_next")),
        ("token_factory_input.csv", "USD per million input tokens", *pivot_tokens(tokens, "input_usd_per_million")),
        ("token_factory_output.csv", "USD per million output tokens", *pivot_tokens(tokens, "output_usd_per_million")),
    ]
    for filename, header, columns, grid in specs:
        if not columns:
            continue
        path = wide / filename
        _write_wide(path, header, columns, grid)
        written.append(path)
    return written
