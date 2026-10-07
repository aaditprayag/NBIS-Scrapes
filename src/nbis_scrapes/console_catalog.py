"""Nebius's own billing catalogue, from the console.

The console's billing pages read every SKU from one gateway call,
`customers/listFullView`. Each page of the answer has the SKUs in `items` and
their prices in `additionalItems.effectivePrices` and `futurePrices`. A SKU
carries its name, which ends in the region, and the date the SKU was made.
A price carries the date it took effect.

The call needs a logged-in console session. Nothing in this repo logs in.
The pages are saved from the browser, and this module reads the saved copy.

The date a region's first SKU was made is the closest public thing to a launch
date. Regions show up here before docs.nebius.com lists them.
"""

from __future__ import annotations

import csv
import json
import re
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "out"
CROSSWALK = ROOT / "data" / "region_crosswalk.json"
INFERENCES = ROOT / "data" / "console_region_inferences.json"

FIELDS = [
    "kind",
    "sku_id",
    "sku_name",
    "title",
    "region",
    "effective_from",
    "price_usd",
    "tiers",
    "currency_enum",
    "price_created",
    "sku_created",
    "pricing_unit",
    "sku_state",
]

REGION_RE = re.compile(r"/([a-z]+-[a-z]+\d)$")
GPU_RE = re.compile(r"^(?:compute/vm|bare-metal/server)\.(gpu-[a-z0-9-]+?|standard-v3-l40s)(?:\.preemptible)?(?:\.(?:gpu|cpu|ram))?/")


def _day(stamp) -> str:
    if not stamp or not stamp.get("seconds"):
        return ""
    return datetime.fromtimestamp(int(stamp["seconds"]), tz=timezone.utc).strftime("%Y-%m-%d")


def region_of(sku_name: str) -> str:
    match = REGION_RE.search(sku_name)
    return match.group(1) if match else ""


def gpu_platform(sku_name: str) -> str:
    match = GPU_RE.match(sku_name)
    return match.group(1) if match else ""


def parse_full_view(pages: list[dict]) -> list[dict]:
    """One row per price in saved listFullView pages."""
    rows = []
    for page in pages:
        extra = page.get("additionalItems") or {}
        skus = {item["metadata"]["id"]: item for item in page.get("items") or []}
        for kind in ("effectivePrices", "futurePrices"):
            for price in extra.get(kind) or []:
                sku = skus.get(price["metadata"]["parentId"]) or {}
                meta = sku.get("metadata") or {}
                labels = (meta.get("labels") or {}).get("labels") or {}
                tiers = ((price.get("spec") or {}).get("rate") or {}).get("tiers") or []
                name = meta.get("name") or ""
                rows.append({
                    "kind": kind,
                    "sku_id": price["metadata"]["parentId"],
                    "sku_name": name,
                    "title": (sku.get("spec") or {}).get("title") or "",
                    "region": region_of(name),
                    "effective_from": _day(price["spec"].get("effectiveTime")),
                    "price_usd": tiers[0]["priceValue"] if tiers else "",
                    "tiers": "|".join(f"{tier['startAmount']}:{tier['priceValue']}" for tier in tiers),
                    "currency_enum": price["spec"].get("currency", ""),
                    "price_created": _day(price["metadata"].get("createdAt")),
                    "sku_created": _day(meta.get("createdAt")),
                    "pricing_unit": labels.get("product__ui_calculator_pricing_unit", ""),
                    "sku_state": (sku.get("status") or {}).get("state", ""),
                })
    return rows


def read_export(path: Path) -> list[dict]:
    """Read a CSV saved from the console with the same columns."""
    with path.open(encoding="utf-8-sig") as handle:
        rows = []
        for record in csv.DictReader(handle):
            row = {field: record.get(field, "") for field in FIELDS}
            row["tiers"] = record.get("tiers") or record.get("tiers(start:price)") or ""
            rows.append(row)
    return rows


def region_table(rows: list[dict], crosswalk: list[dict], inferences: list[dict]) -> list[dict]:
    """First SKU date, GPU platforms, and the place each region is tied to."""
    docs = {entry["region"]: entry for entry in crosswalk}
    ours = {entry["region"]: entry for entry in inferences}
    by_region: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        if row["region"]:
            by_region[row["region"]].append(row)
    table = []
    for region in sorted(by_region, key=lambda name: min(row["sku_created"] for row in by_region[name])):
        skus = by_region[region]
        gpus = [row for row in skus if gpu_platform(row["sku_name"])]
        rubin = [row["sku_created"] for row in gpus if "vera-rubin" in row["sku_name"]]
        priced = [row for row in gpus if row["price_usd"] not in ("", "0")]
        if region in docs:
            place, basis = docs[region]["candidate"], "docs.nebius.com regions page, via data/region_crosswalk.json"
        elif region in ours:
            place, basis = ours[region]["candidate"], f"{ours[region]['confidence']}: {ours[region]['why']}"
        else:
            place, basis = "not tied to a place", "no public source"
        table.append({
            "region": region,
            "first_sku_created": min(row["sku_created"] for row in skus),
            "first_gpu_sku_created": min(row["sku_created"] for row in gpus) if gpus else "no GPU SKU",
            "skus": len(skus),
            "gpu_platforms": ", ".join(sorted({gpu_platform(row["sku_name"]) for row in gpus})) or "none",
            "gpu_skus_with_a_price": len(priced),
            "gpu_skus_at_zero": len(gpus) - len(priced),
            "first_vera_rubin_sku": min(rubin) if rubin else "none",
            "on_docs_regions_page": "yes" if region in docs else "no",
            "place": place,
            "place_basis": basis,
        })
    return table


def write_outputs(rows: list[dict], snapshot_day: str, out_dir: Path | None = None) -> tuple[Path, Path]:
    out_dir = out_dir or OUT
    out_dir.mkdir(parents=True, exist_ok=True)
    catalog = out_dir / f"console_catalog_{snapshot_day}.csv"
    _write(catalog, sorted(rows, key=lambda row: (row["region"], row["sku_name"], row["kind"])), FIELDS)
    crosswalk = json.loads(CROSSWALK.read_text(encoding="utf-8"))
    inferences = json.loads(INFERENCES.read_text(encoding="utf-8"))
    regions = out_dir / "console_regions.csv"
    table = region_table(rows, crosswalk, inferences)
    _write(regions, table, list(table[0].keys()))
    return catalog, regions


def _write(path: Path, rows: list[dict], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
