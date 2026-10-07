"""GPU availability: a daily capture plus whatever history public sources hold.

Price says what a GPU-hour costs. Availability says whether you can get one.
Four public feeds answer that without an account:

- RunPod's public GraphQL. For each GPU it returns a stock level (High,
  Medium, Low, or nothing) per data center and per community-cloud country.
- Vast.ai's offer search. Every listed machine shows up, rentable or not, so
  free GPUs over listed GPUs is a utilization read on the spot market.
- Shadeform's instance-type list. It is a broker over about twenty clouds,
  Lambda and Nebius included, and it flags each (instance type, region) as
  available or not.
- dstack's gpuhunt catalog. It is rebuilt every hour and only keeps RunPod
  offers that are in stock, so one copy a day is a RunPod stock history back
  to April 2024.

Vast's host metrics endpoint is added when VAST_API_KEY is set, as in demand.py.
Lambda's own API and Verda's need keys. They are recorded as skips.

Everything lands in one long CSV, one row per date, source, provider, GPU,
region and metric. A day nobody observed has no row.
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import statistics
import struct
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import zlib
from datetime import datetime, timedelta, timezone
from pathlib import Path

from nbis_scrapes.demand import USER_AGENT, VAST_METRICS_URL, summarize_vast_metrics
from nbis_scrapes.wide_prices import _clean, _write_wide

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "out"
WIDE = OUT / "wide"
SNAPSHOTS = ROOT / "data" / "snapshots"
FIXTURES = ROOT / "data" / "fixtures"
LONG = OUT / "gpu_availability.csv"
CACHE = ROOT / "data" / "cache"
CACHE_LONG = CACHE / "gpu_availability_full.csv"
GPUHUNT_VERSIONS = FIXTURES / "availability_gpuhunt_runpod_versions.txt"

FIELDS = ["date", "source", "provider", "gpu", "region", "metric", "value", "unit", "observed_ts", "source_url", "note"]
KEY = ("date", "source", "provider", "gpu", "region", "metric")

# The chips this file is about. Other GPUs in a feed are ignored.
TARGET_GPUS = ["B200", "H100", "H200", "B300", "GB200", "GB300", "H100 NVL", "H100 PCIe", "H200 NVL"]
WIDE_GPUS = ["B200", "H100", "H200", "B300", "GB200"]

SRC_CHART = "3Fourteen / Silicon Data chart (digitized)"
SRC_GPUHUNT = "dstack gpuhunt RunPod catalog"
SRC_RUNPOD = "RunPod GraphQL"
SRC_VAST = "Vast offer search"
SRC_VAST_METRICS = "Vast metrics API"
SRC_SHADEFORM = "Shadeform instance types"

SRC_FARM = "500.farm Vast marketplace"
SRC_FARM_PRICE = "500.farm Vast prices"
SRC_AKASH = "Akash console API"
SRC_IONET = "io.net explorer"

FARM_URL = "https://500.farm/vastai/grafana.v2/api/datasources/proxy/uid/EdgV2xcnz/api/v1/query_range"
# Vast names as 500.farm's exporter writes them. Consumer cards are left out.
FARM_GPUS = ["B200", "B300", "H100 SXM", "H100 PCIE", "H100 NVL", "H200", "H200 NVL", "A100 SXM4", "A100 PCIE"]
FARM_PRICE_NOTE = "verified, datacenter hosts, 8+ GPU offers, median, 1-day average, before Vast's fee"
AKASH_URL = "https://console-api.akash.network/v1/gpu-breakdown"
IONET_URL = "https://api.io.solutions/v1/io-explorer/network/market-snapshot"
OWN_SOURCES = {SRC_CHART, SRC_GPUHUNT, SRC_RUNPOD, SRC_VAST, SRC_VAST_METRICS, SRC_SHADEFORM,
               SRC_FARM, SRC_FARM_PRICE, SRC_AKASH, SRC_IONET}

RUNPOD_URL = "https://api.runpod.io/graphql"
VAST_BUNDLES_URL = "https://console.vast.ai/api/v0/bundles/"
SHADEFORM_URL = "https://api.shadeform.ai/v1/instances/types"
GPUHUNT_BUCKET = "https://dstack-gpu-pricing.s3.eu-west-1.amazonaws.com"
GPUHUNT_RUNPOD_LATEST = f"{GPUHUNT_BUCKET}/v3/runpod/latest/catalog.zip"
LAMBDA_URL = "https://cloud.lambda.ai/api/v1/instance-types"
VERDA_URL = "https://api.verda.com/v1/instance-availability"

VAST_GPU_NAMES = ["B200", "B300", "GB200", "H100 SXM", "H100 NVL", "H100 PCIE", "H200", "H200 NVL"]

RUNPOD_TYPES_QUERY = "query { countryCodes dataCenters { id listed } }"
RUNPOD_STOCK_QUERY = (
    "query Stock($input: GpuLowestPriceInput) { gpuTypes { id displayName "
    "lowestPrice(input: $input) { stockStatus uninterruptablePrice } } }"
)

# Chips already in gpuhunt's RunPod map, so a missing row is a real zero.
RUNPOD_CHIPS = ["B200", "H100", "H200", "H100 NVL", "H100 PCIe", "H200 NVL"]

SECURE_LOCATION = re.compile(r"^[A-Z]{2,3}-[A-Z]{2,4}-\d+$")
COUNTRY_LOCATION = re.compile(r"^[A-Z]{2}$")


def normalize_gpu(name: str) -> str:
    """One name per chip across RunPod ids, Vast names, Shadeform types, and gpuhunt."""
    text = (name or "").upper().replace("NVIDIA", " ").replace("_", " ").replace("-", " ")
    text = " ".join(text.split())
    for chip in ("GB300", "GB200", "B300", "B200"):
        if chip in text:
            return chip
    for chip in ("H200", "H100"):
        if chip in text:
            if "NVL" in text:
                return f"{chip} NVL"
            if "PCIE" in text:
                return f"{chip} PCIe"
            return chip
    return (name or "").strip()


def _row(date: str, source: str, provider: str, gpu: str, metric: str, value, unit: str = "",
         region: str = "", observed_ts: str = "", source_url: str = "", note: str = "") -> dict:
    if isinstance(value, float):
        value = _clean(value)
    return {
        "date": date, "source": source, "provider": provider, "gpu": gpu, "region": region,
        "metric": metric, "value": "" if value is None else str(value), "unit": unit,
        "observed_ts": observed_ts, "source_url": source_url, "note": note,
    }


# --- 3Fourteen / Silicon Data chart -------------------------------------------------

def parse_digitized_chart(text: str, source_url: str = "") -> list[dict]:
    """B200 on-demand availability, 30-day average, read off the 3Fourteen chart.

    Only points read off the line are kept. Points the digitizer interpolated
    between reads are dropped, so those days stay missing.
    """
    rows = []
    for record in csv.DictReader(io.StringIO(text)):
        if (record.get("basis") or "").strip().lower() != "chart":
            continue
        value = (record.get("b200_on_demand_availability_30d_avg_pct") or "").strip()
        if value == "":
            continue
        rows.append(_row(
            record["date"], SRC_CHART, "market", "B200", "on_demand_availability_30d_avg",
            float(value), "pct", source_url=source_url, note="read off the published chart",
        ))
    return rows


# --- dstack gpuhunt RunPod catalog -----------------------------------------------------

def parse_gpuhunt_runpod(text: str, date: str, source_url: str = "", observed_ts: str = "",
                         seen: set | None = None) -> list[dict]:
    """RunPod offers in a gpuhunt catalog are only the ones in stock when it was built.

    Per GPU: secure-cloud data centers in stock, community-cloud countries in
    stock, distinct (location, GPU count) offers, and the largest GPU count on
    offer. A GPU with no row counts as zero only after it has appeared in an
    earlier catalog (`seen`, updated in place), because gpuhunt added chips to
    its map at different dates. Community counts are written only on days the
    catalog carried community offers at all.
    """
    seen = seen if seen is not None else set()
    secure: dict[str, set] = {}
    country: dict[str, set] = {}
    offers: dict[str, set] = {}
    largest: dict[str, int] = {}
    any_secure = any_country = False
    for record in csv.DictReader(io.StringIO(text)):
        if str(record.get("spot", "")).lower() == "true":
            continue
        if "runpod-cluster" in (record.get("flags") or ""):
            continue
        location = (record.get("location") or "").strip()
        is_secure = bool(SECURE_LOCATION.match(location))
        is_country = bool(COUNTRY_LOCATION.match(location))
        any_secure = any_secure or is_secure
        any_country = any_country or is_country
        gpu = normalize_gpu(record.get("instance_name") or record.get("gpu_name") or "")
        if gpu not in TARGET_GPUS:
            continue
        try:
            count = int(float(record.get("gpu_count") or 0))
        except ValueError:
            continue
        if count < 1:
            continue
        seen.add(gpu)
        if is_secure:
            secure.setdefault(gpu, set()).add(location)
        if is_country:
            country.setdefault(gpu, set()).add(location)
        offers.setdefault(gpu, set()).add((location, count))
        largest[gpu] = max(largest.get(gpu, 0), count)
    rows = []
    if not (any_secure or any_country):
        return rows
    for gpu in [name for name in TARGET_GPUS if name in seen]:
        common = dict(observed_ts=observed_ts, source_url=source_url)
        if any_secure:
            rows.append(_row(date, SRC_GPUHUNT, "runpod", gpu, "secure_data_centers_in_stock",
                             len(secure.get(gpu, ())), "count", **common))
        if any_country:
            rows.append(_row(date, SRC_GPUHUNT, "runpod", gpu, "community_countries_in_stock",
                             len(country.get(gpu, ())), "count", **common))
        rows.append(_row(date, SRC_GPUHUNT, "runpod", gpu, "offers_in_stock",
                         len(offers.get(gpu, ())), "count", **common))
        if largest.get(gpu):
            rows.append(_row(date, SRC_GPUHUNT, "runpod", gpu, "largest_gpu_count_in_stock",
                             largest[gpu], "GPUs", **common))
    return rows


def read_versions(text: str) -> list[tuple[str, str]]:
    """Lines of `date url` from the gpuhunt version list fixture."""
    pairs = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        date, url = line.split()[:2]
        pairs.append((date, url))
    return pairs


# --- RunPod GraphQL ----------------------------------------------------------------

def parse_runpod_stock(payload: dict) -> dict[str, str | None]:
    """GPU -> stock level from one gpuTypes(lowestPrice) answer. None means out of stock."""
    levels: dict[str, str | None] = {}
    for item in ((payload or {}).get("data") or {}).get("gpuTypes") or []:
        gpu = normalize_gpu(item.get("id") or item.get("displayName") or "")
        if gpu not in TARGET_GPUS:
            continue
        status = (item.get("lowestPrice") or {}).get("stockStatus")
        # Several RunPod ids can map to one chip. Keep the best level seen.
        if gpu not in levels or _stock_rank(status) > _stock_rank(levels[gpu]):
            levels[gpu] = status
    return levels


def _stock_rank(status: str | None) -> int:
    return {"Low": 1, "Medium": 2, "High": 3}.get(status or "", 0)


def runpod_rows(date: str, overall: dict, by_data_center: dict, by_country: dict, observed_ts: str = "") -> list[dict]:
    rows = []
    gpus = sorted(set(overall) | {g for levels in by_data_center.values() for g in levels},
                  key=lambda name: TARGET_GPUS.index(name) if name in TARGET_GPUS else 99)
    common = dict(observed_ts=observed_ts, source_url=RUNPOD_URL)
    for gpu in gpus:
        if gpu in overall:
            rows.append(_row(date, SRC_RUNPOD, "runpod", gpu, "secure_stock_level",
                             overall[gpu] or "None", "level", **common))
        if by_data_center:
            in_stock = [dc for dc, levels in sorted(by_data_center.items()) if levels.get(gpu)]
            rows.append(_row(date, SRC_RUNPOD, "runpod", gpu, "secure_data_centers_in_stock",
                             len(in_stock), "count", **common))
            for dc in in_stock:
                rows.append(_row(date, SRC_RUNPOD, "runpod", gpu, "stock_level",
                                 by_data_center[dc][gpu], "level", region=dc, **common))
        if by_country:
            in_stock = [code for code, levels in sorted(by_country.items()) if levels.get(gpu)]
            rows.append(_row(date, SRC_RUNPOD, "runpod", gpu, "community_countries_in_stock",
                             len(in_stock), "count", **common))
    return rows


# --- Vast offer search -------------------------------------------------------------

def summarize_vast_offers(rentable: dict, unrentable: dict) -> dict[str, dict]:
    """Free GPUs over listed GPUs, per chip, from two offer searches.

    Vast lists each machine several times, as 1, 2, 4 and 8 GPU offers.
    A machine's size is num_gpus / gpu_frac. Its free GPUs are the largest
    rentable offer on it. Machines with no rentable offer are fully taken.
    The `rented` field on an offer is about the caller's own rentals, so it
    cannot split taken GPUs into rented and held back. The median price is
    dph_total / num_gpus across rentable offers, each offer counted once.
    """
    machines: dict[tuple, dict] = {}
    prices: dict[str, dict] = {}
    for payload, is_rentable in ((unrentable, False), (rentable, True)):
        for offer in (payload or {}).get("offers") or []:
            gpu = normalize_gpu(offer.get("gpu_name") or "")
            if gpu not in TARGET_GPUS or offer.get("machine_id") is None:
                continue
            count = int(offer.get("num_gpus") or 0)
            fraction = offer.get("gpu_frac") or 1
            size = round(count / fraction) if fraction else count
            entry = machines.setdefault((gpu, offer["machine_id"]), {"size": 0, "free": 0})
            entry["size"] = max(entry["size"], size, count)
            if is_rentable or offer.get("rentable"):
                entry["free"] = max(entry["free"], count)
                if count and offer.get("dph_total") is not None:
                    key = offer.get("id") or (offer["machine_id"], count)
                    prices.setdefault(gpu, {})[key] = float(offer["dph_total"]) / count
    summary: dict[str, dict] = {}
    for (gpu, _machine), entry in machines.items():
        totals = summary.setdefault(gpu, {"machines_listed": 0, "gpus_listed": 0, "gpus_rentable": 0})
        totals["machines_listed"] += 1
        totals["gpus_listed"] += entry["size"]
        totals["gpus_rentable"] += min(entry["free"], entry["size"])
    for gpu, totals in summary.items():
        listed = totals["gpus_listed"]
        totals["gpus_not_rentable"] = listed - totals["gpus_rentable"]
        totals["rentable_pct"] = round(100 * totals["gpus_rentable"] / listed, 2) if listed else None
        per_gpu = sorted((prices.get(gpu) or {}).values())
        totals["offers_rentable"] = len(per_gpu)
        totals["median_usd_per_gpu_hour"] = round(statistics.median(per_gpu), 4) if per_gpu else None
    return summary


def vast_rows(date: str, summary: dict, observed_ts: str = "", note: str = "") -> list[dict]:
    rows = []
    units = {
        "offers_rentable": "count", "machines_listed": "count", "gpus_listed": "GPUs",
        "gpus_rentable": "GPUs", "gpus_not_rentable": "GPUs", "rentable_pct": "pct",
        "median_usd_per_gpu_hour": "USD per GPU-hour",
    }
    for gpu in [name for name in TARGET_GPUS if name in summary]:
        for metric, unit in units.items():
            value = summary[gpu].get(metric)
            if value is None:
                continue
            rows.append(_row(date, SRC_VAST, "vast", gpu, metric, value, unit,
                             observed_ts=observed_ts, source_url=VAST_BUNDLES_URL, note=note))
    return rows


def vast_keys() -> list[tuple[str, str]]:
    """Vast keys on this machine, best first: the vast CLI session key file, then VAST_API_KEY.

    Returns (key, where it came from). The key itself is never written anywhere.
    """
    keys = []
    path = Path(os.path.expanduser("~/.config/vastai/vast_tfa_key"))
    try:
        session = path.read_text(encoding="utf-8").strip()
    except OSError:
        session = ""
    if session:
        keys.append((session, "vast CLI session key"))
    env = os.environ.get("VAST_API_KEY", "").strip()
    if env and env != session:
        keys.append((env, "VAST_API_KEY"))
    return keys


def vast_metric_rows(date: str, parsed: dict, observed_ts: str = "") -> list[dict]:
    rows = []
    for item in parsed.get("gpus") or []:
        gpu = normalize_gpu(item.get("gpu") or "")
        if gpu not in TARGET_GPUS:
            continue
        for metric in ("supply", "demand", "price_median"):
            if item.get(metric) is None:
                continue
            unit = "USD per GPU-hour" if metric.startswith("price") else "count"
            rows.append(_row(date, SRC_VAST_METRICS, "vast", gpu, metric, item[metric], unit,
                             observed_ts=observed_ts, source_url=VAST_METRICS_URL))
    return rows


# --- Shadeform -----------------------------------------------------------------------

def parse_shadeform_types(payload: dict) -> list[dict]:
    """One record per cloud, chip, instance type and region, on-demand only."""
    records = []
    for item in (payload or {}).get("instance_types") or []:
        config = item.get("configuration") or {}
        gpu = normalize_gpu(config.get("gpu_type") or item.get("gpu_type") or "")
        if gpu not in TARGET_GPUS:
            continue
        for slot in item.get("availability") or []:
            if (slot.get("rental_type") or "on_demand") != "on_demand":
                continue
            records.append({
                "cloud": item.get("cloud") or "",
                "gpu": gpu,
                "instance_type": item.get("shade_instance_type") or item.get("cloud_instance_type") or "",
                "num_gpus": item.get("num_gpus") or config.get("num_gpus"),
                "region": slot.get("region") or "",
                "available": bool(slot.get("available")),
            })
    return records


def shadeform_rows(date: str, records: list[dict], observed_ts: str = "") -> list[dict]:
    rows = []
    common = dict(observed_ts=observed_ts, source_url=SHADEFORM_URL)
    groups: dict[tuple, list[dict]] = {}
    for record in records:
        groups.setdefault((record["gpu"], "all"), []).append(record)
        groups.setdefault((record["gpu"], record["cloud"]), []).append(record)
    order = lambda key: (TARGET_GPUS.index(key[0]), key[1] != "all", key[1])
    for gpu, cloud in sorted(groups, key=order):
        group = groups[(gpu, cloud)]
        free = sum(1 for record in group if record["available"])
        rows.append(_row(date, SRC_SHADEFORM, cloud, gpu, "listings", len(group), "count", **common))
        rows.append(_row(date, SRC_SHADEFORM, cloud, gpu, "listings_available", free, "count", **common))
        if cloud == "all":
            rows.append(_row(date, SRC_SHADEFORM, cloud, gpu, "listings_available_pct",
                             round(100 * free / len(group), 2), "pct", **common))
    flags: dict[tuple, bool] = {}
    for record in records:
        key = (record["gpu"], record["cloud"], record["region"])
        flags[key] = flags.get(key, False) or record["available"]
    for (gpu, cloud, region), available in sorted(flags.items(), key=lambda item: (TARGET_GPUS.index(item[0][0]), item[0][1:])):
        rows.append(_row(date, SRC_SHADEFORM, cloud, gpu, "available", int(available), "flag", region=region, **common))
    return rows


# --- 500.farm (Vast marketplace history) ---------------------------------------------

def farm_gpu(name: str) -> str:
    """500.farm's Vast names, in this file's naming. A100s keep their form factor."""
    gpu = normalize_gpu(name)
    return gpu if gpu in TARGET_GPUS else " ".join((name or "").split())


def _farm_series(payload: dict) -> dict[str, dict[str, float]]:
    """gpu -> date -> value from a Prometheus query_range answer."""
    series: dict[str, dict[str, float]] = {}
    for item in (((payload or {}).get("data") or {}).get("result")) or []:
        name = (item.get("metric") or {}).get("gpu_name")
        if not name:
            continue
        for stamp, value in item.get("values") or []:
            day = datetime.fromtimestamp(float(stamp), timezone.utc).strftime("%Y-%m-%d")
            series.setdefault(farm_gpu(name), {})[day] = float(value)
    return series


def parse_500farm_counts(rented: dict, available: dict) -> list[dict]:
    """GPUs rented and not rented on Vast at 00:00 UTC, per chip. A day needs both halves."""
    taken, free = _farm_series(rented), _farm_series(available)
    rows = []
    for gpu in sorted(set(taken) & set(free)):
        for day in sorted(set(taken[gpu]) & set(free[gpu])):
            listed = taken[gpu][day] + free[gpu][day]
            common = dict(source_url=FARM_URL)
            rows.append(_row(day, SRC_FARM, "vast", gpu, "gpus_rented", taken[gpu][day], "GPUs", **common))
            rows.append(_row(day, SRC_FARM, "vast", gpu, "gpus_available", free[gpu][day], "GPUs", **common))
            rows.append(_row(day, SRC_FARM, "vast", gpu, "gpus_listed", listed, "GPUs", **common))
            if listed:
                rows.append(_row(day, SRC_FARM, "vast", gpu, "pct_available",
                                 round(100 * free[gpu][day] / listed, 2), "pct", **common))
    return rows


def parse_500farm_prices(rented: dict, available: dict) -> list[dict]:
    """Median Vast price per GPU-hour for rented and for unrented offers, the headline cut."""
    rows = []
    for payload, metric in ((rented, "median_usd_per_gpu_hour_rented"), (available, "median_usd_per_gpu_hour_available")):
        for gpu, days in sorted(_farm_series(payload).items()):
            for day, value in sorted(days.items()):
                rows.append(_row(day, SRC_FARM_PRICE, "vast", gpu, metric, round(value, 4), "USD per GPU-hour",
                                 source_url=FARM_URL, note=FARM_PRICE_NOTE))
    return rows


def farm_rows_from_csv(counts_text: str = "", prices_text: str = "") -> list[dict]:
    """Rows from the one-off 500.farm pulls: counts by GPU, and the full price tape (headline cut only)."""
    rows = []
    for record in csv.DictReader(io.StringIO(counts_text)) if counts_text else []:
        if record.get("gpu") not in FARM_GPUS:
            continue
        gpu = farm_gpu(record["gpu"])
        note = record.get("metric", "")
        for field, unit in (("gpus_rented", "GPUs"), ("gpus_available", "GPUs"), ("gpus_listed", "GPUs"), ("pct_available", "pct")):
            if record.get(field) not in (None, ""):
                rows.append(_row(record["date"], SRC_FARM, "vast", gpu, field, float(record[field]), unit,
                                 source_url=FARM_URL, note=note))
    for record in csv.DictReader(io.StringIO(prices_text)) if prices_text else []:
        if (record.get("gpu") not in FARM_GPUS or record.get("datacenter") != "yes"
                or record.get("gpu_count_range") != "8+" or record.get("stat") != "median"):
            continue
        state = "rented" if record.get("offer_state") == "rented" else "available"
        rows.append(_row(record["date"], SRC_FARM_PRICE, "vast", farm_gpu(record["gpu"]),
                         f"median_usd_per_gpu_hour_{state}", round(float(record["usd_per_gpu_hour"]), 4),
                         "USD per GPU-hour", source_url=FARM_URL, note=f"{FARM_PRICE_NOTE}; {record.get('metric', '')}"))
    return rows


def farm_queries(gpus: list[str] | None = None) -> dict[str, str]:
    names = "|".join(gpus or FARM_GPUS)
    price = ('avg_over_time(vastai_v2_ondemand_price_median_dollars{{gpu_name=~"{n}",verified="yes",'
             'rented="{r}",datacenter="yes",gpu_count_range="8+"}}[1d])')
    return {
        "count_rented": f'sum by (gpu_name) (vastai_gpu_count{{rented="yes",gpu_name=~"{names}"}})',
        "count_available": f'sum by (gpu_name) (vastai_gpu_count{{rented="no",gpu_name=~"{names}"}})',
        "price_rented": price.format(n=names, r="yes"),
        "price_available": price.format(n=names, r="no"),
    }


def fetch_500farm(days: int = 2, pause: float = 1.2, timeout: int = 60) -> tuple[list[dict], dict]:
    """The last `days` days at 00:00 UTC. Four queries, about one a second."""
    end = int(datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0).timestamp())
    start = end - 86400 * (days - 1)
    answers: dict = {}
    errors = []
    for name, query in farm_queries().items():
        body = urllib.parse.urlencode({"query": query, "start": start, "end": end, "step": 86400}).encode()
        status, raw, _headers = _request(FARM_URL, data=body, timeout=timeout,
                                         headers={"Content-Type": "application/x-www-form-urlencoded"})
        time.sleep(pause)
        try:
            answers[name] = json.loads(raw) if status == 200 else None
        except json.JSONDecodeError:
            answers[name] = None
        if answers[name] is None:
            errors.append(f"{name}: {status} {raw[:200].decode('utf-8', errors='replace')}")
    rows = []
    if answers.get("count_rented") and answers.get("count_available"):
        rows += parse_500farm_counts(answers["count_rented"], answers["count_available"])
    rows += parse_500farm_prices(answers.get("price_rented") or {}, answers.get("price_available") or {})
    return rows, {"url": FARM_URL, "http_status": 200 if not errors else None,
                  "parsed": {"rows": len(rows), "window": [start, end]}, "error": "; ".join(errors) or None}


# --- Akash -------------------------------------------------------------------------------

def parse_akash_breakdown(records: list) -> list[dict]:
    """Akash's daily GPU breakdown: GPUs on the network, GPUs leased, and Akash's own utilization."""
    rows = []
    for record in records or []:
        model = (record.get("model") or "").upper()
        gpu = "A100" if model == "A100" else normalize_gpu(model)
        if gpu not in TARGET_GPUS and gpu != "A100":
            continue
        day = (record.get("date") or "")[:10]
        common = dict(source_url=AKASH_URL, note=f"model {record.get('model')}, all interfaces and memory sizes")
        for field, metric, unit in (("totalGpus", "gpus_total", "GPUs"), ("leasedGpus", "gpus_leased", "GPUs"),
                                    ("gpuUtilization", "utilization_pct", "pct"), ("providerCount", "providers", "count")):
            if record.get(field) is not None:
                rows.append(_row(day, SRC_AKASH, "akash", gpu, metric, float(record[field]), unit, **common))
    return rows


def fetch_akash_history(start: str = "2024-01-01", pause: float = 1.5) -> list[dict]:
    """Akash allows 366 days a call, so a year at a time."""
    rows = []
    year = int(start[:4])
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    while f"{year}-01-01" <= today:
        first = max(start, f"{year}-01-01")
        last = min(today, f"{year}-12-31")
        part, _info = fetch_akash(first, last)
        rows += part
        year += 1
        time.sleep(pause)
    return rows


def fetch_akash(start: str, end: str, timeout: int = 120) -> tuple[list[dict], dict]:
    url = f"{AKASH_URL}?startDate={start}&endDate={end}"
    status, payload, problem = _json(url, timeout=timeout)
    rows = parse_akash_breakdown(payload) if status == 200 and isinstance(payload, list) else []
    return rows, {"url": url, "http_status": status, "parsed": {"rows": len(rows)}, "error": problem}


# --- io.net ------------------------------------------------------------------------------

def parse_ionet_market(payload: dict) -> dict[str, dict]:
    """io.net's market snapshot: GPUs on the network, hired, and idle, per chip, plus its list price."""
    groups: dict[str, list[dict]] = {}
    for item in (payload or {}).get("data") or []:
        if (item.get("type") or "GPU") != "GPU":
            continue
        gpu = normalize_gpu(item.get("hardware_name") or "")
        if gpu in TARGET_GPUS:
            groups.setdefault(gpu, []).append(item)
    summary = {}
    for gpu, items in groups.items():
        total = sum(int(item.get("total") or 0) for item in items)
        hired = sum(int(item.get("hired") or 0) for item in items)
        idle = sum(int(item.get("idle") or 0) for item in items)
        biggest = max(items, key=lambda item: int(item.get("total") or 0))
        summary[gpu] = {"gpus_total": total, "gpus_hired": hired, "gpus_idle": idle,
                        "hired_pct": round(100 * hired / total, 2) if total else None,
                        "price_usd_per_gpu_hour": biggest.get("price"),
                        "names": [item.get("hardware_name") for item in items]}
    return summary


def ionet_rows(date: str, summary: dict, observed_ts: str = "") -> list[dict]:
    rows = []
    units = {"gpus_total": "GPUs", "gpus_hired": "GPUs", "gpus_idle": "GPUs", "hired_pct": "pct",
             "price_usd_per_gpu_hour": "USD per GPU-hour"}
    for gpu in [name for name in TARGET_GPUS if name in summary]:
        for metric, unit in units.items():
            value = summary[gpu].get(metric)
            if value is None:
                continue
            rows.append(_row(date, SRC_IONET, "ionet", gpu, metric, float(value), unit, observed_ts=observed_ts,
                             source_url=IONET_URL, note="; ".join(summary[gpu]["names"])))
    return rows


# --- network -------------------------------------------------------------------------

def _request(url: str, data: bytes | None = None, headers: dict | None = None,
             method: str | None = None, timeout: int = 60) -> tuple[int, bytes, dict]:
    request = urllib.request.Request(url, data=data, method=method,
                                     headers={"User-Agent": USER_AGENT, **(headers or {})})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status, response.read(), dict(response.headers)
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read(), dict(exc.headers or {})
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        return -1, str(exc).encode(), {}


def _json(url: str, body: dict | None = None, headers: dict | None = None, timeout: int = 60) -> tuple[int, object, str | None]:
    data = json.dumps(body).encode() if body is not None else None
    extra = {"Content-Type": "application/json", **(headers or {})} if body is not None else headers
    status, raw, _headers = _request(url, data=data, headers=extra, timeout=timeout)
    if status != 200:
        return status, None, raw[:500].decode("utf-8", errors="replace")
    try:
        return status, json.loads(raw), None
    except json.JSONDecodeError as exc:
        return status, None, f"json: {exc}"


def read_zip_member(url: str, member: str, pause: float = 0.4) -> tuple[bytes | None, str]:
    """Read one file out of a remote zip with HTTP range requests. Returns (bytes, Last-Modified)."""
    status, tail, headers = _request(url, headers={"Range": "bytes=-65536"})
    if status not in (200, 206):
        return None, ""
    modified = headers.get("Last-Modified", "")
    total = int(headers.get("Content-Range", f"/{len(tail)}").split("/")[-1])
    if status == 200 or total <= len(tail):
        with zipfile.ZipFile(io.BytesIO(tail)) as archive:
            return (archive.read(member) if member in archive.namelist() else None), modified
    base = total - len(tail)
    end_record = tail.rfind(b"PK\x05\x06")
    if end_record < 0:
        return None, modified
    cd_size, cd_offset = struct.unpack("<II", tail[end_record + 12:end_record + 20])
    if cd_offset < base:
        time.sleep(pause)
        _status, directory, _h = _request(url, headers={"Range": f"bytes={cd_offset}-{cd_offset + cd_size - 1}"})
    else:
        directory = tail[cd_offset - base:cd_offset - base + cd_size]
    position = 0
    while position + 46 <= len(directory) and directory[position:position + 4] == b"PK\x01\x02":
        method = struct.unpack("<H", directory[position + 10:position + 12])[0]
        size = struct.unpack("<I", directory[position + 20:position + 24])[0]
        name_len, extra_len, comment_len = struct.unpack("<HHH", directory[position + 28:position + 34])
        local = struct.unpack("<I", directory[position + 42:position + 46])[0]
        name = directory[position + 46:position + 46 + name_len].decode()
        position += 46 + name_len + extra_len + comment_len
        if name != member:
            continue
        stop = min(local + 30 + 1024 + size, total)
        if local >= base:
            blob = tail[local - base:]
        else:
            time.sleep(pause)
            _status, blob, _h = _request(url, headers={"Range": f"bytes={local}-{stop - 1}"})
        skip = 30 + sum(struct.unpack("<HH", blob[26:30]))
        data = blob[skip:skip + size]
        return (zlib.decompressobj(-15).decompress(data) if method == 8 else data), modified
    return None, modified


def _ts(last_modified: str) -> str:
    try:
        return datetime.strptime(last_modified, "%a, %d %b %Y %H:%M:%S GMT").strftime("%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        return ""


def gpuhunt_history(versions: list[tuple[str, str]], pause: float = 0.4, log=print) -> list[dict]:
    """Fetch runpod.csv from one gpuhunt catalog per day. Sequential and paced."""
    rows: list[dict] = []
    seen: set = set()
    for date, url in sorted(versions):
        data, modified = read_zip_member(url, "runpod.csv", pause=pause)
        time.sleep(pause)
        if data is None:
            log(f"{date} no runpod.csv")
            continue
        rows.extend(parse_gpuhunt_runpod(data.decode("utf-8", errors="replace"), date, url, _ts(modified), seen))
    return rows


def _vast_search(key: str, pause: float, timeout: int) -> tuple[dict, str | None, int]:
    """Rentable and unrentable offers for the tracked chips. Returns (searches, error, last status)."""
    searches: dict = {}
    code = 0
    for label, rentable in (("rentable", True), ("unrentable", False)):
        query = {"gpu_name": {"in": VAST_GPU_NAMES}, "rentable": {"eq": rentable}, "limit": 5000}
        url = f"{VAST_BUNDLES_URL}?q={urllib.parse.quote(json.dumps(query))}"
        auth = {"Authorization": f"Bearer {key}"} if key else None
        code, payload, problem = _json(url, headers=auth, timeout=timeout)
        if code != 200 or not isinstance(payload, dict):
            return searches, f"{label}: {code} {problem}", code
        searches[label] = payload
        time.sleep(pause)
    return searches, None, code


def collect_live(pause: float = 1.0, timeout: int = 60) -> dict:
    """Pull today's availability from every public feed that answers without a login."""
    taken_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    today = taken_at[:10]
    snapshot = {"taken_at": taken_at, "sources": {}, "rows": []}
    sources = snapshot["sources"]
    rows = snapshot["rows"]

    # RunPod: one call for the overall level, then one per data center and per country.
    status, meta, error = _json(RUNPOD_URL, {"query": RUNPOD_TYPES_QUERY}, timeout=timeout)
    overall: dict = {}
    by_dc: dict = {}
    by_country: dict = {}
    errors = [] if status == 200 else [error]
    if status == 200:
        data = (meta or {}).get("data") or {}
        centers = [item["id"] for item in data.get("dataCenters") or [] if item.get("listed")]
        countries = [code for code in data.get("countryCodes") or [] if code]
        jobs = [("overall", None, {"gpuCount": 1, "secureCloud": True})]
        jobs += [("dc", dc, {"gpuCount": 1, "secureCloud": True, "dataCenterId": dc}) for dc in centers]
        jobs += [("country", code, {"gpuCount": 1, "secureCloud": False, "countryCode": code}) for code in countries]
        for kind, key, variables in jobs:
            time.sleep(pause)
            code, payload, problem = _json(RUNPOD_URL, {"query": RUNPOD_STOCK_QUERY, "variables": {"input": variables}}, timeout=timeout)
            if code != 200 or not isinstance(payload, dict) or payload.get("errors"):
                errors.append(f"{kind} {key}: {problem or (payload or {}).get('errors')}"[:300])
                continue
            levels = parse_runpod_stock(payload)
            if kind == "overall":
                overall = levels
            elif kind == "dc":
                by_dc[key] = levels
            else:
                by_country[key] = levels
        rows.extend(runpod_rows(today, overall, by_dc, by_country, taken_at))
    sources["runpod"] = {
        "url": RUNPOD_URL, "http_status": status,
        "parsed": {"overall": overall, "data_centers": by_dc, "countries": by_country},
        "error": "; ".join(e for e in errors if e) or None,
    }

    # Vast: rentable offers and the rest, for the chips we track. Signed in when a key
    # is on this machine, because a signed-in search returns more of the market.
    time.sleep(pause)
    key, key_origin = "", ""
    searches, vast_error, code = {}, None, 0
    for candidate, origin in vast_keys():
        searches, vast_error, code = _vast_search(candidate, pause, timeout)
        if code not in (401, 403):
            key, key_origin = candidate, origin
            break
    signed_in = bool(key)
    if not signed_in:
        searches, vast_error, code = _vast_search("", pause, timeout)
    for label, payload in searches.items():
        if len(payload.get("offers") or []) >= 5000:
            vast_error = f"{label}: hit the 5000 offer limit, counts would be short"
    how = f"signed in ({key_origin})" if signed_in else "no key"
    summary = summarize_vast_offers(searches.get("rentable"), searches.get("unrentable")) if not vast_error else None
    if summary:
        rows.extend(vast_rows(today, summary, taken_at, note=f"offer search, {how}"))
    sources["vast_offers"] = {
        "url": VAST_BUNDLES_URL, "http_status": 200 if not vast_error else None, "auth": how,
        "parsed": {"gpus": summary, "offers": {k: len(v.get("offers") or []) for k, v in searches.items()}},
        "error": vast_error,
    }

    # Shadeform: every cloud it brokers, flagged available or not per region.
    code, payload, problem = _json(SHADEFORM_URL, timeout=timeout)
    records = parse_shadeform_types(payload) if code == 200 else []
    rows.extend(shadeform_rows(today, records, taken_at))
    sources["shadeform"] = {"url": SHADEFORM_URL, "http_status": code, "parsed": records, "error": problem}

    # gpuhunt: the newest RunPod catalog, which carries the history forward.
    time.sleep(pause)
    code, body, headers = _request(GPUHUNT_RUNPOD_LATEST, timeout=timeout)
    gpuhunt_error = None
    built = ""
    if code == 200:
        try:
            with zipfile.ZipFile(io.BytesIO(body)) as archive:
                text = archive.read("runpod.csv").decode("utf-8", errors="replace")
            built = _ts(headers.get("Last-Modified", ""))
            rows.extend(parse_gpuhunt_runpod(text, built[:10] or today, GPUHUNT_RUNPOD_LATEST, built, set(RUNPOD_CHIPS)))
        except (zipfile.BadZipFile, KeyError) as exc:
            gpuhunt_error = f"zip: {exc}"
    else:
        gpuhunt_error = body[:300].decode("utf-8", errors="replace")
    sources["gpuhunt_runpod"] = {"url": GPUHUNT_RUNPOD_LATEST, "http_status": code, "parsed": {"built": built}, "error": gpuhunt_error}

    # 500.farm: Vast's marketplace at 00:00 UTC, the last two days.
    time.sleep(pause)
    farm, sources["500farm"] = fetch_500farm(days=2)
    rows.extend(farm)

    # Akash: the daily GPU breakdown for the last three days. Today's row is still filling.
    start = (datetime.now(timezone.utc) - timedelta(days=3)).strftime("%Y-%m-%d")
    akash, sources["akash"] = fetch_akash(start, today, timeout=timeout)
    rows.extend(akash)

    # Vast market metrics. They need a key, and Vast only serves them to a host with
    # an active registered machine. Otherwise it answers needs_machine.
    if not key and vast_keys():
        key, key_origin = vast_keys()[0]
    if not key:
        sources["vast_metrics"] = {"url": VAST_METRICS_URL, "http_status": None, "parsed": None,
                                   "error": "skipped: no Vast key (vast CLI session key file or VAST_API_KEY)."}
    else:
        time.sleep(pause)
        code, payload, problem = _json(VAST_METRICS_URL, headers={"Authorization": f"Bearer {key}"}, timeout=timeout)
        parsed = None
        if code == 200 and isinstance(payload, dict) and payload.get("needs_machine"):
            problem = "skipped: Vast answered needs_machine. Market metrics need an active registered host machine."
        elif code == 200 and payload is not None:
            parsed = summarize_vast_metrics(payload)
            rows.extend(vast_metric_rows(today, parsed, taken_at))
        sources["vast_metrics"] = {"url": VAST_METRICS_URL, "http_status": code, "auth": key_origin,
                                   "parsed": parsed, "error": problem}

    for name, url, variable in (("lambda", LAMBDA_URL, "LAMBDA_API_KEY"), ("verda", VERDA_URL, "VERDA_CLIENT_ID")):
        sources[name] = {"url": url, "http_status": None, "parsed": None,
                         "error": f"skipped: needs an account key ({variable}). Lambda and Verda capacity is read through Shadeform."}
    return snapshot


# --- files -----------------------------------------------------------------------------

def read_long(path: Path | None = None) -> list[dict]:
    path = path or LONG
    if not path.exists():
        return []
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def merge_rows(existing: list[dict], new: list[dict]) -> list[dict]:
    """One row per date, source, provider, GPU, region and metric. A new read replaces an old one."""
    table = {tuple(row.get(field, "") for field in KEY): row for row in existing}
    for row in new:
        table[tuple(row.get(field, "") for field in KEY)] = {field: row.get(field, "") for field in FIELDS}
    return sorted(table.values(), key=lambda row: tuple(row.get(field, "") for field in KEY))


def write_long(rows: list[dict], path: Path | None = None) -> Path:
    path = path or LONG
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in FIELDS})
    return path


def append_rows(new: list[dict], path: Path | None = None, cache: Path | None = None) -> tuple[Path, int]:
    """Add rows to the full file in data/cache and the compact one in data/out, without duplicates.

    The cache keeps every row as read. The committed file keeps the series this
    repo reports on, with the per-row URL and raw text left out where the source
    has one fixed address (see `compact_rows`). Without a cache, only the
    committed file is written.
    """
    cache = cache if cache is not None else (CACHE_LONG if path is None else None)
    if cache is not None:
        write_long(merge_rows(read_long(cache), new), cache)
    existing = read_long(path)
    merged = merge_rows(existing, compact_rows(new))
    return write_long(merged, path), len(merged) - len(existing)


# Rows that stay in data/cache only: the parts of a source that the committed
# file can rebuild (rented + available = listed) or that are too fine-grained.
CACHE_ONLY = {
    (SRC_FARM, "gpus_rented"), (SRC_FARM, "gpus_available"),
    (SRC_FARM_PRICE, "median_usd_per_gpu_hour_available"),
    (SRC_AKASH, "gpus_leased"), (SRC_AKASH, "providers"),
}
COMMITTED_FROM = "2024-01-01"
# Research imports kept in the committed file: the ones with 30 days or more on the tracked chips.
COMMITTED_IMPORTS = {"Shadeform instance-types API (Wayback)"}
# Sources whose address is fixed and written once in docs/gpu-availability.md.
FIXED_URL = {SRC_CHART, SRC_GPUHUNT, SRC_FARM, SRC_FARM_PRICE, SRC_AKASH}


def compact_rows(rows: list[dict]) -> list[dict]:
    """The committed subset of rows. Imported research rows lose their per-region detail and raw text."""
    kept = []
    for row in rows:
        if (row.get("source"), row.get("metric")) in CACHE_ONLY:
            continue
        # The committed file is about five chips. Variants (H100 NVL, H200 NVL, A100s) stay in the cache.
        if row.get("gpu") not in WIDE_GPUS:
            continue
        # 500.farm reaches back to 2021. This repo's window starts in 2024; the rest is in the cache.
        if row.get("source") in (SRC_FARM, SRC_FARM_PRICE) and row.get("date", "") < COMMITTED_FROM:
            continue
        imported = row.get("source", "") not in OWN_SOURCES
        if imported and (row.get("region") or row.get("source") not in COMMITTED_IMPORTS):
            continue
        row = dict(row)
        if row.get("source") in FIXED_URL or imported:
            row["source_url"] = ""
        if row.get("source") in (SRC_FARM, SRC_FARM_PRICE, SRC_AKASH):
            row["note"] = ""
        if imported:
            row["note"] = ""
            row["observed_ts"] = ""
        kept.append(row)
    return kept


# --- 500.farm full history -----------------------------------------------------------------

FARM_PRICE_STATS = {"10th_percentile": "p10", "median": "median", "90th_percentile": "p90"}
FARM_PRICE_FIELDS = ["date", "gpu", "datacenter", "gpu_count_range", "offer_state", "p10", "median", "p90", "metric"]


def _farm_post(query: str, start: int, end: int, timeout: int = 120) -> dict | None:
    body = urllib.parse.urlencode({"query": query, "start": start, "end": end, "step": 86400}).encode()
    status, raw, _headers = _request(FARM_URL, data=body, timeout=timeout,
                                     headers={"Content-Type": "application/x-www-form-urlencoded"})
    if status != 200:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def parse_500farm_price_tape(answers: dict) -> list[dict]:
    """One row per date, GPU, segment and offer state, with p10, median and p90 side by side.

    `answers` maps (metric family, stat, rented) to a query_range answer. The
    v1 family has no datacenter or size labels, so its segment is all/all.
    """
    table: dict[tuple, dict] = {}
    for (family, stat, rented), payload in answers.items():
        for item in (((payload or {}).get("data") or {}).get("result")) or []:
            labels = item.get("metric") or {}
            gpu = labels.get("gpu_name")
            if not gpu:
                continue
            segment = (labels.get("datacenter", "all"), labels.get("gpu_count_range", "all"))
            for stamp, value in item.get("values") or []:
                day = datetime.fromtimestamp(float(stamp), timezone.utc).strftime("%Y-%m-%d")
                key = (day, gpu, *segment, "rented" if rented == "yes" else "available", family)
                table.setdefault(key, {})[FARM_PRICE_STATS[stat]] = _clean(round(float(value), 4))
    rows = []
    for key in sorted(table):
        day, gpu, datacenter, size, state, family = key
        rows.append({"date": day, "gpu": gpu, "datacenter": datacenter, "gpu_count_range": size,
                     "offer_state": state, **table[key], "metric": family})
    return rows


def headline_price_tape(rows: list[dict]) -> list[dict]:
    """The committed cut: datacenter chips; datacenter yes/no at 8+ GPUs, and the all/all v1 series."""
    return [row for row in rows if row["gpu"] in FARM_GPUS and (
        (row["datacenter"] in ("yes", "no") and row["gpu_count_range"] == "8+")
        or (row["datacenter"] == "all" and row["gpu_count_range"] == "all"))]


def write_price_tape(rows: list[dict], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FARM_PRICE_FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in FARM_PRICE_FIELDS})
    return path


def _only(payload: dict | None, **labels) -> dict:
    """Keep the series of a query_range answer whose labels match."""
    results = (((payload or {}).get("data") or {}).get("result")) or []
    kept = [item for item in results if all((item.get("metric") or {}).get(k) == v for k, v in labels.items())]
    return {"data": {"result": kept}}


def fetch_500farm_history(start: str = "2021-10-01", pause: float = 1.2, log=print) -> list[dict]:
    """The whole 500.farm record, one point a day: counts and the price tape.

    Fourteen queries, sequential, about one a second. The full price tape goes to
    data/cache/gpu_prices_500farm_full.csv and the headline cut to
    data/out/gpu_prices_500farm.csv. Counts and the headline median come back as rows.
    """
    first = int(datetime.strptime(start, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp())
    end = int(datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0).timestamp())
    names = "|".join(FARM_GPUS)
    counts = {}
    for rented in ("yes", "no"):
        counts[rented] = _farm_post(f'sum by (gpu_name) (vastai_gpu_count{{rented="{rented}",gpu_name=~"{names}"}})', first, end)
        log(f"counts rented={rented} {'ok' if counts[rented] else 'failed'}")
        time.sleep(pause)
    answers = {}
    everything = "|".join(FARM_GPUS + ["L40S", "RTX 4090", "RTX 5090", "A100X", "A100 SXM"])
    for family, prefix in (("vastai_v2", "vastai_v2_ondemand_price"), ("vastai", "vastai_ondemand_price")):
        for stat in FARM_PRICE_STATS:
            for rented in ("yes", "no"):
                query = (f'avg_over_time({prefix}_{stat}_dollars{{gpu_name=~"{everything}",'
                         f'verified="yes",rented="{rented}"}}[1d])')
                answers[(family, stat, rented)] = _farm_post(query, first, end)
                log(f"{family} {stat} rented={rented} {'ok' if answers[(family, stat, rented)] else 'failed'}")
                time.sleep(pause)
    tape = parse_500farm_price_tape(answers)
    if tape:
        write_price_tape(tape, CACHE / "gpu_prices_500farm_full.csv")
        write_price_tape(headline_price_tape(tape), OUT / "gpu_prices_500farm.csv")
    rows = []
    if counts["yes"] and counts["no"]:
        rows += parse_500farm_counts(counts["yes"], counts["no"])
    v2 = {key[2]: value for key, value in answers.items() if key[0] == "vastai_v2" and key[1] == "median"}
    headline = dict(datacenter="yes", gpu_count_range="8+")
    rows += parse_500farm_prices(_only(v2.get("yes"), **headline), _only(v2.get("no"), **headline))
    return rows


def series_label(row: dict) -> str:
    unit = f" ({row['unit']})" if row.get("unit") else ""
    provider = row["provider"]
    if row["source"] in OWN_SOURCES - {SRC_SHADEFORM}:
        head = row["source"]
    else:
        head = f"{row['source']} ({'all clouds' if provider == 'all' else provider})"
    return f"{head}: {row['metric'].replace('_', ' ')}{unit}"


def pivot_gpu(rows: list[dict], gpu: str) -> tuple[list[str], dict[str, dict[str, str]]]:
    """Series down the rows, dates across. Region-level rows stay in the long file."""
    table: dict[str, dict[str, str]] = {}
    for row in rows:
        if row["gpu"] != gpu or row.get("region"):
            continue
        value = row["value"]
        try:
            value = _clean(float(value))
        except ValueError:
            pass
        table.setdefault(series_label(row), {})[row["date"]] = value
    columns = sorted({day for values in table.values() for day in values})
    return columns, {name: table[name] for name in sorted(table)}


# Per-source wide files: (file, source, metric, row label). One row per GPU.
SOURCE_WIDE = [
    ("gpu_availability_500farm_pct.csv", SRC_FARM, ["pct_available"], "Vast GPUs not rented, % of listed"),
    ("gpu_availability_500farm_listed.csv", SRC_FARM, ["gpus_listed"], "Vast GPUs listed"),
    ("gpu_availability_500farm_price.csv", SRC_FARM_PRICE, ["median_usd_per_gpu_hour_rented"],
     "Rented, USD per GPU-hour, " + FARM_PRICE_NOTE),
    ("gpu_availability_akash.csv", SRC_AKASH, ["utilization_pct", "gpus_total"], "Akash GPUs and utilization"),
]


def pivot_source(rows: list[dict], source: str, metrics: list[str]) -> tuple[list[str], dict[str, dict[str, str]]]:
    table: dict[str, dict[str, str]] = {}
    for row in rows:
        if row["source"] != source or row["metric"] not in metrics or row.get("region"):
            continue
        name = row["gpu"] if len(metrics) == 1 else f"{row['gpu']}: {row['metric'].replace('_', ' ')}"
        table.setdefault(name, {})[row["date"]] = _clean(float(row["value"]))
    columns = sorted({day for values in table.values() for day in values})
    order = {gpu: index for index, gpu in enumerate(TARGET_GPUS)}
    names = sorted(table, key=lambda name: (order.get(name.split(":")[0], 50), name))
    return columns, {name: table[name] for name in names}


def build_wide(rows: list[dict] | None = None, out_dir: Path | None = None) -> list[Path]:
    rows = rows if rows is not None else read_long()
    wide = out_dir or WIDE
    written = []
    for gpu in WIDE_GPUS:
        columns, grid = pivot_gpu(rows, gpu)
        if not columns:
            continue
        path = wide / f"gpu_availability_{gpu.lower()}.csv"
        _write_wide(path, "Availability measure", columns, grid)
        written.append(path)
    for filename, source, metrics, header in SOURCE_WIDE:
        columns, grid = pivot_source(rows, source, metrics)
        if not columns:
            continue
        path = wide / filename
        _write_wide(path, header, columns, grid)
        written.append(path)
    return written


def write_snapshot(snapshot: dict) -> Path:
    SNAPSHOTS.mkdir(parents=True, exist_ok=True)
    stamp = snapshot["taken_at"].replace(":", "").replace("-", "")
    path = SNAPSHOTS / f"availability-{stamp}.json"
    path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
    return path


def import_observations(text: str) -> list[dict]:
    """Rows from a research CSV with the long-file columns (raw goes into note). Tracked chips only."""
    rows = []
    for record in csv.DictReader(io.StringIO(text)):
        if not record.get("date") or record.get("value") in (None, ""):
            continue
        row = {field: (record.get(field) or "").strip() for field in FIELDS}
        if not row["note"] and record.get("raw"):
            row["note"] = record["raw"].strip()[:200]
        row["gpu"] = normalize_gpu(row["gpu"])
        if row["gpu"] not in TARGET_GPUS:
            continue
        # "global" means the whole provider. A cloud tier is part of what was measured.
        if row["region"].lower() in ("global", "all"):
            row["region"] = ""
        elif row["region"].lower() in ("secure", "community"):
            row["metric"], row["region"] = f"{row['metric']}_{row['region'].lower()}", ""
        rows.append(row)
    return rows
