"""Nebius GPU prices as other people published them.

Three sources carry a dated copy of Nebius's GPU card that Nebius itself did
not keep:

- dstack's gpuhunt catalog. A GitHub Action rebuilds it about every hour and
  writes a zip to a public S3 bucket. The archive has a Nebius file from
  8 Nov 2023, one version per day.
- SkyPilot's catalog repo. `catalogs/v*/nebius/vms.csv` is regenerated from
  the Nebius billing API, and every change is a git commit.
- Public GitHub price trackers and GPU price sites that list Nebius next to
  other clouds.

None of them sets a Nebius price. A third-party price is checked against the
Nebius list in force that day, from the rate card and the archived price page
already in this repo. If it matches to the cent, that day counts as a day the
Nebius price was seen. If it does not match, the row stays in the file and is
marked as a miss. Monthly roll-ups carry a made-up date, so they never count.
"""

from __future__ import annotations

import csv
import io
import json
import re
import struct
import time
import urllib.error
import urllib.request
import zlib
from collections import defaultdict
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

from nbis_scrapes.wide_prices import OUT, _clean, _write_wide, pivot_docs, pivot_page

USER_AGENT = "NBIS-Scrapes/0.1 (personal research)"
GPUHUNT_BUCKET = "https://dstack-gpu-pricing.s3.eu-west-1.amazonaws.com"
GPUHUNT_LATEST = f"{GPUHUNT_BUCKET}/v3/nebius/latest/catalog.zip"
SKYPILOT_REPO = "skypilot-org/skypilot-catalog"
SKYPILOT_PATHS = ("catalogs/v6/nebius/vms.csv", "catalogs/v7/nebius/vms.csv", "catalogs/v8/nebius/vms.csv", "catalogs/v9/nebius/vms.csv")

GPUHUNT = "dstack gpuhunt catalog"
SKYPILOT = "SkyPilot catalog (git)"

FIELDS = [
    "date",
    "source",
    "source_url",
    "observed_ts",
    "item",
    "region",
    "price_type",
    "gpu_count",
    "price_usd",
    "usd_per_gpu_hour",
    "raw",
    "nebius_sku",
    "nebius_list_usd",
    "matches_nebius_list",
]

# Cents. A tracker that rounds $2.95 to $2.9 is a miss.
TOLERANCE = 0.005
LIST_SKUS = ("H100", "H200", "B200", "B300", "RTX PRO 6000", "L40S-Intel", "L40S-AMD")


def canonical_sku(item: str, raw: str = "") -> str:
    """Name a third-party GPU the way the rate-card files name it."""
    text = f"{item} {raw}".upper().replace("_", " ").replace("-", " ")
    name = item.upper().replace("_", " ").replace("-", " ")
    if "GB300" in name:
        return "GB300"
    if "GB200" in name:
        return "GB200"
    if re.search(r"\bB300\b", name):
        return "B300"
    if re.search(r"\bB200\b", name):
        return "B200"
    if re.search(r"\bH200\b", name):
        return "H200"
    if re.search(r"\bH100\b", name):
        return "H100"
    if re.search(r"RTX ?PRO ?6000|RTX ?6000|RTXPRO6000", name):
        return "RTX PRO 6000"
    if re.search(r"\bL40S\b", name):
        if re.search(r"L40S D\b|\bAMD\b|GENOA", text):
            return "L40S-AMD"
        if re.search(r"L40S A\b|\bINTEL\b|ICE ?LAKE", text):
            return "L40S-Intel"
        return "L40S"
    return item.strip()


def _row(date: str, source: str, url: str, observed_ts: str, item: str, region: str, price_type: str,
         gpu_count: int, price: float, raw: str) -> dict:
    return {
        "date": date,
        "source": source,
        "source_url": url,
        "observed_ts": observed_ts,
        "item": item,
        "region": region,
        "price_type": price_type,
        "gpu_count": gpu_count,
        "price_usd": _clean(price),
        "usd_per_gpu_hour": _clean(round(price / gpu_count, 6)),
        "raw": raw,
    }


def parse_gpuhunt(text: str, date: str, url: str, observed_ts: str) -> list[dict]:
    """One row per GPU, region, and price type: the cheapest preset per GPU."""
    best: dict[tuple, dict] = {}
    for record in csv.DictReader(io.StringIO(text)):
        count = int(float(record.get("gpu_count") or 0))
        if count <= 0 or not record.get("gpu_name"):
            continue
        price = float(record["price"])
        price_type = "preemptible" if record["spot"].strip().lower() == "true" else "on_demand"
        raw = ",".join(record.get(key, "") for key in (
            "instance_name", "location", "price", "cpu", "memory", "gpu_count", "gpu_name", "gpu_memory", "spot"
        ))
        row = _row(date, GPUHUNT, url, observed_ts, record["gpu_name"], record["location"], price_type, count, price, raw)
        key = (record["gpu_name"], record["location"], price_type)
        current = best.get(key)
        rank = (price / count, count, raw)
        if current is None or rank < (float(current["usd_per_gpu_hour"]), current["gpu_count"], current["raw"]):
            best[key] = row
    return sorted(best.values(), key=lambda row: (row["item"], row["region"], row["price_type"]))


def parse_skypilot(text: str, date: str, url: str, observed_ts: str) -> list[dict]:
    """Every GPU instance in one copy of SkyPilot's Nebius vms.csv."""
    rows = []
    for record in csv.DictReader(io.StringIO(text)):
        count = int(float(record.get("AcceleratorCount") or 0))
        if count <= 0 or not record.get("AcceleratorName"):
            continue
        raw = (
            f"{record['InstanceType']},{record['AcceleratorName']},{record['AcceleratorCount']},"
            f"{record['Region']},Price={record['Price']},SpotPrice={record.get('SpotPrice') or ''}"
        )
        for price_type, column in (("on_demand", "Price"), ("preemptible", "SpotPrice")):
            value = (record.get(column) or "").strip()
            if not value:
                continue
            rows.append(_row(date, SKYPILOT, url, observed_ts, record["AcceleratorName"], record["Region"],
                             price_type, count, float(value), raw))
    return rows


def nebius_list(docs_rows: list[dict], page_rows: list[dict]) -> dict[tuple[str, str], list[tuple[str, float]]]:
    """The Nebius list price in force, as dated steps per SKU and price type.

    The rate card wins from its first dated price. Before that, the archived
    price page fills in. Component lines are left out, because the page and
    every third party quote the whole VM. The page kept its old on-demand column after the
    1 Oct 2026 change, so it cannot be trusted once the card has a date.
    """
    steps: dict[tuple[str, str], list[tuple[str, float]]] = {}
    # A line ending ". GPU" is the GPU part of a VM. vCPU and RAM are billed on top.
    docs_rows = [row for row in docs_rows if not row["item"].rstrip().endswith(". GPU")]
    for price_type, preemptible, page_kind in (("on_demand", False, "on_demand"), ("preemptible", True, "preemptible")):
        _columns, card = pivot_docs(docs_rows, preemptible, False)
        _columns, page = pivot_page(page_rows, page_kind)
        for sku in set(card) | set(page):
            if sku.endswith("GPU component") or sku.endswith("me-west1"):
                continue
            card_days = sorted(card.get(sku, {}).items())
            start = card_days[0][0] if card_days else "9999-12-31"
            page_days = [(day, value) for day, value in sorted(page.get(sku, {}).items()) if day < start]
            points = [(day, float(value)) for day, value in page_days + card_days]
            if points:
                steps[(sku, price_type)] = points
    return steps


def list_price_on(steps: dict, sku: str, price_type: str, day: str) -> float | None:
    points = steps.get((sku, price_type)) or []
    price = None
    for start, value in points:
        if start > day:
            break
        price = value
    return price


def corroborate(rows: list[dict], steps: dict) -> list[dict]:
    """Mark each third-party price as a match, a miss, or not comparable."""
    out = []
    for row in rows:
        sku = canonical_sku(row["item"], row.get("raw", ""))
        price_type = row["price_type"]
        if price_type == "spot" and row["date"] < "2026-10-08":
            price_type = "preemptible"
        candidates = ["L40S-Intel", "L40S-AMD"] if sku == "L40S" else [sku]
        listed = []
        if price_type in ("on_demand", "preemptible"):
            for name in candidates:
                price = list_price_on(steps, name, price_type, row["date"])
                if price is not None:
                    listed.append((name, price))
        value = float(row["usd_per_gpu_hour"])
        if sku not in LIST_SKUS and sku != "L40S":
            verdict, shown = "not on the Nebius card", ""
        elif price_type not in ("on_demand", "preemptible"):
            verdict, shown = "price type not comparable", ""
        elif not listed:
            verdict, shown = "no Nebius list on file for that day", ""
        else:
            hit = next(((name, price) for name, price in listed if abs(price - value) <= TOLERANCE + 1e-12), None)
            verdict = "yes" if hit else "no"
            shown = " / ".join(_clean(price) for _name, price in listed)
            if hit:
                shown = _clean(hit[1])
        out.append({**{field: row.get(field, "") for field in FIELDS}, "nebius_sku": sku,
                    "nebius_list_usd": shown or "none", "matches_nebius_list": verdict})
    return out


def counts_as_capture(source: str) -> bool:
    return "monthly" not in source.lower()


def corroborated_days(rows: list[dict], nebius_days: set[str]) -> list[dict]:
    """Days on which at least one third-party price matched the Nebius list."""
    tally: dict[str, dict[str, list[int]]] = defaultdict(lambda: defaultdict(lambda: [0, 0]))
    for row in rows:
        if row["matches_nebius_list"] not in ("yes", "no") or not counts_as_capture(row["source"]):
            continue
        counts = tally[row["date"]][row["source"]]
        counts[1] += 1
        if row["matches_nebius_list"] == "yes":
            counts[0] += 1
    days = []
    for day in sorted(tally):
        matched = {source: counts for source, counts in tally[day].items() if counts[0]}
        if not matched:
            continue
        days.append({
            "date": day,
            "nebius_capture_that_day": "yes" if day in nebius_days else "no",
            "matching_prices": sum(counts[0] for counts in matched.values()),
            "compared_prices": sum(counts[1] for counts in tally[day].values()),
            "sources": "; ".join(f"{source} ({counts[0]} of {counts[1]})" for source, counts in sorted(matched.items())),
        })
    return days


def family(source: str) -> str:
    if source == GPUHUNT:
        return "gpuhunt"
    if source == SKYPILOT:
        return "skypilot"
    return "trackers"


PRICE_LABEL = {"on_demand": "on-demand", "preemptible": "preemptible"}


def pivot_family(rows: list[dict], name: str) -> tuple[list[str], dict[str, dict[str, str]]]:
    table: dict[str, dict[str, str]] = {}
    order: dict[str, tuple] = {}
    rank = {sku: index for index, sku in enumerate(LIST_SKUS)}
    for row in rows:
        if family(row["source"]) != name:
            continue
        label = f"{row['nebius_sku']} {PRICE_LABEL.get(row['price_type'], row['price_type'])}"
        if row["region"]:
            label += f", {row['region']}"
        if name == "trackers":
            label = f"{row['source']}: {label}"
        order[label] = (row["source"] if name == "trackers" else "", rank.get(row["nebius_sku"], 100), label)
        cell = table.setdefault(label, {})
        value = _clean(row["usd_per_gpu_hour"])
        if row["date"] not in cell or float(value) < float(cell[row["date"]]):
            cell[row["date"]] = value
    columns = sorted({day for values in table.values() for day in values})
    return columns, {label: table[label] for label in sorted(table, key=order.get)}


def merge(existing: list[dict], new: list[dict]) -> list[dict]:
    seen = {}
    for row in existing + new:
        key = (row["date"], row["source"], row["item"], row["region"], row["price_type"], str(row["gpu_count"]), row["raw"])
        seen.setdefault(key, row)
    return sorted(seen.values(), key=lambda row: (row["date"], row["source"], row["item"], row["region"], row["price_type"]))


def from_observation_file(path: Path) -> list[dict]:
    """Read a GPU observation CSV in the shared research layout."""
    rows = []
    with path.open(encoding="utf-8") as handle:
        for record in csv.DictReader(handle):
            if record.get("kind", "gpu") != "gpu":
                continue
            source = record["source"]
            if source.lower().startswith(("nebius", "docs.nebius")):
                continue
            count = int(float(record.get("gpu_count") or 1))
            price = float(record["price_usd"])
            per_gpu = record.get("price_per_gpu_hour") or record.get("usd_per_gpu_hour") or price / count
            row = _row(record["date"], source, record["source_url"], record["observed_ts"], record["item"],
                       record.get("region") or "", record["price_type"], count, price, record.get("raw") or "")
            row["usd_per_gpu_hour"] = _clean(per_gpu)
            rows.append(row)
    return rows


def _request(url: str, headers: dict | None = None, method: str = "GET", timeout: int = 60):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, **(headers or {})}, method=method)
    return urllib.request.urlopen(request, timeout=timeout)


def _zip_member(blob: bytes, member: str) -> bytes | None:
    end = blob.rfind(b"PK\x05\x06")
    if end < 0:
        return None
    size, offset = struct.unpack("<II", blob[end + 12:end + 20])
    directory = blob[offset:offset + size]
    position = 0
    while position + 46 <= len(directory) and directory[position:position + 4] == b"PK\x01\x02":
        method = struct.unpack("<H", directory[position + 10:position + 12])[0]
        packed = struct.unpack("<I", directory[position + 20:position + 24])[0]
        name_len, extra_len, comment_len = struct.unpack("<HHH", directory[position + 28:position + 34])
        local = struct.unpack("<I", directory[position + 42:position + 46])[0]
        name = directory[position + 46:position + 46 + name_len].decode("utf-8")
        if name == member:
            local_name, local_extra = struct.unpack("<HH", blob[local + 26:local + 30])
            start = local + 30 + local_name + local_extra
            data = blob[start:start + packed]
            return zlib.decompressobj(-15).decompress(data) if method == 8 else data
        position += 46 + name_len + extra_len + comment_len
    return None


def fetch_gpuhunt_latest() -> list[dict]:
    """Today's Nebius file from the gpuhunt bucket."""
    with _request(GPUHUNT_LATEST) as response:
        blob = response.read()
        modified = parsedate_to_datetime(response.headers["Last-Modified"]).astimezone(timezone.utc)
    text = _zip_member(blob, "nebius.csv")
    if text is None:
        return []
    return parse_gpuhunt(text.decode("utf-8"), modified.strftime("%Y-%m-%d"), GPUHUNT_LATEST, modified.strftime("%Y%m%d%H%M%S"))


def _github_json(url: str) -> list | dict:
    with _request(url, headers={"Accept": "application/vnd.github+json"}) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_skypilot(since: str = "") -> list[dict]:
    """Every commit to SkyPilot's Nebius catalog after `since`, one copy each."""
    rows = []
    for path in SKYPILOT_PATHS:
        page = 1
        while True:
            url = f"https://api.github.com/repos/{SKYPILOT_REPO}/commits?path={path}&per_page=100&page={page}"
            if since:
                url += f"&since={since}T00:00:00Z"
            try:
                commits = _github_json(url)
            except urllib.error.HTTPError:
                break
            if not commits:
                break
            for commit in commits:
                sha = commit["sha"]
                stamp = datetime.fromisoformat(commit["commit"]["committer"]["date"].replace("Z", "+00:00"))
                raw_url = f"https://raw.githubusercontent.com/{SKYPILOT_REPO}/{sha}/{path}"
                try:
                    with _request(raw_url) as response:
                        text = response.read().decode("utf-8")
                except urllib.error.HTTPError:
                    continue
                rows.extend(parse_skypilot(text, stamp.strftime("%Y-%m-%d"), raw_url, stamp.strftime("%Y%m%d%H%M%S")))
                time.sleep(1)
            page += 1
    return rows


def load(path: Path | None = None) -> list[dict]:
    path = path or (OUT / "third_party_gpu_prices.csv")
    if not path.exists():
        return []
    with path.open(encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def nebius_capture_days() -> set[str]:
    days = set()
    for name in ("gpu_price_history.csv", "docs_gpu_prices.csv"):
        path = OUT / name
        if path.exists():
            with path.open(encoding="utf-8") as handle:
                days.update(row["captured_on"] for row in csv.DictReader(handle))
    return days


def write_outputs(rows: list[dict], out_dir: Path | None = None) -> list[Path]:
    out_dir = out_dir or OUT
    wide = out_dir / "wide"
    with (OUT / "docs_gpu_prices.csv").open(encoding="utf-8") as handle:
        docs = list(csv.DictReader(handle))
    with (OUT / "gpu_price_history.csv").open(encoding="utf-8") as handle:
        page = list(csv.DictReader(handle))
    checked = corroborate(rows, nebius_list(docs, page))
    written = []
    long_path = out_dir / "third_party_gpu_prices.csv"
    _write(long_path, checked, FIELDS)
    written.append(long_path)
    days = corroborated_days(checked, nebius_capture_days())
    days_path = out_dir / "third_party_corroborated_days.csv"
    _write(days_path, days, ["date", "nebius_capture_that_day", "matching_prices", "compared_prices", "sources"])
    written.append(days_path)
    for name in ("gpuhunt", "skypilot", "trackers"):
        columns, grid = pivot_family(checked, name)
        if not columns:
            continue
        path = wide / f"third_party_{name}.csv"
        _write_wide(path, "USD per GPU-hour", columns, grid)
        written.append(path)
    return written


def _write(path: Path, rows: list[dict], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
