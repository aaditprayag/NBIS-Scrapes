"""Nebius list prices from archived copies of nebius.com/prices.

The live page only shows the current card. The Internet Archive has the
older cards. Layouts changed:

- 2024: an on-demand column and a reserve column, including 8-GPU rows.
- 2025 to early 2026: one on-demand price per GPU.
- From mid-2026: a preemptible column and an on-demand column.
- October 2026: on-demand, a scheduled next price, and a spot floor.
"""

from __future__ import annotations

import gzip
import html
import json
import re
import time
import urllib.request
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
USER_AGENT = "NBIS-Scrapes/0.1 (personal research; archive.org)"
CDX_URL = (
    "https://web.archive.org/cdx/search/cdx?url=nebius.com/prices"
    "&output=json&fl=timestamp,statuscode,digest,length,original"
    "&filter=statuscode:200&collapse=digest"
)

TOKENS = ("GB300", "GB200", "B300", "B200", "H200", "H100", "RTX PRO 6000", "L40S", "L40s")


def html_to_lines(text: str) -> list[str]:
    text = html.unescape(text)
    if "<" in text:
        cleaned = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", text)
        cleaned = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", cleaned)
        cleaned = re.sub(r"(?i)</(p|div|tr|li|h1|h2|h3|h4|td|th|span|section)>", "\n", cleaned)
        cleaned = re.sub(r"<[^>]+>", "", cleaned)
    else:
        cleaned = text
    cleaned = cleaned.replace("\xa0", " ")
    return [line.strip() for line in cleaned.splitlines() if line.strip()]


def _sku(line: str) -> dict | None:
    if len(line) > 90:
        return None
    clean = line.replace("®", "").replace("*", "").strip()
    counted = re.match(r"(\d+)\s*[×x]\s*(H200|H100|L40S|L40s)", clean, re.IGNORECASE)
    if counted:
        name = counted.group(2).upper()
        return {"sku": "L40S" if name.startswith("L40") else name, "gpus": int(counted.group(1)), "label": clean}
    if not clean.lower().startswith("nvidia"):
        return None
    upper = clean.upper()
    for token in TOKENS:
        if token.upper() in upper:
            sku = "RTX PRO 6000" if token == "RTX PRO 6000" else token.upper()
            if sku.startswith("L40"):
                if "AMD" in upper:
                    sku = "L40S-AMD"
                elif "INTEL" in upper:
                    sku = "L40S-Intel"
                else:
                    sku = "L40S"
            return {"sku": sku, "gpus": 1, "label": clean}
    return None


def _dollars(line: str) -> list[float]:
    return [float(value) for value in re.findall(r"\$([0-9]+(?:\.[0-9]+)?)", line)]


def _per_gpu(amount: float, gpus: int) -> float:
    if gpus > 1:
        return round(amount / gpus, 4)
    return amount


def _mode_after(line: str, mode: str) -> str:
    """Remember which price columns are in force until the table ends."""
    low = line.lower().strip()
    if low.startswith("cpu-only") or low.startswith("storage") or low.startswith("disks,"):
        return "none"
    if "spot price" in low:
        return "spot"
    if "effective october" in low or low.startswith("gpu-hour (effective"):
        return "od_next"
    if "preemptible, gpu-hour" in low:
        return "seen_preempt"
    if low.startswith("on-demand, gpu-hour"):
        return "preempt_od" if mode == "seen_preempt" else "single"
    if low == "price per gpu-hour":
        return "single"
    if low == "on demand":
        return "seen_on_demand"
    if low == "reserve" and mode == "seen_on_demand":
        return "od_reserve"
    if "reserve your best price" in low:
        return "cards"
    return mode


def extract_rates(text: str) -> list[dict]:
    """Return one row per price found on a single archived or live page."""
    lines = html_to_lines(text)
    found = []
    index = 0
    mode = "none"
    while index < len(lines):
        sku = _sku(lines[index])
        if sku is None:
            mode = _mode_after(lines[index], mode)
            index += 1
            continue
        block = []
        nxt = index + 1
        while nxt < len(lines) and _sku(lines[nxt]) is None and len(block) < 8:
            block.append(lines[nxt])
            nxt += 1
        block_text = " ".join(block)
        dollars = []
        for line in block:
            dollars.extend(_dollars(line))
        contact = "contact us" in block_text.lower() or "contact sales" in block_text.lower()
        gpus = sku["gpus"]
        kinds: list[tuple[str, float | None]] = []
        low_block = block_text.lower()
        if "commitment" in low_block or "as low as" in low_block:
            if dollars:
                kinds.append(("commitment", _per_gpu(dollars[0], gpus)))
        elif mode == "preempt_od":
            if len(dollars) >= 2:
                kinds.append(("preemptible", _per_gpu(dollars[0], gpus)))
                kinds.append(("on_demand", _per_gpu(dollars[1], gpus)))
            elif len(dollars) == 1:
                kinds.append(("on_demand", _per_gpu(dollars[0], gpus)))
        elif mode == "od_next":
            if dollars:
                kinds.append(("on_demand", _per_gpu(dollars[0], gpus)))
            if len(dollars) >= 2:
                kinds.append(("scheduled_next", _per_gpu(dollars[1], gpus)))
        elif mode == "od_reserve":
            if dollars:
                kinds.append(("on_demand", _per_gpu(dollars[0], gpus)))
            if len(dollars) >= 2:
                kinds.append(("reserve", _per_gpu(dollars[1], gpus)))
        elif mode == "spot" and dollars:
            kinds.append(("spot_floor", _per_gpu(dollars[0], gpus)))
        elif mode == "single" and dollars:
            kinds.append(("on_demand", _per_gpu(dollars[0], gpus)))
        elif dollars and gpus == 1:
            kinds.append(("unscoped", _per_gpu(dollars[0], gpus)))
        if contact and not kinds:
            kinds.append(("contact_sales", None))
        for kind, price in kinds:
            found.append({
                "sku": sku["sku"],
                "gpus": gpus,
                "kind": kind,
                "usd_per_gpu_hour": price,
                "label": sku["label"],
            })
        for line in block:
            mode = _mode_after(line, mode)
        index = nxt
    return found


def _decode(raw: bytes) -> str:
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    return raw.decode("utf-8", errors="replace")


def _get(url: str, timeout: int = 60) -> bytes:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": USER_AGENT, "Accept-Encoding": "identity"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def cdx_timestamps() -> list[str]:
    payload = json.loads(_decode(_get(CDX_URL)))
    stamps = []
    for row in payload[1:]:
        stamps.append(row[0])
    return stamps


def wayback_url(timestamp: str) -> str:
    return f"https://web.archive.org/web/{timestamp}id_/https://nebius.com/prices"


def fetch_capture(timestamp: str, attempts: int = 4) -> list[dict]:
    last_error: Exception | None = None
    raw = b""
    for attempt in range(attempts):
        try:
            raw = _get(wayback_url(timestamp), timeout=90)
            last_error = None
            break
        except Exception as exc:
            last_error = exc
            time.sleep(3 + attempt * 4)
    if last_error is not None:
        raise last_error
    text = _decode(raw)
    captured = datetime.strptime(timestamp, "%Y%m%d%H%M%S").strftime("%Y-%m-%d")
    rows = []
    for rate in extract_rates(text):
        if rate["gpus"] != 1:
            continue
        rows.append({
            "captured_on": captured,
            "wayback_timestamp": timestamp,
            "wayback_url": f"https://web.archive.org/web/{timestamp}/https://nebius.com/prices",
            **rate,
        })
    return rows


def fetch_all(limit: int | None = None, workers: int = 1, stamps: list[str] | None = None) -> list[dict]:
    stamps = list(stamps) if stamps is not None else cdx_timestamps()
    if limit is not None:
        stamps = stamps[:limit]
    rows: list[dict] = []
    errors: list[str] = []
    # One at a time. The archive refuses bursts.
    for stamp in stamps:
        try:
            rows.extend(fetch_capture(stamp))
        except Exception as exc:
            errors.append(f"{stamp}: {type(exc).__name__}: {exc}")
        time.sleep(1.2)
    rows.sort(key=lambda row: (row["captured_on"], row["sku"], row["kind"]))
    if errors:
        (ROOT / "data" / "out").mkdir(parents=True, exist_ok=True)
        (ROOT / "data" / "out" / "wayback_errors.txt").write_text("\n".join(errors), encoding="utf-8")
    return rows


def on_demand_changes(rows: list[dict]) -> list[dict]:
    """First time each SKU's on-demand rate changes, plus the first print."""
    series: dict[str, list[dict]] = {}
    for row in rows:
        if row["kind"] != "on_demand" or row["usd_per_gpu_hour"] is None:
            continue
        series.setdefault(row["sku"], []).append(row)
    changes = []
    for sku, points in series.items():
        points.sort(key=lambda row: row["captured_on"])
        last = None
        for point in points:
            price = point["usd_per_gpu_hour"]
            if last is None or price != last:
                changes.append({
                    "sku": sku,
                    "captured_on": point["captured_on"],
                    "usd_per_gpu_hour": price,
                    "previous_usd": last,
                    "wayback_url": point["wayback_url"],
                })
                last = price
    return changes
