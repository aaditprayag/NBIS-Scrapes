"""Public reads on GPU price, which chips are still for sale, and Token Factory.

Two public Nebius pages refresh without an account: the GPU price table and
the docs page that lists which GPU platform is sold in which region. Vast.ai's
supply/demand metrics are the market thermometer, and they need a host API key
in VAST_API_KEY. OpenRouter's Nebius provider page is a token-factory readout.
"""

from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SNAPSHOTS = ROOT / "data" / "snapshots"
USER_AGENT = "NBIS-Scrapes/0.1 (personal research)"

ON_DEMAND_SKUS = [
    "NVIDIA GB300 NVL72",
    "NVIDIA HGX B300",
    "NVIDIA GB200 NVL72",
    "NVIDIA HGX B200",
    "NVIDIA HGX H200",
    "NVIDIA HGX H100",
    "NVIDIA RTX PRO 6000",
    "NVIDIA L40S with Intel CPU",
    "NVIDIA L40S with AMD CPU",
]

NEBIUS_PRICES_URL = "https://nebius.com/prices"
NEBIUS_PLATFORMS_URL = "https://docs.nebius.com/compute/virtual-machines/types"
OPENROUTER_PROVIDER_URL = "https://openrouter.ai/provider/nebius"
VAST_METRICS_URL = "https://console.vast.ai/api/v0/metrics/gpu/current/?verified=yes&hosting_type=secure_cloud"


def html_to_lines(text: str) -> str:
    """Turn a Nebius marketing page into the same line layout as the markdown fixture."""
    if "<" not in text:
        return text
    cleaned = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", text)
    cleaned = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", cleaned)
    cleaned = re.sub(r"(?i)</(p|div|tr|li|h1|h2|h3|td|th|section)>", "\n", cleaned)
    cleaned = re.sub(r"<[^>]+>", "", cleaned)
    cleaned = cleaned.replace("&nbsp;", " ").replace("&amp;", "&")
    return cleaned


def _section(text: str, start: str, end: str | None) -> str:
    begin = text.find(start)
    if begin < 0:
        return ""
    if end is None:
        return text[begin:]
    finish = text.find(end, begin + len(start))
    if finish < 0:
        return text[begin:]
    return text[begin:finish]


def _sku_at(line: str) -> str | None:
    for sku in ON_DEMAND_SKUS:
        if line.startswith(sku):
            return sku
    return None


def _parse_sku_blocks(text: str, kind: str) -> list[dict]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    rows = []
    index = 0
    while index < len(lines):
        sku = _sku_at(lines[index])
        if sku is None:
            index += 1
            continue
        block = []
        index += 1
        while index < len(lines) and _sku_at(lines[index]) is None and len(block) < 8:
            block.append(lines[index])
            index += 1
        joined = " ".join(block)
        dollars = [float(value) for value in re.findall(r"\$([0-9]+(?:\.[0-9]+)?)", joined)]
        contact = "contact" in joined.lower()
        row = {"sku": sku, "kind": kind, "contact_sales": contact and not dollars}
        if kind == "on_demand":
            row["usd_per_gpu_hour"] = dollars[0] if dollars else None
            row["next_usd_per_gpu_hour"] = dollars[1] if len(dollars) > 1 else None
        else:
            row["spot_floor_usd_per_gpu_hour"] = dollars[0] if dollars else None
        rows.append(row)
    return rows


def parse_nebius_prices(text: str) -> dict:
    text = html_to_lines(text)
    on_demand = _parse_sku_blocks(
        _section(text, "NVIDIA GPU Instances", "Preemptible"),
        "on_demand",
    )
    spot = _parse_sku_blocks(
        _section(text, "Preemptible", "CPU-only"),
        "spot",
    )
    hikes = []
    for row in on_demand:
        current = row.get("usd_per_gpu_hour")
        nxt = row.get("next_usd_per_gpu_hour")
        if current and nxt and nxt != current:
            hikes.append({
                "sku": row["sku"],
                "from_usd": current,
                "to_usd": nxt,
                "change_pct": round((nxt - current) / current, 4),
            })
    return {"on_demand": on_demand, "spot": spot, "scheduled_hikes": hikes}


def _platforms_from_html(text: str) -> list[dict]:
    rows = []
    for row_html in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", text):
        cells = re.findall(r"(?is)<td[^>]*>(.*?)</td>", row_html)
        if len(cells) < 4:
            continue
        platform = " ".join(re.sub(r"<[^>]+>", " ", cells[1]).split())
        if not platform.startswith("gpu-"):
            continue
        label = " ".join(re.sub(r"<[^>]+>", " ", cells[0]).split())
        regions = re.findall(r"<code>([a-z0-9-]+)</code>", cells[-1])
        rows.append({"platform": platform, "label": label, "regions": regions})
    return rows


def parse_nebius_platforms(text: str) -> list[dict]:
    rows = _platforms_from_html(text)
    if rows:
        seen = set()
        unique = []
        for row in rows:
            if row["platform"] in seen:
                continue
            seen.add(row["platform"])
            unique.append(row)
        return unique
    pattern = re.compile(
        r"\|\s*(NVIDIA[^|\n]+?)\s*\|\s*`?(gpu-[a-z0-9-]+)`?\s*\|[^|\n]*\|\s*([^|\n]+)\|",
        re.IGNORECASE,
    )
    for match in pattern.finditer(text):
        regions = re.findall(r"`([a-z0-9-]+)`", match.group(3))
        if not regions:
            regions = [part.strip(" `") for part in match.group(3).split(",") if part.strip()]
        rows.append({
            "platform": match.group(2),
            "label": " ".join(match.group(1).split()),
            "regions": regions,
        })
    # The docs page repeats platforms in later tables. Keep the first row per id.
    seen = set()
    unique = []
    for row in rows:
        if row["platform"] in seen:
            continue
        seen.add(row["platform"])
        unique.append(row)
    return unique


def parse_openrouter_provider_page(text: str) -> dict:
    match = re.search(r"Access\s+(\d+)\s+models", text, re.IGNORECASE)
    return {
        "models_listed": int(match.group(1)) if match else None,
        "names_nebius": "nebius" in text.lower(),
    }


def summarize_vast_metrics(payload: dict) -> dict:
    """Accept either a list or a dict with a list under a common key."""
    if isinstance(payload, list):
        items = payload
    else:
        items = None
        for key in ("gpus", "data", "metrics", "results"):
            if isinstance(payload.get(key), list):
                items = payload[key]
                break
        if items is None:
            items = []
    parsed = []
    for item in items:
        if not isinstance(item, dict):
            continue
        name = item.get("gpu_name") or item.get("name") or item.get("gpu")
        if not name:
            continue
        parsed.append({
            "gpu": name,
            "supply": item.get("supply") or item.get("available") or item.get("num_available"),
            "demand": item.get("demand") or item.get("rented") or item.get("num_rented"),
            "price_p10": item.get("p10") or item.get("price_p10") or item.get("dph_p10"),
            "price_median": item.get("median") or item.get("price_median") or item.get("dph_median"),
            "price_p90": item.get("p90") or item.get("price_p90") or item.get("dph_p90"),
        })
    return {"gpus": parsed, "count": len(parsed)}


def _get(url: str, headers: dict | None = None, timeout: int = 45) -> tuple[int, str]:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, **(headers or {})})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status, response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        return exc.code, body


def collect_live(timeout: int = 45) -> dict:
    """Fetch the public pages. Vast runs only when VAST_API_KEY is set."""
    taken_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    snapshot = {"taken_at": taken_at, "sources": {}}

    status, prices = _get(NEBIUS_PRICES_URL, timeout=timeout)
    snapshot["sources"]["nebius_prices"] = {
        "url": NEBIUS_PRICES_URL,
        "http_status": status,
        "parsed": parse_nebius_prices(prices) if status == 200 else None,
        "error": None if status == 200 else prices[:500],
    }

    status, platforms = _get(NEBIUS_PLATFORMS_URL, timeout=timeout)
    snapshot["sources"]["nebius_platforms"] = {
        "url": NEBIUS_PLATFORMS_URL,
        "http_status": status,
        "parsed": parse_nebius_platforms(platforms) if status == 200 else None,
        "error": None if status == 200 else platforms[:500],
    }

    status, provider = _get(OPENROUTER_PROVIDER_URL, timeout=timeout)
    snapshot["sources"]["openrouter_nebius"] = {
        "url": OPENROUTER_PROVIDER_URL,
        "http_status": status,
        "parsed": parse_openrouter_provider_page(provider) if status == 200 else None,
        "error": None if status == 200 else provider[:500],
    }

    api_key = os.environ.get("VAST_API_KEY", "").strip()
    if not api_key:
        snapshot["sources"]["vast_metrics"] = {
            "url": VAST_METRICS_URL,
            "http_status": None,
            "parsed": None,
            "error": "skipped: set VAST_API_KEY to a Vast host key. Client keys are rejected.",
        }
    else:
        status, body = _get(
            VAST_METRICS_URL,
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=timeout,
        )
        parsed = None
        error = None
        if status == 200:
            try:
                parsed = summarize_vast_metrics(json.loads(body))
            except json.JSONDecodeError as exc:
                error = f"json: {exc}"
        else:
            error = body[:500]
        snapshot["sources"]["vast_metrics"] = {
            "url": VAST_METRICS_URL,
            "http_status": status,
            "parsed": parsed,
            "error": error,
        }
    return snapshot


def write_snapshot(snapshot: dict | None = None) -> Path:
    snapshot = snapshot or collect_live()
    SNAPSHOTS.mkdir(parents=True, exist_ok=True)
    stamp = snapshot["taken_at"].replace(":", "").replace("-", "")
    path = SNAPSHOTS / f"demand-{stamp}.json"
    path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
    return path
