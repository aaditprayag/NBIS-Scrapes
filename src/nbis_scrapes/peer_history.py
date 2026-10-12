"""Archived on-demand prices for CoreWeave, Crusoe, and Lambda.

Each capture is one published card. The first dollar after a GPU name is the
on-demand rate. A price on an 8-GPU line that is large is a node price and is
divided by the GPU count. A price under $20 on an 8x line is already per GPU,
which is how Lambda wrote "On-demand 8x H100 $2.99".
"""

from __future__ import annotations

import csv
import gzip
import json
import re
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

from nbis_scrapes.market_tapes import nebius_history
from nbis_scrapes.wide_prices import OUT, WIDE, _clean, _write_wide

ROOT = Path(__file__).resolve().parents[2]
USER_AGENT = "NBIS-Scrapes/0.1 (personal research; archive.org)"

SOURCES = (
    ("CoreWeave", "https://www.coreweave.com/pricing/classic"),
    ("CoreWeave", "https://www.coreweave.com/pricing"),
    ("Crusoe", "https://www.crusoe.ai/cloud/pricing"),
    ("Lambda", "https://lambdalabs.com/service/gpu-cloud"),
    ("Lambda", "https://lambda.ai/pricing"),
)

FAMILIES = (
    "GB300", "GB200", "GH200", "B300", "B200", "H200", "H100",
    "L40S", "L40", "A100", "A40", "RTX PRO 6000", "RTX A6000", "A6000",
    "A10", "A30", "V100", "T4",
)


def lines_of(html: str) -> list[str]:
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"\$\s*</[^>]+>\s*", "$", text)
    plain = re.sub(r"<[^>]+>", "\n", text)
    raw = [re.sub(r"\s+", " ", line).strip() for line in plain.splitlines() if line.strip()]
    lines = []
    index = 0
    while index < len(raw):
        if raw[index] == "$" and index + 1 < len(raw) and re.match(r"^[0-9]", raw[index + 1]):
            lines.append("$" + raw[index + 1].split()[0])
            index += 2
        else:
            lines.append(raw[index])
            index += 1
    return lines


def _family(text: str) -> str | None:
    upper = text.upper()
    if "GEFORCE" in upper or "LAPTOP" in upper or "DESKTOP" in upper:
        return None
    for name in FAMILIES:
        if name in upper:
            return name
    return None


def _sku(name: str, following: list[str]) -> str | None:
    family = _family(name)
    if family is None:
        return None
    blob = " ".join([name, *following[:5]]).upper()
    memory = re.search(r"\b(180|141|96|80|48|40|32|24|16)\s*GB\b", blob)
    form = ""
    if re.search(r"\bSXM\b|\bHGX\b|\bNVLINK\b", blob):
        form = "SXM"
    elif re.search(r"\bPCIE\b|\bPCI-E\b", blob):
        form = "PCIe"
    if family in {"H100", "A100", "A40", "A6000"} and (memory or form):
        bits = [family]
        if memory:
            bits.append(memory.group(1) + "GB")
        if form:
            bits.append(form)
        return " ".join(bits)
    if family == "H200" and "NVL" in blob and "NVL72" not in blob:
        return "H200 NVL"
    return family


def _first_price(lines: list[str]) -> float | None:
    for line in lines:
        match = re.search(r"\$\s*([0-9]+(?:\.[0-9]+)?)", line)
        if match:
            return float(match.group(1))
    return None


def _gpu_count(name: str, following: list[str]) -> int | None:
    match = re.search(r"\b(\d+)\s*x\b", name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    for line in following[:3]:
        if re.fullmatch(r"[1-8]", line):
            return int(line)
    if "NVL72" in name.upper():
        return 4
    return None


def _per_gpu(name: str, following: list[str], price: float) -> tuple[float, str] | None:
    count = _gpu_count(name, following)
    # A node price is at least about one dollar per GPU on the line. A small
    # number next to "8x H100" is already the per-GPU on-demand rate.
    if count and count > 1 and price >= count * 0.9:
        per_gpu = price / count
        if per_gpu <= 20:
            return round(per_gpu, 4), f"{price} for {count} GPUs, divided"
    if count and count > 1 and "on-demand" not in name.lower():
        per_gpu = price / count
        if 0.45 <= per_gpu <= 12 and price < 20:
            return round(per_gpu, 4), f"{price} for {count} GPUs, divided"
    if price > 40:
        return None
    return price, "on-demand per GPU"


def extract_nvidia_ondemand(html: str) -> list[dict]:
    lines = lines_of(html)
    found = []
    seen = set()
    for index, line in enumerate(lines):
        if "NVIDIA" not in line.upper() and "A100" not in line.upper():
            continue
        if _family(line) is None and "A100" not in line.upper():
            continue
        following = lines[index + 1:index + 9]
        # Stop the look-ahead at the next GPU name.
        cut = len(following)
        this_family = _family(line)
        for offset, later in enumerate(following):
            other = _family(later)
            if other and other != this_family:
                cut = offset
                break
        following = following[:cut]
        sku = _sku(line, following)
        price = _first_price(following)
        if sku is None or price is None or sku in seen:
            continue
        previous = lines[index - 1] if index else ""
        if re.search(r"reserved|commit", line, re.IGNORECASE):
            continue
        if re.match(r"(?i)reserve\b", previous):
            continue
        price_line = next((item for item in following if "$" in item), "")
        if ("—" in price_line or "–" in price_line) and "on-demand" not in line.lower():
            continue
        per = _per_gpu(line, following, price)
        if per is None:
            continue
        amount, basis = per
        if amount < 0.4 or amount > 25:
            continue
        seen.add(sku)
        found.append({"sku": sku, "usd_per_gpu_hour": amount, "basis": basis})
    return found


def cdx(url: str) -> list[str]:
    query = (
        "https://web.archive.org/cdx/search/cdx?url="
        + urllib.parse.quote(url, safe=":/")
        + "&output=json&fl=timestamp,statuscode,length&filter=statuscode:200&collapse=digest"
    )
    request = urllib.request.Request(query, headers={"User-Agent": USER_AGENT, "Accept-Encoding": "identity"})
    with urllib.request.urlopen(request, timeout=90) as response:
        raw = response.read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    rows = json.loads(raw.decode("utf-8"))
    stamps = []
    for row in rows[1:]:
        if int(row[2] or 0) < 6000:
            continue
        stamps.append(row[0])
    return _monthly(stamps)


def _monthly(stamps: list[str]) -> list[str]:
    chosen = []
    seen = set()
    ordered = sorted(stamps)
    for stamp in ordered:
        key = stamp[:6]
        if key in seen:
            continue
        seen.add(key)
        chosen.append(stamp)
    if ordered and ordered[-1] not in chosen:
        chosen.append(ordered[-1])
    return chosen


def _fetch(stamp: str, url: str) -> str:
    archived = f"https://web.archive.org/web/{stamp}id_/{url}"
    last = None
    for attempt in range(3):
        try:
            request = urllib.request.Request(archived, headers={"User-Agent": USER_AGENT, "Accept-Encoding": "identity"})
            with urllib.request.urlopen(request, timeout=70) as response:
                raw = response.read()
            if raw[:2] == b"\x1f\x8b":
                raw = gzip.decompress(raw)
            return raw.decode("utf-8", errors="replace")
        except Exception as exc:
            last = exc
            time.sleep(2 + attempt * 3)
    raise last


def fetch_company(company: str, url: str) -> list[dict]:
    rows = []
    try:
        stamps = cdx(url)
    except Exception as exc:
        print(f"cdx failed {url}: {exc}")
        return rows
    for stamp in stamps:
        try:
            html = _fetch(stamp, url)
        except Exception as exc:
            print(f"fetch failed {stamp} {url}: {type(exc).__name__}")
            continue
        day = datetime.strptime(stamp, "%Y%m%d%H%M%S").strftime("%Y-%m-%d")
        for item in extract_nvidia_ondemand(html):
            sku = item["sku"]
            if "divided" in item["basis"]:
                sku = f"{sku} (node)"
            rows.append({
                "company": company,
                "sku": sku,
                "captured_on": day,
                "usd_per_gpu_hour": item["usd_per_gpu_hour"],
                "basis": item["basis"],
                "source_url": url,
            })
        time.sleep(0.8)
    return rows


def _prefer(existing: dict, row: dict) -> None:
    key = (row["company"], row["sku"], row["captured_on"])
    current = existing.get(key)
    if current is None:
        existing[key] = row
        return
    # A published per-GPU price beats a divided node price.
    if "divided" in current["basis"] and "divided" not in row["basis"]:
        existing[key] = row


def collect() -> list[dict]:
    chosen: dict[tuple, dict] = {}
    for company, url in SOURCES:
        for row in fetch_company(company, url):
            _prefer(chosen, row)
    for row in nebius_history():
        row = {**row, "source_url": "https://nebius.com/prices"}
        _prefer(chosen, row)
    rows = list(chosen.values())
    rows.sort(key=lambda row: (row["company"], row["sku"], row["captured_on"]))
    return rows


def write_history(rows: list[dict]) -> list[Path]:
    long_path = OUT / "neocloud_prices.csv"
    with long_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["company", "sku", "captured_on", "usd_per_gpu_hour", "basis", "source_url"])
        writer.writeheader()
        writer.writerows(rows)
    written = [long_path]
    by_company: dict[str, dict[str, dict[str, str]]] = {}
    for row in rows:
        by_company.setdefault(row["company"], {}).setdefault(row["sku"], {})[row["captured_on"]] = _clean(row["usd_per_gpu_hour"])
    for company, grid in by_company.items():
        columns = sorted({day for values in grid.values() for day in values})
        path = WIDE / f"neocloud_{company.lower()}.csv"
        _write_wide(path, "USD per GPU-hour, on-demand", columns, grid)
        written.append(path)
    return written



