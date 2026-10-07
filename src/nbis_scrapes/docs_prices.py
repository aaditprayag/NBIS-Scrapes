"""Nebius GPU prices from the official docs, not the marketing page.

The source is the markdown Nebius publishes for agents:

    https://docs.nebius.com/compute/resources/pricing.md
    https://docs.nebius.com/applications/standalone/pricing.md
    https://docs.nebius.com/changelog.md

The live compute page keeps one previous card. Earlier cards stay on older
copies of that same docs URL. Those copies are fetched only for this URL.
"""

from __future__ import annotations

import csv
import re
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
USER_AGENT = "NBIS-Scrapes/0.1 (personal research)"

COMPUTE_MD = "https://docs.nebius.com/compute/resources/pricing.md"
STANDALONE_MD = "https://docs.nebius.com/applications/standalone/pricing.md"
CHANGELOG_MD = "https://docs.nebius.com/changelog.md"

# Official docs page, captured when it still showed a card the live page has replaced.
DOCS_ARCHIVE_STAMPS = (
    "20260527185513",
    "20260715075624",
)

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
}

PRICE_RE = re.compile(r"\\?\$([0-9]+(?:\.[0-9]+)?)")
VINTAGE_RE = re.compile(
    r"(before|from)\s+([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})",
    re.IGNORECASE,
)
ROW_RE = re.compile(
    r"<tr>\s*<t[dh][^>]*>(.*?)</t[dh]>\s*<t[dh][^>]*>(.*?)</t[dh]>\s*<t[dh][^>]*>(.*?)</t[dh]>\s*</tr>",
    re.IGNORECASE | re.DOTALL,
)
HEADING_RE = re.compile(r"<h[1-6][^>]*\bid=\"(nvidia-[^\"]+)\"", re.IGNORECASE)
MD_HEADING_RE = re.compile(r"^#{2,4}\s+(.+)$")
REGION_RE = re.compile(r"`([a-z]{2}-[a-z]+\d)`")


def _strip_tags(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", value)).strip()


def _vintage(label: str) -> tuple[str, str]:
    match = VINTAGE_RE.search(label)
    if not match:
        return "undated", ""
    month = MONTHS.get(match.group(2).lower())
    if month is None:
        return match.group(1).lower(), ""
    return match.group(1).lower(), f"{int(match.group(4)):04d}-{month:02d}-{int(match.group(3)):02d}"


def _price(cell: str) -> float | None:
    match = PRICE_RE.search(cell)
    if not match:
        return None
    return float(match.group(1))


def _row(source_url: str, captured_on: str, platform_id: str, regions: str, item: str, vintage: str, vintage_date: str, price: float) -> dict:
    return {
        "source_url": source_url,
        "captured_on": captured_on,
        "platform_id": platform_id,
        "regions": regions,
        "item": item,
        "preemptible": item.lower().startswith("preemptible"),
        "vintage": vintage,
        "vintage_date": vintage_date,
        "usd_per_gpu_hour": price,
    }


def parse_markdown(text: str, source_url: str, captured_on: str) -> list[dict]:
    platform = ""
    regions: list[str] = []
    vintage, vintage_date = "undated", ""
    currency = ""
    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if line.startswith("<Tab") and "USD" in line:
            currency = "USD"
            continue
        if line.startswith("<Tab") and "ILS" in line:
            currency = "ILS"
            continue
        if line.startswith("</Tab>") or line.startswith("</Tabs>"):
            currency = ""
            continue
        heading = MD_HEADING_RE.match(line)
        if heading and "NVIDIA" in heading.group(1):
            platform = heading.group(1).split(",")[-1].strip().strip("`")
            if platform.startswith("NVIDIA"):
                platform = ""
            regions = []
            vintage, vintage_date = "undated", ""
            continue
        if "available in" in line and "`" in line:
            regions = REGION_RE.findall(line)
        if not line.startswith("|"):
            continue
        if currency == "ILS":
            continue
        cells = [cell.strip().strip("*") for cell in line.strip("|").split("|")]
        if len(cells) < 3:
            continue
        item, price_cell, unit = cells[0], cells[1], cells[2]
        if item.lower().startswith("item"):
            vintage, vintage_date = _vintage(item)
            continue
        if "GPU hour" not in unit:
            continue
        price = _price(price_cell)
        if price is None:
            continue
        rows.append(_row(
            source_url, captured_on, platform, ",".join(regions),
            item, vintage, vintage_date, price,
        ))
    return rows


def parse_docs_html(text: str, source_url: str, captured_on: str) -> list[dict]:
    platform = ""
    vintage, vintage_date = "undated", ""
    rows = []
    events = []
    for match in HEADING_RE.finditer(text):
        events.append((match.start(), "heading", match.group(1)))
    for match in ROW_RE.finditer(text):
        events.append((match.start(), "row", match))
    for _, kind, payload in sorted(events, key=lambda item: item[0]):
        if kind == "heading":
            platform = payload
            vintage, vintage_date = "undated", ""
            continue
        cells = [_strip_tags(cell) for cell in payload.groups()]
        item = cells[0]
        if item.lower().startswith("item"):
            vintage, vintage_date = _vintage(item)
            continue
        if len(cells) < 3 or "GPU hour" not in cells[2]:
            continue
        price = _price(cells[1])
        if price is None:
            continue
        rows.append(_row(source_url, captured_on, platform, "", item, vintage, vintage_date, price))
    deduped = []
    seen = set()
    for row in rows:
        key = (
            row["platform_id"], row["item"], row["vintage"], row["vintage_date"],
            row["usd_per_gpu_hour"], row["captured_on"],
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)
    return deduped


def parse_changelog(text: str) -> list[dict]:
    label = ""
    events = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        label_match = re.match(r"<Update label=\"(.+)\">", line)
        if label_match:
            label = label_match.group(1)
            continue
        if "price" not in line.lower():
            continue
        if not any(token in line for token in ("GPU", "Compute", "Standalone", "Spot pricing")):
            continue
        events.append({
            "changelog_label": label,
            "bullet": _strip_tags(line).lstrip("* ").strip(),
            "source_url": CHANGELOG_MD,
        })
    return events


def _get(url: str, timeout: int = 60) -> str:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept-Encoding": "identity"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        raw = response.read()
    if raw[:2] == b"\x1f\x8b":
        import gzip
        raw = gzip.decompress(raw)
    return raw.decode("utf-8", errors="replace")


def fetch_official(today: str | None = None) -> tuple[list[dict], list[dict]]:
    captured = today or datetime.now(timezone.utc).strftime("%Y-%m-%d")
    prices = []
    for url in (COMPUTE_MD, STANDALONE_MD):
        prices.extend(parse_markdown(_get(url), url, captured))
    for stamp in DOCS_ARCHIVE_STAMPS:
        page = f"https://docs.nebius.com/compute/resources/pricing"
        archived = f"https://web.archive.org/web/{stamp}id_/{page}"
        captured_on = datetime.strptime(stamp, "%Y%m%d%H%M%S").strftime("%Y-%m-%d")
        try:
            html = _get(archived, timeout=90)
        except Exception:
            continue
        prices.extend(parse_docs_html(html, page, captured_on))
    events = parse_changelog(_get(CHANGELOG_MD))
    return prices, events


def _write(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def write_outputs(prices: list[dict], events: list[dict], out_dir: Path | None = None) -> Path:
    out_dir = out_dir or (ROOT / "data" / "out")
    _write(out_dir / "docs_gpu_prices.csv", prices)
    _write(out_dir / "docs_price_events.csv", events)
    return out_dir
