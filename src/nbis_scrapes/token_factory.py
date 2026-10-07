"""Token Factory list prices from Nebius's own catalog.

The JSON at https://tokenfactory.nebius.com/api/public/models_info is the
source the catalog itself calls authoritative. Older days come from copies of
https://tokenfactory.nebius.com/model-catalog.md, which is the same list.
"""

from __future__ import annotations

import csv
import json
import re
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
USER_AGENT = "NBIS-Scrapes/0.1 (personal research)"
MODELS_URL = "https://tokenfactory.nebius.com/api/public/models_info"
CATALOG_URL = "https://tokenfactory.nebius.com/model-catalog.md"

# Unique archived catalogs. The JSON endpoint has no public history of its own.
CATALOG_STAMPS = (
    "20260719234609",
    "20260806034606",
    "20260909111839",
    "20260911125540",
    "20260915184212",
)

FIELD_RE = re.compile(r"([a-z0-9_]+)=([0-9]+(?:\.[0-9]+)?)")


def _num(value):
    if value is None or value == "":
        return None
    return float(value)


def _regions(flavor: dict) -> str:
    names = []
    for region in flavor.get("regions") or []:
        if isinstance(region, dict) and region.get("name"):
            names.append(region["name"])
        elif isinstance(region, str):
            names.append(region)
    return ",".join(names)


def parse_models_info(payload: list, captured_on: str, source_url: str) -> list[dict]:
    rows = []
    for model in payload:
        for flavor in model.get("flavors") or []:
            rows.append({
                "captured_on": captured_on,
                "source_url": source_url,
                "model_id": flavor.get("model_id") or "",
                "model_name": flavor.get("model_name") or model.get("name") or "",
                "vendor": model.get("vendor") or "",
                "model_type": flavor.get("model_type") or model.get("type") or "",
                "regions": _regions(flavor),
                "quantization": flavor.get("quantization") or "",
                "label": flavor.get("label") or "",
                "input_usd_per_million": _num(flavor.get("input_price_per_million_tokens")),
                "cached_input_usd_per_million": _num(flavor.get("input_cached_price_per_million_tokens")),
                "output_usd_per_million": _num(flavor.get("output_price_per_million_tokens")),
                "usd_per_image": _num(flavor.get("price_per_image")),
            })
    return rows


def parse_catalog(text: str, captured_on: str, source_url: str) -> list[dict]:
    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line.startswith("|") or line.startswith("| Model") or line.startswith("| ---"):
            continue
        cells = [cell.strip() for cell in line.strip("|").split("|")]
        if len(cells) < 7:
            continue
        pricing = cells[6].replace("\\_", "_")
        if ":" not in pricing:
            continue
        model_id, rest = pricing.split(":", 1)
        fields = dict(FIELD_RE.findall(rest))
        rows.append({
            "captured_on": captured_on,
            "source_url": source_url,
            "model_id": model_id.strip(),
            "model_name": cells[0],
            "vendor": cells[1],
            "model_type": cells[2],
            "regions": cells[5],
            "quantization": "",
            "label": "",
            "input_usd_per_million": _num(fields.get("input_price_per_million_tokens")),
            "cached_input_usd_per_million": _num(fields.get("input_cached_price_per_million_tokens")),
            "output_usd_per_million": _num(fields.get("output_price_per_million_tokens")),
            "usd_per_image": _num(fields.get("price_per_image")),
        })
    return rows


def price_changes(rows: list[dict]) -> list[dict]:
    series: dict[str, list[dict]] = {}
    for row in rows:
        if not row["model_id"]:
            continue
        series.setdefault(row["model_id"], []).append(row)
    changes = []
    for model_id, points in series.items():
        points.sort(key=lambda row: row["captured_on"])
        previous = None
        for point in points:
            signature = (point["input_usd_per_million"], point["output_usd_per_million"], point["usd_per_image"])
            if previous is None:
                previous = point
                continue
            prev_sig = (
                previous["input_usd_per_million"],
                previous["output_usd_per_million"],
                previous["usd_per_image"],
            )
            if signature != prev_sig:
                changes.append({
                    "model_id": model_id,
                    "captured_on": point["captured_on"],
                    "previous_captured_on": previous["captured_on"],
                    "input_usd_per_million": point["input_usd_per_million"],
                    "previous_input_usd_per_million": previous["input_usd_per_million"],
                    "output_usd_per_million": point["output_usd_per_million"],
                    "previous_output_usd_per_million": previous["output_usd_per_million"],
                    "usd_per_image": point["usd_per_image"],
                    "previous_usd_per_image": previous["usd_per_image"],
                    "source_url": point["source_url"],
                })
            previous = point
    changes.sort(key=lambda row: (row["captured_on"], row["model_id"]))
    return changes


def _get(url: str, timeout: int = 60) -> str:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": USER_AGENT, "Accept-Encoding": "identity"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        raw = response.read()
    if raw[:2] == b"\x1f\x8b":
        import gzip
        raw = gzip.decompress(raw)
    return raw.decode("utf-8", errors="replace")


def _stamp_day(stamp: str) -> str:
    return datetime.strptime(stamp, "%Y%m%d%H%M%S").strftime("%Y-%m-%d")


def _archived_models(stamp: str) -> list[dict] | None:
    url = f"https://web.archive.org/web/{stamp}id_/{MODELS_URL}"
    try:
        payload = json.loads(_get(url, timeout=90))
    except Exception:
        return None
    if not isinstance(payload, list) or not payload:
        return None
    if not isinstance(payload[0], dict) or "flavors" not in payload[0]:
        return None
    return parse_models_info(payload, _stamp_day(stamp), MODELS_URL)


def fetch_history(today: str | None = None) -> list[dict]:
    captured = today or datetime.now(timezone.utc).strftime("%Y-%m-%d")
    rows = []
    for stamp in CATALOG_STAMPS:
        from_json = _archived_models(stamp)
        if from_json:
            rows.extend(from_json)
            continue
        archived = f"https://web.archive.org/web/{stamp}id_/{CATALOG_URL}"
        try:
            text = _get(archived)
        except Exception:
            continue
        rows.extend(parse_catalog(text, _stamp_day(stamp), CATALOG_URL))
    live = json.loads(_get(MODELS_URL))
    rows.extend(parse_models_info(live, captured, MODELS_URL))
    rows.sort(key=lambda row: (row["captured_on"], row["model_id"]))
    return rows


CHANGE_FIELDS = [
    "model_id",
    "captured_on",
    "previous_captured_on",
    "input_usd_per_million",
    "previous_input_usd_per_million",
    "output_usd_per_million",
    "previous_output_usd_per_million",
    "usd_per_image",
    "previous_usd_per_image",
    "source_url",
]


def write_outputs(rows: list[dict], out_dir: Path | None = None) -> Path:
    out_dir = out_dir or (ROOT / "data" / "out")
    out_dir.mkdir(parents=True, exist_ok=True)
    prices = out_dir / "token_factory_prices.csv"
    changes = out_dir / "token_factory_price_changes.csv"
    _write(prices, rows)
    _write(changes, price_changes(rows), CHANGE_FIELDS)
    return out_dir


def _write(path: Path, rows: list[dict], fieldnames: list[str] | None = None) -> None:
    fields = fieldnames or (list(rows[0].keys()) if rows else [])
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
