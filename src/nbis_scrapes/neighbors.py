"""Data centers near each Nebius campus.

PeeringDB supplies the buildings: name, operator, and coordinates. Epoch AI
supplies a power estimate where a frontier campus falls inside the same radius.
The estimate is Epoch's, built from satellite imagery and permits. It is not a
utility meter and it is not a Nebius figure.

A neighbor is another facility within 60 km. The same pin, under 0.4 km, is
dropped. Each campus keeps the 15 nearest so a London listing stays readable,
and the summary counts how many facilities sit inside the radius.
"""

from __future__ import annotations

import csv
import json
import math
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CACHE = ROOT / "data" / "cache"
OUT = ROOT / "data" / "out"
LEDGER = ROOT / "data" / "ledger.json"
PEERINGDB = CACHE / "peeringdb_fac.json"
EPOCH = CACHE / "epoch_data_centers.csv"
GEOCODE = CACHE / "epoch_geocode.json"

RADIUS_KM = 60
SELF_KM = 0.4
KEEP = 15
USER_AGENT = "NBIS-Scrapes/0.1 (personal research; aaditprayag@users.noreply.github.com)"
EPOCH_COUNTRIES = {"United States", "United Kingdom", "Finland", "Iceland"}


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371.0
    p1, l1, p2, l2 = (math.radians(v) for v in (lat1, lon1, lat2, lon2))
    dlat, dlon = p2 - p1, l2 - l1
    h = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    return 2 * radius * math.asin(math.sqrt(h))


def satellite_url(lat: float, lon: float) -> str:
    return f"https://www.google.com/maps?q={lat:.5f},{lon:.5f}&t=k&z=17"


def _clean_owner(value: str) -> str:
    return value.split("#", 1)[0].strip()


def _mw_from_notes(notes: str) -> str:
    import re
    match = re.search(r"(\d+(?:\.\d+)?)\s*MW\b", notes or "", re.IGNORECASE)
    return match.group(1) if match else ""


def load_nebius() -> list[dict]:
    ledger = json.loads(LEDGER.read_text(encoding="utf-8"))
    sites = []
    for site in ledger["sites"]:
        if site["lat"] is None or site["lon"] is None or not site["confirmed"]:
            continue
        sites.append(site)
    return sites


def load_peeringdb() -> list[dict]:
    payload = json.loads(PEERINGDB.read_text(encoding="utf-8"))
    rows = []
    for fac in payload["data"]:
        if fac.get("latitude") in (None, "") or fac.get("longitude") in (None, ""):
            continue
        rows.append({
            "name": fac.get("name") or "",
            "owner": fac.get("org_name") or "",
            "city": fac.get("city") or "",
            "country": fac.get("country") or "",
            "lat": float(fac["latitude"]),
            "lon": float(fac["longitude"]),
            "website": fac.get("website") or "",
            "power_mw": _mw_from_notes(fac.get("notes") or ""),
            "power_basis": "mentioned in the PeeringDB note" if _mw_from_notes(fac.get("notes") or "") else "",
            "source": "PeeringDB",
        })
    return rows


def _geocode(query: str) -> tuple[float, float] | None:
    url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({
        "q": query,
        "format": "json",
        "limit": 1,
    })
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))
    if not payload:
        return None
    return float(payload[0]["lat"]), float(payload[0]["lon"])


def load_epoch(geocode: bool = True) -> list[dict]:
    if not EPOCH.exists():
        return []
    cache = {}
    if GEOCODE.exists():
        cache = json.loads(GEOCODE.read_text(encoding="utf-8"))
    rows = []
    dirty = False
    for record in csv.DictReader(EPOCH.open(encoding="utf-8")):
        if record["Country"] not in EPOCH_COUNTRIES:
            continue
        address = " ".join(record["Address"].split())
        query = address if address else f"{record['Name']}, {record['Country']}"
        if query not in cache and geocode:
            time.sleep(1.1)
            try:
                cache[query] = _geocode(query)
            except Exception:
                cache[query] = None
            dirty = True
        point = cache.get(query)
        if not point:
            continue
        power = record["Current power (MW)"].strip()
        rows.append({
            "name": record["Name"],
            "owner": _clean_owner(record["Owner"]),
            "city": "",
            "country": record["Country"],
            "lat": point[0],
            "lon": point[1],
            "website": "",
            "power_mw": power,
            "power_basis": "Epoch AI estimate, Oct 2026, from satellite imagery and permits",
            "source": "Epoch AI",
        })
    if dirty:
        CACHE.mkdir(parents=True, exist_ok=True)
        GEOCODE.write_text(json.dumps(cache), encoding="utf-8")
    return rows


def _same_building(left: dict, right: dict) -> bool:
    return haversine_km(left["lat"], left["lon"], right["lat"], right["lon"]) < 1.0 and (
        left["name"].lower()[:18] == right["name"].lower()[:18]
        or (left["owner"] and left["owner"].lower() == right["owner"].lower())
    )


def nearby(nebius: list[dict], facilities: list[dict]) -> tuple[list[dict], list[dict]]:
    detail = []
    summary = []
    for site in nebius:
        ranked = []
        for fac in facilities:
            distance = haversine_km(site["lat"], site["lon"], fac["lat"], fac["lon"])
            if distance < SELF_KM or distance > RADIUS_KM:
                continue
            label = f"{fac['name']} {fac['owner']}".lower()
            if "nebius" in label and distance < 25:
                continue
            ranked.append((distance, fac))
        ranked.sort(key=lambda item: item[0])
        # Prefer a record that has a power figure when two sources describe one building.
        chosen = []
        for distance, fac in ranked:
            if any(_same_building(fac, prior) for _, prior in chosen):
                continue
            chosen.append((distance, fac))
        summary.append({
            "nebius_site": site["name"],
            "country": site["country"],
            "facilities_within_60km": len(chosen),
            "nearest": chosen[0][1]["name"] if chosen else "",
            "nearest_km": round(chosen[0][0], 1) if chosen else "",
            "nearest_owner": chosen[0][1]["owner"] if chosen else "",
        })
        for distance, fac in chosen[:KEEP]:
            detail.append({
                "nebius_site": site["name"],
                "nebius_country": site["country"],
                "neighbor": fac["name"],
                "owner": fac["owner"],
                "city": fac["city"],
                "country": fac["country"],
                "distance_km": round(distance, 1),
                "power_mw": fac["power_mw"],
                "power_basis": fac["power_basis"],
                "latitude": round(fac["lat"], 5),
                "longitude": round(fac["lon"], 5),
                "website": fac["website"],
                "satellite_url": satellite_url(fac["lat"], fac["lon"]),
                "source": fac["source"],
            })
    return detail, summary


def to_geojson(detail: list[dict]) -> dict:
    features = []
    for row in detail:
        features.append({
            "type": "Feature",
            "geometry": {"type": "Point", "coordinates": [row["longitude"], row["latitude"]]},
            "properties": {
                "name": row["neighbor"],
                "owner": row["owner"],
                "near": row["nebius_site"],
                "distance_km": row["distance_km"],
                "power_mw": row["power_mw"],
                "power_basis": row["power_basis"],
                "satellite_url": row["satellite_url"],
                "source": row["source"],
            },
        })
    return {"type": "FeatureCollection", "features": features}


def _write(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def build(geocode: bool = True) -> tuple[Path, Path]:
    facilities = load_peeringdb() + load_epoch(geocode=geocode)
    detail, summary = nearby(load_nebius(), facilities)
    detail_path = OUT / "neighbors.csv"
    summary_path = OUT / "neighbors_summary.csv"
    _write(detail_path, detail)
    _write(summary_path, summary)
    (OUT / "neighbors.geojson").write_text(json.dumps(to_geojson(detail)), encoding="utf-8")
    return detail_path, summary_path
