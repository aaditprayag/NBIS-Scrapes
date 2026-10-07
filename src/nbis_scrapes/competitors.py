"""Other AI, cloud, and colocation capacity within 200 km of each Nebius campus.

`neighbors.py` lists buildings within 60 km. This file is about megawatts and
timing. Each record is one phase of a project: a building, a hall, or a power
block, with its own megawatts and the quarter it comes online. Every figure
carries a source URL and a short quote from that page.

Ring 1 is within 50 km of the campus pin. Ring 2 is 50 to 200 km. Distance is
a straight line. The round trip in milliseconds is the distance divided by
100, which is light in fiber there and back with no detours. Real routes are
longer, so read it as a floor.

A phase counts toward the quarterly totals only when its megawatts and its
online quarter are both published. A phase that was running before 2024 counts
from Q1 2024. A year with no quarter is Q4 of that year.
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path

from nbis_scrapes.neighbors import haversine_km
from nbis_scrapes.wide_prices import _clean, _write_wide, quarter_end

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "out"
WIDE = OUT / "wide"
LEDGER = ROOT / "data" / "ledger.json"
PROJECTS = ROOT / "data" / "competitor_projects.jsonl"

RING_1_KM = 50
RING_2_KM = 200
FIRST_YEAR = 2024
LAST_YEAR = 2030

# Research key -> ledger site name. Oklahoma City has no pin in the ledger, and
# Hyderabad is not in it, so those two keep the research pin.
CAMPUSES = {
    "Mantsala": "Mantsala FI #1",
    "Lappeenranta": "Lappeenranta FI (Pajarila / Polarnode)",
    "Independence": "Independence MO (Eastgate)",
    "Highridge": "Butler Twp PA (Highridge)",
    "Birmingham": "Birmingham AL (BHM01, Oxmoor)",
    "Vineland": "Vineland NJ (DataOne)",
    "Kansas City": "Kansas City MO (Patmos)",
    "Oklahoma City": None,
    "Spartanburg": "AIB CLT1 (Moore, Spartanburg County SC; Lockhart Power)",
    "Woodbury": "Minnesota colo (us-north1, Woodbury)",
    "Keflavik": "Keflavik IS (Verne)",
    "Paris": "Paris (Equinix PA10)",
    "Bethune": "Bethune FR (Azur Datacenter, ex-Bridgestone)",
    "Getafe": "Madrid-Getafe ES (Merlin Edged)",
    "Huuru": "Tallinn EE (Greenergy Huuru)",
    "Longcross": "London Longcross (Ark LP01)",
    "Harlow": "Harlow UK (Kao Data KLON-02/03)",
    "Newport": "Newport UK (Vantage CWL1)",
    "Modiin": "Modi'in IL (Mega Or)",
    "Masmiyya": "Masmiyya IL (Mega DC)",
    "Beit Shemesh": "Beit Shemesh IL (Mega DC)",
    "Hyderabad": None,
}

# Address-level research pins, used only where the ledger has none.
RESEARCH_PINS = {
    "Oklahoma City": ("USA", 35.4729886, -97.5170536, "city centre from OpenStreetMap Nominatim; the site address is not public"),
    "Hyderabad": ("India", 17.360589, 78.4740613, "city centre from OpenStreetMap Nominatim; the site is not public"),
}

PHASE_FIELDS = [
    "project_id",
    "project",
    "operator",
    "owner",
    "tenant",
    "type",
    "city",
    "county",
    "country",
    "latitude",
    "longitude",
    "address",
    "coord_basis",
    "coord_url",
    "phase",
    "mw",
    "mw_basis",
    "online_quarter",
    "online_basis",
    "year_only",
    "status",
    "mw_source_url",
    "mw_quote",
    "online_source_url",
    "online_quote",
    "other_sources",
    "notes",
]

SITE_FIELDS = [
    "nebius_site",
    "nebius_country",
    "nebius_pin",
    "project_id",
    "project",
    "operator",
    "type",
    "country",
    "phase",
    "distance_km",
    "ring",
    "rtt_ms",
    "mw",
    "mw_basis",
    "online_quarter",
    "status",
    "counted",
]

GENERIC = {"data", "centre", "center", "centers", "centres", "campus", "the", "and", "park", "phase", "building", "ltd", "limited", "inc"}


def quarters() -> list[str]:
    return [f"Q{quarter} {year}" for year in range(FIRST_YEAR, LAST_YEAR + 1) for quarter in range(1, 5)]


def quarter_index(label) -> int | None:
    """0 is Q1 2024. Anything before 2024 is 0. None means no usable date."""
    if label is None:
        return None
    text = str(label).strip()
    if text.lower().startswith("pre"):
        return 0
    match = re.fullmatch(r"Q([1-4])-(\d{4})", text)
    if not match:
        return None
    quarter, year = int(match.group(1)), int(match.group(2))
    if year < FIRST_YEAR:
        return 0
    return (year - FIRST_YEAR) * 4 + quarter - 1


def ring(distance_km: float) -> int | None:
    if distance_km <= RING_1_KM:
        return 1
    if distance_km <= RING_2_KM:
        return 2
    return None


def rtt_ms(distance_km: float) -> float:
    return round(distance_km / 100.0, 2)


def _tokens(project: dict) -> set[str]:
    text = f"{project.get('operator') or ''} {project.get('project') or ''}".lower()
    return set(re.findall(r"[a-z0-9]{3,}", text)) - GENERIC


def _signature(project: dict) -> list[tuple[str, str]]:
    return sorted((str(phase.get("mw")), str(phase.get("online_quarter"))) for phase in project.get("phases") or [])


def merge_projects(slices: list[list[dict]]) -> tuple[list[dict], list[dict]]:
    """Join research slices. Returns the kept projects and the dropped ones.

    A repeated project id keeps the first slice's copy. A project in another
    slice within 1 km, sharing a name or operator word, whose phases are the
    same megawatts and quarters (or a subset), is the same project twice.
    """
    kept: list[dict] = []
    dropped: list[dict] = []
    seen: set[str] = set()
    for index, projects in enumerate(slices):
        for project in projects:
            pid = project.get("project_id")
            if pid in seen:
                dropped.append({"project_id": pid, "reason": "repeated project id"})
                continue
            seen.add(pid)
            twin = None
            try:
                lat, lon = float(project["lat"]), float(project["lon"])
            except (KeyError, TypeError, ValueError):
                lat = lon = None
            if lat is not None:
                for prior, prior_index in kept:
                    if prior_index == index:
                        continue
                    try:
                        prior_lat, prior_lon = float(prior["lat"]), float(prior["lon"])
                    except (KeyError, TypeError, ValueError):
                        continue
                    if haversine_km(lat, lon, prior_lat, prior_lon) > 1.0 or not (_tokens(project) & _tokens(prior)):
                        continue
                    mine, theirs = _signature(project), _signature(prior)
                    if mine == theirs or set(mine) <= set(theirs):
                        twin = prior
                        break
            if twin is not None:
                dropped.append({"project_id": pid, "reason": f"same project as {twin['project_id']}"})
                continue
            kept.append((project, index))
    return [project for project, _index in kept], dropped


def read_slice(path: Path) -> list[dict]:
    if path.suffix == ".jsonl":
        return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    return json.loads(path.read_text(encoding="utf-8")).get("projects", [])


def load_projects(path: Path | None = None) -> list[dict]:
    path = path or PROJECTS
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def write_projects(projects: list[dict], path: Path | None = None) -> Path:
    path = path or PROJECTS
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        for project in projects:
            handle.write(json.dumps(project, ensure_ascii=False) + "\n")
    return path


def campuses() -> list[dict]:
    ledger = json.loads(LEDGER.read_text(encoding="utf-8"))
    by_name = {site["name"]: site for site in ledger["sites"]}
    found = []
    for key, name in CAMPUSES.items():
        site = by_name.get(name) if name else None
        if site and site["lat"] is not None:
            found.append({"key": key, "name": site["name"], "country": site["country"], "lat": site["lat"],
                          "lon": site["lon"], "pin": f"data/ledger.json, {site.get('geo_precision') or 'city'} level"})
        else:
            country, lat, lon, basis = RESEARCH_PINS[key]
            found.append({"key": key, "name": key, "country": country, "lat": lat, "lon": lon, "pin": basis})
    return found


def _pick(sources: list[dict], field: str) -> dict | None:
    return next((source for source in sources or [] if field in (source.get("field") or "")), None)


def _text(value, missing: str) -> str:
    if value is None or str(value).strip() == "":
        return missing
    return str(value)


def why_not_counted(phase: dict) -> str:
    mw = phase.get("mw")
    index = quarter_index(phase.get("online_quarter"))
    if mw is None:
        return "MW not public"
    try:
        if float(mw) <= 0:
            return "MW not positive"
    except (TypeError, ValueError):
        return "MW not a number"
    if index is None:
        return "no online date public"
    if index >= len(quarters()):
        return f"online after Q4 {LAST_YEAR}"
    return ""


def phase_rows(projects: list[dict]) -> list[dict]:
    rows = []
    for project in projects:
        for phase in project.get("phases") or []:
            sources = phase.get("sources") or []
            mw_source = _pick(sources, "mw") or (sources[0] if sources else {})
            online_source = _pick(sources, "online") or _pick(sources, "status") or mw_source
            others = [source for source in sources if source is not mw_source and source is not online_source]
            rows.append({
                "project_id": project["project_id"],
                "project": project.get("project") or "",
                "operator": _text(project.get("operator"), "not disclosed"),
                "owner": _text(project.get("owner"), "not disclosed"),
                "tenant": _text(project.get("tenant"), "not disclosed"),
                "type": _text(project.get("type"), "not stated"),
                "city": _text(project.get("city"), "not stated"),
                "county": _text(project.get("county"), "not stated"),
                "country": _text(project.get("country"), "not stated"),
                "latitude": _text(project.get("lat"), "not public"),
                "longitude": _text(project.get("lon"), "not public"),
                "address": _text(project.get("address"), "not published"),
                "coord_basis": _text(project.get("coord_basis"), "not stated"),
                "coord_url": _text(project.get("coord_url"), "none"),
                "phase": _text(phase.get("phase"), "whole site"),
                "mw": _clean(phase["mw"]) if isinstance(phase.get("mw"), (int, float)) else _text(phase.get("mw"), "not public"),
                "mw_basis": _text(phase.get("mw_basis"), "not stated"),
                "online_quarter": _text(phase.get("online_quarter"), "not public"),
                "online_basis": _text(phase.get("online_basis"), "not stated"),
                "year_only": "yes" if phase.get("year_only") else "no",
                "status": _text(phase.get("status"), "not stated"),
                "mw_source_url": _text(mw_source.get("url"), "none found"),
                "mw_quote": _text(mw_source.get("quote"), "none found"),
                "online_source_url": _text((online_source or {}).get("url"), "none found"),
                "online_quote": _text((online_source or {}).get("quote"), "none found"),
                "other_sources": " | ".join(f"{source.get('field')}: {source.get('url')}" for source in others) or "none",
                "notes": _text(phase.get("notes"), "none"),
            })
    return rows


def site_rows(projects: list[dict], sites: list[dict]) -> list[dict]:
    """Every project phase within 200 km of every campus."""
    rows = []
    for site in sites:
        for project in projects:
            try:
                lat, lon = float(project["lat"]), float(project["lon"])
            except (KeyError, TypeError, ValueError):
                continue
            distance = haversine_km(site["lat"], site["lon"], lat, lon)
            band = ring(distance)
            if band is None:
                continue
            for phase in project.get("phases") or []:
                reason = why_not_counted(phase)
                mw = phase.get("mw")
                rows.append({
                    "nebius_site": site["name"],
                    "nebius_country": site["country"],
                    "nebius_pin": site["pin"],
                    "project_id": project["project_id"],
                    "project": project.get("project") or "",
                    "operator": _text(project.get("operator"), "not disclosed"),
                    "type": _text(project.get("type"), "not stated"),
                    "country": _text(project.get("country"), "not stated"),
                    "phase": _text(phase.get("phase"), "whole site"),
                    "distance_km": round(distance, 1),
                    "ring": band,
                    "rtt_ms": rtt_ms(distance),
                    "mw": _clean(mw) if isinstance(mw, (int, float)) else _text(mw, "not public"),
                    "mw_basis": _text(phase.get("mw_basis"), "not stated"),
                    "online_quarter": _text(phase.get("online_quarter"), "not public"),
                    "status": _text(phase.get("status"), "not stated"),
                    "counted": reason or "yes",
                })
    rows.sort(key=lambda row: (row["nebius_site"], row["ring"], row["distance_km"], row["project_id"], row["phase"]))
    return rows


def quarterly_mw(rows: list[dict], sites: list[dict], band: int) -> tuple[list[str], dict[str, dict[str, str]]]:
    """Megawatts online by each quarter end, per campus, inside one ring."""
    labels = quarters()
    columns = [quarter_end(label) for label in labels]
    grid: dict[str, dict[str, str]] = {}
    for site in sites:
        totals = [0.0] * len(labels)
        for row in rows:
            if row["nebius_site"] != site["name"] or row["ring"] != band or row["counted"] != "yes":
                continue
            start = quarter_index(row["online_quarter"])
            for index in range(start, len(labels)):
                totals[index] += float(row["mw"])
        grid[site["name"]] = {day: _clean(round(total, 1)) or "0" for day, total in zip(columns, totals)}
    return columns, grid


def build(projects: list[dict] | None = None, out_dir: Path | None = None) -> list[Path]:
    out_dir = out_dir or OUT
    projects = projects if projects is not None else load_projects()
    sites = campuses()
    phases = phase_rows(projects)
    rows = site_rows(projects, sites)
    written = []
    for name, fields, data in (("competitor_phases.csv", PHASE_FIELDS, phases), ("competitors.csv", SITE_FIELDS, rows)):
        path = out_dir / name
        _write(path, data, fields)
        written.append(path)
    for band, header in ((1, "MW online within 50 km"), (2, "MW online 50 to 200 km")):
        columns, grid = quarterly_mw(rows, sites, band)
        path = out_dir / "wide" / f"competitors_ring{band}.csv"
        _write_wide(path, header, columns, grid)
        written.append(path)
    return written


def _write(path: Path, rows: list[dict], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
