"""Build data/ledger.json from the Point72 capacity workbook.

The workbook is the reviewed site panel (5 Oct 2026). This script does not
re-estimate megawatts. It copies the quarterly series and attaches geography,
ownership, and tags so later code can tell a company print from a model cell.

Usage:
  python scripts/seed_ledger.py "C:\\Users\\aadit\\OneDrive\\Nebius Pitch - Point72\\nebius_capacity.xlsx"
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = Path(
    r"C:\Users\aadit\OneDrive\Nebius Pitch - Point72\nebius_capacity.xlsx"
)

# Prefix of the site name in the Quarterly MW sheet. First match wins.
META = [
    ("Vineland NJ", {
        "country": "USA", "ownership": "build-to-suit", "customer": "Microsoft",
        "confirmed": True, "lat": 39.4864, "lon": -75.0257, "geo_precision": "city",
        "note": "FY2025 20-F: first build-to-suit, Vineland, NJ, signed February 2025. Nebius is the tenant. DataOne owns the shell. Microsoft is the customer. Energized rows are IT MW, genset-capped.",
    }),
    ("Independence MO", {
        "country": "USA", "ownership": "greenfield", "customer": "unassigned",
        "confirmed": True, "lat": 39.0956, "lon": -94.3553, "geo_precision": "city",
        "note": "FY2025 20-F: Missouri and Finland are the greenfield sites. Newly owned land, February 2026. 1.2 GW in the Q1 2026 letter. No energized IT through Q2 2026.",
    }),
    ("Butler Twp PA", {
        "country": "USA", "ownership": "greenfield", "customer": "unassigned",
        "confirmed": True, "lat": 40.7612, "lon": -76.2280, "geo_precision": "county",
        "note": "Owned site announced May 2026, after the 20-F. On the 12 Aug 2026 letter map as Pennsylvania. Up to 1.2 GW. Treated as greenfield because the company says it owns the land. Not in the 20-F site list.",
    }),
    ("Birmingham AL", {
        "country": "USA", "ownership": "greenfield", "customer": "unassigned",
        "confirmed": True, "lat": 33.4473, "lon": -86.8376, "geo_precision": "neighborhood",
        "note": "FY2025 20-F: newly owned Alabama site, and the filing groups Alabama with Missouri as a greenfield. The address is a former operations building, which also fits the company's brownfield definition. The filing's word is greenfield. No revenue IT in 1H 2026.",
    }),
    ("Bethune FR", {
        "country": "France", "ownership": "co-location", "customer": "Meta (press, estimated)",
        "confirmed": True, "lat": 50.5303, "lon": 2.6401, "geo_precision": "city",
        "note": "Azur, former Bridgestone plant. Contracted in Q1 2026. Energized base still 0 at Q2 2026.",
    }),
    ("Lappeenranta FI", {
        "country": "Finland", "ownership": "greenfield", "customer": "unassigned",
        "confirmed": True, "lat": 61.0587, "lon": 28.1881, "geo_precision": "city",
        "note": "Owned Finnish site, 310 MW, first customer capacity guided for 2027. Not named in the 20-F bullet list. The 12 Aug 2026 letter adds a Finland site. Classed greenfield because the company says it owns the land.",
    }),
    ("Mantsala FI #1", {
        "country": "Finland", "ownership": "greenfield", "customer": "self-serve",
        "confirmed": True, "lat": 60.6314, "lon": 25.3180, "geo_precision": "city",
        "note": "FY2025 20-F: Mäntsälä is a greenfield data center built to Nebius's design. Oct 2024 release said capacity would triple to 75 MW.",
    }),
    ("Mantsala FI #2", {
        "country": "Finland", "ownership": "greenfield", "customer": "unassigned",
        "confirmed": True, "lat": 60.6550, "lon": 25.2800, "geo_precision": "district",
        "note": "Kapuli, next to the Mäntsälä greenfield. The 20-F does not name this plot separately. Offset from site 1 so both pins are visible.",
    }),
    ("Kansas City MO", {
        "country": "USA", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 39.0997, "lon": -94.5786, "geo_precision": "city",
        "note": "Patmos colo. 5 MW class from early 2025, expandable toward 40 MW.",
    }),
    ("Paris (Equinix", {
        "country": "France", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 48.8400, "lon": 2.3890, "geo_precision": "city",
        "note": "Equinix PA10. Megawatts are a small third-party figure, not a Nebius disclosure.",
    }),
    ("Keflavik IS", {
        "country": "Iceland", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 63.9850, "lon": -22.6060, "geo_precision": "town",
        "note": "Verne, private region, live from Q1 2025.",
    }),
    ("London Longcross", {
        "country": "UK", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 51.3910, "lon": -0.5840, "geo_precision": "town",
        "note": "Ark LP01 at Longcross.",
    }),
    ("Harlow UK", {
        "country": "UK", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 51.7850, "lon": 0.1270, "geo_precision": "city",
        "note": "Kao Data KLON. Ten-year colo.",
    }),
    ("Newport UK", {
        "country": "UK", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 51.5842, "lon": -2.9977, "geo_precision": "city",
        "note": "Vantage CWL1, announced 13 Aug 2026. MW not disclosed on its own.",
        "announced_mw": None,
    }),
    ("UK 3rd", {
        "country": "UK", "ownership": "unknown", "customer": "unassigned",
        "confirmed": False, "lat": 51.5770, "lon": 0.1830, "geo_precision": "lead-only",
        "note": "Lead only: Green Mountain LON-East, Romford. Not a confirmed Nebius campus.",
    }),
    ("Modi'in IL", {
        "country": "Israel", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 31.8980, "lon": 35.0100, "geo_precision": "city",
        "note": "Mega Or. Energized rows are IT MW. Launched Oct 2025.",
    }),
    ("Masmiyya IL", {
        "country": "Israel", "ownership": "co-location", "customer": "unassigned",
        "confirmed": True, "lat": 31.7230, "lon": 34.8380, "geo_precision": "approximate",
        "note": "Mega DC. Pin is approximate.",
    }),
    ("Beit Shemesh IL", {
        "country": "Israel", "ownership": "co-location", "customer": "unassigned",
        "confirmed": True, "lat": 31.7456, "lon": 34.9881, "geo_precision": "city",
        "note": "Mega DC. Landlord completion slides point at 2027.",
    }),
    ("Madrid-Getafe", {
        "country": "Spain", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 40.3057, "lon": -3.7327, "geo_precision": "city",
        "note": "Merlin Edged. Phase 1 went live in Sep 2026, after the history window.",
    }),
    ("Tallinn EE", {
        "country": "Estonia", "ownership": "co-location", "customer": "unassigned",
        "confirmed": True, "lat": 59.4370, "lon": 24.7536, "geo_precision": "city",
        "note": "Greenergy Huuru. Contracted in Q2 2026.",
    }),
    ("Minnesota colo", {
        "country": "USA", "ownership": "co-location", "customer": "self-serve",
        "confirmed": True, "lat": 44.9239, "lon": -92.9594, "geo_precision": "city",
        "note": "FY2025 20-F: co-location in Minnesota, February 2026. The filing does not name the city. Woodbury is the docs region us-north1, not a campus disclosure.",
    }),
    ("Oklahoma colo", {
        "country": "USA", "ownership": "co-location", "customer": "unassigned",
        "confirmed": True, "lat": None, "lon": None, "geo_precision": "unlocated",
        "note": "FY2025 20-F: co-location in Oklahoma, contracts secured February 2026. The town is not stated, so it is not pinned. The site itself is company-reported.",
    }),
    ("France 2nd", {
        "country": "France", "ownership": "co-location", "customer": "unassigned",
        "confirmed": True, "lat": None, "lon": None, "geo_precision": "unlocated",
        "note": "FY2025 20-F: two new French co-location sites in February 2026, in addition to Paris. Towns are not named. Béthune is one landlord-identified town. This row is the other unnamed contract.",
    }),
    ("AIB CLT1", {
        "country": "USA", "ownership": "co-location", "customer": "unassigned",
        "confirmed": True, "lat": 34.8380, "lon": -81.7000, "geo_precision": "county",
        "note": "Not on the 12 Aug 2026 Nebius location map and not in the 20-F. AIB's own 8-K on 30 Sep 2026 says 50 MW IT at its South Carolina site. Co-location. Quarterly history stays empty through Q2 2026.",
        "announced_mw": 50,
        "announced_on": "2026-09-30",
    }),
]


def slug(name: str) -> str:
    out = []
    prev = False
    for ch in name.lower():
        if ch.isalnum():
            out.append(ch)
            prev = False
        elif not prev:
            out.append("-")
            prev = True
    return "".join(out).strip("-")[:48]


def match_meta(name: str) -> dict:
    for prefix, meta in META:
        if name.startswith(prefix):
            return meta
    raise SystemExit(f"no metadata for site: {name}")


def one(company: dict, prefix: str) -> list:
    hits = [k for k in company if k.startswith(prefix)]
    if len(hits) != 1:
        raise SystemExit(f"{prefix!r} matched {hits}")
    return company[hits[0]]


def num(value):
    if value is None or value == "":
        return None
    if isinstance(value, str):
        return value
    return round(float(value), 4)


def connected_tag(quarter: str) -> str:
    if quarter == "Q4 2025":
        return "target"
    if quarter == "Q4 2024":
        return "estimated"
    return "estimated"


def contracted_tag(quarter: str) -> str:
    return {
        "Q4 2025": "reported_floor",
        "Q1 2026": "reported_floor",
        "Q4 2026": "target",
    }.get(quarter, "estimated")


def main() -> None:
    import openpyxl

    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_XLSX
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    ws = wb["Quarterly MW"]
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, max_col=32, values_only=True)))
    quarters = [str(h) for h in headers[4:]]

    sites_raw = []
    current = None
    for row in ws.iter_rows(min_row=3, max_row=122, max_col=32, values_only=True):
        site, block = row[0], row[1]
        if not site or not block:
            continue
        if current is None or current["name"] != site:
            current = {"name": site, "basis": row[2]}
            sites_raw.append(current)
        current[str(block)] = [num(v) for v in row[4:]]

    company = {}
    for row in ws.iter_rows(min_row=124, max_row=148, max_col=32, values_only=True):
        if row[1]:
            company[str(row[1])] = [num(v) for v in row[4:]]
    wb.close()

    sites = []
    for raw in sites_raw:
        meta = match_meta(raw["name"])
        sites.append({
            "id": slug(raw["name"]),
            "name": raw["name"],
            "country": meta["country"],
            "ownership": meta["ownership"],
            "customer": meta["customer"],
            "confirmed": meta["confirmed"],
            "lat": meta["lat"],
            "lon": meta["lon"],
            "geo_precision": meta["geo_precision"],
            "basis_energized": raw.get("basis"),
            "note": meta["note"],
            "announced_mw": meta.get("announced_mw"),
            "announced_on": meta.get("announced_on"),
            "contracted_mw": raw.get("Contracted MW"),
            "energized_low": raw.get("Energized MW low"),
            "energized_base": raw.get("Energized MW base (=midpoint)"),
            "energized_high": raw.get("Energized MW high"),
        })

    def series(prefix: str) -> list[dict]:
        values = one(company, prefix)
        tagger = connected_tag if prefix.startswith("Connected") else None
        rows = []
        for quarter, value in zip(quarters, values):
            if value is None:
                continue
            item = {"quarter": quarter, "value": value}
            if tagger:
                item["tag"] = tagger(quarter)
            rows.append(item)
        return rows

    contracted_rows = []
    for quarter, value in zip(quarters, one(company, "Contracted power reported")):
        if value is None:
            continue
        contracted_rows.append({
            "quarter": quarter,
            "value": value,
            "tag": contracted_tag(quarter),
        })

    ledger = {
        "provenance": {
            "seeded_from": "nebius_capacity.xlsx sheet 'Quarterly MW'",
            "seeded_on": "2026-10-06",
            "workbook_as_of": "2026-10-05",
            "history_end": "Q2 2026",
            "rule": "Cells are copied from the reviewed workbook. Tags say whether a company number is a print, a floor, a target, or an estimate.",
        },
        "quarters": quarters,
        "history_end": "Q2 2026",
        "sites": sites,
        "company": {
            "identified_energized_low": one(company, "Identified sites - energized low"),
            "identified_energized_base": one(company, "Identified sites - energized base"),
            "identified_energized_high": one(company, "Identified sites - energized high"),
            "identified_contracted": one(company, "Identified sites - contracted"),
            "mgmt_power_row": one(company, "Mgmt reported"),
            "mgmt_power_row_warning": "This workbook row mixes a Finland nameplate, year-end 2025 active power, and the low end of the year-end 2026 connected guide. Do not read it as one metric.",
            "connected_low": series("Connected MW low"),
            "connected_base": series("Connected MW base"),
            "connected_high": series("Connected MW high"),
            "contracted_company": contracted_rows,
            "billing_base": one(company, "Billing MW (IT) - base"),
            "billing_bloom_slip": one(company, "Billing MW (IT) - Bloom"),
            "energized_base_it_normalized": one(company, "Energized MW base - IT-normalized"),
            "reported_active": [
                {
                    "quarter": "Q4 2025",
                    "value": 170,
                    "tag": "reported",
                    "source": "Q4 2025 shareholder letter, 12 Feb 2026, and FY2025 20-F.",
                    "url": "https://www.sec.gov/Archives/edgar/data/1513845/000110465926013946/tm266173d1_ex99-2.htm",
                }
            ],
        },
    }

    dest = ROOT / "data" / "ledger.json"
    dest.write_text(json.dumps(ledger, indent=2), encoding="utf-8")
    types = ROOT / "data" / "out" / "site_types.csv"
    types.parent.mkdir(parents=True, exist_ok=True)
    import csv
    with types.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["site", "company_type", "country", "located", "note"])
        writer.writeheader()
        for site in sites:
            writer.writerow({
                "site": site["name"],
                "company_type": site["ownership"],
                "country": site["country"],
                "located": site["lat"] is not None,
                "note": site["note"],
            })
    print(f"wrote {dest} sites={len(sites)} quarters={len(quarters)}")
    print(f"wrote {types}")


if __name__ == "__main__":
    main()
