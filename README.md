# NBIS Scrapes

Private research on Nebius (NBIS) campus power and short-term GPU demand. Public sources only. This is not a price target.

Nebius prints three different power stocks. **Contracted** power is land plus a power commitment. **Connected** power is a fully built and equipped hall, which can sit empty. **Active** power is IT equipment that is on and can bill. The May 2026 shareholder letter is the definition. The company has published one active-power actual in this window, about **170 MW at 31 Dec 2025**, and has published connected power as targets rather than quarter-end actuals.

## Quarterly panel

History is Q1 2024 through Q2 2026. Later quarters are the forward path from the 5 Oct 2026 site ledger. Energized base is that ledger's estimate of live power. It is the stand-in for active power in the quarters Nebius left blank.

| Quarter | Energized base (est.) | Reported active | Connected | Contracted |
|---|---:|---:|---|---|
| Q1–Q3 2024 | 25 | — | — | — |
| Q4 2024 | 30 | — | 25 estimated nameplate | — |
| Q1–Q2 2025 | 50 | — | — | — |
| Q3 2025 | 62.5 | — | — | — |
| Q4 2025 | 147.7 | **170** | **220 target** | >2 GW floor |
| Q1 2026 | 181.5 | — | — | >3.5 GW floor |
| Q2 2026 | 193 | — | — | — |
| Q4 2026 path | 319.5 | — | 941 estimated, guide is 800–1,000 | 5 GW target |

Active over connected, only where both sides are named:

- Year-end 2025: 170 / 220 = **77.3%**. The 220 is the connected target. Nebius never confirmed it was met. They did beat the 100 MW active target.
- Year-end 2026 billing path of about 258 MW over the 800–1,000 MW connected guide is **32% to 26%**. Over the ledger's 941 MW connected reconstruction it is about **27%**. Billing is modeled about a quarter behind energization. Connected running ahead of billing is the build pipeline, not a measured idle-GPU rate.
- From Q4 2026 to Q4 2027 the ledger's energized base rises by about 893 MW, against a company pace of more than 1 GW a year brought to market. That pace is closer to connected power than to IT that is already on.

Full table, site rows, and the ratio file are in `data/out/`. The map is `data/out/map.html`. Open it in a browser. Circle area is contracted megawatts. Green means the model has something pulling in that quarter. Amber means the capacity is booked and energized base is zero. Oklahoma and the unnamed second French colo are not pinned, because the town is not public. Pins are city-level.

## Short-term demand

The filing clock is quarterly. A stockout is a weekly fact. The scrapers worth running, and the ones that are a waste of time, are written up in `docs/demand-proxies.md`.

What actually refreshes:

1. `https://nebius.com/prices` — list price, the published 1 Oct 2026 column, and the public spot floor. NVL72 racks are contact-sales. The page has no history of its own. Fifty archived copies, back to 15 Nov 2024, are in `data/out/gpu_price_history.csv` and written up in `docs/gpu-price-history.md`. On-demand H100 was $2.95 and H200 was $3.50 until at least 7 May 2026. By 12 Jul 2026 those were $3.85 and $4.50. The $4.50 / $5.40 / $8.50 / $9.50 October column shows up on 23 Sep and the first column had not moved to it by 6 Oct.
2. `https://docs.nebius.com/compute/virtual-machines/types` — which GPU platform is still sold in which region. H100 is only `eu-north1`. B300 is in `uk-south1`, `eu-west2`, and `us-north1`. A region id is not a building. Inferences are in `data/region_crosswalk.json` and on the map, kept off the pins.
3. OpenRouter's Nebius provider page — how many models Token Factory is serving. That is the short inference book, not Vineland.
4. Vast.ai `GET /api/v0/metrics/gpu/current/` — market supply, demand, and percentile price, about every five minutes. Needs a Vast **host** key in `VAST_API_KEY`. Client keys are rejected.

## Run

```
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
python -m nbis_scrapes all
```

`panel` and `map` stay offline. `demand` hits the network and writes `data/snapshots/`. Set `VAST_API_KEY` first if you want the market thermometer.

Rebuild the ledger from the capacity workbook, if that file is on this machine:

```
python scripts/seed_ledger.py
```

That step needs openpyxl. Everything else is the standard library.

Design notes: `docs/superpowers/specs/2026-10-06-nbis-power-alt-data-design.md`.
