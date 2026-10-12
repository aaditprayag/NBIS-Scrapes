# Nebius power alt data

Date: 2026-10-06. Status: approved, then extended with a demand proxy and a campus map before the repo was created.

## Question

What is Nebius active power, connected power, and contracted power, by quarter from Q1 2024 through Q2 2026, where are the campuses, what share of connected power is active, and how does that path run forward? Second, which public pages are a proxy for short-term compute demand? Third, draw the campuses and show how much power each one is modeled as pulling.

## Definitions

From the Q1 2026 shareholder letter (6-K, 13 May 2026):

- Contracted power is land plus a power commitment.
- Connected power is power connected into a fully built and equipped data center. The hall can sit idle until the GPUs go in.
- Active power is power consumed by installed, operating IT equipment that can generate revenue.

The Q4 2025 letter used a looser phrase for connected power ("power connected into data centers"). The May 2026 letter is the one the panel uses.

## What is a company print

The only reported active-power stock in this window is about 170 MW at 31 Dec 2025 (Q4 2025 letter, 12 Feb 2026, and the FY2025 20-F).

Connected power is published as targets: 220 MW by year-end 2025, and 800 MW to 1 GW by year-end 2026. A year-end connected actual has not been printed. Q1 2026 and Q2 2026 have no active actual and no connected actual.

Contracted power is a floor or a target, not a precise quarter-end stock: more than 2 GW by early 2026, more than 3.5 GW in the 13 May 2026 letter, year-end 2026 target raised to 5 GW on 12 Aug 2026. From 2027 the company says it will bring more than 1 GW a year to market.

Q1 2024 through most of 2025 has no three-way disclosure. The Finland flagship is the starting stock. The 8 Oct 2024 release said Nebius would triple Mäntsälä to 75 MW, which implies about 25 MW before the expansion. That 25 MW is derived, and it is nameplate, not the later "active power" definition.

## Panel

`data/ledger.json` is the site ledger, seeded from the 5 Oct 2026 capacity workbook and tagged. `src/nbis_scrapes/panel.py` writes:

- `data/out/quarterly_company.csv` — one row per quarter, history through Q2 2026, forward path after that.
- `data/out/quarterly_sites.csv` — site by quarter.
- `data/out/utilization.csv` — only the ratios that have a named numerator and denominator.
- `data/out/SUMMARY.md` — the same panel in markdown.

Rules:

- `reported_active_mw` is filled only for Q4 2025. The workbook row that mixes the Finland nameplate, the 170 MW print, and the 800 MW guide is stored aside and is not copied into that column.
- Connected cells carry a tag: `target` for the 220 MW year-end 2025 figure, `estimated` for the model path including the year-end 2026 reconstruction of the 800 MW–1 GW guide.
- Contracted company cells carry `reported_floor` or `target`.
- Utilization at year-end 2025 is reported active divided by the connected target. It is not an actual utilization.
- Year-end 2026 utilization uses model billing MW over 800, over 1,000, and over the model connected base. Billing lags energization by about a quarter.
- The 2027 check compares the change in identified energized base from Q4 2026 to Q4 2027 with the "more than 1 GW a year" pace. The pace is about capacity brought to market, which is closer to connected than to energized IT, so a miss against 1 GW on the energized series is expected and is labeled that way.

Energized base is the model's estimate of power that is live. It is the closest series to active power. Site rows keep their own basis: Vineland, Modi'in, and AIB are IT megawatts. Most other rows are stated megawatts whose IT-versus-gross basis Nebius does not give.

## Map

`src/nbis_scrapes/mapbuild.py` writes `data/out/map.html` and `data/out/sites.geojson`.

Circle area is contracted megawatts for the selected quarter, or the announced figure when the quarter has no contract. Green means energized base is above zero. Amber means capacity is on the books and energized base is zero. Gray is an unconfirmed lead.

Two views: Q2 2026, the last history quarter, and Q4 2026, the forward path. Oklahoma and the unnamed second French colo have no town, so they are listed and not pinned. Pins are city-level. The page says so.

The GPU-by-region table on the same page is the docs matrix. It is not pinned to buildings. Inferences live in `data/region_crosswalk.json` with confidence `inferred`.

## Demand scrapers

See `docs/demand-proxies.md`. The code is `src/nbis_scrapes/demand.py`.

- Nebius prices and the platform-by-region docs page are fetched with no key.
- OpenRouter's Nebius provider page supplies a model count.
- Vast metrics run only when `VAST_API_KEY` is a host key.

Parsers are tested on fixtures. A failed live fetch is stored as an error and does not invent a number.

## Filing sentences

`src/nbis_scrapes/edgar.py` extracts sentences that use the three defined terms. `fetch_text` caches the raw body under `data/cache/`, which is gitignored. Tests use `data/fixtures/q1_2026_power.txt`.

## Out of scope

Sell-side PDFs, the pitch documents, satellite imagery, and an authenticated console scrape. The ledger can be re-seeded from the workbook with `python scripts/seed_ledger.py PATH` if that file is on disk. openpyxl is required for the seed and is not required to rebuild the panel.

## How to run

```
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
python -m nbis_scrapes all
```

Open `data/out/map.html` in a browser. Tiles come from CARTO.
