# Checking a number

The README has the short version. This is the note for when a cell looks wrong: which column, which file, and what a rebuild will change.

The files were assembled on 7 Oct 2026. The site ledger was copied from a capacity workbook dated 5 Oct 2026, sheet `Quarterly MW`. `scripts/seed_ledger.py` only re-reads that workbook. It does not scrape site power.

## How a wide CSV is laid out

Every file in `data/out/wide/` has the same shape. Row 1 is the fiscal quarter. Nebius reports on a calendar year, so a date from 1 April through 30 June is `Q2 '25`. Row 2 is the column header, a quarter-end or a capture date. Each later row is one series. A blank cell means the source did not print a number.

`power_company.csv` runs Q1 2024 through Q4 2027: energized MW, reported active MW, connected MW and its tag, contracted MW and its tag, billing MW, and reported active divided by connected. History ends Q2 2026. Later columns are the workbook's forward path.

`power_sites_energized.csv` and `power_sites_contracted.csv` are the same quarters, one campus per row. A blank site cell means the campus is not in the books yet.

`power_cost.csv` is quarterly day-ahead power for Finland, France, the UK, Spain, and Estonia (Ember, euros per MWh, all hours) and PJM West peak (EIA ICE, dollars per MWh). The row under each price is the implied cost of one H100-hour at 0.7 kW. PUE is 1.15 in Finland, 1.2 in Estonia, 1.25 in France, the UK, and Spain, and 1.3 on PJM. Peak PJM is not the same product as a Nordic day-ahead average. Q4 2026 is omitted because Ember had only October.

The GPU page files use the day the price was on the page. The rate-card files use the effective date. A price Nebius labels "before 1 Jun 2026" sits on 31 May 2026. "From 1 Jun 2026" sits on 1 Jun 2026. Token Factory input and output are separate files, one model per row.

CoreWeave rows marked `(node)` are a node price divided by the GPU count on the page. Do not average those with the classic per-GPU row. Classic H100 SXM stays near $4.76. The 8-GPU node divided by 8 is about $6.16.

## Company vocabulary

The FY2025 Form 20-F, filed 30 Apr 2026, uses four words.

Greenfield: Nebius owns the land and runs the power. The filing names Mäntsälä, Missouri, and Alabama. Lappeenranta, Kapuli, and Pennsylvania are owned land announced around or after that filing, so the ledger classes them the same way.

Brownfield: an existing building or power asset. Defined, and no site in that list is given the label.

Build-to-suit: a developer owns the land and the power, and Nebius specifies the hall. Vineland, New Jersey, signed February 2025, is the one.

Co-location: leased space. Paris, two further French sites, Keflavík, the UK sites, three Israel sites, Kansas City, Minnesota, and Oklahoma are named this way through February 2026. Spain and Estonia are on the 12 Aug 2026 letter map and are leased.

Oklahoma and one French co-location are company-reported and have no town, so they are not pinned. Minnesota is company-reported only as Minnesota. Woodbury is the public region `us-north1`, not a campus disclosure. South Carolina is not on the 12 Aug map and not in the 20-F. It comes from AIB's 8-K on 30 Sep 2026. Romford is not a reported site.

The 12 Aug 2026 letter also describes an asset-light model. No campus is assigned to it.

Official region names, from `docs.nebius.com/overview/regions`, are in `data/region_crosswalk.json`. `eu-north2` is Iceland, not a second Finland site. `us-central1` is Kansas City. `us-north1` is Woodbury. A region id is not a building.

## Reported, estimated, modeled

The only completed active-power print in the window is about 170 MW at 31 Dec 2025. Connected power is published as targets: 220 MW for year-end 2025, and 800 MW to 1 GW for year-end 2026. Q1 2026 and Q2 2026 have no active actual and no connected actual.

Contracted figures in the letters are often as of the letter date, including a site announced that day.

Energized and billing cells in `data/ledger.json` are the workbook's estimates. The seed script copies them. It does not turn them into company prints. Year-end 2025 utilization of 170/220 uses the connected target. The actual connected stock was not disclosed.

Token Factory price stability is not a demand series. From 19 Jul 2026 to 7 Oct 2026, no model that stayed on the menu changed price. The menu itself changed.

## Rebuild

From the repo root:

```
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
python -m nbis_scrapes panel
python -m nbis_scrapes map
python scripts/build_wide_prices.py
python scripts/build_market_tapes.py
python scripts/fetch_peer_prices.py
python scripts/fetch_docs_prices.py
python scripts/fetch_token_factory_prices.py
python scripts/build_neighbors.py
```

`panel` and `map` stay offline. `fetch_peer_prices.py` hits the Internet Archive and takes several minutes. It keeps one capture per month.

`build_neighbors.py` reads caches under `data/cache/`, which git ignores: a PeeringDB facility dump and the Epoch AI data-center CSV. Epoch coordinates were geocoded with Nominatim and cached at `data/cache/epoch_geocode.json`. Without those caches the neighbor file will not match the committed CSV.

`python scripts/seed_ledger.py PATH` rewrites the ledger. Pass the workbook, or set `NBIS_CAPACITY_XLSX`. It needs openpyxl. Tests are stdlib `unittest`.

## Sources

- Nebius EDGAR, CIK 0001513845. FY2025 20-F: `https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm`. Q2 2026 shareholder letter, 12 Aug 2026: accession `0001104659-26-094568`, exhibit 99.2.
- GPU rate card: `https://docs.nebius.com/compute/resources/pricing.md`. The live page keeps the current card and one previous card. Older cards were read from earlier copies of that same docs URL.
- List prices: `https://nebius.com/prices`, `https://www.coreweave.com/pricing`, `https://www.coreweave.com/pricing/classic`, `https://www.crusoe.ai/cloud/pricing`, `https://lambda.ai/pricing`, `https://lambdalabs.com/service/gpu-cloud`.
- Token Factory: `https://tokenfactory.nebius.com/api/public/models_info`.
- Power prices: Ember monthly European wholesale file. EIA wholesale workbooks `ice_electric-2024final.xlsx`, `ice_electric-2025final.xlsx`, and `ice_electric-2026.xlsx`, hub `PJM WH Real Time Peak`.
- Neighbors: PeeringDB for the operator and the pin. Epoch AI, October 2026, when one of its campuses falls inside 60 km. A neighbor closer than 0.4 km is treated as the same pin. Each Nebius campus keeps the 15 nearest. Lappeenranta has none inside 60 km.

Fingrid's open-data API returned 401 without a key. energy-charts.info was not used. Its license bars redistributing the series.

## Gaps

`src/nbis_scrapes/edgar.py` pulls sentences that mention contracted, connected, or active power out of a saved letter. It does not walk the EDGAR index, and it does not write a new quarter into the panel. A campus announced after 30 Sep 2026 will not appear until the ledger is edited.

There is no bridge from billing megawatts to revenue, capex, and cash. The GPU-hour prices and the power-cost rows are inputs to that bridge.

Vast.ai's own market-metrics API needs a host key in `VAST_API_KEY` and an active host machine. That feed was not pulled. The availability file uses other public reads instead: 500.farm's Vast history, Akash, RunPod as gpuhunt recorded it, archived Shadeform, and a hand-read Silicon Data chart. The 15-minute Nebius spot price is only in the console. Token Factory has no public price history before 19 Jul 2026.

The later notes:

- `docs/third-party-gpu-prices.md` — gpuhunt, SkyPilot, and other trackers, checked against Nebius's list. They never set the price.
- `docs/console-catalog.md` — the 6 Oct 2026 console price list, 2,401 SKUs. No token is stored.
- `docs/competitor-capacity.md` — other projects within 200 km, by phase and quarter. Do not add campuses together.
- `docs/gpu-availability.md` — whether a chip was free on Vast, RunPod, Akash, and Shadeform. A blank day is a day nobody looked.
- `docs/open-model-ceo.md` — CEO comments, 2025 and 2026, on open models. The rows are `data/out/open_model_ceo.csv`. Chesky is the production case. Huang's shares are startups and an unnamed token series, not a Fortune 500 workload.

Two Lambda archive pages failed to download, July 2021 and August 2023. The monthly series still runs from July 2022.
