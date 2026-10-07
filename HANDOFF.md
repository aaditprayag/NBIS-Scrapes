# Handoff

Date: 2026-10-07. Repo: private GitHub `aaditprayag/NBIS-Scrapes`. Local path: `C:\Users\aadit\NBIS-Scrapes`. Branch `master` is the only branch. Do all Nebius work in this repository and push it here. Do not start a second project folder.

This is a one-shot evidence pack for a Point72 stock-pitch competition on Nebius Group (Nasdaq: NBIS). The user will build the model, memo, and slides elsewhere. This repo holds the public-data work those documents should cite. It does not need to be a service that runs again later.

## What the user is trying to show

A 12-month long or short on NBIS, from public information only. The pitch needs one or two business insights that change the earnings or cash-flow view, not a tour of the scrape. The evidence already points at this distinction, which the user has not yet adopted as the recommendation:

Over the next year the P&L is billing megawatts times a yield. Contracted power (about 5 GW guided for year-end 2026) and connected power (800 MW to 1 GW guided) are the backlog. At year-end 2025 the company printed about 170 MW active against a 220 MW connected target. The forward path in the ledger has about 258 MW billing at year-end 2026 against that 800 MW to 1 GW guide.

Wholesale power does not explain the GPU price. In Q3 2026 the implied electricity cost of one H100-hour is about 3 euro cents in Finland and about 10 US cents on PJM West, against a $4 to $6 list price.

## How the wide CSVs are laid out

Every file in `data/out/wide/` has the same shape. Row 1 is the fiscal quarter. Nebius reports on a calendar year, so a date from 1 April through 30 June is `Q2 '25`. Row 2 is the column header: a quarter-end date or a capture date. Each later row is one series. A blank cell means the source did not print a number, not that the number is zero.

## Files to open

Power, the company:

- `data/out/wide/power_company.csv` — Q1 2024 through Q4 2027. Energized MW, reported active MW, connected MW and its tag, contracted MW and its tag, billing MW, and reported active divided by connected. History ends Q2 2026. Later columns are the forward path from the 5 Oct 2026 capacity workbook.
- `data/out/wide/power_sites_energized.csv` and `power_sites_contracted.csv` — same quarters, one campus per row.
- `data/out/site_types.csv` — greenfield, build-to-suit, or co-location.
- `data/out/map.html` — Nebius pins, tan pins for other facilities within 60 km, and a satellite base layer. City-level pins, not surveyed parcels.
- `data/out/neighbors.csv` and `neighbors_summary.csv`.

Power, the electricity:

- `data/out/wide/power_cost.csv` — quarterly day-ahead prices for Finland, France, the UK, Spain, and Estonia (Ember, EUR/MWh, all-hours) and PJM West peak (EIA ICE, USD/MWh). A second row under each price is the implied cost of one H100-hour at 0.7 kW. PUE is 1.15 in Finland, 1.2 in Estonia, 1.25 in France, the UK, and Spain, and 1.3 on PJM. Peak PJM is not the same product as Nordic day-ahead. Q4 2026 is omitted because Ember had only October.

GPU prices, on-demand, by chip:

- `data/out/wide/neocloud_nebius.csv` — 15 Nov 2024 through 6 Oct 2026, from archived `nebius.com/prices`.
- `data/out/wide/neocloud_crusoe.csv` — 18 May 2024 through 3 Oct 2026.
- `data/out/wide/neocloud_lambda.csv` — 1 Jul 2022 through 4 Oct 2026.
- `data/out/wide/neocloud_coreweave.csv` — Jan 2025 through 3 Oct 2026. Rows marked `(node)` are a node price divided by the GPU count on the page. Do not average those with the classic per-GPU row. Classic H100 SXM stays near $4.76. The 8-GPU node divided by 8 is about $6.16.
- `data/out/wide/gpu_rate_card_on_demand.csv` and `gpu_rate_card_preemptible.csv` — Nebius docs card. Columns are the effective dates. "Before 1 Jun 2026" sits on 31 May 2026. "From 1 Jun 2026" sits on 1 Jun 2026.
- `data/out/neocloud_prices.csv` — the long form, with the basis string for every cell.

Token Factory:

- `data/out/wide/token_factory_input.csv` and `token_factory_output.csv` — USD per million tokens, 19 Jul 2026 through 7 Oct 2026. No listed price changed for a model that stayed on the menu. The menu itself changed.

## Company vocabulary

The FY2025 Form 20-F, filed 30 April 2026, uses four words. Greenfield: Nebius owns the land and runs the power. The filing names Mäntsälä, Missouri, and Alabama. Brownfield: an existing building or power asset. Defined, and no site in that list is given the label. Build-to-suit: a developer owns the land and the power, Nebius specifies the hall. Vineland, New Jersey, signed February 2025, is the one. Co-location: leased space. Paris, two further French sites, Keflavík, the UK sites, three Israel sites, Kansas City, Minnesota, and Oklahoma are named this way through February 2026.

Lappeenranta, Kapuli, and Pennsylvania are owned land announced around or after that filing, so the ledger classes them greenfield. Spain and Estonia are on the 12 August 2026 letter map and are leased.

Oklahoma and one French co-location are company-reported and have no town, so they are not pinned. Minnesota is company-reported only as Minnesota. Woodbury is the public region `us-north1`, not a campus disclosure. South Carolina is not on the 12 August map and not in the 20-F. It comes from AIB's 8-K on 30 September 2026. Romford is not a reported site.

The 12 August 2026 letter also describes an asset-light model. No campus is assigned to it.

Official region names, from `docs.nebius.com/overview/regions`, are in `data/region_crosswalk.json`. `eu-north2` is Iceland, not a second Finland site. `us-central1` is Kansas City. `us-north1` is Woodbury. A region id is not a building.

## What is reported, and what is not

The only completed active-power print in the window is about 170 MW at 31 Dec 2025. Connected power is published as targets: 220 MW for year-end 2025, and 800 MW to 1 GW for year-end 2026. Q1 2026 and Q2 2026 have no active actual and no connected actual. Contracted figures in the letters are often as-of the letter date, including a site announced that day, not a prior quarter-end.

`data/ledger.json` was seeded from `nebius_capacity.xlsx` in the user's OneDrive, sheet `Quarterly MW`, workbook dated 5 Oct 2026. Energized and billing cells are that model's estimates. Do not describe them as company prints. The seed script is `scripts/seed_ledger.py`. Re-running it needs that workbook and openpyxl, and it rewrites the ledger.

## How to rebuild

From the repo root, with `PYTHONPATH=src`:

```
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

`fetch_peer_prices.py` hits the Internet Archive and takes several minutes. It keeps one capture per month. `build_neighbors.py` reads caches under `data/cache/` (gitignored): a PeeringDB facility dump and the Epoch AI data-center CSV. Epoch coordinates were geocoded with Nominatim and cached at `data/cache/epoch_geocode.json`. Without those caches the neighbor rebuild will not reproduce the committed CSV.

Tests are stdlib `unittest`. Forty tests were passing on 7 Oct 2026.

## Sources

- Nebius EDGAR, CIK 0001513845. FY2025 20-F: `https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm`. Q2 2026 shareholder letter, 12 Aug 2026: accession `0001104659-26-094568`, exhibit 99.2.
- GPU rate card: `https://docs.nebius.com/compute/resources/pricing.md`. The live page keeps the current card and one previous card. Older cards were read from earlier copies of that same docs URL.
- List prices: `https://nebius.com/prices`, `https://www.coreweave.com/pricing`, `https://www.coreweave.com/pricing/classic`, `https://www.crusoe.ai/cloud/pricing`, `https://lambda.ai/pricing`, `https://lambdalabs.com/service/gpu-cloud`.
- Token Factory: `https://tokenfactory.nebius.com/api/public/models_info`. The catalog page calls this JSON authoritative.
- Power prices: Ember monthly European wholesale file, CC-licensed, EUR/MWh. EIA wholesale workbooks `ice_electric-2024final.xlsx`, `ice_electric-2025final.xlsx`, and `ice_electric-2026.xlsx`, hub `PJM WH Real Time Peak`.
- Neighbors: PeeringDB facilities for the operator and the pin. Epoch AI, October 2026, for a power estimate when one of its campuses falls inside 60 km. A neighbor closer than 0.4 km is treated as the same pin. Each Nebius campus keeps the 15 nearest. Lappeenranta has none inside 60 km.

Fingrid's open-data API returned 401 without a key. energy-charts.info was not used. Its license bars redistributing the series.

## Do not

- Do not put this work in another folder.
- Do not describe an estimated megawatt as a company print. Read the tag column.
- Do not average a CoreWeave `(node)` row with the classic per-GPU row.
- Do not treat Token Factory price stability as a demand series. The menu changed. The prices of surviving models did not.
- Do not add more nearby-datacenter or satellite work unless it changes a twelve-month earnings view. The user asked for creative alt data and then chose the neocloud price tape and the power-cost tape.
- Do not walk a judge through the scrape. The presentation should carry the evidence for one insight.

## Still open

The ledger is not a live filing parser. `src/nbis_scrapes/edgar.py` extracts sentences that contain the three power terms. It does not walk the EDGAR index or write a new quarter into the panel. A new campus after 30 Sep 2026 will not appear until the ledger is edited.

There is no twelve-month bridge from billing megawatts to revenue, capex, and cash. That was offered and the user asked for the price tape and the power-cost tape instead. The GPU-hour prices and the power-cost rows are the inputs to that bridge. They are not the bridge.

Vast.ai supply and demand needs a host API key in `VAST_API_KEY`. It was not pulled. The 15-minute Nebius spot price is only in the console. Token Factory has no public price history before 19 Jul 2026.

Two Lambda archive pages failed to download, July 2021 and August 2023. The monthly series still runs from July 2022.
