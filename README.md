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

1. Official rate card: `https://docs.nebius.com/compute/resources/pricing.md`, the standalone-applications price page, and `changelog.md`. `python scripts/fetch_docs_prices.py` writes `data/out/docs_gpu_prices.csv`. Nebius keeps the current card and the previous one. Older cards are read from earlier copies of that same docs URL. On-demand H100 goes 2.118 (GPU line only, before 1 Oct 2025) → 2.95 → 3.85 on 1 Jun 2026 → 4.50 on 1 Oct 2026. H200 goes 2.668 → 3.50 → 4.50 → 5.40. B200 goes 4.5432 → 5.50 → 7.15 → 8.50. B300 goes 6.10 → 7.85 → 9.50. The write-up is `docs/gpu-price-history.md`.
2. `https://nebius.com/prices` — the marketing page. It is a second copy of the list price, not the rate card. Its archive is `data/out/gpu_price_history.csv`.
3. `https://docs.nebius.com/compute/virtual-machines/types` — which GPU platform is still sold in which region. H100 is only `eu-north1`. B300 is in `uk-south1`, `eu-west2`, and `us-north1`. A region id is not a building. Inferences are in `data/region_crosswalk.json` and on the map, kept off the pins.
4. OpenRouter's Nebius provider page — how many models Token Factory is serving. That is the short inference book, not Vineland.
5. Vast.ai `GET /api/v0/metrics/gpu/current/` — market supply, demand, and percentile price, about every five minutes. Needs a Vast **host** key in `VAST_API_KEY`. Client keys are rejected.

## Token Factory

The price source is `https://tokenfactory.nebius.com/api/public/models_info`. The catalog page says that JSON is the machine-readable list. Archived copies of that endpoint run from 19 Jul 2026 through 15 Sep 2026, plus a live pull. `python scripts/fetch_token_factory_prices.py` writes two CSVs:

- `data/out/token_factory_prices.csv` — one row per model per date, USD per million tokens.
- `data/out/token_factory_price_changes.csv` — only when an input or output price moves.

Across those dates, no model that stayed listed changed its input or output price. The catalog did change: nine model IDs from 19 Jul were gone by 7 Oct, and eight new ones were added. Prices for the ones that remained, including Qwen3-235B at $0.20 / $0.60 and GLM-5.1 at $1.40 / $4.40, are the same on every capture.

## Site types

Nebius uses four words in the FY2025 Form 20-F, filed 30 Apr 2026: greenfield, brownfield, build-to-suit, and co-location. Brownfield is defined and no site in that list is given the label. `data/out/site_types.csv` is the current classification.

Greenfield, as the company uses the word: Mäntsälä, and the owned US sites in Missouri and Alabama. Lappeenranta, Kapuli, and Pennsylvania are owned land announced around or after that filing, so they are classed the same way. Vineland is the one build-to-suit. Paris, Keflavík, Kansas City, the UK sites, Israel, Minnesota, Oklahoma, the later French sites, Spain, and Estonia are co-location. Oklahoma and one French site are company-reported and still have no town, so they are not pinned. South Carolina is not on the 12 Aug 2026 company map. It comes from AIB's 30 Sep 2026 8-K. Romford is not a reported site.

## Wide price CSVs

`python scripts/build_wide_prices.py` writes one CSV per series under `data/out/wide/`. The series names run down the rows. Each date is a column. The row above the dates is the fiscal quarter. Nebius reports on a calendar year, so every date from 1 Apr 2025 through 30 Jun 2025 is labeled `Q2 '25`.

`power_company.csv` is the quarterly power exhibit from Q1 2024 through Q4 2027. `power_sites_energized.csv` and `power_sites_contracted.csv` are the same quarters, one campus per row. A blank cell means the campus is not in the books yet. History runs through Q2 2026. Later columns are the forward path.

The rate-card files use the effective date of each published price. A price Nebius labels "before 1 Jun 2026" sits on 31 May 2026. A price labeled "from 1 Jun 2026" sits on 1 Jun 2026. Both are `Q2 '26`. The page files use the day that price was on the page. Token Factory input and output are separate files, with one row per model.

## Neocloud price tape and power cost

`python scripts/build_market_tapes.py` writes two tapes.

The price tape is the published on-demand GPU-hour card for Nebius, CoreWeave, Crusoe, and Lambda. Nebius is the archived page history. The other three are the card fetched with this build. CoreWeave publishes a node price, so H100 is $49.24 divided by 8 GPUs. Wide files are `data/out/wide/neocloud_nebius.csv`, `neocloud_coreweave.csv`, `neocloud_crusoe.csv`, and `neocloud_lambda.csv`. The long file is `data/out/neocloud_prices.csv`.

On the October 2026 card, an H100 is $3.85 at Nebius, $3.90 at Crusoe, $3.99 at Lambda, and about $6.16 at CoreWeave. An H200 is $4.50 at Nebius, $4.29 at Crusoe, and about $6.31 at CoreWeave. Nebius is priced with the cheap group, not with CoreWeave.

The power tape is `data/out/wide/power_cost.csv`. Finland, France, the UK, Spain, and Estonia are quarterly averages of Ember's monthly day-ahead prices, in EUR per MWh. PJM West is the quarterly average of EIA's ICE peak prices, in USD per MWh, and it is the public hub for the New Jersey and Pennsylvania sites. A second row under each price turns it into the electricity cost of one H100-hour at 0.7 kW, with a stated PUE (1.15 in Finland, 1.30 on PJM). In Q3 2026 that electricity cost is about 3 euro cents in Finland and about 10 US cents on PJM West, against a $4 to $6 GPU-hour list price. Wholesale power is a small slice of the posted GPU price. Finland's Q1 2026 average, 94 EUR/MWh, was the expensive quarter.

## Nearby data centers

`data/out/neighbors.csv` lists other facilities within 60 km of each confirmed Nebius campus. The operator comes from PeeringDB. A power figure is filled only when Epoch AI publishes an estimate for that campus, or when the PeeringDB note itself states a megawatt. Each row has a satellite-photo link. The map’s Satellite layer is the aerial photo. Tan pins are the neighbors. `neighbors_summary.csv` is the count per Nebius campus.

Lappeenranta has no other listed facility inside 60 km. Kansas City’s nearest listed building is the Patmos hall itself, 0.6 km from the city pin. Keflavík’s neighbor is Nscale, about 1.7 km away, with an Epoch estimate of 5 MW. AWS Berwick is 37 km from the Butler Township site, with an Epoch estimate of 108 MW.

## Third-party GPU prices

`python scripts/fetch_third_party_prices.py` reads Nebius GPU prices that other people published: dstack's gpuhunt catalog, SkyPilot's Nebius catalog in git, and 31 GitHub trackers and price sites. Rows are in `data/out/third_party_gpu_prices.csv`. Wide files are `data/out/wide/third_party_gpuhunt.csv`, `third_party_skypilot.csv`, and `third_party_trackers.csv`.

A third-party price never sets a Nebius price. It is checked against the rate card for that day, or the page archive before the card. A match to the cent counts that day as a day the price was seen. A miss stays in the file and is marked. gpuhunt has a Nebius file on 608 days from 8 Nov 2023, and its H100, H200, and B200 prices step on 1 Jun 2026 and 1 Oct 2026, the same days as the card. The matches give 524 days, and 489 of them have no Nebius capture in this repo. They are in `data/out/third_party_corroborated_days.csv`. The write-up is `docs/third-party-gpu-prices.md`.

## Console billing catalogue

The console's billing pages list every SKU in every region, with the date each price took effect and the date each SKU was made. They need a login, so this repo keeps the reader and one saved copy, not a scraper. `data/out/console_catalog_2026-10-06.csv` is 2,401 SKUs in 20 regions. `data/out/console_regions.csv` is the first SKU date per region.

Ten regions are on `docs.nebius.com/overview/regions` and keep the place in `data/region_crosswalk.json`. The other ten are not on the docs page yet. Vera Rubin SKUs sit at a price of zero in six regions, starting with `eu-north1` on 20 Aug 2026 and then `eu-north3`, `in-south1`, `me-west3`, `us-central2`, and `us-south2` in September. The places for the new regions are inferences, in `data/console_region_inferences.json`. How to save the pages from a logged-in browser, with no token stored, is in `docs/console-catalog.md`.

## Competitor capacity within 200 km

`python scripts/build_competitors.py` reads `data/competitor_projects.jsonl`, 651 AI, cloud, and colocation projects researched around the 22 Nebius campuses, split into 883 phases. Each phase has its megawatts, the quarter it comes online, and a source URL and quote for both. Ring 1 is within 50 km of the campus pin in `data/ledger.json`. Ring 2 is 50 to 200 km. The round trip is the straight-line distance divided by 100, in milliseconds, so it is a floor.

`data/out/competitor_phases.csv` is one row per phase. `data/out/competitors.csv` is one row per campus and phase, with the distance, ring, round trip, and whether the phase is in the totals. `data/out/wide/competitors_ring1.csv` and `competitors_ring2.csv` are the megawatts online by each quarter end from Q1 2024 through Q4 2030, one campus per row. A phase is in the totals only when its megawatts and its date are both public.

By the end of 2030, Paris has about 2.5 GW of dated capacity within 50 km, and Longcross and Harlow about 2.3 GW each. Lappeenranta, Béthune, Hüüru, and Spartanburg County have almost nothing within 50 km, and their competition sits 50 to 200 km out. The write-up is `docs/competitor-capacity.md`.

## GPU availability

`python scripts/fetch_availability.py` adds today's GPU availability to `data/out/gpu_availability.csv`: the last two days of 500.farm's Vast history, Akash's leased share, RunPod's stock level per data center, Vast's rentable share of listed GPUs and median price, and Shadeform's available flag for each cloud and region it resells, Lambda and Nebius included. Only the Vast search uses a key, the Vast CLI session key on this machine, and it falls back to a keyless search. Vast's own market metrics need an active host machine, so they are recorded as a skip.

The history is 500.farm's daily count of rented and free GPUs on Vast (B200 from 26 Jun 2025, H100 from Jan 2024) and its price tape, Akash's daily breakdown from March 2024, RunPod's in-stock offers as dstack's gpuhunt catalog recorded them (544 days from 2 Apr 2024), 22 archived copies of Shadeform's availability API, and a hand-read copy of 3Fourteen's chart of Silicon Data's B200 on-demand availability (264 days from 1 Jan 2026). A day nobody observed is missing, not filled. These are marketplaces, not Nebius. They read how tight the neoclouds are, one step removed.

The chart falls from 62% on 1 Jan 2026 to under 1% by mid-March. Vast's free B200 share falls from 43% in January to 7% in September while the median rented B200 price goes from $2.87 in February to $8.03 in October. RunPod's B200 stock fell from about 5 data centers a day in February to almost none in April. Wide files are `data/out/wide/gpu_availability_b200.csv` and the rest. The full rows and the full price tape are in `data/cache/`, and `python scripts/fetch_availability.py --farm-history --akash-history` rebuilds them. The scheduled task `NBIS Daily Availability` runs the daily capture at 08:45. The write-up is `docs/gpu-availability.md`.

## Run

```
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
python -m nbis_scrapes all
python scripts/fetch_docs_prices.py
python scripts/fetch_token_factory_prices.py
python scripts/build_wide_prices.py
python scripts/fetch_third_party_prices.py
python scripts/build_console_catalog.py
python scripts/build_competitors.py
python scripts/fetch_availability.py
```

`panel` and `map` stay offline. `demand` hits the network and writes `data/snapshots/`. Set `VAST_API_KEY` first if you want the market thermometer.

Rebuild the ledger from the capacity workbook, if that file is on this machine:

```
python scripts/seed_ledger.py
```

That step needs openpyxl. Everything else is the standard library.

Design notes: `docs/superpowers/specs/2026-10-06-nbis-power-alt-data-design.md`.
