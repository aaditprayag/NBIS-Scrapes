# NBIS

Filings, Nebius price pages, archived copies of those pages, and a campus map. The CSVs are already in `data/out/`. There is no price target in here.

`data/out/map.html` is the sites. `data/out/wide/power_company.csv` is the megawatts. In every wide file, row 1 is the fiscal quarter. Nebius reports on a calendar year, so 1 Apr through 30 Jun is `Q2 '25`. A blank cell is a figure the source did not print.

Power history runs through Q2 2026. Later columns are the path in a site workbook dated 5 Oct 2026. Read the tag before quoting a connected or contracted cell. `target` and `reported_floor` are company language. `estimated` is the workbook.

## Power

The May 2026 shareholder letter defines the three.

Contracted power is land plus a power commitment. The year-end 2026 guide is about 5 GW. The letters said above 2 GW at year-end 2025 and above 3.5 GW in Q1 2026. Those figures are often as of the letter date, and a site announced that day is in the number.

Connected power is a hall that is built and equipped. It can sit empty. What Nebius has published is a target: 220 MW for the end of 2025, and 800 MW to 1 GW for the end of 2026. Q1 2026 and Q2 2026 have no connected actual.

Active power is IT that is on and can bill. One company print in this window, about 170 MW at 31 Dec 2025. That beat a 100 MW active target. Nebius has not said whether the 220 MW connected target was met. 170 / 220 is 77%, against the target.

The site model has about 258 MW billing at the end of 2026. That is 32% to 26% of the 800 MW to 1 GW connected guide, and about 27% of the model's own 941 MW connected estimate. `data/out/SUMMARY.md` has the quarter-by-quarter table.

From Q4 2026 to Q4 2027 the same model adds about 893 MW of energized base. The company has talked about bringing more than 1 GW a year to market from 2027.

- `data/out/wide/power_company.csv` runs Q1 2024 through Q4 2027.
- `data/out/wide/power_sites_energized.csv` and `power_sites_contracted.csv` are the same quarters, one campus per row. A blank means the campus is not in the books yet.
- `data/out/site_types.csv` is greenfield, build-to-suit, or co-location, using the words in the FY2025 20-F.
- `data/out/map.html`. Circle size is contracted MW, or the announced figure when that quarter has no contract. Green means the model has energized MW in the quarter selected. Amber means the capacity is booked and energized MW is zero. Tan pins are other facilities within 60 km. Oklahoma and one French site have no public town, so they are left off. Pins are the town.

The 20-F was filed 30 Apr 2026. Greenfield means Nebius owns the land and runs the power. The filing names Mäntsälä, Missouri, and Alabama. Build-to-suit means a developer owns the land and the power and Nebius specifies the hall. Vineland, New Jersey, signed February 2025, is the one. Co-location is leased space: Paris, the later French sites, Keflavík, the UK, Israel, Kansas City, Minnesota, Oklahoma. Brownfield is defined, and no site in that list gets the label.

Lappeenranta, Kapuli, and Pennsylvania are owned land announced around or after that filing, so the ledger classes them with the greenfield sites. Spain and Estonia are on the 12 Aug 2026 letter map and are leased. South Carolina is not on that map. It comes from AIB's 8-K on 30 Sep 2026. Romford is not a reported site.

`eu-north2` is Iceland. `us-central1` is Kansas City. `us-north1` is Woodbury, Minnesota, which is the name of the public region, not a campus disclosure. `data/region_crosswalk.json` is the list. Which chips are still for sale, by region, is `docs/demand-proxies.md`.

## Prices

On the docs card, on-demand H100 goes from $2.95 on 1 Oct 2025 to $3.85 on 1 Jun 2026 and $4.50 on 1 Oct 2026. Before 1 Oct 2025 the chip line alone was $2.118, with CPU and RAM billed separately, and the docs say the hourly total did not change. H200 goes $3.50, $4.50, $5.40. B200 goes $5.50, $7.15, $8.50. B300 goes $6.10, $7.85, $9.50. The path is written up in `docs/gpu-price-history.md`. Effective dates are in `data/out/wide/gpu_rate_card_on_demand.csv` and `gpu_rate_card_preemptible.csv`. A price Nebius labels "before 1 Jun 2026" sits on 31 May.

The marketing page captured on 6 Oct 2026 still showed H100 at $3.85 in the on-demand column, and $4.50 in a second column effective 1 Oct. The comparison here uses the 6 Oct on-demand column. Crusoe was $3.90, Lambda $3.99, and a CoreWeave 8-GPU node divided by 8 was about $6.16. Rows marked `(node)` are that division. Don't average them with a per-GPU row. Classic CoreWeave H100 SXM stays near $4.76. H200 on that capture was $4.50 at Nebius, $4.29 at Crusoe, and about $6.31 at CoreWeave on the node basis.

List prices. The Microsoft and Meta deals are contracts, and the 15-minute spot price lives in the console.

Archived marketing pages, on demand: `data/out/wide/neocloud_nebius.csv` from 15 Nov 2024 through 6 Oct 2026, plus `neocloud_crusoe.csv`, `neocloud_lambda.csv`, and `neocloud_coreweave.csv`. `data/out/neocloud_prices.csv` is the long file. Every cell has a note on what was divided by what.

gpuhunt, SkyPilot, and 31 other trackers are a check on the dates. They do not set a Nebius price. A match to the cent counts as a day the price was seen. A miss stays in the file and is marked. gpuhunt has a Nebius file on 608 days from 8 Nov 2023, and its H100, H200, and B200 prices step on 1 Jun 2026 and 1 Oct 2026, the same days as the card. 524 days match, and 489 of those have no Nebius capture in this repo. `data/out/third_party_gpu_prices.csv`, `data/out/third_party_corroborated_days.csv`, `docs/third-party-gpu-prices.md`.

The console billing pages list every SKU in every region, with the day the price took effect and the day the SKU was created. Wider than the docs card. The saved copy is `data/out/console_catalog_2026-10-06.csv`, 2,401 SKUs in 20 regions. `data/out/console_regions.csv` is the first SKU date per region. Ten of the regions are on the public docs page. Vera Rubin SKUs are priced at zero in six regions, starting with `eu-north1` on 20 Aug 2026, then `eu-north3`, `in-south1`, `me-west3`, `us-central2`, and `us-south2` in September. Places for the new regions are guessed in `data/console_region_inferences.json` and marked as guesses. The pages need a login. No token is stored. `docs/console-catalog.md` is how the copy was saved.

## Electricity

`data/out/wide/power_cost.csv`. Finland, France, the UK, Spain, and Estonia are Ember day-ahead prices, all hours, euros per MWh. PJM West is the EIA reprint of ICE peak prices, dollars per MWh. That hub is the public one for the New Jersey and Pennsylvania sites. Peak PJM and an all-hours Nordic price are different products. The row under each price is the electricity in one H100-hour at 0.7 kW. PUE is 1.15 in Finland, 1.2 in Estonia, 1.25 in France, the UK, and Spain, and 1.3 on PJM.

In Q3 2026 that is about 3 euro cents in Finland and about 10 US cents on PJM West, against a $4 to $6 list price for the same hour. Finland's expensive quarter in the file is Q1 2026, at 94 EUR/MWh. Q4 2026 is left out because Ember only had October.

## Token Factory

`data/out/wide/token_factory_input.csv` and `token_factory_output.csv`. Dollars per million tokens, 19 Jul 2026 through 7 Oct 2026. Nothing earlier is on the public endpoint. No model that stayed on the menu changed price. The menu did. Nine IDs from 19 Jul were gone by 7 Oct, and eight had been added. Qwen3-235B stayed $0.20 in and $0.60 out. GLM-5.1 stayed $1.40 and $4.40.

## Other sites

`data/out/neighbors.csv` is facilities within 60 km. The operator comes from PeeringDB. A megawatt is filled only when Epoch AI or the PeeringDB note prints one. Lappeenranta has none inside 60 km. Nscale at Keflavík is about 1.7 km from Verne, with an Epoch figure of about 5 MW. AWS Berwick is about 37 km from Butler Township, Epoch 108 MW.

Out to 200 km, with an online quarter on each phase: 651 other AI, cloud, and colocation projects around the 22 campuses, 883 phases. Each phase has a source URL and a short quote for the megawatts and another for the date. If either is missing, the phase stays in the CSV and stays out of the totals. Ring 1 is within 50 km of the pin in `data/ledger.json`. Ring 2 is 50 to 200 km. The round-trip column is straight-line kilometers divided by 100, in milliseconds. Real fiber routes are longer.

By the end of 2030, Paris has about 2.5 GW of dated capacity within 50 km. Longcross and Harlow are about 2.3 GW each. Lappeenranta, Béthune, Hüüru, and Spartanburg County have almost nothing inside 50 km. The quarters are `data/out/wide/competitors_ring1.csv` and `competitors_ring2.csv`. Neighboring campuses share projects, so the rows do not add. `docs/competitor-capacity.md`.

## Availability

Nebius does not publish a fill rate for the self-serve cloud. CoreWeave, Lambda, and Crusoe do not either. `data/out/gpu_availability.csv` is what is public around them: 500.farm's Vast history, RunPod as gpuhunt recorded it, Akash, archived copies of Shadeform, and a 3Fourteen chart of Silicon Data's B200 availability, read off the line by hand. A day nobody observed is blank.

The chart is 62% on 1 Jan 2026 and under 1% by mid-March. Silicon Data does not say which clouds are in it. A reading is good to about a point. On Vast, the share of listed B200s that were free goes from 43% in January 2026 to 7% in September, and the median price of a rented B200 goes from $2.87 in February to $8.03 in October. RunPod's in-stock B200 offers go from about 5 data centers a day in February to almost none in April. `data/out/wide/gpu_availability_b200.csv`. H100, H200, B300, and the holes in the history are in `docs/gpu-availability.md`.

## Sources

- FY2025 Form 20-F, filed 30 Apr 2026: <https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm>
- GPU rate card: <https://docs.nebius.com/compute/resources/pricing.md>
- Marketing page: <https://nebius.com/prices>
- Token Factory: <https://tokenfactory.nebius.com/api/public/models_info>
- European power: Ember's monthly wholesale file. US hub: EIA ICE workbooks, `PJM WH Real Time Peak`.
- Neighbors within 60 km: PeeringDB, and Epoch AI (October 2026) where it prints a megawatt.
- Other prices: dstack gpuhunt, SkyPilot, and the trackers in `docs/third-party-gpu-prices.md`.
- Availability: 500.farm, Akash's console API, gpuhunt's RunPod file, Shadeform's archive, the 3Fourteen chart.

Rebuilding needs Python 3.11 or newer. The panel and the map stay offline. Fetching prices hits the web, and the older cards come from the Internet Archive.

```
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
python -m nbis_scrapes map
python scripts/build_wide_prices.py
python scripts/fetch_third_party_prices.py
python scripts/build_console_catalog.py
python scripts/build_competitors.py
python scripts/fetch_availability.py
```

`python scripts/seed_ledger.py PATH` rewrites `data/ledger.json` from the capacity workbook. Set `NBIS_CAPACITY_XLSX` if you would rather not pass the path. That step needs openpyxl. The committed ledger is enough for the other commands.

Vast's own market-metrics API wants a host key in `VAST_API_KEY` and an active host machine. That feed is not here. The availability file uses the sources above.
