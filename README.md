# Nebius (NBIS) public data

Evidence for a Nebius Group pitch. Filings, Nebius's own price pages, archived copies of those pages, a campus map, competitor megawatts near each site, and public reads on whether a GPU is actually free. No price target lives in this repo.

Open [`data/out/map.html`](data/out/map.html) in a browser for the picture. Open [`data/out/wide/power_company.csv`](data/out/wide/power_company.csv) for the megawatts. A blank cell means the source did not print a number. It does not mean zero.

History in the power files runs through Q2 2026. Later quarters are the forward path from a site ledger dated 5 Oct 2026. Those later cells are estimates. The tag on each connected and contracted figure says whether Nebius printed it.

## Three different megawatts

Nebius uses three words, and they are not interchangeable. The May 2026 shareholder letter is the definition.

**Contracted** power is land plus a power commitment. It is the backlog. The year-end 2026 guide is about 5 GW. At year-end 2025 the company said contracted power was above 2 GW, and in Q1 2026 above 3.5 GW.

**Connected** power is a hall that is built and equipped. It can still be empty. Nebius has published this as a target, not as a quarter-end actual: 220 MW for the end of 2025, and 800 MW to 1 GW for the end of 2026.

**Active** power is IT equipment that is on and can bill. In this window there is one company print: about **170 MW at 31 Dec 2025**. That beat a 100 MW active target. Nebius has not said the 220 MW connected target was met. 170 / 220 is 77%.

The site model, which is not a company print, has about **258 MW billing at the end of 2026**. Against the 800 MW to 1 GW connected guide, that is 32% to 26%. Against the model's own 941 MW connected estimate, it is about 27%. The 5 GW contracted figure is a third object. It is not 5 GW of equipment sending invoices.

From Q4 2026 to Q4 2027 the same model adds about 893 MW to its energized base. The company has talked about bringing more than 1 GW a year to market from 2027. That pace sits next to connected power, not next to IT that is already on.

The full table, with a tag on every connected and contracted cell, is [`data/out/SUMMARY.md`](data/out/SUMMARY.md).

## Where to look

Power:

- [`data/out/wide/power_company.csv`](data/out/wide/power_company.csv) — company series, Q1 2024 through Q4 2027. Row 1 of every wide file is the fiscal quarter (`Q2 '25`). Nebius reports on a calendar year.
- [`data/out/wide/power_sites_energized.csv`](data/out/wide/power_sites_energized.csv) and [`power_sites_contracted.csv`](data/out/wide/power_sites_contracted.csv) — the same quarters, one campus per row.
- [`data/out/site_types.csv`](data/out/site_types.csv) — greenfield, build-to-suit, or co-location, using the words in the FY2025 Form 20-F.
- [`data/out/map.html`](data/out/map.html) — circle size is contracted megawatts. Green means the model has something on in that quarter. Amber means the capacity is booked and energized megawatts are zero. Tan pins are other facilities within 60 km. Oklahoma and one French site have no public town, so they are not on the map. Pins are the town, not a surveyed parcel.

GPU list prices, on demand, by chip:

- [`data/out/wide/neocloud_nebius.csv`](data/out/wide/neocloud_nebius.csv) — archived `nebius.com/prices`, 15 Nov 2024 through 6 Oct 2026.
- [`data/out/wide/neocloud_crusoe.csv`](data/out/wide/neocloud_crusoe.csv), [`neocloud_lambda.csv`](data/out/wide/neocloud_lambda.csv), [`neocloud_coreweave.csv`](data/out/wide/neocloud_coreweave.csv) — the same idea for the other three public cards.
- [`data/out/wide/gpu_rate_card_on_demand.csv`](data/out/wide/gpu_rate_card_on_demand.csv) and [`gpu_rate_card_preemptible.csv`](data/out/wide/gpu_rate_card_preemptible.csv) — Nebius's own docs card, on the day the price took effect. A price labeled "before 1 Jun 2026" sits on 31 May 2026.
- [`data/out/neocloud_prices.csv`](data/out/neocloud_prices.csv) — the long file. Every cell has a note on what was divided by what.
- The written price path is [`docs/gpu-price-history.md`](docs/gpu-price-history.md).

On the docs card, on-demand H100 goes from $2.95 (the all-in rate from 1 Oct 2025; the chip-only line before that was $2.118) to $3.85 on 1 Jun 2026 and $4.50 on 1 Oct 2026. H200 goes $3.50, then $4.50, then $5.40. B200 goes $5.50, then $7.15, then $8.50. B300 goes $6.10, then $7.85, then $9.50.

The marketing page captured on 6 Oct 2026 still showed the old on-demand column, H100 at $3.85, next to a second column of $4.50 effective 1 Oct 2026. The peer comparison below uses that 6 Oct on-demand column. Crusoe was $3.90, Lambda $3.99, and a CoreWeave 8-GPU node divided by 8 was about $6.16. Rows marked `(node)` are that division. Leave them off any average with a per-GPU row. CoreWeave's classic H100 SXM stays near $4.76. H200 on the same capture was $4.50 at Nebius, $4.29 at Crusoe, and about $6.31 at CoreWeave on the node basis.

These are list prices. They are not the Microsoft or Meta contract, and they are not the 15-minute spot print inside the Nebius console.

Electricity:

- [`data/out/wide/power_cost.csv`](data/out/wide/power_cost.csv) — quarterly wholesale power, and under each price the electricity cost of one H100-hour at 0.7 kW.
- Finland, France, the UK, Spain, and Estonia are Ember day-ahead prices, all hours, in euros per MWh. PJM West is the EIA reprint of ICE peak prices, in dollars per MWh, and it is the public hub for the New Jersey and Pennsylvania sites. Peak PJM is a different product from an all-hours Nordic day-ahead price.
- The assumed PUE is 1.15 in Finland, 1.2 in Estonia, 1.25 in France, the UK, and Spain, and 1.3 on PJM. In Q3 2026 that H100-hour of electricity is about 3 euro cents in Finland and about 10 US cents on PJM West. The list price of the same hour is $4 to $6. Finland's expensive quarter in this file is Q1 2026, at 94 euros per MWh.

Token Factory, Nebius's inference menu:

- [`data/out/wide/token_factory_input.csv`](data/out/wide/token_factory_input.csv) and [`token_factory_output.csv`](data/out/wide/token_factory_output.csv) — dollars per million tokens, 19 Jul 2026 through 7 Oct 2026. Nothing earlier is on the public endpoint.
- No model that stayed listed changed its input or output price. The menu did. Nine model IDs from 19 Jul were gone by 7 Oct, and eight new ones had been added. Qwen3-235B stayed $0.20 in and $0.60 out. GLM-5.1 stayed $1.40 and $4.40.

Other facilities within 60 km:

- [`data/out/neighbors.csv`](data/out/neighbors.csv) — operator from PeeringDB. A megawatt is filled only when Epoch AI or the PeeringDB note prints one. [`neighbors_summary.csv`](data/out/neighbors_summary.csv) is the count per campus.
- Lappeenranta has no other listed facility inside 60 km. Nscale at Keflavík is about 1.7 km from Verne, with an Epoch estimate of about 5 MW. AWS Berwick is about 37 km from the Butler Township site, with an Epoch estimate of 108 MW.
- The wider cut, out to 200 km and split into phases with an online quarter, is below.

Which chips are still offered, and in which Nebius region, is written up in [`docs/demand-proxies.md`](docs/demand-proxies.md). A region id is not a building. `eu-north2` is Iceland. `us-central1` is Kansas City. `us-north1` is Woodbury, Minnesota, the name of the public region, not a separate campus disclosure. The crosswalk is [`data/region_crosswalk.json`](data/region_crosswalk.json).

## What Nebius actually printed

The only completed active-power number in the window is about 170 MW at 31 Dec 2025. Q1 2026 and Q2 2026 have no active actual and no connected actual. Connected figures in the files are a target or an estimate. Energized megawatts and billing megawatts are the site model. Read the tag before you quote a cell.

Contracted figures in the shareholder letters are often as of the letter date, including a site announced that day. They are not always a prior quarter-end.

The FY2025 Form 20-F, filed 30 Apr 2026, sorts sites into four words. Greenfield: Nebius owns the land and runs the power. The filing names Mäntsälä, Missouri, and Alabama. Build-to-suit: a developer owns the land and the power, and Nebius specifies the hall. Vineland, New Jersey, signed February 2025, is the one. Co-location: leased space. Paris, further French sites, Keflavík, the UK, Israel, Kansas City, Minnesota, and Oklahoma are named that way. Brownfield is defined, and no site in that list is given the label.

Lappeenranta, Kapuli, and Pennsylvania are owned land announced around or after that filing, so the ledger classes them with the greenfield sites. Spain and Estonia are on the 12 Aug 2026 letter map and are leased. South Carolina is not on that map. It comes from AIB's 8-K on 30 Sep 2026. Romford is not a reported site.

## Prices other people wrote down

Nebius's own pages set the price. gpuhunt, SkyPilot, and 31 other trackers are only a check, and a way to fill days this repo never captured directly. A match to the cent counts that day as a day the price was seen. A miss stays in the file and is marked as a miss.

gpuhunt has a Nebius file on 608 days from 8 Nov 2023. Its H100, H200, and B200 prices step on 1 Jun 2026 and 1 Oct 2026, the same days as Nebius's card. Those matches cover 524 days, and 489 of them have no Nebius capture here. Files: [`data/out/third_party_gpu_prices.csv`](data/out/third_party_gpu_prices.csv), [`data/out/third_party_corroborated_days.csv`](data/out/third_party_corroborated_days.csv), and [`docs/third-party-gpu-prices.md`](docs/third-party-gpu-prices.md).

## The console's own price list

The billing pages inside the Nebius console list every SKU in every region, with the day the price took effect and the day the SKU was created. That list is wider than the public docs card. The saved copy is [`data/out/console_catalog_2026-10-06.csv`](data/out/console_catalog_2026-10-06.csv): 2,401 SKUs in 20 regions. [`data/out/console_regions.csv`](data/out/console_regions.csv) is the first SKU date in each region.

Ten of those regions are on the public docs page. Ten are not. Vera Rubin SKUs sit at a price of zero in six regions, starting with `eu-north1` on 20 Aug 2026, then `eu-north3`, `in-south1`, `me-west3`, `us-central2`, and `us-south2` in September. The places guessed for the new regions are in [`data/console_region_inferences.json`](data/console_region_inferences.json), and they are labeled as guesses.

The catalogue needs a login. This repo does not log in and does not store a token. [`docs/console-catalog.md`](docs/console-catalog.md) says how the pages were saved from a browser that was already signed in.

## Competitor megawatts near each campus

[`data/out/neighbors.csv`](data/out/neighbors.csv) is the buildings within 60 km. This cut is the megawatts, out to 200 km, by the quarter each phase comes online.

651 other AI, cloud, and colocation projects around the 22 Nebius campuses, split into 883 phases. Each phase has its megawatts, the quarter it comes online, and a source URL and a quote for both. A phase is in the totals only when both the megawatts and the date are public. Ring 1 is within 50 km of the campus pin. Ring 2 is 50 to 200 km. The round-trip figure is the straight-line distance divided by 100, in milliseconds. Real fiber is longer, so that number is a floor.

By the end of 2030, Paris has about 2.5 GW of dated capacity within 50 km, and Longcross and Harlow about 2.3 GW each. Lappeenranta, Béthune, Hüüru, and Spartanburg County have almost nothing within 50 km. Their competition sits in the outer ring. The quarters are in [`data/out/wide/competitors_ring1.csv`](data/out/wide/competitors_ring1.csv) and [`competitors_ring2.csv`](data/out/wide/competitors_ring2.csv). The write-up is [`docs/competitor-capacity.md`](docs/competitor-capacity.md). Do not add campuses together. Neighboring sites share projects.

## Whether you can get the chip

A list price can sit still while the chip behind it is sold out. Nebius does not publish how full its self-serve cloud is. Neither do CoreWeave, Lambda, or Crusoe. What is public is the marketplace around them.

[`data/out/gpu_availability.csv`](data/out/gpu_availability.csv) holds those reads. A day nobody looked is left blank. Nothing is carried forward.

The B200 picture, from sources that are not Nebius:

- A hand-read 3Fourteen chart of Silicon Data's on-demand availability starts at 62% on 1 Jan 2026 and is under 1% by mid-March. Silicon Data does not say which clouds are in the chart. Treat a reading as good to about a point.
- On Vast, the share of listed B200s that were free goes from 43% in January 2026 to 7% in September. The median price of a rented B200 goes from $2.87 in February to $8.03 in October.
- RunPod's in-stock B200 offers go from about 5 data centers a day in February to almost none in April.

Wide files start at [`data/out/wide/gpu_availability_b200.csv`](data/out/wide/gpu_availability_b200.csv). The long version, including H100, H200, and B300, is [`docs/gpu-availability.md`](docs/gpu-availability.md). These series say whether the market around Nebius was tight. They do not say how full a Nebius hall was.

## Sources

- FY2025 Form 20-F, filed 30 Apr 2026: <https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm>
- GPU rate card: <https://docs.nebius.com/compute/resources/pricing.md>
- Marketing prices: <https://nebius.com/prices>
- Token Factory, the JSON the catalog calls authoritative: <https://tokenfactory.nebius.com/api/public/models_info>
- European wholesale power: Ember's monthly file, CC-licensed. US hub: EIA ICE workbooks, hub `PJM WH Real Time Peak`.
- Neighbors: PeeringDB for the operator and the pin. Epoch AI, October 2026, for a power estimate when one of its campuses falls inside 60 km.
- Third-party prices: dstack gpuhunt, SkyPilot's Nebius catalog, and the trackers named in [`docs/third-party-gpu-prices.md`](docs/third-party-gpu-prices.md).
- Availability: 500.farm's public Vast history, Akash's console API, RunPod as gpuhunt recorded it, archived Shadeform, and the 3Fourteen chart.

The CSVs and the map are already built. Rebuilding them needs Python 3.11 or newer and the standard library.

## Run

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

`python scripts/seed_ledger.py` rewrites `data/ledger.json` from the capacity workbook. Pass the workbook path, or set `NBIS_CAPACITY_XLSX`. That step needs openpyxl. The committed ledger is enough for every other command.

Fetching a fresh price page hits the public web and, for the older cards, the Internet Archive. The commands are at the bottom of [`docs/gpu-price-history.md`](docs/gpu-price-history.md) and [`docs/demand-proxies.md`](docs/demand-proxies.md). Vast.ai's own market-metrics API needs a host key in `VAST_API_KEY`. That feed is not in this repo. The availability file uses the public sources named above.
