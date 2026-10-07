# GPU availability

A price says what a GPU-hour costs. Availability says whether you can get one at that price. A list price can sit still for months while the chip behind it is sold out. This file is about the second question, for B200 first, then H100, H200, and B300.

The script is `scripts/fetch_availability.py`. The long file is `data/out/gpu_availability.csv`, one row per date, source, provider, GPU, region, and metric. Each daily run also writes `data/snapshots/availability-<time>.json`.

A day nobody observed has no row in the long file. In the wide files that day is a blank cell, the same as every other wide file in this repo. Nothing is carried forward or filled in.

## Why marketplaces

Nebius does not publish how full its self-serve cloud is. Nor does CoreWeave, Lambda, or Crusoe. What is public is the spot and marketplace layer around them: RunPod, Vast, Akash, and the brokers that resell neocloud capacity, Shadeform among them. Those markets are where a team that cannot get a B200 from a neocloud goes next, and where a neocloud with idle B200s can sell them. When the neoclouds are tight, the marketplaces fill up and their prices rise. When the neoclouds have slack, B200s show up on the marketplaces and sit there.

So a marketplace series is a read on neocloud and hyperscaler tightness one step removed. It is not Nebius. The mix is different: Vast is mostly independent hosts, Akash is a small decentralized network, and a few hundred B200s on Vast is a sliver of one Nebius campus. Read the direction and the turns, not the level.

## The sources in the committed file

**3Fourteen / Silicon Data chart.** 3Fourteen Research published a chart of B200 on-demand availability from Silicon Data, as a 30-day average. It was read off the chart by hand into a CSV. 264 days are points read off the line, from 1 Jan 2026 through 28 Sep 2026. Seven more days in the hand-read file were filled in between reads. Those seven are left out, so 2 to 6 Jan, 23 Feb, and 27 Feb are missing. The series starts at 62% on 1 Jan, is 30% by the end of January, 6% by the end of February, and under 1% from mid-March. It recovers to about 4.7% at the end of May and goes back to zero in September. It is one chart, read by eye, so treat a reading as good to about a point. Silicon Data does not say which clouds are in it. The source CSV is not in this repo. Load it with `--chart`.

**500.farm, Vast marketplace.** 500.farm runs an exporter over Vast's public offer list and keeps the result in Prometheus. Its public charts read that database through an anonymous Grafana proxy, `https://500.farm/vastai/grafana.v2/api/datasources/proxy/uid/EdgV2xcnz/api/v1/query_range`, and so does this repo: a form POST with `query`, `start`, `end`, and `step=86400`, one value per day at 00:00 UTC. No token is used or stored. robots.txt only closes `/vastai-exporter/`, which this repo never touches. Requests go one a second.

The count query is `sum by (gpu_name) (vastai_gpu_count{rented="yes"|"no", gpu_name=~"..."})`, summed over verified and unverified hosts. Rows per chip are GPUs listed and the percent not rented, which is GPUs available over GPUs listed. B200 starts on 26 Jun 2025, H100 SXM on 10 Jan 2024, H200 on 13 Nov 2024, and B300 on 2 Jun 2026. The history reaches back to Oct 2021 for A100. The committed file starts on 1 Jan 2024 and keeps B200, H100 SXM, H200, and B300. GPUs rented and available, the A100s, and the PCIe and NVL variants are in the cache.

Fleet size matters. Vast had 8 to 11 B200s listed in July 2025 and between 300 and 560 from April 2026, so the early B200 percentages move a lot on a few machines. The level is not the same as the Silicon Data chart. The B200 30-day average to 28 Sep 2026 is 8.9% on 500.farm against zero on the chart. The direction is the same: 43% free in January 2026, 38% in March and April, 16% in July, 7% in September.

**500.farm, Vast prices.** The same database has Vast's on-demand price percentiles, before Vast's fee, as the charts show them: `avg_over_time(vastai_v2_ondemand_price_<10th_percentile|median|90th_percentile>_dollars{verified="yes", rented="yes"|"no"}[1d])`. The v2 series split by datacenter host or not and by offer size (1–3, 4–7, 8+ GPUs). An older v1 series has no split and is kept as `all`/`all`. `data/out/gpu_prices_500farm.csv` is the headline cut: the nine datacenter chips, verified hosts, datacenter yes and no at 8+ GPUs plus the v1 series, with p10, median, and p90 side by side for rented and for unrented offers. The full tape, every chip and segment, is `data/cache/gpu_prices_500farm_full.csv`. The long file carries one price series per chip, the median of rented 8+ GPU offers on datacenter hosts. For B200 that is $4.70 when it starts on 8 Sep 2025, $2.87 on average in February 2026, and $8.03 in October 2026.

**dstack gpuhunt, RunPod.** dstack's open-source GPU catalog is rebuilt every hour and written as a zip to a public S3 bucket, `dstack-gpu-pricing`. Its RunPod file is built from RunPod's public GraphQL, and since the first RunPod build in April 2024 it drops every offer whose stock status is empty. So each copy of `runpod.csv` is a list of what RunPod had in stock when it was built. One copy a day is kept. Per GPU, four counts come out of it: secure-cloud data centers with any in-stock offer, community-cloud countries with any in-stock offer, distinct (location, GPU count) offers in stock, and the largest GPU count in stock in one offer.

There are 544 days, from 2 Apr 2024 through 5 Oct 2026. The 2024 copies stop on 12 Jun 2024. No copy with a RunPod file was found between then and 20 Feb 2025. GitHub no longer lists the build runs from that period, and stepping through the run numbers found nothing. May 2026 has 7 days and July 2026 has 9, because the combined catalog was only published when every provider built cleanly. From 18 Aug 2026 the RunPod file has its own key, `v3/runpod/<date>-<run>/catalog.zip`. The copy used for each day is listed in `data/fixtures/availability_gpuhunt_runpod_versions.txt`, so those rows carry no URL of their own.

A GPU gets a zero only after it has shown up in an earlier copy, because gpuhunt added chips to its map at different dates. B200 starts on 11 Jun 2025, the day after gpuhunt added it. H100 starts on 2 Apr 2024 and H200 on 20 Feb 2025. Three changes in gpuhunt itself move the counts: community cloud was dropped from 23 Oct 2024 and added back on 4 Mar 2025, and from 20 Nov 2025 offers on hosts below CUDA 12.8 are dropped. Community counts are written only on days the file has community offers at all. The daily run reads `v3/runpod/latest/catalog.zip`, so the series keeps growing on the same method. Its date is the day that catalog was built.

**Akash.** Akash's console API is public and documented at `https://console-api.akash.network/v1/doc`. `GET /v1/gpu-breakdown?startDate=&endDate=` returns, per day and GPU model, the GPUs on the network, the GPUs leased, the number of providers, and Akash's own utilization, leased over total. It takes up to 366 days a call. The long file keeps utilization and GPUs on the network. Leased GPUs, provider counts, and A100 are in the cache. H100 starts on 4 Mar 2024, H200 on 9 Jan 2025, and B300 on 25 Aug 2026. Akash has never listed a B200. The model name covers every form factor and memory size. Akash is small, 72 H100s in October 2026, and its utilization swings when one provider joins or leaves: H100 on Akash was 230 to 400 GPUs from November 2024 to November 2025 and fell to about 70 by March 2026.

**Shadeform, archived.** Shadeform resells about twenty clouds, Lambda, Nebius, Verda, Crusoe, Hyperstack, and Voltage Park among them, and its instance-types API marks each instance type in each region as available or not. The Internet Archive has 22 changed copies of that API from 10 Mar 2025 to 7 Oct 2026, B200 on 5 of them. Copies were sampled, one a day at most, one request a second. Rows are per cloud, as Shadeform reports them.

## The daily capture

The scheduled run reads five things.

1. 500.farm, the last two days at 00:00 UTC: counts and the headline prices, four queries.
2. Akash, the last three days. Today's row is still filling, so tomorrow's run replaces it.
3. RunPod stock. `https://api.runpod.io/graphql` answers without a key. The run asks for every GPU's stock level once overall (secure cloud, one GPU), once per listed data center (34 on 7 Oct 2026), and once per community country (10). The level is High, Medium, Low, or nothing. Rows are the overall level, the count of data centers and countries in stock, and the level in each data center that has stock. This is the question gpuhunt asks, without the CUDA filter, so the live count can run a little above gpuhunt's. The newest gpuhunt RunPod catalog is read too.
4. Vast offers. `https://console.vast.ai/api/v0/bundles/` is the search behind `vastai search offers`. It returns every listed machine, rentable or not. The run signs in with the Vast CLI session key in `~/.config/vastai/vast_tfa_key` when that file exists, then tries `VAST_API_KEY`, then searches with no key. Signed in, Vast returns more of the market: at 05:00 UTC on 7 Oct 2026 it showed 32 B200 machines, against 13 for the keyless search half an hour earlier. The `note` column says which way each day was read. The key is read at run time and never written to a file, a log, or a snapshot. A machine is listed several times, as 1, 2, 4, and 8 GPU offers. Its size is `num_gpus / gpu_frac`, and its free GPUs are its largest rentable offer. Rows per GPU are rentable offers, machines listed, GPUs listed, GPUs rentable, GPUs not rentable, rentable as a percent of listed, and the median price per GPU-hour across rentable offers (`dph_total / num_gpus`). The offer's `rented` field is about the caller's own rentals, so a GPU that is not rentable cannot be split into rented and held back by its host. On the 05:00 run, 46 of 228 listed B200s were rentable, 20.2%, at a median $7.81 a GPU-hour. H100 SXM was 42 of 175 at $2.40, H200 11 of 138 at $5.06, and B300 17 of 48 at $7.50. This series starts on 7 Oct 2026. Its history is the 500.farm series, which counts the same market another way.
5. Shadeform, `https://api.shadeform.ai/v1/instances/types`, which needs no key. Rows per GPU are listings, listings available, and the available share across all clouds, the same two counts per cloud, and a 1 or 0 per cloud and region. On 7 Oct 2026, none of the 15 B200 listings was available, including all 8 at Lambda. Shadeform shows only the capacity it can sell, so a cloud can have stock it does not offer through Shadeform.

The run also calls Vast's `metrics/gpu/current`, the market supply and demand that `demand.py` asks for, with the same key. Vast serves those numbers only to a host with an active registered machine. With a host account and no machine it answers 200 with `needs_machine: true` and no GPUs, and the run records that as a skip. The history endpoint has the same gate. Lambda's own API and Verda's need an account key and are recorded as skips. Their capacity is read through Shadeform.

## Kept in the cache, not committed

These were pulled once. Their rows are in `data/cache/gpu_availability_full.csv` and can be loaded again with `--import`. They are short, or they do not cover the four chips for 30 days.

- GitHub repositories that log availability, read through their commit history: RunPod stock levels (toolshedlabs-hash/gpu-price-feed, 23 Jul to 6 Oct 2026), Lambda regions with capacity (summitkhatiwada22/ai-compute-tracker, 29 days from 8 Sep 2026, B200 in no region on any day), Vast offer counts from five trackers (May to Oct 2026), and an hourly multi-cloud lake (gustofied/the-compute-bazaar, 25 Jul to 21 Aug 2026). Nothing on GitHub logs availability before April 2026.
- GetDeploying's B200 and B300 pages with each provider's AVAILABLE or UNAVAILABLE flag (25 days in the Internet Archive, 9 Mar to 2 Oct 2026) and its weekly dataset of providers listing each GPU (53 weeks from 6 Oct 2025). The weekly dataset counts listings, not stock.
- Prime Intellect's availability API (10 archived days from 28 Dec 2024), Vast's search API (13 archived days), and ComputePrices' out-of-stock badges (2 days).
- io.net's market snapshot, `https://api.io.solutions/v1/io-explorer/network/market-snapshot`, has GPUs total, hired, and idle per chip with no login, but no history. On 7 Oct 2026 all 512 of its H100s and all 8 B300s were hired. The parser is in the module. It is not in the daily run.

## Not pursued

- hashrate.no's Vast pages carry about 90 days of hourly rented and available counts per GPU, with B200 back to 21 Aug 2025. That is the same Vast market 500.farm already covers back to 2025, so it was not added.
- wovenai.ca has a live snapshot of Vast utilization, the same numbers as hashrate.no. Its history is in a paid plan.
- hostlens.app and SF Compute's price index need a login.
- Vast's own market metrics and history need an active host machine.

## What it shows

Four independent sources agree on the B200 path. The Silicon Data chart falls from 62% on 1 Jan 2026 to under 1% by mid-March. RunPod's B200 stock, as gpuhunt saw it, averaged 4.9 data centers in stock in February 2026, 2.8 in March, and 0.2 in April. B200 offers in stock went from about 22 a day in January and February to under 1 in April. On Vast, the share of B200s not rented fell from 43% in January to 7% in September while the fleet grew from about 60 to about 330, and the median rented B200 price went from $2.87 in February to $8.03 in October. On Shadeform, no B200 listing was available on 7 Oct 2026.

H100 went the other way on RunPod. It had H100 in stock in about 4 data centers a day from December 2025 through February 2026, and 9 to 10 from May 2026. H200 on RunPod held between 5.5 and 8.2 data centers a month from March 2025 on. On Vast both tightened in 2026. The free H100 share was 40–45% in November and December 2025 and 15–28% a month from March 2026, except June at 41%. H200 was 40–59% from October 2025 to February 2026 and 16–29% from March, except June at 47%.

## Files

| File | What it is |
|---|---|
| `data/out/gpu_availability.csv` | The committed long file: B200, H100, H200, B300. 500.farm rows from 1 Jan 2024. |
| `data/out/gpu_prices_500farm.csv` | 500.farm's Vast price tape, headline cut. |
| `data/out/wide/gpu_availability_<gpu>.csv` | One file per chip (`b200`, `h100`, `h200`, `b300`). Every series for that chip down the rows, dates across, the fiscal quarter above. |
| `data/out/wide/gpu_availability_500farm_pct.csv`, `_500farm_listed.csv`, `_500farm_price.csv` | 500.farm per chip: percent not rented, GPUs listed, and the median rented price. |
| `data/out/wide/gpu_availability_akash.csv` | Akash per chip: utilization and GPUs on the network. |
| `data/cache/gpu_availability_full.csv` | Every row as read, with URLs and raw text. Not committed. |
| `data/cache/gpu_prices_500farm_full.csv` | 500.farm's full price tape. Not committed. |

Rows whose address is fixed (the chart, gpuhunt, 500.farm, Akash) leave `source_url` blank in the committed file. The addresses are the ones in this document.

## Run

```
$env:PYTHONPATH = "src"
python scripts/fetch_availability.py
python -m nbis_scrapes availability
```

Both pull today's reads, add them to the long file and the cache without duplicates (a second read on the same day replaces the first), write a snapshot, and rebuild the wide files. A daily run is about 55 requests, about one a second.

To rebuild the history and the cache:

```
python scripts/fetch_availability.py --farm-history --akash-history
python scripts/fetch_availability.py --history --chart "B200 Availability (3Fourteen chart, digitized).csv"
```

`--farm-history` is fourteen queries to 500.farm and rewrites `data/out/gpu_prices_500farm.csv` and the full tape in `data/cache/`. `--akash-history` is one call a year. `--history` reads `runpod.csv` out of every catalog in the version list with range requests, about a thousand requests at 0.4 seconds apart. `--import` takes a research CSV with the long-file columns (`date, source, provider, gpu, region, metric, value, unit, observed_ts, source_url`, and `raw` in place of `note`), which is how the cached research rows came in. `--wide` only rebuilds the wide files.

The scheduled task `NBIS Daily Availability` runs `python scripts/fetch_availability.py` from this folder at 08:45 every day, with `PYTHONPATH=src`, through `pythonw` so no window opens. Its log is `%LOCALAPPDATA%\NBIS-Scrapes\availability.log`.
