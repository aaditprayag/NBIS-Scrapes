# Third-party Nebius GPU prices

Nebius's own pages are the price. The archive of `nebius.com/prices` and the docs rate card are in `docs/gpu-price-history.md`. This file is about the other places a Nebius GPU price gets written down, and what they are good for. They are good for dates. They never set a price.

The script is `scripts/fetch_third_party_prices.py`. Rows are in `data/out/third_party_gpu_prices.csv`. The days they add are in `data/out/third_party_corroborated_days.csv`.

## The sources

**dstack gpuhunt.** dstack's open-source GPU catalog is rebuilt by a GitHub Action and written as a zip to a public S3 bucket, `dstack-gpu-pricing`. Old versions are never deleted. The 2023–24 builds sit at `v1/<date>/catalog.zip`. Later builds sit at `v2/<date>-<run>/catalog.zip` and, from Aug 2026, `v3/nebius/<date>-<run>/catalog.zip`, where the run number is the workflow run that built it. The bucket cannot be listed, so a past day is found by trying the run numbers GitHub reports for that day, or by stepping about 24 runs a day back from a known version. The zip is read with range requests, so only `nebius.csv` comes down.

One version per day is kept. The file has 608 days from 8 Nov 2023 through 7 Oct 2026. There is no Nebius file between 15 Apr 2024 and 1 Apr 2025, and two gaps of about a month in 2026 (29 Apr to 25 May, 24 Jun to 23 Jul) where no build could be found. The 2023–24 rows are the old Nebius AI cloud in `eu-north1-c`: H100 at $4.12, A100 at $2.66, L40 at $1.50, L4 at $0.50.

**SkyPilot.** SkyPilot keeps its Nebius catalog in `skypilot-org/skypilot-catalog` at `catalogs/v*/nebius/vms.csv`. A bot regenerates it from the Nebius billing API. Each commit is a copy of the card on the commit date. There are 19 dated copies from 3 Feb 2025 through 1 Oct 2026. SkyPilot writes every preset, so one GPU has several rows. The `SpotPrice` column is the preemptible price.

**GitHub trackers and price sites.** Thirty-one other sources list Nebius next to other clouds: GPU price trackers kept in git (one row per commit that holds a Nebius price), GPU price sites read live, and archived copies of those sites. They are noisy. One tracker lists an H100 on demand at $2.15, which is the preemptible price. One site still showed the $2.95 H100 four months after the June increase. That is why none of them sets a price.

## The rule

A third-party price is checked against the Nebius list price in force that day, for the same GPU and the same price type. The list is the rate card from its first dated price, and the archived price page before that. Component lines (the GPU part of a VM before 1 Oct 2025, and the L40S GPU line) are left out, because every third party quotes the whole VM.

`matches_nebius_list` is one of:

- `yes` — within half a cent of the list.
- `no` — off the list. `nebius_list_usd` shows the list it missed.
- `no Nebius list on file for that day` — the GPU is on the card, but this repo has no list price that early. Most of these are the 2023–24 gpuhunt rows. The old Nebius AI price page is not in this repo.
- `price type not comparable` — reserve terms, unlabeled prices, and spot after 8 Oct 2026, when spot stops being a fixed preemptible price.
- `not on the Nebius card` — A100, L40, and L4 from 2023–24.

A day goes into `third_party_corroborated_days.csv` only when at least one third-party price that day is a `yes`. A miss never makes a day. Monthly roll-ups (gpufinder's monthly floor) carry a made-up date and never make a day either. `nebius_capture_that_day` says whether the page archive or the rate card was already captured that day.

Of 9,781 rows, 6,830 match. The matches give 524 days, and 489 of them are days with no Nebius capture in this repo. gpuhunt is one of the matching sources on 459 of them.

## What it shows

gpuhunt and SkyPilot moved on the same days as the rate card. gpuhunt has H100 in `eu-north1` at $2.95 in every build from 1 Apr 2025 through 31 May 2026, $3.85 from 1 Jun 2026, and $4.50 from 1 Oct 2026. H200 goes 3.50 → 4.50 → 5.40 on the same days, in `eu-north1` and `us-central1`. B200 in `us-central1` goes 5.50 → 7.15 → 8.50. That puts the June increase on 1 Jun 2026, the date on the rate card, inside the 7 May to 12 Jul gap in the page archive.

Preemptible moved once, on 1 Jun 2026: H100 1.25 → 2.15, H200 1.45 → 2.45, B200 2.90 → 3.95. It did not move on 1 Oct 2026.

## Wide files

`data/out/wide/third_party_gpuhunt.csv` and `third_party_skypilot.csv` have one row per GPU, price type, and region. Where a source lists several presets for one GPU, the cell is the cheapest per-GPU price. `third_party_trackers.csv` has one row per source, GPU, price type, and region. The row above the dates is the fiscal quarter, as in the other wide files.

## Refresh

```
python scripts/fetch_third_party_prices.py
```

That adds the latest gpuhunt build and any SkyPilot commits after the last one on file, then re-checks every row. `--offline` only re-checks. `--import` takes observation CSVs in the same layout (`date, source, source_url, observed_ts, kind, item, region, price_type, gpu_count, price_usd, price_per_gpu_hour, raw`) and skips any row whose source is a Nebius page. The trackers were pulled once that way. Their parsers are one per repo and are not kept here.
