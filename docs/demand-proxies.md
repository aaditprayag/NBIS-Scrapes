# Short-term compute demand

Nebius sells two different products into two different clocks. The Microsoft and Meta contracts are multi-year and show up in contracted megawatts and remaining performance obligations. Short-term demand is the other book: GPUs a customer can turn on this week, and tokens Nebius is willing to serve tonight. Those move in days. Filings move in quarters.

A useful proxy has to be public, repeatable, and closer to a stockout than to a press release. Five sources, in the order they earn their keep.

## 1. Nebius on-demand prices, and the hike already on the page

`https://nebius.com/prices`

This is the company telling you what a GPU-hour costs with no commitment. As of the page fetched on 6 Oct 2026, the table prints today's rate and a second rate effective 1 Oct 2026:

| SKU | On-demand | From 1 Oct 2026 |
|---|---:|---:|
| HGX B300 | $7.85 | $9.50 |
| HGX B200 | $7.15 | $8.50 |
| HGX H200 | $4.50 | $5.40 |
| HGX H100 | $3.85 | $4.50 |
| RTX PRO 6000 | $1.80 | $1.80 |
| GB300 NVL72 and GB200 NVL72 | Contact sales | Contact sales |

A scheduled hike on the liquid SKUs, with the small RTX card unchanged and the NVL72 racks held back for sales, is a demand print. Commitment discounts go to 35% off, so the list price is the top of the book, not the price Microsoft pays. Still, list price is the only price Nebius publishes for the self-serve book.

The same page prints spot floors (`from $0.79` on H100 and H200, `from $0.99` on B200 and B300). The footnote says the live spot rate moves with utilization and can change every 15 minutes, and that the cap is one cent under the on-demand rate. The live print sits in the console price calculator. That page needs a login, so this repo stores the public floor and leaves the 15-minute series for a later authenticated scrape.

`parse_nebius_prices` in `src/nbis_scrapes/demand.py` reads both tables. Run `python -m nbis_scrapes demand` to refresh.

## 2. Which chips are actually for sale, in which region

`https://docs.nebius.com/compute/virtual-machines/types`

This is the better supply snapshot, and it is easy to miss because it looks like documentation. The platform table is a stock list:

| Platform | Regions on the docs page |
|---|---|
| B300 (`gpu-b300-sxm`) | `uk-south1`, `eu-west2`, `us-north1` |
| B200 | `us-central1`, and a separate platform only in `me-west1` |
| H200 | `eu-north1`, `eu-north2`, `eu-west1`, `us-central1` |
| H100 | `eu-north1` only |
| L40S | `eu-north1` only |

H100 surviving in a single region, while B300 shows up in the UK, a second European region, and `us-north1`, is the generation mix of the self-serve cloud. It is not the Microsoft campus. Vineland can be full of GB300s for a contract and never appear on this table.

What to store each week: the set of `(platform, region)` pairs. A region dropping a platform is a stockout or a migration. A new region appearing is a campus coming online for self-serve, which is the moment energized megawatts start to matter for the short book.

Do not glue a region id onto a building. The docs page does not name cities. `data/region_crosswalk.json` holds inferences (Finland, Israel, Minnesota) and marks every one of them as inferred. The map draws that table beside the pins so the two are not confused.

The docs also mention a capacity advisor that checks quotas against physical capacity. That call is inside a project, so it needs an account. It is the right next scrape if a console login is available. Until then the public matrix is the scrape.

## 3. Vast.ai supply, demand, and the percentile price

`GET https://console.vast.ai/api/v0/metrics/gpu/current/`

Vast publishes, for hosts, a current count of supply and demand and a P10 / median / P90 price, refreshed about every five minutes, plus a history endpoint. That is the cleanest public thermometer for short-term GPU demand in the whole market: rented machines over listed machines, and whether the median moved.

It is not Nebius. It is the spot market Nebius's list price has to clear. A Nebius H200 hike from $4.50 to $5.40 means one thing if Vast's H200 median also jumped, and another thing if Vast is soft and Nebius hiked alone.

The endpoint requires a Vast **host** API key in `VAST_API_KEY`. A normal client key is rejected. The repo calls it when the key is present and records a skip when it is not. `summarize_vast_metrics` accepts the current-snapshot JSON. History is `GET /api/v0/metrics/gpu/history/` with `gpu_name`, `start`, `end`, and `step`. Pull H100, H200, and B200 daily. That series is the one worth keeping.

The offers search (`POST /api/v0/bundles/`) is a worse substitute. It needs the same key, it returns machines you could rent rather than a demand count, and the sample changes with the query. Use it only to audit the metrics endpoint.

## 4. OpenRouter, for the token factory rather than the training clusters

`https://openrouter.ai/provider/nebius`

Nebius Token Factory is listed as a provider. On 6 Oct 2026 the page title was "Access 12 models served by Nebius Token Factory." Count of models, and later the per-model input and output price versus the same model at other providers, is the inference-demand proxy.

This will not tell you whether Vineland is full. Token factory is a slice of GPUs pointed at short requests. It matters because management has been explicit that short-duration capacity was pricing above $40 million per MW per year in the Q3 window, against a 2026 installed-base yield near $12 million. If Nebius pulls models off OpenRouter or prices through it rise faster than peer providers, that slice is tight. If the model count grows while list-price GPUs stay in stock on the docs page, the short book has slack.

`parse_openrouter_provider_page` stores the model count from the page text. A later pass can hit OpenRouter's model endpoints and keep Nebius's price next to the cheapest other provider for the same model.

## 5. What not to scrape for this question

- Contracted gigawatts. That is the five-year land-and-power book. It answers a different question, and it is already in the quarterly panel.
- Job postings. They lag, and a commissioning role is a supply signal.
- Reddit and Discord. They are anecdotes. They are not a time series.
- Satellite photos. They can confirm a hall exists. They do not update weekly with a megawatt, and they do not see a stockout inside a live hall.
- A paid datacenter model. Useful as a manual override column. Not a script.

## What "tight" means in the snapshot

Read the four series together, once a week is enough, daily if Vast is keyed:

1. Nebius list price and the scheduled next price, by SKU.
2. Nebius spot floor, as the only public spot number.
3. The set of regions still selling each platform.
4. Vast median and demand/supply for the same SKU, if the key is set.
5. OpenRouter model count, and later Nebius's price gap versus peers.

A tight week is a hike, or contact-sales on a SKU that used to have a price, plus a region losing that platform, plus Vast's median up. A loose week is the opposite: platforms still listed everywhere they were, spot floor well under list, Vast median down. One of those without the others is a maybe.

Snapshots land in `data/snapshots/`. Raw fetches are not committed. The parsers are covered by fixtures in `data/fixtures/` so the tests do not need the network.
