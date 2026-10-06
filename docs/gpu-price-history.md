# Nebius GPU list-price history

Source: every unique Internet Archive copy of `https://nebius.com/prices` from 15 Nov 2024 through 6 Oct 2026 (50 captures). The script is `scripts/fetch_wayback_prices.py`. Rows are in `data/out/gpu_price_history.csv`. A price is recorded on the day the archive crawled the page, so the change happened on or before that day, and after the previous crawl.

These are published list prices per GPU-hour. They are not the Microsoft or Meta contract price, and they are not the console's 15-minute spot print.

## On-demand

| SKU | First archived print | Still that price through | Next print |
|---|---|---|---|
| H100 | $2.95 on 15 Nov 2024 | 7 May 2026 | $3.85 by 12 Jul 2026 |
| H200 | $3.50 on 15 Nov 2024 | 7 May 2026 | $4.50 by 12 Jul 2026 |
| B200 | $5.50 on 25 Jul 2025 | 7 May 2026 | $7.15 by 12 Jul 2026 |
| B300 | $6.10 on 2 Apr 2026 | 7 May 2026 | $7.85 by 12 Jul 2026 |
| RTX PRO 6000 | $1.80 by 12 Jul 2026 | still $1.80 on 6 Oct 2026 | no change |

No archive copy between 7 May and 12 July 2026 contains the table, so that step is dated to the window, not to a day. Two December 2025 copies had no price table at all.

## The October 2026 column

From the 23 Sep 2026 capture, the page adds a second column, labeled effective 1 Oct 2026, and leaves the on-demand column where it was. That second column is still the second column on the 6 Oct capture. The docs page calls the same numbers the price "from October 1, 2026," and the old numbers the price "before."

| SKU | On-demand column, still, on 6 Oct | Effective 1 Oct column, from 23 Sep |
|---|---:|---:|
| H100 | $3.85 | $4.50 |
| H200 | $4.50 | $5.40 |
| B200 | $7.15 | $8.50 |
| B300 | $7.85 | $9.50 |

## Other rates on the same pages

- Reserve / commitment cards, Nov 2024 through 2025: H100 $2.00, H200 $2.30. The Nov 2024 table also prints those as the reserve column next to on-demand.
- Fixed preemptible, from the 12 Jul 2026 capture onward: H100 $2.15, H200 $2.45, B200 $3.95, B300 $4.30. The docs page shows those preemptible rates unchanged across 1 Oct.
- Spot floor, from 23 Sep 2026: "from $0.79" on H100 and H200, "from $0.99" on B200 and B300. That floor is not the fixed preemptible rate. The live spot price moves in the console and is not in the archive.

L40S rows swap which name sits on $1.55 and which sits on $1.82 between the 2025 pages and the April 2026 page. The vCPU bands move with the prices, so that is a label swap on the page, not a clean price cut. Those rows are in the CSV and are not in the table above.

Vast's own history endpoint is real (`/api/v0/metrics/gpu/history/`) and is the way to get a daily market series. It needs a Vast host key in `VAST_API_KEY`. This pull does not include it.
