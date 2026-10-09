# Nebius research – GPU cloud price scrapers

## 1. Live re-scrape (dated CSV)
```
python3 /workspace/nebius/market/scrapers/gpu_prices.py                # all providers
python3 /workspace/nebius/market/scrapers/gpu_prices.py --providers nebius,coreweave,aws_od --out /tmp/x
```
Writes `/workspace/nebius/market/data/price_snapshot_<today>.csv` (same 12 columns as the workbook).
Providers: nebius, coreweave, lambda, crusoe, together, runpod, aws_cb (EC2 Capacity Blocks page), aws_od (AWS public
on-demand price JSON, us-east-1), azure (Retail Prices API), oracle (OCI price-list API), gcp (accelerator-optimized
page), vast (Vast.ai public offer API; median of live verified on-demand offers).
A provider that fails is written as a row with `pricing_type=ERROR` so nothing fails silently. Only stdlib needed.

Parsers are regexes on page text, so they break when sites redesign. After a run, eyeball the CSV against the source
pages. Known caveats:
- Nebius shows two columns; the parser treats the "Effective <date>" column as current and tags the other one "superseded".
- RunPod: the static HTML doesn't say whether prices are the Community or Secure Cloud tier.
- Azure: reservations come back as total-term prices, converted with total/(8760*yrs)/GPUs. Spot meters equal to on-demand are flagged and left out.
- Not scrapeable (contact sales / blocked): Voltage Park, IREN, Nscale, Hyperbolic, FluidStack. These are covered by hand in rows_batch1/2.py.

## 2. Curated snapshot -> workbook
`rows_batch1.py` and `rows_batch2.py` hold the hand-verified 2026-10-04 rows, including contact-sales rows and notes.
```
cd /workspace/nebius && python3 build.py rows_batch1 rows_batch2
```
This rebuilds sheets `Price_Snapshot_2026-10-04` and `Price_Summary` in `/workspace/nebius/nebius_market.xlsx`. Other
sheets are left alone. It also rewrites `/workspace/nebius/price_snapshot_2026-10-04.csv`.

## 3. Chart
`python3 /workspace/nebius/market/scrapers/make_chart.py` reads the curated CSV and writes
`/workspace/nebius/market/charts/gpu_price_ondemand_2026-10-04.png`.

## 4. Comps charts
- comps_charts.py: builds charts/ev_per_contracted_mw_vs_growth.png and charts/ev_ebitda_and_cost_of_debt.png from parts/comps_values.json (run parts/build_comps.py first).

## 5. Demand signals
- See demand/README.md (jobs, Tranco, OpenRouter, customer stories, GitHub, HN mentions, docs changelog -> parts/demand_signals.xlsx + charts/demand_*.png).
