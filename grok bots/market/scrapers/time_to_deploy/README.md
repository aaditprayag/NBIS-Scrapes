# Time-to-deploy study (Nebius GPU Market & Comps)

Question: does Nebius make each new NVIDIA generation available faster than its peers, and so capture launch-premium pricing?

## Pipeline (run with /workspace/nebius/.venv_ph/bin/python)
1. `cdx_full.py`: pulls digest-collapsed Wayback CDX listings for the Nebius, CoreWeave and Lambda pricing pages and writes them to `/workspace/nebius/wayback_raw/cdx_full_*.json`.
2. `fetch_snapshots.py URL TS [TS...]`: fetches raw captures using the `id_` modifier, at about 1 request per second. Captures are cached under `/workspace/nebius/wayback_raw/<safe_url>/<ts>.html`. This is the same cache the price-history workstream uses (shared helper in `/workspace/nebius/src/wb/`).
3. `totext.py`: converts the HTML to `cache_text/`. It is incremental.
4. `parse_prices.py`: writes `out/price_obs.csv`, one row per provider×chip×capture with $/GPU-hr, the basis text and the archive URL. It handles Nebius "from Oct 1" announced columns, CoreWeave instance/8 math and Lambda per-GPU rows.
5. `data_curated.py`: hand-curated NVIDIA milestones, the reference GA dates and the provider announcement, preview and GA dates. Every entry has a source URL and a tag.
6. `analyze.py`: computes the lag table, launch premium, as-of prices at +3/+6/+12 months, log-linear decay and the pricing-page change log.
7. `build_workbook.py [--partial]`: writes `/workspace/nebius/parts/time_to_deploy.xlsx`.
8. `charts.py`: writes PNGs to `/workspace/nebius/market/charts/`.

## Methods
- **Lag** = provider GA (or first public price) − NVIDIA reference GA, in days.
  - Providers that are not yet GA are treated as censored at today − reference.
  - Alternative NVIDIA definitions are shown in their own columns.
- **As-of price**: a step function over captures. A change is dated to the first capture that shows it, so every date is an upper bound.
- **Decay**: the step function is sampled monthly from the first listed price. The fit is OLS of ln(price) on month, and %/mo = exp(b) − 1. The fit is only run when there are at least 3 points.
- Together AI's parse is noisy, so it is used for availability only. AWS figures are Capacity Block (reserved) prices and are not comparable to on-demand prices.
