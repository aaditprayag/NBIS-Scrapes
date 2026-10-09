# Wayback GPU price-history scraper (Nebius equity research)

Re-runnable pipeline that rebuilds `/workspace/nebius/market/gpu_price_history.xlsx` (+ copies in `/workspace/nebius/parts/`) from archived pricing pages.

Python env: `/workspace/nebius/.venv_ph` (pandas, openpyxl, matplotlib, requests, beautifulsoup4, lxml, brotli).

| step | script | what it does |
|---|---|---|
| 1 | `fetch.py <label> <url> [url...]` | CDX API (`statuscode:200`, `collapse=timestamp:6`, 2024+) -> one snapshot/month; fetches `web.archive.org/web/<ts>id_/<url>`, ~1 s sleep, retries on 429/5xx; caches HTML in `/workspace/nebius/wayback_raw/<url-slug>/<ts>.html`; logs every attempt to `wayback_raw/fetchlog_<label>.jsonl` |
| 2 | `gapfill.py <label> <prefix> [...]` | fills missing months from query-string / alternate URL variants (prefix match) |
| 3 | `parse.py` | per-provider regex parsers over visible text (`txt.py` handles gzip/brotli bodies). Emits only prices literally present in the snapshot |
| 4 | `run_parse.py [labels]` | applies parsers to all logged snapshots; prints headline series & coverage gaps (debug) |
| 5 | `build.py` | normalises to $/GPU-hr, adds live captures/APIs (labelled), writes workbook, CSVs, charts |

Labels -> providers: nebius, coreweave, coreweave2024, cwcks, lambda, together, crusoe, runpod, vast, nscale, aws (Capacity Blocks), vantage_aws / vantage_az (instances.vantage.sh = ec2instances.info), oracle, oracle_api (OCI price-list API), gcp, gcpvm, azure, awsp5.

Rebuild everything from cache: `cd /workspace/nebius && .venv_ph/bin/python market/scrapers/wayback/build.py` (~40 s).
Add a month: rerun `fetch.py` for the label (cached snapshots are skipped), then `build.py`.

Key conventions: instance prices / GPU count (AWS p5 /8, CoreWeave HGX /8, GB200 tray /4); per-minute x60; "starting at / as low as / from" tagged as floors in `price_basis`; `headline=TRUE` marks the one representative row per provider/SKU/pricing_type/term used in pivots. Forward-filled data lives only in `Monthly_FFill` and charts.

## Stage 2 additions
- New fetch labels (run_parse.py): awsp4, lambda2020, coreweave2021, runpod2023, runpod_models, vast_gpu, nscale_nodes, gcpframe, gcpvm, gcpvm2020, vantage_a100, vantage_az2. `WB_FROM=<year>` env var sets fetch.py start year.
- New parsers (parse.py): runpod_old, runpod_model, vast_gpu, aws_p4, gcp_vm, gcp_gpu_frame, contact_rows; Nebius reserve terms read from page header; Lambda 2022-23 layout.
- Helper scripts: cdxprobe.py (CDX coverage probe), gcp_frames.py (fetch GCP gpus-pricing iframes).
- New sheets: SKU_Nebius_vs_Medians, A100, A100_Decay, A100_Monthly_2020_2026_FFill/Obs, L40S, No_Public_Price, Coverage_by_SKU, Delisted_Series.
