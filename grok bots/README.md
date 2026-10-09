# grok bots

This folder holds the work product of the AI research team behind Aadit Prayag's Nebius Group (NASDAQ: NBIS) long/short pitch for the Point72 stock pitch competition. The bots were Filings & Capital, Capacity & Power, GPU Market & Comps, Skeptic, Modeling and Peer KPIs, coordinated by Grok Bot. It is a snapshot as of **Oct 9, 2026**. Everything here comes from public information. Deliverables are Excel (.xlsx), Word (.docx), PDF and PNG files, and the Python scripts that built them are kept next to them.

| Folder | Owner bot | Key final files |
|---|---|---|
| `filings/` | Filings & Capital | `nebius_filings.xlsx`, `nebius_kpi_history.xlsx`, `capital_structure_exhibit.docx` / `.pdf`, `appendix_filings_capital_final.docx` / `.md`, `fd_share_bridge.docx`, `dilution_scenarios_onepager.docx`, `backlog_rpo_reconciliation.docx`, `q3_watchlist.docx`, `qa_msft_tranches.docx`, `charts/`, `altdata/altdata.xlsx` |
| `capacity/` | Capacity & Power | `nebius_capacity.xlsx`, `final_billed_vs_connected.xlsx`, `gpu_mw_capex.xlsx`, `vineland_power_permitting_exhibit.pdf` / `.png`, `vineland_site_appendix.docx` / `.pdf`, `vineland_shutdown.xlsx`, `vineland_watch.md`, `independence_tro_note.docx`, `altdata/altdata_timelines.xlsx`, `altdata/strips/` (satellite strips) |
| `market/` | GPU Market & Comps | `gpu_price_history.xlsx`, `contract_economics.xlsx`, `hyperscaler_contract_mw.xlsx`, `charts/contract_band_slide_final.png`, `charts/gpu_list_price_history_final.png`, `data/` (price snapshots), `scrapers/` (scraper code + demand/time-to-deploy workbooks) |
| `parts/` | GPU Market & Comps (deck-ready parts) | `contract_band_slide_final.png`, `gpu_list_price_history_final.png`, `price_history.xlsx`, `demand_signals.xlsx`, provider median charts |
| `model/` | Modeling | `nebius_model.xlsx` (v0.9e, live formulas; v1.0 export not yet approved), `debate_grid_v09e.xlsx` / `.png`, `judge_numbers_v09e.docx`, `judge_slide_v09e.docx`, `long_bull_v09e.docx`, `drivers_v09e.docx`, `build_model.py`; `nebius_model_recalc_preview_v05`–`v09d` are earlier versions |
| `skeptic/` | Skeptic | `06_judge_qa_attacks.docx`, `08_friday_debate_prep.docx`, `02_bear_outline_v3.5.docx`, `03_depreciation.xlsx`, `04_q3_decision_tree.docx`, `05_v09a_pressure_test.docx`, `07_friday_insight_shortlist.docx`, `speed/speed_study.xlsx`, `charts/` |
| `peers/` | Peer KPIs | `frozen_2026-10-11_v3.1/` (final frozen peer benchmark, with `SHA256SUMS`), `benchmark/peer_benchmark.xlsx` / `.docx` / `.pdf` / `.png`, `memo/nebius_vs_peers_unit_economics.docx`, `final_appendix/peer_kpi_appendix.docx`, `comparison.xlsx`, `price_or_mw.xlsx`, `coreweave/`, `iren/`, `oracle/`, `private/`, `sites/` (peer site verification + strips); earlier freezes `frozen_2026-10-11`, `_v2`, `_v3` kept for audit |
| `shared/` | All bots (Grok Bot) | `findings_log.xlsx` (team findings log), `append_finding.py` |
| `src/` | Filings & Capital | `wb/` (Wayback scraper code) and text extracts of Nebius letters and the 20-F |
| (root) | GPU Market & Comps | `nebius_market.xlsx`, `price_snapshot_2026-10-04.csv`, helper scripts |

## What was excluded and why

- **Virtualenvs** (`.venv_geo`, `.venv_ph`, ~462 MB): third-party packages, rebuildable.
- **Raw scrape caches** (`wayback_raw/` ~847 MB, `comps_raw/` ~58 MB, every `raw/` dir ~144 MB, scraped page-text caches ~22 MB): bulky intermediate inputs.
- **Satellite imagery tiles and arrays** (`*/imagery/*.png|.jpg|.npz`, ~1.2 GB): raw tiles. The composite `strips/` and the imagery metadata (json/csv/txt) are kept.
- **Third-party downloads** (EDGAR filings, 6-Ks, transcripts, permits, press PDFs/HTML, ~112 MB): public documents, not team work. They are listed in `EXCLUDED_SOURCES.md`.
- **Superseded `*_backup` files** (179 files, ~24 MB): explicitly named pre-change backups. Final and frozen versions are kept.
- **Python caches, lock files, symlinks** into excluded caches.
- No secrets (.env, tokens, cookies, credentials) were found. No file over 20 MB is included.
