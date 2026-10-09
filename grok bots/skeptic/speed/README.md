# Nebius Skeptic / Red Team: speed premium, residual value, credit

Output folder: `/workspace/nebius/skeptic/speed/`. Built 2026-10-04 (ET).

## Files
- `speed_study.xlsx`: workbook with live formulas. Yellow cells are inputs. Every row carries a source URL and a tag (Reported / Derived / Estimated).
  - **S1_PriceSeries / S1_Steps:** 470+ dated posted prices (Nebius, CoreWeave, Lambda, AWS, indices) and the step changes in each series.
  - **S1_Timeline:** first-availability dates. Lags are computed with formulas. Superseded by the GPU Market study at `/workspace/nebius/parts/time_to_deploy.xlsx`; see S1_Reconcile.
  - **S1_Decay:** launch premium versus trough, and half-life.
  - **S1_PrevGen:** ratios of each generation's price to the generation before it.
  - **S1_HS_Entry:** prices before and after each hyperscaler-entry event, with a summary computed by COUNTIF.
  - **S1_HeadStart:** NPV value of an N-month head start. Inputs, a 120-month table and a sensitivity grid.
  - **S1_Reconcile:** cross-check against the GPU Market teammate's lag table and A100 series. That data is reused, not re-scraped.
  - **S2_Residual:** used and refurbished H100/A100 marks versus 4/5/6-year book NBV.
  - **S2_RentalPrecedent:** A100/V100 rental paths, plus H100 projections for years 3–5.
  - **S3_Credit:** CRWV bond yields and spreads; NBIS and CRWV convert terms and bond-only yields; implied spread vs vol; mark-to-model; the exchange of the legacy notes; Merton.
  - **Sources:** source list.
- `speed_summary.md`: findings and verdict.
- `charts/*.png`: s1_h100_price_paths, s1_nebius_all_chips, s1_headstart_heatmap, s2_residual_vs_book, s2_rental_precedent, s3_convert_implied_spread.
- `data/`:
  - `price_series.csv`: parsed Wayback captures. Each row has its archive URL.
  - `hs_entry_test.csv`
  - `credit_model_out.json`
  - `ust_2026.csv`: Treasury par curves.
  - `nbis_aug26_pricing.*`: SEC pricing exhibit.
- `code/`: the scripts that build all of the above (see Rebuild).

## Rebuild (python3, openpyxl, matplotlib, numpy, scipy)
```
cd /workspace/nebius/skeptic/speed
python3 code/wb2txt.py          # /workspace/nebius/wayback_raw/*.html -> txt/ (teammate's cached captures; no re-scrape)
python3 code/parse_series.py    # -> data/price_series.csv
python3 code/hs_test.py         # -> data/hs_entry_test.csv
python3 code/credit_model.py    # -> data/credit_model_out.json (bond floor + Black-Scholes call on accreted strike; Merton)
python3 code/build_workbook.py  # -> speed_study.xlsx (reads /workspace/nebius/market/price_history_monthly.csv and time_to_deploy lag_summary.json)
python3 code/charts.py          # -> charts/
# optional recalc check: soffice --headless --convert-to xlsx --outdir /tmp/recalc speed_study.xlsx
```
- Hand-curated inputs are in `code/curated.py`. Each has a URL and a tag.
- MINIFS in S1_Timeline needs Excel 2019+ or LibreOffice.
- Python-solved values (convert-implied spread and vol, Merton) are pasted into the workbook as values and labelled Estimated.
