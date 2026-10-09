# NBIS charts (deck-ready PNGs, 200 dpi)

Regenerate: `python3 make_charts.py [share_price]`. The default price is 242.81, the Oct 2, 2026 close supplied by Aadit, not a filing figure. Requires matplotlib and numpy.

| File | What it shows | Data basis |
|---|---|---|
| `dilution_waterfall.png` | Issued shares (6/30/26), then + Aug-26 exchange shares, + NVIDIA pre-funded warrant, giving basic. Then + RSUs, + options (treasury stock method), + each convert if-converted (OTM notes = 0), + ATM remaining, giving fully diluted | Same math as the `CapStructure` tab in `../nebius_filings.xlsx`. At $242.81: 308.7M basic, 381.5M FD (+23.6%) |
| `kpi_revenue_by_segment.png` | Quarterly revenue by segment, Q3'24 to Q2'26 (AI cloud, TripleTen, Avride, eliminations) | `Quarterly` tab segment rows. Q4'24 and Q4'25 are Derived (FY less 9M). H1'24 split not disclosed |
| `kpi_arr.png` | Company-stated ARR, Q1'24 to Q2'26 | `Quarterly` tab ARR row. 2024 values from the Oct-24 investor presentation chart |
| `kpi_adj_ebitda.png` | Group adjusted EBITDA and margin | `Quarterly` tab. Q1'24 is Derived (H1 less Q2) |
| `kpi_capex_vs_da.png` | Purchases of PP&E/intangibles vs D&A | `Quarterly` tab. Q1'24 is Derived |
| `kpi_rpo_deferred_revenue.png` | RPO and deferred revenue, YE25 to Q2'26 | `Quarterly` tab (RPO first disclosed at YE25) |

Every value is cited cell by cell in `nebius_filings.xlsx` (Quarterly_Sources / CapStructure source columns): SEC form, date, accession and URL.
Style: navy/teal palette, no top/right spines, source footnote on every chart, sized 11x6 in (13x6.5 in for the waterfall). They drop straight into 16:9 slides.
If the workbook values change, update the lists at the top of `make_charts.py`. Values are hard-coded copies of the workbook for reproducibility.
