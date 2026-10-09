# Oracle OCI (ORCL) peer KPIs on the Nebius/CoreWeave template, Cloud Infrastructure perimeter

Built 2026-10-05 ET. Light polish 2026-10-05 afternoon: EDGAR re-check — no new ORCL 8-K/10-Q since 2026-10-04; KPIs unchanged. Abilene / Port Washington site cross-links added. Files:
- Workbook: `/workspace/nebius/peers/oracle/oracle_kpis.xlsx` (recalculated; 0 formula errors).
- Builder: `/workspace/nebius/peers/oracle/build_orcl_kpis.py`.
- Shared engine: `/workspace/nebius/peers/peer_kpi_template.py`.
- Debt-table parser: `oracle/work/parse_orcl_debt.py` → `orcl_debt_tables.json`.
- Raw filings: `oracle/raw/` and `comps_raw/orcl_*`.

**Template.** The row IDs, definitions, tags and sourcing rules match CRWV and IREN. Oracle-only memo rows are I40+ and M01+.

**Fiscal mapping.** Oracle's fiscal year ends May-31, so each fiscal quarter ends one month before the calendar quarter-end. Each calendar column uses the Oracle quarter that ends one month earlier:
- Q1'24 = FQ3 FY24 (Feb-24), Q2'24 = FQ4 FY24 (May-24), …, Q2'26 = FQ4 FY26 (May-26).
- An extra column, **Q3'26† = FQ1 FY27 (Aug-26)**, is the latest filed quarter.

**OCI is not a reported segment.** AI-infra (Stargate / GB200–B300 clusters) sits inside Cloud infrastructure (IaaS) revenue and cannot be separated in filings — treat I01 as OCI (incl. non-AI), and use call-based MW delivered (I09N / M07) as the AI-capacity proxy only.
- I01 is the "Cloud infrastructure" (IaaS) revenue line, which includes non-AI OCI.
- Capex, depreciation, PP&E, RPO and debt are company-wide.
- **Segment-allocation estimates are tagged Estimated (orange):**
  - capex is attributed 100% to OCI (parameter P05, an upper bound);
  - incremental OCI revenue is attributed to delivered AI MW (K43);
  - MW delivered comes from earnings calls, with two quarters estimated.

## Q2'26 (FQ4 FY26, May-26) and latest (FQ1 FY27, Aug-26)

| KPI | Q2'26 (May-26) | Latest (Aug-26) | Tag | CRWV Q2'26 |
|---|---|---|---|---|
| OCI revenue (I01) | $5,787M (+93% YoY) | $7,388M (+121%) | R | 2,575 |
| K01/K02 revenue per installed MW | n/d (no installed-MW disclosure) | n/d | – | 6.87 / 8.24 |
| K43 incremental OCI revenue per delivered MW (YoY) | 9.3 $M/MW-yr (on ">1.2 GW" delivered in FY26) | 8.3 | E | 5.29 |
| K10 capex per MW delivered (100% allocation) | 56.9 (16.5B / ~290 MW) | 33.5 (28.5B / 850 MW) | E | 12.8 |
| K13 trailing-4Q capex per trailing-4Q MW | 46.4 (55.7B / 1.2 GW) | 39.0 | E | 20.0 |
| M06 capex per $ of incremental OCI run-rate | 4.98× | 4.68× | E | – |
| K16 server/network useful life | 6 years (5→6 effective FY25; 4→5 in FY23) | 6 | R | 6 |
| K17 depreciation yield | n/d quarterly (FY26 annual M05: 12.1%) | 13.4% | D | 15.5% |
| K20 RPO (company) | $638B | $664B | R | 103.7 |
| K42 / K22 RPO within 12 months / 24 months | 12% / ~29% (24-month figure estimated) | 13% / ~31.5% | R / E | n/d / 41% |
| K23 RPO / total revenue run-rate | 8.3 years | 8.6 years | D | 10.1 |
| K24–K27 concentration | no customer ≥10% of FY revenue (count 0) | – | R | 36% / 72% / 3 |
| K28 weighted effective cost of debt (10-K table) | 4.83% (senior notes 4.81%, CP 4.35%, term loans 5.29%) | n/d (annual only) | D | 8.41% |
| K36 interest × 4 / average carrying debt | 4.36% | 4.48% | D | 7.8% |
| K38 borrowings (carrying) | $129.5B | $125.3B | R | 35.6 |
| K39 debt / RPO | 0.20× | 0.19× | D | 0.34× |
| K41 reported GPU utilization | 97.5% | 97.9% | R (call) | n/d |
| M04 GPUs per delivered MW | – | ~353 (>300k GPUs / 850 MW) | D | – |
| M08 not-yet-commenced DC lease commitments | $260B | $288B | R | – |

## Quarterly trends (calendar column = Oracle fiscal quarter ending 1 month earlier; ᴱ = estimated)

| ID | KPI | Unit | Q1'24 | Q2'24 | Q3'24 | Q4'24 | Q1'25 | Q2'25 | Q3'25 | Q4'25 | Q1'26 | Q2'26 | Q3'26† |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| K10 | Cash capex per new active MW | $M/MW | n/d | n/d | n/d | n/d | n/d | n/d | 77.3ᴱ | 30.1ᴱ | 46.6ᴱ | 56.9ᴱ | 33.5ᴱ |
| K13 | Trailing-4Q cash capex per trailing-4Q new active MW | $M/MW | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 46.4ᴱ | 39.0ᴱ |
| K16 | Useful life - GPU / server equipment | years | 5.00 | 5.00 | 6.00 | 6.00 | 6.00 | 6.00 | 6.00 | 6.00 | 6.00 | 6.00 | 6.00 |
| K17 | Depreciation yield (D&A x4 / avg in-service gross PP&E) | % | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 13.4 |
| K19 | Depreciation as % of revenue (company / total revenue) | % | 6.08 | 5.68 | 6.04 | 6.46 | 7.10 | 7.24 | 9.05 | 10.6 | 12.5 | 12.6 | 16.3 |
| K20 | RPO (ASC 606, EoP) | $B | 80.2 | 97.9 | 99.1 | 97.3 | 130 | 138 | 455 | 523 | 553 | 638 | 664 |
| K22 | RPO % recognized within 24 months | % | 60.5ᴱ | 57.0ᴱ | 57.0ᴱ | 58.5ᴱ | 51.0ᴱ | 53.5ᴱ | 22.5ᴱ | 25.0ᴱ | 27.5ᴱ | 29.0ᴱ | 31.5ᴱ |
| K23 | RPO / annualized quarterly total revenue | years | 1.51 | 1.71 | 1.86 | 1.73 | 2.30 | 2.17 | 7.63 | 8.15 | 8.04 | 8.31 | 8.58 |
| K27 | Number of customers >=10% of revenue | # | n/d | 0.00 | n/d | n/d | n/d | 0.00 | n/d | n/d | n/d | 0.00 | n/d |
| K28 | Weighted effective cost of debt (principal-weighted, EoP) - all facilities | % | n/d | 4.13 | n/d | n/d | n/d | 4.39 | n/d | n/d | n/d | 4.83 | n/d |
| K36 | Implied cash cost: interest x4 / avg principal | % | n/d | 4.02 | 3.93 | 4.00 | 3.86 | 4.14 | 4.02 | 4.24 | 3.89 | 4.36 | 4.48 |
| K38 | Total debt principal (EoP) | $B | 88.0 | 86.9 | 84.5 | 88.6 | 96.3 | 92.6 | 91.3 | 108 | 135 | 130 | 125 |
| K39 | Debt principal / RPO | x | 1.10 | 0.89 | 0.85 | 0.91 | 0.74 | 0.67 | 0.20 | 0.21 | 0.24 | 0.20 | 0.19 |
| K41 | Company-reported GPU utilization (physical) | % | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 97.5 | 97.9 |
| K42 | RPO % recognized within 12 months | % | 43.0 | 39.0 | 38.0 | 39.0 | 31.0 | 33.0 | 10.0 | 10.0 | 12.0 | 12.0 | 13.0 |
| K43 | Incremental annualized revenue per incremental MW (YoY) | $M/MW-yr | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 9.31ᴱ | 8.33ᴱ |
| M01 | OCI share of total revenue | % | 13.4 | 13.8 | 16.2 | 17.3 | 18.8 | 18.8 | 22.4 | 25.4 | 28.4 | 30.2 | 38.2 |
| M02 | OCI revenue growth YoY | % | n/d | n/d | n/d | n/d | 49.0 | 52.0 | 55.4 | 67.6 | 84.3 | 93.2 | 121 |
| M03 | Company capex / total revenue | % | 12.6 | 19.6 | 17.3 | 28.2 | 41.5 | 57.1 | 57.0 | 74.9 | 108 | 86.0 | 147 |
| M04 | GPUs per delivered MW (FQ1 FY27) | GPUs/MW | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 353 |
| M05 | Annual depreciation yield (trailing-4Q depreciation / avg in-service gross PP&E, FY) | % | n/d | n/d | n/d | n/d | n/d | 10.7 | n/d | n/d | n/d | 12.1 | n/d |
| M06 | Trailing-4Q OCI-attributed capex per $ of incremental annualized OCI revenue | x | n/d | n/d | n/d | n/d | 4.28ᴱ | 5.18ᴱ | 5.74ᴱ | 5.39ᴱ | 5.39ᴱ | 4.98ᴱ | 4.68ᴱ |
| M07 | Cumulative AI MW delivered since Jun-25 | MW | n/d | n/d | n/d | n/d | n/d | n/d | 110ᴱ | 510ᴱ | 910ᴱ | 1200ᴱ | 2050ᴱ |
| M08 | Not-yet-commenced DC lease commitments | $B | n/d | n/d | n/d | n/d | 48.4 | 43.4 | 99.8 | 248 | 261 | 260 | 288 |
| M09 | RPO increase QoQ | $B | 14.7 | 17.7 | 1.20 | -1.80 | 32.9 | 7.60 | 318 | 68.0 | 29.3 | 85.4 | 26.0 |


MW delivered (I09N):
- Q3'25 ~110ᴱ (residual)
- Q4'25 ~400 ("close to 400")
- Q1'26 >400
- Q2'26 ~290ᴱ (Q1 FY27 was "almost 3×")
- Q3'26† 850 (>300k GPUs; Abilene 618 MW, 75% of the site)

## Contract and RPO context (sourced in Notes_Gaps / Sources)
- **RPO jumps:** +$317B in FQ1 FY26, and +$85B in FQ4 FY26. The FY26 10-K attributes the increase to "certain significant cloud contracts".
- **Customer announcements in PRs and calls:**
  - FQ3 FY25: OpenAI, xAI, Meta, NVIDIA and AMD; the first Stargate contract was expected at that time.
  - FQ2 FY26: Meta and others.
  - The ~$300B OpenAI figure is media-reported only.
- **Off-balance-sheet commitments at Aug-26:** $288B of data-center lease commitments not yet commenced, plus $19B of cloud-infrastructure purchase commitments entered after May-26.



## Site cross-links (Peer KPIs alt-data)

Stargate capacity is split across peer packs so MW is not double-counted:
- **Abilene TX (Lancium / Crusoe-built, Oracle/OpenAI tenant):** counted under **Crusoe** in `peers/sites/` and `peers/private/` (C37). Evidenced-active IT L/B/H **200 / 421 / 421**; YE26 **421 / 632 / 843**; Grade Green (B1–4) / Yellow (B5–8). Do **not** add Abilene into Oracle active MW here.
- **Port Washington WI (Vantage "Lighthouse", Oracle/OpenAI):** Oracle's own key site in `peers/sites/` §6. Announced ~902 MW IT / 1.3 GW; **active 0 / 0 / 0**; YE26 0; handover from end-2027; ATC grid upgrade Dec-2026–Dec-2027. Grade **Yellow/Red**.

## Remaining n/d (flagged)

| Gap | Why | Watch |
|---|---|---|
| K01–K09, K14–K15 | No installed-MW or GPU fleet disclosure | Call commentary / future segment |
| OCI-only capex, D&A, PP&E, EBITDA | OCI not a segment; allocations are E | — |
| MW basis (IT vs gross) | ~353 GPUs/MW suggests facility/gross | Clarifying disclosure |
| K28 latest quarter | Debt tables annual (10-K) | FY27 10-K |
| Customer ≥10% (intra-year) | Only annual "none ≥10%" | — |
| No new filing since Oct 4 | EDGAR empty 2026-10-04→10-05 | FQ2 FY27 |


## Biggest gaps
1. **OCI is not a segment.** There is no OCI capex, D&A, PP&E, EBITDA, cost of revenue or customer data, so every OCI capex/MW and revenue/MW figure is an allocation estimate.
2. **No installed MW or GPU fleet count.** Oracle gives only MW delivered per quarter, from calls; some of those figures are lower bounds or estimates. This leaves K01–K09 and K14–K15 n/d.
3. **The MW basis is unclear.** The ~353 GPUs/MW density points to gross/facility MW rather than IT MW, which would understate $/MW against IT-basis peers.
4. **Debt tables are annual only, and RPO is company-wide with no OCI or customer split.**
5. **OCI list prices exist only for Aug–Oct 2026**, so there is no list-price history.
6. **The calendar mapping is offset by one month** (the Q2'26 column is Mar–May 2026).
