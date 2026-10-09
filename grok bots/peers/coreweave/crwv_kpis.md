# CoreWeave (CRWV): peer KPI history, Q1'24 to Q2'26, on Nebius definitions

_Built Sun Oct 4, 2026 (ET); polished Mon Oct 5, 2026 (ET) — Notes_Gaps A/B/C, K02b≡K02, EDGAR scan (no new ops). Workbook: `crwv_kpis.xlsx` (same folder). It has tabs KPIs, Inputs, Debt_Facilities, Sources, Definitions, Notes_Gaps and Comparison, and the formulas are live. Every input carries a tag and a source (filing, section, URL).
Tags: **R** = reported, **D** = derived (formula shown), **E** = estimated (relies on an assumption), **n/d** = not disclosed._

## Definitions mirrored from the Nebius work
The KPI formulas come from `model/nebius_model.xlsx` (KPIs, Capex_DA and Hist tabs). The MW basis and GPU-density method come from `capacity/gpu_mw_capex.md/.xlsx` and `parts/comps_rev_per_mw.csv`. The cost-of-debt method comes from `parts/comps_sources.csv`. Concentration and RPO follow `filings/nebius_filings.xlsx`. The Definitions tab gives the exact Nebius row behind each KPI.

- **Rev/MW** = quarterly revenue × 4 ÷ MW, in $M per MW-year. It is computed on period-end active MW, average active MW, and contracted MW.
- **Rev/GPU** = revenue × 4000 ÷ average GPUs. **$/GPU-hr** = revenue ÷ (average GPUs × hours).
- **Implied utilization** = $/GPU-hr ÷ fleet-weighted on-demand list price. Nebius itself flags that this ratio mixes discounting with utilization.
- **Capex/MW** = cash capex ÷ new active MW. Variants: in-service additions per new MW (this matches CRWV's own capex definition), trailing 4 quarters, and net or gross PP&E per active MW.
- **Depreciation yield** = D&A × 4 ÷ average in-service gross PP&E.
- **Cost of debt** = principal-weighted effective rate from the Note 10 debt table.
- **MW basis:** CRWV never says whether "active power" is IT or gross MW. At YE24 it ran >250k GPUs on ~360 MW, about 694 GPUs/MW. That is close to Nebius's Hopper base of 713 per IT-MW and well above 548 per gross MW, so CRWV's MW most likely means IT load. That is the same basis as Nebius's "active" MW.

## Latest quarter (Q2'26) and trend

| KPI | Q2'26 | Trend Q1'24 → Q2'26 |
|---|---|---|
| Rev / active MW, period-end (D) | **$6.9M** | $8.3M (Q4'24) → peak $10.3M (Q2'25) → $7.4M (Q4'25) → $8.3M (Q1'26) → $6.9M |
| Rev / active MW, avg MW (D; 2024 = E) | **$8.2M** | ~$7.1M (Q1'24 E) → $9.2M (Q4'24 E) → $10.9M peak (Q2'25) → $8.7–9.0M (Q4'25–Q1'26) → $8.2M |
| Rev / contracted MW (D) | $2.8M | $2.3M (Q4'24) → $1.9M (Q3'25) → $2.8M |
| Adj. EBITDA / active MW, avg (D) | $4.8M | ~$6.0M (2024 E) → $6.8M (Q2'25) → ~$5.0M (Q4'25–Q2'26) |
| Rev / GPU, avg (E) | $14.5K/GPU-yr ($1.66/GPU-hr) | $9.7K (Q1'24) → $13.3K (Q4'24) → $16.4K (Q2'25) → $14.5K |
| Implied utilization (E) | ~18% | 18% → 24% (2024) → 27% (Q2'25) → 18%. This is a price-realization index, not physical utilization |
| Cash capex / new active MW (D) | $12.8M (single quarter) | Noisy: $13–51M per quarter. Trailing 4Q: $21.0M (FY25) → $28.6M → $20.0M. FY24 = $30.0M |
| In-service additions / new MW (D) | $18.7M | $15–59M per quarter. FY24 = $28.7M |
| Net PP&E / active MW (D) | $31.2M | $33.1M (YE24) → $36.4M (Q1'26) → $31.2M |
| Useful life, tech equipment (R) | 6 yrs | Unchanged throughout. The last change was 5→6 yrs, effective 1-Jan-2023 |
| Depreciation yield (D) | 15.5% (implied life ~6.5 yrs) | 16.7% (Q1'25) → 15.5%. D&A/revenue rose from 42% to 54% |
| RPO (R) | $103.7B; 41% due within 24 months | $15.1B (YE24) → $60.7B (YE25) → $103.7B. The share due within 24 months fell from 54% to 41%. Backlog is $104.2B, excluding >$25B signed in early Q3 |
| RPO / annualized revenue (D) | 10.1x | 5.0x (Q4'24) → 10.1x |
| Top customer % of revenue (R) | 36% | 33% → 59–72% (Q2'24–Q3'25) → 45% → 36% |
| Customers ≥10% (R) | 3: 36% / 26% / 10%, 72% combined | 3 → 1 (all of 2025) → 3 |
| Weighted effective cost of debt (D) | **8.4%** | 14.7% (YE23) → 12.0% (YE24) → 9.1% (YE25) → 8.4%. The company's own figures are 14.9 / 12.2 / 9.0 / 8.3%, and 7.8% pro forma |
| — DDTL / senior notes / converts / OEM | 9.6% / 9.8% / 2.0% / 10.7% | DDTL fell from 12.4% to 9.6%. Notes stayed ~10% and converts ~2%. The DDTL share of principal fell from 73% to 38% |
| Implied cash cost (contractual interest ÷ avg principal) (D) | 7.8% | 11.1% (Q1'25) → 7.8% |

## Key findings
1. **Revenue per MW peaked in mid-2025 and has been diluted since.**
   - On average MW it was ~$10–11M/MW in Q2–Q3'25 and is ~$8.2M now. CRWV added ~650 MW in H1'26 (850 → 1,500 MW), and new capacity bills with a lag.
   - The Sep-26 convert deck says prices rose ~70% across SKUs in July 2026. It also cites ~$40M/MW on short-dated deals. Both are marginal figures and point to rev/MW recovering in H2'26.
   - CRWV (~$7–8M/MW) is still well below Nebius's quoted $20–25M ACV/MW on new deals. Nebius's own $13.7M/MW is computed on stale YE25 MW.
2. **Capex per MW stays at about $20–30M per incremental IT-MW** (trailing 4Q and FY24). Single quarters swing between $13M and $51M because GPU purchases lead energization.
   - The level sits below Nebius's GB300 GPU+network model (~$40M/IT-MW, capacity/gpu_mw_capex.md) and inside the $20.5–34.5M per connected MW implied by Nebius's 2026 guide. CRWV leases its data-center shells, so its capex is mostly GPUs and network.
3. **Depreciation is the clearest point of contrast with Nebius.**
   - CRWV depreciates tech equipment over 6 years, versus Nebius's 5 years from 2026 (4 years before).
   - The depreciation yield of ~15.5–17% confirms a ~6-year blended life. D&A is now 54% of revenue.
4. **The backlog is getting larger and longer-dated.**
   - RPO grew ~7x in 6 quarters and now covers ~10 years of current revenue.
   - The share due within 24 months fell from 54% to 41%, so near-term conversion is a smaller part of the book.
5. **Concentration has fallen sharply.**
   - Microsoft-type exposure dropped from ~70% of revenue (all of 2025) to 36% in Q2'26. Two other customers are now at 26% and 10%.
   - CRWV is still more concentrated than Nebius (top customer 24%, top 3 = 59%).
6. **Funding cost fell ~360bp in 18 months** (12.0% → 8.4%).
   - Mix drove most of it: 2% converts, investment-grade non-recourse DDTL 4.0 at 7%, and paydown of the 15% DDTL 1.0.
   - The marginal secured cost is still ~9–10%: senior notes, DDTL 5.0 at 9%, and DDTL 5.5 at SOFR+5.50% in Aug-26.
   - Nebius's blended cost is ~5.1%, about 330bp cheaper, but it is mostly converts.

## Caveats and biggest gaps
- **GPU counts:** CRWV has disclosed none since YE24. Rev/GPU, $/GPU-hr and implied utilization after YE24 are **Estimated** as MW × Nebius base densities: GB200 538/IT-MW for 2025 additions and GB300 491 for 2026. The density choice alone moves them ±15–20%.
- **2024 interim data:** CRWV was pre-IPO, so the 2024 quarterly P&L exists but there is no MW, balance sheet or debt table for Q1–Q3'24. 2024 rev/MW uses MW interpolated between YE23 and YE24 (Estimated). Depreciation yield and cost of debt are n/d for Q1–Q3'24.
- **MW figures:** they are often "more than" or "approximately" (850+, 1 GW+), and IT vs gross is unstated. Single-quarter ΔMW, and therefore capex/MW, is noisy.
- **Implied utilization:** it blends contract discounts (reserved and committed pricing sits well below on-demand list) with idle capacity. List prices are flat H100 $6.155 and GB200 $10.50, from the GPU Market agent's series. Do not read the ratio as physical utilization.
- **Debt rates:** the effective rates are rounded to whole percentages and ranges are entered at their midpoints (±0.5pp). Leases (~$10B+ operating) are excluded.
- **Quarterly Q4 values:** Q4 concentration and Q4 interest are FY minus 9M. For concentration this carries ±2pp of rounding.

## Reuse
- Copy the workbook, keep the row IDs (I01–I27, P01–P03, K01–K40), and replace the Inputs and Debt_Facilities tabs for IREN, Oracle OCI or private peers.
- The Comparison tab already links CRWV Q2'26 live and holds cited Nebius values where the Nebius workspace has them. The IREN, OCI and private-peer columns are blank.
- IREN note: its fiscal year ends Jun-30, so map its fiscal quarters to calendar columns.


## Remaining n/d (Oct 5 polish — full detail in Notes_Gaps)

| ID | Item |
|---|---|
| A1 | Q1–Q3'24 MW / BS / debt tables (pre-IPO) |
| A2 | GPU counts after YE24 |
| A3 | IT vs gross MW basis (unstated; density leans IT) |
| A4 | Physical GPU utilization (K09 = price-realization only) |
| A5 | Separate billed-IT MW series (→ K02b≡K02) |
| A6 | Q1–Q3'24 RPO / concentration |
| A7 | Adj. EBITDA segment bridge (N/A — single perimeter) |

EDGAR through 2026-10-05: latest ops filing still Q2'26 10-Q; Oct 2–3 = Form 4/144 only.
