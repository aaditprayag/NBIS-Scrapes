# NBIS: depreciation life, EBITDA vs EBIT vs FCF, and what the price needs
*Nebius Skeptic (Red Team) for Aadit Prayag · draft for Wed Oct 7, 2026 · built Sun Oct 4, 2026 (ET) · workbook `03_depreciation.xlsx` (live formulas: blue = inputs, black = formulas; values verified by LibreOffice recalculation against a Python mirror, `dep/model.py`) · charts in `charts/`*

Tags: **R** Reported · **D** Derived · **E** Estimated. Base EV is $69.62B, with 287.66M shares at $242.81 (corrected Oct 4).

## 1. Starting point: Nebius's stated lives and peers
| Company | Server/GPU book life | Change | Tag / source |
|---|---|---|---|
| **Nebius** | **5 yrs from 1/1/26** (was 4) | −$43.0M Q2'26 D&A; −$86.1M 1H26; ~−$167.6M FY26 est. on YE25 assets | R: [Q2 FS Note 1](https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm), [20-F p.56](https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm), `filings/nebius_fact_base.md` |
| CoreWeave | 6 yrs (5→6 in 2023) | | R via [secondary](https://gpuinsights.net/server-useful-life-gpu-depreciation-2026/) |
| Alphabet / Oracle | 6 yrs | Oracle 5→6 in FY25 (−$733M opex) | R via [secondary](https://www.yieldtheory.app/research/michael-burry-ai-depreciation) |
| Microsoft | 2–6 yrs | 4→6 in FY23 | R via secondary |
| Meta | 5–5.5 yrs | →5.5 in Jan-25 (~−$2.9B) | R via secondary |
| Amazon | 6 yrs; a subset 5 | subset cut 6→5 in 2025 citing AI pace (+$1.4B D&A) | R via [CNBC](https://www.cnbc.com/2025/11/14/ai-gpu-depreciation-coreweave-nvidia-michael-burry.html) |
| IREN / Lambda | not found | | |

Control flag (R, 20-F FY25): there is a material weakness over fixed assets, including **depreciation start dates**, with remediation targeted by end-2026. FY26 is Deloitte's first audit (replacing Reanda).

## 2. Method (cohort/vintage)
- **Asset base at 6/30/26 (R):** $5,197.7M of server gross in service, depreciated over the remaining life (life − ~4 quarters average age, E). There is also $7,836.4M not yet in use, placed in service 50/50 in Q3/Q4'26 (E). Other in-service PP&E is $1,115.4M.
- **Capex:** FY26 $23.2B / FY27 $29.1B / FY28 $38.9B consensus (MarketScreener). It is flat thereafter (E; ≈1.3GW/yr at ~$30M/MW, in line with the ">1GW/yr" guide). 79% is GPU/server (E: management via a secondary transcript summary) and goes into service with a 2-quarter lag. Each quarter is a cohort (Blackwell → Vera Rubin → next gen) depreciated straight-line over **4 / 5 / 6 yrs** using a half-quarter convention. Non-GPU is depreciated over a 12-yr blended life (E).
- **EBITDA:** quarterly consensus Q3'26–Q4'27, FY28 annual consensus, then FY29+ = the FY28-implied EBITDA yield (27.6% of gross GPU in service) × the prior quarter's gross (E). A yield override input exists.
- **FCF** = EBITDA − capex − cash interest (coupons $313M/yr + 7% on funding beyond pro forma cash). It excludes prepayments and working capital (Q2 OCF of +$2.25B was prepayment-driven), taxes (NOLs), SBC and leases. **Book life changes EBIT, not FCF.**
- **Sanity check (D):** Q3'26 modeled D&A is $369M at 5 yrs vs $259.7M actual in Q2. The step-up comes from $7.8B of assets entering service.

## 3. Annual roll-up ($M; FY26 includes 1H actuals, and 1H D&A is restated to each life using 344.4 = 4 × 86.1)
| | FY26 | FY27 | FY28 | FY29 (E) | FY30 (E) |
|---|---|---|---|---|---|
| Revenue | 3,337 | 12,177 | 22,983 | 33,222 | 47,164 |
| Adj. EBITDA | 1,391 | 7,113 | 13,978 | 20,205 | 28,685 |
| D&A @ 4 / 5 / 6 yr | 1,675 / 1,382 / 1,186 | 6,291 / 5,131 / 4,358 | 12,845 / 10,481 / 8,905 | 20,293 / 17,091 / 14,523 | 27,617 / 23,396 / 20,324 |
| **EBIT @ 4 yr** | (284) | 822 | 1,133 | (88) | 1,068 |
| **EBIT @ 5 yr (Nebius policy)** | 9 | 1,982 | 3,497 | 3,114 | 5,289 |
| **EBIT @ 6 yr** | 205 | 2,755 | 5,073 | 5,682 | 8,361 |
| Capex | 23,207 | 29,135 | 38,885 | 38,885 | 38,885 |
| FCF (FY26 = H2 only) | (14,208) | (22,945) | (27,543) | (23,152) | (16,111) |
| External funding need beyond pro forma cash (YE, cumulative) | 0 | 22,655 | 50,198 | 73,350 | 89,461 |

Cross-check: consensus FCF is −$23.1B for FY27 and −$25.5B for FY28 (MarketScreener), close to the model. Consensus YE28 net debt of $50.7B matches the model's ~$50.2B funding need.
**Read:** Going from 5 to 4 yrs takes **$1.16B (FY27) and $2.36B (FY28) out of EBIT**, which is −$4.0 and −$8.2 per share pre-tax. At the old 4-yr policy, consensus EBITDA produces roughly **zero to $1.1B of EBIT every year through 2030**, before SBC (Q2 $102.5M) and interest (Q2 $119.1M).

## 4. Horizon views (quarter × 4 run-rate, $M)
| Mark | Quarter | EBITDA | EBIT 4y / 5y / 6y | FCF | Cum. external need |
|---|---|---|---|---|---|
| 6-month (Apr-27) | Q1'27 | 5,108 | 1,220 / 1,939 / 2,418 | (24,340) | 5,796 |
| 1-year (Oct-27) | Q3'27 | 7,532 | 430 / 1,739 / 2,612 | (22,734) | 17,362 |
| 3-year (Oct-29) | Q3'29 | 21,265 | 576 / 3,321 / 6,016 | (22,317) | 68,203 |

## 5. Reverse DCF: implied expectations (E; WACC 11%, g 3%, tax 25.8% terminal, steady-state replacement = D&A at the life)
- **Using consensus-implied asset yields (~28%) the model is negative under every life**: 5-yr ≈ −$323/share, 4-yr ≈ −$416, 6-yr ≈ −$265. The flat $38.9B/yr capex earns less than its replacement cost plus WACC, so more growth destroys value. Read this as "consensus numbers don't support the price," not as a price target.
- **Required steady-state EBITDA yield on gross GPU capital to justify the $69.6B EV:** 54.5% (4-yr) / **48.4% (5-yr)** / 44.0% (6-yr). That converts to **~$21.2M / $18.9M / $17.2M of revenue per MW** sustained for the asset life. Reference points: installed base ~$12M/MW (≈31%), Q2 deals $20–25M/MW (≈51–64%), Q3 short-term >$40M/MW (R, Q2 letter).
- **Value/share by FY29+ yield (5-yr life):** 30% → −$259 · 35% → −$126 · 40% → $7 · 45% → $141 · **50% → $274** · 55% → $407. At 4-yr: 50% → $181 (−$93 vs 5-yr). At 6-yr: 50% → $332 (+$58).
- **WACC sensitivity** at a 51.3% yield (Q2-deal pricing, 5-yr): 9% → $525 · 10% → $401 · 11% → $309 · 12% → $238. Required 5-yr yield: 41.9% / 45.1% / 48.4% / 51.9%.
- **Simple ROIC per MW (E; $30M/MW, 60.8% margin, gross capital):** after-tax at 5 yrs is 5.0% at $12M/MW, 9.5% at $15M and 17.1% at $20M. Breakeven against an 11% WACC is ≈ **$16M/MW** (≈$18M at 4 yrs, ≈$15M at 6 yrs).
**So what:** The stock is priced as if the whole fleet earns Q2-deal economics (~$19M/MW) for its full life. Consensus EBITDA implies roughly installed-base economics. The useful-life debate is worth ~$60–90/share at those yields, through replacement capex rather than reported EBIT.

## Charts
- `charts/03a_ebit_by_life_quarterly.png`: quarterly EBITDA vs EBIT at 4/5/6 yrs
- `charts/03b_annual_ebitda_ebit_fcf.png`: annual EBITDA / EBIT / FCF
- `charts/03c_reverse_dcf_yield.png`: value/share vs fleet yield by life, with the price line
- `charts/03d_da_by_cohort_5yr.png`: D&A by cohort at 5 yrs

## Caveats / gaps
- No GPU vintage roll-forward is disclosed, so cohort ages are estimates. Not-yet-in-use includes prepayments and DC work, and the 79% GPU share is applied to it too.
- FY29–30 have no consensus. They are extrapolated by yield, and the capex path is an input.
- Prepayments are excluded from FCF. They shift funding, not asset yield.
- WACC, g, $30M/MW and the 79% share are assumptions. Peer lives are from secondary summaries.
- An impairment toggle exists (default 0%). No impairment has been recorded (R, Q2 FS Note 7).
