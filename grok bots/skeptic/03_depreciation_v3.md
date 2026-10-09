# 03 Depreciation & unit economics, v2 (NBIS)
*Nebius Skeptic for Aadit Prayag · v2 compiled Sun Oct 4, 2026, ~6:45 PM ET · due Wed Oct 7 · v1 kept as 03_depreciation_v1.xlsx / .md*

Workbook: `03_depreciation.xlsx`, with live formulas (blue = inputs, black = formulas). Tabs: Assumptions · Sites · MW · Microsoft · Rev_MW · Model · Annual · Horizons · Reverse_DCF · Cohorts · Charts. Charts: `charts/v2_ebit_by_life.png`, `v2_reverse_dcf.png`, `v2_fy27_revenue.png`. Builder: `dep_v2/build.py`; scenario runner: `dep_v2/scen.py` → `scen.json`.

## What changed vs v1
- **Shares/EV:** 308.69M economic basic shares (incl. the 21.07M NVIDIA warrant), EV $74.73B, 381.5M FD (Filings reversal).
- **MW input:** Capacity's `Quarterly MW` energized low/base/high by site and quarter. Sites with no stated basis are treated as gross; IT sites (Vineland, AIB) are ×1.25.
- **Billing:** energized → billing lag of 1 quarter (Filings: the "few months" lag runs from connected power to revenue; 2 quarters tested). Capex → energized lag stays at 2 quarters.
- **Capex per IT MW:** Capacity base case. GB300 is $40.2M GPU/net plus $12.5M shell (owned), or $41.3M colo all-in (leased). Vera Rubin capex runs from Q3'27.
- **Owned share of capex:** owned sites' share of each fiscal year's energized additions (Capacity tracker): 0% in 2026, 44% in 2027, 88% in 2028. All-leased and all-owned variants are run as well.
- **Opex:** power at $139/MWh (range $93–185, ≈$1.1M per gross MW-yr; Bloom is a power purchase, i.e. opex). Leased rent is $2.25M per IT MW-yr. Other opex is calibrated to the Q2 segment margin of 49.7%.
- **Microsoft tranche sheet:** includes an SLA haircut and a per-tranche termination toggle. Scenarios: letter, Capacity base, Capacity bear.
- **Prepayments:** 70% of deals carry prepayments of 55% of capex, recognized over 20 quarters. Includes the $5.98B deferred-revenue opening balance; can be toggled off.
- **Off-balance-sheet EV variant:** adds $5.32B of energy commitments, $12.05B of signed-not-commenced leases and $1.63B of on-balance-sheet leases, for $93.7B.
- **Residual value:** provisional A100-proxy index × 60% utilization.

## Headline results (base: Capacity base MW, 1-qtr billing lag, 5-yr policy life)
**Annual ($M; FY26 = 1H actual + 2H model; consensus through FY28, extrapolated after)**

| | FY26 | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|
| Revenue (cons. → ext.) | 3,337 | 12,177 | 22,983 | 35,046 | 49,509 |
| Adj. EBITDA | 1,392 | 7,113 | 13,978 | 21,315 | 30,111 |
| Capex | 23,207 | 29,135 | 38,885 | 38,885 | 38,885 |
| EBIT margin, 4-yr | −8.2% | 7.4% | 5.6% | 5.1% | 8.0% |
| EBIT margin, 5-yr (policy) | 0.6% | 17.0% | 16.1% | 14.0% | 16.2% |
| EBIT margin, 6-yr | 4.8% | 23.5% | 23.1% | 21.3% | 22.1% |
| Pts added by 4→5 change | 8.8 | **9.6** | **10.5** | 9.0 | 8.2 |
| FCF (prepayments on; FY26 = H2) | −9,436 | −14,944 | −17,568 | −14,519 | −9,603 |
| FCF ex-prepayments | −14,207 | −22,403 | −26,143 | −20,025 | −12,114 |
| Liquidity at YE (start $17.2B, Filings) | 7,764 | −7,179 | −24,748 | −39,267 | −48,869 |
| Bottom-up revenue from billing MW | 2,619 | **9,937** | 25,286 | 42,473 | 60,400 |

*Note: the FY27 revenue row is MarketScreener's quarterly path ($12.18B). The FY27 consensus total is $12.3B.*

**Horizons (quarter at the horizon):** at 6 months (Q1'27), revenue is $2.18B, EBITDA $1.28B, EBIT at 5 yrs $0.50B (4 yrs $0.32B), and liquidity $3.8B. At 1 year (Q3'27): $3.30B / $1.88B / $0.46B (4 yrs $0.13B), and liquidity is −$3.8B (negative from Q2'27). At 3 years (Q3'29): $9.24B / $5.62B / $1.31B (4 yrs $0.63B), with billing MW at ~2,294 gross and liquidity at −$36.1B. *(Values from the Horizons tab.)*

### 1) Governance leg: EBIT margin points from the life change
In FY27 the life change lifts EBIT margin from 7.4% at 4 yrs to 17.0% at 5 yrs (+9.6 pts). In FY28 it goes from 5.6% to 16.1% (+10.5 pts). The 20–30% target was framed on 4 yrs (May-25 letter) and reaffirmed without that caveat on Feb 12, 2026, the same call that moved servers to 5 yrs (Guidance_Log, `filings/nebius_filings.xlsx`). Even with the change, consensus sits below the target range through FY28. Earnings-quality dollars: the life change contributes ~$43M a quarter ($167.6M FY26 guided; ~14% of Q2 D&A, since Q2 would have been $302.7M). Capitalized interest was $15.7M in Q1 and $40.2M in Q2, vs $107.5M of gross convert interest.

### 2) Reverse DCF on billing capital (Modeling pushback)
| Life | Implied FY29+ yield (EV $74.7B) | Implied (EV incl. off-BS $93.7B) | FY28 cons. yield, billing basis | v1 in-service basis | DCF EV at 31% | $/share gap |
|---|---|---|---|---|---|---|
| 4 yrs | 52.1% | 54.6% | 31.0% | 27.2% | −$82.7B | −$510 |
| 5 yrs | 47.9% | 50.5% | 31.0% | 27.2% | −$51.7B | −$410 |
| 6 yrs | 45.2% | 47.7% | 31.0% | 27.2% | −$31.1B | −$343 |

**Restated "negative under every life":** the result still holds on the billing basis. The billing basis raises the FY28 yield from 27% to 31%, which is not enough. Sensitivities:
- FY29–30 capex −50%/yr: implied 51/47/44%; EV at 31% is −$48B / −$23B / −$6B.
- Same with WACC 9%: 45/41/38%; 6 yrs turns slightly positive at +$12B.
- 2-quarter billing lag: FY28 yield 35.9%; implied 54/50/47%.
- All-leased fleet: FY28 yield 28.9%, implied 41% at 5 yrs. All-owned: 33.0% vs 49%.

The off-balance-sheet EV variant double-counts against an after-rent EBITDA, so treat it as a bound only. Per billing MW, FY28 consensus equals **$11.0M of revenue and $6.7M of EBITDA per gross MW**, below the $12M installed-book ACV. The billing-MW row from Capacity lands Tue and will replace the energized+lag proxy.

### 3) Leg 3: FY27 revenue on the quarterly billing-MW path (vs $12.3B consensus)
Method: billing MW = energized MW shifted by the lag. Non-Microsoft billing MW is priced at $12M for the Q2'26 base and $22.5M for increments, plus the Oct 1 hike on the ~4% on-demand share. Microsoft revenue comes from the tranche sheet.

| Scenario | FY27 rev | Gap | $/share (6.08x) |
|---|---|---|---|
| Capacity low | $7.01B | −$5.29B | −$104 |
| **Capacity base** | **$9.94B** | **−$2.36B** | **−$46** |
| Capacity high | $12.86B | +$0.56B | +$11 |
| Base, 2-qtr lag | $7.62B | −$4.68B | −$92 |
| Base, Microsoft bear (T5–T9 Q3'27–Q1'28) | $9.36B | −$2.94B | −$58 |
| Base, Microsoft T5–T9 cancelled | $8.78B | −$3.52B | −$69 |
| Base, Microsoft "all delivered" | $10.71B | −$1.59B | −$31 |
| Base + 25 MW/qtr unseen sites | $11.35B | −$0.95B | −$19 |
| Low + 2q lag + Microsoft bear | $5.16B | −$7.14B | −$141 |

Closing the base-case gap needs **~42 gross MW/qtr of unseen sites** (Oklahoma and Minnesota are candidates) from Q4'26. Each 1 MW/qtr adds ~$56.7M of FY27 revenue. FY28 bottom-up ($25.3B) is above consensus ($23.0B), so this is a 2027 timing call.

### 4) Microsoft / Vineland
Equal split: $96.6M/qtr per tranche. FY27 Microsoft revenue is $3.48B (letter) / $2.71B (base) / $2.13B (bear) / $1.55B (T5–T9 cancelled = 4-tranche genset ceiling, ~$1.55B/yr). PV at 11%: $13.4B / $12.3B / $11.8B / $5.8B. Q2 check: T1+T2 at an equal split is $193M/qtr vs the ~$140M top customer, so the conflict remains open.

### 5) Cohort unit economics (per gross MW; Cohorts tab)
| Cohort (capex $M/gross MW) | Life | Project IRR owned / leased | Payback (yrs) o / l | Breakeven EBITDA margin o / l | Contract margin o / l | Equity IRR, prepaid deal o / l |
|---|---|---|---|---|---|---|
| Installed book $12M (Hopper: 28.7 o / 21.5 l) | 4 / 5 / 6 | 10% / 15% / 15% · 4% / 12% / 12% | 3.4–3.7 · 3.6 | 66/57/56% · 58/49/47% | 65% · 50% | 23/30/29% · 11/25/24% |
| Q2 deals $20M (GB300: 42.2 / 33.0) | 4 / 5 / 6 | 16 / 22 / 22% · 16 / 23 / 24% | 3.1 · 2.8 | 60/51/51% · 53/45/43% | 68% · 59% | 36/42/42% · 39/47/48% |
| Q2 deals $25M | 4 / 5 / 6 | 28 / 32 / 32% · 31 / 38 / 38% | 2.4 · 2.1 | 48/41/41% · 43/36/35% | 69% · 62% | 59/63/63% · 69/75/76% |
| Short-term >$40M (1-yr, then re-rent) | 4 / 5 / 6 | 11 / 13 / 14% · 8 / 10 / 11% | 3.2–3.5 · 2.5 | 69/66/65% · 61/58/55% | 71% · 67% | 17/18/19% · 13/16/18% |
| Microsoft $11.6M (300 MW gross; leased) | 4 / 5 / 6 | n/m / −5% / −6% | never | 92/77/75% | 49% | n/m / −18% / n/m |
| Microsoft $17.4M (200 MW basis; leased) | 4 / 5 / 6 | 8 / 15 / 16% | 3.3 | 61/51/50% | 57% | prepayment > capex (no equity at risk) |

Equity return treats a prepayment (55% of capex; Microsoft 40% of TCV) as customer-funded capex, repaid through lower cash revenue over the term. The blended book is 70% prepaid. A short-term deal is modeled as a full new build, which is pessimistic if it reuses installed GPUs. Owned vs leased: leased capex is lower, which offsets rent, so at Q2-deal pricing leased sites match or beat owned on IRR.

## Caveats
- Residual index is PLACEHOLDER (no A100 decay curve yet).
- Capacity's billing-MW row is due Tue.
- Microsoft tranche sizes are not disclosed.
- Capex is extrapolated flat after FY28 (a fade case is shown).
- FY26 H1 D&A is actual; in-service additions use the 6/30 not-in-use balance ($7.84B) for Q3/Q4'26.
- Tax at 25.8% on positive EBIT only.
- No Q3 balance sheet yet.
