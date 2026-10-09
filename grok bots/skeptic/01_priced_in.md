# NBIS: what's priced in (Red Team / skeptic snapshot), v3
**As of:** close Fri Oct 2, 2026 (compiled Sun Oct 4, 2026, ET) · **Analyst:** Nebius Skeptic (bear side)
**Tags:** **R** = Reported · **D** = Derived (math shown) · **E** = Estimated (method + source). Full row-level sourcing is in `01_priced_in.xlsx` (sheets: Priced_in, Price_history, Quarterly_consensus).

> **Staleness flag:** The latest filed balance sheet is **Jun 30, 2026** (Q2'26 PR/6-K, Aug 12). Three things happened after it and are **not yet in any filed balance sheet**: the $775M secured loan (Jul 17), the **$5.75B Aug-24 converts**, and the $800M note-for-~15.8M-share exchange. Q3'26 results (est. **Nov 10, 2026**) will be the first to show them. Pro forma cash below does **not** include Q3 operating cash flow or capex, because neither has been disclosed. **Q2 capex alone was $5.66B**, so true Sep-30 cash is probably materially below the pro forma figure (Q2 OCF was +$2.25B, but that came from prepayments).

> **v3 correction (Oct 4, ~6:45 PM ET; v2 kept as 01_priced_in_v2.*):** Filings has reversed the share-count correction, with high confidence. The 271.86M figure is *issued* shares and does **not** include the 21.07M NVIDIA pre-funded warrant, so the warrant is added back. Economic basic shares are **308.69M** (271.86 + 15.77 exchange + 21.07 warrant), market cap is **$74.95B** and base EV is **$74.73B** (v2: 287.66M / $69.85B / $69.62B). Fully diluted at $242.81 is **381.5M** (Filings CapStructure). The v2 fixes that stand: CRWV is compared ex-leases ($79.52B) and convert principal is $13.43B. Sources: [Filings fact base](/workspace/nebius/filings/nebius_fact_base.md), `/workspace/nebius/filings/nebius_filings.xlsx` (CapStructure, SOTP) and team comps (`/workspace/nebius/parts/comps.xlsx`).

## 1. Price, shares, market cap
| Item | Value | Tag | As-of | Source |
|---|---|---|---|---|
| Close | **$242.81** (+4.53%) | R | 2026-10-02 | [StockAnalysis](https://stockanalysis.com/stocks/nbis/history/) |
| Issued & outstanding (A+B) | 271,855,218 (238.40M A + 33.46M B) | R | 2026-06-30 | [Q2 PR](https://assets.nebius.com/assets/72a8c258-bbb7-4df7-ab9d-8698f6cb88fc/PR.pdf) |
| Aug-24 exchange shares | 15.77M (PR: 'approximately 15.8 million') | R | 2026-08-24 | [6-K ex99.1](https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex99-1.htm) |
| NVIDIA pre-funded warrant | 21,065,936 sh @ $0.0001: **NOT included in the 271.86M issued** (Filings v3 reversal), so it is added to the economic basic count; no exercise reported through 10/1 | R | 2026-06-30 | [Filings fact base](/workspace/nebius/filings/nebius_fact_base.md) |
| ATM sales after Jun 30 | none disclosed (12,270,507 sh remain of 25M program; Q2 sold 12.73M @ $223.60) | R | through 2026-10-01 | [Filings fact base](/workspace/nebius/filings/nebius_fact_base.md) |
| **Economic basic shares** | **308.69M** = 271.86 + 15.77 + 21.07 | D | 2026-10-02 | above |
| Converts outstanding (orig. principal) | **$13.43B**: 2029/2031 ~$87.5M each @ $51.45 · 2030 1.00% & 2032 2.75% ~$1.58B each @ $138.75 · 2031 1.25% $2.59B @ $183.22 · 2033 2.625% $1.75B @ $180.31 · 0.50% 2/2030 $3.45B @ $313.46 · 4.50% 2/2034 $2.30B @ $324.65. **No capped calls.** | D | 2026-10-02 | [Filings fact base](/workspace/nebius/filings/nebius_fact_base.md) |
| ITM convert shares @ $242.81 | 50.02M (1.70M + 1.70M + 11.40M + 11.40M + 14.12M + 9.71M) | D | 2026-10-02 | above |
| OTM convert shares | 18.09M (11.01M + 7.08M) | D | 2026-10-02 | above |
| Options / RSUs | 6.74M options (WAEP $88.65), so 4.28M net under the treasury-stock method / 6.23M RSUs | R / D | 2026-06-30 | [Filings fact base](/workspace/nebius/filings/nebius_fact_base.md) |
| **Fully diluted @ $242.81** | **381.5M** = 308.69 + RSUs 6.23 + options 4.28 (TSM) + ITM converts 50.02 + ATM capacity 12.27 | D (Filings CapStructure) | 2026-10-02 | `nebius_filings.xlsx` CapStructure |
| All converts regardless of price | 399.6M (+ OTM 18.09M). Under net-share settlement: 349.9M, which needs ~$7.7B of cash. Default settlement is shares, with **no capped calls** | D | 2026-10-02 | CapStructure |
| Issuer calls | Possible above 130% of the conversion price; the first is Dec 2026 (leftover 2029s). Dilution steps down at $183, $139 and $51 | R | | Indenture_Terms |
| Mkt cap, headline aggregator | $66.55B (274.10M sh) | R (agg.) | 2026-10-04 | [StockAnalysis](https://stockanalysis.com/stocks/nbis/statistics/) |
| Mkt cap, company count | $66.01B = 271.86M × $242.81 | D | | |
| Markets Insider | $64.13B (counts Class A only, so it is understated) | R (agg.), flagged | | [BI](https://markets.businessinsider.com/stocks/nbis-stock) |
| **Mkt cap, economic basic (base)** | **$74.95B** = 308.69M × $242.81 | D | | |
| Mkt cap, fully diluted | $92.63B = 381.5M × $242.81 | D | | |

Assumption: the 2033-note greenshoe ($262.5M) was **not** exercised. I inferred this because H1'26 convert proceeds were exactly $4,337.5M and Q2 shows no convert proceeds.

## 2. Enterprise value
| Component | USD M | Tag | As-of |
|---|---|---|---|
| Cash & equivalents | 8,042.1 | R | 2026-06-30 ([PR](https://assets.nebius.com/assets/72a8c258-bbb7-4df7-ab9d-8698f6cb88fc/PR.pdf)) |
| Restricted cash (excluded) | 1,056.0 | R | 2026-06-30 |
| Debt current + non-current (mostly converts, carrying) | 46.7 + 8,499.0 = 8,545.7 | R | 2026-06-30 |
| + Secured GPU facility | +775 (SOFR+2.50%, due 10/2030; assumed fully drawn) | R | 2026-07-17 ([PR](https://assets.nebius.com/assets/cf774b13-960d-4c5e-a78f-a0278799cca1/17072026%C2%A0Nebius%20raises%20%24775%20million%20in%20first%20secured%20debt%20financing%20to%20accelerate%20global%20buildout.pdf)) |
| + Aug-24 converts | +5,750 ($3.45B 0.50% 2030 + $2.30B 4.50% 2034; accrete to 110%/125%) | R | 2026-08-24 ([6-K](https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex99-1.htm)) |
| − Exchanged 2029/2031 notes | −800 (at face; carrying value not disclosed) | R | 2026-08-24 |
| **Pro forma debt (ex-leases)** | **14,270.7** | D | |
| **Pro forma cash** | **14,497.1** = 8,042.1 + 775 + ~5,680 (company's net-proceeds estimate) | E. **Stale: excludes Q3 burn** (Q2 capex was $5,657M) | |
| Operating leases (non-curr / curr) | 1,510.4 / 114.9 | R | 2026-06-30 |
| Deferred revenue (prepayments) | 5,975.2 | R | 2026-06-30 |

| EV version | USD B | Math |
|---|---|---|
| Headline (StockAnalysis) | 68.71 | 66.55 + 10.20 (debt incl. leases) − 8.04; balance sheet predates the Aug financing |
| **Base (pro forma, ex-leases)** | **74.73** | 74.95 + 14.27 − 14.50 |
| Base + started leases | 76.35 | + 1.625 |
| Incl. off-BS commitments (bound) | 93.73 | + energy commitments 5.32 + leases signed not commenced 12.05 + started leases 1.63 (undiscounted) |
| Fully diluted (ITM converts as equity, ATM proceeds as cash) | 81.75 | 92.63 + (14.27 − 7.675 ITM face) − (14.50 + 2.98 ATM proceeds) |
| Bear view (prepayments as debt-like) | 80.70 | 74.73 + 5.975 |
| EV per contracted MW | $21.4M/MW on >3.5GW (D); $14.9M on 5GW. Adding signed-not-started ($12.05B) and started ($1.63B) leases against contracted MW: $25.3M / $17.7M. CRWV: ~$21.5M (ex-lease, 3.7GW) | |
| EV per energized MW (started leases only) | ~$413M per gross MW = (74.73 + 1.63) / ~185 MW energized at 6/30 (Capacity Quarterly MW, Q2'26 base) | |
| SOTP non-core stakes (Filings) | ~$1.6B carrying value, ~$4.3B upper bound, about $4–11 per fully diluted share | `nebius_filings.xlsx` SOTP |

**So what:** On v3 math, base EV ($74.7B) is ~$6.0B above the aggregator headline ($68.7B). The gap is the 21.07M warrant shares (~$5.1B) plus the 15.77M exchange shares (~$3.8B), partly offset by net cash after the Aug financing. Diluted exposure is $81.8B once ITM converts, options, RSUs and the ATM are counted. Q3 burn is not captured.

## 3. Consensus vs guidance
| Metric (USD B) | FY2026 | FY2027 | FY2028 | Source / tag |
|---|---|---|---|---|
| Revenue (mean) | 3.34 (20 analysts; 3.00–3.59) | 12.30 (21; **8.81–17.99**) | 22.98 | E: [Yahoo](https://finance.yahoo.com/quote/NBIS/analysis/), [MarketScreener](https://www.marketscreener.com/quote/stock/NEBIUS-GROUP-8037501/finances/) |
| EBITDA (adj., MarketScreener) | 1.387 (41.6%) | 6.728 (54.7%); other sources ~6.36 | 13.978 (60.8%) | E: MarketScreener; analyst count for EBITDA not shown (~10 on page) |
| EPS (USD) | −2.28 (8; −6.41 to +0.49) | −5.64 (9; −15.31 to +5.27) | +3.43 | E: Yahoo / MarketScreener |
| Capex | 23.2 | 29.1 | 38.9 | E: MarketScreener |
| FCF | −17.5 | −23.1 | −25.5 | E: MarketScreener |
| Net debt (YE) | 13.7 | 34.3 | 50.7 | E: MarketScreener |
| FY2029–2030 | not disclosed/not found | | | |

**Quarterly consensus (E, MarketScreener; Yahoo agrees on revenue):** Q3'26 rev $904M / EBITDA $355M · Q4'26 $1,452M / $671M · Q1'27 $2,175M / $1,277M · Q2'27 $2,646M / $1,581M (Q3'27 $3,304M / $1,883M; Q4'27 $4,052M / $2,372M).

**Company guidance (R, Aug 12, 2026, [letter](https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf)):** FY26 revenue $3.0–3.4B · YE26 ARR $7–9B · ~40% adj. EBITDA margin · capex $20–25B · YE26 connected power 800MW–1GW · contracted power **5GW** (raised from >4GW) · >$9B prepayments in 2026 · >1GW/yr deployments from 2027 · >$40B commitments from investment-grade customers (MSFT, META). Q2'26 actuals (R, [Filings fact base](/workspace/nebius/filings/nebius_fact_base.md)): revenue $582.3M · adj. EBITDA $236.2M (41%) · D&A $259.7M · SBC $102.5M · interest $119.1M · capex $5,657M · OCF +$2,246M (prepayment-driven; deferred revenue +$1.2B) · PP&E $13,045M · deferred revenue $5,975M · RPO $37.5B (36% within 24 months) · top-3 customers 24/21/14% of revenue · ARR $3.0B. Microsoft: 9 tranches, $17.39B owed regardless of use, ~$6.96B upfront, billed through Oct-2031. The letter says all tranches are delivered, but **that conflicts** with a top customer of only ~$140M/qtr vs the ~$725–870M/qtr full delivery implies (open). Meta: #1 ~$2.88B/5yr; #2 $12B firm on Vera Rubin, tranches from early 2027, plus up to $15B backstop (likely outside RPO). ~70% of deals carry prepayments covering 50–60% of related capex. ACV/MW: 2026 base ~$12M, Q2 deals >$20M ($20–25M), Q3 short-term deals >$40M. Server useful life was extended from 4 to **5 years** on 1/1/26. That cut D&A by $43.0M in Q2 and ~$167.6M est. for FY26 (company figure). Pricing caveat: the installed base runs at ~$12M/MW. The new $20–25M/MW deals are priced on forecast costs and include capacity not yet built. Older-gen GPU pricing rose >30% QoQ in Q2.

## 4. Sell-side price targets
| Aggregator | n | Low | Mean | Median | High | As-of |
|---|---|---|---|---|---|---|
| S&P Global ([StockAnalysis](https://stockanalysis.com/stocks/nbis/forecast/)) | 19 | $84 | **$283.58** | $300 | $415 | 2026-09-25 |
| [MarketBeat](https://www.marketbeat.com/stocks/NASDAQ/NBIS/forecast/) | 20 | $84 | **$241.80** | n/a | $410 | 2026-10-04 (includes stale targets) |
| Markets Insider | 43 | $30 | n/a | $155.98 | $410 | stale-looking, not used |

Ratings: MarketBeat has 1 Sell / 6 Hold / 11 Buy / 2 Strong Buy (Moderate Buy). S&P's Sep-26 tally is 4 Strong Buy / 1 Buy / 3 Hold / 0 Sell / 1 Strong Sell.
Recent changes: **William Blair** initiated Outperform, no PT (9/30) · **BNP Paribas Exane** upgraded to Outperform, $260→$399 (9/24) · **Rothschild & Co Redburn** initiated **Sell, $84** (9/21) · **Northland** reiterated $410 (9/18; raised from $248 on 7/20) · **Truist** initiated Buy, $355 (9/9) · **Citi** $278→$324 (8/14) · **Baird** $250→$340 (8/13) · **DA Davidson** Neutral, $175 (8/13) · **Piper Sandler** initiated Neutral, $224 (8/3) · BofA $240→$280 (6/8) · Morgan Stanley EW, $144 (5/14, stale). The firm behind the $415 high: not found.
**So what:** The mean target implies 0% to +17% upside depending on the aggregator. The stock already trades at or near consensus targets.

## 5. Short interest (FINRA, Sep 15, 2026 settlement)
46,779,540 shares short (R). That is +3.3% vs 45.27M on Aug 31 but **down ~23% from the 61.07M peak on Jul 15**. Short % of float is 18.5% (MarketBeat) / 20.5% (StockAnalysis) / 22.05% (ChartMill) / 22.8% (Gainbot); the spread comes from different float definitions. Days to cover is 3.6 on FINRA ADV (2.2–2.8 elsewhere). Borrow fee is **~0.3%**, so the stock is easy to borrow ([CurvedTrading](https://curvedtrading.com/stocks/nbis/)). Next FINRA data comes Oct 9 (Sep 30 settlement). [MarketBeat history](https://www.marketbeat.com/stocks/NASDAQ/NBIS/short-interest/).
**So what:** Cheap borrow plus $13.43B of converts outstanding (no capped calls, so convert-arb funds must delta-hedge by shorting the stock) means a large share of the short interest is probably convert-arb hedging rather than directional bear conviction (this is an inference, not measured). The Aug 12 print moved the stock +34% in a day, so squeeze risk into the Nov 10 print is real.

## 6. Bear pieces / short reports
- **No activist short-seller report found.** I searched Kerrisdale, Spruce Point, Grizzly, Culper, Fuzzy Panda and general queries.
- **Michael Burry** (Substack "Cassandra Unchained"): disclosed a short at $211.77 on Aug 6, 2026. On Sep 28 he covered the common short and switched to **June puts at double-digit strikes**, saying he was "moving timelines up." Core claims: GPU depreciation/useful life, debt-funded capex. [24/7 Wall St](https://247wallst.com/investing/2026/09/29/nebius-climbs-4-despite-michael-burrys-bigger-bearish-bet-coreweave-ticks-up-cloudflare-slips/), [Oninvest](https://en.oninvest.com/article/an-investor-from-short-game-bet-against-nebius-ahead-of-its-quarterly-report)
- **Rothschild & Co Redburn**, Sell $84 (Sep 21): neocloud earnings are "cyclically inflated"; credit markets are pricing risks equities ignore; hyperscaler off-balance-sheet leverage limits the next leg of spending; the market prices "the most optimistic outcome from the compute-repricing thesis." [TipRanks/TheFly](https://www.tipranks.com/news/the-fly/nebius-initiated-with-a-sell-at-rothschild-co-redburn-thefly-news)
- **Seeking Alpha:** Cavenagh Research, "A Blowout Quarter Just Made Burry's Short More Interesting" (Aug 13; D&A is 45% of revenue, FCF is negative) [link](https://seekingalpha.com/article/4936071-nebius-stock-blowout-quarter-just-made-michael-burry-short-more-interesting) · Bears of Wall Street, "The Math Doesn't Work" (Sep 3; funding gap, ~30% downside) [link](https://seekingalpha.com/article/4942787-nebius-math-does-not-work) · "It's Getting Worse And Worse" (Sep 15; dilution, 20% issuance authority, convert interest, pricing normalization in 2027–28) [link](https://seekingalpha.com/article/4946606-nebius-its-getting-worse-and-worse)
- **DA Davidson (Gil Luria):** Neutral, $175 (Aug 13).

## 7. Implied multiples
| | EV (USD B) | EV/FY26 rev | EV/FY27 rev | EV/FY27 EBITDA |
|---|---|---|---|---|
| NBIS headline EV (aggregator, incl. leases) | 68.71 | 20.6x | 5.59x | 10.2x |
| **NBIS base (pro forma, ex-leases)** | **74.73** | **22.4x** | **6.08x** | **11.1x** |
| NBIS fully diluted | 81.75 | 24.5x | 6.65x | 12.2x |
| **CoreWeave (pro forma, ex-leases)** | **79.52** (team comps; aggregator $95.46B incl. ~$16.5B leases not used) | 6.2x (12.87) | **3.03x** (26.26) | **4.78x** (16.64) |

CRWV consensus comes from [Yahoo](https://finance.yahoo.com/quote/CRWV/analysis/) and [MarketScreener](https://ca.marketscreener.com/quote/stock/COREWEAVE-INC-185380761/calendar/). On a consistent ex-lease basis, NBIS trades at **~2.0x CRWV on FY27 revenue (+101%) and ~2.3x on FY27 EBITDA (+132%)**. Other comps from the team's comps.xlsx: revenue per active MW is CRWV ~$6.9–8.2M and IREN AI ~$7.1M (NBIS's active MW is stale at ~170MW YE25, so it can't be compared). Cost of debt is NBIS 5.1% vs CRWV 8.4%. Illustrative cross-checks (D):
- At CRWV's 4.78x, NBIS's FY27 EBITDA supports about $32.2B of EV, i.e. ~$32.4B equity, or **~$105/share (−57%)** on 308.69M shares. At CRWV's 3.03x FY27 revenue: ~$37.3B EV, or ~$121/share (−50%).
- 8x FY28 consensus EBITDA ($13.98B) minus consensus YE28 net debt ($50.7B) gives ~$61B of equity (~$198/share on 308.69M) vs the $74.95B economic market cap.

## 8. Price history: how implied expectations moved (D unless noted)
| Qtr-end | Close | Shares (M) | Mkt cap ($M) | Net cash/(debt) incl. leases ($M) | EV ($M) | Qtr rev ($M) | EV / (qtr rev×4) |
|---|---|---|---|---|---|---|---|
| Q1–Q3 2024 | trading halted 2022-02-28 → 2024-10-21 ([Nasdaq](https://www.nasdaq.com/press-release/nasdaq-resumes-trading-nebius-group-nv-2024-10-18)) | | | | | | n/a |
| Q4 2024 | 27.70 | 235.75 | 6,530 | 2,385 | 4,145 | 37.9 | 27.3x |
| Q1 2025 | 21.11 | 238.11 | 5,027 | 1,259 | 3,768 | 55.3 | 17.0x |
| Q2 2025 | 55.33 | 238.71 | 13,208 | 455 | 12,753 | 105.1 | 30.3x |
| Q3 2025 | 112.27 | 251.81 | 28,271 | 226 | 28,045 | 146.1 | 48.0x |
| Q4 2025 | 83.71 | 255.16 | 21,359 | (1,220) | 22,579 | 227.7 | 24.8x |
| Q1 2026 | 103.76 | 256.10* | 26,573 | (217) | 26,790 | 399.0 | 16.8x |
| Q2 2026 | 276.17 | 271.86* | 75,078 | (2,153) | 77,231 | 582.3 | 33.2x |
| Q3 2026 (Sep 30) | 235.88 (R) | not yet reported | | | | | |

Sources: quarter-end closes from FINRA report-date prices on [MarketBeat](https://www.marketbeat.com/stocks/NASDAQ/NBIS/short-interest/) (single source; not cross-checked for every quarter). Shares and net cash from [StockAnalysis balance sheet](https://stockanalysis.com/stocks/nbis/financials/balance-sheet/?p=quarterly). Quarterly revenue from MarketScreener and the company.  Q1/Q2'26 share counts: StockAnalysis basis; Q2 (company count) includes the warrant shares. Today: base EV of $69.6B is **30x Q2 run-rate revenue, ~23x the $3.0B June ARR, and ~8.7x the midpoint of the YE26 ARR guide ($8B)**.

## Red Team read: what the market is implicitly assuming
1. **A ~4x revenue step in 2027 is in the price.** *v2 check (03_depreciation v2, Rev_MW):* on Capacity's quarterly billing-MW path, FY27 revenue comes out at $9.9B in the base case, $7.0B low and $12.9B high, vs $12.3B consensus. Only the high case clears it. Consensus FY27 revenue of $12.3B (range $8.8–18.0B) against YE26 ARR guidance of $7–9B requires average 2027 ARR of about $12.3B. With a linear ramp from ~$8B, that means **exiting 2027 at ~$16–17B ARR** (my estimate). In MW terms (rough; connected power ≠ billable IT load), $12.3B needs ~1.0GW billing all year at the 2026 base ~$12M/MW, or ~615MW at the >$20M/MW Q2-deal pricing. That is essentially **all of the YE26 connected power (0.8–1.0GW) billing for the full year at no slippage**, plus new capacity.
2. **Margin expansion and pricing durability.** FY27 consensus EBITDA margin is 54.7% (60.8% in FY28), vs the ~40% FY26 guide and 41% in Q2. That needs the $20M+/MW repricing to hold even though management itself flagged that pricing could normalize in 2027–28 (per the SA bear piece), and even though Rothschild says it is cyclical.
3. **Funding stays open.** Consensus has **$29B of capex and −$23B FCF in FY27, with net debt reaching $34B (YE27) and $51B (YE28)**. Equity holders are assuming continued access to converts, asset-backed debt and ATM issuance without an equity-value-destructive raise. That is ~$13.43B of converts already, 20% issuance authority, and 12.27M ATM shares left.
4. **A premium to the closest peer.** At ~6.1x FY27 revenue and ~11.1x FY27 EBITDA (ex-leases), NBIS trades at about 2.0–2.3x CoreWeave (3.0x / 4.8x). If it re-rated to CRWV's EBITDA multiple on the same consensus, EV would be about $32B (~$105/share). On EV per contracted MW the two are essentially equal (~$21.4M vs ~$21.5M). Both are mostly leased (NBIS started leases $1.63B plus $12.05B signed but not commenced), so the premium is about how fast MW converts to revenue at high $/MW, not about the land-and-power bank.
5. **Accounting tailwinds persist.** The 5-year server life (previously 4) lowers D&A by ~$43M a quarter. In FY27/FY28 it adds ~9.6/10.5 points of EBIT margin (03 v2). If GPU economic lives are shorter (Burry's thesis), EBIT and FCF payback are worse than the adj. EBITDA framing suggests.

## Gaps / caveats
- No Q3'26 balance sheet exists yet, so pro forma cash excludes Q3 burn and prepayments. Carrying value of the new converts and of the exchanged notes is not disclosed (face values used).
- NVIDIA warrant exercise after Sep 11: not reported. The warrant is now added to economic basic shares (Filings v3); the Q2'26 history row uses issued shares and excludes it. No post-Jun-30 ATM sales disclosed.
- Microsoft "all tranches delivered" vs Q2 customer revenue (~$140M/qtr top customer): unresolved.
- No breakdown of how many analysts sit behind each EBITDA estimate. FY27 EBITDA ranges $6.36–6.73B across aggregators. No FY2029–30 consensus found.
- Price-target aggregators conflict (mean $241.80 vs $283.58). Markets Insider looks stale. The firm behind the $415 high was not identified.
- Short % of float varies with the float definition. The convert-arb share of short interest is not measurable from public data.
- CRWV EV is now the team comps pro forma ex-lease figure ($79.52B). The Oracle/IREN multiples live in comps.xlsx and are not repeated here.
- Quarter-end prices come from a single source (MarketBeat). No full Seeking Alpha text was read (paywall), only the summaries.
