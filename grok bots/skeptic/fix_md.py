p='01_priced_in.md'; s=open(p).read()
FB='[Filings fact base](/workspace/nebius/filings/nebius_fact_base.md)'
def rep(a,b):
    global s
    assert a in s, a[:70]; s=s.replace(a,b)
i=s.find('> **Staleness flag:**')
rep(s[i:s.find('\n',i)], s[i:s.find('\n',i)]+' **Q2 capex alone was $5.66B**, so true Sep-30 cash is probably materially below the pro forma figure (Q2 OCF was +$2.25B, but that came from prepayments).\n\n> **Correction (Oct 4, 6:10 PM ET):** v1 double-counted the NVIDIA pre-funded warrant (21.07M sh). It is **already inside** the 271.86M Jun-30 basic count. v1 also compared NBIS ex-leases with a CoreWeave EV that included ~$16.5B of leases. Both are fixed below. Pro forma shares are now 287.66M (was 308.72M), base EV is $69.62B (was $74.73B), and the CRWV EV is $79.52B ex-leases (was $95.46B). Convert principal is $13.43B. Source: '+FB+' and team comps (`/workspace/nebius/parts/comps.xlsx`).')
a=s.find('| NVIDIA pre-funded warrant |'); b=s.find('Assumption: the 2033-note')
s=s[:a]+'''| NVIDIA pre-funded warrant | 21,065,936 sh @ $0.0001: **already included in the 271.86M** (Q2 wtd-avg basic of 280.4M includes it); no exercise reported through 10/1 | R | 2026-06-30 | '''+FB+''' |
| ATM sales after Jun 30 | none disclosed (12,270,507 sh remain of 25M program; Q2 sold 12.73M @ $223.60) | R | through 2026-10-01 | '''+FB+''' |
| **Pro forma economic shares** | **287.66M** = 271.86 + 15.8 | D | 2026-10-02 | above |
| Converts outstanding (orig. principal) | **$13.43B**: 2029/2031 ~$87.5M each @ $51.45 · 2030 1.00% & 2032 2.75% ~$1.58B each @ $138.75 · 2031 1.25% $2.59B @ $183.22 · 2033 2.625% $1.75B @ $180.31 · 0.50% 2/2030 $3.45B @ $313.46 · 4.50% 2/2034 $2.30B @ $324.65. **No capped calls.** | D | 2026-10-02 | '''+FB+''' |
| ITM convert shares @ $242.81 | 50.02M (1.70M + 1.70M + 11.40M + 11.40M + 14.12M + 9.71M) | D | 2026-10-02 | above |
| OTM convert shares | 18.09M (11.01M + 7.08M) | D | 2026-10-02 | above |
| Options / RSUs | 6.74M options (WAEP $88.65), so 4.28M net under the treasury-stock method / 6.23M RSUs | R / D | 2026-06-30 | '''+FB+''' |
| **Fully diluted** | **348.19M** = 287.66 + 50.02 + 4.28 + 6.23 (max incl. OTM converts + remaining ATM: ~378.6M) | D | 2026-10-02 | |
| Mkt cap, headline aggregator | $66.55B (274.10M sh) | R (agg.) | 2026-10-04 | [StockAnalysis](https://stockanalysis.com/stocks/nbis/statistics/) |
| Mkt cap, company count | $66.01B = 271.86M × $242.81 | D | | |
| Markets Insider | $64.13B (counts Class A only, so it is understated) | R (agg.), flagged | | [BI](https://markets.businessinsider.com/stocks/nbis-stock) |
| **Mkt cap, pro forma economic (base)** | **$69.85B** = 287.66M × $242.81 | D | | |
| Mkt cap, fully diluted | $84.54B = 348.19M × $242.81 | D | | |

'''+s[b:]
rep('| **Pro forma cash** | **14,497.1** = 8,042.1 + 775 + ~5,680 (company\'s net-proceeds estimate) | E, stale (no Q3 burn) | |','| **Pro forma cash** | **14,497.1** = 8,042.1 + 775 + ~5,680 (company\'s net-proceeds estimate) | E. **Stale: excludes Q3 burn** (Q2 capex was $5,657M) | |')
a=s.find('| **Base (pro forma, ex-leases)** | **74.73**'); b=s.find('**So what:** Aggregator EVs')
s=s[:a]+'''| **Base (pro forma, ex-leases)** | **69.62** | 69.85 + 14.27 − 14.50 |
| Base + leases | 71.24 | + 1.625 |
| Fully diluted (ITM converts treated as equity) | 76.64 | 84.54 + (14.27 − 7.675 ITM face) − 14.50 |
| Bear view (prepayments as debt-like) | 75.59 | 69.62 + 5.975 |
| EV per contracted MW | $19.9M/MW on >3.5GW (D); $13.9M on the 5GW YE26 target. CRWV: ~$21.5M (ex-lease, 3.7GW) | |

'''+s[b:]
i=s.find('**So what:** Aggregator EVs')
rep(s[i:s.find('\n',i)],'**So what:** On consistent math, pro forma EV ($69.6B) is only ~$0.9B above the aggregator headline ($68.7B). The ~15.8M exchange shares add ~$3.8B of market cap, and that is roughly offset by net cash after the Aug financing. Real diluted exposure is $76.6B once ITM converts, options and RSUs are counted. The Aug financing is close to net-debt-neutral until the cash is spent on capex, and that is already happening (Q3 burn is not captured).')
rep('Q2\'26 actuals: revenue $582.3M, adj. EBITDA $236.2M (41%), ARR $3.0B, capex $5.66B.','Q2\'26 actuals (R, '+FB+'): revenue $582.3M · adj. EBITDA $236.2M (41%) · D&A $259.7M · SBC $102.5M · interest $119.1M · capex $5,657M · OCF +$2,246M (prepayment-driven; deferred revenue +$1.2B) · PP&E $13,045M · deferred revenue $5,975M · RPO $37.5B (36% within 24 months) · top-3 customers 24/21/14% of revenue · ARR $3.0B. Microsoft: 9 tranches, $17.39B owed regardless of use, ~$6.96B upfront, billed through Oct-2031. The letter says all tranches are delivered, but **that conflicts** with a top customer of only ~$140M/qtr vs the ~$725–870M/qtr full delivery implies (open). Meta: #1 ~$2.88B/5yr; #2 $12B firm on Vera Rubin, tranches from early 2027, plus up to $15B backstop (likely outside RPO). ~70% of deals carry prepayments covering 50–60% of related capex.')
rep('Server useful life was extended from 4 to **5 years** in 2026.','Server useful life was extended from 4 to **5 years** on 1/1/26. That cut D&A by $43.0M in Q2 and ~$167.6M est. for FY26 (company figure). Pricing caveat: the installed base runs at ~$12M/MW. The new $20–25M/MW deals are priced on forecast costs and include capacity not yet built. Older-gen GPU pricing rose >30% QoQ in Q2.')
rep('$13.45B of converts outstanding','$13.43B of converts outstanding (no capped calls, so convert-arb funds must delta-hedge by shorting the stock)')
a=s.find('| NBIS headline EV | 68.71'); b=s.find('## 8. Price history')
s=s[:a]+'''| NBIS headline EV (aggregator, incl. leases) | 68.71 | 20.6x | 5.59x | 10.2x |
| **NBIS base (pro forma, ex-leases)** | **69.62** | **20.8x** | **5.66x** | **10.35x** |
| NBIS fully diluted | 76.64 | 22.9x | 6.23x | 11.4x |
| **CoreWeave (pro forma, ex-leases)** | **79.52** (team comps; aggregator $95.46B incl. ~$16.5B leases not used) | 6.2x (12.87) | **3.03x** (26.26) | **4.78x** (16.64) |

CRWV consensus comes from [Yahoo](https://finance.yahoo.com/quote/CRWV/analysis/) and [MarketScreener](https://ca.marketscreener.com/quote/stock/COREWEAVE-INC-185380761/calendar/). On a consistent ex-lease basis, NBIS trades at **~1.9x CRWV on FY27 revenue (+87%) and ~2.2x on FY27 EBITDA (+117%)**. Other comps from the team's comps.xlsx: revenue per active MW is CRWV ~$6.9–8.2M and IREN AI ~$7.1M (NBIS's active MW is stale at ~170MW YE25, so it can't be compared). Cost of debt is NBIS 5.1% vs CRWV 8.4%. Illustrative cross-checks (D):
- At CRWV's 4.78x, NBIS's FY27 EBITDA supports about $32.2B of EV, i.e. ~$32.4B equity, or **~$113/share (−54%)**. At CRWV's 3.03x FY27 revenue: ~$37.3B EV, or ~$130/share.
- 8x FY28 consensus EBITDA ($13.98B) minus consensus YE28 net debt ($50.7B) gives ~$61B of equity (~$212/share on today's count) vs the $69.9B pro forma market cap.

'''+s[b:]
rep('Today: base EV of $74.7B is **32x Q2 run-rate revenue, ~25x the $3.0B June ARR, and ~9.3x the midpoint of the YE26 ARR guide ($8B)**.','Q1/Q2\'26 share counts: StockAnalysis basis; Q2 (company count) includes the warrant shares. Today: base EV of $69.6B is **30x Q2 run-rate revenue, ~23x the $3.0B June ARR, and ~8.7x the midpoint of the YE26 ARR guide ($8B)**.')
rep('*Excludes the 21.07M NVIDIA warrant shares, so Q1/Q2\'26 market cap is understated by about $2.2B and $5.8B.','')
rep('That is ~$13.45B of converts already, 20% issuance authority, and 12.3M ATM shares left.','That is ~$13.43B of converts already, 20% issuance authority, and 12.27M ATM shares left.')
rep('At ~6.1x FY27 revenue and ~11.1x FY27 EBITDA, NBIS trades at about 1.7–1.9x CoreWeave. If it re-rated to CRWV\'s multiple on the same consensus, EV would be about $39B (~$126/share).','At ~5.7x FY27 revenue and ~10.4x FY27 EBITDA (ex-leases), NBIS trades at about 1.9–2.2x CoreWeave (3.0x / 4.8x). If it re-rated to CRWV\'s EBITDA multiple on the same consensus, EV would be about $32B (~$113/share). On EV per contracted MW the two are close (~$19.9M vs ~$21.5M), so the premium is about how fast MW converts to revenue at high $/MW, not about the land-and-power bank.')
rep('- Whether the NVIDIA warrant was exercised after Sep 11, and any ATM sales after Jun 30: not found.\n- Options/RSU/PSU count not found, so the fully diluted figure excludes it.','- NVIDIA warrant exercise after Sep 11: not reported (warrant is in the basic count either way). No post-Jun-30 ATM sales disclosed.\n- Microsoft "all tranches delivered" vs Q2 customer revenue (~$140M/qtr top customer): unresolved.')
rep('- Oracle/IREN comps not done. CRWV EV is from an aggregator (Jun-30 balance sheet).','- CRWV EV is now the team comps pro forma ex-lease figure ($79.52B). The Oracle/IREN multiples live in comps.xlsx and are not repeated here.')
open(p,'w').write(s)
