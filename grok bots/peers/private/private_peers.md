# Private peer KPI comps: Nscale, Lambda, Crusoe

Built Sun 2026-10-04 (ET); refreshed Mon Oct 5 2026 (ET) by `build_private_peers.py`. Workbook: `private_peers.xlsx`, with tabs README, Summary, Nscale, Lambda, Crusoe, Debt_Facilities, Sources, Definitions and Notes_Gaps. Row IDs and definitions match `/workspace/nebius/peers/coreweave/crwv_kpis.xlsx` (I01-I27, K01-K40, D1-D8). Rows that exist only for private peers are PX01-PX14.

**Tags:** R = reported by the company or a filing; P = press-reported (secondhand); D = derived (formula in the workbook); E = estimated (rests on an assumption or a third-party estimate such as Sacra); n/d = not disclosed. A D row built on press inputs shows `(inputs P)`. Every figure has its source document, date and URL on the company tabs.

## Source base (EDGAR check 2026-10-05; no new filings since Oct 4)
- **Nscale:** public **S-1 filed 2026-09-18** (CIK 0002110365; DRS 2026-02-17 plus 3 DRS/A; **no S-1/A as of 2026-10-05**, no price range). Only audited source among the three. Monarch WV site status folded from `peers/sites/` (N60–N63).
- **Lambda:** **no S-1, F-1 or public DRS** on EDGAR (re-check 2026-10-05). Lambda, Inc. (CIK 0001954966) has only Form D filings (2022, 2024). Financials = press (The Information 2025, Bloomberg 2026) and Sacra estimates.
- **Crusoe:** Form D only (latest 2026-09-17: $3.11B of $3.90B Series F sold; no new Form D as of 2026-10-05). Metrics from company PRs (gross MW, TCV) and press revenue. Abilene site status: `peers/sites/` / C37.

## Latest values (Summary tab)

| ID | KPI | Unit | Nscale | Lambda | Crusoe | CRWV Q2'26 |
|---|---|---|---|---|---|---|
| PX01 | Revenue run-rate / forward revenue | $M/yr | 281 (D, 2026-06-30) | 1,500 (P, 2026-12-31) | 2,000 (P, 2026-12-31) |  |
| I01 | Revenue (latest reported period) | $M | 141 (R, 2026-06-30) | 140 (P, 2025-06-30) | 250 (P, 2024-12-31) |  |
| I09 | Active MW | MW | 81.0 (D, 2026-08-31) | 47.0 (P, 2025-06-30) | 1,000 (R, 2026-09-17) | 1,500 |
| I10 | Contracted MW (active + contracted) | MW | 1,370 (R, 2026-08-31) | n/d | 6,000 (R, 2026-09-17) | 3,700 |
| I11 | GPU count (active) | GPUs | 25,000 (R, 2026-08-31) | 25,000 (R, 2025-02-24) | n/d |  |
| PX04 | GPUs active + contracted | GPUs | 461,000 (R, 2026-08-31) | n/d | n/d |  |
| I14 | RPO (ASC 606) | $B | 56.4 (R, 2026-06-30) | n/d | n/d | 103.7 |
| I15 | Backlog / TCV (company metric) | $B | 103 (R, 2026-08-31) | 50.0 (P, 2026-09-30) | 140 (R, 2026-09-17) |  |
| PX05 | Weighted avg contract life | years | 5.70 (R, 2026-08-31) | n/d | n/d |  |
| K01 | Revenue per active MW | $M/MW-yr | 3.47 (D, 2026-06-30) | 9.70 (D (inputs P), 2025-06-30) | 2.00 (E (inputs P/R), 2026-09-17) | 6.9 |
| K03 | Revenue per contracted MW | $M/MW-yr | 0.21 (D, 2026-08-31) | n/d | 0.33 (E (inputs P/R), 2026-09-17) | 2.8 |
| PX10 | Contract ACV per contracted MW | $M/MW-yr | 13.2 (E, 2026-08-31) | 16.7 (D (inputs P), 2026-09-01) | n/d | Nebius new deals ~20-25 (ref) |
| PX11 | TCV per contracted MW | $M/MW | 75.5 (D, 2026-08-31) | n/d | 23.3 (D, 2026-09-17) |  |
| K04 | Contracted / active MW | x | 16.9 (D, 2026-08-31) | n/d | 6.00 (D, 2026-09-17) |  |
| K05 | Adj. EBITDA per active MW | $M/MW-yr | -4.92 (D, 2026-06-30) | n/d | n/d | 4.8 |
| K06 | Revenue per GPU | $K/GPU-yr | 11.2 (E, 2026-06-30) | 18.2 (E (inputs P/R), 2025-06-30) | n/d | 14.5 |
| K07 | Implied blended $/GPU-hr | $/GPU-hr | 1.29 (E, 2026-06-30) | 2.09 (E (inputs P/R), 2025-06-30) | n/d | 1.66 |
| PX13 | Contract-implied $/GPU-hr | $/GPU-hr | 4.49 (E, 2026-08-31) | 3.71 (D (inputs P), 2025-09-08) | n/d |  |
| K08 | List price (fleet-weighted / proxy) | $/GPU-hr | n/d | 2.99 (R, 2025-06-16) | 3.90 (R, 2026-10-04) |  |
| K09 | Implied utilization | % | n/d | 69.8 (E (inputs P/R), 2025-06-30) | n/d | ~18 |
| K14 | Net PP&E (or project cost) per active MW | $M/MW | 70.7 (D, 2026-06-30) | n/d | 12.5 (D, 2025-05-21) | 31.2 |
| K15 | Gross in-service PP&E per active MW | $M/MW | 20.8 (D, 2026-06-30) | n/d | n/d |  |
| K16 | Useful life - technology equipment | years | 5.50 (R, 2026-06-30) | n/d | n/d | 6 |
| K17 | Depreciation yield | % | 15.7 (D, 2026-06-30) | n/d | n/d | 15.5 |
| K19 | D&A % revenue | % | 124 (D, 2026-06-30) | n/d | n/d |  |
| K23 | RPO / annualized revenue | years | 201 (D, 2026-06-30) | n/d | n/d | 10.1 |
| I17 | Largest customer % revenue | % | 52.0 (R, 2026-06-30) | n/d | n/d |  |
| K26 | Top-N customers combined % | % | n/d | 50.0 (P, 2025-06-30) | n/d |  |
| K28 | Weighted effective cost of debt | % | 18.7 (D, 2026-06-30) | n/d | n/d | 8.4 |
| PX09 | Marginal secured debt cost (latest facility) | % | 6.25 (E, 2026-08-31) | 6.78 (R, 2026-10-01) | 9.10 (E (inputs P), 2026-09-10) |  |
| K38 | Total debt | $B | 0.14 (D, 2026-06-30) | n/d | n/d |  |
| PX06 | Equity raised (latest round) | $M | 3,360 (R, 2026-09-25) | 1,500 (R, 2025-11-18) | 3,108 (R, 2026-09-17) |  |
| PX07 | Latest post-money valuation | $M | 14,600 (R, 2026-03-09) | 5,430 (P, 2025-11-18) | 30,900 (R, 2026-09-17) |  |
| PX08 | Recent debt commitments (sum) | $M | 6,712 (D, 2026-09-04) | 2,934 (D, 2026-10-01) | 1,150 (D (inputs R/E), 2025-06-11) |  |
| PX14 | Valuation / forward revenue | x | n/d | n/d | 15.4 (E (inputs P), 2026-09-17) |  |

The reference column comes from crwv_kpis.md. PX10 for CRWV is the Nebius ~$20-25M ACV/MW benchmark on new deals.

## Findings
**Nscale (S-1).**
- H1'26 revenue was $140.6M (R), a $281M run-rate.
- The contract book dwarfs revenue: RPO $56.4B (R, 6/30/26) and TCV $103.4B (R, 8/31/26). Only $2.6B of TCV is active. Roughly 86% of TCV is Microsoft (up to $43.8B) plus Anthropic (up to $44.6B; financing not yet committed).
- 1,370 IT MW is active + contracted (R). The S-1 does not quantify active MW; the 5 active sites sum to 81 MW (D, upper bound).
- Revenue per active MW is about 3.5 (D). On a contracted basis, ACV is about $13.2M/MW-yr (E: TCV / 5.7-yr life / 1,370 MW), below the ~$20-25M Nebius benchmark.
- Contract-implied pricing is about $4.49 per GPU-hr (E) across 461k GPUs. Realized H1'26 revenue was $1.29 per GPU-hr (E), reflecting the ramp.
- Technology equipment useful life is 5-6 years (R). The PP&E depreciation yield is 15.7%, an implied life of 6.4 years (D), in line with CRWV's 15.5%.
- Largest customer was 52% of H1'26 revenue (R).
- Funding cost: the reported weighted effective rate on 6/30/26 debt was 18.7% (D), but that debt is only $137M of legacy Macquarie and GPU-facility borrowings. The August 2026 GPU facilities price at SOFR+2.375% (about 6.2%, E), with $6.7B of new commitments after June (D). $3.36B of pre-IPO converts and a $14.6B Series C valuation are both R.
- **Monarch WV footnote (for Modeling v0.9a):** contracted **460 IT MW** (N60, R) / evidenced-active **0** (N61, R from peers/sites). Anthropic up to $44.6B (N35); implied ACV **~$16.2M/IT MW-yr** if 6-yr press term (N62, E). Grade **Yellow** for 460 by H1'28 / Yellow-Red for 1.37 GW (N63). Off-grid gas microgrid; air permit still draft; financing uncommitted. **Cite this pack + peers/sites/site_verification.md §7** — contract price stands; do not treat as live capacity.

**Lambda (press only).**
- Q2'25 revenue was ~$140M, of which ~$114M cloud (P). Sacra estimates Dec-2025 ARR at ~$760M (E). Bloomberg reports 2026 revenue expected above $1.5B (P).
- Q2'25 capacity was 47 MW (P), giving about $9.7M cloud revenue per MW-yr (D on P inputs).
- Revenue per GPU was about $2.09 per GPU-hr (E), against a $2.99 H100 list price, which implies ~70% utilization or price realization (E).
- Contracts: NVIDIA's leaseback implies $3.71 per GPU-hr (P). The press-reported Anthropic deal ($35B over 6 years, ~350 MW) works out to about $16.7M ACV per MW-yr (D on P inputs).
- Backlog (company metric, per a Lambda investor letter reported by WSJ 2026-10-06) rose from $15B (Jun-26) to $50B (Sep-26) (P, L38/L39); ~$35B of the increase is Anthropic. Lambda is raising up to $4B at $14.5B pre-money (Blackstone + Coatue; talks, P, L37) and now targets a 2027 IPO.
- 2026 debt has moved to investment-grade-style structures: a $1.008B DDTL at 6.78% fixed (A(low)/Baa1, R), a $926M TLB at SOFR+300 (R), and a $1B credit facility (R). Total commitments are $2.93B (D). Outstanding debt is n/d.

**Crusoe (company PR plus press).**
- Series F: $3.9B at $30.9B post-money (R). TCV is $140B+, gross contracted capacity is 6 GW+, and ~1 GW gross is delivered (R, 2026-09-17).
- 2026 revenue is 'on track for $2B' (P, Newcomer), which implies about $2.0M per gross MW-yr (E). That figure mixes landlord and cloud revenue and includes JV capacity, so it is not comparable to cloud peers.
- TCV per gross contracted MW is about $23M (D), but no term is disclosed. Valuation is about 15x 2026E revenue (E).
- The Abilene JV costs about $12.5M per gross MW for shell and power, excluding GPUs (D).
- Debt pricing: the Iceland GPU DDTL price talk is SOFR+475-500 (P), about 9.1% all-in (E). An AMD-backstopped loan is reported at ~6% (P).

## Gaps (structural vs watch)

### Structurally unknowable (until S-1 / audited disclosure)
- **Lambda:** no S-1 → no audited revenue, MW total, GPU count since Feb-2025, ASC 606 RPO (press backlog only: $50B Sep-26), debt outstanding, useful lives or concentration. Latest hard revenue = Q2'25 press.
- **Crusoe:** no financials; revenue press-only; MW gross incl. JV/landlord (Abilene GPUs = Oracle/OpenAI). No TCV term, GPU count, concentration, debt balance, useful lives.
- **Nscale:** company-wide active MW not quantified (81 MW = site-sum upper bound). No list price → K08/K09 n/d. Half-year only.
- **All three:** K02 and K10–K13 n/d (no MW history). Every KPI is a single dated point.

### Known caveats (handle carefully)
- Nscale Anthropic $44.6B / Monarch 460 MW: financing not committed; campus civil-only (Grade Yellow). $16.2M ACV/MW is a contract print, not live capacity (N60–N63; peers/sites §7).
- Crusoe Rev/MW mixes landlord + cloud and includes JV capacity — not comparable to pure cloud peers. Abilene counted under Crusoe (C37 / peers/sites).
- Lambda Aug-28 ~$1B private debt may overlap the Oct-1 IG DDTL — not double-counted in PX08.

### Watch next
- Nscale S-1/A (price range; Q3 active MW); Anthropic financing close; WVDEP final R14-0042; Monarch first roofs (peers/sites). IPO talk ~Oct-15 / ~$35B remains P/unconfirmed (EDGAR 2026-10-05: no S-1/A).
- Lambda public DRS/S-1 (IPO now 2027 per WSJ); up-to-$4B / $14.5B pre-money round close (WSJ 10/6 talks); KC Blitz go-live.
- Crusoe Iceland DDTL final pricing; Abilene B5–8 energization (peers/sites YE26).
