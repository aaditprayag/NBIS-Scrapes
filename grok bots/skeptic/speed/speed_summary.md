# NBIS bear case: speed premium, residual value, credit (Red Team)
Prepared 2026-10-04 (ET). The workbook is `speed_study.xlsx`. Each figure below is tagged R = Reported, D = Derived or E = Estimated, and its source URL is in the sheet named in brackets.

## Verdict
**The speed thesis does not hold.**
- Nebius is not first to market. Per the GPU Market study (`/workspace/nebius/parts/time_to_deploy.xlsx`), it ranked 8th of 10 on H100 and 4th of 8 on H200. It trails the first neocloud by about 2.6–7.4 months on every generation and came after at least one hyperscaler on each one. On B200 it trailed CoreWeave by only about 0.5–2 months, but trailed Lambda and GCP by about 3–4 months.
- The only launch premium Nebius ever showed was competed away within about a year:
  - H100: about 12 months.
  - H200: under 2 months, and the cut came before GA.
- Nebius's 2026 price increases (H100 +53% since May, B200 +55%) are a market-wide shortage. AWS, Lambda and the index panels rose at the same time. This is not a reward for speed.
- Even on a generous model, a 3–6 month head start is worth about 3–6% of a GPU's lifetime revenue, and under 1% at fleet level.

The more useful bear angles are elsewhere:
- **(a) Rental-price regime risk.** The October 2026 prices look like the A100's 2022 peak, and the A100 then fell about 43% from peak to trough.
- **(b) Credit.** NBIS's own convertibles imply high-yield spreads.

## Study 1: speed premium (only the additions to the GPU Market study, per steering)
**Launch premium and half-life** [S1_Decay]

| Series | Launch price | Trough | Launch premium | Half-life |
|---|---|---|---|---|
| Nebius H100 | $4.85 (2023-11-01) | $2.95 (by 2024-11-15) | +64% (D) | 10.2 months (D). Gone entirely by ~12.5 months. |
| Nebius H200 | $4.85 pre-GA list (2024-10-06) | $3.50 at GA (2024-11-15) | +39% (D) | ~1.3 months (D) |

- Note on the H200: the teammate's Launch_Premium_Decay sheet starts at $3.50, so it misses the $4.85 pre-GA price.
- B200, B300 and GB200: no decay can be measured. Prices were flat and then rose during the 2026 shortage. Nebius has never posted a GB200 price ("Contact us").
- Hyperscalers cut in single steps rather than decaying gradually:
  - AWS P5 on-demand −44% at about 22 months.
  - Capacity Blocks −20% at about 5–11 months.
  - P5en −25% at about 6 months.

**Hyperscaler-entry test** [S1_HS_Entry]
- 36 neocloud price pairs could be measured around 12 entry events. Only 4 stepped down by 5% or more.
- All 4 were in the Aug–Nov 2024 H200 wave and hit the previous generation: Nebius H100 −39%, Lambda H100 −14%.
- In 2025, 0 of 28 pairs moved. This is correlation, not causation.
- Generation price ratios are stable (Nebius H200/H100 about 1.17–1.20x; B200/H100 about 1.86–1.89x) [S1_PrevGen]. That reads as a performance ladder, not a speed premium that decays.

**NPV of an N-month head start** [S1_HeadStart] (E)
- Inputs:
  - π0 = 64%, half-life 10.2 months, 48-month life (Nebius 20-F, 4-year GPU depreciation).
  - Steady-state price $2.95, 70% utilization, 12% discount rate.
- Results:
  - **N = 3 months:** excess premium is 3.2% of per-GPU lifetime revenue undiscounted, 3.8% in PV, about $2.7k per GPU.
  - **N = 6 months:** about 5.7% undiscounted.
- At fleet level, with about 15% of a generation's fleet online in the early window (an illustrative assumption), the value is about 0.6% of the generation's lifetime revenue.
- The input that matters is that Nebius's actual head start is **negative**.

## Study 2: residual value [S2_Residual, S2_RentalPrecedent, S1_Reconcile]
**Current used marks (asks and surveys, not sold prices)**

| Item | Mark | Age |
|---|---|---|
| H100 8-GPU server, used | $150–180k vs $246–315k new = 53–64% (R/D; hashrateindex 2026-04-08) | ~3.0 years |
| H100 PCIe, used | $32.3k vs $29.5k new = ~109% (R/D; Clustermelt 2026-09-01) | — |
| H100 SXM, Mercatus survey | 75–85% at 18–30 months; 40–75% at 36 months; 25–35% at 60+ months | — |
| A100 SXM4 80GB, used | $9.9k ≈ 40% of the DGX-A100 per-GPU launch price ($199k / 8) | ~6.3 years |

- **Against book value:** at about 3.25 years, 4-year straight-line book is about 19% and 6-year is about 46%. **Market marks are at or above book today**, so depreciation lives look conservative at current marks. This undercuts the "GPUs are worthless in 3 years" short.

**Rents for older generations** (CCIR panel, plus the GPU Market teammate's Wayback A100 series)
- A100 at 4–6 years: $1.5–2.8/hr (Lambda $1.79 → $2.79; Crusoe $1.65–2.00; RunPod $1.59; CoreWeave $2.70).
- V100 at about 8.5 years: $0.81 (CCIR fit −21%/yr).
- Both are above an estimated cash cost of about $0.40/hr, which implies an economic life of more than 6 years at posted prices.
- Azure retired V100 about 7.5 years after launch.

**The bear angle is cyclical.**
- The A100 fell 43% from its 2022 peak to its 2025 trough (CCIR $3.05 → $1.75).
- Applying precedent paths to Nebius's October 2026 H100 price of $4.50 gives about $2.3–3.5 at age 5, which is 22–48% lower (D/E):
  - A100 trough/peak ratio: $2.58.
  - V100 decline rate: $2.32.
  - A100 endpoint CAGR: $3.49.
- Residual marks are regime-dependent. Today's figures are shortage figures.

## Study 3: credit [S3_Credit]
**Trading levels**
- NBIS convert secondary prices: **not found**. The notes are 144A, there is no FINRA access, and no press quotes were found.
- CRWV straight bonds (R; bondterminal, 2026-10-01): 11.45–13.41% YTW, G+651 to 838bp, prices 85–93.

**NBIS August 2026 converts** (R; SEC 6-K, priced 2026-08-19)
- 0.50% due 2030: $3.45B, accretes to 110%.
- 4.50% due 2034: $2.3B, accretes to 125%.
- Bond-only yield at par (D): 2030 3.23% (UST −104bp); 2034 7.09% (UST +259bp before any option value).

**Convert-implied credit spread** (E; bond floor plus Black-Scholes call)
- Calibration: CRWV's new 2.875% 2033 convert implies 55% vol at its own bond spread (836bp), which is 0.77x its 30-day implied vol. Applying the same ratio to NBIS's 30-day IV of 74.6% gives about 57–58% long-dated vol.
- At that vol, NBIS's converts imply about **690–740bp (2030s)** and **950–1,000bp (2034s)**. That is CoreWeave-like or wider, even though NBIS has much lower leverage (pro forma debt $14.3B vs $75B equity).

**Merton 5-year spreads** (E)
- NBIS about 303bp; CRWV about 497bp.
- Both are below credit-market levels (CRWV bonds G+651 to 838bp; NBIS's convert-implied levels). **Credit is pricing more risk than equity.**

**Mark-to-model on 2026-10-02** (E, not a trade)

| Note | Model price (50–60% vol, 600–836bp) | Parity |
|---|---|---|
| NBIS 2030 | 95.5–106.4 | 77.5 |
| NBIS 2034 | 103.8–120.4 | 74.8 |

**Other reference points**
- Rates are up about 70bp since pricing.
- The legacy June-2025 notes were exchanged at about 1.6% over parity.
- The $775M secured facility is priced at SOFR+250.

## Gaps
- No NBIS convert trading prices.
- No eBay sold comps (scripted access was blocked). Used prices are asks or surveys.
- B200 and GB200 premium decay cannot be measured.
- Some GCP GA dates are approximate.
- f, the discount rate and utilization are illustrative.
- Posted prices are not transacted prices.
- Calibrated vol is low-to-medium confidence.
