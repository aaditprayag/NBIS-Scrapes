# RECONCILE — v2 freeze vs deck (Nebius Peer KPIs, Fri Oct 9, 2026 ~4:45 PM ET; desk research only)

Read-only inputs: this folder's v2 copies; locked deck line `skeptic/08_friday_debate_prep.md` l.24 (band line); slide `market/charts/contract_band_slide_final.png` (built by `market/scrapers/final_charts_1009.py`, hard-coded values; its source data is `market/contract_economics.xlsx` Print_Table). BM = benchmark/peer_benchmark.xlsx, CMP = comparison.xlsx, POM = price_or_mw.xlsx, APX = final_appendix/peer_kpi_appendix.xlsx.

## 1. IREN live MW — recommend **~40 MW [R, Jun 30, 2026]**
- The deck line compares company-level contracted vs live MW. ~40 MW is IREN's only company-level number ("operating AI Cloud Services capacity ~40MW", FY26 10-K + FY26 results 8-K, both Aug 27, 2026). Childress ~75 MW (50–100) is a single-site, team-derived estimate [D] (sites/appendix_iren_childress §2). It's the wrong scope for this line, so keep it as context only.
- Newer primary check (EDGAR, Oct 9): the only 8-Ks since Jun 30 are Jul 1, Jul 20, Aug 4, Aug 13 (Horizon 1 accepted, 50 IT MW) and Aug 27 (FY26). The only later filing is a 13G (Sep 30). No investor update gives a newer company-level live MW. Next print is Q1 FY27 (~early–mid Nov, est.).
- Ratio: 260 / 40 = 6.5x is an **upper bound**, because the live figure predates Horizon 1. Show **"≤6.5x [D]"**, not "6.5x" (APX shows 6.5x [D]).
- Footnote: "IREN live = ~40 MW AI Cloud operating capacity at Jun 30, 2026 (FY26 10-K, R), company-wide. It excludes Horizon 1 (50 IT MW, accepted Aug 13). Childress ~75 MW active (Oct, D) is one site and isn't used. Contracted ~260 IT = Microsoft 200 + NVIDIA ~60 (D). Ratio is an upper bound."

## 2. Oracle contracted — recommend **">10 GW [R, FQ1 FY27 call Sep 10, 2026]"**
- CMP Chart_Data E34 stores 10,000 only as a chart value. The primary wording is "'>10 GW secured'" (BM Benchmark row 6). Oracle files no contracted-MW figure, and "secured" isn't the same as contracted. Don't show "10,000".
- Ratio: show **n/d**, not a lower bound. The numerator is a floor (>10 GW). The denominator, ~2,050 [E], is cumulative MW delivered since Jun-25, not the installed base (n/d). The biases run in opposite directions, so no bound holds. This matches CMP K04 (n/d). APX shows ">4.9x [D]", so change that in a v3 (owner: Peers).
- Footnote: "Oracle: '>10 GW secured' (FQ1 FY27 call, Sep 10; R, not a filed contract figure). Live ~2,050 = AI MW delivered since Jun-25 (E); installed base n/d. Contracted/live n/d."

## 3. Deck-match check (deck line + slide vs v2 freeze)
| # | Number | Deck / slide | v2 value | File/cell (v2) | Match |
|---|---|---|---|---|---|
| 1 | Band low | $9.7M (line, slide) | 9.70 | BM Comp_Band C11, C24 | Y |
| 2 | Band top | $16.7M (line, slide) | 16.67 | BM Comp_Band C12, C25 | Y |
| 3 | IREN–Microsoft, 200 IT MW | $9.70M **[R]** | 9.70 **[D]**, 200 [R] | BM Contracts K8/L8, H8; CMP P1 W10 | **N (tag)** |
| 4 | Meta–Nebius #1, ~50 MW | $11.52M [E] | 11.52 [E], 50 | BM K28/L28; CMP N2 W6 | Y |
| 5 | Nscale–MSFT Narvik | $11.71M [D] | 11.71 [D] | BM K16/L16 | Y |
| 6 | Nscale–MSFT all SOWs | $13.35M **[D]** | 13.35 **[P]** | BM K15/L15 | **N (tag)** |
| 7 | MSFT–Nebius, 240 IT MW | $14.49M **[D on E]** | 14.49 **[E, house]** | BM K27/L27; CMP N1 W4 | **N (tag)** |
| 8 | OpenAI–Oracle (4.5 GW ÷1.2) | $16.00M **[E]** | 16.00 **[P]** | BM K11/L11 (O2-IT) | **N (tag)** |
| 9 | Anthropic–Nscale, 460 IT MW | $16.16M **[R/D]** | 16.16 **[P]** | BM K14/L14 | **N (tag)** |
| 10 | Anthropic–Lambda, ~350 MW | $16.67M **[R (secondary)]** | 16.67 **[P]** | BM K17/L17 | **N (tag)** |
| 11 | Nebius Q2'26 deals | $20–25M [R] (line, slide) | 20–25 [R] | BM KPI_Prints D9:E9 | Y |
| 12 | Bitdeer A201, ~15 MW | ≤$22.4M **[D]** (line, slide) | 22.38 **[D (basis P)]**, 15.19 | BM K24/L24, H24 | **N (tag, minor)** |
| 13 | Clear-spot high | $22.68M [D] | 22.68 [D] | BM Nebius_Locks E32 | Y |
| 14 | Clear-spot base | $24.51M [D] (line, slide) | 24.51 [D] | BM Nebius_Locks E31 | Y |
| 15 | Street non-MSFT hurdle | ~$27.6M [D] (line, slide) | 27.63 [D] | POM Nebius_Check B37; BM Nebius_Locks C30 | Y ("~27.6" = 27.63 rounded) |

**Result: 8 / 15 match. All 7 mismatches are tag differences; there are no value or rounding mismatches.** The deck line's PT figures ($0, $0–15, $39–63, $237) are model outputs, not peer numbers. They aren't in the peer files and weren't checked (owner: Modeling).
- **Cause:** the slide tags the MW/input basis. The v2 freeze tags the result by its weakest input (BM legend: "P = press-only, or derived with any press input"). Mismatches: #6 (term is a 5–6-yr midpoint), #9 (6-yr term is press), #10 and #8 (press TCV), and #3 (all inputs R, so the result is D).
- **Fix #3:** change the slide to [R/D]. Its own source, Print_Table row 6, already says "R / D". Owner: **GPU Market** (`final_charts_1009.py` / slide).
- **Fix #6, #8, #9, #10, #12:** use weakest-input tags on the deck slide: #6 [D; term P], #8 [P; IT read E], #9 [R/D; term P], #10 [P], #12 [D; basis P]. Owner: **GPU Market**. The v2 values stay as they are.
- **Fix #7:** "D on E" and "E, house" mean the same thing; pick one. Recommend [E, house] to match the memo, benchmark and appendix. Owner: **GPU Market** (slide + Print_Table J10).
- **Side note (Peers, frozen):** the CMP N2 note still says "unsourced secondary ~50 MW". v2 now sources it (Print_Table #3, slide, model Contracts!D23:D24). Fix the wording in a v3 only.
