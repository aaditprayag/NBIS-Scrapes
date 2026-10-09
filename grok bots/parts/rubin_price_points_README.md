# Rubin price points: README (built 2026-10-04)

File: `rubin_price_points.xlsx` (sheets: `Rubin_price_points`, `GPUs_per_MW_check`). Build script: `build_rubin_price_points.py`.
Reference GB300 rates: $5.10 (new deal) and $4.65 (IREN 3-yr). Model density: GB300 491 / Rubin 320 GPUs per IT MW.

## Rubin price points found (public sources only)
| # | Point | $/GPU-hr | vs $5.10 | vs $4.65 | $/IT MW-yr (@320) | Type / term | Tag |
|---|---|---|---|---|---|---|---|
| 1 | Anthropic–Nscale (existing anchor) | 5.77 | 1.13x | 1.24x | $16.16M | Contract, 6 yr, ~460 MW | Derived ($44.6B/6/460 MW ÷ 320 GPU/MW ÷ 8,760) |
| 2 | **WhiteFiber–Prime Intellect (NEW)** | **7.15** | **1.40x** | **1.54x** | $20.04M | Contract, 3 yr, 576 VR200, Canada, start Q2-27 | Derived ($108.2M ÷ (576 × 8,760 × 3)) |
| 3 | Civo early-access "from" price (NEW) | 11.00 | 2.16x | 2.37x | $30.84M | List / indicative, reserved, term unspecified, UK, Q1-27 | Reported (unit assumed per GPU) |

Same-provider checks (not apples-to-apples): WhiteFiber B300 5-yr Iceland = $3.47/GPU-hr ($87.5M ÷ (576 × 8,760 × 5)), so VR200 is 2.06x. Civo B300 36-mo = $5.49 and on-demand = $7.49, so the Civo VR "from" price is 2.00x and 1.47x.

Cost side (Estimated): SemiAnalysis TCO puts Rubin at $3.57 vs GB300 $2.36/GPU-hr, or **1.51x** (1.56x at hyperscaler volume). Morgan Stanley BOM puts the VR200 rack at $7.8M vs ~$4.0M for GB300, or **1.95x** (1.68x if SOCAMM is bought direct). Bernstein goes up to $9.1M. Tom's Hardware quotes conflict (0.8–1.2x).

Not derivable: Nebius–Meta $12B/5-yr Rubin (no MW or GPU count disclosed). Nscale–Microsoft 1.35 GW LOI (no dollar value). No public Rubin rate card at CoreWeave, Lambda, Crusoe, Together, Nebius, Oracle, the hyperscalers, or marketplaces (SAVRN tracker: 0 providers). No Ornn or Silicon Data Rubin index exists. The Nebius Q2 auction "+15%" was for **Blackwell**, not Rubin.

## GPUs per IT MW (320 check)
NVIDIA's own page gives 40K Rubin GPUs per 100 MW AI factory, or **400/MW** (DSX MaxLPS design point; IT vs utility basis not stated). Rack power of ~190 kW (Max-Q) to ~230 kW (Max-P) gives 379 to 313 GPUs per MW of rack power, or roughly 300–365 per IT MW once overhead is added. A third-party figure of 430k GPUs per 1.35 GW (319/MW) is unverified. **320 is plausible but sits at the conservative (low-density) end.** The break-even premium is 491/density: 1.53x at 320, 1.40x at 350, 1.29x at 380, 1.23x at 400.

## Recommendation (per-GPU premium vs $5.10, holding 320 GPU/MW)
- **Bear 1.10x (~$5.60):** hyperscaler/frontier-lab anchor pricing on long tenor (Nscale 1.13x is a floor by construction).
- **Base 1.25x (~$6.40):** a blend of the two contract prints (1.13x on 6-yr hyperscale, 1.40x on 3-yr AI-native). It sits close to the current 1.2x case.
- **Bull 1.40x (~$7.15):** the WhiteFiber 3-yr print. 1.5x has no contract support; only Civo's list "from" price exceeds it.

So-what: no contract print reaches the 1.53x break-even. At 320 GPU/MW, Rubin revenue per IT MW comes to 0.74–0.91x of GB300 at $5.10. Meanwhile the cost ratio is 1.5–1.95x, so per-GPU margins are squeezed unless density is higher than 320 or pricing firms up. **Model tip:** drive Rubin off $/IT MW-yr ($16–20M), not a per-GPU premium, because the per-GPU figure swings ±20% with the GPUs/MW assumption.

Confidence: **Low–Medium.** That is 2 contract prints, both derived from TCV under ratable, 100%-billed assumptions. Point 1 rests on press-sourced MW figures. Point 2 is small (~$108M, about 2–3 MW).


## Monarch WV site status (2026-10-05)

SITE STATUS FOOTNOTE (Peer KPIs 2026-10-05): $16.2M/IT MW-yr is an owned-hall CONTRACT print (Nscale S-1 Anthropic order; ~$44.6B ÷ 6 yr ÷ ~460 IT MW). Monarch WV is civil-only as of late Sep 2026 (graded pads, no hall roofs, air permit draft, Caterpillar microgrid not live, financing not committed) — Grade Yellow for 460 MW by H1'28. Underwrite PRICE from the contract; do NOT treat Monarch MW as already-powered comps. Source: peers/sites/site_verification.md §7.
