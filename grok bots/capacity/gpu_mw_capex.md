# Nebius (NBIS): GPUs per MW and capex per MW by NVIDIA generation

_As of Sun Oct 4, 2026 (ET). Public information only. Model: `gpu_mw_capex.xlsx`, with tabs GPUsPerMW, CapexPerMW, Nebius checks, Peers and Sources. Inputs are blue and formulas are live. Every figure is tagged in the workbook as Reported, Derived or Estimated._

## 1. GPUs per MW (Estimated; specs Reported)
Method: all-in IT kW/GPU = (system or rack max kW ÷ GPUs) × (1 + overhead for scale-out network, storage and mgmt). GPUs/IT-MW = 1000 ÷ that figure. Gross (facility) MW = IT MW × PUE.

| Generation | IT-MW low / base / high | Gross-MW low / base / high | PUE used (L/B/H) | Unit kW basis |
|---|---|---|---|---|
| Hopper H100/H200 (HGX 8) | 665 / 713 / 813 | 458 / 548 / 739 | 1.45 / 1.30 / 1.10 | 10.2 kW DGX max (NVIDIA) |
| B200 (HGX 8) | 474 / 509 / 575 | 327 / 391 / 523 | 1.45 / 1.30 / 1.10 | 14.3 kW (NVIDIA) |
| GB200 NVL72 | 487 / 538 / 571 | 361 / 449 / 519 | 1.35 / 1.20 / 1.10 | 120–132 kW; base 1.86 kW/GPU ≈ SemiAnalysis 1.87 |
| B300 (HGX 8) | 452 / 502 / 534 | 312 / 386 / 486 | 1.45 / 1.30 / 1.10 | 14–15 kW DGX B300 (NVIDIA) |
| GB300 NVL72 | 453 / 491 / 519 | 335 / 409 / 472 | 1.35 / 1.20 / 1.10 | 132–142 kW (NVIDIA/HPE/Lenovo) |
| Vera Rubin NVL72 (72 packages; "NVL144" counts dies) | 280 / 320 / 361 | 207 / 267 / 328 | 1.35 / 1.20 / 1.10 | 190–230 kW (Kuo estimate; Supermicro sizes 227 kW). NVIDIA hasn't published a figure |

## 2. Owned all-in capex per IT MW (Estimated, US$M)
Each figure is GPU/server + networking/storage + DC shell/power, using base density.

| Generation | Low / Base / High | Base split: GPU-server / network+storage / DC | Base per gross MW | Colo capex (no shell) base | Colo incl. PV of 10-yr lease (base) |
|---|---|---|---|---|---|
| Hopper | 28 / 36 / 46 | 20.5 / 5.4 / 10 | 27.6 | 26.9 | 40.9 |
| B200 | 34 / 42 / 53 | 27.3 / 5.1 / 10 | 32.6 | 33.4 | 47.5 |
| GB200 NVL72 | 33 / 42 / 51 | 23.2 / 6.0 / 12.5 | 34.7 | 30.2 | 44.3 |
| B300 | 38 / 50 / 62 | 34.5 / 5.6 / 10 | 38.6 | 41.1 | 55.2 |
| GB300 NVL72 (v2) | 39 / 49 / 60 | 30.7 / 6.1 / 12.5 | 41.1 | 37.8 | 51.9 |
| Vera Rubin NVL72 (v2) | 42 / 53 / 65 | 34.7 / 5.3 / 12.5 | 43.8 | 41.1 | 55.1 |

These figures line up with the external benchmarks:
- Jensen Huang puts all-in cost at $50–60B per GW.
- GPU plus networking comes to 70–80% of owned capex in the model, a little above Barclays' 65–70%.
- Rubin's $/MW looks lower only because power per GPU jumps about 1.5x. Rubin pricing is the least certain input.

## 3. Nebius-implied cross-checks (Derived)
- **Modi'in:** 4,000 B200 on 8 MW = 500/MW. That matches the B200 model (509/IT-MW), so the 8 MW reads as IT load.
- **Longcross:** 4,000 B300 on 16 MW = 250/MW (phase 1 not full). Adding phase 2 gives 7,000 GPUs: 438/MW on 16 MW or 368/MW on 19 MW. Those numbers fit a gross basis (~386).
- **GPU counts run ahead of MW:**
  - YE24: >20k GPUs on 25 MW is about 800/MW.
  - End-Q1'25: ~38k GPUs on ~35 MW is about 1,086/MW, which is physically impossible for Hopper.
  - Patmos' claim of 35k GPUs on 40 MW (875/MW) is also above the physical maximum.
  - Conclusion: Nebius GPU counts include GPUs delivered but not yet powered. Don't back MW out of GPU counts.
- **YE25 balance sheet:** gross servers + network of $3,124M ÷ 170 MW active = **$18.4M/MW**, below the Hopper model's ~$25.9M/IT-MW. Possible reasons: active MW is close to gross, purchase prices were low, or some active hardware still sits in assets-not-in-use. Including the $2,417M of assets not yet in use gives $32.6M/MW (an upper bound).
- **2026 guide:** $16–20B capex for +580–780 MW connected (220 → 800–1,000 MW) = **$20.5–34.5M per connected MW**.
  - This is understated, because GPUs for capacity connected late in 2026 are bought in 1H27.
  - DC build (≈20% of capex, per the CFO) works out to ~$5.3M/MW. That is below the $12.5M owned-build cost, which fits a colo-heavy 2026.
- **Payback:** Q2'26 deals run at $20–25M ACV/MW with a 1y10m payback. That implies capex of ~$18–32M/MW (base ~$25M at 60% contribution margin), below the GB300 GPU+net model ($40M/IT-MW). So Nebius's per-MW denominator is likely gross or contracted MW, or the payback excludes the shell.
- **Microsoft:** [RETIRED Oct 9: gross-MW basis; use $14.49M per IT MW = $17.4B ÷ 5 ÷ 240 IT MW, E house] $17.4B over 5 years ÷ 300 MW = $11.6M ACV/MW, which matches the ~$12M "2026 base" ACV/MW in the Q2 letter.

## 4. MW basis (team request)
- **Nebius wording:**
  - "Active" = "power being consumed by installed, operational IT equipment". That reads as an **IT basis** (Reported wording).
  - "Connected" = power connected into built, equipped DCs. The basis is ambiguous; it is probably provisioned or facility power (Estimated).
  - "Contracted" = land + power commitments, i.e. gross.
  - Nebius never labels site MW as IT or gross.
- **Vineland "300 MW" most likely means facility/gross power, so roughly 230–250 MW IT** (Estimated, medium confidence). Evidence:
  - Bloom capacity is 328 MW installed (Nebius PR) and ~250 MW guaranteed (secondary source), plus 5 MW of grid power for mechanical load.
  - DataOne says the fuel cells feed "data halls and associated mechanical loads".
  - If 300 MW were IT, then at PUE ≥1.2 the site would need ≥360 MW, more than the 333 MW installed.
- **Counterpoints:**
  - A secondary source (Measured AI) says "315 MW IT".
  - DataOne cites a "maximum capacity of 350MW".
  - Courier-Post says the plant is "capable of 374 MW".
  - The planning-board agendas give square footage only, with no MW.
  - The ">150k GB300 at Vineland" claim would need ~305 MW IT / ~366 MW gross, which doesn't fit the power supply. Treat it as unverified.

## 5. Peers (Derived from disclosures)
- **IREN:**
  - $5.8B Dell GB300 order ÷ 200 MW IT = $29M/MW IT (GPU + ancillary only).
  - Horizon 1 DC cost $6–7M/MW IT.
  - New contracts bring >$20M revenue per IT MW with a ~2-year payback.
- **CoreWeave:**
  - Net PP&E per active MW: $36M (YE25) and $31M (Q2'26).
  - The ~$37B 2026 capex guide (secondary source) works out to ~$37M per incremental active MW.
- **Nscale:**
  - 461k GPUs on 1.37 GW = 338/MW.
  - 66k Vera Rubin on 200 MW = 330/MW.
  - Monarch site: 8 GW gross vs 6.5 GW IT (ratio 1.23).
- **Oracle/OpenAI Abilene:** 450k GB200 on 1.2 GW = 375/MW, which looks like a gross basis.
- **Crusoe:** shell/power costs $12.5M/MW ($15B for 1.2 GW).
- **Lambda Kansas City:** ≥417 B300/MW.
- **Supermicro Vera Rubin blueprint:** 230 GPUs per MW of envelope.
- **Industry cost benchmarks:**
  - Cushman: $17.6M/MW greenfield, excluding GPUs.
  - Colocation rent: $165–205/kW-month ≈ $2.0–2.5M per MW-year.

## 6. Key caveats
1. Nebius never states IT vs gross, so every per-MW ratio carries a 1.1–1.45x PUE ambiguity.
2. NVIDIA doesn't publish Vera Rubin rack power or any list prices. Rack and server prices come from press or third parties (Tom's, SemiAnalysis, Data Gravity, vendor blogs). The H200 server price ($230k base) is an Estimated input with no primary source.
3. The colocation fit-out of $0.5–2M/MW is an Estimated assumption.
4. Some sources are secondary and still need checking: CoreWeave's capex guide, JLL's $/MW, Measured AI's Vineland details, and Morgan Stanley's Rubin BOM via Locsic.
5. Capex timing (GPUs bought around when power is energized) distorts any capex ÷ MW ratio within a single year.

## 7. v2 update (Oct 4 2026): sourced GB300 rack price; Rubin revised
v1 files are kept as gpu_mw_capex_v1.xlsx, build_gpu_mw_v1.py and gpu_mw_capex_v1.md. The new sheet **'GB300 rack price'** drives the GB300 rows in CapexPerMW.

**GB300 NVL72 rack (incl. in-rack NVLink trays, CX-8 NICs, in-rack cooling; excl. scale-out, storage, install):**

| Source | Price/rack | Notes |
|---|---|---|
| Loop Capital, Mar-25 (S37) | $3.7–4.0M | Apple ~$1B ≈ 250 systems |
| Morgan Stanley BOM, May-26 (S35) | ~$4.0M | ODM price |
| Wolfe, Jan-26 (S36) | ~$4.3M | "reported prices" |
| Data Gravity POs, Aug-26 (S9) | just under $5.0M | small orders, EXW |
| Tom's Hardware, Mar-26 (S8) | $6–6.5M | inference config; treat as stress case |
| IREN Aug-25 (S39) | ~$5.04M per 72 GPUs **all-in** | includes InfiniBand |
| IREN–Dell (S25/S42) | ~$5.5M per 72 GPUs **all-in** | GPU count unverified |

**Selected rack price:** low $4.0M / **base $4.5M** / high $5.0M (stress $6.25M). Network+storage+install stays at $0.9M/rack (= $6.1M/IT MW). That's in line with SemiAnalysis's GB200 all-in uplift of +$0.8M (S6). Scale-out fabric alone is only ~$1.6–4k per GPU (S45).

**GB300 compute + network per IT MW (491 GPUs/MW):** low **31.4** / base **36.8** / high **42.3** (stress 50.8). Leased (adding fit-out) is 31.9 / 37.8 / 44.3.

**Against the $33M IT-reading cutoff:**
- The base case fails by about $3.8M.
- The low case passes by $1.6M.
- Breakeven is a $3.94M rack, or ≤440 GPUs per IT MW at base prices.

**Peer and Nebius checks:**
- IREN–Dell all-in works out to $29M/IT MW. That passes the cutoff, but at about 380 GPUs per IT MW.
- Nebius Q2'26: $5,197.7M ÷ 183 = $28.4M per active MW. That is consistent with an IT reading of active MW for a fleet that is mostly Hopper.

**Rubin (v2):**

| Case | Rack | Network etc. per rack | Leased per IT MW | Matches |
|---|---|---|---|---|
| Low | $6.7M | $0.9M | 34.3 | — |
| Base | $7.8M | $1.2M | 41.1 | team base $41M |
| High | $9.1M | $1.5M | 49.2 | team bear $48.5M |

In v1, Rubin network at $3.6M/MW ($0.8M/rack) was too low. Rubin's CX-9 1.6T NICs double per-GPU scale-out bandwidth, so the network line should cost more per rack than GB300's $0.9M, not less. Its in-rack NAND only offsets storage.
