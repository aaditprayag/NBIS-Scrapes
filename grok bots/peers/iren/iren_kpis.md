# IREN (NASDAQ: IREN) peer KPIs on the Nebius/CoreWeave template, AI Cloud perimeter

Built 2026-10-05 ET. Light polish 2026-10-05 afternoon: EDGAR re-check — no new 8-K/10-Q/6-K for IREN since 2026-10-04; KPIs unchanged. Childress site cross-link added. Files:
- Workbook: `/workspace/nebius/peers/iren/iren_kpis.xlsx` (recalculated; 0 formula errors).
- Builder: `/workspace/nebius/peers/iren/build_iren_kpis.py`.
- Shared engine: `/workspace/nebius/peers/peer_kpi_template.py`.
- Raw filings: `peers/iren/raw/` and `comps_raw/iren_*`.

**Template.** The row IDs (I01–I27, K01–K40), definitions D1–D8, and R/D/E/n-d tags (blue/black/orange/grey) are the same as in `crwv_kpis.xlsx`. Phase-2 extensions:
- Standard rows I28–I32 and K41–K45: useful life, reported GPU utilization, RPO share within 12 months, incremental revenue per incremental MW, revenue per dollar of GPU hardware, and segment gross margin.
- IREN-only memo rows I40+ and M01+, used for the Bitcoin-mining split.

**Fiscal mapping is exact.** IREN's fiscal year ends Jun-30, so Q1'24 = FQ3 FY24 (Mar-24) through Q2'26 = FQ4 FY26 (Jun-26). Details are in the Fiscal_Map tab.

**Accounting basis.** The 6-K/20-F periods through Mar-25 are IFRS. The FY25 10-K onward is US GAAP, with FY24 and FY25 recast.

## Q2'26 (FQ4 FY26, Jun-26), or the latest available

| KPI | IREN | Tag | CRWV Q2'26 |
|---|---|---|---|
| K01 AI revenue per operating AI MW (period-end) | **7.05 $M/MW-yr** (70.5 × 4 / ~40 MW) | D | 6.87 |
| K02 per average MW | n/d (latest Q3'25: 10.5) | E | 8.24 |
| K06 revenue per GPU | n/d (latest Q3'25: $14.7K/GPU-yr). Memo M07 Q2'26 on estimated period-end GPUs: ~$13.1K (range 8.8–14.4) | D / E | 14.5 |
| K07 $/GPU-hr | n/d (latest Q3'25: 1.66) | D | 1.66 |
| K09 implied utilization (vs IREN's own $2.00 assumption) | n/d (latest Q3'25: 83%) | E | 18% (vs list) |
| K10–K15 capex/MW and PP&E/MW | n/d. Capex is company-wide and there is no quarterly AI MW series. Memo M06: owned GPU hardware per AI MW = $50.9M/MW | – | 12.8 / 18.7 / 20.0 / 31.2 / 27.1 |
| K16 GPU (HPC hardware) useful life | 5 years (unchanged FY24–FY26; mining hardware 4 years) | R | 6 |
| K17 depreciation yield (company-wide) | 14.4% | D | 15.5% |
| K20 RPO | $5.1B | R | 103.7 |
| K21 backlog including ASC 842 lease value | $16.6B (Microsoft ~$9.7B ≈ 58%, NVIDIA $3.4B ≈ 20%, derived) | R | 104.2 |
| K22 RPO share within 24 months | 43% | D | 41% |
| K42 RPO share within 12 months | 18% | D | n/d |
| K24–K27 customer concentration | n/d (not disclosed) | – | 36% / 72% / 3 |
| K28 weighted effective cost of debt | 1.85% (converts 1.13%, DDTL 7.12%, USPP 7.05%) | D | 8.41% |
| K36 finance expense × 4 / average debt | 1.71% | D | 7.8% |
| K38 debt | $7.71B | R | 35.6 |
| K43 incremental revenue per incremental MW (YoY) | 6.8 $M/MW-yr | E | 5.29 |
| K44 revenue / average GPU hardware | 0.20× | D | 0.34× (tech equipment) |
| K45 AI segment gross margin (excl. D&A) | 86.9% | D | n/d |
| M01 AI share of total revenue | 51.4% (Bitcoin $66.7M vs AI $70.5M) | D | – |

## Quarterly trends (calendar column = IREN fiscal quarter; ᴱ = estimated)

| ID | KPI | Unit | Q1'24 | Q2'24 | Q3'24 | Q4'24 | Q1'25 | Q2'25 | Q3'25 | Q4'25 | Q1'26 | Q2'26 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| K01 | Revenue per active MW - period-end MW | $M/MW-yr | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 7.05 |
| K02 | Revenue per active MW - average MW | $M/MW-yr | n/d | 8.87ᴱ | 11.1ᴱ | 5.60ᴱ | 5.39ᴱ | 10.5ᴱ | 10.5ᴱ | n/d | n/d | n/d |
| K06 | Revenue per GPU ($K per GPU-year, avg GPUs) | $K/GPU-yr | n/d | 12.4 | 15.6ᴱ | 7.85ᴱ | 7.55 | 14.7 | 14.7 | n/d | n/d | n/d |
| K07 | Implied blended realized price ($/GPU-hr) | $/GPU-hr | n/d | 1.42 | 1.77ᴱ | 0.89ᴱ | 0.87 | 1.68 | 1.66 | n/d | n/d | n/d |
| K09 | Implied utilization (realized $/GPU-hr / weighted list) | % | n/d | 71.2ᴱ | 88.5ᴱ | 44.4ᴱ | 43.7ᴱ | 84.0ᴱ | 83.2ᴱ | n/d | n/d | n/d |
| K16 | Useful life - GPU / server equipment | years | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 |
| K17 | Depreciation yield (D&A x4 / avg in-service gross PP&E) | % | n/d | 29.2 | 21.6 | 14.8 | 14.7 | 15.2 | 18.0 | 18.1 | 18.6 | 14.4 |
| K20 | RPO (ASC 606, EoP) | $B | n/d | n/d | n/d | n/d | n/d | n/d | 0.20 | 0.29 | 0.71 | 5.10 |
| K21 | Revenue backlog (company metric / incl. contracted non-RPO, EoP) | $B | n/d | n/d | n/d | n/d | n/d | n/d | n/d | 9.99 | 10.4 | 16.6 |
| K22 | RPO % recognized within 24 months | % | n/d | n/d | n/d | n/d | n/d | n/d | 95.8 | 98.7 | 80.7 | 43.1 |
| K23 | RPO / annualized quarterly revenue | years | n/d | n/d | n/d | n/d | n/d | n/d | 6.64 | 4.18 | 5.28 | 18.1 |
| K28 | Weighted effective cost of debt (principal-weighted, EoP) - all facilities | % | n/d | n/d | n/d | 3.87 | 3.87 | 4.06 | 4.06 | 1.08 | 1.08 | 1.85 |
| K36 | Implied cash cost: interest x4 / avg principal | % | n/d | n/d | n/d | 11.4 | 3.74 | 2.90 | 3.75 | 1.80 | 1.58 | 1.71 |
| K38 | Total debt principal (EoP) | $B | 0.00 | 0.00 | 0.00 | 0.44 | 0.44 | 0.99 | 0.99 | 3.75 | 3.75 | 7.71 |
| K39 | Debt principal / RPO | x | n/d | n/d | n/d | n/d | n/d | n/d | 5.08 | 12.9 | 5.27 | 1.51 |
| K42 | RPO % recognized within 12 months | % | n/d | n/d | n/d | n/d | n/d | n/d | 47.9 | 51.9 | 43.4 | 17.6 |
| K43 | Incremental annualized revenue per incremental MW (YoY) | $M/MW-yr | n/d | n/d | n/d | n/d | 7.96ᴱ | 11.6ᴱ | 9.24ᴱ | n/d | n/d | 6.81ᴱ |
| K44 | Annualized revenue / avg gross GPU-compute hardware in service | x | n/d | 0.32 | 0.38 | 0.21 | 0.21 | 0.38 | 0.32 | 0.21 | 0.20 | 0.20 |
| K45 | Segment gross margin (revenue - cost of revenue excl. D&A) / revenue | % | n/d | n/d | 92.7 | 89.7 | 90.6 | 93.2 | 90.1 | 86.2 | 86.3 | 86.9 |
| M01 | AI Cloud share of total revenue | % | 1.05 | 4.47 | 6.04 | 2.29 | 2.47 | 3.72 | 3.06 | 9.37 | 23.2 | 51.4 |
| M02 | Bitcoin mining revenue per mining MW (annualized) | $M/MW-yr | n/d | 0.84ᴱ | 0.64ᴱ | 1.05ᴱ | 0.97ᴱ | 0.98ᴱ | 1.15ᴱ | n/d | n/d | n/d |
| M03 | AI MW share of operating data-center MW | % | 0.44ᴱ | 0.44ᴱ | 0.32ᴱ | 0.52ᴱ | 0.40ᴱ | 0.33ᴱ | 0.36ᴱ | n/d | n/d | 9.52 |
| M04 | Installed hashrate | EH/s | 10.0 | 10.0 | 21.0 | 31.0 | 37.0 | 50.0 | 50.0 | 46.0 | 38.0 | 23.2 |
| M06 | Owned GPU hardware (gross) per AI MW | $M/MW | 25.9ᴱ | 29.1ᴱ | 29.4ᴱ | 25.6ᴱ | 26.1ᴱ | 28.5ᴱ | 36.0ᴱ | n/d | n/d | 50.9 |
| M08 | Company Adj. EBITDA margin | % | 40.4 | 21.5 | 4.74 | 53.9 | 57.5 | 65.1 | 38.2 | 40.8 | 41.1 | 14.0 |


## AI cloud vs Bitcoin mining split (as far as disclosure allows)

- **Revenue:** fully split by segment (I01 AI, I40 Bitcoin). AI rose from ~1% of revenue to 51% in Q2'26. Bitcoin revenue peaked at $233M in Q3'25 and fell to $67M by Q2'26.
- **MW:**
  - AI MW is reported only for Jun-26 (~40 MW, basis unstated). Earlier AI MW is estimated as GPUs / 713 per IT-MW.
  - Bitcoin MW is reported only for Jun-26 (~380 MW at 23.2 EH/s). Earlier Bitcoin MW is estimated as total operating MW minus AI MW.
  - Total operating data-center MW is reported, but on a gross basis and often "as of PR date".
- **GPUs:** operating counts run 816 → 1,896 → ~1,900 → ~2,100 (Sep-25). After that IREN gives only installed/contracted targets.
- **Assets:** GPU hardware at cost (I31) is split from mining hardware (I48): $2.03B vs $0.60B at Jun-26, plus $0.30B of GPU finance-lease right-of-use assets. Capex, D&A and EBITDA are not split.



## Site cross-link (Peer KPIs alt-data)

Childress TX is IREN's key AI site in `/workspace/nebius/peers/sites/` (`site_verification.md` §2; `sites.xlsx`).
- **Announced:** 750 MW grid; Horizons 1–4 = 4 × 50 MW IT liquid-cooled (+H5–6 2027).
- **Contracted:** 200 MW IT Microsoft + 60 MW NVIDIA (from early 2027).
- **Evidenced-active AI IT (as of 2026-10-04):** L/B/H **50 / 75 / 100** (Horizon 1 delivered/accepted 2026-08-13; H2 commissioning). Grade **Green** (H1–2) / **Yellow** (H3–4 by Q4'26).
- **Bitcoin vs AI:** site still carries legacy mining load winding down; grid MW ≠ AI IT. Keep AI Cloud KPIs (I01/K01+) separate from Bitcoin memo rows (I40/M01–M04) — as this pack already does.

## Remaining n/d (flagged)

| Gap | Why | Watch |
|---|---|---|
| K02, K06–K09 for Q4'25–Q1'26 | No AI MW / GPU counts those quarters | Next ops update with AI MW and GPU fleet |
| K05, K10–K15 on AI perimeter | Capex/D&A/PP&E/EBITDA company-wide | Segment split (unlikely near-term) |
| K08 list price | IREN uses $2.00 assumption, not a public list | — |
| K24–K27 concentration | Not disclosed | 10-K/20-F |
| No new filing since Oct 4 | EDGAR empty 2026-10-04→10-05 | FY27 Q1 / monthly updates |


## Biggest gaps
1. **No AI MW or GPU counts for Dec-25 and Mar-26, and only one AI MW datapoint (Jun-26).** As a result, K02 and K06–K09 stop at Q3'25.
2. **Capex, D&A, PP&E and Adj. EBITDA are company-wide.** This leaves K05 and K10–K15 n/d on the AI perimeter; the GPU-hardware proxies (K44, M06) are used instead.
3. **No customer revenue concentration.** Backlog concentration can only be derived.
4. **No public list price.** K08/K09 use IREN's own $2.00/GPU-hr revenue assumption, so they are not comparable with the list-based CRWV/Nebius figures.
5. **Interest is only partly quarterly.** Contractual interest is annual only; quarterly finance expense includes amortization. The IFRS→GAAP basis break falls at Jun-25.
6. **Post-period:** Aug-26 $2.4B GPU financing at 9.0% fixed. The marginal cost of debt is far above the 1.85% blend.
