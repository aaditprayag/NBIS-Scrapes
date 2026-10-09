# FINAL — YE26 / FY27 billed IT MW vs connected MW (locked for Sunday)
_Thu Oct 8, 2026, ~9:45 AM ET (re-cut v3: FY27 base and high on billed-tranche basis per team decision). Nebius Capacity & Power. Public sources only; nobody contacted. Tags: **R** Reported · **D** Derived · **E** Estimated. Companion: `final_billed_vs_connected.xlsx` (live formulas; tie checks in Final_Table)._

**For judges: connected is not billing.** We estimate ~**941 MW connected** at YE26 [E] (management guide **800 MW–1 GW** [R]) vs only ~**258 MW billing** [D], about **27%**. Revenue runs on billed MW.

## Definitions
- **Connected** [R definition; E numbers]: power delivered into fully built, equipped data-center halls (shell + fit-out ready). It can sit idle until GPUs arrive (Q2'25 letter: "fully provisioned… can be activated immediately upon GPU installation"; Q3'25 letter: "power connected to built data centers"; Q1'26 letter). It is **NOT** the ~5 GW of contracted land + power, **NOT** substation or grid access alone, and **NOT** billing. Gross/shell (facility-power) basis.
- **Active** [R]: power consumed by installed, operational IT.
- **Billed** [D]: revenue-generating IT MW, meaning active IT lagged ~1 quarter. Vineland is counted at delivered Microsoft tranches.

## Company totals (tie to locks)
| Metric | Base | High | Low | Tag | Ties? | Confidence |
|---|---:|---:|---:|---|---|---|
| YE26 connected | **941** | 1,195 | 561 | E (guide 800–1,000 R) | Base ✔ (941), low ✔ (561.2); **high NON-TIE: sites 1,156 vs 1,195** (39 MW unattributed: France 2nd n/d + range) | Low–Med (site splits E; Nebius gives no per-site figure) |
| YE26 energized IT | 301.7 | 414.0 | 189.5 | D | ✔ 301.72 | Med |
| **YE26 billed IT** | **257.7** | **284.2** | — | D | ✔ 257.66 / 284.16 | Med (Vineland tranche MW E; 1Q lag assumption) |
| **FY27 avg billed IT** | **536.5** (536.48) | **663.8** (663.80) | — | D | ✔ both (Vineland at billed tranches, base Bloom timing, in base and high) | Low–Med (Bloom T5–T9, Béthune, Oklahoma ramps; 26.7 IT/tranche is E) |
| *Memo: superseded Oct 8* | *475.0 (474.96)* | *645.4 (645.36): Vineland energized-lagged in high* | | *D* | *Did not match the revenue engine* | |
| *Memo: bull Bloom timing* | | *696.8 (696.83)* | | *D* | *T5–T6 from Q1'27, T7–T9 from Q2'27 (Vineland 220.0)* | |
| YE27 connected | 1,882 | 2,517 | 1,249 | E | company-level only | Low |
| Ratios | billed ÷ connected YE26 **27%**; gap **683 MW**; connected IT-equivalent ÷1.35–1.20 = **697–784 MW**, still ~2.7–3.0× billed | | | D | | |

**Method (D):** YE26 billed uses capacity v2 row 147: each site's IT-normalized energized MW at Q3'26 (1Q lag), plus Vineland at 4 delivered tranches × 26.7 IT (E). FY27 average: **Vineland base uses the billed-tranche basis**, the same as the revenue engine: T1–T4 107 IT MW for the full year ($1,551.5M) plus T5–T9 starting Q2/Q3'27 ($1,159.5M, ~80 MW-yr) = $2,711M ÷ $14.5M = **187.0** (model v0.9d Judge_Slide A19). Every other site uses model v0.9d's quarter-average method: (E_Q3'26 + 2E_Q4'26 + 2E_Q1'27 + 2E_Q2'27 + E_Q3'27)/8. **The Vineland high uses the same 187.0** (base Bloom timing). Other sites' high paths stay on the model method (476.8). IT-normalized means stated MW ÷ Gross/IT factor (1.1–1.35) where the stated figure is treated as gross.

## Site table
| Site | YE26 connected (E) | YE26 energized IT (D, ref) | YE26 billed base (D) | YE26 billed high (D) | FY27 avg billed base (D) | FY27 avg billed high (D) | YE27 connected (E, fixed sites) | Confidence | Reason |
|---|---:|---:|---:|---:|---:|---:|---:|---|---|
| Vineland NJ (DataOne) | 300 | 81.8 | 106.7 | 106.7 | **187.0** (tranche) | **187.0** (tranche) | 300 | Med (YE26) / Low–Med (FY27) | 4 of 9 tranches delivered (R) but MW/tranche is E (26.7 IT); YE26 billed 106.7 (billed-tranche basis) > energized 81.8; FY27 hinges on Bloom (Case B -$966M @~35%) |
| Independence MO (Eastgate) | 100 | 0 | 0 | 0 | 4.6 | 9.3 | n/d | High (billed) / Low (connected) | No IT before ~Q3'27, so YE26 billed 0 is solid; YE26 connected 100 = utility feed/shell plug only |
| Butler Twp PA (Highridge) | 0 | 0 | 0 | 0 | 0 | 0 | n/d | High | PPL 230 kV switchyard/laterals late-2027 to 2029; nothing bills through FY27 |
| Birmingham AL (BHM01, Oxmoor) | 50 | 0 | 0 | 0 | 11.6 | 20.8 | n/d | Med | 115 kV NLT Dec 31 2026, 230 kV Oct 2027 (R); first revenue MW 2027; litigation Mar 2027 |
| Bethune FR (Azur) | 120 | 23.1 | 0 | 0 | 54.5 | 86.5 | n/d | Low-Med | Mgmt says 120 MW by YE26 (R) vs model energized 30; ICPE decision pending; 1Q lag puts YE26 billed at 0 |
| Lappeenranta FI | 0 | 0 | 0 | 0 | 7.6 | 15.2 | n/d | Med | First building H1'27; substation spring 2027 |
| Mantsala FI #1 | 75 | 68.2 | 68.2 | 68.2 | 68.2 | 68.2 | 75 | High | 75 MW built early 2026 (R); IT-normalized at PUE 1.1 (E) |
| Mantsala FI #2 (Kapuli) | 0 | 0 | 0 | 0 | 0 | 0 | n/d | Med | Substation autumn 2027; nothing bills in FY27 on 1Q lag |
| Kansas City MO (Patmos) | 40 | 20 | 20 | 35 | 20 | 35 | n/d | Med | Operational; Nebius share of facility not disclosed (5-35 range) |
| Paris (Equinix PA10) | 5 | 5 | 5 | 5 | 5 | 5 | 5 | Med | Operational; 5 MW is third-party figure |
| Keflavik IS (Verne) | 10 | 10 | 10 | 10 | 10 | 10 | 10 | High | 10 MW live since Mar 2025 (R) |
| London Longcross (Ark LP01) | 19 | 13.5 | 12.3 | 12.3 | 14.0 | 14.3 | 19 | High | 16->19 MW contract (R); stated basis treated as gross (/1.3) |
| Harlow UK (Kao) | 22 | 2.2 | 0 | 0 | 9.9 | 14.8 | n/d | Med | 22 MW contract (R); ramp timing E |
| Newport UK (Vantage CWL1) | 10 | 5 | 0 | 0 | 14.2 | 22.9 | n/d | Low | Nebius MW not disclosed; UK residual 43 MW (D) split E |
| UK 3rd (GM LON-East lead) | 5 | 0 | 0 | 0 | 1.4 | 2.8 | n/d | Low | Site identity unconfirmed (GM 14 MW to unnamed neocloud) |
| Modi'in IL (Mega Or) | 8 | 8 | 8 | 8 | 8 | 8 | 8 | High | 8 MW live Oct 2025 (R) |
| Masmiyya IL (Mega DC) | 22 | 5.5 | 0 | 0 | 11 | 16.5 | n/d | Med | 22 MW contracted (R); ramp E |
| Beit Shemesh IL (Mega DC) | 0 | 0 | 0 | 0 | 15.9 | 22 | n/d | Med | 58 MW; landlord completion Q2'27 (R) |
| Madrid-Getafe ES | 18 | 13.5 | 7.5 | 9 | 15.6 | 16.9 | n/d | Med | 18 MW contract (R); ramp E |
| Tallinn EE (Greenergy Huuru) | 22 | 11 | 0 | 0 | 15.1 | 19.2 | n/d | Med | 22 MW first phase later 2026 (R) |
| Minnesota colo | 15 | 10 | 10 | 15 | 10 | 15 | n/d | Low | Tenant MW n/d; IronGate MSP01 inference |
| Oklahoma colo (unnamed) | 100 | 25 | 10 | 15 | 51.2 | 71.2 | n/d | Low | Map '>100 MW' category only; town/landlord n/d |
| France 2nd colo | n/d | n/d | n/d | n/d | n/d | n/d | n/d | n/d | Region live (R) but site/MW not disclosed; not invented |
| AIB CLT1 (Moore SC) | 0 | 0 | 0 | 0 | 1.6 | 3.1 | n/d | Med | 50 MW critical IT (R, AIB PR Sep 30 2026); halls ~2027 |
| **COMPANY TOTAL (sites; France 2nd n/d excluded)** | 941 | 301.7 | 257.7 | 284.2 | **536.5** | **663.8** | 1,882 (co.) | see checks |  |

Tags per row: connected = E (site split from `ye26_active_vs_connected.xlsx`); billed = D, built on E site paths; Vineland tranches delivered = R, MW per tranche = E. Sources are per row in the xlsx.

## MW needed at $16M (D; Street $12.3B and $16M are inputs, both E)
Street FY27 $12.3B ÷ $16M = **768.75 MW** of FY27 average billed IT needed.

| Path | FY27 avg billed MW | Gap to 768.75 | Needed ÷ path |
|---|---:|---:|---:|
| Base (tranche basis) | 536.5 | **~232** (232.3) | **1.43x** |
| High (tranche basis, base Bloom timing) | 663.8 | **~105** (105.0) | **1.16x** |
| Bull Bloom timing (memo) | 696.8 | **~72** (71.9) | 1.10x |

**Source-of-truth note (per Grok):** the model engine is the source of truth. Its Contracts tab has Microsoft 2027 revenue at ~$2,643.8M, ≈ ~182 MW-yr (billing MW average 185.6). A19's 187.0 overstates that by ~1–5 MW because of the T3–T4 billing ramp and the Q1'27 genset cap. **The final pair may land a few MW lower once Modeling links A19 to Contracts.**

## Basis flags (read before quoting site numbers)
1. **Vineland YE26 billed 106.7 > energized 81.8.** The lock bills 4 delivered tranches × 26.7 IT (E). If the early tranches are smaller (the Q2'26 top-customer 24% check points that way), true billed is nearer ~82. **The deck pair stays 300 connected vs ~82 active** [E / D]. The 218 is only the gap, and ~107 IT is the T1–T4 genset ceiling (62 × 1.982 ÷ 1.2, D).
2. **RESOLVED (Oct 8): FY27 revenue and MW are now both on the billed-tranche basis.** Vineland FY27 is **187.0** [D], replacing 125.5, and the company FY27 base is **536.48** (was 474.96, now kept only as a superseded memo).
   - **Non-Vineland MW** rises from ~288 (the slide's implied 475 − 187) to **~349.5** [D]. That cuts the slide's implied non-Vineland $/MW by **~17.6%**, from ~$33.3M to ~$27.4M (Modeling to confirm the dollars).
   - **Street $12.3B ÷ FY27 base** falls from ~$25.9M to **~$22.9M per billed MW** [D] (Modeling to confirm).
   - **High is now also on the tranche basis (team decision, Oct 8): 663.8.** 645.4 is a superseded memo (Vineland energized-lagged in high) and 696.8 is the bull Bloom-timing memo. The model's A27 memo (646.9) used a quarter-average of quarter-end tranches, which put Vineland at 170.
   - **Contracts-engine cross-check:** Microsoft 2027 billing IT MW averages 185.6. 2027 revenue is $2,643.8M (the T3–T4 billing ramp plus the Q1'27 genset cap at 102.4), which is ≈ 182 MW-yr at $14.49M. That is within ~1–5 MW of 187, so immaterial.
   - **Caveat:** 26.7 IT per tranche is E. If early tranches are smaller, Vineland revenue and MW fall together, so $/MW holds but both levels drop.
3. Connected is on a gross/shell basis and billed on an IT basis, so part of the gap is basis. Even IT-equivalent connected (697–784) is ~2.7–3.0× billed.
4. Only the company totals are locks. Site-level billed is Derived, and Nebius discloses no per-site connected, active or billed MW.

## Changes vs Oct 5
- **Changed (team decision, Oct 8 ~9:45 AM ET):** FY27 average billed base **474.96 → 536.48** (Vineland 125.5 → 187.0, billed-tranche basis). The 475 is kept as a memo row: "superseded Oct 8: energized-lagged basis, did not match revenue engine". FY27 average billed high **645.4 → 663.8** (Vineland in high moved to the tranche basis, base Bloom timing). 645.4 is kept as a superseded memo and 696.8 as the bull-timing memo. MW needed at $16M: 768.75 (gaps ~232 / ~105 / ~72).
- **Unchanged:** YE26 billed (257.7/284.2; already tranche basis for Vineland), energized (302/190/414), and connected (941/561/1,195). So are the Vineland pair, MSFT 4 of 9, and Case B (−$966M @ ~35%).
- **New:** a site-level split of billed MW (YE26 and FY27) that ties to the locks; YE27 connected filled only for fixed/fully built sites (Vineland 300 shell, Mäntsälä 75, Longcross 19, Keflavik 10, Modi'in 8, Paris 5), with the company figure at 1,882 [E]; and flag #2 above (now resolved).
- **Q3 date convention:** late Oct–mid Nov (Oct 22 or Nov 10 per third-party calendars; not company-announced) [E]. Short form: late Oct–mid Nov (est.). The Q3 letter may land before the ~Nov 6–9 NJDEP genset filing deadline.

## Files
- `/workspace/nebius/capacity/final_billed_vs_connected.xlsx` (Final_Table, Quarterly_IT, Definitions, Change_vs_Oct5, Sources)
- `/workspace/nebius/capacity/final_billed_vs_connected.md`
- Build script: `/workspace/nebius/capacity/_build/build_final_billed.py`
