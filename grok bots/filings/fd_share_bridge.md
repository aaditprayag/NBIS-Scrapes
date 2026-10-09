# Fully diluted share bridge: 6/30/26 → pro forma post-Aug-26 (as of Thu Oct 8, 2026)

**Owner:** Filings & Capital · **Public information only** · Workbook: `fd_share_bridge.xlsx` (live formulas; inputs in blue on `Inputs`; sheets `FD_Bridge`, `PT_Consistency`, `Headroom`)
**Tags:** **[R]** Reported · **[D]** Derived (math shown) · **[E]** Estimated / assumption · **ND** = not disclosed
**FD definition (team):** Class A outstanding + Class B + NVIDIA warrant + if-converted shares on all convertible notes. RSUs and options are excluded (memo below).

## A. Bridge

| # | Step | Shares | Tag | Source / math |
|---|---|---:|---|---|
| 1 | Class A outstanding 6/30/26 | 238,400,165 | R | Q2 PR (6-K 8/12/26, acc 0001104659-26-094568) |
| 2 | Class B outstanding 6/30/26 | 33,455,053 | R | Q2 PR |
| 3 | NVIDIA pre-funded warrant (unexercised, $0.0001) | 21,065,936 | R | Q2FS Note 14 (acc 0001104659-26-094844) |
| 4 | If-converted, all convert series (Q2 weighted avg, anti-dilutive) | 65,416,804 | R | Q2FS Note 1 |
| **FD0** | **FD at 6/30/26** | **358,337,958** | **D** | Sum of rows 1–4 |
| 5 | Less: conversion shares, $400M Jun-2029 2.00% notes exchanged | (7,774,520) | D | 400,000 bonds × 19.4363 [R rate] |
| 6 | Less: conversion shares, $400M Jun-2031 3.00% notes exchanged | (7,774,520) | D | 400,000 bonds × 19.4363 [R rate] |
| 7 | Plus: Aug exchange shares, 2029 notes | 7,873,876 | D (day count E) | 400,000 × ($4,403.18 [R] + $1,000 × 2.00% × 76/360) ÷ $223.90 [R], rounded down |
| 8 | Plus: Aug exchange shares, 2031 notes | 7,895,512 | D (day count E) | 400,000 × ($4,413.18 [R] + $1,000 × 3.00% × 76/360) ÷ $223.90 [R], rounded down |
| 9 | Plus: if-converted, new $3,450M 0.50% Feb-2030 notes | 11,006,190 | D (rate R) | 3,450,000 bonds × 3.1902 [R] ($313.46) |
| 10 | Plus: if-converted, new $2,300M 4.50% Feb-2034 notes | 7,084,460 | D (rate R) | 2,300,000 bonds × 3.0802 [R] ($324.65) |
| 11 | Post-10/5/26 changes disclosed | 0 | D | EDGAR + newsroom re-check 10/8: none (`edgar_recheck_2026-10-08.md`) |
| **FD1** | **FD pro forma after Aug-26 transactions** | **376,648,956** | **D** | FD0 + rows 5–11 |
| | Change vs 6/30 FD | +18,310,998 (+5.11%) | D | |

**Exchange shares (rows 7+8): 15,769,388 [D]** vs company "approximately 15.8 million" [R].
- **Formula source:** Form of Exchange Agreement, Ex.4.5 to 6-K 8/24/26 (acc 0001104659-26-100347). Shares = (Notes Exchange Price × amount + accrued and unpaid interest from 6/5/26 to settlement) ÷ Share Price. Share Price = last reported sale price on the pricing date = $223.90 on 8/19/26 [R, 6-K 8/20/26]. Rounded down once per series [R].
- **Assumption flagged [E]:** 30/360 day count and T+2 settlement 8/21 (76 days). With 8/24 settlement (79 days), the count rises by **744 shares**, which is immaterial.
- **Inducement vs straight conversion:** rows 5–8 net to **+220,348** shares [D].
- **Reported-only alternative:** using the company's ~15.8M instead of the derived count gives FD ≈ **376,679,568** [R/D].

**Memo (excluded from FD; footnote):**
- **Unvested RSUs 6,234,091** [R] and **options 6,740,600** [R] (WAEP $88.65) at 6/30 (Q2FS Note 13). The team FD excludes both. A gross upper bound including both (no treasury-stock method) would be **389,623,647** [D].
- **Point-in-time cross-check of row 4:** 6/30 principal × rates ≈ 65,570,979 [D], vs the 65,416,804 weighted average [R] (~0.2% higher).
- **Q3 baseline Class A** ≈ 238,400,165 + 15,769,388 = **254,169,553** [D], before ATM and RSU deliveries.

## B. PT bridge consistency (shares ↔ net debt)

**Rule:** if a note series is counted as if-converted shares, remove that series from debt. If a series is kept as debt, drop its shares. Team default = all series if-converted (FD 376.6M).

| Item | $M | Tag | Treatment |
|---|---:|---|---|
| Convertible notes carrying, 6/30 (all old series) | 8,499.0 | R | Remove from net debt (old series are in the 65.4M) |
| of which exchanged $800M (pro-rata carrying: 400/487.5 × (504.1 + 498.4)) | 822.6 | D | **Already gone post-Aug** (retired for shares, no cash). Don't remove it twice when rolling the balance sheet forward |
| Old converts remaining carrying post-exchange | 7,676.4 | D | Remove (if-converted) |
| **New notes face to remove from debt** ($3,450M + $2,300M) | **5,750.0** | R | Q3 carrying value ND. Accreted at maturity = 3,450 × 1.10 + 2,300 × 1.25 = **6,670.0** [D] (claim if not converted) |
| **Cash to KEEP: new-note net proceeds** | **~5,680** | R (company estimate) | "approximately $5.68 billion if the initial purchasers fully exercise" [R, 6-K 8/20]; fully exercised [R, 6-K 8/24]. Exact net ND |
| Exchange cash | 0 | R | Notes-for-shares; accrued interest settled in shares |
| Capped calls / hedges | 0 | R | **None** on any series |
| Keep as debt: $775M secured facility (drawn ND) | 775 | R | Full draw is an assumption [E] |
| Keep as debt: current debt 6/30 | 46.7 | R | |
| Pro forma cash (8,042.1 + ~5,680) | ~13,722.1 | D | Excl. restricted 1,056.0 and Q3 OCF/capex |

**Moneyness at $237.15** (10/7 close [R, third-party quote]): Jun-29/31 ($51.45), Sep-30/32 ($138.75), Mar-31 ($183.22) and Mar-33 ($180.31) are in the money. The **new Feb-2030 ($313.46) and Feb-2034 ($324.65) are out of the money** [D].
- **Alternative convention:** treat the new notes as debt, so FD = **358,558,306** [D] and debt carries the $5.75B face (accreting to $6.67B).
- Whichever convention the model uses, it must use it on both the share and the debt side.

## C. Net no-vote headroom (restated)

| | (i) Exchange shares from treasury | (ii) Exchange shares newly issued | Tag |
|---|---:|---:|---|
| Class A issued | 288,585,891 | 304,355,279 | R / D |
| Treasury | 34,416,338 | 50,185,726 | D |
| AGM 20% pool = 20% × (Class A issued + Class B 33,455,053) | 64,408,189 | 67,562,066 | D |
| Gross headroom | 98,824,527 | 117,747,792 | D |
| Less: unexercised NVIDIA warrant | (21,065,936) | (21,065,936) | R |
| **Net no-vote headroom** | **77,758,591 (~78M)** | **96,681,856 (~97M)** | D |

**Reasoning:**
- The exchange settled **8/24/26, one day before the 8/25 AGM authority** [R]. Whether those shares came from treasury or new issuance is **ND**.
- Q2 ATM shares were delivered **from treasury** [R, Q2FS Note 14]. That precedent makes case (i), **~78M**, the conservative planning figure.
- Case (ii), **~97M**, applies only if the shares were newly issued under the prior authority, which raises issued capital and therefore the 20% pool.
- Both cases fall share-for-share with any Q3 ATM sales (ND) and RSU/option deliveries from treasury (ND, small).
- Compared with the team's 6/30 snapshot of 93.5M, the conservative case is ~15.8M lower.
- **The Q3 interim treasury count will resolve this.**

## D. Change vs 376.6M

**No change.** The 10/8 re-check found no issuer share-count disclosures since 10/5. The GS 13G/A (9/30, 5.1%) implies a 252.6M–257.6M denominator [D]. That is consistent with the ~254.2M pro forma Class A but is third-party data, not an issuer figure.
