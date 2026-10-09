# Dilution / convert scenarios — one-pager for Modeling

**As of:** Mon Oct 5, 2026 (ET)  
**Audience:** Modeling (wire companion `dilution_scenarios.xlsx`)  
**Public information only.** Tags: **[R]** Reported · **[D]** Derived (math shown) · **[E]** Estimated (labelled placeholder).

**FROZEN for the Sunday pack (Thu Oct 8, 2026, 3:10 PM ET).** Sections A–E are the 6/30 base as written on 10/5. Where they differ, these v0.9e figures supersede them:

- **PT share count:** base FD **358,558,306** [D] (6/30 FD 358,337,958, less the exchanged Jun-29/31 conversion shares, plus the ~15.8M exchange shares; new Feb-30/Feb-34 notes treated as debt). It's **369,564,496** for PTs between $313.46 and $324.65, and **376,648,956** above $324.65 [D]. See `fd_share_bridge.md`.
- **No-vote headroom:** net **~78–97M** [D]. That's ~78M if the exchange shares came from treasury and ~97M if they were newly issued; the settlement source isn't disclosed. This replaces the 93.5M net figure in section B.
- **Funding gap:** case (a) needs ~215M shares of authority, **~2.8×** the ~77.8M net no-vote headroom [D, model v0.9e].
- **Older converts:** **$7,675.0M** face remaining (Jun-29/31 $87.5M each after the Aug exchange). The maturity claim is **$9,051.9M** [D], which matches model v0.9e (see correction at end).
- **EDGAR re-check 3:05 PM ET 10/8:** there's no new issuer 6-K, 424B5/ATM, F-3/S-8, or Q3 date PR since this morning. The only new third-party filing is the GS 13G/A (5.1%, event 9/30), already logged in `edgar_recheck_2026-10-08.md`.

---

## A. Share-count base at 6/30/26

| Item | Shares | Tag | Source / note |
|------|--------|-----|---------------|
| Class A outstanding | **238,400,165** | [R] | Q2 PR / outstanding shares |
| Class B outstanding | **33,455,053** | [R] | Q2 PR (10 votes/share) |
| Class A issued | **288,585,891** | [R] | Q2FS Note 14 |
| Treasury Class A | **50,185,726** | [D] | 288,585,891 − 238,400,165 |
| Issued Class A + Class B | **271,855,218** | [R] | Outstanding shares line; excl. warrant |
| NVIDIA pre-funded warrant (unexercised) | **21,065,936** | [R] | Q2FS Note 14; $0.0001 strike |
| Warrant in **basic WASO** | Yes, from issue date | [R] | Q2FS Note 1 EPS policy — included in basic/diluted WASO even though unexercised |
| Warrant in **issued & outstanding** | No | [R] | Not in the 271,855,218 issued count until exercise; settles from treasury or authorized capital (FY25 20-F) |
| ATM sold in Q2 2026 | **12,729,493** | [R] | Note 14; WA **$223.60** [R]; net **$2,811.1M** [R] |
| ATM remaining (of 25M program) | **12,270,507** | [R] | 25,000,000 − 12,729,493 |
| Q2 basic = diluted WASO (net loss) | **280,381,615** | [R] | = basic; converts anti-dilutive |

---

## B. No-vote Class A funding capacity

From `funding_capacity_note.md` (AGM Items 13–14 PASSED 8/25/26):

| Component | Shares | Tag | Math |
|-----------|--------|-----|------|
| Treasury 6/30/26 | 50,185,726 | [D] | issued − outstanding |
| AGM 20% pool (snapshot) | 64,408,189 ≈ 64.4M | [D] | 20% × (288,585,891 + 33,455,053) = 20% × 322,040,944 |
| **Gross headroom (no new vote)** | **114,593,915 ≈ 114.6M** | [D] | 50,185,726 + 64,408,189 |
| Less unexercised NVIDIA warrant | (21,065,936) | [R] | Settlement from treasury or authorized |
| **Net headroom** | **93,527,979 ≈ 93.5M** | [D] | 114,593,915 − 21,065,936 |

Modeling shorthand: ~**94–115M** Class A funding-capacity shares without a new AGM vote.

---

## C. Convertible notes & dilution scenarios

### Convert stack at 6/30/26 (Q2FS Note 12) — pre Aug-26 exchange/new issue

| Metric | $M | Tag |
|--------|----|-----|
| Carrying amount | **8,499.0** | [R] |
| Accreted principal at maturity | **10,041.8** | [R] |
| Fair value (Level 3) | **20,823.7** | [R] |

| Series | Remaining orig. principal 6/30 ($M) | Conv. rate (sh/$1k) | Conv. price ($) | Note 12 accreted ($M) | Carrying ($M) | Tag |
|--------|-------------------------------------|--------------------|-----------------|-----------------------|---------------|-----|
| 2.00% Jun-2029 | 487.5 | 19.4363 | 51.45 | 587.5 | 504.1 | [R]/[D] remaining |
| 3.00% Jun-2031 | 487.5 | 19.4363 | 51.45 | 612.5 | 498.4 | [R]/[D] remaining |
| 1.00% Sep-2030 | 1581.25 | 7.2072 | 138.75 | 1818.4 | 1588.2 | [R]/[D] remaining |
| 2.75% Sep-2032 | 1581.25 | 7.2072 | 138.75 | 1818.4 | 1575 | [R]/[D] remaining |
| 1.25% Mar-2031 | 2587.5 | 5.4579 | 183.22 | 3105 | 2588.6 | [R]/[D] remaining |
| 2.625% Mar-2033 | 1750 | 5.546 | 180.31 | 2100 | 1744.7 | [R]/[D] remaining |

**If-converted shares excluded from diluted EPS (Q2 loss):** **65,416,804** [R] — Q2FS Note 1: “The weighted average number of incremental shares under the if-converted method for the convertible notes excluded from the computation because their effect was anti-dilutive was **65,416,804** for the three months … ended June 30, 2026.” Cross-check [D]: initial rates × pre-exchange orig. principal ≈ 65.57M (within ~0.2%).

Post-period (not in 6/30 BS): Aug-26 exchange retired $400M of each Jun-2029/2031 note for ~15.8M shares [R]; new $3,450M 0.50% Feb-2030 + $2,300M 4.50% Feb-2034 [R]. Use CapStructure / Converts sheet for post-period waterfall.

### Fully diluted (FD) definition used here

**FD = Class A outstanding + Class B + NVIDIA warrant + if-converted converts**  
= 238,400,165 + 33,455,053 + 21,065,936 + 65,416,804  
= **358,337,958** [D]

(Excludes RSUs/options/ATM remainder and post-6/30 Aug exchange shares — those sit in CapStructure waterfall separately.)

### Scenario table — equity issuance / ATM

Shares issued under each scenario; gross proceeds = shares × price. Dilution % = shares issued ÷ base (Class A out or FD). Headroom scenarios are incremental Class A from treasury + AGM 20% pool (no new vote).

| Scenario | Shares issued | Tag shares | Gross proceeds @ $223.60 [R] ($M) | @ $150 [E] ($M) | @ $100 [E] ($M) | Dilution vs Class A out | Dilution vs FD |
|----------|---------------|------------|----------------------------------|----------------|----------------|-------------------------|----------------|
| Status quo | 0 | — | 0 | 0 | 0 | 0% | 0% |
| ATM exhaust remainder | 12,270,507 | [D]/[R] | 2,743.7 | 1,840.6 | 1,227.1 | **5.1%** | **3.4%** |
| Issue 50% of net headroom (~47M) | 46,763,990 | [D]/[R] | 10,456.4 | 7,014.6 | 4,676.4 | **19.6%** | **13.1%** |
| Issue 100% of net headroom (~94M) | 93,527,979 | [D]/[R] | 20,912.9 | 14,029.2 | 9,352.8 | **39.2%** | **26.1%** |
| Issue 50% of gross headroom (~57M) | 57,296,958 | [D]/[R] | 12,811.6 | 8,594.5 | 5,729.7 | **24.0%** | **16.0%** |
| Issue 100% of gross headroom (~115M) | 114,593,915 | [D]/[R] | 25,623.2 | 17,189.1 | 11,459.4 | **48.1%** | **32.0%** |

**Read of key cells (Modeling wire):**
- ATM exhaust @ $223.60 [R]: **12,270,507** sh → **$2,743.7M** gross; **5.1%** vs Class A / **3.4%** vs FD.
- 100% net headroom (93,527,979) @ $223.60: **$20,912.9M** gross; **39.2%** / **26.1%**.
- 100% gross headroom (114,593,915) @ $223.60: **$25,623.2M** gross; **48.1%** / **32.0%**.
- Same share counts @ $100 [E] halve proceeds vs $200; dilution % is **price-invariant** (share count only).

---

## D. One-line guidance for Modeling

- **PT share count:** use **FD** (Class A out + Class B + warrant + if-converted converts = **358,337,958** [D] at 6/30) for per-share valuation / price-target math; cite **basic** Class A outstanding (**238,400,165** [R]) only when the slide is explicitly “per Class A share outstanding.” Do **not** use WASO alone for PT without stating warrant inclusion.
- **Cash / credit timing stays separate from dilution:** deferred-revenue recognition and customer prepay credit schedules are **not** share-count inputs — see `payback_recon_inputs.md` / `deferred_rpo_rollforward.xlsx`. Do not bake recognition timing into the dilution waterfall.

---

## E. Gaps

- **Post-6/30 ATM** sales: not yet in balance sheet / Note 14; remaining capacity may already be lower.
- **NVIDIA warrant exercise:** undisclosed through latest 6-Ks (through 10/1/26 Inferize); still treated as unexercised.
- **$775M senior secured facility** (signed 10-Jul-2026): drawn amount **undisclosed**.
- Aug-26 ~15.8M exchange shares and new converts are **post-period** vs this 6/30 base — refresh after next interim.

---

## File pointers

- Companion workbook: `dilution_scenarios.xlsx`
- Funding math: `funding_capacity_note.md`
- Converts / CapStructure: `nebius_filings.xlsx`
- Fact base: `nebius_fact_base.md`

**Correction (10/8/26):** the Jun-2029 / Jun-2031 "Note 12 accreted" figures ($587.5M / $612.5M) are the redemption amounts on the **$500M original** issue (117.5% / 122.5%), not on the $487.5M remaining. On the $87.5M left after the Aug exchange, the claims at maturity are $102.8M / $107.2M [D]. The older six total **$9,051.9M** [D], which matches model v0.9e.

---
*Change log 2026-10-09 (1): Funding-gap ratio restated text-only from "~215M new shares, ~2.2–2.8× the no-vote headroom" to "~215M shares of authority, ~2.8× the ~77.8M net no-vote headroom" (team lock 10/9: authority needed vs 77.8M available, net of NVIDIA warrant; Grok OK 8:47 AM). No numbers changed.*
