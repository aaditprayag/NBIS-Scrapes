# Post-6/30/26 capital update: ATM, NVIDIA warrant, $775M facility, other issuance

**As of:** Mon Oct 5, 2026, ~1:00 PM ET · **Owner:** Filings & Capital · **Public information only**
**Tags:** **[R]** Reported (source given) · **[D]** Derived (math shown) · **[E]** Estimated (labelled) · **ND** = not disclosed

**Where I looked (all items):** the live EDGAR submissions API `https://data.sec.gov/submissions/CIK0001513845.json`, pulled 10/5/26 with a descriptive User-Agent. The full filing list (498 filings, no overflow pages) covers every 6-K, 6-K/A, 424B*, F-3*, S-8, 13G/13D and 20-F filed since 6/25/26. I also checked local copies of every post-6/30 6-K (`edgar/`, `edgar/post/`), the nebius.com newsroom and financials pages (fetched 10/5), the Q2'26 call transcript, the 9/8/26 Goldman Communacopia transcript, and web/news searches.

**Every filing since 6/30/26:** 6-K 7/17 + 6-K/A 7/17 ($775M facility) · 13G 7/20 (NVIDIA) · 6-K 7/27 (AGM notice) · 13G 8/7 + 13G/A 8/7 (Goldman Sachs) · 6-K 8/12 ×2 (Q2 results/letter; interim FS) · 6-K 8/19 (convert launch) · 6-K 8/20 (pricing) · 6-K 8/24 (closing + indentures + form of exchange agreement) · 6-K 8/26 (AGM results) · 6-K 9/8 (Palantir) · 6-K 10/1 (Inferize) · Forms 3/4/144 (insiders). There is **no 424B\*, F-3, S-8 or new prospectus supplement after 11/12/2025**.

---

## Summary table

| Item | Status | Number + source | Impact on FD 358.3M / headroom |
|---|---|---|---|
| (a) ATM sales after 6/30 | **Not disclosed** (none reported); indirect evidence of **~zero sales 7/1–7/28** | 12,270,507 shares remaining at 6/30 [R, Q2FS Note 14]. AGM record date **7/28/26** [R, AGM notice 6-K 7/27]: Class A eligible to vote **238,402,543** [R, AGM results 6-K 8/26] vs **238,400,165** outstanding 6/30 [R] → **+2,378** [D]. ATM shares are delivered from treasury [R, Note 14], so any July ATM sale would have raised this count. | None proven. Remaining capacity is still 12.27M [R] unless Q3 sales took place 7/29–9/30 (**ND**; first visible in the Q3 interim) |
| (a) New ATM program / prospectus supplement | **Not disclosed / none filed** | The only ATM supplement is the 424B5 of **11/12/2025** (up to 25,000,000 Class A; agents GS, MS, BofA, Citi; under F-3ASR 333-286932) [R]. No 424B5 since. | None |
| (b) NVIDIA warrant exercise / 13G changes | **Not disclosed** (no exercise reported; no NVIDIA 13G/A) | 13G 7/20/26: **22,256,412** sh beneficial (9.3%) = 1,190,476 shares + **21,065,936** warrant shares; exercise and sale barred **before 9/11/26** [R]. Exercise also conditioned on HSR clearance or an exemption [R, Q2FS Note 14]. | **FD-neutral**: the warrant is already in FD and basic WASO. Exercise would use treasury or authorized shares, and the 21.07M is already netted in the 93.5M net headroom. |
| (c) $775M facility draw / amendment | **Drawn amount not disclosed; no amendment filed** | Signed 7/10/26, ~$775M senior secured term loan, Term SOFR+2.50%, matures 10/31/2030, DSCR 1.15x + minimum liquidity [R, 6-K/A 7/17]. PR headline "Nebius raises $775 million" [R]. Q2FS Note 16 (subsequent events) gives no draw [R]. | No share impact. Debt rises by up to $775M (draw **ND**; full draw is an assumption [E]) |
| (d) Aug-26 exchange: exact share count | **Exact count not disclosed**; "approximately 15.8 million" [R] | **Derived ≈ 15,769,388** shares (range 15.769–15.770M depending on settlement date) from the Ex.4.5 formula [D, below] | FD **+220k** net from the exchange alone [D] |
| (d) New converts (Aug-26) | **Disclosed** | $3,450M 0.50% Feb-2030 (3.1902 sh/$1k) + $2,300M 4.50% Feb-2034 (3.0802 sh/$1k) [R, 6-K 8/20 and 8/24] | FD **+18,090,650** if-converted [D] |
| (d) Other equity/debt issuance | **None found** | No other 6-K, 424B or registration statement. Goldman Sachs 13G (6/30: 15,891,853 sh, 7.2%) and 13G/A (7/31: 23,104,462 sh, 10.5%) [R] are GS's *own* holdings (shared power; GS is an ATM agent). **Not** an issuer share-count disclosure. | None |
| Other post-6/30 capital-relevant | AIB Data Centers 50 MW, 12-yr capacity agreement (9/30/26, **AIB** press release; no Nebius 6-K): Nebius **prepays** AIB [R, AIB PR] | Prepayment amount **ND** | Cash use / lease commitment, not shares |

---

## (d) Aug-2026 exchange share count: derivation [D]

Source: Form of Exchange Agreement, Ex.4.5 to 6-K 8/24/26 (acc 0001104659-26-100347), https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex4-5.htm

- "Notes Exchange Price" [R]: **$4,403.18** per $1,000 of 2029 Notes; **$4,413.18** per $1,000 of 2031 Notes.
- Shares = (Exchange Price × amount + accrued interest from 6/5/26 to settlement) ÷ Share Price. Share Price = last reported sale price on the pricing date [R] = **$223.90** on 8/19/26 [R, pricing PR 6-K 8/20].
- Accrued interest uses 30/360 [E, standard convention]; 76 days to 8/21 (T+2) or 79 days to 8/24 (PR settlement date).
  - 2029: 400,000 × (4,403.18 + 4.22) ÷ 223.90 = **7,873,876**
  - 2031: 400,000 × (4,413.18 + 6.33) ÷ 223.90 = **7,895,512**
  - **Total ≈ 15,769,388** (15,770,132 at 79 days), which matches the "approximately 15.8 million" [R]
- Versus straight conversion at 19.4363 sh/$1k: 800,000 × 19.4363 = **15,549,040** [D] → inducement ≈ **220k shares** (~$51–61 per $1,000 bond above conversion value at $223.90) [D].
- Remaining Jun-2029 / Jun-2031 principal after exchange: **$87.5M each** (487.5 − 400) [D] → 3,401,353 if-converted shares [D].

---

## FD impact (pro forma for Aug-26 transactions) [D]

| Step | Shares | Tag |
|---|---:|---|
| FD at 6/30/26 (A 238,400,165 + B 33,455,053 + warrant 21,065,936 + if-converted 65,416,804) | 358,337,958 | [D] team base |
| − If-converted shares of the $800M exchanged notes (800,000 × 19.4363) | (15,549,040) | [D] |
| + Exchange shares issued (formula above) | +15,769,388 | [D] |
| + New 2030 notes if-converted (3,450,000 × 3.1902) | +11,006,190 | [D from R rates] |
| + New 2034 notes if-converted (2,300,000 × 3.0802) | +7,084,460 | [D from R rates] |
| **FD pro forma after Aug-26** | **376,648,956 ≈ 376.6M** | **[D]** |

Net change **+18.3M (+5.1%)** [D]. Caveats:
1. The 65,416,804 is a Q2 *weighted-average* if-converted figure. A point-in-time 6/30 count is ~65.57M (team cross-check), which would add ~0.15M.
2. Excludes Q3 ATM (ND), RSUs/options (6,234,091 RSUs + 6,740,600 options at 6/30 [R]) and Class B→A conversions (FD-neutral).
3. Conversion counts use the initial rates on *original* principal [R]. Accretion does not change the share count; settlement is cash/shares at the company's election [R].

**Class A outstanding pro forma** (no Q3 ATM) ≈ 238,400,165 + 15,769,388 = **254,169,553** [D]. This is the Q3 baseline: anything above ~254.2M (plus small RSU/option deliveries) at 9/30 means ATM or other issuance.

---

## Headroom impact (no-vote authority) [D]

The exchange shares settled **8/24/26, one day before the 8/25 AGM** authority took effect. Whether they came from **treasury** or new issuance is **ND**. (Q2 ATM shares came from treasury [R]; warrant settlement can be either [R].)

| Scenario | Treasury | 20% pool | Gross headroom | Net of warrant |
|---|---:|---:|---:|---:|
| Team base (6/30 snapshot) | 50,185,726 | 64,408,189 | 114,593,915 | 93,527,979 |
| (i) Exchange shares from treasury [D] | 34,416,338 | 64,408,189 | **98,824,527** | **77,758,591** |
| (ii) Exchange shares newly issued (pool = 20% × (304,355,279 + 33,455,053)) [D] | 50,185,726 | 67,562,066 | **117,747,792** | **96,681,856** |

**Read:** no-vote net headroom is **~78M–97M** [D] depending on settlement source (ND), versus 93.5M at the 6/30 snapshot. For the short, use the **lower ~78M** as the conservative capacity figure until the Q3 interim shows treasury. Any Q3 ATM sales come off further, share for share.

---

## Sources (URLs)

- Live EDGAR submissions JSON (pulled 10/5/26): https://data.sec.gov/submissions/CIK0001513845.json
- ATM 424B5 (11/12/25): https://www.sec.gov/Archives/edgar/data/1513845/000110465925110025/tm2530771-1_424b5.htm
- Q2'26 interim FS (Note 14 ATM, warrant; Note 16 subsequent events): https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm
- AGM notice (record date 7/28/26): 6-K 7/27/26, acc 0001104659-26-087080 (Ex.99.1, local `edgar/tm2621374d1_ex99-1.htm.txt`)
- AGM results (238,402,543 Class A eligible): https://www.sec.gov/Archives/edgar/data/1513845/000110465926101076/tm2624077d1_ex99-1.htm
- NVIDIA 13G 7/20/26: https://www.sec.gov/Archives/edgar/data/1513845/000104581026000062/xslSCHEDULE_13G_X01/primary_doc.xml
- Goldman 13G / 13G/A 8/7/26: https://www.sec.gov/Archives/edgar/data/1513845/000088698226000305/ ; https://www.sec.gov/Archives/edgar/data/1513845/000088698226000308/
- $775M facility 6-K/A 7/17/26: https://www.sec.gov/Archives/edgar/data/1513845/000110465926084606/tm2620779d1_6ka.htm
- Convert pricing 6-K 8/20/26: https://www.sec.gov/Archives/edgar/data/1513845/000110465926098924/tm2623617d1_ex99-1.htm
- Convert closing + Ex.4.5 6-K 8/24/26: https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex99-1.htm
- AIB–Nebius 50 MW (9/30/26): https://www.globenewswire.com/news-release/2026/09/30/3371855/0/en/aib-data-centers-signs-contract-with-nebius-for-ai-data-center-capacity.html

**Confidence:** High that no ATM program, prospectus supplement, warrant exercise, facility amendment or other issuance was *filed* through 10/5/26. High on the derived 15.77M exchange count (formula is Reported; only the interest day count is assumed). Medium on "zero July ATM" (record-date inference; 7/29–9/30 not observable).
