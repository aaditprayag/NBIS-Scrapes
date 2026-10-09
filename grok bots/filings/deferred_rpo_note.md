# Deferred revenue & RPO — finding note (NBIS SHORT)

**As of:** Mon Oct 5, 2026 (ET)  
**Workbook:** `deferred_rpo_rollforward.xlsx`  
**Tags:** [R] Reported · [D] Derived · [E] Estimated

---

## Finding (lead)

**Strategic prepays are structurally long-dated, not a near-term revenue reservoir.** At 6/30/26 deferred revenue was **$5,975.2M** [R] (current **$979.4M** + non-current **$4,995.8M**), yet only **$130.8M** [R] of Q2’26 revenue came from the opening deferred balance — a **2.7%** [D] quarterly release rate on the **$4,778.1M** [R] Q1 close / Q2 open. The company’s own current/non-current split says only **~16.4%** [D] of the stock is expected to recognize within 12 months. That pattern is **capacity/delivery-linked recognition over 1–5 years** [R note], not “cash arrived on the last day of the quarter” alone, and **not** a disclosed IREN-style after-month-24 credit schedule (Nebius has not disclosed one).

**RPO has no ≤12-month split.** GAAP RPO at 6/30/26 was **$37,490.6M** [R]: **36%** in 24 months ending 6/30/28 (~**$13.5B** [D]), **40%** in months 25–48 (~**$15.0B** [D]), remainder thereafter; excludes original duration ≤1 year. Call/letter “~$40B backlog/commitments” is management rounding around that figure. Versus Street FY27 ~**$12.3B** [E] and FY26 guide **$3.0–3.4B** [R], a flat 12/24 cut of the 0–24m bucket implies only ~**$6.7B** [E] contracted into calendar 2027 — leaving a multi-billion unsigned/delivery-mix gap to Street if that crude cut holds.

---

## Deferred roll-forward (total, $M)

| Quarter | Open | Adds¹ | Recog from open | Close | Release rate | Current / Non-current |
|---------|------|-------|-----------------|-------|--------------|------------------------|
| Q4’24 | 17.4 [R] | −0.9 [D] net | ND | 16.3 [R] | ND | 16.3 / — |
| Q1’25 | 16.3 [R] | 2.7 [D] net | ND | 19.0 [R] | ND | 19.0 / — |
| Q2’25 | 19.0 [R] | 3.5 [D] | 3.2 [R] | 19.3 [R] | **16.8%** [D] | 19.3 / — |
| Q3’25 | 19.3 [R] | −3.7 [D] net | ND | 15.6 [R] | ND | 15.6 / — |
| Q4’25 | 15.6 [R] | 1,561.9 [D] net | ND | 1,577.5 [R] | ND | 275.5 / 1,302.0 [R] |
| Q1’26 | 1,577.5 [R] | 3,224.1 [D] | 23.5 [D]² | 4,778.1 [R] | **1.5%** [D] | 685.6 / 4,092.5 [R] |
| Q2’26 | 4,778.1 [R] | 1,327.9 [D] | 130.8 [R] | 5,975.2 [R] | **2.7%** [D] | 979.4 / 4,995.8 [R] |

¹ Adds = Close − Open + Recog − Other. Where Recog is ND, Recog is treated as 0 so Adds = net change (gross prepays understated).  
² Q1’26 recog = 1H’26 $154.3 − Q2’26 $130.8 [D] from Q2FS Note.

**Sources:** YE25 20-F Deferred Revenue / RPO note; Q1’26 & Q2’26 interim FS Note (~p.F-15 / F-18); quarterly 6-K PR BS/CF (`tm2427209`, `tm257185`, `tm2515580`, `tm2522866`, `tm2530882`, `tm266173`).

---

## Credit-timing inference

| Hypothesis | Verdict | Evidence |
|------------|---------|----------|
| Prepays credited only as capacity goes live / POs fulfilled | **Supported (high)** | Note: deferred when cash received “before performance obligations are fulfilled”; expected recognition **1–5 calendar years** [R]; MSFT non-upfront invoiced monthly through Oct-2031 [R]. |
| Large prepays late in quarter vs undelivered capacity | **Partially explanatory (medium)** | Q1–Q2’26 adds >> recog; ending stock jumped on prepay waves. Explains **flow** mismatch in the arrival quarter, **not** why **84%** of stock is non-current. |
| Back-loaded credit after month 24 (IREN–MSFT peer) | **Not evidenced for NBIS (low)** | No Nebius credit-vs-billing schedule in filings/letters/calls (`prepay_credit_schedule_search.md`). Peer analogy only. |
| Tiny release ⇒ deferred is “fake” / won’t convert | **Overstated bear (for SHORT: use carefully)** | Current portion **$979.4M** [R] is real near-term; unbilled RPO **~$31.5B** [D] is the bigger contracted stock. Issue for SHORT is **timing/execution** into Street FY27, not absence of contracts. |

**Confidence:** High on long-dated recognition horizon from current/NC split; medium on delivery-linked mechanism; none on a specific month-24 credit schedule.

---

## RPO timing (or lack of ≤12-month split)

| As-of | RPO $M [R] | 0–24m | 25–48m | Thereafter | ≤12m split? |
|-------|------------|-------|--------|------------|-------------|
| 12/31/25 | 21,333.0 | 28% (~5,973 [D]) | 39% (~8,320 [D]) | 33% (~7,040 [D]) | **None** |
| 3/31/26 | 33,585.3 | 29% (~9,740 [D]) | 39% (~13,098 [D]) | 32% (~10,747 [D]) | **None** |
| 6/30/26 | 37,490.6 | 36% (~13,497 [D]) | 40% (~14,996 [D]) | 24% (~8,998 [D]) | **None** |

Practical expedient: excludes original duration ≤1 year; includes estimated variable consideration [R]. 6/30/26 fully AI cloud [R].

**Best near-term proxies when no ≤12m RPO exists:** current deferred **$979.4M** [R]; MSFT monthly remainder **~$10.4B** [D] through Oct-2031; Meta 2026 dedicated **$12B** [R] from early 2027.

---

## FY27 coverage math (crude)

- Street FY27 ~**$12.3B** [E] · FY26 guide **$3.0–3.4B** [R] · 1H26 actual **$981.3M** [R] → implied 2H26 mid ~**$2.2B** [D].  
- 0–24m RPO ~**$13.5B** [D] through 6/30/28 includes 2H26 + FY27 + 1H28.  
- Flat 12/24 of that bucket → FY27 contracted ~**$6.7B** [E] ≈ **55%** of Street [E] → ~**$5.6B** still unsigned / mix / steeper ramp [E].  
- If recognition inside 0–24m is back-loaded (consistent with only 16% of deferred classified current), FY27 contracted share rises — **modeling assumption, not disclosed**.

---

## Gaps

- No quarterly recognized-from-opening for Q4’24, Q1’25, Q3’25, Q4’25 (2025 PRs lack full interim notes).  
- No invoice-level prepayment credit schedule.  
- No ≤12-month / 1–2 / 2–5 / >5 RPO buckets.  
- Q4’24 open $17.4 unrestated vs YE24 close $16.3 restated — immaterial bridge.  
- “$40B” call language vs $37.5B GAAP RPO — not reconciled by the company.
