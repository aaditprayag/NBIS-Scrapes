# Payback_Recon inputs — handoff for Modeling

**As of:** Mon Oct 5, 2026 (ET)  
**From:** Filings & Capital → **Modeling Payback_Recon**  
**Primary workbook:** `deferred_rpo_rollforward.xlsx`  
**Companion note:** `deferred_rpo_note.md`  
Tags: **[R]** Reported · **[D]** Derived · **[E]** Estimated · **ND** = not disclosed.

---

## Pull exactly these from `deferred_rpo_rollforward.xlsx`

### 1. Sheet `Deferred_Rollforward` — quarterly open / adds / recog / close ($M)

| Quarter | Open [R] | Adds [D] | Recog from open | Close [R] |
|---------|----------|----------|-----------------|-----------|
| Q4’24 | 17.4 | net (ND recog→0) | ND | 16.3 |
| Q1’25 | 16.3 | net | ND | 19.0 |
| Q2’25 | 19.0 | 3.5 | **3.2 [R]** | 19.3 |
| Q3’25 | 19.3 | net | ND | 15.6 |
| Q4’25 | 15.6 | net (~1,561.9) | ND | 1,577.5 |
| Q1’26 | 1,577.5 | 3,224.1 | **23.5 [D]** (= 1H 154.3 − Q2 130.8) | 4,778.1 |
| Q2’26 | 4,778.1 | 1,327.9 | **130.8 [R]** | 5,975.2 |

Use live formulas in the xlsx (Adds / Check columns) — do not hardcode adds if the sheet already computes them.

### 2. Current vs non-current (same sheet, lower block)

| Quarter | Current [R] | Non-current [R] | Total | Current % [D] |
|---------|-------------|-----------------|-------|---------------|
| Q4’25 | 275.5 | 1,302.0 | 1,577.5 | — |
| Q1’26 | 685.6 | 4,092.5 | 4,778.1 | — |
| Q2’26 | **979.4** | **4,995.8** | **5,975.2** | **16.4%** |

Q4’24–Q3’25: current ≈ total (NC = 0 on PR BS).

### 3. Sheet `Release_Rate`

| Quarter | Opening | Recog | Release rate [D] |
|---------|---------|-------|------------------|
| Q2’25 | 19.0 | 3.2 [R] | **16.8%** |
| Q1’26 | 1,577.5 | 23.5 [D] | **~1.5%** |
| Q2’26 | 4,778.1 | 130.8 [R] | **~2.7%** |

Stock-view ≤12m proxy: Current / Total at 6/30/26 = **16.4%** [D from R].

### 4. Sheet `RPO_Timing` — 0–24 / 25–48 / later ($ and %)

| As-of | RPO $M [R] | 0–24m % [R] | 0–24m $ [D] | 25–48m % [R] | 25–48m $ [D] | Thereafter |
|-------|------------|-------------|-------------|--------------|--------------|------------|
| 12/31/25 | 21,333.0 | 28% | ~5,973 | 39% | ~8,320 | ~33% / ~7,040 |
| 3/31/26 | 33,585.3 | 29% | ~9,740 | 39% | ~13,098 | ~32% |
| 6/30/26 | **37,490.6** | **36%** | **~13,497** | **40%** | **~14,996** | **~24% / ~8,998** |

**No ≤12-month RPO split** disclosed. Unbilled RPO at 6/30/26 ≈ **31,515.4** [D] (= 37,490.6 − 5,975.2).

### 5. Sheet `FY27_Coverage` (optional context)

Street FY27 ~$12.3B [E]; FY26 guide $3.0–3.4B [R]; crude even 12/24 of 0–24m bucket ≈ **$6.7B** [E] contracted into calendar 2027.

---

## Team lock — cash-credit timing (do not override)

| Statement | Status |
|-----------|--------|
| **Even-from-go-live** recognition fits Q2 evidence (tiny release on multi-billion open; 16.4% current) | **Team lock** — use for Payback_Recon base |
| Cash-credit schedule (how prepays credit against invoices — pro rata, after month 24, etc.) | **Not disclosed** for Nebius |
| Any specific back-loaded credit curve (incl. IREN peer analogy) | **Labelled [E] only** — not evidenced in NBIS filings |
| Conclusion for Modeling | **Cash-credit timing is NOT evidenced as back-loaded.** Do not force an IREN-style after-month-24 credit into Payback_Recon as [R]/[D]. |

Detail: `deferred_rpo_note.md` § Credit-timing inference; `prepay_credit_schedule_search.md`.

---

## Still missing (do not invent)

- Recognized-from-opening for Q4’24, Q1’25, Q3’25, Q4’25 (ND → Adds = net change only).
- ≤12-month / 1–2 / 2–5 / >5 RPO buckets.
- Invoice-level prepayment credit schedule.
- Same-period additions recognized within the quarter (never disclosed separately).
- Company reconciliation of call “~$40B” vs GAAP RPO $37.5B.

---

## Dilution vs Payback

Share-count / convert dilution (`dilution_scenarios.xlsx`) is **orthogonal** to deferred/RPO timing. Do not mix cash-credit assumptions into the dilution waterfall.
