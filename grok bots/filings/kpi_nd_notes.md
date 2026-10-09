# KPI n.d. notes — remaining blanks in nebius_kpi_history.xlsx

**As of:** Mon Oct 5, 2026 (ET)  
**Owner:** Filings & Capital  
**Rule:** Do **not** invent fill values. Rebuild via `kpi/kpi_build.py` only when real Reported cells are added.

Tags: **[R]** Reported · **[D]** Derived · **[E]** Estimated · **n.d.** = not disclosed.

---

## 1. RPO before YE25

| Question | Answer |
|----------|--------|
| First disclosed RPO amount | **$21,333.0M** as of **12/31/25** [R] |
| First disclosure document | YE25 Form **20-F** (filed ~2026-04-30), Note — Remaining performance obligations: “As of December 31, 2025, the aggregate transaction price allocated to the RPO was **$21,333.0**, of which **28%** is expected to be recognized over the 24 months ending December 31, 2027, **39%** between months 25 and 48, and the remainder thereafter.” |
| Path | `edgar/nbis-20251231x20f.htm.txt` |
| Earlier 6-Ks / FY24 20-F | **No** RPO dollar amount found. Local EDGAR text search for “remaining performance obligation” hits only: YE25 20-F, Q1’26 interim FS, Q2’26 interim FS. FY24 20-F (`nbis-20241231x20f.htm.txt`) and 2025 quarterly PRs/letters (Q1–Q3’25) do **not** disclose an RPO figure. |
| KPI grid | Row “Remaining performance obligations”: **n.d.** for Q1’24–Q3’25; populated YE25 / Q1’26 / Q2’26. **Leave n.d.** — no backfill. |

**So-what:** Pre-YE25 RPO is a permanent n.d. for the KPI history, not a missing pull.

---

## 2. Mid-2025 PP&E Note 7 (servers / ANIU roll)

| Period | Servers & network equipment, gross | Assets not yet in use (ANIU) | Disclosed? |
|--------|-------------------------------------|------------------------------|------------|
| YE24 (12/31/24) | **794.9** [R] | **163.8 / 163.9** [R] | Yes — FY24 20-F Note (PPE) |
| Q1’25 | n.d. | n.d. | **No** — Q1’25 PR/letter only; no interim FS Note 7 roll |
| Q2’25 / H1’25 | n.d. | n.d. | **No** — Q2’25 letter + H1’25 OFR discuss depreciation narrative only; **no** servers/ANIU table |
| Q3’25 | n.d. | n.d. | **No** — Q3’25 PR/letter; no Note 7 breakdown |
| YE25 (12/31/25) | **3,123.6** [R] | **2,417.4** [R] | Yes — YE25 20-F Note 7 |
| Q1’26 (3/31/26) | **4,692.7** [R] | **2,427.9** [R] | Yes — Q1’26 interim FS Note (PPE) |
| Q2’26 / H1’26 (6/30/26) | **5,197.7** [R] | **7,836.4** [R] | Yes — Q2’26 interim FS Note 7 |

**Quarters that lack the servers/ANIU roll:** **Q1’25, Q2’25, Q3’25** (all mid-2025 interim periods). Also Q1–Q3’24 lack the roll in the KPI grid (only YE24 20-F starts the series).

**What YE25 and H1’26 do disclose (cite):**
- YE25 20-F Note 7: PPE table with Server and network equipment and Assets not yet in use at 12/31/24 and 12/31/25 (`edgar/nbis-20251231x20f.htm.txt`).
- Q1’26 interim: same captions at 12/31/25 and 3/31/26 (`edgar/nbis-20260331xex99d2.htm.txt`).
- Q2’26 / H1’26 interim Note 7: same captions at 12/31/25 and 6/30/26 — servers **5,197.7**, ANIU **7,836.4** [R] (`edgar/nbis-20260812xex99d2.htm.txt`).

**KPI grid:** Rows “Server and network equipment, gross” / “Assets not yet in use” correctly show **n.d.** for mid-2025. Do **not** interpolate. QoQ “servers placed in service” / “change in ANIU” stay n.d. across those gaps.

---

## 3. Other related n.d. (unchanged; pointer)

See KPI `Checks` / `Sources` and fact_base “Not disclosed / gaps”: post-Mar-2025 company GPU count; active MW after YE25; customer concentration for several quarters; post-6/30 ATM; warrant exercise; $775M facility draw.

---

## Action

No `kpi_build.py` rebuild this turn — no new Reported cells added. This note documents why the n.d. cells remain.
