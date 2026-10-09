# Nebius (NBIS) alt-data from disclosures (Filings & Capital workstream)

Built Oct 4 2026 (ET) for Aadit Prayag's Point72 pitch. Public sources only. Every number cites a URL and date. Blank / "not available" means the company did not disclose it.

## Status
| Task | Sheet / file | Status |
|---|---|---|
| 2 Guidance hit rate + MW wording history | `Guidance_HitRate`, `MW_Wording`, `guidance_hitrate.png` | DONE (v1) |
| 4 Prepayments / RPO / commitments | `Prepay_Commitments`, `prepay_commitments.png` | DONE (v1) |
| 1 Language text-mining | `Language` | Done: `language_trends.png`; mined text in `mined_text/` |
| 3 Raise to deploy | `Raise_to_Deploy` | Done |
| 5 Insiders (Form 4/144) | `Insiders` | Done: `insiders.png` |

## Sources
* SEC EDGAR (CIK 1513845): quarterly results 6-Ks (ex 99.1 press release, ex 99.2 shareholder letter), interim financial statements 6-Ks (Q1-26, Q2-26 ex 99.2), the 20-Fs for FY2024 and FY2025, and financing 6-Ks. Accessed with curl, header `User-Agent: Aadit Prayag research aadit.prayag@nyu.edu`, under 10 req/s. Text copies of the exhibits are reused from the main fact-base folder `/workspace/nebius/filings/edgar/`. Only the FY2024 20-F was newly pulled, into `raw/`.
* Earnings-call transcripts: StockAnalysis.com (transcript data from Quartr), free pages, saved verbatim in `transcripts/transcript_<q>.txt` with the source URL in the header. There are 8 calls, Q3-24 to Q2-26. **No earnings calls were held for Q1-24 or Q2-24.** In that period Yandex N.V. was being divested and NBIS trading was halted until Oct 21 2024. Q3-24 and Q4-24 also had no shareholder letter (press release only).
* All URLs are in `scripts/sources.py`. Every EDGAR URL was checked HTTP 200 on 2026-10-04.

## Method
### Task 2: guidance hit rate
* Every quantified forward-looking statement on ARR, revenue, capex, MW (contracted/connected/active), GPU deployments, EBITDA and prepayments is in `scripts/guidance_data.py`. The sources are the Oct-18-2024 investor presentation, the press releases, the shareholder letters and the call transcripts.
* Each quote is verbatim. `scripts/verify_quotes.py` checks after whitespace normalization that every statement quote and every "actual" quote appears in the local source text. Result on the last run: 40 statements plus 15 wording quotes, 0 failures.
* Status values:
  * BEAT/HIT: met by the target date.
  * MISS: not met.
  * SLIPPED: met late; slip is counted in quarters.
  * PENDING: the target date is in the future.
  * NOT VERIFIABLE: the actual was never disclosed.
* Hit rate = (HIT+BEAT) / resolved statements. A version excluding contracted-MW rows is also given, because contracted MW is about 1% of capex according to the CEO on the Nov-11-2025 call.
* A statement that was superseded is still scored against its own target. The latest guide is shown as a separate row.
### Task 4: prepayments and commitments
* Quarter-end values come from the press-release balance sheets (Q2-24 to Q3-25), the 20-F notes (YE24, YE25) and the interim statements notes (Q1-26, Q2-26). Where a 20-F restated a value, the restated value is used, and the note says so.
* RPO is disclosed only from YE25. The amount due within 24 months = RPO x the disclosed %.
* GPU/hardware purchase obligations are **not disclosed** as a quantified commitment.


### Task 1: language text-mining
- Corpus: 8 earnings-call transcripts, Q3-24 to Q2-26, from free StockAnalysis pages (URLs in `scripts/sources.py`, keys `CALL_*`; full text in `transcripts/`). Shareholder letters (6-K ex 99.2) cover Q1-25 to Q2-26. For Q2-24 to Q4-24 no letter was published, so the earnings press release (ex 99.1) is used and flagged `release*`. Q1-24 has no call and no letter (Yandex divestment period).
- Cleaning: alphabetic tokens only, with the forward-looking-statement boilerplate removed. Financing and share-price usages are masked (EXCL list). Supply phrases are matched first and then masked, so "demand exceeds supply" counts only as supply-constrained.
- Phrase families: the full regex dictionary is in `scripts/language.py` (`FAM`, `POS_PRICE`, `NEG_PRICE`, `EXCL`) and is reproduced at the bottom of the `Language` sheet. Counts are normalised per 10,000 words.
- Outputs: `scripts/language_results.json`, `mined_text/<Q>_<kind>.txt` (each with a source URL header) and `language_trends.png`.

### Task 3: raise to deploy
- Each funding source F0 to F12 (starting cash, private placement, converts, equity, NVIDIA, ATM, secured facility, customer prepayments) is logged with its date, amount, source key and quote (`scripts/raise_sheet.py`).
- Deployment is matched FIFO against quarterly cash capex (CF "purchase of P&E and intangibles"). The lag is the number of quarters from the raise quarter to the quarter in which that raise is fully consumed.
- These are overlaid with PP&E net growth and announced MW/ARR milestones.
- A regression of capex against the change in ARR at lags k=0..3 is trend-dominated (r=0.86 to 0.94 at every k), so it cannot identify the lag. Management's stated 6-12 weeks from connected power to revenue is cited instead.

### Task 5: insiders
- Filing list: `raw/insider_filings_list.json`, from submissions JSON `recent` (filings since 2024-01-01: Form 4 x37, 4/A x2, 144 x31, 3 x15). The raw XML for each filing is in `raw/insiders/`.
- Parsing: `scripts/parse_insiders.py` writes `scripts/insiders.json`.
  - Form 4: code S = sale. Each sale is classified as RSU tax sell-to-cover (from the footnote), 10b5-1 (from the `aff10b5One` box or the footnote), or other.
  - A 4/A replaces the original for the same owner and date.
  - Form 144: shares, aggregate market value, approximate sale date, plan adoption date (flagged as 10b5-1) and remarks.
- Coverage: no Form 4 exists before Mar-31-2026. NBIS is a foreign private issuer, and FPI insiders were exempt from Section 16 until the Holding Foreign Insiders Accountable Act took effect on Mar-18-2026 (SEC FAQ: https://www.sec.gov/about/divisions-offices/division-corporation-finance/holding-foreign-insiders-accountable-act-frequently-asked-questions). Form 3s were filed that day. 10% holders of FPIs remain exempt. Form 144 notices (proposed sales) exist from Mar-2025. They are not proof of execution.
- Outputs: the `Insiders` sheet (by insider and quarter, a Form 144 table, an events overlay and transaction detail) and `insiders.png` (`scripts/chart_insiders.py`).

## Scripts (`scripts/`)
`sources.py` (URL registry) · `guidance_data.py` · `prepay_data.py` · `verify_quotes.py` · `build_xlsx.py` (writes `altdata.xlsx`, auto-adds the other sheets once their modules exist) · `charts_t2_t4.py` · `kwic.py` (keyword-in-context helper, at top level).
Rebuild: `cd /workspace/nebius/filings/altdata && python3 scripts/verify_quotes.py && python3 scripts/build_xlsx.py && python3 scripts/charts_t2_t4.py`

Full rebuild: `python3 scripts/verify_quotes.py && python3 scripts/language.py && python3 scripts/parse_insiders.py && python3 scripts/build_xlsx.py && python3 scripts/charts_t2_t4.py && python3 scripts/chart_language.py && python3 scripts/chart_insiders.py`

## Caveats
* Capex = cash "purchases of property and equipment (and intangibles)" from the cash-flow statement. Accrual capex (including vendor-financed equipment) may be higher.
* FY2024 and Oct-2024 guidance included Toloka. Toloka was deconsolidated in May 2025, so the FY25 revenue comparisons are not like-for-like.
* MW-1 (Finland 75 MW) relies on the Capacity teammate's verification (`/workspace/nebius/capacity/site_list.md`).
