import re,sys
sys.path.insert(0,'_build'); from md2docx import convert
def demote(t):
    out=[]
    for l in t.split('\n'):
        m=re.match(r'^(#{1,5}) (.*)',l)
        out.append(('#'*(len(m.group(1))+1)+' '+m.group(2)) if m else l)
    return '\n'.join(out)
wl=demote(open('q3_watchlist.md').read()); dl=demote(open('dilution_scenarios_onepager.md').read())
head='''# Appendix: Filings & Capital (NBIS short), Q3 watch-list and dilution / headroom

**Final for the Sunday Oct 11, 2026 pack** · built Fri Oct 9, 2026 (ET) · **Owner:** Filings & Capital · **Public information only**
**Tags:** **[R]** Reported · **[D]** Derived (math shown) · **[E]** Estimated / team judgement · **ND** = not disclosed.

This appendix merges the frozen `q3_watchlist` (Part A) and the frozen `dilution_scenarios_onepager` (Part B) without changing any number. Part C lists the primary sources. Where Part B's 6/30 base sections differ from the v0.9e figures at its top, the v0.9e figures govern.

## Locked figures used in both parts (team locks, 10/9/26)

- **PT share count:** 358,558,306 FD for any PT below $313.46 (new Feb-30/Feb-34 notes as debt); 369,564,496 for $313.46–$324.65; 376,648,956 above $324.65 [D]. Every PT the team quotes sits in the first tier.
- **No-vote authority available:** 77.76M ≈ **77.8M** = 64.41M AGM 20% pool + 34.42M treasury − 21.07M NVIDIA pre-funded warrant [D]. The ~97M upper case (exchange shares newly issued) is background only; settlement source ND.
- **Ratios are shares of authority needed ÷ 77.8M** (authority needed = new shares + convert reserve; Modeling v0.9e): Case (a) ~215M = 2.8×; A4 ~230M = 3.0×; D3 ~226M = 2.9×; high-path mixed routes 2.9–3.5×; Bitdeer at $24.6M ~209–213M = ~2.7×; Lambda at $21.7M ~322–326M = 4.1–4.2× [D].
- **Comps band:** $9.7–16.7M per MW **as reported** (Lambda and Bitdeer bases unstated). Management's four Q2'26 deals: $20–25M per MW, avg TCV >$1B each, MW basis and term not disclosed [R, Q2'26 letter p.3–4].
- **Q3 results date:** not company-announced as of Fri 10/9/26 8:40 AM ET (EDGAR + nebius.com newsroom). Date releases have come 12–20 days before results; a late date is treated as neutral, not a signal [D/E].

---

# Part A. Q3 2026 watch-list
'''
mid='''

---

# Part B. Dilution / convert scenarios and no-vote headroom
'''
src='''

---

# Part C. Primary sources (SEC EDGAR, Nebius Group N.V., CIK 0001513845, unless noted)

| Document | Date | Used for | Link |
|---|---|---|---|
| Q2'26 results PR, 6-K Ex.99.1 | 8/12/26 | Revenue, capex, outstanding shares | https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-1.htm |
| Q2'26 shareholder letter, 6-K Ex.99.2 | 8/12/26 | ARR, Q2 deal ACV $20–25M/MW, payback, MSFT tranches | https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm |
| Q2'26 interim financial statements | 8/12/26 | RPO, deferred revenue, ATM (Note 14), warrant, converts, concentration | https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm |
| $775M secured facility, 6-K/A | 7/17/26 | Facility terms; draw ND | https://www.sec.gov/Archives/edgar/data/1513845/000110465926084606/tm2620779d1_6ka.htm |
| AGM notice, 6-K (acc 0001104659-26-087080) | 7/27/26 | Record date 7/28, items 13–14 | EDGAR accession 0001104659-26-087080 |
| Convert pricing, 6-K Ex.99.1 | 8/20/26 | Feb-30 / Feb-34 notes, conversion rates | https://www.sec.gov/Archives/edgar/data/1513845/000110465926098924/tm2623617d1_ex99-1.htm |
| Convert closing + Form of Exchange Agreement (Ex.4.5) | 8/24/26 | Exchange share formula, ~15.8M exchange shares | https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex4-5.htm |
| AGM results, 6-K Ex.99.1 | 8/26/26 | Items 13–14 passed (20% pool) | https://www.sec.gov/Archives/edgar/data/1513845/000110465926101076/tm2624077d1_ex99-1.htm |
| ATM prospectus supplement, 424B5 | 11/12/25 | ATM programme | https://www.sec.gov/Archives/edgar/data/1513845/000110465925110025/tm2530771-1_424b5.htm |
| NVIDIA Schedule 13G | 7/20/26 | 22,256,412 sh incl. 21,065,936 warrant shares | https://www.sec.gov/Archives/edgar/data/1513845/000104581026000062/xslSCHEDULE_13G_X01/primary_doc.xml |
| Goldman Sachs 13G / 13G/A | 8/7/26 | Holder history | https://www.sec.gov/Archives/edgar/data/1513845/000088698226000305/ |
| Goldman Sachs 13G/A No.2 (event 9/30) | 10/7/26 | 13,010,758 sh, 5.1% | https://www.sec.gov/Archives/edgar/data/1513845/000088698226000526/ |
| Latest issuer 6-K (Inferize) | 10/1/26 | No capital change since | https://www.sec.gov/Archives/edgar/data/1513845/000110465926112824/tm2626792d1_6k.htm |
| FY25 Form 20-F | 2026 | MSFT Vineland contract wording, revenue model | Nebius FY25 20-F (`edgar/nbis-20251231x20f.htm`) |
| Q2'26 earnings call transcript (third party) | 8/12/26 | YE26 ARR and capex guidance remarks | https://stockanalysis.com/stocks/nbis/transcripts/655490-q2-2026/ |
| EDGAR submissions feed | checked 10/9/26 | Filing-by-filing re-check | https://data.sec.gov/submissions/CIK0001513845.json |
| Nebius newsroom | checked 10/9/26 | Q3 results-date release watch | https://nebius.com/newsroom |

Supporting workpapers (on the shared drive): `fd_share_bridge.md/.xlsx` (FD and headroom math), `post630_capital_update.md`, `q3_watchlist.xlsx` (change log rows 1–9), `dilution_scenarios.xlsx`, `edgar_recheck_2026-10-08.md`.
'''
open('appendix_filings_capital_final.md','w').write(head+wl.split('\n',1)[1]+mid+dl.split('\n',1)[1]+src)
convert('appendix_filings_capital_final.md','appendix_filings_capital_final.docx',landscape=True)
print('built')
