import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
U=dict(
FB="/workspace/nebius/filings/nebius_fact_base.md (team Filings fact base; nebius_filings.xlsx Converts/FactBase sheets)",
Q2FS="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm",
COMPS="/workspace/nebius/parts/comps.xlsx (team GPU Market comps; EV_Bridge sheet)",
PR="https://assets.nebius.com/assets/72a8c258-bbb7-4df7-ab9d-8698f6cb88fc/PR.pdf",
SHL="https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf",
SAH="https://stockanalysis.com/stocks/nbis/history/",
SA="https://stockanalysis.com/stocks/nbis/statistics/",
SAF="https://stockanalysis.com/stocks/nbis/forecast/",
SABS="https://stockanalysis.com/stocks/nbis/financials/balance-sheet/?p=quarterly",
MBF="https://www.marketbeat.com/stocks/NASDAQ/NBIS/forecast/",
MBS="https://www.marketbeat.com/stocks/NASDAQ/NBIS/short-interest/",
MI="https://markets.businessinsider.com/stocks/nbis-stock",
YF="https://finance.yahoo.com/quote/NBIS/analysis/",
MS="https://www.marketscreener.com/quote/stock/NEBIUS-GROUP-8037501/finances/",
CLOSE="https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex99-1.htm",
PRICE="https://www.sec.gov/Archives/edgar/data/1513845/000110465926098924/tm2623617d1_ex99-1.htm",
NV="https://www.sec.gov/Archives/edgar/data/1513845/000110465926026163/tm268532d1_6k.htm",
MAR="https://nebius.com/newsroom/nebius-group-announces-closing-of-private-offering-of-convertible-senior-notes-with-aggregate-gross-proceeds-of-approximately-4-3-billion",
MAR33="https://www.sec.gov/Archives/edgar/data/1513845/000110465926032735/tm268409d4_ex4-3.htm",
SEP25="https://nebius.com/newsroom/nebius-group-announces-pricing-of-upsized-private-offering-of-2-75-billion-of-convertible-senior-notes",
JUN25="https://www.sec.gov/Archives/edgar/data/1513845/000110465925055197/tm2516825d1_6k.htm",
LOAN="https://assets.nebius.com/assets/cf774b13-960d-4c5e-a78f-a0278799cca1/17072026%C2%A0Nebius%20raises%20%24775%20million%20in%20first%20secured%20debt%20financing%20to%20accelerate%20global%20buildout.pdf",
CHM="https://www.chartmill.com/stock/quote/NBIS/ownership/short-interest",
CT="https://curvedtrading.com/stocks/nbis/",
EQ="https://equibles.com/stocks/nbis/short-interest",
CRWV="https://stockanalysis.com/stocks/crwv/statistics/",
CRWVY="https://finance.yahoo.com/quote/CRWV/analysis/",
CRWVMS="https://ca.marketscreener.com/quote/stock/COREWEAVE-INC-185380761/calendar/",
ROTH="https://www.tipranks.com/news/the-fly/nebius-initiated-with-a-sell-at-rothschild-co-redburn-thefly-news",
NASDAQ="https://www.nasdaq.com/press-release/nasdaq-resumes-trading-nebius-group-nv-2024-10-18",
)
R=[]  # section, metric, value, unit, tag, asof, source, notes
def a(*x): R.append(x)
S="1 Price & shares"
a(S,"Share price (close)",242.81,"USD","Reported","2026-10-02",U["SAH"],"Corroborated by MarketScreener, MarketBeat, Yahoo")
a(S,"Shares issued & outstanding (Class A+B, ex-treasury)",271.855218,"M sh","Reported","2026-06-30",U["PR"],"238,400,165 Class A + 33,455,053 Class B; excludes 50,185,726 treasury")
a(S,"Shares outstanding per StockAnalysis",274.10,"M sh","Reported (aggregator)","2026-10-04",U["SA"],"Basis not disclosed; differs from company figure")
a(S,"Shares issued in Aug-24 note exchange",15.77,"M sh","Reported","2026-08-24",U["CLOSE"]+" ; "+U["FB"]+" (CapStructure)","~15.77M Class A shares for $800M of 2029/2031 notes (press release: 'approximately 15.8 million')")
a(S,"NVIDIA pre-funded warrant shares",21.065936,"M sh","Reported","2026-03-11",U["NV"],"v3 2026-10-04 (Filings reversal, high confidence): NOT included in the 271.86M ISSUED shares - added to economic basic. (v2 had wrongly removed it.) Exercise not reported through 10/1/26")
a(S,"ATM shares sold after 2026-06-30","none disclosed","M sh","Reported (absence)","2026-10-01",U["Q2FS"]+" ; "+U["FB"],"12,729,493 sold in Q2 at WA $223.60; 12,270,507 remain under 25M program; no post-6/30 sales disclosed through 10/1/26 6-K")
a(S,"Share options outstanding",6.7406,"M","Reported","2026-06-30",U["Q2FS"]+" ; "+U["FB"],"WAEP $88.65 (Note 13)")
a(S,"Unvested RSUs",6.234091,"M","Reported","2026-06-30",U["Q2FS"]+" ; "+U["FB"],"WA grant FV $58.22 (Note 13)")
a(S,"Economic basic shares (v3)",308.69,"M sh","Derived","2026-10-02",U["FB"]+" (CapStructure)","271.86 issued + 15.77 Aug exchange + 21.07 NVIDIA pre-funded warrant = 308.69 (v2 287.66 superseded). No post-6/30 ATM disclosed")
a(S,"Convertible notes outstanding (orig. principal, pro forma)",13425,"USD M","Derived","2026-10-02",U["FB"],"CORRECTED: $13.43B = 87.5+87.5 (2029/2031) + 1,581.25x2 (2030 1%/2032 2.75%) + 2,587.5 (2031) + 1,750 (2033) + 3,450 (2/2030) + 2,300 (2/2034). NO capped calls disclosed")
a(S,"In-the-money convert shares (if-converted @ $242.81)",50.02,"M sh","Derived","2026-10-02",U["FB"],"CORRECTED: 2029/2031 ~$87.5M each @$51.45 (1.70M each); 2030 1%/2032 2.75% $1,581.25M each @$138.75 (11.40M each); 2031 1.25% $2,587.5M @$183.22 (14.12M); 2033 2.625% $1,750M @$180.31 (9.71M). Initial conv. rates; accretion features ignored")
a(S,"Out-of-the-money convert shares",18.09,"M sh","Derived","2026-10-02",U["PRICE"],"2030 0.50% $3,450M @$313.46 (11.01M); 2034 4.50% $2,300M @$324.65 (7.08M)")
a(S,"Options, treasury-stock method @ $242.81",4.28,"M sh","Derived","2026-10-02",U["FB"],"6.7406 x (1 - 88.65/242.81)")
a(S,"Fully diluted shares @ $242.81 (Filings CapStructure)",381.5,"M sh","Derived","2026-10-02",U["FB"]+" (CapStructure)","308.69 + RSUs 6.23 + options 4.28 (TSM) + ITM converts 50.02 + ATM capacity 12.27 = 381.5")
a(S,"Shares incl. ALL converts regardless of price",399.6,"M sh","Derived","2026-10-02",U["FB"]+" (CapStructure)","381.5 + OTM converts 18.09 = 399.6")
a(S,"Shares under net-share settlement (all converts)",349.9,"M sh","Derived","2026-10-02",U["FB"]+" (CapStructure)","Would need ~$7.7B cash to settle principal. Default settlement is SHARES; NO capped calls")
a(S,"Issuer call provisions","callable above 130% of conversion price; first = Dec 2026 (leftover 2029s)","","Reported","2026-08-24",U["FB"]+" (Indenture_Terms)","Dilution steps down at $183 (2031/2033), $139 (2030/2032), $51 (2029/2031)")
a(S,"Market cap - headline (StockAnalysis)",66.55,"USD B","Reported (aggregator)","2026-10-02",U["SA"],"274.10M x $242.81")
a(S,"Market cap - company share count",66.01,"USD B","Derived","2026-10-02",U["PR"],"271.855M x $242.81")
a(S,"Market cap - Markets Insider",64.13,"USD B","Reported (aggregator)","2026-10-02",U["MI"],"Uses 238.40M (Class A only) - understated; flagged")
a(S,"Market cap - economic basic (base case)",74.95,"USD B","Derived","2026-10-02","see rows above","v3: 308.69M x $242.81 (v2 69.85 superseded)")
a(S,"Market cap - fully diluted",92.63,"USD B","Derived","2026-10-02","see rows above","v3: 381.5M x $242.81")
S="2 Enterprise value"
a(S,"Cash & cash equivalents",8042.1,"USD M","Reported","2026-06-30",U["PR"],"Latest filed balance sheet (Q2'26 6-K/PR)")
a(S,"Restricted cash (excluded from EV cash)",1056.0,"USD M","Reported","2026-06-30",U["PR"],"Excluded (conservative)")
a(S,"Debt, current",46.7,"USD M","Reported","2026-06-30",U["PR"],"")
a(S,"Debt, non-current",8499.0,"USD M","Reported","2026-06-30",U["PR"],"Mostly convertible notes (carrying value)")
a(S,"Operating lease liabilities (non-current / current)", "1,510.4 / 114.9","USD M","Reported","2026-06-30",U["PR"]+" ; "+U["SABS"],"Current portion from StockAnalysis (PR does not break out)")
a(S,"Deferred revenue (customer prepayments) current+non-current",5975.2,"USD M","Reported","2026-06-30",U["PR"],"979.4 + 4,995.8; debt-like service obligation (bear adjustment)")
a(S,"+ Secured GPU-backed facility (Jul 17 2026)",775,"USD M","Reported","2026-07-17",U["LOAN"],"SOFR+2.50%, matures 2030-10-31. NOT in a filed balance sheet; assumed fully drawn")
a(S,"+ Aug-24 converts (2030 0.50% $3.45B + 2034 4.50% $2.30B)",5750,"USD M","Reported","2026-08-24",U["CLOSE"],"Original principal; accretes to 110%/125% at maturity. NOT in a filed balance sheet")
a(S,"- 2029/2031 notes exchanged for shares",-800,"USD M","Reported","2026-08-24",U["CLOSE"],"$400M each; face value used (carrying not disclosed)")
a(S,"Pro forma total debt (ex-leases)",14270.7,"USD M","Derived","2026-10-02","rows above","46.7 + 8,499.0 + 775 + 5,750 - 800")
a(S,"Pro forma cash",14497.1,"USD M","Estimated","2026-10-02",U["PRICE"],"8,042.1 + 775 + ~5,680 est. net note proceeds. FLAG: EXCLUDES Q3 burn - Q2 capex alone was $5,657M (offset partly by prepayment-driven OCF, Q2 +$2,246M). True 9/30 cash likely materially lower; Q3 6-K (~Nov 10) needed")
a(S,"EV - headline (StockAnalysis)",68.71,"USD B","Reported (aggregator)","2026-10-04",U["SA"],"$66.55B + $10.20B debt incl. leases - $8.04B cash; Jun-30 balance sheet, pre-Aug converts/exchange")
a(S,"EV - base case (pro forma, ex-leases)",74.73,"USD B","Derived","2026-10-02","rows above","v3: 74.95 + 14.27 - 14.50 = 74.73 (v2 69.62 superseded). Ex-lease basis = same as CRWV comp")
a(S,"EV - base incl. started operating leases",76.35,"USD B","Derived","2026-10-02","rows above","74.73 + 1.625")
a(S,"EV - incl. off-BS commitments (energy $5.32B + leases signed not commenced $12.05B) and started leases",93.73,"USD B","Derived","2026-10-02",U["Q2FS"],"74.73 + 5.324 + 12.054 + 1.625 (undiscounted; bound only)")
a(S,"EV - fully diluted (ITM converts as equity, ATM proceeds as cash)",81.75,"USD B","Derived","2026-10-02","rows above","v3: 92.63 + (14.27 - 7.675 ITM face) - (14.50 + 2.98 ATM proceeds 12.27M x $242.81) = 81.75")
a(S,"EV - base + deferred revenue as debt-like (bear view)",80.70,"USD B","Derived","2026-10-02","rows above","74.73 + 5.975")
a(S,"EV per contracted MW",21.4,"USD M/MW","Derived","2026-10-02",U["FB"],"v3: 74,727 / >3,500 MW (May-26 contracted); 14.9 on 5 GW target; incl. signed-not-started leases 12.05B + started 1.63B: 25.3 / 17.7. CRWV ~21.5 (79,520/3,700, ex-lease) per comps.xlsx")
a(S,"EV per energized MW (started leases only)",413,"USD M/MW","Derived","2026-10-02","/workspace/nebius/capacity/nebius_capacity.xlsx (Quarterly MW, Q2 2026 base)","(74,727 + 1,625 started leases) / ~185 gross MW energized at 6/30 (173 MW mixed basis, IT sites x1.25)")
a(S,"SOTP non-core stakes (Filings)","~$1.6B carrying / ~$4.3B upper bound","USD B","Derived (Filings)","2026-06-30",U["FB"]+" (SOTP)","~$4-11 per fully diluted share (381.5M)")
S="3 Consensus & guidance"
for m,v,n,src,nt in [("Revenue FY2026 cons (mean)",3.34,"20 analysts; range 3.0-3.59",U["YF"],""),
 ("Revenue FY2027 cons (mean)",12.30,"21 analysts; range 8.81-17.99",U["YF"],"MarketScreener 12.302 agrees"),
 ("Revenue FY2028 cons",22.983,"analyst count not shown",U["MS"],""),
 ("Adj. EBITDA FY2026 cons",1.387,"MarketScreener (~10 analysts on page)",U["MS"],"EBITDA definition (adjusted vs reported) not specified"),
 ("Adj. EBITDA FY2027 cons",6.728,"MarketScreener",U["MS"],"implies 54.7% margin. Other aggregators show ~6.36B (finanzen.net) - range 6.36-6.73"),
 ("Adj. EBITDA FY2028 cons",13.978,"MarketScreener",U["MS"],"60.8% margin"),
 ("EPS FY2026 cons",-2.28,"8 analysts; range -6.41 to 0.49",U["YF"],"USD/share"),
 ("EPS FY2027 cons",-5.64,"9 analysts; range -15.31 to 5.27",U["YF"],"USD/share"),
 ("EPS FY2028 cons",3.434,"",U["MS"],"USD/share"),
 ("Capex FY2026 cons",23.207,"",U["MS"],""),("Capex FY2027 cons",29.135,"",U["MS"],""),("Capex FY2028 cons",38.885,"",U["MS"],""),
 ("FCF FY2027 cons",-23.148,"",U["MS"],""),("Net debt YE2026 cons",13.659,"",U["MS"],""),("Net debt YE2027 cons",34.316,"",U["MS"],""),("Net debt YE2028 cons",50.728,"",U["MS"],"")]:
    a(S,m,v,"USD B" if "EPS" not in m else "USD","Estimated (consensus)","2026-10-04",src,(n+"; "+nt).strip("; "))
a(S,"FY2029-2030 consensus","not disclosed/not found","","n/a","2026-10-04",U["MS"],"No public source found beyond FY2028")
for m,v in [("Guide: FY26 group revenue","$3.0-3.4B"),("Guide: YE26 ARR","$7-9B"),("Guide: FY26 adj. EBITDA margin","~40%"),("Guide: FY26 capex","$20-25B"),("Guide: YE26 connected power","800MW-1GW"),("Guide: YE26 contracted power","5GW (raised from >4GW)"),("Guide: 2026 customer prepayments",">$9B"),("Guide: deployment pace",">1GW/yr starting 2027")]:
    a(S,m,v,"","Reported (guidance)","2026-08-12",U["SHL"],"Reaffirmed/raised Aug 12 2026")
a(S,"Q2'26 revenue / adj. EBITDA / ARR","582.3 / 236.2 / 3,000","USD M","Reported","2026-06-30",U["PR"],"ARR = last-month AI cloud revenue x12")
a(S,"Q2'26 capex (PP&E+intangibles)",5657.4,"USD M","Reported","2026-06-30",U["PR"],"H1'26 8,130.3")
a(S,"ACV per MW: 2026 base / Q2 deals / Q3 short-term","$12M / >$20M ($20-25M) / >$40M","USD M per MW","Reported (approx.)","2026-08-12",U["SHL"],"")
a(S,"Customer commitments (IG customers e.g. MSFT, META)",">$40B","USD","Reported","2026-08-12",U["SHL"],"")
a(S,"Server/network useful life","5 yrs (from 4 yrs pre-2026)","","Reported","2026-08-12",U["SHL"],"Lowers D&A vs prior policy")
a(S,"Q2'26 D&A / SBC / interest expense / OCF","259.7 / 102.5 / 119.1 / 2,246.1","USD M","Reported","2026-06-30",U["PR"]+" ; "+U["FB"],"OCF prepayment-driven (deferred revenue +1,197.0; A/R -1,186.8). Useful-life 4->5yr cut Q2 D&A by $43.0M; FY26 est. -$167.6M (company)")
a(S,"PP&E net / RPO","13,045.2 / 37,490.6","USD M","Reported","2026-06-30",U["Q2FS"],"RPO 36% within 24 months")
a(S,"Top-3 customers % of Q2 revenue","24% / 21% / 14%","%","Reported","2026-06-30",U["Q2FS"],"Unnamed")
a(S,"Microsoft contract","TCV $17,392.9M; 9 tranches; ~$6,958M upfront; billed through Oct-2031","USD M","Reported","2026-04-30",U["FB"],"Q2 letter says all tranches delivered; CONFLICT: top customer only ~$140M/qtr in Q2 vs ~$725-870M/qtr fully-delivered run-rate (Capacity) - open")
a(S,"Meta contracts","#1 ~$2,881M/5yr; #2 $12B firm (Vera Rubin, tranches from early 2027) + up to $15B backstop (likely outside RPO)","USD M","Reported","2026-03-16",U["FB"],"")
a(S,"Prepayment coverage","~70% of deals carry prepayments covering 50-60% of related capex","","Reported (company)","2026-08-12",U["SHL"]+" ; "+U["FB"],"Older-gen GPU pricing >30% higher Q2 vs Q1; new Q2 deals ($20-25M/MW) priced on forecast costs incl. capacity not yet built")
S="4 Price targets"
a(S,"PT avg / median / low / high (S&P Global via StockAnalysis)","283.58 / 300 / 84 / 415","USD","Reported (aggregator)","2026-09-25",U["SAF"],"19 analysts")
a(S,"PT avg / low / high (MarketBeat)","241.80 / 84 / 410","USD","Reported (aggregator)","2026-10-04",U["MBF"],"20 analysts; includes stale targets (e.g. GS $120 from 2025-09)")
a(S,"PT median / low / high (Markets Insider)","155.98 / 30 / 410","USD","Reported (aggregator)","2026-10-02",U["MI"],"43 analysts - appears stale; flagged, not used")
a(S,"Ratings (MarketBeat)","1 Sell / 6 Hold / 11 Buy / 2 Strong Buy","","Reported (aggregator)","2026-10-04",U["MBF"],"Consensus Moderate Buy")
a(S,"Ratings Sep-26 (S&P Global)","4 Strong Buy / 1 Buy / 3 Hold / 0 Sell / 1 Strong Sell","","Reported (aggregator)","2026-09-25",U["SAF"],"9 ratings in month")
for d,f,t in [("2026-09-30","William Blair","Initiate Outperform, no PT"),("2026-09-24","BNP Paribas Exane","Upgrade to Outperform, PT $260->$399"),("2026-09-21","Rothschild & Co Redburn","Initiate Sell, PT $84"),("2026-09-18","Northland","Reiterate Outperform $410 (raised from $248 on 2026-07-20)"),("2026-09-09","Truist","Initiate Buy, PT $355"),("2026-08-14","Citi","Buy, PT $278->$324"),("2026-08-13","Baird","Outperform, PT $250->$340"),("2026-08-13","DA Davidson","Neutral, PT $175"),("2026-08-03","Piper Sandler","Initiate Neutral, PT $224"),("2026-06-08","BofA","Buy, PT $240->$280"),("2026-05-14","Morgan Stanley","Equal Weight, PT $126->$144 (stale)")]:
    a(S,"PT change: "+f,t,"","Reported","{}".format(d),U["MBF"] if f!="Truist" else U["SAF"],"")
S="5 Short interest"
a(S,"Shares short",46.77954,"M sh","Reported (FINRA)","2026-09-15",U["MBS"],"Prev 45.27M (Aug 31); peak 61.07M (Jul 15)")
a(S,"Short % of float","18.5% / 20.5% / 22.05% / 22.8%","%","Reported (aggregators)","2026-09-15",U["MBS"]+" ; "+U["SA"]+" ; "+U["CHM"],"Differences = float denominator")
a(S,"Days to cover","3.6 / 2.76 / 2.23","days","Reported (aggregators)","2026-09-15",U["EQ"]+" ; "+U["SA"]+" ; "+U["CHM"],"FINRA ADV basis = 3.6")
a(S,"Borrow fee","~0.3%","%","Reported","2026-10-02",U["CT"],"Easy to borrow")
a(S,"Next FINRA release","2026-10-09 (Sep 30 settlement)","","Reported","2026-10-04",U["EQ"],"")
S="6 Implied multiples"
for nm,e in [("headline",68.71),("base",74.73),("fully diluted",81.75)]:
    a(S,f"EV/FY26 rev | EV/FY27 rev | EV/FY27 EBITDA ({nm} EV ${e}B)",f"{e/3.34:.1f}x | {e/12.30:.2f}x | {e/6.728:.1f}x","x","Derived","2026-10-02","EV rows + consensus rows",f"{e}/3.34; {e}/12.30; {e}/6.728")
a(S,"CRWV EV - aggregator incl. ~$16.5B leases (NOT used)","95.46","USD B","Reported (aggregator)","2026-10-02",U["CRWV"],"CORRECTED: inconsistent with NBIS ex-lease basis")
a(S,"CRWV EV - pro forma ex-leases (used)",79.52,"USD B","Derived","2026-10-02",U["COMPS"],"Mkt cap 49.43 + debt 39.27 (incl. Sep-26 $4.2B converts) - cash 9.18 (comps.xlsx EV_Bridge)")
a(S,"CRWV cons rev FY26 / FY27; EBITDA FY26 / FY27","12.87 / 26.26; 7.569 / 16.64","USD B","Estimated (consensus)","2026-10-04",U["CRWVY"]+" ; "+U["CRWVMS"],"40 analysts rev")
a(S,"CRWV EV/FY26 rev | EV/FY27 rev | EV/FY27 EBITDA (ex-lease)","6.2x | 3.03x | 4.78x","x","Derived","2026-10-02","CRWV rows","CORRECTED: 79.52/12.87; /26.26; /16.64 (was 7.4x/3.64x/5.74x)")
a(S,"NBIS premium to CRWV (ex-lease, base EV)","EV/FY27 rev +101% | EV/FY27 EBITDA +132%","","Derived","2026-10-02","rows above","v3: 6.08/3.03 = 2.01x ; 11.11/4.78 = 2.32x")
a(S,"Revenue per active MW (annualized Q2)","NBIS n/m (active MW stale) | CRWV 6.9-8.2 | IREN AI ~7.1","USD M/MW","Derived","2026-06-30",U["COMPS"],"Comps Rev-per-MW block; NBIS active MW last disclosed ~170 MW at YE25")
a(S,"Wtd effective cost of debt","NBIS 5.1% | CRWV 8.4%","%","Derived","2026-06-30",U["COMPS"],"Principal-weighted effective rates")
a(S,"Oracle / IREN comps","not found (not attempted in depth)","","n/a","","","Gap")
a(S,"Illustrative: NBIS at CRWV FY27 EV/EBITDA",32.15,"USD B EV","Derived","2026-10-02","rows above","v3: 6.728 x 4.78 = 32.15B EV + 0.23B net cash = 32.38B equity => ~$105/sh on 308.69M (-57%). At CRWV EV/FY27 rev 3.03x: 37.27B EV => ~$121/sh (-50%)")
a(S,"Illustrative: 8x FY28 EBITDA less YE28 cons net debt",61.1,"USD B equity","Derived","2026-10-02",U["MS"],"8 x 13.978 - 50.728 = 61.1B vs 74.95B economic mkt cap (~$198/sh on 308.69M; lower after future issuance)")
# history
H=[("Q1 2024","2024-03-31",None,None,None,None,"Trading halted 2022-02-28 to 2024-10-21 (last trade $18.94 as YNDX)"),
("Q2 2024","2024-06-30",None,None,None,None,"Halted"),("Q3 2024","2024-09-30",None,None,None,None,"Halted"),
("Q4 2024","2024-12-31",27.70,235.75,2385,37.9,""),("Q1 2025","2025-03-31",21.11,238.11,1259,55.3,""),
("Q2 2025","2025-06-30",55.33,238.71,455,105.1,""),("Q3 2025","2025-09-30",112.27,251.81,225.8,146.1,"Post Sep-2025 equity+convert raise"),
("Q4 2025","2025-12-31",83.71,255.16,-1220,227.7,""),("Q1 2026","2026-03-31",103.76,256.1,-217.2,399.0,"Aggregator count; whether it includes 21.07M NVIDIA warrant shares (issued Mar-11) unverified"),
("Q2 2026","2026-06-30",276.17,271.855,-2153,582.3,"Company ISSUED share count; EXCLUDES 21.07M NVIDIA warrant (Filings v3)"),
("Q3 2026 (partial)","2026-09-30",235.88,None,None,None,"Share count/net cash not yet reported (Q3 results Nov 10)")]
wb=Workbook(); ws=wb.active; ws.title="Priced_in"
hdr=["Section","Metric","Value","Unit","Tag","As-of","Source URL","Notes / math"]
ws.append(hdr)
for r in R: ws.append(list(r))
h=wb.create_sheet("Price_history")
h.append(["Quarter","Quarter-end date","Close (USD)","Shares out (M)","Mkt cap (USD M)","Net cash(+)/debt(-) incl. leases (USD M)","EV (USD M)","Quarter revenue (USD M)","EV / annualized qtr revenue (x)","Tag","Sources","Notes"])
for q,d,p,s,nc,rev,nt in H:
    if p and s:
        mc=p*s; ev=mc-nc; mult=ev/(rev*4)
        h.append([q,d,p,s,round(mc),nc,round(ev),rev,round(mult,1),"Derived (price x shares; EV = mcap - net cash)",U["MBS"]+" (quarter-end close) ; "+U["SABS"]+" (shares, net cash) ; "+U["MS"]+" (qtr revenue)",nt])
    elif p:
        h.append([q,d,p,"not disclosed/not found","","","","","","Reported price",U["SAH"],nt])
    else:
        h.append([q,d,"n/a (halted)","","","","","","","n/a",U["NASDAQ"],nt])
qc=wb.create_sheet("Quarterly_consensus")
qc.append(["Quarter","Revenue cons (USD M)","Adj. EBITDA cons (USD M)","EBITDA margin","Tag","Source","Notes"])
for q,rv,eb,n in [("Q3 2026",904.1,354.7,"Yahoo rev 904.09M (16 analysts; range 876.6-942.6)"),("Q4 2026",1452,670.7,"Yahoo 1.45B (16; range 1.09-1.71B); x4 = 5.8B vs 7-9B ARR guide"),("Q1 2027",2175,1277,""),("Q2 2027",2646,1581,""),("Q3 2027",3304,1883,""),("Q4 2027",4052,2372,"")]:
    qc.append([q,rv,eb,f"{eb/rv:.1%}","Estimated (consensus)",U["MS"],n])
for w in wb.worksheets:
    for c in w[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F3864"); c.alignment=Alignment(wrap_text=True)
    for col in w.columns:
        w.column_dimensions[col[0].column_letter].width=min(60,max(10,max(len(str(c.value or "")) for c in col)*0.9))
    w.freeze_panes="A2"
wb.save("/workspace/nebius/skeptic/01_priced_in_v3.xlsx")
with open("/workspace/nebius/skeptic/01_priced_in_v3.csv","w",newline="") as f:
    cw=csv.writer(f); cw.writerow(hdr); cw.writerows(R)
for row in h.iter_rows(values_only=True): print(row[:9])
print(len(R),"rows")
