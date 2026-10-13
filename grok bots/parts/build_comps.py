#!/usr/bin/env python3
"""Build /workspace/nebius/parts/comps.xlsx (+CSVs). All $ in USD millions unless noted.
Every hard-coded input is listed in SRC with a source URL. Derived cells are Excel formulas."""
import csv, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

PRICE_DATE = "2026-10-02 (close)"
Y = "https://query1.finance.yahoo.com/v8/finance/chart/{t}?range=10d&interval=1d"
U = dict(
 NB_PR="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-1.htm",
 NB_LET2="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm",
 NB_FS="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm",
 NB_LET1="https://www.sec.gov/Archives/edgar/data/1513845/000110465926059872/tm2614392d1_ex99-2.htm",
 NB_LET4="https://www.sec.gov/Archives/edgar/data/1513845/000110465926013946/tm266173d1_ex99-2.htm",
 NB_20F="https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm",
 NB_CONV="https://www.sec.gov/Archives/edgar/data/1513845/000110465926100347/tm2623863d1_ex99-1.htm",
 NB_CONVP="https://www.sec.gov/Archives/edgar/data/1513845/000110465926098924/tm2623617d1_ex99-1.htm",
 NB_CALL="https://www.fool.com/earnings/call-transcripts/2026/08/19/nebius-nbis-q2-2026-earnings-call-transcript/",
 NB_CONS="https://www.marketscreener.com/quote/stock/NEBIUS-GROUP-8037501/finances/",
 SA_NB="https://stockanalysis.com/stocks/nbis/statistics/",
 CR_PR="https://www.sec.gov/Archives/edgar/data/1769628/000176962826000362/coreweave2q26earningspress.htm",
 CR_10Q="https://www.sec.gov/Archives/edgar/data/1769628/000176962826000366/crwv-20260630.htm",
 CR_Q4="https://www.sec.gov/Archives/edgar/data/1769628/000176962826000094/coreweave4q25earningspress.htm",
 CR_XBRL="https://data.sec.gov/api/xbrl/companyfacts/CIK0001769628.json",
 CR_CONV="https://www.sec.gov/Archives/edgar/data/1769628/000176962826000432/ex991pricing.htm",
 CR_CONV8K="https://www.sec.gov/Archives/edgar/data/1769628/000176962826000432/crwv-20260917.htm",
 CR_DDTL="https://www.sec.gov/Archives/edgar/data/1769628/000176962826000357/ex991pr.htm",
 CR_CALL="https://s205.q4cdn.com/133937190/files/doc_financials/2026/q2/CRWV-US-CORRECTED-TRANSCRIPT-CoreWeave-Q2-2026-Earnings-Call-11August2026.pdf",
 CR_CONS="https://finance.yahoo.com/quote/CRWV/analysis/",
 SA_CR="https://stockanalysis.com/stocks/crwv/statistics/",
 IR_PR="https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm",
 IR_10K="https://www.sec.gov/Archives/edgar/data/1878848/000187884826000052/iren-20260630.htm",
 IR_XBRL="https://data.sec.gov/api/xbrl/companyfacts/CIK0001878848.json",
 IR_0720="https://www.sec.gov/Archives/edgar/data/1878848/000114036126028871/ef20078253_ex99-1.htm",
 OR_PR="https://www.sec.gov/Archives/edgar/data/1341439/000119312526387905/orcl-ex99_1.htm",
 OR_10Q="https://www.sec.gov/Archives/edgar/data/1341439/000119312526389274/orcl-20260831.htm",
 OR_FEB="https://www.sec.gov/Archives/edgar/data/1341439/000119312526037840/d23732d8k.htm",
 OR_XBRL="https://data.sec.gov/api/xbrl/companyfacts/CIK0001341439.json",
 NS_S1="https://www.sec.gov/Archives/edgar/data/2110365/000119312526395475/ck0002110365-20260918.htm",
 NS_C="https://www.nscale.tech/press-releases/nscale-series-c",
 NS_CONV="https://www.prnewswire.com/news-releases/nscale-raises-3-36b-in-pre-ipo-convertible-financing-302890199.html",
 CU_F="https://www.globenewswire.com/news-release/2026/09/17/3364326/0/en/crusoe-raises-3-9-billion-series-f-for-its-vertically-integrated-ai-infrastructure-platform.html",
 CU_REV="https://www.newcomer.co/p/crusoe-is-on-track-for-2-billion",
 CU_BBG="https://www.bloomberg.com/news/articles/2026-09-03/crusoe-raises-over-3-billion-in-funding-at-30-billion-valuation",
 LA_BBG="https://www.bloomberg.com/news/articles/2026-08-24/ai-cloud-provider-lambda-in-talks-for-3-billion-pre-ipo-round",
 LA_TFN="https://techfundingnews.com/lambda-eyes-12b-valuation-with-3b-pre-ipo-raise-as-ai-cloud-race-heats-up/",
 LA_SACRA="https://sacra.com/research/no-side-quests-for-lambda/",
 LA_DCD="https://www.datacenterdynamics.com/en/news/microsoft-signs-multi-billion-dollar-cloud-capacity-deal-with-lambda/",
 LA_DEBT="https://lambda.ai/blog/lambda-closes-1-billion-senior-secured-fixed-rate-financing",
)

def wavg(pairs): return sum(p*r for p,r in pairs)/sum(p for p,_ in pairs)
NB_EFF = wavg([(587.5,7.06),(1818.4,4.15),(612.5,6.87),(1818.4,4.88),(3105.0,4.98),(2100.0,5.20)])
CR_EFF = wavg([(1300,15),(3190,11),(3000,9),(2215,9),(1101,9),(2000,10),(1750,10),(2750,10),(1250,10),(2279,9),(2588,2),(4000,2),(4220,11),(189,12),(2837,7),(882,9)])
IR_EFF = wavg([(212.31,3.87),(233.389,4.22),(1000,0.37),(1150,0.45),(1150,1.18),(3000,1.18),(413,7.12),(525,7.05)])
OR_EFF = 1428*4/((7199+122342+7625+117712)/2)*100   # annualized Q1 interest exp / avg debt
OR_FEB = wavg([(3000,4.55),(3500,4.95),(3000,5.35),(5000,5.70),(2250,6.55),(5000,6.70),(2750,6.85)])

# Column spec: key, header, kind ('in' input / 'f' formula / 'txt')
COLS = [
 ("company","Company","txt"),("status","Ticker / status","txt"),("period","Latest period used","txt"),("pend","Period end","txt"),
 ("backlog","Revenue backlog / RPO ($M)","in"),("backlog_basis","Backlog basis","txt"),
 ("cmw","Contracted power (MW)","in"),("amw","Active / billing power (MW)","in"),("amw_basis","MW basis / date","txt"),
 ("rev_ltm","Revenue LTM ($M)","in"),("rev_q","Revenue latest Q ($M)","in"),("rev_qa","Revenue latest Q annualized ($M)","f"),
 ("yoy","Revenue YoY latest Q (%)","in"),("rev_mw","Revenue per active MW (annualized Q, $M/MW)","f"),
 ("ebitda_ltm","Adj. EBITDA LTM ($M)","in"),("ebitda_q","Adj. EBITDA latest Q ($M)","in"),
 ("m_ltm","Adj. EBITDA margin LTM (%)","f"),("m_q","Adj. EBITDA margin latest Q (%)","f"),
 ("ebitda_ntm","Adj. EBITDA NTM consensus ($M)","in"),("rev_fwd","Forward revenue: FY guide / cons / reported est. ($M)","in"),("rev_fwd_basis","Forward revenue basis","txt"),
 ("cod","Wtd avg effective cost of debt (%)","in"),("cod_recent","Recent issuance coupons / spreads","txt"),
 ("debt_bs","Total debt, latest BS ($M, ex-leases)","in"),("cash_bs","Cash & equivalents, latest BS ($M, ex-restricted)","in"),("pref","Preferred stock ($M)","in"),
 ("debt_pf_adj","Pro forma debt adj. after BS date ($M)","in"),("cash_pf_adj","Pro forma cash adj. after BS date ($M)","in"),("pf_bridge","Pro forma bridge","txt"),
 ("debt_pf","Pro forma total debt ($M)","f"),("cash_pf","Pro forma cash ($M)","f"),
 ("top_cust","Top customer (% of revenue, latest period)","in"),("anchors","Named anchor customers","txt"),
 ("price","Share price ($)","in"),("price_date","Price date","txt"),
 ("sh_bs","Shares, latest reported basis (M)","in"),("sh_pf","Shares, pro forma economic (M)","in"),
 ("mcap_bs","Market cap / equity value, reported basis ($M)","f"),("mcap_pf","Market cap / equity value, pro forma ($M)","f"),
 ("ev_agg","Aggregator EV ($M, StockAnalysis; incl. leases)","in"),
 ("ev_bs","EV, reported basis ($M)","f"),("ev_pf","EV, pro forma ($M) [used for multiples]","f"),
 ("ev_cmw","EV / contracted MW ($M/MW)","f"),("ev_amw","EV / active MW ($M/MW)","f"),
 ("ev_ebitda_ltm","EV / LTM adj. EBITDA (x)","f"),("ev_ebitda_qa","EV / annualized Q adj. EBITDA (x)","f"),("ev_ebitda_ntm","EV / NTM adj. EBITDA consensus (x)","f"),
 ("ev_rev_ltm","EV / LTM revenue (x)","f"),("ev_rev_qa","EV / annualized Q revenue (x)","f"),("ev_rev_fwd","EV / forward revenue (x)","f"),
 ("tag","Reported vs estimate","txt"),("notes","Notes / gaps","txt"),
]
K = {c[0]:get_column_letter(i+1) for i,c in enumerate(COLS)}

R = []  # rows
R.append(dict(company="Nebius Group", status="NBIS (Nasdaq)", period="Q2 2026 (6-K filed 2026-08-12)", pend="2026-06-30",
 backlog=37490.6, backlog_basis="ASC 606 RPO", cmw=3500, amw=170, amw_basis="Contracted '>3.5 GW' (Q1'26 letter, 2026-05-13; YE26 target 5 GW). Active ~170 MW at YE2025 (latest actual disclosed)",
 rev_ltm=1355.1, rev_q=582.3, yoy=454, ebitda_ltm=375.5, ebitda_q=236.2, ebitda_ntm=5411.7,
 rev_fwd=3200, rev_fwd_basis="FY26 guidance midpoint $3.0-3.4B (reaffirmed Q2)",
 cod=round(NB_EFF,2), cod_recent="Aug-26 converts: 0.50% 2030 ($3.45B) & 4.50% 2034 ($2.3B); Jul-26 $775M GPU-backed facility SOFR+2.50%; Mar-26 converts 1.25% 2031 / 2.63% 2033",
 debt_bs=8545.7, cash_bs=8042.1, pref=0, debt_pf_adj=5725.0, cash_pf_adj=6455.0,
 pf_bridge="+$775M secured facility (assumed fully drawn; drawn amount not disclosed); +$5,750M Aug converts; -$800M 2029/31 notes exchanged (face; carrying ~= $802M); cash +$775M +~$5,680M est. net convert proceeds (Skeptic sheet basis). Ignores Q3 op. cash flow/capex/ATM",
 top_cust=24, anchors="Microsoft, Meta (strategic LT contracts); Reflection, Cohere (Q2 deals); Customer C = 24%, D = 21%, E = 14% of Q2 rev (unnamed)",
 price=242.81, price_date=PRICE_DATE, sh_bs=292.921154, sh_pf=308.72, ev_agg=68710,
 tag="Reported (filings); NTM = consensus estimate; MW = company disclosure",
 notes="LTM = H1'26 + (FY25 529.8 - H1'25 156.0); LTM EBITDA = 365.7 + (-64.9 - -74.7). Reported-basis shares = 271.86M + 21.07M NVIDIA pre-funded warrants. Active MW is stale (YE25) so EV/active MW and rev/MW overstate; no period-end active/connected MW disclosed for Q2. YE26 connected guidance 800MW-1GW. NTM consensus = Q4'26-Q3'27 quarterly cons (MarketScreener via Skeptic sheet)."))
R.append(dict(company="CoreWeave", status="CRWV (Nasdaq)", period="Q2 2026 (10-Q filed 2026-08-12)", pend="2026-06-30",
 backlog=103700, backlog_basis="ASC 606 RPO $103.7B (company 'revenue backlog' $104.2B; excl. >$25B added early Q3)", cmw=3700, amw=1500, amw_basis="~3.7 GW contracted / 1.5 GW active at 2026-06-30 (Q2 PR); YE26 active guide >1.85 GW",
 rev_ltm=7590, rev_q=2575, yoy=round((2575/1212-1)*100,1), ebitda_ltm=4400, ebitda_q=1510, ebitda_ntm=14372,
 rev_fwd=12800, rev_fwd_basis="FY26 guidance midpoint $12.4-13.2B (Q2 call)",
 cod=round(CR_EFF,2), cod_recent="Sep-26 2.875% converts 2033 ($4.2B incl. option); Aug-26 DDTL 5.5 $2.6B (Ba2/BB+); Jun-26 9.625% notes 2032, 8.50% EUR notes 2032; Apr-26 9.75% notes 2031, 1.75% converts 2032; DDTL 5.0 SOFR+4.50%",
 debt_bs=35068, cash_bs=5539, pref=0, debt_pf_adj=4200, cash_pf_adj=3638.2,
 pf_bridge="+$4,200M Sep-22 2.875% converts (option exercised); cash +$4,137.0M est. net proceeds - $498.8M base capped call (additional capped call cost not disclosed). DDTL 5.5 undrawn: excluded. Shares already at 2026-07-31 count",
 top_cust=36, anchors="Customer A 36% (historically Microsoft; unnamed in 10-Q), B 26%, C 10% of Q2 rev; OpenAI (up to ~$6.5B order + earlier), Meta (~$21.0B Mar-26 order)",
 price=89.62, price_date=PRICE_DATE, sh_bs=551.536602, sh_pf=551.536602, ev_agg=95460,
 tag="Reported (filings); NTM = calendarized consensus estimate",
 notes="LTM = FY25 5,131 - H1'25 2,194 + H1'26 4,653. LTM adj. EBITDA = FY25 3,093 - 1,360 + 2,667. Debt = recourse + non-recourse carrying (ex leases). Cash incl. $15M marketable securities. NTM EBITDA = 0.25xFY26 cons 7,569 + 0.75xFY27 cons 16,640 (Skeptic sheet: Yahoo/MarketScreener). Aggregator EV $95.46B includes ~$16.5B lease liabilities -> not comparable."))
R.append(dict(company="IREN", status="IREN (Nasdaq)", period="Q4 FY26 / FY ended 2026-06-30 (10-K filed 2026-08-27)", pend="2026-06-30",
 backlog=16600, backlog_basis="RPO $5.1B + contracted lease value $11.4B (ASC 842)", cmw=5000, amw=420, amw_basis="~5 GW grid connection agreements (10-K); operating = ~380 MW Bitcoin + ~40 MW AI Cloud at 2026-06-30",
 rev_ltm=707.0, rev_q=137.2, yoy=round((137.2/187.292-1)*100,1), ebitda_ltm=245.7, ebitda_q=19.2, ebitda_ntm=None,
 rev_fwd=None, rev_fwd_basis="No revenue guidance; $4bn contracted ARR target for YE26 capacity (ARR, not revenue)",
 cod=round(IR_EFF,2), cod_recent="May-26 1.00% converts 2033 ($3.0B); USPP 5.96% fixed; DDTL SOFR+2.25% (swapped ~3.80%+spread); Aug-26 $2.4B Blue Owl/PIMCO GPU financing at 9.0% fixed (staged, undrawn at 10-K)",
 debt_bs=7593.0, cash_bs=5895.6, pref=0, debt_pf_adj=0, cash_pf_adj=0,
 pf_bridge="Shares 380.19M (6/30) -> 394.06M (8/14), incl. 12.6M for Mirantis acquisition; post-6/30 ATM proceeds not separately disclosed -> no cash adj.; Aug-26 $2.4B financing undrawn -> excluded",
 top_cust=None, anchors="Microsoft ($9.7B, 5-yr), NVIDIA ($3.4B, 5-yr) = 'substantial majority' of contracted revenue; also Perplexity, Cohere, Figure AI, Together AI, Fluidstack",
 price=41.76, price_date=PRICE_DATE, sh_bs=380.193608, sh_pf=394.058648, ev_agg=None,
 tag="Reported (filings)",
 notes="Top-customer % of revenue not disclosed (FY26 revenue 82% Bitcoin mining, no customer). Revenue/EBITDA still Bitcoin-heavy, so revenue- and EBITDA-based multiples are not comparable to AI-cloud peers. AI-only Q4 rev $70.5M / 40 MW = ~$7.1M per MW annualized. Contracted MW is grid power (gross), not IT MW."))
R.append(dict(company="Oracle (consolidated)", status="ORCL (NYSE)", period="Q1 FY27 (10-Q filed 2026-09-11)", pend="2026-08-31",
 backlog=664000, backlog_basis="ASC 606 RPO (company-wide; predominantly OCI per mgmt)", cmw=None, amw=None, amw_basis="Not disclosed in filings (Q1 PR: delivered 850 MW additional capacity in the quarter)",
 rev_ltm=71776, rev_q=19345, yoy=round((19345/14926-1)*100,1), ebitda_ltm=38752, ebitda_q=11213, ebitda_ntm=None,
 rev_fwd=90000, rev_fwd_basis="FY27 (May-27) revenue guidance 'at least $90B'",
 cod=round(OR_EFF,2), cod_recent="Feb-26 $25B notes: 4.55%% 2029 ... 6.85%% 2066 (fixed-tranche wtd coupon %.2f%%) + FRN 2029; 6.50%% Series D mandatory convertible preferred" % OR_FEB,
 debt_bs=125337, cash_bs=37077, pref=4954, debt_pf_adj=0, cash_pf_adj=0, pf_bridge="Share count already 2026-09-07 (post $20B ATM); no material post-BS debt events found",
 top_cust=None, anchors="OpenAI (Stargate) and other AI labs; customer % not disclosed",
 price=142.30, price_date=PRICE_DATE, sh_bs=3023.736, sh_pf=3023.736, ev_agg=None,
 tag="Reported (filings); EBITDA derived",
 notes="Adj. EBITDA (derived, not company-reported) = GAAP op. income + depreciation + amortization + SBC; LTM = FY26 - Q1FY26 + Q1FY27 (XBRL). Cost of debt = annualized Q1 interest expense / avg debt (net of capitalized interest -> understated). EV includes $4,954M preferred. Whole-company EV; OCI not separable."))
R.append(dict(company="Oracle - OCI segment (memo)", status="ORCL segment", period="Q1 FY27", pend="2026-08-31",
 backlog=None, backlog_basis="RPO not split by segment", cmw=None, amw=None, amw_basis="n/d",
 rev_ltm=22142, rev_q=7388, yoy=round((7388/3347-1)*100,1), ebitda_ltm=None, ebitda_q=None, ebitda_ntm=None, rev_fwd=None, rev_fwd_basis="",
 cod=None, cod_recent="", debt_bs=None, cash_bs=None, pref=None, debt_pf_adj=None, cash_pf_adj=None, pf_bridge="",
 top_cust=None, anchors="", price=None, price_date="", sh_bs=None, sh_pf=None, ev_agg=None,
 tag="Reported (cloud infrastructure revenue line)",
 notes="Cloud infrastructure (IaaS) revenue only. LTM = FY26 18,101 - Q1FY26 3,347 + Q1FY27 7,388. No segment EBITDA, debt or EV -> valuation multiples left blank."))
R.append(dict(company="Nscale", status="Private (S-1 filed 2026-09-18)", period="H1 2026 (S-1)", pend="2026-06-30",
 backlog=103400, backlog_basis="TCV of signed contracts at 2026-08-31 (not ASC 606 RPO); ~$2.6B active", cmw=1370, amw=None, amw_basis="~1.37 GW active + contracted at 2026-08-31; active MW not split out (~25k active GPUs of 461k)",
 rev_ltm=163.2, rev_q=None, yoy=round((140.6/10.4-1)*100,0), ebitda_ltm=-265.6, ebitda_q=None, ebitda_ntm=None, rev_fwd=None, rev_fwd_basis="",
 cod=None, cod_recent="Macquarie senior facility SOFR+7.50% (eff. 21.5%); Iceland DDTL SOFR+5.50%; Sep-26 $3.36B pre-IPO convertible notes (terms not disclosed)",
 debt_bs=137.4, cash_bs=1478.4, pref=0, debt_pf_adj=0, cash_pf_adj=0, pf_bridge="Sep-26 $2.36B convertible notes funded (+$1B NVIDIA Nov) not added: convert to equity at IPO; valuation basis stale vs that round",
 top_cust=52, anchors="Microsoft (up to $43.8B through 2033), Anthropic (up to $44.6B, Aug-26) = ~85% of TCV; OpenAI",
 price=None, price_date="Series C 2026-03-09", sh_bs=None, sh_pf=None, ev_agg=None, mcap_override=14600,
 tag="Reported (S-1) for financials; valuation = last private round",
 notes="Equity value = Series C valuation $14.6B (Mar-26), pre-convert; FT-reported IPO target ~$35B not used. LTM = H1'26 140.6 + FY25 33.0 - H1'25 10.4; LTM adj. EBITDA = -199.2 + -81.7 - (-15.3). YoY is H1/H1 (no quarterly split). Weighted avg cost of debt not computable (debt small, mixed)."))
R.append(dict(company="Crusoe", status="Private", period="Series F (2026-09-17)", pend="2026-09-17",
 backlog=140000, backlog_basis="'$140B+' total contract value (company PR; term not disclosed)", cmw=6000, amw=1000, amw_basis="'6GW+ gross contracted', '~1 GW gross delivered and operational' (Series F PR)",
 rev_ltm=None, rev_q=None, yoy=None, ebitda_ltm=None, ebitda_q=None, ebitda_ntm=None,
 rev_fwd=2000, rev_fwd_basis="'On track for $2B' 2026 revenue (Newcomer, from financials obtained) - ESTIMATE",
 cod=None, cod_recent="Not disclosed", debt_bs=None, cash_bs=None, pref=None, debt_pf_adj=None, cash_pf_adj=None, pf_bridge="No balance sheet disclosed",
 top_cust=None, anchors="OpenAI, Microsoft, Meta (Bloomberg)",
 price=None, price_date="Series F 2026-09-17", sh_bs=None, sh_pf=None, ev_agg=None, mcap_override=30900,
 tag="Company PR / press; revenue = press estimate",
 notes="EV proxy = $30.9B post-money equity only (debt/cash not disclosed) -> EV understated vs. levered peers. Gross MW (not IT MW); revenue includes data-center development mix."))
R.append(dict(company="Lambda", status="Private", period="Series E (Nov-2025); press Aug-2026", pend="2025-11 / 2026-08",
 backlog=None, backlog_basis="Not disclosed", cmw=None, amw=None, amw_basis="Not disclosed (target '3GW' liquid-cooled capacity; ~15 US data centers)",
 rev_ltm=None, rev_q=None, yoy=None, ebitda_ltm=None, ebitda_q=None, ebitda_ntm=None,
 rev_fwd=1500, rev_fwd_basis="'>$1.5B' 2026 revenue (Bloomberg-reported expectation) - ESTIMATE; $760M ARR end-2025 (Sacra estimate)",
 cod=None, cod_recent="Oct-26 $1B IG fixed-rate DDTL (Baa1/A(low); rate n/d); Aug-26 ~$1B broadly syndicated loan; earlier $917M GPU-backed debt",
 debt_bs=None, cash_bs=None, pref=None, debt_pf_adj=None, cash_pf_adj=None, pf_bridge="No balance sheet disclosed",
 top_cust=None, anchors="Microsoft (multi-billion, multi-year GB300 deal), NVIDIA (~$1.5B / 18k GPU leaseback)",
 price=None, price_date="Series E 2025-11", sh_bs=None, sh_pf=None, ev_agg=None, mcap_override=5900,
 tag="Press / estimates only",
 notes="Equity value = $5.9B Series E post-money (PitchBook cites $5.43B); in talks for up to $3B at $12B+ (Bloomberg, not closed). No MW, EBITDA, or debt totals -> most multiples blank."))

def f(col,r,expr): return expr
def formulas(r):
    g=lambda k: f"{K[k]}{r}"
    return {
     "rev_qa": f'=IF({g("rev_q")}="","",{g("rev_q")}*4)',
     "rev_mw": f'=IF(OR({g("rev_q")}="",{g("amw")}=""),"",{g("rev_q")}*4/{g("amw")})',
     "m_ltm": f'=IF(OR({g("ebitda_ltm")}="",{g("rev_ltm")}=""),"",{g("ebitda_ltm")}/{g("rev_ltm")}*100)',
     "m_q": f'=IF(OR({g("ebitda_q")}="",{g("rev_q")}=""),"",{g("ebitda_q")}/{g("rev_q")}*100)',
     "debt_pf": f'=IF({g("debt_bs")}="","",{g("debt_bs")}+N({g("debt_pf_adj")}))',
     "cash_pf": f'=IF({g("cash_bs")}="","",{g("cash_bs")}+N({g("cash_pf_adj")}))',
     "mcap_bs": f'=IF({g("price")}="","",{g("price")}*{g("sh_bs")})',
     "mcap_pf": f'=IF({g("price")}="","",{g("price")}*{g("sh_pf")})',
     "ev_bs": f'=IF({g("mcap_bs")}="","",{g("mcap_bs")}+N({g("debt_bs")})-N({g("cash_bs")})+N({g("pref")}))',
     "ev_pf": f'=IF({g("mcap_pf")}="","",{g("mcap_pf")}+N({g("debt_pf")})-N({g("cash_pf")})+N({g("pref")}))',
     "ev_cmw": f'=IF(OR({g("ev_pf")}="",{g("cmw")}=""),"",{g("ev_pf")}/{g("cmw")})',
     "ev_amw": f'=IF(OR({g("ev_pf")}="",{g("amw")}=""),"",{g("ev_pf")}/{g("amw")})',
     "ev_ebitda_ltm": f'=IF(OR({g("ev_pf")}="",{g("ebitda_ltm")}=""),"",IF({g("ebitda_ltm")}<=0,"n/m",{g("ev_pf")}/{g("ebitda_ltm")}))',
     "ev_ebitda_qa": f'=IF(OR({g("ev_pf")}="",{g("ebitda_q")}=""),"",IF({g("ebitda_q")}<=0,"n/m",{g("ev_pf")}/({g("ebitda_q")}*4)))',
     "ev_ebitda_ntm": f'=IF(OR({g("ev_pf")}="",{g("ebitda_ntm")}=""),"",{g("ev_pf")}/{g("ebitda_ntm")})',
     "ev_rev_ltm": f'=IF(OR({g("ev_pf")}="",{g("rev_ltm")}=""),"",{g("ev_pf")}/{g("rev_ltm")})',
     "ev_rev_qa": f'=IF(OR({g("ev_pf")}="",{g("rev_q")}=""),"",{g("ev_pf")}/({g("rev_q")}*4))',
     "ev_rev_fwd": f'=IF(OR({g("ev_pf")}="",{g("rev_fwd")}=""),"",{g("ev_pf")}/{g("rev_fwd")})',
    }

def pyvals(d):
    v=dict(d); n=lambda x: x if x is not None else 0
    v["rev_qa"]= d["rev_q"]*4 if d.get("rev_q") is not None else None
    v["rev_mw"]= d["rev_q"]*4/d["amw"] if d.get("rev_q") is not None and d.get("amw") else None
    v["m_ltm"]= d["ebitda_ltm"]/d["rev_ltm"]*100 if d.get("ebitda_ltm") is not None and d.get("rev_ltm") else None
    v["m_q"]= d["ebitda_q"]/d["rev_q"]*100 if d.get("ebitda_q") is not None and d.get("rev_q") else None
    v["debt_pf"]= d["debt_bs"]+n(d.get("debt_pf_adj")) if d.get("debt_bs") is not None else None
    v["cash_pf"]= d["cash_bs"]+n(d.get("cash_pf_adj")) if d.get("cash_bs") is not None else None
    mo=d.get("mcap_override")
    v["mcap_bs"]= d["price"]*d["sh_bs"] if d.get("price") else mo
    v["mcap_pf"]= d["price"]*d["sh_pf"] if d.get("price") else mo
    v["ev_bs"]= v["mcap_bs"]+n(d.get("debt_bs"))-n(d.get("cash_bs"))+n(d.get("pref")) if v["mcap_bs"] is not None else None
    v["ev_pf"]= v["mcap_pf"]+n(v["debt_pf"])-n(v["cash_pf"])+n(d.get("pref")) if v["mcap_pf"] is not None else None
    ev=v["ev_pf"]
    dv=lambda a,b: (a/b if (a is not None and b) else None)
    v["ev_cmw"]=dv(ev,d.get("cmw")); v["ev_amw"]=dv(ev,d.get("amw"))
    e=d.get("ebitda_ltm"); v["ev_ebitda_ltm"]= None if (ev is None or e is None) else ("n/m" if e<=0 else ev/e)
    e=d.get("ebitda_q"); v["ev_ebitda_qa"]= None if (ev is None or e is None) else ("n/m" if e<=0 else ev/(e*4))
    v["ev_ebitda_ntm"]=dv(ev,d.get("ebitda_ntm")); v["ev_rev_ltm"]=dv(ev,d.get("rev_ltm"))
    v["ev_rev_qa"]=dv(ev,v["rev_qa"]); v["ev_rev_fwd"]=dv(ev,d.get("rev_fwd"))
    return v

wb=Workbook(); ws=wb.active; ws.title="Comps"
hdr_fill=PatternFill("solid",fgColor="1F3864"); fx_fill=PatternFill("solid",fgColor="EAF1FB")
ws.append([c[1] for c in COLS])
for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=hdr_fill; c.alignment=Alignment(wrap_text=True,vertical="top")
VALS=[]
for i,d in enumerate(R):
    r=i+2; fm=formulas(r); v=pyvals(d); VALS.append(v)
    for j,(k,h,kind) in enumerate(COLS):
        cell=ws.cell(row=r,column=j+1)
        if kind=="f":
            # privates: market cap is an input (valuation) not price x shares
            if k in("mcap_bs","mcap_pf") and d.get("mcap_override") is not None:
                cell.value=d["mcap_override"]
            else:
                cell.value=fm[k]; cell.fill=fx_fill
        else:
            cell.value=d.get(k)
        if isinstance(cell.value,(int,float)) or kind=="f": cell.number_format='#,##0.0'
ws.freeze_panes="B2"; ws.row_dimensions[1].height=75
for j,(k,h,kind) in enumerate(COLS):
    ws.column_dimensions[get_column_letter(j+1)].width = 60 if k in("notes","pf_bridge","cod_recent","anchors","amw_basis") else (28 if kind=="txt" else 14)
n=len(R)+3
ws.cell(row=n,column=1,value="Legend: blue-shaded cells are Excel formulas; others are sourced inputs (see Comps_Sources). Multiples use PRO FORMA EV. EV = equity value + debt (ex-leases, carrying) - unrestricted cash + preferred. Private-company 'EV' = last post-money equity value only (no net debt). Prices: Yahoo Finance close %s." % PRICE_DATE)
ws.cell(row=n+1,column=1,value="Blank = not disclosed / not found (reason in Notes). 'n/m' = negative EBITDA. MW definitions differ (IT vs gross/grid power); see MW basis column.")

# ---- Revenue per MW block ----
b=n+3
ws.cell(row=b,column=1,value="Revenue per MW block (latest-quarter annualized revenue / active billing MW)").font=Font(bold=True)
hdr=["Company","Annualized Q revenue ($M)","Active MW","Revenue per active MW ($M/MW)","Basis / caveat","Source"]
for j,h in enumerate(hdr): ws.cell(row=b+1,column=j+1,value=h).font=Font(bold=True)
blk=[
 ("CoreWeave (period-end MW)",2575*4,1500,"Q2'26 rev x4 / 1.5 GW active at 6/30 (understates: MW ramped ~1.0->1.5 GW in Q2)",U["CR_PR"]),
 ("CoreWeave (avg Q2 MW est.)",2575*4,1250,"ESTIMATE: avg of ~1.0 GW (1.5 GW less 'nearly 500 MW' added) and 1.5 GW",U["CR_PR"]),
 ("IREN - AI Cloud only",70.5*4,40,"Q4 FY26 AI Cloud rev x4 / ~40 MW operating AI capacity at 6/30",U["IR_10K"]),
 ("IREN - total",137.2*4,420,"Includes Bitcoin mining (~380 MW); not comparable",U["IR_PR"]),
 ("Nebius (stale MW)",582.3*4,170,"Q2'26 rev x4 / ~170 MW active at YE2025 (latest disclosed) -> OVERSTATES; Q2 active MW not disclosed",U["NB_LET4"]),
 ("Crusoe (estimate)",2000,1000,"ESTIMATE: 2026E $2B (press) / ~1 GW gross operational; gross not IT MW",U["CU_F"]),
]
for i,(c,rv,mw,basis,src) in enumerate(blk):
    r=b+2+i
    ws.cell(row=r,column=1,value=c); ws.cell(row=r,column=2,value=rv); ws.cell(row=r,column=3,value=mw)
    x=ws.cell(row=r,column=4,value=f"=B{r}/C{r}"); x.fill=fx_fill; x.number_format='0.0'
    ws.cell(row=r,column=5,value=basis); ws.cell(row=r,column=6,value=src)
r=b+2+len(blk)+1
ws.cell(row=r,column=1,value="Company-disclosed contracted yields (new deals, not realized): Nebius Q2 deals $20-25M ACV per MW (Q2 letter); IREN recent 3-yr contracts >$20M revenue per IT MW (FY26 PR). Test vs ~$12M/MW implied by price (Skeptic).")
BLK=blk

# ---- EV bridge sheet ----
eb=wb.create_sheet("EV_Bridge")
eb.append(["Company","Basis","Market cap ($M)","+ Debt ($M)","- Cash ($M)","+ Preferred ($M)","EV ($M)","Shares (M)","Notes","Source"])
for c in eb[1]: c.font=Font(bold=True)
rowmap={d["company"]:i+2 for i,d in enumerate(R)}
for d in R[:3]+[R[3]]:
    rr=rowmap[d["company"]]
    eb.append([d["company"],"Aggregator (StockAnalysis, incl. leases)",None,None,None,None,d.get("ev_agg"),None,"From Skeptic sheet; debt incl. operating/finance leases" if d.get("ev_agg") else "Not retrieved (site blocks automated fetch)", U["SA_NB"] if d["company"].startswith("Neb") else (U["SA_CR"] if d["company"]=="CoreWeave" else "")])
    eb.append([d["company"],"Reported basis (latest BS)",f"=Comps!{K['mcap_bs']}{rr}",f"=Comps!{K['debt_bs']}{rr}",f"=Comps!{K['cash_bs']}{rr}",f"=N(Comps!{K['pref']}{rr})",f"=Comps!{K['ev_bs']}{rr}",f"=Comps!{K['sh_bs']}{rr}",d["period"],""])
    eb.append([d["company"],"Pro forma (post-BS events)",f"=Comps!{K['mcap_pf']}{rr}",f"=Comps!{K['debt_pf']}{rr}",f"=Comps!{K['cash_pf']}{rr}",f"=N(Comps!{K['pref']}{rr})",f"=Comps!{K['ev_pf']}{rr}",f"=Comps!{K['sh_pf']}{rr}",d["pf_bridge"],""])
eb.append([]); eb.append(["Reconciliation vs Skeptic sheet (/workspace/nebius/skeptic/01_priced_in.xlsx)"])
for line in [
 "Nebius pro forma EV: this file %.0f vs Skeptic $74,730M - same bridge (308.72M sh, debt 14,270.7, cash 14,497.1); difference = rounding." % VALS[0]["ev_pf"],
 "Nebius $775M facility: 6-K FS note dates agreement 2026-07-10; Skeptic cites 2026-07-17 press release (announcement date). Same facility.",
 "Nebius exchanged notes: Skeptic deducts $800M face; 6/30 carrying of 80% of 2029/2031 series = 0.8x(504.1+498.4) = $802M -> immaterial.",
 "CoreWeave EV: Skeptic uses aggregator $95.46B (debt incl. ~$16.5B leases, pre-Sep converts). This file: pro forma ex-leases %.0f. Skeptic's 3.6x FY27 rev / 5.7x FY27 EBITDA multiples are on the lease-inclusive EV; on this file's EV: %.2fx FY27 rev / %.2fx FY27 EBITDA." % (VALS[1]["ev_pf"], VALS[1]["ev_pf"]/26260, VALS[1]["ev_pf"]/16640),
 "Nebius shares: StockAnalysis 274.1M vs company 271.86M (6/30) - aggregator basis undisclosed; this file uses company count + warrants.",
 "Nebius row reconciled vs Filings & Capital fact base (/workspace/nebius/filings/nebius_fact_base.md): revenue, adj. EBITDA, RPO 37,490.6, debt 8,545.7, cash 8,042.1, customers C/D/E 24/21/14%, shares, warrants, MW (>3.5 GW May-26; ~170 MW active YE25) all match. Fact base notes $775M drawn amount not disclosed.",
 "Nebius aggregator EV ($68.71B) uses 6/30 BS incl. leases, excludes warrants/Aug converts/exchange shares -> understated vs pro forma.",
]: eb.append([line])
for col,w in zip("ABCDEFGHIJ",[24,34,14,12,12,12,14,11,90,60]): eb.column_dimensions[col].width=w

# ---- Sources sheet ----
src=wb.create_sheet("Comps_Sources")
src.append(["company","metric","value","period","source_url","note","tag"])
for c in src[1]: c.font=Font(bold=True)
S=lambda *a: src.append(list(a))
# Nebius
S("Nebius","RPO",37490.6,"2026-06-30",U["NB_FS"],"Note: Revenue - RPO $37,490.6M; 36% within 24 months","Reported")
S("Nebius","Contracted power (MW)",3500,"as of 2026-05-13",U["NB_LET1"],"'Contracted capacity already exceeds 3.5 GW'; Q2 letter raises YE26 target to 5 GW (current figure not stated)","Reported (floor)")
S("Nebius","YE26 contracted target / connected guidance","5,000 MW / 800-1,000 MW","FY2026 guidance",U["NB_LET2"],"Connected 800MW-1GW reaffirmed on Q2 call: "+U["NB_CALL"],"Guidance")
S("Nebius","Active power (MW)",170,"2025-12-31",U["NB_LET4"],"'~170 MW of active power' at YE25; no later actual disclosed","Reported (stale)")
S("Nebius","Revenue Q2'26",582.3,"Q2 2026",U["NB_PR"],"","Reported")
S("Nebius","Revenue H1'26 / H1'25",  "981.3 / 156.0","H1",U["NB_PR"],"","Reported")
S("Nebius","Revenue FY2025",529.8,"FY2025",U["NB_20F"],"XBRL Revenues CY2025 (data.sec.gov companyfacts CIK0001513845)","Reported")
S("Nebius","Revenue LTM",1355.1,"LTM to 2026-06-30","derived","981.3 + 529.8 - 156.0","Derived")
S("Nebius","Revenue YoY Q2",454,"Q2 2026",U["NB_PR"],"","Reported")
S("Nebius","Adj. EBITDA Q2'26 / H1'26 / H1'25","236.2 / 365.7 / -74.7","",U["NB_PR"],"","Reported")
S("Nebius","Adj. EBITDA FY2025",-64.9,"FY2025",U["NB_LET4"],"Q4'25 letter table","Reported")
S("Nebius","Adj. EBITDA LTM",375.5,"LTM","derived","365.7 + (-64.9 - -74.7)","Derived")
S("Nebius","NTM adj. EBITDA consensus",5411.7,"Q4'26-Q3'27",U["NB_CONS"],"Sum of quarterly cons 670.7+1277+1581+1883 as captured in Skeptic sheet 2026-10-04","Estimate (consensus)")
S("Nebius","FY26 revenue guidance","3,000-3,400","FY2026","https://www.tradingkey.com/news/transcripts/262106466-tradingkey","Reaffirmed Q2; adj. EBITDA margin ~40%","Guidance")
S("Nebius","Wtd effective cost of debt (%)",round(NB_EFF,2),"2026-06-30",U["NB_FS"],"Principal-weighted effective rates of 6 convert series (Note 12)","Derived")
S("Nebius","Aug-26 converts 0.50% 2030 / 4.50% 2034","3,450 / 2,300","2026-08-24",U["NB_CONV"],"","Reported")
S("Nebius","Secured facility SOFR+2.50%",775,"2026-07-10",U["NB_FS"],"Subsequent events note","Reported")
S("Nebius","Debt (current + non-current, carrying)",8545.7,"2026-06-30",U["NB_PR"],"46.7 + 8,499.0","Reported")
S("Nebius","Cash & equivalents",8042.1,"2026-06-30",U["NB_PR"],"Excl. restricted cash 1,056.0","Reported")
S("Nebius","Est. net proceeds Aug converts",5680,"2026-08",U["NB_CONVP"],"Per Skeptic sheet (company estimate incl. option)","Estimate (company)")
S("Nebius","Top customer % revenue",24,"Q2 2026",U["NB_FS"],"Customer C 24%, D 21%, E 14% (unnamed)","Reported")
S("Nebius","Anchors Microsoft/Meta",None,"",U["NB_LET1"],"Q1 letter names Microsoft, Meta strategic contracts; Q2 letter names Reflection, Cohere","Reported")
S("Nebius","Shares outstanding",271.855218,"2026-06-30",U["NB_PR"],"","Reported")
S("Nebius","NVIDIA pre-funded warrant shares",21.065936,"2026-03-11",U["NB_FS"],"Note 14; exercise price $0.0001","Reported")
S("Nebius","Exchange shares",15.8,"2026-08-24",U["NB_CONV"],"~15.8M Class A for $800M of 2029/2031 notes","Reported")
S("Nebius","Share price",242.81,PRICE_DATE,Y.format(t="NBIS"),"Matches Skeptic sheet","Market")
S("Nebius","Aggregator EV",68710,"2026-10-04",U["SA_NB"],"Via Skeptic sheet","Aggregator")
# CoreWeave
S("CoreWeave","RPO",103700,"2026-06-30",U["CR_XBRL"],"XBRL RevenueRemainingPerformanceObligation; PR 'revenue backlog' ~$104B ($104.2B deck)","Reported")
S("CoreWeave","Contracted / active power (MW)","3,700 / 1,500","2026-06-30",U["CR_PR"],"'approximately 3.7 GW' contracted; active 'reach 1.5 GW'","Reported")
S("CoreWeave","Revenue Q2'26 / Q2'25",  "2,575 / 1,212","Q2",U["CR_PR"],"","Reported")
S("CoreWeave","Revenue FY2025",5131,"FY2025",U["CR_XBRL"],"","Reported")
S("CoreWeave","Revenue LTM",7590,"LTM","derived","5,131 - 2,194 + 4,653","Derived")
S("CoreWeave","Adj. EBITDA Q2 / H1'26 / H1'25","1,510 / 2,667 / 1,360","",U["CR_PR"],"","Reported")
S("CoreWeave","Adj. EBITDA FY2025",3093,"FY2025",U["CR_Q4"],"","Reported")
S("CoreWeave","Adj. EBITDA LTM",4400,"LTM","derived","3,093 - 1,360 + 2,667","Derived")
S("CoreWeave","NTM adj. EBITDA (calendarized cons)",14372,"Oct-26-Sep-27",U["CR_CONS"],"0.25 x FY26 7,569 + 0.75 x FY27 16,640 (Skeptic sheet; MarketScreener/Yahoo)","Estimate (consensus)")
S("CoreWeave","FY26 revenue guidance","12,400-13,200","FY2026",U["CR_CALL"],"","Guidance")
S("CoreWeave","Wtd effective cost of debt (%)",round(CR_EFF,2),"2026-06-30",U["CR_10Q"],"Principal-weighted effective rates, Note 10 debt table","Derived")
S("CoreWeave","Sep-26 2.875% converts 2033",4200,"2026-09-22",U["CR_CONV8K"],"$3.7B + $0.5B option exercised; net proceeds est. $3,644.5M/$4,137.0M; capped call $498.8M (base): "+U["CR_CONV"],"Reported")
S("CoreWeave","DDTL 5.5",2600,"2026-08-10",U["CR_DDTL"],"Delayed draw; excluded from pro forma","Reported")
S("CoreWeave","Debt (carrying, recourse + non-recourse)",35068,"2026-06-30",U["CR_PR"],"6,235+1,278+25,170+2,385; excl. finance leases 221","Reported")
S("CoreWeave","Cash + marketable securities",5539,"2026-06-30",U["CR_PR"],"5,524 + 15; excl. restricted 1,380","Reported")
S("CoreWeave","Top customer % revenue",36,"Q2 2026",U["CR_10Q"],"A 36%, B 26%, C 10%; OpenAI & Meta orders named in risk factors","Reported")
S("CoreWeave","Shares (A+B)",551.536602,"2026-07-31",U["CR_10Q"],"458,871,690 A + 92,664,912 B","Reported")
S("CoreWeave","Share price",89.62,PRICE_DATE,Y.format(t="CRWV"),"","Market")
S("CoreWeave","Aggregator EV",95460,"2026-10-02",U["SA_CR"],"Via Skeptic sheet; incl. leases","Aggregator")
# IREN
S("IREN","RPO + contracted lease value",16600,"2026-06-30",U["IR_10K"],"RPO $5.1B + lease $11.4B","Reported")
S("IREN","Grid-connected power agreements (MW)",5000,"2026-06-30",U["IR_10K"],"'approximately 5GW of total power capacity'","Reported")
S("IREN","Operating capacity (MW)","~380 BTC + ~40 AI","2026-06-30",U["IR_10K"],"","Reported")
S("IREN","Revenue FY26 / Q4 FY26",  "707.0 / 137.2","FY26",U["IR_PR"],"","Reported")
S("IREN","Revenue Q4 FY25",187.292,"Q4 FY25",U["IR_XBRL"],"FY25 501.023 - 9M FY25 313.731","Derived")
S("IREN","Adj. EBITDA FY26 / Q4",  "245.7 / 19.2","FY26",U["IR_PR"],"","Reported")
S("IREN","Wtd effective cost of debt (%)",round(IR_EFF,2),"2026-06-30",U["IR_10K"],"Principal-weighted effective rates, Note 23 (converts 0.37%-4.22%, DDTL 7.12%, USPP 7.05%)","Derived")
S("IREN","Microsoft GPU financing wtd rate (%)",6.0,"FY26",U["IR_PR"],"USPP + DDTL excl. fees; Blue Owl/PIMCO $2.4B at 9.0% fixed","Reported")
S("IREN","Debt (carrying)",7593.0,"2026-06-30",U["IR_PR"],"169.4 + 7,423.6; excl. finance leases 243.8","Reported")
S("IREN","Cash & equivalents",5895.6,"2026-06-30",U["IR_PR"],"Excl. restricted 1,724.0","Reported")
S("IREN","Anchors",None,"",U["IR_10K"],"Microsoft $9.7B, NVIDIA $3.4B = substantial majority of contracted revenue; customer list: "+U["IR_0720"],"Reported")
S("IREN","Shares 6/30 / 8/14","380.19 / 394.06","",U["IR_10K"],"","Reported")
S("IREN","Share price",41.76,PRICE_DATE,Y.format(t="IREN"),"","Market")
# Oracle
S("Oracle","RPO",664000,"2026-08-31",U["OR_10Q"],"13% within 12 months","Reported")
S("Oracle","Revenue Q1 FY27 / Q1 FY26 / FY26","19,345 / 14,926 / 67,357","",U["OR_PR"],"","Reported")
S("Oracle","OCI revenue Q1FY27 / Q1FY26 / FY26","7,388 / 3,347 / 18,101","",U["OR_PR"],"Cloud infrastructure line","Reported")
S("Oracle","Adj. EBITDA LTM (derived)",38752,"LTM to 2026-08-31",U["OR_XBRL"],"Op inc 23,057 + dep 9,428 + amort 1,453 + SBC 4,814 (FY26 - Q1FY26 + Q1FY27)","Derived")
S("Oracle","Adj. EBITDA Q1 FY27 (derived)",11213,"Q1 FY27",U["OR_PR"],"6,728 + 3,156 + 202 + 1,127","Derived")
S("Oracle","Cost of debt (%)",round(OR_EFF,2),"Q1 FY27",U["OR_PR"],"Interest exp 1,428 x4 / avg debt (net of capitalized interest)","Derived")
S("Oracle","Feb-26 notes wtd fixed coupon (%)",round(OR_FEB,2),"2026-02-04",U["OR_FEB"],"$25B issue: 4.55%-6.85% + FRN","Reported")
S("Oracle","Debt",125337,"2026-08-31",U["OR_PR"],"7,625 + 117,712","Reported")
S("Oracle","Cash + marketable securities",37077,"2026-08-31",U["OR_PR"],"36,369 + 708","Reported")
S("Oracle","Preferred (6.50% Series D mandatory convertible)",4954,"2026-08-31",U["OR_10Q"],"","Reported")
S("Oracle","Shares",3023.736,"2026-09-07",U["OR_XBRL"],"dei:EntityCommonStockSharesOutstanding (10-Q cover)","Reported")
S("Oracle","FY27 revenue guidance",90000,"FY27",U["OR_PR"],"'at least $90 billion'","Guidance")
S("Oracle","Share price",142.30,PRICE_DATE,Y.format(t="ORCL"),"","Market")
# Nscale
S("Nscale","TCV",103400,"2026-08-31",U["NS_S1"],"TCV, not RPO; ~$2.6B active","Reported (S-1)")
S("Nscale","Active + contracted capacity (MW)",1370,"2026-08-31",U["NS_S1"],"1 GW owned, 200 MW leased, 165 MW colo","Reported (S-1)")
S("Nscale","Revenue H1'26 / H1'25 / FY25","140.6 / 10.4 / 33.0","",U["NS_S1"],"","Reported (S-1)")
S("Nscale","Adj. EBITDA H1'26 / H1'25 / FY25","-199.2 / -15.3 / -81.7","",U["NS_S1"],"","Reported (S-1)")
S("Nscale","Largest customer % revenue",52,"H1 2026",U["NS_S1"],"73% in FY25","Reported (S-1)")
S("Nscale","Debt / cash","137.4 / 1,478.4","2026-06-30",U["NS_S1"],"","Reported (S-1)")
S("Nscale","Valuation (Series C)",14600,"2026-03-09",U["NS_C"],"","Reported (company)")
S("Nscale","Pre-IPO convertible notes",3360,"2026-09-25",U["NS_CONV"],"$2.36B funded + $1B NVIDIA (Nov)","Reported (company)")
# Crusoe
S("Crusoe","Post-money valuation",30900,"2026-09-17",U["CU_F"],"Series F initial close $3.9B","Reported (company)")
S("Crusoe","Contracted / operational capacity (MW)","6,000+ / ~1,000 (gross)","2026-09-17",U["CU_F"],"","Reported (company)")
S("Crusoe","TCV","140,000+","2026-09-17",U["CU_F"],"","Reported (company)")
S("Crusoe","2026 revenue","~2,000","2026E",U["CU_REV"],"'on track for $2 billion' (financials obtained by Newcomer)","Estimate (press)")
S("Crusoe","Anchors",None,"",U["CU_BBG"],"OpenAI, Microsoft, Meta","Press")
# Lambda
S("Lambda","Post-money valuation",5900,"2025-11",U["LA_TFN"],"Series E $1.5B; PitchBook cites $5.43B; talks at $12B+: "+U["LA_BBG"],"Press")
S("Lambda","2026 revenue expectation","1,500+","2026E",U["LA_BBG"],"Bloomberg-reported","Estimate (press)")
S("Lambda","ARR end-2025",760,"2025-12",U["LA_SACRA"],"Sacra estimate","Estimate (third party)")
S("Lambda","Debt financings",None,"2026",U["LA_DEBT"],"$1B IG DDTL (Oct-26); ~$1B BSL (Aug-26); rates/total not disclosed","Reported (company)")
S("Lambda","Anchors",None,"",U["LA_DCD"],"Microsoft multi-billion deal; NVIDIA ~$1.5B leaseback","Press")
src.column_dimensions["B"].width=40; src.column_dimensions["E"].width=90; src.column_dimensions["F"].width=70

out="/workspace/nebius/parts/_rebuild/comps.xlsx"; wb.save(out)
# CSVs (values)
with open("/workspace/nebius/parts/comps.csv","w",newline="") as fh:
    w=csv.writer(fh); w.writerow([c[1] for c in COLS])
    for v in VALS: w.writerow([("" if v.get(k) is None else (round(v[k],2) if isinstance(v.get(k),float) else v[k])) for k,_,_ in COLS])
with open("/workspace/nebius/parts/comps_sources.csv","w",newline="") as fh:
    w=csv.writer(fh)
    for row in src.iter_rows(values_only=True): w.writerow(row)
with open("/workspace/nebius/parts/comps_rev_per_mw.csv","w",newline="") as fh:
    w=csv.writer(fh); w.writerow(["Company","Annualized Q revenue ($M)","Active MW","Revenue per active MW ($M/MW)","Basis","Source"])
    for c,rv,mw,basis,s in BLK: w.writerow([c,rv,mw,round(rv/mw,2),basis,s])
json.dump([{k:v.get(k) for k,_,_ in COLS} for v in VALS],open("/workspace/nebius/parts/comps_values.json","w"),indent=1,default=str)
import subprocess, shutil, os
subprocess.run(["soffice","--headless","--calc","--convert-to","xlsx","--outdir","/tmp/lo_comps",out],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=180)
if os.path.exists("/tmp/lo_comps/comps.xlsx"): shutil.copy("/tmp/lo_comps/comps.xlsx",out)  # cache formula results
from openpyxl import load_workbook as _lw
_wb=_lw(out,data_only=True)
with open("/workspace/nebius/parts/comps_ev_bridge.csv","w",newline="") as fh:
    w=csv.writer(fh)
    for row in _wb["EV_Bridge"].iter_rows(values_only=True): w.writerow(row)
for v in VALS:
    print(v["company"][:22].ljust(23), "EVpf=%s EV/cMW=%s EV/aMW=%s EV/EBITDA ltm=%s ntm=%s EV/revQa=%s cod=%s" % tuple(
      (("%.1f"%x) if isinstance(x,float) else str(x)) for x in [v["ev_pf"],v["ev_cmw"],v["ev_amw"],v["ev_ebitda_ltm"],v["ev_ebitda_ntm"],v["ev_rev_qa"],v["cod"]]))
