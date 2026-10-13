from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
E='https://www.sec.gov/Archives/edgar/data/1513845/'
S={ # key: (form, date, accession, exhibit/doc, description)
'Q2PR':('6-K','2026-08-12','0001104659-26-094568','tm2622968d1_ex99-1.htm','Q2 2026 results press release (Ex.99.1)'),
'Q2LTR':('6-K','2026-08-12','0001104659-26-094568','tm2622968d1_ex99-2.htm','Q2 2026 shareholder letter (Ex.99.2)'),
'Q2FS':('6-K','2026-08-12','0001104659-26-094844','nbis-20260812xex99d2.htm','Q2 2026 unaudited interim financial statements (Ex.99.2)'),
'Q2MDA':('6-K','2026-08-12','0001104659-26-094844','nbis-20260812xex99d1.htm','Q2 2026 interim MD&A/OFR (Ex.99.1)'),
'Q1PR':('6-K','2026-05-13','0001104659-26-059872','tm2614392d1_ex99-1.htm','Q1 2026 results press release (Ex.99.1)'),
'Q1LTR':('6-K','2026-05-13','0001104659-26-059872','tm2614392d1_ex99-2.htm','Q1 2026 shareholder letter (Ex.99.2)'),
'Q1FS':('6-K','2026-05-20','0001104659-26-064092','nbis-20260331xex99d2.htm','Q1 2026 unaudited interim financial statements (Ex.99.2)'),
'Q4PR':('6-K','2026-02-12','0001104659-26-013946','tm266173d1_ex99-1.htm','Q4/FY2025 results press release (Ex.99.1)'),
'Q4LTR':('6-K','2026-02-12','0001104659-26-013946','tm266173d1_ex99-2.htm','Q4 2025 shareholder letter (Ex.99.2)'),
'20F25':('20-F','2026-04-30','0001104659-26-052948','nbis-20251231x20f.htm','Annual report FY2025'),
'20FA25':('20-F/A','2026-05-22','0001104659-26-065681','nbis-20251231x20fa.htm','Amendment to FY2025 20-F'),
'Q3PR25':('6-K','2025-11-12','0001104659-25-109806','tm2530882d1_ex99-1.htm','Q3 2025 results press release (Ex.99.1)'),
'Q3LTR25':('6-K','2025-11-12','0001104659-25-109806','tm2530882d1_ex99-2.htm','Q3 2025 shareholder letter (Ex.99.2)'),
'Q3FS25':('6-K','2025-11-13','0001104659-25-110831','tm2531034d1_ex99-1.htm','Q3 2025 interim OFR/financial info (Ex.99.1)'),
'Q2PR25':('6-K','2025-08-07','0001104659-25-075028','tm2522866d1_ex99-1.htm','Q2 2025 results press release (Ex.99.1)'),
'Q2LTR25':('6-K','2025-08-07','0001104659-25-075028','tm2522866d1_ex99-2.htm','Q2 2025 shareholder letter (Ex.99.2)'),
'H1FS25':('6-K','2025-09-10','0001104659-25-088860','tm2525580d2_ex99-1.htm','H1 2025 interim OFR/financial info (Ex.99.1)'),
'Q1PR25':('6-K','2025-05-20','0001104659-25-050975','tm2515580d1_ex99-1.htm','Q1 2025 results press release (Ex.99.1)'),
'Q1LTR25':('6-K','2025-05-20','0001104659-25-050975','tm2515580d1_ex99-2.htm','Q1 2025 shareholder letter (Ex.99.2)'),
'Q4PR24':('6-K','2025-02-20','0001104659-25-015480','tm257185d1_ex99-1.htm','Q4/FY2024 results press release (Ex.99.1)'),
'20F24':('20-F','2025-04-30','0001558370-25-005991','nbis-20241231x20f.htm','Annual report FY2024'),
'Q3PR24':('6-K','2024-10-31','0001104659-24-112875','tm2427209d1_ex99-1.htm','Q3 2024 results press release (Ex.99.1)'),
'PRES24':('6-K','2024-10-18','0001104659-24-109850','tm2426436d1_ex99-1.htm','Investor presentation Oct 2024 (Ex.99.1)'),
'Q2PR24':('6-K','2024-08-01','0001513845-24-000024','tmb-20240630xex99d1.htm','Q2 2024 results press release (Ex.99.1)'),
'JUN25':('6-K','2025-06-02','0001104659-25-055197','tm2516825d1_ex99-1.htm','June 2025 $1B convertible notes PR (Ex.99.1)'),
'SEP25P':('6-K','2025-09-12','0001410578-25-002149','tm2525503d8_ex99-2.htm','Sept 2025 convert pricing PR (Ex.99.2)'),
'SEP25C':('6-K','2025-09-15','0001104659-25-089969','tm2525580d4_ex99-1.htm','Sept 2025 convert/equity closing PR (Ex.99.1)'),
'MAR26P':('6-K','2026-03-18','0001104659-26-029863','tm268409d3_ex99-1.htm','Mar 2026 convert pricing PR (Ex.99.1)'),
'MAR26C':('6-K','2026-03-20','0001104659-26-032735','tm268409d4_ex99-1.htm','Mar 2026 convert closing PR (Ex.99.1); indentures Ex.4.1/4.3'),
'AUG26L':('6-K','2026-08-19','0001104659-26-098590','tm2623513d1_ex99-1.htm','Aug 2026 convert launch PR (Ex.99.1)'),
'AUG26P':('6-K','2026-08-20','0001104659-26-098924','tm2623617d1_ex99-1.htm','Aug 2026 convert pricing + exchange PR (Ex.99.1)'),
'AUG26C':('6-K','2026-08-24','0001104659-26-100347','tm2623863d1_ex99-1.htm','Aug 2026 convert closing PR (Ex.99.1); indentures Ex.4.1 (2030), Ex.4.3 (2034), exchange agmt Ex.4.5'),
'ABL':('6-K/A','2026-07-17','0001104659-26-084606','tm2620779d1_6ka.htm','Senior secured facility ($775M) 6-K/A + PR Ex.99.1'),
'MSFT':('6-K','2025-09-08','0001104659-25-088312','tm2525580d1_6k.htm','Microsoft commercial agreement 6-K'),
'META26':('6-K','2026-03-16','0001104659-26-027886','tm268879d1_ex99-1.htm','Meta $27B agreement PR (Ex.99.1)'),
'META25':('6-K','2025-11-12','0001104659-25-109803','tm2530882d2_6k.htm','Meta commercial agreement 6-K (Nov 2025)'),
'META26K':('6-K','2026-03-16','0001104659-26-027886','tm268879d1_6k.htm','Meta $27B agreement 6-K body'),
'NVDA':('6-K','2026-03-11','0001104659-26-026163','tm268532d1_ex99-1.htm','NVIDIA partnership / $2B investment PR (Ex.99.1)'),
'NVDA13G':('SC 13G','2026-07-20','0001045810-26-000062','primary_doc.xml','NVIDIA Schedule 13G on NBIS'),
'AGM':('6-K','2026-08-26','0001104659-26-101076','tm2624077d1_ex99-1.htm','AGM results (share counts at record date)'),
'PLTR':('6-K','2026-09-08','0001104659-26-105749','tm2624958d1_ex99-1.htm','Palantir partnership PR'),
'INFZ':('6-K','2026-10-01','0001104659-26-112824','tm2626792d1_ex99-1.htm','Inferize acquisition PR'),
'CALL':('Transcript','2026-08-12','n/a','fool.com','Q2 2026 earnings call transcript (Motley Fool, 3rd party)'),
'PP24':('6-K','2024-12-02','0001104659-24-124570','tm2429858d1_ex99-1.htm','Dec 2024 $700M private placement PR (Ex.99.1)'),
'SEP25G':('6-K','2025-09-22','0001104659-25-091939','tm2526625d1_ex99-1.htm','Sept 2025 equity offering greenshoe exercise PR (Ex.99.1)'),
'IND29':('6-K','2025-06-02','0001104659-25-055197','tm2516825d1_ex99-4.htm','Indenture - 2.00% notes due 2029 (Ex.99.4)'),
'IND31':('6-K','2025-06-02','0001104659-25-055197','tm2516825d1_ex99-5.htm','Indenture - 3.00% notes due 2031 (Ex.99.5)'),
'IND30S':('6-K','2025-09-15','0001104659-25-089969','tm2525580d4_ex4-1.htm','Indenture - 1.00% notes due Sep-2030 (Ex.4.1)'),
'IND32':('6-K','2025-09-15','0001104659-25-089969','tm2525580d4_ex4-3.htm','Indenture - 2.75% notes due Sep-2032 (Ex.4.3)'),
'IND31M':('6-K','2026-03-20','0001104659-26-032735','tm268409d4_ex4-1.htm','Indenture - 1.25% notes due Mar-2031 (Ex.4.1)'),
'IND33':('6-K','2026-03-20','0001104659-26-032735','tm268409d4_ex4-3.htm','Indenture - 2.625% notes due Mar-2033 (Ex.4.3)'),
'IND30F':('6-K','2026-08-24','0001104659-26-100347','tm2623863d1_ex4-1.htm','Indenture - 0.50% notes due Feb-2030 (Ex.4.1)'),
'IND34':('6-K','2026-08-24','0001104659-26-100347','tm2623863d1_ex4-3.htm','Indenture - 4.50% notes due Feb-2034 (Ex.4.3)'),
'EXAG':('6-K','2026-08-24','0001104659-26-100347','tm2623863d1_ex4-5.htm','Form of exchange agreement, 2029/2031 notes for shares (Ex.4.5)'),
'ABLPR':('6-K/A','2026-07-17','0001104659-26-084606','tm2620779d1_ex99-1.htm','Senior secured facility PR (Ex.99.1)'),
}
def url(k):
    f,d,a,doc,desc=S[k]
    if k=='CALL': return 'https://www.fool.com/earnings/call-transcripts/2026/08/19/nebius-nbis-q2-2026-earnings-call-transcript/'
    return E+a.replace('-','')+'/'+doc
def cite(k,ref=''):
    f,d,a,doc,desc=S[k]
    return f"{f} {d} acc {a} {doc}{(' '+ref) if ref else ''} | {url(k)}"
BLUE=Font(color='0000FF'); BOLD=Font(bold=True); HDR=PatternFill('solid',fgColor='DDDDDD')
wb=Workbook()
# ---------------- FactBase
fb=[ # (category, fact, value, unit, tag, source key, ref, note)
('Revenue','Group revenue Q2 2026',582.3,'$M','Reported','Q2PR','p.1 & Stmt of Ops','+454% YoY; +46% QoQ'),
('Revenue','Nebius AI cloud segment revenue Q2 2026',574.9,'$M','Reported','Q2FS','Note 15, p.F-33','~98% of group'),
('Revenue','Avride segment revenue Q2 2026',1.0,'$M','Reported','Q2FS','Note 15, p.F-33',''),
('Revenue','TripleTen segment revenue Q2 2026',10.0,'$M','Reported','Q2FS','Note 15, p.F-33',''),
('Revenue','Intersegment eliminations Q2 2026',-3.6,'$M','Reported','Q2FS','Note 15, p.F-33','Toloka deconsolidated Q2 2025 (equity method; discontinued ops)'),
('Revenue','Group revenue 1H 2026',981.3,'$M','Reported','Q2PR','p.1',''),
('ARR','ARR end of June 2026 (AI cloud June revenue x12)',3000,'$M','Reported','Q2LTR','p.9 (also p.3)','+598% YoY; $1.9B end-Mar'),
('Guidance','FY2026 ARR guidance (YE26)','7,000-9,000','$M','Reported (call)','CALL','CFO remarks','Reaffirmed 8/12/26; also Q1 2026 letter p.2'),
('Guidance','FY2026 group revenue guidance','3,000-3,400','$M','Reported (call)','CALL','CFO remarks','Q1 2026 letter p.2: "On track to achieve $3.0B-$3.4B revenue"'),
('Guidance','FY2026 group adj. EBITDA margin guidance','~40%','%','Reported (call)','CALL','CFO remarks','Q1 2026 letter p.2 "~40% Adj. EBITDA margin in 2026"'),
('Guidance','FY2026 capex guidance','20,000-25,000','$M','Reported (call)','CALL','CFO remarks','Not in SEC filings; call transcript only'),
('Guidance','Customer prepayments expected in 2026','>9,000','$M','Reported','Q2LTR','p.3, p.10',''),
('Profitability','Group adjusted EBITDA Q2 2026',236.2,'$M','Reported','Q2PR','Non-GAAP recon','41% margin; Q1 2026 129.5'),
('Profitability','Nebius AI cloud adjusted EBITDA Q2 2026',285.7,'$M','Reported','Q2FS','Note 15, p.F-33','49.7% margin'),
('Profitability','Avride adj. EBITDA Q2 2026',-40.1,'$M','Reported','Q2FS','Note 15',''),
('Profitability','TripleTen adj. EBITDA Q2 2026',-9.4,'$M','Reported','Q2FS','Note 15',''),
('Profitability','Net income/(loss) from continuing ops Q2 2026',-190.4,'$M','Reported','Q2PR','Stmt of Ops','EPS basic/diluted -0.68'),
('Profitability','D&A Q2 2026',259.7,'$M','Reported','Q2PR','Stmt of Ops',''),
('Profitability','Share-based compensation Q2 2026',102.5,'$M','Reported','Q2PR','p.1','incl. $115.9M non-recurring Eigen AI SBC/post-combination in product dev (letter p.10)'),
('Profitability','Interest expense Q2 2026 (net of capitalized)',119.1,'$M','Reported','Q2PR','Stmt of Ops','convert-related 67.3 per Note 12; $40.2M capitalized'),
('Capex','Purchases of PP&E and intangibles Q2 2026',5657.4,'$M','Reported','Q2PR','Cash flow','1H26: 8,130.3'),
('Cash flow','Operating cash flow Q2 2026 (continuing)',2246.1,'$M','Reported','Q2PR','Cash flow','driven by +1,197.0 deferred revenue and +1,186.8 A/R decrease'),
('Balance sheet','Cash & cash equivalents 30-Jun-2026',8042.1,'$M','Reported','Q2PR','Balance sheet','plus restricted cash 1,056.0 (total 9,098.1)'),
('Balance sheet','Debt, current 30-Jun-2026',46.7,'$M','Reported','Q2PR','Balance sheet',''),
('Balance sheet','Debt, non-current 30-Jun-2026 (convertible notes carrying value)',8499.0,'$M','Reported','Q2PR','Balance sheet','Note 12: accreted principal at maturity 10,041.8; fair value 20,823.7'),
('Balance sheet','Total debt 30-Jun-2026',8545.7,'$M','Derived','Q2PR','46.7+8,499.0',''),
('Balance sheet','PP&E net 30-Jun-2026',13045.2,'$M','Reported','Q2PR','Balance sheet','assets not yet in use 7,836.4 (Note 7)'),
('Balance sheet','Deferred revenue (current+non-current) 30-Jun-2026',5975.2,'$M','Reported','Q2FS','Note 1, p.F-19','979.4 current + 4,995.8 non-current'),
('Backlog','Remaining performance obligations 30-Jun-2026',37490.6,'$M','Reported','Q2FS','Note 1, p.F-19','36% within 24 months; 40% months 25-48; 100% AI cloud. YE25: 21,333.0; Q1 26: 33,585.3'),
('Customers','Customer concentration Q2 2026','C 24% / D 21% / E 14%','% rev','Reported','Q2FS','Note 1, p.F-17','Customers not named'),
('Shares','Shares issued & outstanding 30-Jun-2026 (total)',271855218,'shares','Reported','Q2PR','Outstanding Shares','Issued Class A + Class B only. Does NOT include the 21,065,936 NVIDIA pre-funded warrant shares (unexercised; no exercise reported through 10/1/26). Add warrant shares once for basic/fully diluted counts'),
('Shares','Class A outstanding 30-Jun-2026',238400165,'shares','Reported','Q2PR','Outstanding Shares','excl. 50,185,726 treasury'),
('Shares','Class B outstanding 30-Jun-2026',33455053,'shares','Reported','Q2PR','Outstanding Shares','10 votes/share'),
('Shares','Weighted avg basic shares Q2 2026',280381615,'shares','Reported','Q2PR','Stmt of Ops','includes NVIDIA pre-funded warrant shares (from issue date, per Q2FS Note 1) even though they are not in the 271,855,218 issued count'),
('Shares','Weighted avg diluted shares Q2 2026',280381615,'shares','Reported','Q2PR','Stmt of Ops','Equals basic (net loss); 65,416,804 weighted if-converted shares excluded as anti-dilutive (Q2FS Note 1). 1H26 diluted 302,618,666'),
('Shares','NVIDIA pre-funded warrant',21065936,'shares','Reported','Q2FS','Note 14, p.F-31','NOT included in the 271,855,218 issued & outstanding count (unexercised warrant, not issued shares). Included in basic and diluted weighted-average shares from issue date (Q2FS Note 1 & Note 14). Add once in basic/fully diluted counts; do not double count. $0.0001 strike; $2.0B gross; 13G 7/20/26: 22,256,412 sh beneficial (9.3%); exercise not reported through 10/1/26'),
('Shares','Share options outstanding 30-Jun-2026',6740600,'options','Reported','Q2FS','Note 13, p.F-29','WAEP $88.65'),
('Shares','Unvested RSUs 30-Jun-2026',6234091,'RSUs','Reported','Q2FS','Note 13, p.F-29','WA grant FV $58.22'),
('Shares','ATM sold Q2 2026',12729493,'shares','Reported','Q2FS','Note 14, p.F-31','WA $223.60; net $2,811.1M; 12,270,507 remain of 25M program'),
('Shares','Exchange shares issued Aug 2026 (approx.)',15800000,'shares','Reported','AUG26C','Ex.99.1','for $400M 2029 + $400M 2031 notes'),
('Debt','Senior secured facility (signed 10-Jul-2026)',775,'$M','Reported','ABL','6-K/A body','SOFR+2.50%; matures 31-Oct-2030; DSCR 1.15x; secured by Nebius Compute II assets'),
('Power','Contracted power target YE2026','5 GW','MW','Reported','Q2LTR','p.3, p.5','up from >4 GW (May-26); (5) contracted = land+power commitments'),
('Power','Contracted power secured (May 2026)','>3.5 GW','MW','Reported','Q1LTR','p.3, p.6','>75% owned'),
('Power','Connected power target YE2026','800 MW - 1 GW','MW','Reported','Q1LTR','p.6','reiterated on Q2 call'),
('Power','Active power 31-Dec-2025','~170 MW','MW','Reported','Q4LTR','Highlights','latest active-MW datapoint; Q1/Q2 2026 active MW not disclosed'),
('Power','Deployment pace from 2027','>1 GW per year','MW','Reported','Q2LTR','p.3, p.5',''),
('GPUs','GPU count (installed)','not disclosed','GPUs','Not disclosed','Q2LTR','','Not disclosed since FY24 20-F; last company actual = ~30,000 GPUs deployed at 31-Mar-2025, primarily H200 (20-F FY2024, acc 0001558370-25-005991). The ~38k from the Feb-25 call was a forward-looking target.'),
('GPUs','GPUs deployed at 31-Mar-2025 (last company actual)','30,000','GPUs','Reported','20F24','Item 4 (Business): "As of March 31, 2025, we have deployed approximately 30,000 GPUs across our global data center footprint, primarily NVIDIA H200s."','Approximate; primarily H200'),
('Pricing','Q2 2026 deal ACV per MW','$20-25M (avg >$20M)','$M/MW','Reported','Q2LTR','p.3-4','2026 base ~$12M; Q3 short-term deals >$40M; payback 1y10m'),
('Contracts','Microsoft: total contract value',17392.9,'$M','Reported','20F25','Note 1, p.F-17','up to ~$19.4B with additional capacity (6-K 9/8/25); 5-yr; 9 tranches 2025-26; upfront ~$6,958.1M; monthly through Oct-2031'),
('Contracts','Microsoft: delivery status','all tranches delivered','','Reported','Q2LTR','p.5','tranche 1 Nov-25, tranche 2 Feb-26 (20-F)'),
('Contracts','Meta 2025 agreement (1-Nov-2025) TCV',2880.7,'$M','Reported','20F25','Note 1, p.F-18','6-K 11/12/25 (acc 0001104659-25-109803): ~$2.9B, 2 clusters, 5-yr, tranches Dec-25 & Feb-26; delivered (Item 10 p.76 says first tranche Nov-25); termination on late tranche after grace period; site not disclosed'),
('Contracts','Meta 2026 agreement (13-Mar-2026) dedicated capacity',12000,'$M','Reported','META26','Ex.99.1','5-yr; NVIDIA Vera Rubin; deliveries from early 2027 / tranches through 1H27 (20-F p.75)'),
('Contracts','Meta 2026 agreement backstop (unsold capacity)',15000,'$M','Reported','META26','Ex.99.1','up to; Meta buys unsold capacity; total up to ~$27B'),
('Contracts','Other Q2 2026 deals','4 deals avg TCV >$1B (Reflection, Cohere, US neolab, quant firm)','','Reported','Q2LTR','p.4',''),
('Contracts','Palantir partnership (8-Sep-2026)','preferred sovereign AI infra partner; no $ disclosed','','Reported','PLTR','Ex.99.1',''),
('M&A','Inferize acquisition (1-Oct-2026)','terms not disclosed','','Reported','INFZ','Ex.99.1',''),
('D&A policy','Server & network equipment useful life','5 years from 1-Jan-2026 (was 4)','','Reported','Q2FS','Note 1, p.F-10','Q2 26 depreciation -$43.0M; 1H26 -$86.1M; FY26 est. -$167.6M (20-F p.56)'),
]
ws=wb.active; ws.title='FactBase'
ws.append(['Category','Fact','Value','Unit','Tag','Form','Filing date','Accession','Document','Page/section','Note','URL'])
for c in ws[1]: c.font=BOLD; c.fill=HDR
for r in fb:
    k=r[5]; f,d,a,doc,desc=S[k]
    ws.append([r[0],r[1],r[2],r[3],r[4],f,d,a,doc,r[6],r[7],url(k)])
    ws.cell(ws.max_row,3).font=BLUE
# ---------------- Converts
cv=wb.create_sheet('Converts')
cv.append(['Instrument','Issue/close date','Orig principal issued ($M)','Coupon','Maturity','Initial conv. rate (sh per $1k)','Initial conv. price ($)','Accretion at maturity','Eff. conv. price at maturity ($)','Capped call','Converted/exchanged ($M orig)','Remaining orig principal after Aug-26 exchange ($M)','Note 12 "principal amount outstanding" 6/30/26 ($M, accreted basis)','Carrying 6/30/26 ($M)','Fair value 6/30/26 ($M)','Earliest issuer redemption','Tag','Sources'])
for c in cv[1]: c.font=BOLD; c.fill=HDR
rows=[
['2.00% Conv Sr Notes due Jun-2029','2025-06-05',500,'2.00%','2029-06-05',19.4363,51.45,'120%',61.74,'none disclosed','12.5 converted 2025 (242,953 sh); 400 exchanged Aug-26',87.5,587.5,504.1,2928.7,'on/after 2026-12-10 (130% trigger)','Reported; remaining = Derived (500-12.5-400)','JUN25; 20F25 p.F-47/48; Q2FS Note 12; AUG26C'],
['3.00% Conv Sr Notes due Jun-2031','2025-06-05',500,'3.00%','2031-06-05',19.4363,51.45,'125%',64.31,'none disclosed','12.5 converted 2025 (242,953 sh); 400 exchanged Aug-26',87.5,612.5,498.4,2908.4,'on/after 2027-06-10 (130% trigger) per indenture Sec. heading (20-F text implies Dec-2026 for both)','Reported; remaining = Derived','JUN25; 20F25; Q2FS Note 12; AUG26C'],
['1.00% Conv Notes due Sep-2030','2025-09-15','~1,581 (PR: ~$1.58B; 1,818.4/1.15 = 1,581.2 Derived)','1.00%','2030-09-15',7.2072,138.75,'115%','~159.6 (Derived 138.75x1.15)','none disclosed','none disclosed','~1,581','1818.4',1588.2,3555.5,'on/after 2028-09-20 (130% trigger)','Reported/Derived','SEP25P; SEP25C; Q2FS Note 12'],
['2.75% Conv Notes due Sep-2032','2025-09-15','~1,581','2.75%','2032-09-15',7.2072,138.75,'115%','~159.6 (Derived)','none disclosed','none disclosed','~1,581','1818.4',1575.0,3488.0,'on/after 2028-09-20 (130% trigger)','Reported/Derived','SEP25P; SEP25C; Q2FS Note 12'],
['1.25% Conv Sr Notes due Mar-2031','2026-03-20',2587.5,'1.25%','2031-03 (Mar-2031)',5.4579,183.22,'120%',219.86,'none disclosed','none',2587.5,3105.0,2588.6,4696.5,'on/after 2029-03-20','Reported','MAR26P; MAR26C; Q2FS Note 12'],
['2.625% Conv Sr Notes due Mar-2033','2026-03-20',1750,'2.625%','2033-03 (Mar-2033)',5.5460,180.31,'120%',216.37,'none disclosed','none (greenshoe $262.5M not exercised)',1750,2100.0,1744.7,3246.6,'on/after 2030-03-20','Reported','MAR26P; MAR26C; Q2FS Note 12'],
['0.50% Conv Sr Notes due Feb-2030','2026-08-24',3450,'0.50%','2030-02-15',3.1902,313.46,'110%',344.81,'none disclosed','n/a (new)',3450,'n/a (post-period)','n/a','n/a','2028-02-21 (130% trigger)','Reported','AUG26P; AUG26C (Ex.4.1 indenture)'],
['4.50% Conv Sr Notes due Feb-2034','2026-08-24',2300,'4.50%','2034-02-15',3.0802,324.65,'125%',405.82,'none disclosed','n/a (new)',2300,'n/a (post-period)','n/a','n/a','2028-08-21 (130%; 150% until 2029-08-21)','Reported','AUG26P; AUG26C (Ex.4.3 indenture)'],
['TOTAL converts (orig principal) post Aug-26','', '~14,225 issued','', '', '', '', '', '', '', '825 retired', '~13,425 (Derived)','10,041.8 at 6/30','8,499.0','20,823.7','','Derived','sum of rows'],
['Senior secured term loan (Nebius Compute II)','signed 2026-07-10',775,'SOFR+2.50%','2030-10-31','n/a','n/a','n/a','n/a','n/a','n/a',775,'n/a','n/a','n/a','prepayable w/o premium','Reported','ABL'],
]
for r in rows:
    cv.append(r)
    for j in (3,4,5,6,7,12): cv.cell(cv.max_row,j).font=BLUE
cv.append([]); cv.append(['Notes: No capped call / call spread is disclosed in any convert press release, 6-K, 20-F or interim FS reviewed (text search for "capped call" returned zero hits). Settlement: indentures set Default Settlement Method = Physical Settlement (shares); company may elect Cash, Physical or Combination (cash up to a Specified Dollar Amount per $1k + shares) and may irrevocably fix/eliminate methods; Sec. 3.10 (Jun-25 onward excl.) deems Physical Settlement where it avoids Dutch withholding. See Indenture_Terms sheet. Interest: 2029/2031 semi-annual Jun 5/Dec 5; 2030S/2032 Mar 15/Sep 15; Mar-26 notes Mar 15/Sep 15; Aug-26 notes Feb 15/Aug 15 from Feb 15 2027.'])
cv.append(['Potential conversion shares at initial rates on remaining orig principal (Derived): 2029 87.5x19.4363=1.70M; 2031 1.70M; Sep-30 1,581.25x7.2072=11.40M; Sep-32 11.40M; Mar-31 2,587.5x5.4579=14.12M; Mar-33 1,750x5.546=9.71M; Feb-30 3,450x3.1902=11.01M; Feb-34 2,300x3.0802=7.08M; total ~68.1M shares (before make-whole adjustments).'])
# ---------------- Quarterly
Q=['Q1 2024','Q2 2024','Q3 2024','Q4 2024','Q1 2025','Q2 2025','Q3 2025','Q4 2025','Q1 2026','Q2 2026']
ND='not disclosed'
def R(v,k,ref='',tag='Reported',note=''):
    if tag not in ('Reported','Derived','Estimated'): note=tag+(('; '+note) if note else ''); tag='Reported'  # positional-arg fix: 4th arg was used as a note
    return (v,tag,k,ref,note)
def D(v,k,ref,note): return (v,'Derived',k,ref,note)
def N(note='',k=None): return (ND,'Not disclosed',k,'',note)
rows={}
rows['Revenue - group (continuing ops, latest restated basis, $M)']=[D(9.7,'Q2PR25','H1 24.2 - Q2 14.5','restated ex-Toloka; orig. reported incl. Toloka 11.4'),R(14.5,'Q2PR25','comparative'),R(32.1,'Q3PR25','comparative','orig. 43.3 incl. Toloka'),R(35.2,'Q4PR','comparative','orig. 37.9 incl. Toloka'),R(50.9,'Q1PR','comparative','orig. 55.3 incl. Toloka'),R(105.1,'Q2PR25'),R(146.1,'Q3PR25'),R(227.7,'Q4PR'),R(399.0,'Q1PR'),R(582.3,'Q2PR')]
rows['Revenue - as originally reported (incl. Toloka through Q1-25, $M)']=[D(11.4,'Q2PR24','H1 36.3 - Q2 24.9','matches Q1 2025 PR comparative'),R(24.9,'Q2PR24'),R(43.3,'Q3PR24'),R(37.9,'Q4PR24'),R(55.3,'Q1PR25'),R(105.1,'Q2PR25'),R(146.1,'Q3PR25'),R(227.7,'Q4PR'),R(399.0,'Q1PR'),R(582.3,'Q2PR')]
rows['Segment revenue - Nebius AI cloud ($M)']=[N('H1-24 total 17.4 (H1FS25); quarterly split not disclosed','H1FS25'),N('H1-24 total 17.4','H1FS25'),D(27.1,'Q3FS25','9M 44.5 - H1 17.4; also reported Q3 comparative','reported 27.1'),D(23.8,'20F25','FY24 68.3 - 9M 44.5',''),R(41.4,'Q1FS','Note segments'),R(93.7,'Q2FS','Note 15'),R(131.1,'Q3FS25'),D(214.1,'20F25','FY25 480.3 - 9M 266.2',''),R(389.7,'Q1FS'),R(574.9,'Q2FS','Note 15')]
rows['Segment revenue - Avride ($M)']=[N('H1-24 total 0.1','H1FS25'),N('H1-24 total 0.1','H1FS25'),R(0.1,'Q3FS25'),D(0.1,'20F25','FY 0.3 - 9M 0.2',''),R(0.2,'Q1FS'),R(0.2,'Q2FS'),R(0.2,'Q3FS25'),D(0.7,'20F25','FY 1.3 - 9M 0.6',''),R(0.9,'Q1FS'),R(1.0,'Q2FS')]
rows['Segment revenue - TripleTen ($M)']=[N('H1-24 total 10.2','H1FS25'),N('H1-24 total 10.2','H1FS25'),R(6.3,'Q3FS25'),D(12.3,'20F25','FY 28.8 - 9M 16.5',''),R(10.5,'Q1FS'),R(12.3,'Q2FS'),R(16.2,'Q3FS25'),D(15.1,'20F25','FY 54.1 - 9M 39.0',''),R(11.6,'Q1FS'),R(10.0,'Q2FS')]
rows['Segment revenue - Eliminations ($M)']=[N('H1-24 total -3.5','H1FS25'),N('H1-24 total -3.5','H1FS25'),R(-1.4,'Q3FS25'),D(-1.0,'20F25','FY -5.9 - 9M -4.9',''),R(-1.2,'Q1FS'),R(-1.1,'Q2FS'),R(-1.4,'Q3FS25'),D(-2.2,'20F25','FY -5.9 - 9M -3.7',''),R(-3.2,'Q1FS'),R(-3.6,'Q2FS')]
rows['Segment revenue - Toloka']=[(ND+' (discontinued ops; reclassified)','Not disclosed','20F25','',"Toloka deconsolidated Q2 2025; prior periods in discontinued ops")]*10
rows['ARR - Nebius AI cloud, as stated ($M)']=[R(32,'PRES24','ARR chart (1Q24)','investor presentation'),R(45,'PRES24','ARR chart (2Q24)','Q2-24 PR: ">$80M as of July 2024"'),R(121,'PRES24','ARR chart (3Q24)','Q3-24 PR: "surpassing $120M as of September 2024"'),R(90,'Q4PR24','highlights','"December 2024 ARR ... $90 million, below previous guidance"'),R(249,'Q1LTR25'),R(430,'Q2PR25','highlights'),R(551,'Q3LTR25','"approximately $551 million"'),R(1250,'Q4LTR'),R(1920,'Q1LTR','p.9 ($1.92B)'),R(3000,'Q2LTR','p.9')]
rows['ARR guidance outstanding at quarter end (as stated)']=[N(),N(),R('$170-190M (Dec-24)','PRES24'),R('$750M-1B (Dec-25)','Q4PR24'),R('$750M-1B (Dec-25)','Q1LTR25'),R('$900M-1.1B (Dec-25)','Q2PR25'),R('$900M-1.1B (Dec-25)','Q3PR25'),R('$7-9B (Dec-26)','Q4LTR'),R('$7-9B (Dec-26)','Q1LTR','p.2'),R('$7-9B (Dec-26)','CALL')]
rows['Adjusted EBITDA - group ($M)']=[D(-58.4,'Q2PR25','H1 -116.5 - Q2 -58.1','restated basis'),R(-58.1,'Q2PR25'),R(-45.9,'Q3PR25'),R(-63.9,'Q4PR'),R(-53.7,'Q1PR'),R(-21.0,'Q2PR25'),R(-5.2,'Q3PR25'),R(15.0,'Q4PR'),R(129.5,'Q1PR'),R(236.2,'Q2PR')]
rows['Adjusted EBITDA - Nebius AI cloud segment ($M)']=[N('H1-24 total -68.5','H1FS25'),N('H1-24 total -68.5','H1FS25'),R(-21.5,'Q3FS25'),D(-38.5,'20F25','FY -128.5 - 9M -90.0',''),R(-27.4,'Q1FS'),R(9.5,'Q2FS'),R(25.1,'Q3FS25'),D(51.8,'20F25','FY 59.0 - 9M 7.2',''),R(174.0,'Q1FS'),R(285.7,'Q2FS')]
rows['Adjusted EBITDA - Avride ($M)']=[N('H1-24 -34.4','H1FS25'),N('H1-24 -34.4','H1FS25'),R(-15.4,'Q3FS25'),D(-17.2,'20F25','FY -67.0 - 9M -49.8',''),R(-16.9,'Q1FS'),R(-17.3,'Q2FS'),R(-20.8,'Q3FS25'),D(-27.7,'20F25','FY -82.7 - 9M -55.0',''),R(-34.1,'Q1FS'),R(-40.1,'Q2FS')]
rows['Adjusted EBITDA - TripleTen ($M)']=[N('H1-24 -13.6','H1FS25'),N('H1-24 -13.6','H1FS25'),R(-9.0,'Q3FS25'),D(-8.2,'20F25','FY -30.8 - 9M -22.6',''),R(-9.4,'Q1FS'),R(-13.2,'Q2FS'),R(-9.5,'Q3FS25'),D(-9.1,'20F25','FY -41.2 - 9M -32.1',''),R(-10.4,'Q1FS'),R(-9.4,'Q2FS')]
rows['Net income/(loss) - continuing ops ($M)']=[D(-68.6,'Q2PR25','H1 -185.5 - Q2 -116.9',''),R(-116.9,'Q2PR25'),R(-43.6,'Q3PR25'),R(-122.9,'Q4PR'),R(-104.3,'Q1PR'),R(502.5,'Q2PR25','','incl. $597.4M equity revaluation gain'),R(-119.6,'Q3PR25'),R(-249.6,'Q4PR'),R(621.2,'Q1PR','','incl. $780.6M equity revaluation gain'),R(-190.4,'Q2PR')]
rows['Net income/(loss) - total ($M)']=[D(-316.5,'Q2PR25','H1 -414.0 - Q2 -97.5','matches Q1-25 PR comparative'),R(-97.5,'Q2PR25'),R(-94.2,'Q3PR25'),R(-133.2,'Q4PR'),R(-113.5,'Q1PR'),R(584.4,'Q2PR25'),R(-119.6,'Q3PR25'),R(-249.6,'Q4PR'),R(621.2,'Q1PR'),R(-190.4,'Q2PR')]
rows['Depreciation & amortization ($M)']=[D(8.9,'Q2PR25','H1 20.3 - Q2 11.4','Q1-25 PR comparative also 8.9'),R(11.4,'Q2PR25'),R(23.5,'Q3PR25'),R(33.3,'Q4PR'),R(49.1,'Q1PR','','Q1-25 PR originally 49.2'),R(75.2,'Q2PR25'),R(99.0,'Q3PR25'),R(180.7,'Q4PR'),R(212.0,'Q1PR'),R(259.7,'Q2PR','','4->5yr life change cut Q2 depreciation by $43.0M')]
rows['Capex - purchases of PP&E and intangibles ($M)']=[D(58.9,'Q2PR25','H1 217.9 - Q2 159.0',''),R(159.0,'Q2PR25'),R(172.1,'Q3PR25'),R(417.5,'Q4PR'),R(543.9,'Q1PR'),R(510.6,'Q2PR25'),R(955.5,'Q3PR25'),R(2056.0,'Q4PR'),R(2472.9,'Q1PR'),R(5657.4,'Q2PR')]
rows['PP&E, net (end of period, $M)']=[N('pre-divestiture (Yandex N.V.) balance sheet not comparable'),R(328.3,'Q2PR24','BS','as reported Aug-24 (incl. Toloka)'),R(475.3,'Q3PR24','BS','as reported'),R(846.7,'Q4PR','BS comparative','orig. 849.3 (Q4-24 PR)'),R(1334.1,'Q1PR25','BS','as reported (incl. Toloka)'),R(1789.4,'Q2PR25'),R(3314.4,'Q3PR25'),R(5553.3,'Q2PR','BS comparative','Q4-25 PR originally 5,572.5'),R(7131.7,'Q1PR'),R(13045.2,'Q2PR')]
rows['Cash & cash equivalents (end of period, $M)']=[N('pre-divestiture balance sheet not comparable'),R(2325.5,'Q2PR24','BS','as reported (incl. Toloka)'),R(2288.2,'Q3PR24','BS','continuing ops 2,277.8 per Q3-25 PR'),R(2434.7,'Q4PR','BS comparative','orig. 2,449.6 incl. Toloka'),R(1447.0,'Q1PR25','BS','continuing ops 1,439.7 per Q1-26 PR'),R(1679.3,'Q2PR25'),R(4794.8,'Q3PR25'),R(3678.1,'Q4PR'),R(9298.2,'Q1PR'),R(8042.1,'Q2PR','','restricted cash +1,056.0')]
rows['Total debt (current + non-current, carrying, $M)']=[N('pre-divestiture'),R(6.5,'Q2PR24','BS current portion','non-current not shown'),R(6.2,'Q3PR24','BS current portion',''),R(6.1,'Q4PR','BS',''),R(6.2,'Q1PR25','BS current portion',''),D(986.2,'Q2PR25','8.0 + 978.2','June-25 converts'),D(4106.8,'Q3PR25','16.0 + 4,090.8',''),D(4127.7,'Q4PR','24.5 + 4,103.2',''),D(8450.4,'Q1PR','18.4 + 8,432.0',''),D(8545.7,'Q2PR','46.7 + 8,499.0','')]
rows['Shares outstanding - total (period end)']=[N('Yandex N.V. pre-divestiture'),R(293850159,'Q2PR24','','before final divestiture closing/treasury'),R(199340118,'Q3PR24'),R(235753600,'Q4PR24'),R(238108831,'Q1PR25'),R(238705092,'Q2PR25'),R(251807222,'Q3PR25'),R(253016971,'Q4PR'),R(253898194,'Q1PR'),R(271855218,'Q2PR')]
rows['Shares outstanding - Class A']=[N(),R(258151485,'Q2PR24'),R(163641444,'Q3PR24'),R(200054926,'Q4PR24'),R(202410157,'Q1PR25'),R(203006418,'Q2PR25'),R(218158548,'Q3PR25'),R(219465088,'Q4PR'),R(220406311,'Q1PR'),R(238400165,'Q2PR')]
rows['Shares outstanding - Class B']=[N(),R(35698674,'Q2PR24'),R(35698674,'Q3PR24'),R(35698674,'Q4PR24'),R(35698674,'Q1PR25'),R(35698674,'Q2PR25'),R(33648674,'Q3PR25'),R(33551883,'Q4PR'),R(33491883,'Q1PR'),R(33455053,'Q2PR')]
def EST(v,k,ref,note): return (v,'Estimated',k,ref,note)
rows['Weighted avg shares - basic']=[N('not disclosed quarterly (Yandex N.V. pre-divestiture; no EPS table in PR). FY2024 basic = 281,005,226 (20-F FY24)','20F24'),N('not disclosed quarterly; FY2024 basic = 281,005,226; period-end Q2-24 count 293,850,159 is pre-Second Completion','20F24'),EST(246595139,'Q3PR24','avg(293,850,159 Q2-end; 199,340,118 Q3-end)','ESTIMATE only: weighted count not disclosed; 94.9M treasury shares received 12-Jul-2024 so true average is lower than simple avg'),EST(217546859,'Q4PR24','avg(199,340,118; 235,753,600)','ESTIMATE only: not disclosed; 33.3M placement shares issued Dec-2024 so true average is closer to Q3-end count'),R(237916047,'Q1PR','comparative'),R(238520244,'Q2PR','comparative'),EST(245256157,'Q3PR25','avg(238,705,092; 251,807,222)','ESTIMATE only: not disclosed in Q3-25 PR/OFR. Cross-check: FY25 basic 242,531,291 x4 - Q1 - Q2 = 493,688,873 implied H2 sum vs estimate sum 497,668,254 (+0.8%)'),EST(252412097,'Q4PR','avg(251,807,222; 253,016,971)','ESTIMATE only: not disclosed; see Q3-25 cross-check'),R(258298911,'Q1PR'),R(280381615,'Q2PR','','includes 21,065,936 NVIDIA pre-funded warrant shares; excludes 2,743,655 unvested Eigen AI shares')]

rows['Weighted avg shares - diluted']=[N('not disclosed quarterly; FY2024 diluted = 281,005,226 (net loss)','20F24'),N('not disclosed quarterly; FY2024 diluted = 281,005,226','20F24'),EST(246595139,'Q3PR24','= basic estimate (net loss quarter -> anti-dilutive)','ESTIMATE only'),EST(217546859,'Q4PR24','= basic estimate (net loss)','ESTIMATE only'),R(237916047,'Q1PR','comparative'),R(247266034,'Q2PR','comparative'),EST(245256157,'Q3PR25','= basic estimate (net loss from continuing ops)','ESTIMATE only; FY25 diluted 247,679,946 per 20-F'),EST(252412097,'Q4PR','= basic estimate (net loss)','ESTIMATE only'),R(308971701,'Q1PR'),R(280381615,'Q2PR','','net loss -> diluted = basic; 65,416,804 weighted if-converted shares excluded as anti-dilutive (Q2FS Note 1)')]
rows['Deferred revenue, total (end of period, $M)']=[N(),N(),N(),N(),N(),N(),N(),R(1577.5,'Q2FS','Note 1'),R(4778.1,'Q1FS','Note 1'),R(5975.2,'Q2FS','Note 1')]
rows['RPO (end of period, $M)']=[N(),N(),N(),N(),N(),N(),N(),R(21333.0,'20F25','Note (revenue)'),R(33585.3,'Q1FS','Note 1'),R(37490.6,'Q2FS','Note 1')]
rows['Active power (MW)']=[N(),N(),N(),N(),N(),N(),N(),R('~170','Q4LTR','highlights','vs 100MW target'),N('Q1-26 letter defines active power but gives no figure','Q1LTR'),N('','Q2LTR')]
rows['Connected power (MW) / target']=[N(),N(),N(),N(),N(),R('target 220 MW by YE25','Q2LTR25'),N(),N('target 800MW-1GW by YE26','Q4LTR'),N('target 800MW-1GW by YE26','Q1LTR'),N('target 800MW-1GW by YE26 (call)','CALL')]
rows['Contracted power (as stated)']=[N(),N(),R('Finland expansion to 75 MW (target)','Q3PR24'),R('Kansas City 5MW initial (exp. to 40MW)','Q4PR24'),N(),R('NJ 200 MW contracted in construction; Finland ~75MW','Q2LTR25'),R('>1 GW (Aug-25 per chart); target >2.5 GW YE26','Q1LTR','p.2 chart'),R('>2 GW secured; target >3 GW','Q4LTR'),R('>3.5 GW (>75% owned); target >4 GW','Q1LTR','p.3, p.6'),R('target 5 GW YE26','Q2LTR','p.3, p.5')]
rows['GPU count (as stated)']=[N(),N(),R('Finland site capacity up to 60,000 GPUs (potential)','Q3PR24'),R('intends to deploy >22,000 NVIDIA Blackwell GPUs (US)','Q4PR24'),N(),N(),N(),N(),N(),N()]
rows['Headcount']=[N(),N(),R('500+ employees (~400 engineers) - Nebius core','PRES24'),N(),N(),N(),N(),R('~1,500 employees (31-Dec-25)','20F25'),N(),N()]
rows['Managed customers (as stated)']=[N(),R('>30 managed clients','Q2PR24'),R('>40 managed clients','Q3PR24'),N(),N(),N(),N(),N(),N(),N()]
rows['Customer concentration (>10% of revenue, quarter)']=[N(),N(),N(),N(),N(),R('A 39%, B 15%','Q2FS','Note 1'),N(),R('FY25: A 25%, B 15%','20F25'),N(),R('C 24%, D 21%, E 14%','Q2FS','Note 1')]
qs=wb.create_sheet('Quarterly'); src=wb.create_sheet('Quarterly_Sources')
qs.append(['Metric']+Q); src.append(['Cell','Metric','Quarter','Value','Tag','Source (form date acc doc ref | URL)','Note/derivation'])
for c in qs[1]: c.font=BOLD; c.fill=HDR
for c in src[1]: c.font=BOLD; c.fill=HDR
for m,vals in rows.items():
    qs.append([m]+[v[0] for v in vals]); r=qs.max_row
    for j,v in enumerate(vals):
        cell=qs.cell(r,j+2); cell.font=BLUE if v[1] in('Reported','Derived') else (Font(color='C65911',italic=True) if v[1]=='Estimated' else Font(color='808080',italic=True))
        cl=f"{get_column_letter(j+2)}{r}"
        s=cite(v[2],v[3]) if v[2] else ''
        note=(v[4]+(' | derivation: '+v[3] if v[1] in ('Derived','Estimated') else '')).strip(' |')
        src.append([cl,m,Q[j],v[0],v[1],s,note])
qs.append([]); qs.append(['Tag legend: blue = Reported/Derived input; grey italic = not disclosed; orange italic = Estimated (method in Quarterly_Sources). Every cell is sourced in Quarterly_Sources (keyed by cell). 2024 quarters on restated continuing-ops basis (Toloka moved to discontinued ops in Q2 2025) unless row says "as reported".'])
qs.column_dimensions['A'].width=58
for j in range(2,12): qs.column_dimensions[get_column_letter(j)].width=16
# ---------------- MW_Disclosures
mw=wb.create_sheet('MW_Disclosures')
mw.append(['Date','Metric type','Value','Exact quote','Document','Page/section','URL'])
for c in mw[1]: c.font=BOLD; c.fill=HDR
MW=[
('2024-10-31','Planned capacity / GPUs','75 MW; up to 60,000 GPUs','"Nebius announced a major expansion of its data center in Finland to triple capacity to 75 MW by the end of 2025 or early 2026. This will enable Nebius to accommodate up to 60,000 GPUs(3) at this location alone, with annual revenue potential of over USD 1 billion at full capacity utilization"','Q3PR24','Business highlights'),
('2025-02-20','GPU plan / site MW','>22,000 Blackwell GPUs; Kansas City 5MW (exp. 40MW)','"In December 2024, Nebius announced its intention to deploy over 22,000 NVIDIA Blackwell GPUs in data centers in the U..." ; "NVIDIA GPU cluster in Kansas City (5MW initial phase), set to go live at the end of Q1 2025 ... The facility is expandable to 40MW"','Q4PR24','Highlights'),
('2025-08-07','Connected power target / contracted','220 MW connected YE25; NJ 200 MW; Finland ~75 MW','"We expect to have 220 MW of connected power, which includes both active power and power that is fully provisioned in our data centers..." ; "We now have 200 MW of contracted power that is in active construction in New Jersey, of which 100 MW will be made available this year, with the rest following in 2026"','Q2LTR25','Group Business Update (p.3-4)'),
('2026-02-12','Active power','~170 MW at YE25','"We ended the year with ~170 MW of active power, significantly exceeding our 100 MW target."','Q4LTR','Letter / capacity section'),
('2026-02-12','Contracted / connected power','>2 GW secured; >3 GW target; 800MW-1GW connected YE26','"We have already secured more than 2 GW of contracted power ... we now expect to have more than 3 GW of contracted power by year-end." ; "We also remain well on track to reach 800 MW to 1 GW of connected power by the end of 2026."','Q4LTR','Highlights'),
('2026-05-13','Definitions','contracted / connected / active','"(1) contracted power, or power secured by contracted land and power commitments; (2) connected power, or power connected into fully built and equipped data centers; and (3) active power, or power being consumed by installed, operational IT equipment and available for revenue generation."','Q1LTR','p.6'),
('2026-05-13','Contracted power','>3.5 GW; >75% owned; target >4 GW','"Our contracted capacity now exceeds 3.5 GW, with owned capacity representing more than 75% of the total. Given this momentum, we are raising our contracted power guidance to more than 4 GW by year-end"','Q1LTR','p.6 (also p.3)'),
('2026-05-13','Owned sites','Missouri 1.2 GW; Pennsylvania up to 1.2 GW; Finland 310 MW; 3 GW owned across 5 sites','"new second owned giga-watt-scale site in the United States — in Pennsylvania, with up to 1.2 GW of power ... alongside the 1.2 GW AI factory in Independence, Missouri" ; "new owned location in Finland that when fully deployed at 310 MW" ; "our owned facilities will deliver 3 GW of capacity across five sites"','Q1LTR','p.6'),
('2026-05-13','Connected power target','800MW-1GW YE26','"We continue to expect 800MW to 1GW of connected power by year-end."','Q1LTR','p.6'),
('2026-08-12','Contracted power target','5 GW YE26','"We now anticipate ending 2026 with 5 GW of contracted power(5), up from the +4 GW we indicated last quarter." (5) "Contracted power, secured by contracted land and power commitments."','Q2LTR','p.5 (also p.3)'),
('2026-08-12','Deployment pace','>1 GW/yr from 2027','"The pace at which we bring capacity to market will accelerate as we plan to deploy more than 1 GW per year, starting in 2027."','Q2LTR','p.5'),
('2026-08-12','Pricing per MW','$20-25M ACV/MW; >$40M short-term; $12M 2026 base','"won deals averaging above $20 million in revenue per MW" ; "annual contract value of $20-25 million per MW" ; chart: "2026 base $12M ... Q2 deals >$20M ... Q3 short-term capacity deals >$40M"','Q2LTR','p.2 chart, p.3-4'),
('2026-08-12','Microsoft delivery','all tranches delivered','"To date, we\'ve delivered all capacity tranches to Microsoft under the contract, including two tranches over the past month."','Q2LTR','p.5'),
('2026-08-12','Meta 2027 capacity','second Meta deal capacity early 2027','"We also continued to progress the buildout supporting our second Meta agreement, with capacity on track to come online in early 2027."','Q2LTR','p.5'),
('2026-04-30','Microsoft contract - capacity terms','9 tranches; no GPU/MW count','"The Microsoft GPU Services are being deployed in nine tranches during 2025 and 2026 ... estimated to be up to $17,392.9 ... including aggregate upfront payments of approximately $6,958.1 ... through October 2031."','20F25','Note 1, p.F-17'),
('2026-04-30','Meta 2026 contract - capacity terms','$12B dedicated, tranches 1H27; $15B backstop; no GPU/MW count','"Certain of the Orders are for dedicated GPU capacity clusters across multiple locations ... with deployments in tranches throughout the first half of 2027 ... These Orders have a total contract value of $12 billion."','20F25','Item 10 Material Contracts, p.75'),
('2025-11-12','Meta #1 contract terms','2 clusters, 2 tranches Dec-25/Feb-26, ~$2.9B, 5 yrs','"the Company will provide Meta access to two dedicated GPU infrastructure capacity clusters (each, a GPU Service) over a five-year term. The GPU Services will be deployed in two tranches during December 2025 and February 2026 ... The Order has a total contract value of approximately $2.9 billion."','META25','6-K body'),
('2026-03-16','Meta #2 contract terms','$12B dedicated (tranches from early 2027) + up to $15B backstop','"Certain of the Orders are for dedicated GPU capacity clusters across multiple locations, and for a duration of 5 years, with deployments in tranches starting early 2027 ... These Orders have a total contract value of $12 billion."','META26K','6-K body'),
('n/a','GPU count / MW tied to Microsoft or Meta contract','not disclosed','No filing reviewed ties Microsoft or Meta contract value to a GPU count or MW figure. Microsoft capacity is at Vineland NJ (200 MW NJ contracted per Q2-25 letter) but the contract share of that MW is not disclosed.','20F25',''),
]
for r in MW: mw.append([r[0],r[1],r[2],r[3],cite(r[4]).split(' | ')[0],r[5],url(r[4])])
# ---------------- D&A_Policy
da=wb.create_sheet('D&A_Policy')
da.append(['Item','Value / quote','Document','Page/section','URL'])
for c in da[1]: c.font=BOLD; c.fill=HDR
DA=[
('Useful lives table (FY2025)','Server and network equipment 4.0 years; Infrastructure systems and equipment 3.0-10.0 years; Office furniture and equipment 3.0 years; Buildings 20.0 years; Other property and equipment 2.0-10.0 years; Land not depreciated','20F25','Note 2 PP&E policy, p.F-22'),
('Useful lives table (FY2024)','Server and network equipment 4.0 years; Infrastructure systems and equipment 3.0-10.0 years; Office furniture and equipment 3.0 years; Buildings 20.0 years; Other property and equipment 2.0-5.0 years (FY2025 widened Other to 2.0-10.0)','20F24','Note 2 PP&E policy, p.F-19'),
('GPU useful life statement (FY2024 20-F)','"Our main assets, GPUs, are currently depreciated over four years." No change in estimate recorded in FY2024; no rationale for life beyond usage given.','20F24','Item 5 Critical accounting estimates, p.48'),
('Power gear useful life','Not separately disclosed (power/electrical likely within "Infrastructure systems and equipment 3.0-10.0 years")','20F25','p.F-22'),
('GPU useful life statement','"Our main assets, GPUs, are currently depreciated over four years."','20F25','Item 5 Critical accounting estimates, p.56'),
('Change in estimate (announced)','"In January 2026, the Company completed an assessment of the useful lives of servers and network equipment ... should be extended from four to five years ... we expect this change in accounting estimate will reduce the depreciation expenses for fiscal year 2026 by approximately $167.6 million."','20F25','p.56; also Note 2 p.F-15'),
('Change in estimate (applied) - impact','"Management applied this change in accounting estimate prospectively beginning January 1, 2026 ... for the six months ended June 30, 2026 included a reduction in depreciation expense of $86.1 and an increase in net income ... of $75.7 ... for the three months ended June 30, 2026 ... reduction in depreciation expense of $43.0 and a decrease in net loss ... of $34.1."','Q2FS','Note 1, p.F-10'),
('Shareholder letter on life change','"We use a five-year useful life for our server and network equipment based on usage patterns and current utilization commitments, up from the four-year useful life used prior to 2026."','Q2LTR','p.10'),
('Impairment of PP&E','"No impairment of property and equipment was identified or recorded during the three and six months ended June 30, 2025 and 2026."','Q2FS','Note 7, p.F-24'),
('Capitalized interest','Capitalization of interest started Q3 2025; $28.0M capitalized in FY2025 (20-F p.F-38); $40.2M Q2-26 and $59.6M 1H-26 (Q2FS Note 12)','Q2FS','Note 12, p.F-28'),
('PP&E gross by class 31-Dec-2024','Server & network 794.9; Land/bldgs 48.5; Infrastructure 63.1; Other 12.7; Not yet in use 163.8; Total 1,083.0; Acc. dep. (236.3); Net 846.7','20F25','Note 7, p.F-38'),
('PP&E gross by class 31-Dec-2025','Server & network 3,123.6; Land/bldgs 381.4; Infrastructure 199.1; Other 68.8; Not yet in use 2,417.4; Total 6,190.3; Acc. dep. (637.0); Net 5,553.3','20F25','Note 7, p.F-38'),
('PP&E gross by class 30-Jun-2026','Server & network 5,197.7; Land/bldgs 593.4; Infrastructure 332.5; Other 189.5; Not yet in use 7,836.4; Total 14,149.5; Acc. dep. (1,104.3); Net 13,045.2','Q2FS','Note 7, p.F-23'),
('Gross additions by class (Derived)','FY2025: server & network +2,328.7; land/bldgs +332.9; infra +136.0; other +56.1; not-yet-in-use +2,253.6 (net change in gross, ignores disposals/transfers). 1H2026: server & network +2,074.1; land/bldgs +212.0; infra +133.4; other +120.7; not-yet-in-use +5,419.0','20F25 / Q2FS','Derived from tables above'),
('Depreciation of PP&E','FY2023 27.9; FY2024 75.2; FY2025 411.0 (20-F p.F-38); Q2-25 72.6; Q2-26 253.8; 1H-25 121.2; 1H-26 462.6 (Q2FS p.F-24)','Q2FS','Note 7'),
('Formal rollforward (additions/disposals by vintage)','not disclosed - only gross balances by class; no GPU vintage schedule','Q2FS',''),
]
for r in DA:
    k=r[2] if r[2] in S else 'Q2FS'
    da.append([r[0],r[1],cite(k).split(' | ')[0] if r[2] in S else r[2],r[3],url(k)])
exec(open('/workspace/nebius/filings/tabs2.py').read())
exec(open('/workspace/nebius/filings/tabs3.py').read())
exec(open('/workspace/nebius/filings/tabs4.py').read())
exec(open('/workspace/nebius/filings/tabs5.py').read())
# ---------------- Sources
so=wb.create_sheet('Sources')
so.append(['Key','Form','Filing date','Accession','Document','Description','URL'])
for c in so[1]: c.font=BOLD; c.fill=HDR
for k,v in S.items(): so.append([k,*v,url(k)])
for sh in wb.worksheets:
    for col in range(1,sh.max_column+1):
        if sh.title!='Quarterly': sh.column_dimensions[get_column_letter(col)].width=min(60,max(12,sh.column_dimensions[get_column_letter(col)].width or 12))
wb.calculation.fullCalcOnLoad=True
wb.save('/workspace/nebius/filings/nebius_filings.xlsx')
import json; json.dump({'fb':fb},open('/workspace/nebius/filings/fb.json','w'))
json.dump({'rows':rows,'S':S,'ND':ND},open('/workspace/nebius/filings/quarterly_rows.json','w'),default=str)
print('saved', len(rows),'quarterly rows')
