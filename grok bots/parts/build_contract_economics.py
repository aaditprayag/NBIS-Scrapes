#!/usr/bin/env python3
"""Build contract_economics.xlsx + chart. Nebius GPU Market & Comps, 2026-10-04 ET."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

H = 8760.0
GPM = {'GB300_IT': 491, 'GB300_GROSS': 409, 'VR_IT': 320, 'VR_GROSS': 267}
CAPFILE = '/workspace/nebius/capacity/gpu_mw_capex.xlsx (GPUsPerMW tab, base case)'
R = 0.08  # discount rate for prepay PV adjustment (Estimated)
rm = R / 12
def AF(n): return (1 - (1 + rm) ** -n) / rm

# ---------------- sources ----------------
S = dict(
 NB_20F='https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm',
 NB_MS6K='https://www.sec.gov/Archives/edgar/data/1513845/000110465925088312/tm2525580d1_6k.htm',
 NB_META1='https://www.sec.gov/Archives/edgar/data/1513845/000110465925109803/tm2530882d2_6k.htm',
 NB_META2='https://www.sec.gov/Archives/edgar/data/1513845/000110465926027886/tm268879d1_6k.htm',
 NB_META2PR='https://www.sec.gov/Archives/edgar/data/1513845/000110465926027886/tm268879d1_ex99-1.htm',
 NB_Q2='https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm',
 DCD_100K='https://www.datacenterdynamics.com/en/news/microsoft-to-get-100000-nvidia-gb300s-under-nebius-deal-has-signed-more-than-33bn-in-capacity-agreements-with-neoclouds/',
 DCD_300MW='https://www.datacenterdynamics.com/en/news/microsoft-to-use-nebius-gpu-data-centers-in-deal-worth-174bn-over-five-years/',
 KYI='https://knowyourinvestments.substack.com/p/thoughts-on-the-new-nebius-contract',
 IREN_8K='https://www.sec.gov/Archives/edgar/data/1878848/000114036125040072/ef20058139_8k.htm',
 IREN_PR='https://www.sec.gov/Archives/edgar/data/1878848/000114036125040072/ef20058139_ex99-1.htm',
 IREN_10K='https://www.sec.gov/Archives/edgar/data/1878848/000187884826000052/iren-20260630.htm',
 IREN_FY26='https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm',
 NS_S1='Nscale Ltd Form S-1 filed 2026-09-18 (local copy /workspace/nebius/comps_raw/nscale_s1.htm); EDGAR filer CIK 0002110365',
 NS_PR='https://www.nscale.tech/press-releases/nscale-microsoft-2025',
 AKER='https://akernscale.com/news/microsoft-signs-multi-billion-dollar-ai-infrastructure-agreement-with-aker-nscale',
 LAMBDA='https://lambda.ai/blog/lambda-announces-multibillion-dollar-agreement-with-microsoft-to-deploy-ai-infrastructure-powered-by-tens-of-thousands-of-nvidia-gpus',
 CW_META1='https://www.sec.gov/Archives/edgar/data/1769628/000176962825000050/crwv-20250925.htm',
 CW_META1_PRESS='https://www.reuters.com/technology/coreweave-signs-14-billion-ai-deal-with-meta-bloomberg-news-reports-2025-09-30/',
 CW_META2='https://www.sec.gov/Archives/edgar/data/1769628/000176962826000154/crwv-20260331.htm',
 CW_META2PR='https://www.coreweave.com/news/coreweave-and-meta-announce-21-billion-expanded-ai-infrastructure-agreement',
 CW_10Q='https://www.sec.gov/Archives/edgar/data/1769628/000176962826000366/crwv-20260630.htm',
 CW_OAI1='https://www.coreweave.com/news/coreweave-announces-agreement-with-openai-to-deliver-ai-infrastructure',
 CW_OAI1X='https://www.sec.gov/Archives/edgar/data/1769628/000119312525052207/d899798dex1025.htm',
 CW_OAI2='https://www.reuters.com/business/coreweave-signs-new-4-bln-deal-with-openai-filing-shows-2025-05-15/',
 CW_OAI3='https://www.sec.gov/Archives/edgar/data/1769628/000119312525216497/d17274d8k.htm',
 ANT_BBG='https://www.bloomberg.com/news/articles/2026-08-26/anthropic-to-pay-nscale-45-billion-for-ai-computing-power',
 ANT_TCD='https://www.techceodaily.com/big-tech/anthropic-nscale-45-billion-compute-deal',
 ORCL_WSJ='https://www.wsj.com/business/openai-oracle-sign-300-billion-computing-deal-among-biggest-in-history-ff27c8fe',
 ORCL_DCD='https://www.datacenterdynamics.com/en/news/openai-confirmed-to-be-behind-30bn-a-year-oracle-cloud-deal-45gw-expected-across-multiple-data-center-sites/',
 CCIR='https://ccir.io/research/implied-committed-rates',
 SAVRN='https://savrn.com/ai-index/pricing/gpus/vera-rubin',
 WULF='https://www.sec.gov/Archives/edgar/data/1083301/000110465925078084/tm2523008d2_ex99-1.htm',
 CIVO='https://www.civo.com/ai/vera-rubin-gpu',
 NB_PRICES='https://nebius.com/prices',
 CW_PRICES='https://www.coreweave.com/pricing',
)

# ---------------- prepayment PV adjustments (Estimated) ----------------
N = 60
nb_P = 6958.1; nb_TCV = 17392.9
nb_gain = nb_P - nb_P / N * AF(N)                  # upfront received t0, credited ratably over 60m (assumption)
nb_equiv = nb_TCV + nb_gain * N / AF(N)
nb_adj = nb_equiv / nb_TCV
ir_TCV = 9700.0; ir_P = 0.2 * ir_TCV
ir_pv_credit = ir_P / 36 * (AF(60) - AF(24))       # credited pro rata after month 24 (8-K)
ir_gain = ir_P - ir_pv_credit
ir_equiv = ir_TCV + ir_gain * N / AF(N)
ir_adj = ir_equiv / ir_TCV

def gpuhr(tcv, gpus, yrs): return tcv * 1e6 / (gpus * H * yrs)

# ---------------- DEALS (raw terms) ----------------
# columns
DC = ['ID','Deal / tranche','Provider','Customer','Signed / announced','TCV ($M)','TCV tag','Upside / optional TCV ($M)','Term (yrs)','Term detail',
      'MW','MW basis (IT / gross / unlabeled)','MW tag','GPU type','GPU count','GPU tag','Start / ramp','Prepayment','Option / extension',
      'Take-or-pay / termination','Scope (who pays power & facility)','Primary source (page/section)','Secondary source','Confidence','Notes']
deals = [
 ['N1','Microsoft–Nebius (Vineland NJ), committed','Nebius','Microsoft','2025-09-08 (eff. 2025-09-07)',17392.9,'Reported',19400,5,'5-yr term per tranche; 9 tranches deployed 2025-26; monthly invoicing through Oct-2031',
  300,'Unlabeled site capacity ("up to 300 MW"); Capacity team reads it as gross/facility (~230-250 MW IT)','Reported (site MW); basis Estimated',
  'NVIDIA GB300 NVL72 (press)', 100000,'Press (Bloomberg via DCD): ">100,000" GB300 - a floor, attributed to the $19.4B deal',
  'T1 Nov-2025, T2 Feb-2026; Q2-26 letter: all tranches due to date delivered (Capacity team: >=4 of 9)',
  '~$6,958.1M aggregate upfront payments (40% of TCV) - Reported','Microsoft may buy additional services/capacity -> ~$19.4B TCV',
  'Fees committed irrespective of utilization (take-or-pay); per-tranche termination if delivery missed after grace and no alt capacity; LDs/service credits',
  'Full-stack dedicated GPU capacity; Nebius pays power (Bloom fuel cells/gensets) and DataOne build-to-suit lease',
  S['NB_20F']+' (Note 1, p.F-17: $17,392.9M, $6,958.1M upfront, 9 tranches, through Oct-2031); '+S['NB_MS6K']+' (6-K: 5-yr, ~$17.4B/~$19.4B)',
  S['DCD_100K']+' (GPU count); '+S['DCD_300MW']+' (300 MW site)','High (TCV/prepay); Medium (GPUs); Low-Med (MW basis)',
  'No GPU count, MW or pricing in any Nebius filing.'],
 ['N1o','Microsoft–Nebius incl. optional capacity','Nebius','Microsoft','2025-09-08',19400,'Reported (up to)',None,5,'as N1',
  300,'as N1','as N1','GB300 (press)',100000,'Press, floor; extra $2.0B likely buys extra GPUs (not disclosed)','as N1','as N1','-','as N1','as N1',S['NB_MS6K'],S['DCD_100K'],'Medium','Upside case only'],
 ['N2','Meta–Nebius #1 (Nov-2025 order)','Nebius','Meta','2025-11-01 (6-K 2025-11-12)',2880.7,'Reported',None,5,'5-yr term; 2 clusters/tranches',
  None,'not disclosed','-','not disclosed (secondary source says B300; unverified)',None,'not disclosed',
  'Tranches Dec-2025 & Feb-2026 (delivered)','not disclosed','Meta may extend term and/or add capacity',
  'Fees irrespective of utilization; Meta may terminate order if T1 late after grace; T2-only termination thereafter; discounted fees for late delivery',
  'Full-stack dedicated clusters + storage/connectivity; site not disclosed',
  S['NB_META1']+' (6-K Item: ~$2.9B, 5-yr, 2 tranches); '+S['NB_20F']+' (Note 1, p.F-18: $2,880.7M)',S['KYI']+' (unsourced "~50 MW, B300")','High (TCV); Low (MW/GPU)',
  'Nebius letter: contract size was limited to available capacity.'],
 ['N3','Meta–Nebius #2 (Mar-2026), dedicated Vera Rubin','Nebius','Meta','2026-03-13 (6-K 2026-03-16)',12000,'Reported',None,5,'Series of 5-yr orders; tranches through 1H27',
  None,'not disclosed','-','NVIDIA Vera Rubin NVL72','not disclosed','not disclosed','Deliveries from early 2027; tranches throughout 1H27 (20-F p.75); Q2 letter: on track',
  'not disclosed','Up to $15B backstop order (row N4)','Per-tranche SLA/termination for delivery delays (20-F)','Full-stack dedicated clusters, multiple locations (not named)',
  S['NB_META2']+' (6-K: $12B dedicated, 5-yr); '+S['NB_META2PR']+' (PR: Vera Rubin)',S['KYI']+' (estimate ~135 MW for $12B; Estimated by author)','High (TCV); n.d. (MW/GPU)',''],
 ['N4','Meta–Nebius #2 backstop (unsold capacity)','Nebius','Meta','2026-03-13',15000,'Reported (up to, contingent)',None,5,'5 yrs from deployment of each unsold cluster',
  None,'not disclosed','-','not disclosed (certain upcoming clusters)',None,'not disclosed','Upcoming clusters','not disclosed','-',
  'Meta buys only capacity Nebius fails to sell to third parties','Full-stack',S['NB_META2'],'','High (terms); contingent value','Contingent; excluded from normalized comps'],
 ['P1','Microsoft–IREN (Childress TX, Horizon 1-4)','IREN','Microsoft','2025-11-02 (8-K 2025-11-03)',9700,'Reported',None,5,'"five-year average term"; through 2031',
  200,'IT (critical IT load, "combined IT load of approximately 200MW")','Reported',
  'NVIDIA GB300 (NVL72 per 10-K)',None,'not disclosed (Estimated in Normalized)',
  'H1 accepted by Microsoft Aug-2026; H2-H4 targeted Q4-2026, grace to start of Q2-2027 (10-K)',
  '20% of each tranche paid before delivery, credited pro rata against fees after month 24 (8-K)','-',
  'Customary termination incl. missed delivery after cure; delay credits','Full-stack: IREN owns DC and buys power (ERCOT; fixed-price hedges May-2026); $5.8B Dell GPU+ancillary capex',
  S['IREN_8K']+' (Item 1.01); '+S['IREN_PR'],S['IREN_10K']+' (Risk factors: H1 Aug-2026); '+S['IREN_FY26']+' (four 50MW (IT) deployments)','High',''],
 ['P2','Microsoft–Nscale, all SOWs (Sep-2025 to Apr-2026)','Nscale','Microsoft','2025-09 to 2026-04',43800,'Reported (up to)',None,5.5,'GPU service terms 5-6 yrs from acceptance of each tranche; through Dec-2033; +6-month option',
  None,'Site IT MW in S-1 table not mapped to customer (see Notes)','-',
  '~191,600 GB300 (Texas 104k, Narvik 52k, Loughton 23k, Sines 12.6k) + >66,000 Vera Rubin NVL72 (Sines SIN02)',257600,'Reported (press releases/S-1); sum Derived',
  'Sines GB300 from Q1-26; Texas from Q3-26; Loughton Q1-27; Narvik staged from 2026; Sines VR from late 2027',
  'not disclosed per contract (S-1: "meaningful proportion of TCV provided as upfront payments" historically)','Optional 6-month extension; Microsoft option on 700 MW Texas phase 2 (not in TCV)',
  'Payments subject to delivery/availability; termination for cause','Full-stack GPU infra + colocation + cloud services; mix of owned, leased and colo sites',
  S['NS_S1']+' (Business - Customers and Business Model, p.114)',S['NS_PR']+' (site GPU counts); '+S['AKER'],'Medium (TCV high; GPU/MW mapping medium)',
  'Narvik JV rolled up into Nscale (S-1 "JV Roll-Up"), so Narvik is assumed inside the $43.8B.'],
 ['P2a','  of which Microsoft–Aker/Nscale Narvik','Aker-Nscale JV (now Nscale)','Microsoft','2025-09-17',6200,'Reported',None,5,'5-yr binding agreement; staged from 2026',
  None,'not disclosed (S-1 lists two 100 IT MW owned Norway sites; mapping Estimated)','-','NVIDIA GB300',52000,'Reported (Nscale PR 2025-10-15)',
  'Staged deployments beginning 2026','not disclosed','-','"standard commercial terms and safeguards"','Full-stack; cheap Norwegian hydropower',
  S['AKER'],S['NS_PR']+' (52,000 GB300); '+S['CCIR'],'High (TCV, GPUs)','Cleanest public GB300 print (CCIR grade A).'],
 ['P3','Microsoft–Lambda','Lambda','Microsoft','2025-11-03','not disclosed','"multibillion-dollar" only',None,'not disclosed','"multi-year"',
  None,'not disclosed','-','GB300 NVL72 among "tens of thousands" of NVIDIA GPUs','not disclosed','"tens of thousands"','not disclosed','not disclosed','not disclosed','not disclosed','not disclosed',
  S['LAMBDA'],'https://www.datacenterdynamics.com/en/news/microsoft-signs-multi-billion-dollar-cloud-capacity-deal-with-lambda/','Low (no economics public)','Lambda has no S-1 on file (as of 2026-10-04).'],
 ['P4','Meta–CoreWeave #1 (Sep-2025 order)','CoreWeave','Meta','2025-09-25 (8-K 2025-09-30)',14200,'Reported (up to)',None,round((2031+349/365)-(2025+268/365),2),'Through 2031-12-14 (span from signing; ramp shortens billed term)',
  None,'not disclosed','-','GB300 (press)','not disclosed','not disclosed','not disclosed','not disclosed','Option to materially expand through 2032 (exercised in P5)',
  'Subject to delivery/availability; termination for cause','Full-stack reserved capacity',S['CW_META1']+' (Item 1.01)',S['CW_META1_PRESS'],'High (TCV); n.d. (MW/GPU)',''],
 ['P5','Meta–CoreWeave #2 (Mar-2026 order)','CoreWeave','Meta','2026-03-31 (announced 2026-04-09)',21000,'Reported (up to)',None,round((2032+354/365)-(2026+89/365),2),'New capacity through 2032-12-20; includes exercise of prior option through 2032-04-10',
  None,'not disclosed','-','Multiple; includes some initial Vera Rubin deployments','not disclosed','not disclosed','not disclosed','not disclosed','Incl. exercised option',
  'Subject to delivery/availability','Full-stack, multiple locations',S['CW_10Q']+' (Risk factors: customer concentration); '+S['CW_META2'],S['CW_META2PR'],'High (TCV); n.d. (MW/GPU)',''],
 ['P6','OpenAI–CoreWeave tranche 1 (Mar-2025)','CoreWeave','OpenAI','2025-03-10',11900,'Reported (up to)',None,5,'5-yr',
  None,'not disclosed','-','not disclosed','not disclosed','not disclosed','not disclosed','not disclosed; OpenAI received $350M of CRWV stock','-','not disclosed','Full-stack',S['CW_OAI1'],S['CW_OAI1X']+' (Ex.10.25)','High (TCV)','$350M equity to OpenAI is a price concession not in TCV'],
 ['P7','OpenAI–CoreWeave tranche 2 (May-2025)','CoreWeave','OpenAI','2025-05-15',4000,'Reported (up to)',None,3.9,'Through Apr-2029',
  None,'not disclosed','-','not disclosed','not disclosed','not disclosed','not disclosed','not disclosed','-','not disclosed','Full-stack',S['CW_OAI2'],'','Medium-High',''],
 ['P8','OpenAI–CoreWeave tranche 3 (Sep-2025)','CoreWeave','OpenAI','2025-09-23 (8-K 2025-09-25)',6500,'Reported (up to)',None,round((2031+151/365)-(2025+266/365),2),'Through 2031-05-31',
  None,'not disclosed','-','not disclosed','not disclosed','not disclosed','not disclosed','not disclosed','-','Subject to delivery/availability; termination for cause','Full-stack',S['CW_OAI3']+' (Item 1.01)','','High (TCV)','OpenAI total ~$22.4B'],
 ['O1','Anthropic–Nscale (Monarch WV), Vera Rubin','Nscale','Anthropic','2026-08-25',44600,'Reported (up to)',None,6,'"six years" (Bloomberg); S-1: multi-year terms per tranche, extension options',
  460,'IT (S-1 site #17: US owned, 460 IT MW contracted; Bloomberg ~460 MW)','Reported (S-1 table) / press mapping',
  'NVIDIA Vera Rubin NVL72',None,'not disclosed (redacted)','Late 2027 start; 4 tranches','not disclosed','Extension options at Anthropic election',
  'Anthropic may terminate late tranches without liability; uptime-based termination','Full-stack; Nscale-owned campus, on-site gas generation',
  S['NS_S1']+' (p.114 Customers; Properties table p.120)',S['ANT_BBG']+'; '+S['ANT_TCD'],'Medium',
  'Financing not yet committed (S-1).'],
 ['O2','OpenAI–Oracle (Stargate)','Oracle','OpenAI','2025-07/09 (reported)',300000,'Press (WSJ); not confirmed by Oracle filings by name',None,5,'~5 yrs starting ~2027 (press)',
  4500,'unlabeled ("4.5 GW")','Press / OpenAI capacity statements','Mixed (GB200 etc.)','not disclosed','-','~2027','not disclosed','-','not disclosed','Full-stack OCI',
  S['ORCL_WSJ'],S['ORCL_DCD'],'Low-Medium','Optional peer row'],
 ['O3','Anthropic–SpaceX/xAI (Colossus)','SpaceX (xAI)','Anthropic','2026 (424B4 2026-06-12)','$1.25B/month through May-2029','Reported (prospectus, via CCIR)',None,'~2.9 (Jun-26 to May-29)','Either party may terminate on 90 days notice after initial 3 months',
  None,'not disclosed','-','unspecified NVIDIA',325000,'Reported (prospectus, via CCIR)','Ramped May-Jun 2026 at reduced fee','not disclosed','-','Terminable on 90 days notice','Full-stack',
  'SpaceX Form 424B4, acc. 0001628280-26-042639 (as quoted by CCIR)',S['CCIR'],'Medium-Low','Short-tenor in substance; chip unknown'],
 ['REF','TeraWulf–Fluidstack colocation (Google backstop) - facility-only reference','TeraWulf','Fluidstack (Google backstop $1.8B)','2025-08-14',3700,'Reported',8700,10,'10-yr initial + two 5-yr options',
  200,'IT (">200 MW of critical IT load (~250 MW gross)")','Reported','n/a (tenant brings GPUs)',None,'-','40 MW 1H26; 200+ MW by YE26','not disclosed','Two 5-yr extensions -> ~$8.7B',
  'Lease; Google backstop','Facility/shell only (modified gross lease); no GPUs',S['WULF'],'','High','Used only to size the facility component of full-stack pricing'],
]

# ---------------- NORMALIZED ----------------
NC = ['ID','Deal','Chip','TCV ($M)','Term (yrs)','TCV / yr ($M)','MW used','MW basis','MW tag / method','$M per MW-yr (as stated MW)','$ per MW-hr (as stated MW)',
      'IT MW (normalized)','IT-MW method','$M per IT MW-yr','$ per IT MW-hr','GPUs used','GPU tag / method','Implied $ per GPU-hr','Prepay PV factor (8%)',
      'ADJ $M per IT MW-yr (prepay-adjusted)','ADJ $ per GPU-hr (prepay-adjusted)','Discount vs CoreWeave GB200 list $10.50','Discount vs Nebius B300 OD $9.50',
      'Chip-mismatch flag','Confidence','Math shown']
norm = []
def add(**k): norm.append(k)

# N1
tcvyr = 17392.9/5
it_mw = 100000/GPM['GB300_IT']
add(ID='N1',Deal='Microsoft–Nebius (committed)',Chip='GB300',TCV=17392.9,T=5,Y=tcvyr,MW=300,MWb='Unlabeled (likely gross)',MWt='Reported site MW',
    pmw=tcvyr/300,pmh=tcvyr/300*1e6/H,IT=it_mw,ITm='Derived: 100,000 GB300 / 491 GPUs per IT MW',pit=tcvyr/it_mw,pith=tcvyr/it_mw*1e6/H,
    G=100000,Gt='Press (">100k", floor)',gh=gpuhr(17392.9,100000,5),pf=nb_adj,apit=tcvyr*nb_adj/it_mw,agh=gpuhr(17392.9,100000,5)*nb_adj,
    flag='Same chip/buyer/tenor as P1, P2a',conf='Medium',
    math=f'17,392.9/5 = {tcvyr:,.1f}/yr; /300 MW = ${tcvyr/300:.2f}M/MW-yr = ${tcvyr/300*1e6/H:,.0f}/MW-hr (team figure verified, 300-MW basis); 100,000/491 = {it_mw:.1f} IT MW -> ${tcvyr/it_mw:.2f}M/IT MW-yr; $/GPU-hr = 17,392.9e6/(100,000 x 8,760 x 5) = {gpuhr(17392.9,100000,5):.2f}; prepay factor {nb_adj:.3f} (upfront $6,958.1M at t0 credited ratably over 60m, 8%)')
# N1o
add(ID='N1o',Deal='Microsoft–Nebius incl. option ($19.4B)',Chip='GB300',TCV=19400,T=5,Y=19400/5,MW=300,MWb='Unlabeled (likely gross)',MWt='Reported site MW',
    pmw=19400/5/300,pmh=19400/5/300*1e6/H,IT=it_mw,ITm='as N1 (upper bound: extra $2B may add GPUs)',pit=19400/5/it_mw,pith=19400/5/it_mw*1e6/H,
    G=100000,Gt='Press (floor)',gh=gpuhr(19400,100000,5),pf=None,apit=None,agh=None,flag='Upside; GPU count probably > 100k',conf='Low-Medium',
    math=f'19,400/5/300 = {19400/5/300:.2f}; 19,400e6/(100,000 x 8,760 x 5) = {gpuhr(19400,100000,5):.2f}')
# N2 Meta 1
m1y = 2880.7/5
add(ID='N2',Deal='Meta–Nebius #1',Chip='not disclosed',TCV=2880.7,T=5,Y=m1y,MW=None,MWb='not disclosed',MWt='Estimated scenario only: ~48-50 MW',
    pmw=None,pmh=None,IT=50,ITm='Estimated (LOW): ~50 MW per unsourced secondary (KYI); = 48 MW if priced at Nebius "2026 base ~$12M ACV/MW" (Q2-26 letter p.3)',
    pit=m1y/50,pith=m1y/50*1e6/H,G=None,Gt='not disclosed',gh=None,pf=None,apit=None,agh=None,flag='Chip unknown',conf='Low',
    math=f'2,880.7/5 = {m1y:.1f}/yr; / 50 MW (Estimated) = ${m1y/50:.2f}M/MW-yr; $/GPU-hr not computed (chip & count not disclosed). If GB300 at 50 IT MW (24,550 GPUs): ${gpuhr(2880.7,50*491,5):.2f} (illustrative only)')
# N3 Meta 2
m2y = 12000/5
add(ID='N3',Deal='Meta–Nebius #2 (Vera Rubin, firm)',Chip='Vera Rubin',TCV=12000,T=5,Y=m2y,MW=None,MWb='not disclosed',MWt='not disclosed',
    pmw=None,pmh=None,IT=None,ITm=f'not disclosed. Break-even IT MW to match Anthropic–Nscale VR $16.2M/IT MW-yr = {m2y/ (44600/6/460):.0f} MW; at Nebius 2026 base $12M = {m2y/12:.0f} MW; at Q2-26 deals $20M = {m2y/20:.0f} MW',
    pit=None,pith=None,G=None,Gt='not disclosed',gh=None,pf=None,apit=None,agh=None,flag='Rubin vs GB300 - compare per MW only',conf='n.d.',
    math=f'12,000/5 = {m2y:,.0f}/yr. At VR 320 GPUs/IT MW, each 100 IT MW = 32,000 GPUs; if 148 MW -> ${gpuhr(12000,148*320,5):.2f}/GPU-hr; if 200 MW -> ${gpuhr(12000,200*320,5):.2f}; if 120 MW -> ${gpuhr(12000,120*320,5):.2f}')
# P1 IREN
iy=9700/5; ig=200*GPM['GB300_IT']
add(ID='P1',Deal='Microsoft–IREN',Chip='GB300',TCV=9700,T=5,Y=iy,MW=200,MWb='IT (critical)',MWt='Reported',
    pmw=iy/200,pmh=iy/200*1e6/H,IT=200,ITm='Reported',pit=iy/200,pith=iy/200*1e6/H,G=ig,Gt=f'Estimated: 200 IT MW x 491 = {ig:,}; alt. 360-410/MW (Capacity team, from $5.8B Dell capex) -> 72-82k GPUs -> ${gpuhr(9700,82000,5):.2f}-{gpuhr(9700,72000,5):.2f}',
    gh=gpuhr(9700,ig,5),pf=ir_adj,apit=iy*ir_adj/200,agh=gpuhr(9700,ig,5)*ir_adj,flag='Same chip/buyer/tenor as N1',conf='High (per MW); Medium (per GPU)',
    math=f'9,700/5 = {iy:,.0f}/yr; /200 IT MW = ${iy/200:.2f}M/IT MW-yr = ${iy/200*1e6/H:,.0f}/MW-hr; 9,700e6/({ig:,} x 8,760 x 5) = {gpuhr(9700,ig,5):.2f}; prepay factor {ir_adj:.3f} (20% paid pre-delivery, credited pro rata months 25-60, 8%)')
# P2 Nscale agg
ns_it = 191600/GPM['GB300_IT'] + 66000/GPM['VR_IT']; nsy=43800/5.5
add(ID='P2',Deal='Microsoft–Nscale (all SOWs)',Chip='GB300 + Vera Rubin mix',TCV=43800,T=5.5,Y=nsy,MW=None,MWb='n/a',MWt='-',pmw=None,pmh=None,
    IT=ns_it,ITm=f'Derived: 191,600/491 + 66,000/320 = {ns_it:.0f} IT MW (S-1 site-table mapping alt. ~680 IT MW -> ${nsy/680:.1f}M)',pit=nsy/ns_it,pith=nsy/ns_it*1e6/H,
    G=257600,Gt='Reported counts (sum)',gh=gpuhr(43800,257600,5.5),pf=None,apit=None,agh=None,flag='Blended GB300/VR - $/GPU-hr not same-chip',conf='Medium-Low',
    math=f'43,800/5.5 = {nsy:,.0f}/yr; / {ns_it:.0f} IT MW = ${nsy/ns_it:.2f}M; 43,800e6/(257,600 x 8,760 x 5.5) = {gpuhr(43800,257600,5.5):.2f}')
# P2a Narvik
nvy=6200/5; nv_it=52000/GPM['GB300_IT']
add(ID='P2a',Deal='Microsoft–Nscale/Aker Narvik',Chip='GB300',TCV=6200,T=5,Y=nvy,MW=None,MWb='n/a',MWt='-',pmw=None,pmh=None,IT=nv_it,ITm='Derived: 52,000/491',
    pit=nvy/nv_it,pith=nvy/nv_it*1e6/H,G=52000,Gt='Reported',gh=gpuhr(6200,52000,5),pf=None,apit=nvy/nv_it,agh=gpuhr(6200,52000,5),flag='Same chip/buyer/tenor as N1',conf='High (per GPU); Medium (per MW)',
    math=f'6,200/5 = {nvy:,.0f}/yr; 52,000/491 = {nv_it:.1f} IT MW -> ${nvy/nv_it:.2f}M; 6,200e6/(52,000 x 8,760 x 5) = {gpuhr(6200,52000,5):.2f}; prepay not disclosed -> no adj.')
# P2b Nscale ex Narvik
rx_it = (191600-52000)/GPM['GB300_IT'] + 66000/GPM['VR_IT']; rxy=(43800-6200)/5.5
add(ID='P2b',Deal='Microsoft–Nscale ex-Narvik (residual)',Chip='GB300 + Vera Rubin mix',TCV=43800-6200,T=5.5,Y=rxy,MW=None,MWb='n/a',MWt='-',pmw=None,pmh=None,
    IT=rx_it,ITm='Derived: 139,600/491 + 66,000/320',pit=rxy/rx_it,pith=rxy/rx_it*1e6/H,G=257600-52000,Gt='Reported (sum)',gh=gpuhr(37600,205600,5.5),
    pf=None,apit=None,agh=None,flag='Blended; assumes Narvik inside $43.8B',conf='Low',math=f'37,600/5.5/{rx_it:.0f} = ${rxy/rx_it:.2f}M; 37,600e6/(205,600 x 8,760 x 5.5) = {gpuhr(37600,205600,5.5):.2f}')
# P3 Lambda
add(ID='P3',Deal='Microsoft–Lambda',Chip='GB300 (partly)',TCV=None,T=None,Y=None,MW=None,MWb='not disclosed',MWt='-',pmw=None,pmh=None,IT=None,ITm='not disclosed',pit=None,pith=None,
    G=None,Gt='"tens of thousands"',gh=None,pf=None,apit=None,agh=None,flag='-',conf='n.d.',math='No TCV, term, MW or GPU count public.')
# CoreWeave
for id_,nm,chip,tcv,t in [('P4','Meta–CoreWeave #1','GB300 (press)',14200,deals[9][8]),('P5','Meta–CoreWeave #2','incl. Vera Rubin',21000,deals[10][8]),
                          ('P6','OpenAI–CoreWeave T1','not disclosed',11900,5),('P7','OpenAI–CoreWeave T2','not disclosed',4000,3.9),('P8','OpenAI–CoreWeave T3','not disclosed',6500,deals[13][8])]:
    add(ID=id_,Deal=nm,Chip=chip,TCV=tcv,T=t,Y=tcv/t,MW=None,MWb='not disclosed',MWt='-',pmw=None,pmh=None,IT=None,ITm='not disclosed',pit=None,pith=None,
        G=None,Gt='not disclosed',gh=None,pf=None,apit=None,agh=None,flag='-',conf='High (TCV/yr floor); n.d. per unit',
        math=f'{tcv:,.0f}/{t} = {tcv/t:,.0f}/yr (span from signing; ramp means billed-year value is higher)')
# O1 Anthropic Nscale
ay=44600/6; ag=460*GPM['VR_IT']
add(ID='O1',Deal='Anthropic–Nscale (Vera Rubin)',Chip='Vera Rubin',TCV=44600,T=6,Y=ay,MW=460,MWb='IT',MWt='Reported (S-1 table) / press',pmw=ay/460,pmh=ay/460*1e6/H,
    IT=460,ITm='Reported (S-1 site #17)',pit=ay/460,pith=ay/460*1e6/H,G=ag,Gt=f'Estimated: 460 x 320 = {ag:,}',gh=gpuhr(44600,ag,6),pf=None,apit=ay/460,agh=None,
    flag='Rubin - per-GPU not comparable to GB300',conf='Medium',math=f'44,600/6 = {ay:,.0f}/yr; /460 = ${ay/460:.2f}M/IT MW-yr = ${ay/460*1e6/H:,.0f}/MW-hr; 44,600e6/({ag:,} x 8,760 x 6) = {gpuhr(44600,ag,6):.2f} (SAVRN est. $5.83)')
# O2 Oracle
oy=300000/5
add(ID='O2',Deal='OpenAI–Oracle (Stargate)',Chip='Mixed',TCV=300000,T=5,Y=oy,MW=4500,MWb='Unlabeled',MWt='Press',pmw=oy/4500,pmh=oy/4500*1e6/H,
    IT=4500/1.2,ITm='Estimated: if 4.5 GW is gross, /PUE 1.2',pit=oy/(4500/1.2),pith=oy/(4500/1.2)*1e6/H,G=None,Gt='not disclosed',gh=None,pf=None,apit=None,agh=None,
    flag='Mixed chips; MW basis unknown',conf='Low-Medium',math=f'300,000/5 = {oy:,.0f}/yr; /4,500 = ${oy/4500:.2f}M/MW-yr; if gross -> /3,750 IT = ${oy/3750:.2f}M')
# O3 xAI
add(ID='O3',Deal='Anthropic–SpaceX/xAI',Chip='unspecified NVIDIA',TCV=None,T=2.9,Y=15000,MW=None,MWb='not disclosed',MWt='-',pmw=None,pmh=None,IT=None,ITm='not disclosed',pit=None,pith=None,
    G=325000,Gt='Reported',gh=1250e6/(325000*730),pf=None,apit=None,agh=None,flag='Chip unknown; 90-day termination',conf='Medium-Low',math='1,250e6/(325,000 x 730 hrs/month) = 5.27; TCV/yr = 1.25 x 12 = $15,000M')
# REF
add(ID='REF',Deal='TeraWulf–Fluidstack colo (facility-only)',Chip='n/a',TCV=3700,T=10,Y=370,MW=200,MWb='IT',MWt='Reported',pmw=1.85,pmh=1.85e6/H,IT=200,ITm='Reported',pit=1.85,pith=1.85e6/H,
    G=None,Gt='-',gh=None,pf=None,apit=None,agh=None,flag='Facility only - not a compute price',conf='High',math='3,700/10/200 = $1.85M/IT MW-yr (~$154/kW-month)')

# ---------------- ANCHORS ----------------
AC=['Chip','Anchor','$ per GPU-hr','$M per MW-yr','Basis','Type','Tag','Source','Confidence','Notes']
anch=[
 ['GB300','Microsoft–Nebius implied (5-yr take-or-pay)',round(gpuhr(17392.9,100000,5),2),round(tcvyr/it_mw,2),'per IT MW via 491/MW','Contract (implied)','Derived (GPU count press)',S['NB_20F']+'; '+S['DCD_100K'],'Medium','$4.43 at $19.4B; ceiling because ">100k" GPUs'],
 ['GB300','Microsoft–Nscale/Aker Narvik implied (5-yr)',round(gpuhr(6200,52000,5),2),round(nvy/nv_it,2),'per IT MW via 491/MW','Contract (implied)','Derived (all inputs disclosed)',S['AKER']+'; '+S['NS_PR'],'High','CCIR grade A'],
 ['GB300','Microsoft–IREN implied (5-yr)',round(gpuhr(9700,ig,5),2),9.70,'IT MW reported','Contract (implied)','Derived/Estimated (GPU count from 491/MW)',S['IREN_8K'],'Medium',f'Range ${gpuhr(9700,ig,5):.2f}-{gpuhr(9700,72000,5):.2f} for 491-360 GPUs/IT MW'],
 ['GB300 / B300 (IT MW)','IREN recent 3-yr contracts',round(20/(GPM['GB300_IT']*H)*1e6,2),20.0,'per IT MW (">$20m revenue per MW (IT)")','Contract (company KPI)','Reported $/MW; $/GPU-hr Derived at 491/MW',S['IREN_FY26']+' (p.1)','Medium','Active discussions at ~$25M/IT MW -> ~$5.81/GPU-hr'],
 ['Mixed (new-gen)','Nebius Q2-26 deals ACV',round(20/(GPM['GB300_IT']*H)*1e6,2),20.0,'per MW (basis not labeled); ">$20M"','Contract (company KPI)','Reported $/MW; $/GPU-hr Derived at 491/MW (=$5.58 at 409/gross MW)',S['NB_Q2']+' (p.3-4)','Medium','2026 base ~$12M/MW; Q3 short-term >$40M ($40-50M opportunity)'],
 ['GB300','Astra Intellectus GB300 NVL72, 30 racks, 3-yr lease (published)',3.70,None,'per GPU-hr','Posted reserved price','Reported (provider page via SAVRN)',S['SAVRN'],'Medium',''],
 ['GB300','Verda GB300 on-demand / 24-month',10.11,None,'per GPU-hr','List','Reported (via SAVRN)',S['SAVRN'],'Medium','24-month: $7.58 (25% off)'],
 ['GB300','Oracle OCI BM.GPU.GB300.4 on-demand',18.00,None,'per GPU-hr','List','Reported (via SAVRN)',S['SAVRN'],'Medium',''],
 ['GB200','CoreWeave GB200 NVL72 on-demand',10.50,None,'per GPU-hr ($42/4-GPU instance)','List','Reported',S['CW_PRICES'],'High',''],
 ['GB200','AWS u-p6e-gb200x72 Capacity Block',10.58,None,'per GPU-hr','List','Reported','https://aws.amazon.com/ec2/capacityblocks/pricing/','High',''],
 ['GB200','Azure ND GB200 v6, 3-yr reserved (East US)',11.90,None,'per GPU-hr','List (reserved)','Reported (Azure retail API via SAVRN)','https://prices.azure.com/api/retail/prices','High','On-demand $27.04'],
 ['B300','Nebius HGX B300 on-demand (eff. 2026-10-01)',9.50,None,'per GPU-hr','List','Reported',S['NB_PRICES'],'High','GB200/GB300 NVL72 = "Contact us"'],
 ['Vera Rubin','Anthropic–Nscale implied',round(gpuhr(44600,ag,6),2),round(ay/460,2),'IT MW reported; GPUs at 320/IT MW','Contract (implied)','Derived/Estimated',S['NS_S1']+'; '+S['ANT_BBG'],'Medium','SAVRN independently estimates $5.83'],
 ['Vera Rubin','Meta–Nebius #2 scenario band',f'{gpuhr(12000,200*320,5):.2f}-{gpuhr(12000,120*320,5):.2f}','12-20 (scenario)','scenario: 120-200 IT MW','Contract (scenario)','Estimated - MW not disclosed',S['NB_META2'],'Low','Not a data point; shows range'],
 ['Vera Rubin','Civo VR NVL72 early access',"from 11.00 (unit '/hr', likely per GPU)",None,'reservation-only','Posted early-access price','Reported (unit ambiguous)',S['CIVO'],'Low','Varies with deposit/term; delivery Q1-2027'],
 ['Vera Rubin','NVIDIA list price','not disclosed',None,'-','-','not disclosed','-','-','NVIDIA publishes no Rubin list price; rack HW est. $7.8M (Morgan Stanley via Yahoo) to $9.1M (Bernstein via TrendForce), per SAVRN'],
 ['unspecified','Anthropic–SpaceX/xAI',5.27,None,'per GPU-hr','Contract (implied)','Derived (prospectus terms via CCIR)',S['CCIR'],'Medium-Low','90-day termination; chip unknown'],
 ['Facility only','TeraWulf–Fluidstack colo',None,1.85,'per IT MW','Lease','Derived',S['WULF'],'High','Shell+power infrastructure only ~$154/kW-mo'],
]

# ---------------- NOTES ----------------
notes = [
 ['Purpose','Put every hyperscaler/anchor GPU-capacity deal on one basis and test whether Nebius signed cheap anchors vs peers. Public information only. Built 2026-10-04 (ET) by Nebius GPU Market & Comps.'],
 ['Tags','Reported = stated in filing/press release/credible press (URL given). Derived = arithmetic on Reported inputs (math in "Math shown"). Estimated = needs an assumption (method stated). "not disclosed" = no public figure; we do not guess.'],
 ['Hours','8,760 hours/yr. $/GPU-hr = TCV / (GPUs x 8,760 x term yrs). Assumes ratable spend over the full term, 100% billed hours, all TCV = compute. Ramps, non-compute scope (storage/network) and partial billing make true rates HIGHER, so implied rates are floors.'],
 ['GPUs per MW',f'From {CAPFILE}: GB300 NVL72 491 per IT MW / 409 per gross MW; Vera Rubin NVL72 320 per IT MW / 267 per gross MW; Hopper 713 / 548. Estimated (rack-power specs Reported; overhead and PUE Estimated).'],
 ['MW basis','Main comparison column = $M per IT MW-yr. Where IT MW is disclosed we use it (IREN 200, Anthropic–Nscale 460). Where only GPU count is public we derive IT MW = GPUs / GPUs-per-IT-MW (Nebius MSFT, Narvik, Nscale). This makes $/IT MW-yr and $/GPU-hr mutually consistent for the GB300 rows.'],
 ['Microsoft–Nebius $1,324/MW-hr check','VERIFIED as arithmetic on a 300 MW basis: 17,392.9/5/300 = $11.595M/MW-yr; /8,760 = $1,323.6/MW-hr. CORRECTION on basis: the 300 MW is unlabeled site capacity that Capacity & Power reads as gross (~230-250 MW IT). On an IT basis the same deal is $14.5M/IT MW-yr (240 MW IT) to $17.1M/IT MW-yr ($1,950/IT MW-hr) at the GPU-derived 204 MW IT (100k GB300 / 491). Do not compare $11.6M (gross-ish) to IREN $9.7M (IT) - that flatters Nebius by ~30-45%.'],
 ['Prepayment adjustment (Estimated)',f'Converts each deal into an equivalent level-pay TCV with equal PV at an 8% annual discount rate (monthly). Nebius: $6,958.1M upfront treated as received at start and credited ratably over 60 months -> factor {nb_adj:.3f}. IREN: 20% of TCV paid before delivery, credited pro rata over months 25-60 (8-K) -> factor {ir_adj:.3f}. Others: prepayment not disclosed -> no adjustment (their adjusted values are understated if they have prepay). Real tranche timing differs; adjustment is directional.'],
 ['Who pays power/facility','All compute rows are full-stack GPU services: provider supplies GPUs, network, DC and power (no bare-metal-plus-customer-power deals found). Facility component reference: TeraWulf–Fluidstack colo $1.85M/IT MW-yr (~$154/kW-mo), i.e. ~10-20% of full-stack $/MW-yr. Power cost differences (Narvik hydro vs NJ fuel cells/gensets vs ERCOT) hit provider margin, not the price column. No adjustment applied to prices; flagged in Deals sheet.'],
 ['Chip mismatch','$/GPU-hr is only apples-to-apples within a chip. Same-chip GB300 set: N1, P1, P2a (all Microsoft, all ~5-yr). Vera Rubin rows (N3, O1, P2 part) compared per IT MW only. CoreWeave rows lack MW/GPU and appear as $/yr only.'],
 ['Term','TCV/yr uses stated term; for CoreWeave orders the span from signing to end date (so $/yr is a floor). Nscale MSFT terms are 5-6 yrs per tranche; 5.5 used.'],
 ['Nscale MW mapping','S-1 Properties table gives IT MW per site but not per customer. Microsoft-linked sites (Estimated): Portugal colo 30, US leased 200 (Texas), Portugal owned 200 (SIN02), UK owned 50 (Loughton), Norway owned 100+100 (Narvik) = ~680 IT MW vs 596 IT MW GPU-derived. Both shown.'],
 ['Meta–Nebius MW','Neither Meta deal discloses MW or GPU count. Meta #1 ~50 MW comes from an unsourced substack - LOW confidence and illustrative. Meta #2 shown as break-even MW needed to match peers; we do not assert a $/MW.'],
 ['Not double counting','P2a (Narvik) sits inside P2 (Nscale aggregate). N1o is an upside case of N1. N4 backstop is contingent and excluded. REF is not a compute deal.'],
 ['Key gaps','No public GPU count/MW for: Meta–Nebius #1 & #2, all CoreWeave orders, Microsoft–Lambda (no TCV either). No public Rubin list price. Nebius never labels MW as IT vs gross. Microsoft–Nebius GPU count is press (">100k"), so its $/GPU-hr is a ceiling: at 150k GPUs (unverified claim; Capacity team says power cannot support it) it falls to $2.65.'],
 ['Files','Script: /workspace/nebius/parts/build_contract_economics.py. Chart: /workspace/nebius/market/charts/contract_usd_per_mw_yr.png'],
]

# ---------------- WRITE ----------------
wb = Workbook()
hdr_fill = PatternFill('solid', fgColor='1F3864'); hdr_font = Font(bold=True, color='FFFFFF')
thin = Side(style='thin', color='BFBFBF'); border = Border(left=thin, right=thin, top=thin, bottom=thin)
nb_fill = PatternFill('solid', fgColor='E2EFDA')
def style_sheet(ws, widths, numfmt=None, wrap_cols=()):
    for c in ws[1]:
        c.fill = hdr_fill; c.font = hdr_font; c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = border
    ws.row_dimensions[1].height = 45
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = border; c.alignment = Alignment(wrap_text=True, vertical='top')
            if numfmt and c.column in numfmt and isinstance(c.value, (int, float)): c.number_format = numfmt[c.column]
        if isinstance(row[0].value, str) and row[0].value.startswith('N'):
            for c in row: c.fill = nb_fill
    ws.freeze_panes = 'C2'

ws = wb.active; ws.title = 'Deals'
ws.append(DC)
for d in deals: ws.append(d)
style_sheet(ws, [6,36,14,12,18,12,16,12,9,30,8,30,18,28,12,30,34,34,28,36,36,60,50,22,36],
            numfmt={6:'#,##0.0',8:'#,##0',9:'0.0#',11:'#,##0',15:'#,##0'})

ws2 = wb.create_sheet('Normalized')
ws2.append(NC)
keys=['ID','Deal','Chip','TCV','T','Y','MW','MWb','MWt','pmw','pmh','IT','ITm','pit','pith','G','Gt','gh','pf','apit','agh','dcw','dnb','flag','conf','math']
for n in norm:
    same = 'GB300' in n['Chip'] and 'Rubin' not in n['Chip']
    n['dcw'] = (1 - n['gh']/10.50) if (n.get('gh') and same) else None
    n['dnb'] = (1 - n['gh']/9.50) if (n.get('gh') and same) else None
    ov = {'P2':{'MW':'n/a (no stated MW)','pmw':'n/a','pmh':'n/a'},'P2a':{'MW':'n/a (no stated MW)','pmw':'n/a','pmh':'n/a','pf':'prepay n.d.'},
          'P2b':{'MW':'n/a','pmw':'n/a','pmh':'n/a'},'REF':{'gh':'n/a (no GPUs)'},'O3':{'TCV':'$1.25B/month','pmw':'not disclosed','pit':'not disclosed'},
          'O1':{'pf':'prepay n.d.'}}.get(n['ID'],{})
    row=[]
    for k in keys:
        v=n.get(k)
        if k in ov: v=ov[k]
        elif v is None and k in ('pmw','pit','gh','TCV','MW','IT'): v='not disclosed'
        row.append(v)
    ws2.append(row)
style_sheet(ws2, [6,34,16,11,8,11,8,16,22,11,11,10,34,11,11,10,30,10,9,12,12,11,11,26,18,70],
            numfmt={4:'#,##0',5:'0.0#',6:'#,##0',7:'#,##0',10:'$0.00',11:'$#,##0',12:'#,##0',14:'$0.00',15:'$#,##0',16:'#,##0',18:'$0.00',19:'0.000',20:'$0.00',21:'$0.00',22:'0%',23:'0%'})
r = ws2.max_row + 2
ws2.cell(r,1,'Units: TCV and TCV/yr in $M; per-MW-yr columns in $M; per-MW-hr and per-GPU-hr in $. Green rows = Nebius. ADJ columns = prepayment PV-adjusted (see Notes_Method).').font = Font(italic=True)

ws3 = wb.create_sheet('Rubin_GB300_Anchors')
ws3.append(AC)
for a in anch: ws3.append(a)
style_sheet(ws3, [16,48,14,12,30,22,30,60,12,50], numfmt={3:'$0.00',4:'$0.00'})

ws4 = wb.create_sheet('Notes_Method')
ws4.append(['Topic','Note'])
for nn in notes: ws4.append(nn)
verdict = [
 ['VERDICT','No - not cheap vs peer anchors on a like-for-like basis. Same buyer (Microsoft), same chip (GB300), same ~5-yr tenor: Nebius ~$3.97/GPU-hr ($4.31 prepay-adjusted) vs Narvik $2.72 and IREN ~$2.25-3.07 -> Nebius is ~30-75% HIGHER per GPU-hr; per IT MW-yr $17.1M vs $9.7M (IREN) / $11.7M (Narvik) / $13.4M (Nscale blended). Only the team\'s 300-MW (gross-ish) basis makes Nebius look in line ($11.6M).'],
 ['BUT','All hyperscaler anchors - Nebius\'s included - are priced well below today\'s market: Nebius 2026 base ~$12M/MW vs Q2-26 deals >$20M/MW and Q3 short-term >$40M/MW (Q2 letter p.3-4); IREN Microsoft $9.7M/IT MW vs IREN recent 3-yr >$20M/IT MW. Implied GB300 anchor rates ($2.3-4.0) are 60-80% below GB200/GB300 list ($10.50-18.00). The anchors are cheap vs the 2026 spot/term market, not vs each other.'],
 ['Meta','Meta #1 (~$576M/yr) and Meta #2 ($2.4B/yr firm, Vera Rubin) cannot be judged per MW - not disclosed. To match Anthropic–Nscale VR economics ($16.2M/IT MW-yr), Meta #2 must cover <=~148 IT MW (~47k VR GPUs); if it covers ~200 MW it is priced at the $12M base, i.e. cheap.'],
]
for v in verdict: ws4.append(v)
style_sheet(ws4, [28,160])
for row in ws4.iter_rows(min_row=2):
    if row[0].value in ('VERDICT','BUT','Meta'):
        for c in row: c.font = Font(bold=True)

out = '/workspace/nebius/parts/_rebuild/contract_economics.xlsx'  # 10/8 lock guard: locked file has in-place edits
wb.save(out); print('saved', out)
print('nb_adj', nb_adj, 'ir_adj', ir_adj)
for n in norm: print(n['ID'], n['Deal'], n.get('pmw'), n.get('pit'), n.get('gh'), n.get('apit'), n.get('agh'))

import json
json.dump(norm, open('/workspace/nebius/parts/contract_economics_norm.json','w'), indent=1, default=str)

# Re-append the ShortTerm_Mix sheet (added 2026-10-04) so rebuilds keep it
import subprocess as _sp
_sp.run(['python3','/workspace/nebius/parts/add_shortterm_mix.py'], check=True)
