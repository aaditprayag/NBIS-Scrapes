#!/usr/bin/env python3
"""Build nebius_capacity.xlsx (Capacity & Power analyst). Run: python3 build_capacity.py
Inputs blue font, formulas black. Every written cell is logged to Tags_Sources."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

OUT = '/workspace/nebius/capacity/nebius_capacity.xlsx'
BLUE = Font(color='0000FF'); BLACK = Font(color='000000'); BOLD = Font(bold=True)
HDR = PatternFill('solid', fgColor='DDEBF7')
GRADEFILL = {'Green': 'C6EFCE', 'Yellow': 'FFEB9C', 'Red': 'FFC7CE', 'Yellow/Red': 'F8CBAD'}

Q = [f'Q{q} {y}' for y in range(2024, 2031) for q in range(1, 5)]
QI = {q: i for i, q in enumerate(Q)}
HIST_END = QI['Q2 2026']

# ---------------- source shorthands ----------------
S = dict(
 q325='https://www.sec.gov/Archives/edgar/data/1513845/000110465925109806/tm2530882d1_ex99-2.htm',
 q425='https://www.sec.gov/Archives/edgar/data/1513845/000110465926013946/tm266173d1_ex99-2.htm',
 q126='https://www.sec.gov/Archives/edgar/data/1513845/000110465926059872/tm2614392d1_ex99-2.htm',
 q226='https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm',
 q226call='https://www.fool.com/earnings/call-transcripts/2026/08/19/nebius-nbis-q2-2026-earnings-call-transcript/ (cached /workspace/nebius/src/q2t.txt)',
 f20='https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm',
 ms6k='https://www.sec.gov/Archives/edgar/data/1513845/000110465925088312/tm2525580d1_6k.htm',
 q225='https://www.sec.gov/Archives/edgar/data/1513845/000110465925075028/tm2522866d1_ex99-2.htm',
 alt='/workspace/nebius/capacity/altdata/altdata_timelines.xlsx (Timeline, Vineland_MW)',
 enr='https://www.enr.com/articles/63695-inside-the-vineland-data-center-power-dispute',
 flood='https://newjerseymonitor.com/2026/10/01/nj-data-center-polluted-air-for-nearly-a-year-before-regulators-intervened/',
 flood2='https://floodlightnews.org/new-jersey-dataones-illegal-still-running/',
 njgov='https://www.nj.gov/governor/news/2026/20260922a.shtml',
 hntr='https://hntrbrk.com/breaking-news/vineland',
 whyy='https://whyy.org/articles/vineland-planning-board-approves-data-center-plan/',
 vlagenda='https://www.vinelandcity.org/Archive/Planning%20Board/Agendas/2026/Agenda%208-17-26%20%28Special%20Meeting%29.pdf',
 indpdf='https://www.independencemo.gov/sites/default/files/2026-02/Nebius%20x%20Independence%2C%20MO.cleaned.pdf',
 indcp='https://citizenportal.ai/articles/7477161 (Independence PUAB: IPP PPA/ESA timeline)',
 indneb='https://nebius.com/independencemo',
 indgb='https://ded.mo.gov/press-room/nebius-breaks-ground-gigawatt-scale-ai-factory-independence',
 butneb='https://nebius.com/schuylkillcountypa',
 butrh='https://www.republicanherald.com/2026/09/01/nebius-gets-green-light-to-expand-its-data-center-plan-in-butler-twp/',
 butppl='https://www.republicanherald.com/2026/04/24/butler-twp-switchyard-to-anchor-data-center-in-ppl-powerline-overhaul/',
 butdcd='https://www.datacenterdynamics.com/en/news/nebius-details-plans-for-12gw-data-center-campus-in-pennsylvania/',
 bhmltr='https://abc3340.com/resources/pdf/9222f76a-3063-46e4-aee7-796b30ed611f-NebiusResponsestoCityLetter_Redacted.pdf',
 bhmneb='https://nebius.com/birminghamal',
 bhmlaw='https://www.wbhm.org/local-news/2026-09-18/action-picks-up-in-lawsuits-targeting-birmingham-data-center',
 bhmperm='https://bhamnow.com/2026/05/18/40-million-permits-birmingham-data-center/',
 bet_vdn='https://www.lavoixdunord.fr/1740138/article/2026-09-24/la-premiere-ia-factory-des-hauts-de-france-va-bientot-prendre-vie-bethune',
 bet_nfi='https://www.nordfranceinvest.com/success-story/dutch-company-nebius-is-setting-up-the-first-ai-data-center-in-the-hauts-de-france-region/',
 bet_ppve='https://presseagence.fr/bethune-numerique-un-projet-de-datacenter-soumis-a-consultation-publique/',
 bet_reu='https://www.reuters.com/technology/amsterdam-based-ai-firm-nebius-build-240mw-data-centre-near-lille-france-2026-02-12/',
 lap_fira='https://fira.fi/en/tiedote/construction-and-project-development-company-fira-selected-as-main-contractor-for-one-of-europes-largest-ai-data-center-projects/',
 lap_pr='https://nebius.com/newsroom/nebius-to-construct-310-mw-ai-factory-in-finland',
 lap_cav='https://www.caverion.com/newsroom/releases/2026/caverion-to-deliver-eur-15-million-substation-for-nebius-data-center-in-finland-creating-jobs-for-more-than-80-professionals/',
 man_pr='https://assets.nebius.com/assets/6590d0a0-0fc4-4396-bc3f-f52a9ccb770e/20260331%20Nebius%20to%20construct%20310%20MW%20AI%20factory%20in%20Finland.pdf',
 man2_cav='https://www.assemblincaverion.com/news-and-media/news/2026/caverion-finland-strengthens-its-position-in-the-data-centre-market-and-will-deliver-a-eur-15-million-substation-for-ai-cloud-company-nebius/',
 man_dcd='https://www.datacenterdynamics.com/en/news/nebius-to-deploy-5mw-nvidia-h200-cluster-at-patmos-data-center-in-kansas-city-missouri/',
 kc_dcd='https://www.datacenterdynamics.com/en/news/patmos-adds-10mw-of-capacity-at-kansas-city-data-center/',
 kc_job='https://careers.nebius.com/?gh_jid=4975997101',
 paris='https://dcatlas.io/en/explore/facilities/nebius-equinix-paris (3rd-party, low confidence)',
 kef='https://nebius.com/blog/posts/300-mw-new-jersey-and-iceland-regions',
 lon='https://datacentrenews.uk/story/nebius-fills-ark-s-longcross-data-centre-to-capacity',
 lon2='https://www.datacenterdynamics.com/en/news/ark-dc-to-add-new-building-to-longcross-data-center-campus-outside-london-uk/',
 kao='https://www.datacenterdynamics.com/en/news/nebius-signs-22mw-capacity-agreement-with-kao-data-in-the-uk/',
 kao2='https://www.linkedin.com/posts/infratil_klon-03-development-update-activity-7404073674942656512-l6mm',
 ukpr='https://nebius.com/newsroom/nebius-expands-in-uk-with-more-nvidia-powered-infrastructure-more-customers-and-more-cloud-capabilities-for-agentic-and-enterprise-ai',
 isr='https://www.datacenterdynamics.com/en/news/nebius-signs-80mw-data-center-lease-with-mega-or-in-israel/',
 isr2='https://www.calcalistech.com/ctechnews/article/25ue6vi49',
 modiin='https://nebius.com/newsroom/nebius-brings-nvidia-blackwell-to-israel-with-one-of-the-country-s-first-ai-infrastructure-deployments',
 mad='https://www.linkedin.com/feed/update/urn:li:activity:7505906816590073857',
 mad2='https://www.merlinproperties.com/dia-a-dia/firmamos-un-contrato-de-18-mw-con-nebius-para-nuestro-date-center-madrid-getafe/',
 tal='https://news.err.ee/1610109073/baltics-first-ai-factory-to-open-in-estonia',
 tal2='https://www.datacenterdynamics.com/en/news/nebius-expands-european-presence-announces-deployment-in-estonia-and-second-data-center-in-m%C3%A4nts%C3%A4l%C3%A4-finland/',
 docs='https://docs.nebius.com/overview/regions (accessed 2026-10-04)',
 aib='http://archive.fast-edgar.com/20260930/AB2ZM22DL222P26L222Q2MY2R6GDZ222ZQ82/ea030723301ex99-1.htm (AIB 8-K Ex.99.1, Sep 30 2026)',
 gpumw='/workspace/nebius/capacity/gpu_mw_capex.xlsx',
 sitelist='/workspace/nebius/capacity/site_list.xlsx / site_list.md',
 filings='/workspace/nebius/filings/nebius_filings.xlsx (Quarterly, MW_Disclosures)',
)

# ---------------- GPUs per MW (values reproduced from gpu_mw_capex.xlsx GPUsPerMW formulas, Base rows) ----------------
def gpm(D, E, G, J):
    it = 1000 / (D / E * (1 + G)); return round(it, 1), round(it / J, 1)
GEN = {  # gen: (row in GPUsPerMW, D,E,G,J)
 'Hopper H100/H200': (6, 10.2, 8, 0.10, 1.30),
 'B200 (HGX)':       (9, 14.3, 8, 0.10, 1.30),
 'B300 (HGX)':       (15, 14.5, 8, 0.10, 1.30),
 'GB300 NVL72':      (18, 137, 72, 0.07, 1.20),
 'Vera Rubin NVL72': (21, 230, 72, None, 1.20),
}
GPM = {}
for g, (r, D, E, G, J) in GEN.items():
    if g == 'Vera Rubin NVL72':
        GPM[g] = (320.0, 267.0)  # gpu_mw_capex.md table base values (IT 320 / gross 267)
    else:
        GPM[g] = gpm(D, E, G, J)

# ---------------- Sites ----------------
# fields: key, name, country, own, contract, contract_ev, contracted_MW, basis, pue, mgmt_golive, my_golive, grade, gen, sources
SITES = [
 dict(k='VIN', name='Vineland NJ (DataOne)', c='USA', own='Build-to-suit (DataOne owns shell; Nebius tenant)', tag='Microsoft', tagev='Reported: 6-K Sep 2025 / 20-F: MSFT served from Vineland', cmw=300, basis='Contracted 300 = gross (Estimated); energized rows = IT MW (altdata power-bounded)', pue='not disclosed (IT est. uses PUE 1.2-1.35, Estimated)',
      mg='Phase 1 summer 2025; all 9 MSFT tranches during 2025-26; Vineland in 2026 connected power (Q2-26 call)', my='~107 MW IT ceiling on rented gensets through YE26; Bloom-powered remainder 1H-2H 2027', grade='Yellow/Red', gen='GB300 NVL72 (assumed; not disclosed by Nebius)', basis_it=True,
      src=f"{S['ms6k']}; {S['f20']}; {S['q226']}; {S['q226call']}; {S['enr']}; {S['flood']}; {S['njgov']}; {S['hntr']}; {S['whyy']}; {S['alt']}"),
 dict(k='IND', name='Independence MO (Eastgate)', c='USA', own='Owned greenfield', tag='unassigned (Meta #2 candidate - assumed)', tagev='assumed: Meta #2 = multiple undisclosed locations; Phase 1 Q3-27 is after Meta #2 1H27 tranches', cmw=1200, basis='unknown (likely gross/utility)', pue='not disclosed',
      mg='Operational 2027 (Q1-26 letter); IPL initial power 2H26', my='Phase 1 ~Q3 2027; 600 MW not before 2029', grade='Green', gen='Vera Rubin NVL72 (assumed; company-level VR statement)',
      src=f"{S['indneb']}; {S['indpdf']}; {S['indcp']}; {S['indgb']}; {S['q126']}; {S['alt']}"),
 dict(k='BUT', name='Butler Twp PA (Highridge)', c='USA', own='Owned greenfield', tag='unassigned', tagev='not disclosed', cmw=1200, basis='unknown (PPL ESA = utility/gross)', pue='not disclosed',
      mg='260 MW by Oct 2027; 660 MW Jan 2029; 1,000+ MW Nov 2029', my='First MW Q4 2027 at best; base ~Q2 2028-2029 (PPL lines late-27 to early-29; 3-yr build)', grade='Red', gen='Vera Rubin NVL72 (assumed)',
      src=f"{S['butneb']}; {S['butrh']}; {S['butppl']}; {S['butdcd']}; {S['alt']}"),
 dict(k='BHM', name='Birmingham AL (BHM01, Oxmoor)', c='USA', own='Owned brownfield', tag='unassigned', tagev='not disclosed', cmw=300, basis='unknown (300 MW is press figure, not Nebius)', pue='not disclosed',
      mg='Initial-phase energization Q4 2026 (Nebius page); operational 2027 (Q1-26 letter)', my='115 kV NLT Dec 31 2026; first revenue MW 2027; 230 kV Oct 2027', grade='Yellow', gen='Vera Rubin NVL72 (assumed)',
      src=f"{S['bhmneb']}; {S['bhmltr']}; {S['bhmperm']}; {S['bhmlaw']}; {S['alt']}"),
 dict(k='BET', name='Bethune FR (Azur Datacenter, ex-Bridgestone)', c='France', own='Colocation (20-F); Azur finances/builds', tag='Meta (#2 likely) - press', tagev="La Voix du Nord 2026-09-24: part of site 'fera travailler l'IA de Meta' (Estimated mapping to Meta #2)", cmw=240, basis='unknown', pue='not disclosed',
      mg='First capacity late summer 2026; ~120 MW YE26; 240 MW 2027', my='Start Q4 2026 (~27 MW initial); 120 MW ~1H27; prefect ICPE decision pending', grade='Yellow', gen='GB300 NVL72 (Estimated: EUR4M/rack per press)',
      src=f"{S['bet_reu']}; {S['bet_vdn']}; {S['bet_nfi']}; {S['bet_ppve']}; {S['alt']}"),
 dict(k='LAP', name='Lappeenranta FI (Pajarila / Polarnode)', c='Finland', own="Owned per Nebius; Polarnode says Nebius 'operates' its campus (conflict)", tag='unassigned', tagev='not disclosed', cmw=310, basis='unknown', pue='not disclosed',
      mg='First capacity to customers 2027', my='First building H1 2027 (Fira); substation spring 2027', grade='Green', gen='Vera Rubin NVL72 (assumed; PR: Blackwell & Rubin)',
      src=f"{S['lap_pr']}; {S['lap_fira']}; {S['lap_cav']}"),
 dict(k='MA1', name='Mantsala FI #1', c='Finland', own='Owned', tag='self-serve (assumed); Meta #1 site not disclosed', tagev='assumed', cmw=75, basis='unknown; PUE ~1.1 (company)', pue='~1.1 (Nebius, company statement)',
      mg='75 MW end-2025/early-2026', my='75 MW built (completed early 2026)', grade='Green', gen='Hopper (first 25 MW) + GB300 NVL72 (expansion, Estimated split)',
      src=f"{S['man_pr']}; {S['alt']}"),
 dict(k='MA2', name='Mantsala FI #2 (Kapuli)', c='Finland', own='Owned (implied)', tag='unassigned', tagev='not disclosed', cmw=70, basis='unknown (70 MW substation)', pue='not disclosed',
      mg="Capacity 'from 2027'", my='Q4 2027 earliest (substation completes autumn 2027)', grade='Yellow', gen='Vera Rubin NVL72 (assumed)',
      src=f"{S['man2_cav']}; {S['tal2']}"),
 dict(k='KC', name='Kansas City MO (Patmos)', c='USA', own='Colocation', tag='self-serve (assumed)', tagev='assumed', cmw=40, basis='unknown', pue='not disclosed',
      mg='5 MW live Apr 2025; expandable to 40 MW', my='Operational; Nebius share of 35 MW facility not disclosed', grade='Green', gen='H200 (first 5 MW) + B200 (rest)',
      src=f"{S['man_dcd']}; {S['kc_dcd']}; {S['kc_job']}"),
 dict(k='PAR', name='Paris (Equinix PA10)', c='France', own='Colocation', tag='self-serve (assumed)', tagev='assumed', cmw=5, basis='unknown; MW not disclosed (5 = 3rd-party)', pue='not disclosed',
      mg='Live Nov 2024', my='Operational', grade='Green', gen='Hopper H100/H200', src=f"{S['paris']}; {S['docs']}"),
 dict(k='KEF', name='Keflavik IS (Verne)', c='Iceland', own='Colocation', tag='self-serve (assumed)', tagev='assumed', cmw=10, basis='unknown', pue='not disclosed',
      mg='Live Mar 2025', my='Operational (private region)', grade='Green', gen='Hopper H100/H200', src=f"{S['kef']}; {S['docs']}"),
 dict(k='LON', name='London Longcross (Ark LP01)', c='UK', own='Colocation', tag='self-serve (assumed)', tagev='assumed', cmw=19, basis='unknown (GPU/MW fit suggests gross)', pue='not disclosed',
      mg='Launched Nov 2025; 16 MW live + 3 MW in delivery (Jun 2026); LP02 36 MW (Nebius share n/d)', my='19 MW by ~Q4 2026; LP02 share/timing not disclosed', grade='Green', gen='B300 (HGX)', src=f"{S['lon']}; {S['lon2']}; {S['docs']}"),
 dict(k='HAR', name='Harlow UK (Kao Data KLON-02/03)', c='UK', own='Colocation (10-yr)', tag='self-serve (assumed)', tagev='assumed', cmw=22, basis='unknown', pue='not disclosed',
      mg='UK 65 MW fully ramped 2027', my='KLON-03 completion end-2026; MW early 2027', grade='Green', gen='B300 (HGX) (assumed; docs uk-south2 lists RTX PRO 6000)', src=f"{S['kao']}; {S['kao2']}; {S['ukpr']}"),
 dict(k='UK3', name='UK 3rd new deployment (unnamed)', c='UK', own='not disclosed', tag='unassigned', tagev='not disclosed', cmw=None, basis='unknown', pue='not disclosed',
      mg='Part of 65 MW UK by 2027', my='No evidence found', grade='Red', gen='not disclosed', src=S['ukpr']),
 dict(k='MOD', name="Modi'in IL (Mega Or)", c='Israel', own='Colocation', tag='self-serve (assumed); 25% to Israel Innovation Authority', tagev='Reported (IIA share)', cmw=8, basis='IT (4,000 B200 / 8 MW = 500/MW fits IT model)', pue='not disclosed',
      mg='Launched Oct 2025', my='Operational', grade='Green', gen='B200 (HGX)', basis_it=True, src=f"{S['modiin']}; {S['isr']}"),
 dict(k='MAS', name='Masmiyya IL (Mega DC)', c='Israel', own='Colocation (5-yr)', tag='unassigned', tagev='not disclosed', cmw=22, basis='unknown', pue='not disclosed',
      mg='Delivery Q3 2026 (landlord)', my='Delivery not confirmed as of Oct 4 2026; Q4 2026-Q1 2027', grade='Yellow', gen='GB300 NVL72 (assumed)', src=f"{S['isr']}; {S['isr2']}"),
 dict(k='BEI', name='Beit Shemesh IL (Mega DC)', c='Israel', own='Colocation (5-yr)', tag='unassigned', tagev='not disclosed', cmw=58, basis='unknown', pue='not disclosed',
      mg='Staged Q3 2026 - Q1 2027 (landlord)', my='Staged Q4 2026 - Q3 2027', grade='Yellow', gen='GB300 NVL72 (assumed)', src=f"{S['isr']}; {S['isr2']}"),
 dict(k='MAD', name='Madrid-Getafe ES (Merlin Edged)', c='Spain', own='Colocation', tag='self-serve (assumed)', tagev='assumed', cmw=18, basis='unknown', pue='not disclosed',
      mg="Phase 1 live Sep 2026; phases 2-3 'coming weeks'", my='Phase 1 live; 18 MW by Q4 2026-Q1 2027', grade='Green', gen='B300 (HGX) (RTX PRO 6000 phase 1 not modeled)', src=f"{S['mad']}; {S['mad2']}; {S['docs']}"),
 dict(k='TAL', name='Tallinn EE (Greenergy Huuru)', c='Estonia', own='Colocation', tag='unassigned', tagev='not disclosed', cmw=22, basis='unknown', pue='not disclosed',
      mg='22 MW first phase later in 2026', my='Q4 2026-Q2 2027; Estonian govt security review discussed at Cabinet (Aug 2026)', grade='Yellow', gen='GB300 NVL72 (assumed)', src=f"{S['tal']}; {S['tal2']}"),
 dict(k='MN', name='Minnesota colo (us-north1, Woodbury)', c='USA', own='Colocation (20-F)', tag='self-serve (assumed)', tagev='assumed', cmw=None, basis='unknown', pue='not disclosed',
      mg='Secured Feb 2026', my='Region live in docs (B300); MW not disclosed -> in plug', grade='Yellow', gen='B300 (HGX)', src=f"{S['f20']}; {S['docs']}"),
 dict(k='OK', name='Oklahoma colo (unnamed)', c='USA', own='Colocation (20-F)', tag='unassigned', tagev='not disclosed', cmw=None, basis='unknown', pue='not disclosed',
      mg='Secured Feb 2026', my='No region in docs; no landlord/MW -> in plug', grade='Red', gen='not disclosed', src=f"{S['f20']}; {S['docs']}"),
 dict(k='FR2', name='France 2nd colo (unnamed; possibly eu-west2)', c='France', own='Colocation (20-F)', tag='unassigned', tagev='not disclosed', cmw=None, basis='unknown', pue='not disclosed',
      mg='Secured Feb 2026', my='eu-west2 (France, B300) live in docs - mapping is inference; MW n/d -> in plug', grade='Yellow', gen='B300 (HGX) (inference)', src=f"{S['f20']}; {S['docs']}"),
 dict(k='AIB', name='AIB CLT1 (South Carolina)', c='USA', own='Colocation/lease (12-yr)', tag='unassigned', tagev='not disclosed', cmw=50, basis='IT (50 MW critical IT on 65 MW utility ESA)', pue='65/50 = 1.30 implied (Derived)',
      mg='Hall 1 (25 MW) ~month 10, Hall 2 ~month 14 after Oct 1 2026 ESA start', my='Hall 1 ~Q3-Q4 2027; permits only submitted; project financing not arranged', grade='Red', gen='Vera Rubin NVL72 (assumed)', basis_it=True, src=S['aib']),
]
SK = {s['k']: s for s in SITES}

# ---------------- MW anchors: (quarter, low, high, tag, source) -> forward filled ----------------
A = {
 'VIN': [('Q4 2025', '=ROUND((5+16*1.982)*0.8/1.35,1)', '=ROUND((5+16*1.982)/1.2,1)', 'Estimated', 'altdata Vineland_MW row 2025-10 (16 gensets x 1.982 MW + 5 MW grid; low x0.8/1.35, high /1.2)'),
         ('Q1 2026', 30, 45, 'Estimated', 'altdata Vineland_MW row 2026-02 (32 gensets, ~63 MW nameplate); capped at/below end-Q2 observed running range (33-45) so series is monotonic'),
         ('Q2 2026', 33, 45, 'Estimated', 'altdata Vineland_MW end-Q2 (~25 of 62 running, thermal) per alt-data researcher summary'),
         ('Q3 2026', 57, 75, 'Estimated', 'altdata Vineland_MW Q3 (>=45 of 62 running) per alt-data researcher summary'),
         ('Q4 2026', 57, '=ROUND((5+62*1.982)/1.2,1)', 'Estimated', 'Low = Q3 run-rate held (Bloom not yet energized); High = generator ceiling (62x1.982+5)/1.2 = ~107 MW IT. Assumes gensets permitted by NJDEP ~Nov 2026 deadline'),
         ('Q1 2027', 57, 180, 'Estimated', 'High: first Bloom blocks (328 MW installed/250 guaranteed) energized; low: Bloom slips'),
         ('Q2 2027', 107, 230, 'Estimated', 'Bloom ramp; IT of 300 MW gross ~230-250 (gpu_mw_capex.md)'),
         ('Q3 2027', 160, 250, 'Estimated', 'Bloom ramp'),
         ('Q4 2027', 230, 250, 'Estimated', 'Full 300 MW gross (~230-250 IT) lands in 2027'),
        ],
 'IND': [('Q3 2027', 0, 100, 'Estimated', 'Phase 1 ~Q3 2027 (KSHB/altdata); IPL 200 MW interim'),
         ('Q4 2027', 50, 200, 'Estimated', 'IPP Blue Valley Ph1 200-250 MW Oct 2027'),
         ('Q2 2028', 100, 250, 'Estimated', ''), ('Q4 2028', 200, 250, 'Estimated', 'Power capped ~250 MW until 2029 study'),
         ('Q2 2029', 300, 600, 'Estimated', 'IPL 600 MW early 2029 (study)'), ('Q4 2029', 500, 800, 'Estimated', ''),
         ('Q4 2030', 700, 1200, 'Estimated', '1.2 GW needs on-site generation (evaluating)')],
 'BUT': [('Q4 2027', 0, 130, 'Estimated', 'Mgmt 260 MW Oct 2027; final land-dev plan pending; PPL 230 kV lines late-27 to early-29'),
         ('Q1 2028', 0, 260, 'Estimated', ''), ('Q4 2028', 130, 400, 'Estimated', ''),
         ('Q2 2029', 260, 660, 'Estimated', 'Mgmt 660 MW Jan 2029'), ('Q4 2029', 400, 1000, 'Estimated', 'Mgmt 1,000+ MW Nov 2029'),
         ('Q4 2030', 660, 1200, 'Estimated', '')],
 'BHM': [('Q1 2027', 0, 25, 'Estimated', '115 kV NLT Dec 31 2026; no closed roof Sep 20 2026 imagery'),
         ('Q2 2027', 0, 50, 'Estimated', ''), ('Q3 2027', 25, 75, 'Estimated', ''), ('Q4 2027', 50, 100, 'Estimated', '230 kV switching station Oct 2027'),
         ('Q2 2028', 100, 200, 'Estimated', ''), ('Q4 2028', 200, 300, 'Estimated', 'Full build-out through 2028 (Nebius page)'), ('Q4 2029', 300, 300, 'Estimated', '')],
 'BET': [('Q4 2026', 0, 60, 'Estimated', 'Start "dans quelques semaines" (Sep 24); ~27 MW initial (Nord France Invest); ICPE prefect decision pending'),
         ('Q1 2027', 27, 120, 'Estimated', ''), ('Q2 2027', 60, 180, 'Estimated', ''), ('Q4 2027', 120, 240, 'Estimated', ''),
         ('Q4 2028', 240, 240, 'Estimated', 'Full 240 MW by 2027-2028 (Nord France Invest)')],
 'LAP': [('Q2 2027', 0, 35, 'Estimated', 'First building H1 2027 (Fira); substation 70 MW spring 2027'),
         ('Q3 2027', 0, 70, 'Estimated', ''), ('Q4 2027', 35, 70, 'Estimated', ''), ('Q2 2028', 70, 140, 'Estimated', ''),
         ('Q4 2028', 140, 200, 'Estimated', ''), ('Q4 2029', 220, 310, 'Estimated', ''), ('Q4 2030', 310, 310, 'Estimated', '')],
 'MA1': [('Q1 2024', 25, 25, 'Reported', 'Pre-expansion 25 MW site capacity (DCD Nov 2024)'),
         ('Q3 2025', 25, 50, 'Estimated', 'Phase 1 expansion targeted late Q3-25 (Q1-25 call)'),
         ('Q4 2025', 50, 75, 'Estimated', 'Expansion "should be completed by end of year" (Q3-25 letter)'),
         ('Q1 2026', 75, 75, 'Reported', '75 MW completed earlier this year (Mar 31 2026 PR)')],
 'MA2': [('Q4 2027', 0, 35, 'Estimated', 'Substation completion autumn 2027 (Caverion)'), ('Q1 2028', 35, 70, 'Estimated', ''), ('Q2 2028', 70, 70, 'Estimated', '')],
 'KC':  [('Q2 2025', 5, 5, 'Reported', '5 MW phase 1 live Apr 2025'),
         ('Q3 2025', 5, 15, 'Estimated', 'Phase 2 added Jul 2025, MW n/d'),
         ('Q4 2025', 5, 35, 'Estimated', 'Landlord: facility 35 MW total Sep 2025; Nebius share n/d'),
         ('Q4 2027', 20, 40, 'Estimated', 'Nebius stated expandable to 40 MW')],
 'PAR': [('Q4 2024', 5, 5, 'Estimated', 'MW never disclosed; 5 MW = DC Atlas (low confidence)')],
 'KEF': [('Q1 2025', 10, 10, 'Reported', '10 MW cluster live Mar 2025')],
 'LON': [('Q4 2025', 16, 16, 'Reported', '16 MW launched Nov 2025'), ('Q4 2026', 16, 19, 'Estimated', '+3 MW in delivery (Ark, Jun 2026)'), ('Q1 2027', 19, 19, 'Estimated', '')],
 'HAR': [('Q4 2026', 0, 4.4, 'Estimated', 'KLON-02 share 4.4 MW = 3rd-party inference'), ('Q1 2027', 4.4, 22, 'Estimated', 'KLON-03 on track end-2026'), ('Q3 2027', 22, 22, 'Estimated', '')],
 'MOD': [('Q4 2025', 8, 8, 'Reported', '8 MW, 4,000 B200, launched Oct 20 2025')],
 'MAS': [('Q4 2026', 0, 22, 'Estimated', 'Landlord delivery Q3 2026 not confirmed'), ('Q1 2027', 11, 22, 'Estimated', ''), ('Q2 2027', 22, 22, 'Estimated', '')],
 'BEI': [('Q4 2026', 0, 39, 'Estimated', 'Staged Q3-26..Q1-27; high = 2/3 by YE26 (linear)'), ('Q1 2027', 20, 58, 'Estimated', ''), ('Q2 2027', 39, 58, 'Estimated', ''), ('Q3 2027', 58, 58, 'Estimated', '')],
 'MAD': [('Q3 2026', 6, 9, 'Estimated', 'Phase 1 live mid-Sep 2026 (MW n/d; 1/3-1/2 of 18)'), ('Q4 2026', 9, 18, 'Estimated', 'Phases 2-3 "coming weeks"'), ('Q1 2027', 18, 18, 'Estimated', '')],
 'TAL': [('Q4 2026', 0, 22, 'Estimated', '"Later in 2026"'), ('Q1 2027', 11, 22, 'Estimated', ''), ('Q2 2027', 22, 22, 'Estimated', '')],
 'MN': [], 'OK': [], 'FR2': [], 'UK3': [],
 'AIB': [('Q3 2027', 0, 25, 'Estimated', 'Hall 1 ~month 10 after Oct 1 2026 (~Aug 2027)'), ('Q4 2027', 0, 50, 'Estimated', 'Hall 2 ~month 14 (~Dec 2027)'),
         ('Q1 2028', 25, 50, 'Estimated', 'Financing not arranged; permits submitted only'), ('Q3 2028', 50, 50, 'Estimated', '')],
}
# contracted MW first quarter of announcement/agreement
CONTRACT_Q = {'VIN': ('Q1 2025', 300), 'IND': ('Q4 2025', 1200), 'BUT': ('Q2 2026', 1200), 'BHM': ('Q3 2025', 300), 'BET': ('Q1 2026', 240),
 'LAP': ('Q1 2026', 310), 'MA1': ('Q4 2024', 75), 'MA2': ('Q2 2026', 70), 'KC': ('Q4 2024', 40), 'PAR': ('Q3 2024', 5), 'KEF': ('Q4 2024', 10),
 'LON': ('Q2 2025', 16), 'HAR': ('Q2 2026', 22), 'MOD': ('Q2 2025', 8), 'MAS': ('Q1 2026', 22), 'BEI': ('Q1 2026', 58), 'MAD': ('Q2 2026', 18),
 'TAL': ('Q2 2026', 22), 'AIB': ('Q3 2026', 50)}
CONTRACT_STEP = {'LON': [('Q2 2026', 19)]}
MGMT = {  # site: [(quarter, value, source)]
 'VIN': [('Q4 2025', 100, 'Q2-25 letter: 100 MW available in 2025'), ('Q4 2026', 300, 'Q2-25 letter: rest in 2026 (300 MW site; 200 MW contracted then)')],
 'MA1': [('Q4 2025', 75, 'Oct 2024 PR: 75 MW end-2025/early-2026')],
 'KC': [('Q1 2025', 5, 'Feb 2025: 5 MW live end Q1 2025')],
 'BET': [('Q3 2026', 'first capacity', 'Reuters Feb 12 2026'), ('Q4 2026', 120, 'Reuters/DCD: ~120 MW YE26'), ('Q4 2027', 240, 'Reuters: 240 MW in 2027')],
 'BHM': [('Q4 2026', 'initial energization (MW n/d)', 'nebius.com/birminghamal')],
 'IND': [('Q3 2026', 'IPL 200 MW late Q3 2026', 'City of Independence deck Feb 2026'), ('Q4 2027', 'operational 2027', 'Q1-26 letter')],
 'BUT': [('Q4 2027', 260, 'Nebius site page: 260 MW by Oct 2027'), ('Q1 2029', 660, 'Jan 2029'), ('Q4 2029', 1000, '1,000+ MW Nov 2029')],
 'LAP': [('Q4 2027', 'first capacity 2027', 'Mar 31 2026 PR')],
 'MA2': [('Q4 2027', 'from 2027', 'DCD Aug 13 2026')],
 'HAR': [('Q4 2027', 22, 'UK 65 MW fully ramped 2027')],
 'MAS': [('Q3 2026', 22, 'Mega Or: delivery Q3 2026')],
 'BEI': [('Q1 2027', 58, 'Mega Or: staged Q3-26..Q1-27')],
 'MAD': [('Q3 2026', 18, 'LinkedIn/Merlin: live Sep 2026 (phase 1)')],
 'TAL': [('Q4 2026', 22, 'Q2-26: later in 2026')],
 'LON': [('Q4 2025', 16, 'Target Q4-25'), ('Q4 2026', 19, 'Ark Jun 2026: +3 MW in delivery')],
 'AIB': [('Q3 2027', 25, 'AIB: Hall 1 ~month 10'), ('Q4 2027', 50, 'AIB: Hall 2 ~month 14')],
}


# =====================================================================================
# V2 REVISIONS (2026-10-04/05 ET): new sites, KC history, Vineland 240 IT cap, Israel slip,
# basis/PUE per site, IT/gross/connected/billing series, tranche tags, alt path, summary.
# =====================================================================================
S.update(
 f24='https://www.sec.gov/Archives/edgar/data/1513845/000155837025005991/nbis-20241231x20f.htm',
 shl226pdf='https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf',
 gh='https://boards-api.greenhouse.io/v1/boards/nebius/jobs?content=true (accessed 2026-10-04)',
 woodbury='https://woodburynewsnet.org/7933/business/what-woodbury-residents-should-know-about-irongates-two-data-centers/',
 vantage='https://vantage-dc.com/news/vantage-data-centers-and-nebius-expand-uk-ai-infrastructure-with-first-deployment-in-south-wales-ai-growth-zone/',
 gmlon='https://www.datacenterdynamics.com/en/news/green-mountain-secures-neocloud-customer-at-data-center-in-london-uk/',
 megaor='https://www.deeptase.co.il/en/analysis/4873',
 aib10k='https://www.sec.gov/Archives/edgar/data/2070542/000121390026036810/ea0282624-10k_blockchain.htm',
 aibpr='https://www.globenewswire.com/news-release/2026/09/30/3371855/0/en/aib-data-centers-signs-contract-with-nebius-for-ai-data-center-capacity.html',
 peers='/workspace/nebius/peers/comparison.xlsx (Contracts_$MWyr, Sites tabs)',
 unseen='/workspace/nebius/capacity/unseen_capacity.md',
 altpath='/workspace/nebius/capacity/mw_path_proposed.csv',
 shut='/workspace/nebius/capacity/vineland_shutdown.xlsx',
)
_new = [
 dict(k='NEW', name='Newport UK (Vantage CWL1)', c='UK', own='Colocation (Vantage)', tag='self-serve (assumed)', tagev='assumed', cmw=None, basis='not stated (MW not disclosed)', pue='not disclosed',
      mg='Announced Aug 13 2026; part of UK 65 MW fully ramped 2027', my='Newport + UK 3rd = 65 - 22 (Harlow) = 43 MW by 2027 (Derived); Newport ~29 if LON-East is the 3rd site (low)', grade='Yellow', gen='not disclosed (B300/GB300 assumed)',
      src=f"{S['vantage']}; {S['ukpr']}; {S['gh']} (Newport jobs 4883814101 Jun 19, 4959471101 Aug 27)"),
]
SITES[:] = SITES[:[x['k'] for x in SITES].index('UK3')] + _new + SITES[[x['k'] for x in SITES].index('UK3'):]
for x in SITES:
    if x['k'] == 'UK3':
        x.update(name='UK 3rd new deployment (lead: Green Mountain LON-East, Romford)', my='Lead only: Green Mountain LON-East 14 MW to unnamed neocloud (Sep 1 2026) + Nebius East London job (Feb 2); low confidence', src=f"{S['ukpr']}; {S['gmlon']}")
    if x['k'] == 'OK':
        x.update(cmw='>100 (map category)', my='Q2-26 letter map: Sites >100MW; 8 on-site roles since Mar 2 2026; town/landlord not found', grade='Yellow/Red', src=f"{S['f20']}; {S['shl226pdf']} (p.2 map); {S['gh']}")
    if x['k'] == 'MN':
        x.update(my='us-north1 Woodbury; likely IronGate MSP01 (15 MW live, fully leased; inference)', src=f"{S['f20']}; {S['docs']}; {S['woodbury']}; {S['gh']}")
    if x['k'] == 'AIB':
        x.update(name='AIB CLT1 (Moore, Spartanburg County SC; Lockhart Power)', src=f"{S['aib']}; {S['aibpr']}; {S['aib10k']}")
    if x['k'] == 'MAS': x.update(my='Mega Or Q2-26 deck: Yoav construction completion Q4 2026 -> Q4-26/Q2-27', src=f"{S['isr']}; {S['isr2']}; {S['megaor']}")
    if x['k'] == 'BEI': x.update(my='Mega Or Q2-26 deck: completion Q2 2027 -> staged Q1-Q4 2027', src=f"{S['isr']}; {S['isr2']}; {S['megaor']}")
SK = {x['k']: x for x in SITES}
# --- anchors for previously n/d sites (Estimated; see unseen_capacity.md) ---
A['MN'] = [('Q1 2026', 5, 15, 'Estimated', 'us-north1 live (docs); DC technician "onsite in our Minnesota data center" Feb 20 2026; IronGate MSP01 15 MW live fully leased (Woodbury News Net Sep 9 2026) -> 5-15 MW (inference)')]
A['OK'] = [('Q2 2026', 0, 20, 'Estimated', 'On-site DC IT technician roles Mar 2 / May 28 2026 ("onsite in our Oklahoma data center"); scale n/d'),
           ('Q3 2026', 5, 15, 'Estimated', 'Security (Jun 23), admin services mgr (Aug 17), warehouse (Sep 15)'),
           ('Q4 2026', 10, 40, 'Estimated', 'Site manager (Sep 30) + liquid-cooling critical-infra engineer (Oct 2); map ">100MW" = eventual size'),
           ('Q1 2027', 25, 75, 'Estimated', ''), ('Q2 2027', 50, 100, 'Estimated', ''), ('Q3 2027', 75, 125, 'Estimated', ''),
           ('Q4 2027', 75, 125, 'Estimated', 'Q2-26 letter map category "Sites >100MW" (Reported category; MW n/d)')]
A['NEW'] = [('Q4 2026', 0, 10, 'Estimated', 'Jobs Jun/Aug 2026; announced Aug 13 2026'), ('Q1 2027', 5, 25, 'Estimated', ''), ('Q2 2027', 10, 35, 'Estimated', ''),
            ('Q3 2027', 15, 43, 'Estimated', 'UK 65 MW fully ramped 2027 minus Harlow 22 = 43 (Derived) is the cap for Newport+UK3'), ('Q4 2027', 20, 43, 'Estimated', '')]
A['UK3'] = [('Q2 2027', 0, 5, 'Estimated', 'Lead only (Green Mountain LON-East 14 MW)'), ('Q3 2027', 0, 12, 'Estimated', ''), ('Q4 2027', 0, 14, 'Estimated', '14 MW = LON-East neocloud lease (if Nebius; low)')]
# --- KC history recalibration: start Q1'25 at 5-15 MW (FY24 20-F ~30k GPUs at 3/31/25 needs ~55 gross MW of Hopper) ---
A['KC'] = [('Q1 2025', 5, 15, 'Estimated', 'Recalibrated v2: 5 MW phase 1 live end-Q1 2025 (Reported floor) to 15 MW; FY24 20-F: ~30,000 GPUs deployed at Mar 31 2025 implies ~55 gross MW of Hopper vs 40 MW at other sites'),
           ('Q2 2025', 5, 15, 'Estimated', 'Phase 1 5 MW (Reported floor); held flat to keep series monotonic'),
           ('Q3 2025', 5, 15, 'Estimated', 'Phase 2 added Jul 2025, MW n/d'),
           ('Q4 2025', 5, 35, 'Estimated', 'Landlord: facility 35 MW total Sep 2025; Nebius share n/d'),
           ('Q4 2027', 20, 40, 'Estimated', 'Nebius stated expandable to 40 MW')]
# --- Vineland cap at 240 IT MW (300 gross / 1.25) ---
A['VIN'] = [a if a[0] not in ('Q3 2027', 'Q4 2027') else (a[0], min(a[1], 240) if isinstance(a[1], (int, float)) else a[1], 240, a[3], a[4] + '; v2: capped at 240 MW IT (300 gross / 1.25)') for a in A['VIN']]
# --- Israel slip per Mega Or Q2-26 presentation ---
A['MAS'] = [('Q4 2026', 0, 11, 'Estimated', 'Mega Or Q2-26 deck: Yoav construction completion Q4 2026 (slip vs Q3-26 delivery)'), ('Q1 2027', 0, 22, 'Estimated', ''), ('Q2 2027', 11, 22, 'Estimated', ''), ('Q3 2027', 22, 22, 'Estimated', '')]
A['BEI'] = [('Q1 2027', 0, 20, 'Estimated', 'Mega Or Q2-26 deck: Beit Shemesh completion Q2 2027'), ('Q2 2027', 20, 39, 'Estimated', ''), ('Q3 2027', 39, 58, 'Estimated', ''), ('Q4 2027', 58, 58, 'Estimated', '')]
CONTRACT_Q['AIB'] = ('Q3 2026', 50)
MGMT['OK'] = [('Q2 2026', '>100 MW site (map)', 'Q2-26 letter map ' + S['shl226pdf'])]
MGMT['NEW'] = [('Q4 2027', '43 incl. UK3 (Derived)', 'UK PR Jun 8 2026: 3 new UK sites = 65 MW fully ramped 2027; minus Harlow 22')]
# --- per-site basis & conversion (gross/IT factor). treat: how the stated/energized figure is treated for normalization ---
BASIS = {  # k: (stated basis, evidence, treat_as, factor gross/IT, factor tag, factor basis)
 'VIN': ('IT (energized rows); contracted 300 = gross', 'Energized rows are power-bounded IT (gensets/1.2); 300 MW gross vs Bloom 328 installed + 5 grid supply (gpu_mw_capex.md)', 'IT', 1.20, 'Estimated', 'Operating PUE 1.2 for liquid-cooled GB300 (team ceiling formula); design ratio 300/240 = 1.25'),
 'IND': ('not stated', 'Utility (IPL/IPP) power figures', 'gross', 1.35, 'Estimated', 'Peer utility-to-IT design ratios 1.30-1.52 (AIB 65/50, Crusoe 900/672, Denton 394/262, Helios 800/526; peers comparison.xlsx); operating PUE ~1.2-1.3'),
 'BUT': ('not stated', 'PPL ESA = utility load', 'gross', 1.35, 'Estimated', 'as IND'),
 'BHM': ('not stated', '300 MW press figure; Alabama Power 115/230 kV', 'gross', 1.35, 'Estimated', 'as IND'),
 'BET': ('not stated', '240 MW grid connection (press)', 'gross', 1.30, 'Estimated', 'Northern France climate; peer range'),
 'LAP': ('not stated', '310 MW site/substation figures', 'gross', 1.15, 'Estimated', 'Nordic free cooling (Mantsala-class)'),
 'MA1': ('not stated', '75 MW site figure; Nebius cites "one of the world\'s leading PUE"', 'gross', 1.10, 'Estimated', 'PUE ~1.1 (company statements per site_list; 20-F gives no number)'),
 'MA2': ('not stated', '70 MW substation (Caverion)', 'gross', 1.15, 'Estimated', 'Nordic'),
 'KC':  ('not stated', 'Colo; Patmos facility MW', 'IT', 1.35, 'Estimated', 'Air-cooled Hopper/B200 retrofit; colo leases usually priced on critical IT'),
 'PAR': ('not stated', 'Colo; 5 MW third-party', 'IT', 1.35, 'Estimated', 'Equinix urban colo'),
 'KEF': ('not stated', 'Colo (Verne)', 'IT', 1.15, 'Estimated', 'Iceland free cooling'),
 'LON': ('gross (Derived)', '7,000 B300 on 19 MW = 368/MW fits gross density (gpu_mw_capex.xlsx Nebius checks)', 'gross', 1.30, 'Estimated', 'Air-cooled HGX B300'),
 'HAR': ('not stated', 'Colo (Kao Data)', 'IT', 1.20, 'Estimated', 'Kao Data design PUE ~1.2 (Estimated)'),
 'NEW': ('not stated', 'MW not disclosed', 'IT', 1.25, 'Estimated', 'UK colo'),
 'UK3': ('not stated', 'Lead only', 'IT', 1.25, 'Estimated', 'UK colo'),
 'MOD': ('IT (Derived)', '4,000 B200 / 8 MW = 500/MW fits IT density', 'IT', 1.30, 'Estimated', 'Israel climate'),
 'MAS': ('not stated', 'Mega DC lease MW', 'IT', 1.30, 'Estimated', 'Israel climate'),
 'BEI': ('not stated', 'Mega DC lease MW', 'IT', 1.30, 'Estimated', 'Israel climate'),
 'MAD': ('not stated', 'Merlin 18 MW contract', 'IT', 1.30, 'Estimated', 'Madrid climate'),
 'TAL': ('not stated', 'Greenergy 22 MW first phase', 'IT', 1.20, 'Estimated', 'Baltic climate'),
 'MN':  ('not stated', 'IronGate MW (inference)', 'IT', 1.30, 'Estimated', 'US colo'),
 'OK':  ('not stated', 'Map category >100MW', 'IT', 1.30, 'Estimated', 'US colo, liquid cooling'),
 'FR2': ('not stated', 'n/d', 'IT', 1.30, 'Estimated', 'n/d'),
 'AIB': ('IT (Reported)', '"50 MW of critical IT capacity" on 65 MW utility ESA (AIB PR Sep 30 2026)', 'IT', 1.30, 'Derived', '65/50 = 1.30'),
}
OWN = {k: ('owned' if k in ('IND', 'BUT', 'BHM', 'LAP', 'MA1', 'MA2') else 'leased') for k in BASIS}

def ffill(key):
    """Return per-quarter list of (low, high, tag, src) with forward fill; 0 before first anchor."""
    out = [(0, 0, 'Derived', 'Not yet energized (0 before first anchor)')] * len(Q)
    anchors = sorted(A[key], key=lambda a: QI[a[0]])
    for j, (q, lo, hi, tag, src) in enumerate(anchors):
        end = QI[anchors[j + 1][0]] if j + 1 < len(anchors) else len(Q)
        for i in range(QI[q], end):
            t = tag if i == QI[q] else ('Estimated' if i > HIST_END else tag)
            s = src if i == QI[q] else f'carried forward from {q}'
            out[i] = (lo, hi, t, s)
    return out

# ---------------- workbook ----------------
wb = openpyxl.Workbook()
LOG = []  # (sheet, cell, tag, source, note)
def put(ws, r, c, v, tag=None, src=None, inp=None, note=''):
    cell = ws.cell(r, c, v)
    isf = isinstance(v, str) and v.startswith('=')
    if inp is None: inp = not isf
    if v is not None and not (isinstance(v, str) and not isf and tag is None):
        cell.font = BLUE if (inp and not isf) else BLACK
    if tag:
        LOG.append((ws.title, cell.coordinate, v if not isf else 'formula: ' + v, tag, src or '', note))
    return cell

def header(ws, r, vals):
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.font = BOLD; c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical='top')

# ===== README (filled at end, created first for tab order) =====
wsR = wb.active; wsR.title = 'README'
wsS = wb.create_sheet('Sites'); wsQ = wb.create_sheet('Quarterly MW'); wsG = wb.create_sheet('GPUs deployed')
wsV = wb.create_sheet('GPU vintage'); wsT = wb.create_sheet('Vineland tranche track'); wsL = wb.create_sheet('Tags_Sources')

# ===== Sites =====
cols = ['Site', 'Country', 'Ownership type', 'Contract tag', 'Contract tag evidence', 'Contracted MW', 'MW basis (IT/gross/unknown)', 'PUE (where disclosed)',
        'Mgmt go-live', 'My evidence-based go-live', 'Risk grade', 'GPU generation', 'YE26 energized low', 'YE26 energized high', 'YE26 energized base', 'Key sources']
header(wsS, 1, cols)
SITEROW = {}
for i, s in enumerate(SITES, 2):
    SITEROW[s['k']] = i
    put(wsS, i, 1, s['name']); put(wsS, i, 2, s['c']); put(wsS, i, 3, s['own'], 'Reported', s['src'])
    put(wsS, i, 4, s['tag'], 'Reported' if 'Reported' in s['tagev'] else ('Estimated' if 'press' in s['tag'] else 'Estimated'), s['tagev'])
    put(wsS, i, 5, s['tagev'])
    put(wsS, i, 6, s['cmw'] if s['cmw'] is not None else 'not disclosed', 'Reported' if s['k'] not in ('BHM',) else 'Reported (press)', s['src'])
    put(wsS, i, 7, s['basis'], 'Estimated', S['gpumw'] + ' sec.4; ' + S['sitelist'])
    put(wsS, i, 8, s['pue'], 'Reported' if 'Nebius' in s['pue'] else ('Derived' if 'Derived' in s['pue'] else 'Reported'), s['src'])
    put(wsS, i, 9, s['mg'], 'Reported', s['src']); put(wsS, i, 10, s['my'], 'Estimated', s['src'])
    c = put(wsS, i, 11, s['grade'], 'Estimated', 'see site_verification.md'); c.fill = PatternFill('solid', fgColor=GRADEFILL[s['grade']])
    put(wsS, i, 12, s['gen'], 'Estimated' if 'assum' in s['gen'] or 'Estimated' in s['gen'] or 'inference' in s['gen'] else 'Reported', s['src'])
    put(wsS, i, 16, s['src'])
for j, w in enumerate([34, 9, 30, 26, 40, 10, 30, 18, 40, 40, 10, 30, 10, 10, 10, 80], 1):
    wsS.column_dimensions[CL(j)].width = w

# ===== Quarterly MW =====
QC0 = 5  # first quarter column (E)
header(wsQ, 1, ['Site', 'Block', 'Basis', 'Tag (hist / proj)'] + Q)
wsQ.cell(2, 1, "History Q1'24-Q2'26 (tags per cell in Tags_Sources); projections Q3'26-Q4'30 = Estimated. Base = midpoint of low/high (formula). Blue = input, black = formula.").font = Font(italic=True)
r = 3
ROWS = {}
for s in SITES:
    k = s['k']; ff = ffill(k) if A[k] else None
    basis = 'IT' if s.get('basis_it') else 'stated/unknown'
    blocks = ['Contracted MW', 'Energized MW low', 'Energized MW high', 'Energized MW base (=midpoint)', 'Mgmt figure']
    ROWS[k] = {}
    for b in blocks:
        put(wsQ, r, 1, s['name']); put(wsQ, r, 2, b); put(wsQ, r, 3, basis if 'Energized' in b else ('gross/unknown' if b == 'Contracted MW' else 'as stated'))
        ROWS[k][b] = r
        for i, q in enumerate(Q):
            col = QC0 + i
            if b == 'Contracted MW':
                if k in CONTRACT_Q:
                    q0, v = CONTRACT_Q[k]
                    for qs, vs in CONTRACT_STEP.get(k, []):
                        if i >= QI[qs]: v = vs
                    if i >= QI[q0]:
                        put(wsQ, r, col, v, 'Reported' if i <= HIST_END else 'Reported (carried)', s['src'])
                else:
                    put(wsQ, r, col, 'n/d', 'Reported', 'MW not disclosed')
            elif b in ('Energized MW low', 'Energized MW high'):
                if ff is None:
                    put(wsQ, r, col, 'n/d', 'Reported', 'MW not disclosed -> captured in plug row')
                else:
                    lo, hi, tag, src = ff[i]
                    v = lo if b.endswith('low') else hi
                    put(wsQ, r, col, v, tag, src)
            elif b.startswith('Energized MW base'):
                if ff is None:
                    put(wsQ, r, col, 'n/d', 'Reported', 'MW not disclosed')
                else:
                    L = f'{CL(col)}{ROWS[k]["Energized MW low"]}'; H = f'{CL(col)}{ROWS[k]["Energized MW high"]}'
                    put(wsQ, r, col, f'=({L}+{H})/2', 'Derived', 'midpoint of low and high')
            elif b == 'Mgmt figure':
                for (mq, mv, msrc) in MGMT.get(k, []):
                    if mq == q: put(wsQ, r, col, mv, 'Reported', msrc + ' | ' + s['src'][:200])
        r += 1
    r += 0
# company rows
r += 1
CO = {}
def sumrow(label, block, note):
    global r
    put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, label); put(wsQ, r, 3, 'mixed (see README)')
    for i in range(len(Q)):
        col = QC0 + i; refs = ','.join(f'{CL(col)}{ROWS[s["k"]][block]}' for s in SITES)
        put(wsQ, r, col, f'=SUM({refs})', 'Derived', note)
    CO[label] = r; r += 1
sumrow('Identified sites - energized low', 'Energized MW low', 'sum of site low rows (n/d excluded)')
sumrow('Identified sites - energized high', 'Energized MW high', 'sum of site high rows')
sumrow('Identified sites - energized base', 'Energized MW base (=midpoint)', 'sum of site base rows')
sumrow('Identified sites - contracted MW', 'Contracted MW', 'sum of site contracted MW (named sites only)')
# reported/guide row
put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, 'Mgmt reported / guided MW (power, company)'); put(wsQ, r, 3, 'as stated')
REP = {'Q4 2024': (25, 'Reported', 'Connected MW YE24 25 MW (Q3-25 letter chart) ' + S['q325']),
       'Q4 2025': (170, 'Reported', '~170 MW ACTIVE power YE25 (Q4-25 letter) ' + S['q425']),
       'Q4 2026': (800, 'Reported (guide low end)', '800 MW-1 GW CONNECTED by YE26 (Q3-25/Q4-25/Q1-26 letters; Q2-26 call) ' + S['q126']),
       'Q4 2027': ('=800+1000', 'Derived', 'YE26 guide low (800) + ">1 GW per year starting 2027" (Q2-26 letter) ' + S['q226'])}
for q, (v, t, src) in REP.items():
    put(wsQ, r, QC0 + QI[q], v, t, src)
CO['rep'] = r; r += 1
put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, 'Plug: undisclosed sites (MN, OK, France 2nd, UK 3rd, unattributed) = reported - identified base'); put(wsQ, r, 3, 'mixed')
for q in REP:
    col = CL(QC0 + QI[q])
    put(wsQ, r, QC0 + QI[q], f'={col}{CO["rep"]}-{col}{CO["Identified sites - energized base"]}', 'Derived', 'plug so that identified base + plug = mgmt reported/guided figure; NOT an estimate of what exists')
CO['plug'] = r; r += 1
put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, 'Total incl. plug (ties to mgmt where reported)'); put(wsQ, r, 3, 'mixed')
for i in range(len(Q)):
    col = CL(QC0 + i)
    put(wsQ, r, QC0 + i, f'={col}{CO["Identified sites - energized base"]}+N({col}{CO["plug"]})', 'Derived', 'identified base + plug')
CO['tot'] = r; r += 1
put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, 'Contracted power reported/target (company)'); put(wsQ, r, 3, 'gross')
for q, v, t, src in [('Q4 2025', 2000, 'Reported', '>2 GW secured (Feb 12 2026) ' + S['q425']), ('Q1 2026', 3500, 'Reported', '>3.5 GW (May 13 2026) ' + S['q126']),
                     ('Q4 2026', 5000, 'Reported (target)', '5 GW YE26 target (Aug 12 2026) ' + S['q226'])]:
    put(wsQ, r, QC0 + QI[q], v, t, src)
CO['contr'] = r; r += 1
put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, 'Unattributed contracted MW (company - named sites)'); put(wsQ, r, 3, 'gross')
for q in ['Q4 2025', 'Q1 2026', 'Q4 2026']:
    col = CL(QC0 + QI[q]); put(wsQ, r, QC0 + QI[q], f'={col}{CO["contr"]}-{col}{CO["Identified sites - contracted MW"]}', 'Derived', 'company contracted - sum of named-site contracted MW')
r += 1
wsQ.column_dimensions['A'].width = 36; wsQ.column_dimensions['B'].width = 34; wsQ.column_dimensions['C'].width = 14; wsQ.column_dimensions['D'].width = 4
wsQ.freeze_panes = 'E3'
YE26 = CL(QC0 + QI['Q4 2026'])
for s in SITES:
    rr = SITEROW[s['k']]
    for cc, b in [(13, 'Energized MW low'), (14, 'Energized MW high'), (15, 'Energized MW base (=midpoint)')]:
        put(wsS, rr, cc, f"='Quarterly MW'!{YE26}{ROWS[s['k']][b]}", 'Derived', 'link to Quarterly MW Q4 2026')

# ===== GPUs deployed =====
header(wsG, 1, ['Generation', 'GPUs per IT MW (base)', 'GPUs per gross MW (base)', 'Source cell', 'Tag'])
GROW = {}
for i, (g, (row, *_)) in enumerate(GEN.items(), 2):
    GROW[g] = i
    put(wsG, i, 1, g)
    put(wsG, i, 2, GPM[g][0], 'Estimated', f"{S['gpumw']} GPUsPerMW!I{row} (base row; value reproduced from its formula)")
    put(wsG, i, 3, GPM[g][1], 'Estimated', f"{S['gpumw']} GPUsPerMW!L{row} (base row)" + ('; VR values from gpu_mw_capex.md table' if 'Rubin' in g else ''))
    put(wsG, i, 4, f"GPUsPerMW!I{row} / L{row}"); put(wsG, i, 5, 'Estimated')
wsG.cell(8, 1, 'RTX PRO 6000 not modeled in gpu_mw_capex.xlsx -> Madrid phase 1 / KC RTX share treated as B300/B200 (Estimated).').font = Font(italic=True)
GEN_OF = {'VIN': 'GB300 NVL72', 'IND': 'Vera Rubin NVL72', 'BUT': 'Vera Rubin NVL72', 'BHM': 'Vera Rubin NVL72', 'BET': 'GB300 NVL72', 'LAP': 'Vera Rubin NVL72',
          'MA2': 'Vera Rubin NVL72', 'PAR': 'Hopper H100/H200', 'KEF': 'Hopper H100/H200', 'LON': 'B300 (HGX)', 'HAR': 'B300 (HGX)', 'MOD': 'B200 (HGX)',
          'MAS': 'GB300 NVL72', 'BEI': 'GB300 NVL72', 'MAD': 'B300 (HGX)', 'TAL': 'GB300 NVL72', 'AIB': 'Vera Rubin NVL72',
          'MN': 'B300 (HGX)', 'OK': 'GB300 NVL72', 'NEW': 'GB300 NVL72', 'UK3': 'GB300 NVL72'}
gr = 10
header(wsG, gr, ['Site', 'Generation', 'MW basis col', 'Note'] + Q); gr += 1
GPU_ROWS = []
def gref(g, it): return f"${'B' if it else 'C'}${GROW[g]}"
for s in SITES:
    k = s['k']
    if not A[k]: continue
    it = bool(s.get('basis_it'))
    put(wsG, gr, 1, s['name']); put(wsG, gr, 3, 'IT' if it else 'gross')
    base_r = ROWS[k]['Energized MW base (=midpoint)']
    for i in range(len(Q)):
        col = QC0 + i; m = f"'Quarterly MW'!{CL(col)}{base_r}"
        if k == 'MA1':
            f = f"=ROUND(MIN({m},25)*{gref('Hopper H100/H200',0)}+MAX({m}-25,0)*{gref('GB300 NVL72',0)},0)"; gname = 'Hopper first 25 MW + GB300 above'
        elif k == 'KC':
            f = f"=ROUND(MIN({m},10)*{gref('Hopper H100/H200',0)}+MAX({m}-10,0)*{gref('B200 (HGX)',0)},0)"; gname = 'H200 first 10 MW + B200 above (v2)'
        else:
            gname = GEN_OF[k]; f = f"=ROUND({m}*{gref(gname,it)},0)"
        put(wsG, gr, col, f, 'Derived', 'energized base MW x GPUs/MW (gpu_mw_capex.xlsx base); generation assumed per Sites tab')
    put(wsG, gr, 2, gname)
    GPU_ROWS.append(gr); gr += 1
put(wsG, gr, 1, 'TOTAL modeled GPUs (identified sites)')
for i in range(len(Q)):
    col = QC0 + i; put(wsG, gr, col, f"=SUM({CL(col)}{GPU_ROWS[0]}:{CL(col)}{GPU_ROWS[-1]})", 'Derived', 'sum')
gtot = gr; gr += 1
put(wsG, gr, 1, 'Reported GPU counts (company)'); put(wsG, gr, 4, 'GPU counts have run AHEAD of powered MW (delivered-not-powered); do not back MW out of them')
for q, v, t, src in [('Q3 2024', 10000, 'Estimated', 'site_list KPI pivot: ~10k (">20k by YE24" / "nearly doubled"); no primary source found'),
                     ('Q4 2024', 20000, 'Estimated', '">20,000 GPUs by YE2024" target (Q3-24 call); Q4-24 release: Hopper base "nearly doubled QoQ" (Reported, qualitative)'),
                     ('Q1 2025', 30000, 'Reported', 'FY24 20-F: "As of March 31, 2025, we have deployed approximately 30,000 GPUs ... primarily NVIDIA H200s" ' + S['f24'] + ' (v2: replaces ~38k call figure)')]:
    put(wsG, gr, QC0 + QI[q], v, t, src)
gr += 1
put(wsG, gr, 1, "Reported site GPU counts"); put(wsG, gr, 4, "Modi'in 4,000 HGX B200 (Oct 2025); Longcross 4,000 B300 phase 1 (+3,000 phase 2, timing n/d); Mantsala 'up to 60,000' = potential, not deployed")
put(wsG, gr, QC0 + QI['Q4 2025'], '=4000+4000', 'Reported', "Modi'in 4,000 B200 + Longcross 4,000 B300 " + S['modiin'] + '; ' + S['lon'])
wsG.column_dimensions['A'].width = 34; wsG.column_dimensions['B'].width = 26; wsG.column_dimensions['D'].width = 30

# ===== GPU vintage =====
header(wsV, 1, ['Site', 'Row type', 'Generation (assumed)', 'Note'] + Q)
vr = 2
for s in SITES:
    k = s['k']
    if not A[k]: continue
    base_r = ROWS[k]['Energized MW base (=midpoint)']
    gname = s['gen'] if k in ('MA1', 'KC') else GEN_OF[k]
    put(wsV, vr, 1, s['name']); put(wsV, vr, 2, 'Generation in service'); put(wsV, vr, 3, gname)
    for i in range(len(Q)):
        col = QC0 + i; m = f"'Quarterly MW'!{CL(col)}{base_r}"
        if k == 'MA1':
            f = f'=IF({m}>25,"Hopper+GB300",IF({m}>0,"Hopper",""))'
        elif k == 'KC':
            f = f'=IF({m}>10,"H200+B200",IF({m}>0,"H200",""))'
        else:
            f = f'=IF({m}>0,"{gname}","")'
        put(wsV, vr, col, f, 'Derived', 'generation label when energized base > 0')
    vr += 1
    put(wsV, vr, 1, s['name']); put(wsV, vr, 2, 'MW added this quarter (vintage cohort)'); put(wsV, vr, 3, gname)
    for i in range(len(Q)):
        col = QC0 + i; m = f"'Quarterly MW'!{CL(col)}{base_r}"
        prev = f"'Quarterly MW'!{CL(col-1)}{base_r}" if i > 0 else '0'
        put(wsV, vr, col, f'={m}-{prev}', 'Derived', 'delta of energized base MW = in-service cohort for vintage depreciation (5-yr life from Jan 2026 per 20-F)')
    vr += 1
wsV.column_dimensions['A'].width = 34; wsV.column_dimensions['B'].width = 30; wsV.column_dimensions['C'].width = 26

# ===== Vineland tranche track =====
T = wsT
T.cell(1, 1, 'Vineland (Microsoft) tranche track - 9 tranches, $17,392.9M take-or-pay (20-F), ~$6,958.1M upfront, billing through Oct 2031').font = BOLD
inputs = [('Tranche count', 9, 'Reported', S['f20'] + ' Note 1 F-17'),
          ('Contract value $M', 17392.9, 'Reported', S['f20']),
          ('Upfront payments $M', 6958.1, 'Reported', S['f20']),
          ('Site contracted MW (gross)', 300, 'Reported', S['kef']),
          ('IT MW at full build (mid of 230-250, Estimated)', 240, 'Estimated', S['gpumw'] + ' / gpu_mw_capex.md sec.4'),
          ('Genset count', 62, 'Reported', S['enr'] + ' (NJDEP)'),
          ('Genset MW each', 1.982, 'Reported', S['enr']),
          ('Grid MW', 5, 'Reported (secondary)', S['alt']),
          ('PUE for ceiling', 1.2, 'Estimated', S['alt']),
          ('Q2-26 group revenue $M', 582.3, 'Reported', S['filings']),
          ('Top customer share Q2-26 (customer C)', 0.24, 'Reported', S['filings'] + ' (identity not disclosed)')]
IR = {}
for i, (lab, v, t, src) in enumerate(inputs, 3):
    put(T, i, 1, lab); put(T, i, 2, v, t, src); IR[lab] = f'$B${i}'
dr = 3 + len(inputs) + 1
derived = [('MW per tranche, gross (EQUAL-SPLIT ASSUMPTION)', f"={IR['Site contracted MW (gross)']}/{IR['Tranche count']}"),
           ('MW per tranche, IT (equal split assumption)', f"={IR['IT MW at full build (mid of 230-250, Estimated)']}/{IR['Tranche count']}"),
           ('Generator ceiling, gross MW', f"={IR['Genset count']}*{IR['Genset MW each']}+{IR['Grid MW']}"),
           ('Generator ceiling, IT MW', f"=ROUND(({IR['Genset count']}*{IR['Genset MW each']}+{IR['Grid MW']})/{IR['PUE for ceiling']},1)"),
           ('Max tranches supportable on gensets (IT basis, equal split)', f"=ROUNDDOWN(B{dr+3}/B{dr+1},1)"),
           ('TCV per tranche $M (equal split)', f"={IR['Contract value $M']}/{IR['Tranche count']}"),
           ('Quarterly revenue per tranche $M (TCV/tranche / 20 qtrs)', f"=B{dr+5}/20"),
           ('Full run-rate $M/qtr (TCV/20)', f"={IR['Contract value $M']}/20"),
           ('Q2-26 implied top-customer revenue $M', f"={IR['Q2-26 group revenue $M']}*{IR['Top customer share Q2-26 (customer C)']}"),
           ('Q2-26 equal-split expectation for T1+T2 $M', f"=2*B{dr+6}"),
           ('Q2-26 top customer as % of full MSFT run-rate', f"=B{dr+8}/B{dr+7}")]
for j, (lab, f) in enumerate(derived):
    put(T, dr + j, 1, lab); put(T, dr + j, 2, f, 'Derived', lab)
tr = dr + len(derived) + 2
header(T, tr, ['Tranche', 'Delivery (reported)', 'Physical evidence / date', 'Power source', 'Est. MW gross (equal split, assumption)', 'Est. MW IT (equal split)', 'Cumulative IT MW', 'Within genset ceiling?',
               'Est. billing start - Low (Bear)', 'Est. billing start - Base', 'Est. billing start - High (Bull)', 'Risk type', 'Tag', 'Source'])
TR = [
 ('T1', 'Nov 2025 (20-F)', 'Oct 2025: 16 gensets (31.7 MW) start operating (DataOne letter to NJDEP)', 'Rented Cat G3516H gensets', 'Q4 2025', 'Q4 2025', 'Q4 2025', 'Delivered; SLA/availability exposure if gensets must stop (NJDEP ~Nov 2026 deadline)', 'Reported (delivery); Estimated (MW)', S['f20'] + '; ' + S['flood']),
 ('T2', 'Feb 2026 (20-F)', 'Feb 2026: +16 gensets (32 total, 63.4 MW); city report "Phase 1 generators are online"', 'Rented gensets', 'Q1 2026', 'Q1 2026', 'Q1 2026', 'Delivered; same SLA exposure', 'Reported (delivery); Estimated (MW)', S['f20'] + '; ' + S['flood']),
 ('T3', 'Jul-Aug 2026 ("two tranches over the past month", Q2-26 letter Aug 12)', 'Jun 2026: all 62 gensets fired (122.9 MW); >=45 running Aug (thermal drone); DataOne: >=3 of 14 halls operating early Aug', 'Rented gensets', 'Q3 2026', 'Q3 2026', 'Q3 2026', 'Delivered; same SLA exposure', 'Reported (delivery); Estimated (MW, quarter)', S['q226'] + '; ' + S['flood2'] + '; ' + S['alt']),
 ('T4', 'Jul-Aug 2026 (as T3)', 'As T3. Cum. 4 x 26.7 IT = 107 MW = genset ceiling: no headroom', 'Rented gensets', 'Q3 2026', 'Q3 2026', 'Q3 2026', 'Delivered; at ceiling - SLA risk if any units fail/curtailed', 'Reported (delivery); Estimated (MW)', S['q226'] + '; ' + S['enr']),
 ('T5', 'Not yet ("continue to deliver the remaining tranches", Q2-26 call)', 'Bloom units being placed (Sep 8 2026 imagery, low conf.); stop-construction order Aug 10 (lift not publicly confirmed); site plan approved Aug 17', 'Bloom fuel cells (328 MW installed / ~250 MW guaranteed)', 'Q3 2027', 'Q2 2027', 'Q1 2027', 'TERMINATION/SLA risk (delivery after "during 2025 and 2026" window + grace period), not demand risk', 'Estimated', S['q226call'] + '; ' + S['hntr'] + '; ' + S['whyy']),
 ('T6', 'Not yet', 'As T5', 'Bloom', 'Q3 2027', 'Q2 2027', 'Q1 2027', 'Termination/SLA risk', 'Estimated', S['alt']),
 ('T7', 'Not yet', 'As T5', 'Bloom', 'Q4 2027', 'Q3 2027', 'Q2 2027', 'Termination/SLA risk', 'Estimated', S['alt']),
 ('T8', 'Not yet', 'As T5', 'Bloom', 'Q4 2027', 'Q3 2027', 'Q2 2027', 'Termination/SLA risk', 'Estimated', S['alt']),
 ('T9', 'Not yet', 'As T5', 'Bloom', 'Q1 2028', 'Q3 2027', 'Q2 2027', 'Termination/SLA risk', 'Estimated', S['alt']),
]
for j, row in enumerate(TR):
    rr = tr + 1 + j
    name, deliv, ev, pw, bl, bb, bh, risk, tag, src = row
    put(T, rr, 1, name); put(T, rr, 2, deliv, 'Reported' if 'Not' not in deliv else 'Reported (call)', src); put(T, rr, 3, ev, tag, src); put(T, rr, 4, pw, 'Reported', src)
    put(T, rr, 5, f'=$B${dr}', 'Derived', 'equal split assumption (300/9)'); put(T, rr, 6, f'=$B${dr+1}', 'Derived', 'equal split assumption (240/9)')
    put(T, rr, 7, f'=SUM($F${tr+1}:F{rr})', 'Derived', 'cumulative')
    put(T, rr, 8, f'=IF(G{rr}<=$B${dr+3}+0.5,"yes","NO - needs Bloom")', 'Derived', 'vs generator ceiling IT')
    put(T, rr, 9, bl, 'Estimated', 'Billing = delivery for T1-T4 (contract invoiced monthly per tranche term); T5-T9 = Bloom energization + >=1 qtr commissioning lag'); put(T, rr, 10, bb, 'Estimated', 'as col I'); put(T, rr, 11, bh, 'Estimated', 'as col I')
    put(T, rr, 12, risk); put(T, rr, 13, tag); put(T, rr, 14, src)
nr = tr + 12
notes = [
 'CONCLUSION: The Q2-26 letter says "all capacity tranches" delivered, but the call says "all the tranches that we were required to deliver ... up to date" and "continue to deliver the remaining tranches". Read: >=4 of 9 delivered (T1 Nov-25, T2 Feb-26, two in Jul/Aug-26); 5 remain.',
 'Physical cap: 62 rented gensets (122.9 MW) + ~5 MW grid = ~128 MW gross / ~107 MW IT ceiling. That supports ~4 equal tranches (4 x 26.7 IT = 107). T5-T9 (~5/9 of TCV) need Bloom fuel cells, which were stop-ordered Aug 10 2026 and not yet confirmed energized.',
 'Team hypothesis ("delivered = handed over; last tranche billing ~Q4-26") is NOT supported: the full run-rate needs Bloom; base case full billing Q3 2027.',
 'Q2-26 revenue check: top customer 24% x $582.3M = ~$140M vs equal-split T1+T2 = ~$193M. Either T1/T2 are smaller than 1/9 each (consistent with 16 gensets per tranche, ~27 IT MW) or customer C is not Microsoft (not disclosed).',
 'Dated risk: NJDEP AO ($1.07M, ~Sep 22 2026) gives 45 days to file for and obtain air permits or stop running the generators -> ~early/mid Nov 2026. If the gensets stop before Bloom is live, T1-T4 hit availability failures -> SLA credits / tranche-termination rights (20-F). This is a delivery/termination/SLA risk, NOT demand risk.',
 'Billing lag: T1-T4 billing assumed to start at delivery (reported dates). T5-T9 billing = Bloom energization + >=1 quarter (Estimated).',
]
for j, n in enumerate(notes):
    T.cell(nr + j, 1, n).alignment = Alignment(wrap_text=False)
T.column_dimensions['A'].width = 52; T.column_dimensions['B'].width = 22; T.column_dimensions['C'].width = 60
for c in 'DEFGHIJKLMN': T.column_dimensions[c].width = 16


# =====================================================================================
# V2 ADDITIONS: Basis tab, IT/gross/connected/billing rows, Bloom-slip, tranche tags, alt path, summary
# =====================================================================================
import csv as _csv
wsB = wb.create_sheet('Basis'); wsC = wb.create_sheet('Capacity tranches'); wsA = wb.create_sheet('Alt path'); wsU = wb.create_sheet('Summary')
# ---- Basis tab ----
header(wsB, 1, ['Site', 'Key', 'Ownership', 'Stated MW basis (IT / gross / not stated)', 'Evidence for basis', 'Treated as (for normalization)', 'Gross/IT factor (PUE-based)', 'Factor tag', 'Factor basis / PUE assumption', 'Source'])
BR = {}
for i, x in enumerate(SITES, 2):
    k = x['k']; b = BASIS[k]; BR[k] = i
    put(wsB, i, 1, x['name']); put(wsB, i, 2, k); put(wsB, i, 3, OWN[k], 'Reported', x['src'][:300])
    put(wsB, i, 4, b[0], 'Reported' if 'Reported' in b[0] else ('Derived' if 'Derived' in b[0] or 'energized' in b[0] else 'Reported'), x['src'][:300])
    put(wsB, i, 5, b[1]); put(wsB, i, 6, b[2], 'Estimated', 'colo/lease figures treated as critical IT unless evidence says gross; owned utility figures treated as gross')
    put(wsB, i, 7, b[3], b[4], b[5] + ' | ' + S['peers']); put(wsB, i, 8, b[4]); put(wsB, i, 9, b[5]); put(wsB, i, 10, S['peers'] if b[4] != 'Derived' else S['aibpr'])
n = len(SITES) + 3
wsB.cell(n, 1, 'Peer conversion evidence (peers/comparison.xlsx): AIB CLT1 65 MW utility / 50 MW IT = 1.30; Crusoe-Microsoft Abilene 900 gross / 672 IT = 1.34; Crusoe-Oracle Abilene 1.2 GW / ~843 IT = 1.42; Vantage Port Washington ~1.3 GW / ~902 IT = 1.44; Core Scientific Denton 394 grid / 262 IT = 1.50; Galaxy Helios 800 / 526 IT = 1.52. These are DESIGN (utility-to-IT) ratios incl. headroom; operating PUE for liquid-cooled builds is lower (~1.15-1.3). Factors here sit between the two (Estimated).').alignment = Alignment(wrap_text=True)
wsB.cell(n + 1, 1, 'IT-normalized MW = energized MW / factor if treated as gross, else as stated. Gross MW = energized MW x factor if treated as IT, else as stated. n/d sites excluded.')
for j, w in enumerate([40, 6, 10, 34, 60, 14, 12, 10, 60, 50], 1): wsB.column_dimensions[CL(j)].width = w
def fac(k): return f"Basis!$G${BR[k]}"
def norm_terms(block, mode, i):
    col = CL(QC0 + i); t = []
    for x in SITES:
        k = x['k']; ref = f"N({col}{ROWS[k][block]})"
        if mode == 'IT': t.append(ref if BASIS[k][2] == 'IT' else f"{ref}/{fac(k)}")
        else: t.append(ref if BASIS[k][2] == 'gross' else f"{ref}*{fac(k)}")
    return '=' + '+'.join(t)
r += 1
put(wsQ, r, 1, 'V2 NORMALIZED SERIES (see Basis tab)').font = BOLD; r += 1
for mode in ('IT', 'gross'):
    for lab, block in [('low', 'Energized MW low'), ('base', 'Energized MW base (=midpoint)'), ('high', 'Energized MW high')]:
        put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, f'Energized MW {lab} - {mode}-normalized'); put(wsQ, r, 3, mode)
        for i in range(len(Q)):
            put(wsQ, r, QC0 + i, norm_terms(block, mode, i), 'Derived', f'sum of site {block} converted to {mode} basis with Basis tab factors')
        CO[f'{mode}_{lab}'] = r; r += 1
# connected MW line (Reported history + Estimated bottom-up Q3'26-Q4'27, from mw_path_proposed.csv / unseen_capacity.md)
CONN = {'low': [295, 561, 626.4, 763.4, 990, 1249], 'base': [353, 941, 1078, 1243, 1382, 1882], 'high': [412, 1195, 1322, 1577, 2017, 2517]}
for lab in ('low', 'base', 'high'):
    put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, f'Connected MW {lab} (built & equipped; gross/shell basis)'); put(wsQ, r, 3, 'gross/shell')
    put(wsQ, r, QC0 + QI['Q4 2024'], 25, 'Reported', 'YE24 connected 25 MW (Q3-25 letter chart) ' + S['q325'])
    put(wsQ, r, QC0 + QI['Q4 2025'], 220, 'Reported (target)', 'Q3-25 letter: on track for 220 MW connected by YE25 (actual connected not restated) ' + S['q325'])
    for j, q in enumerate(['Q3 2026', 'Q4 2026', 'Q1 2027', 'Q2 2027', 'Q3 2027', 'Q4 2027']):
        put(wsQ, r, QC0 + QI[q], CONN[lab][j], 'Estimated', 'bottom-up connected by site (Vineland 300 gross shell, Independence IPL, Birmingham 115 kV, Bethune 120, OK >100 ...) ' + S['altpath'] + '; ' + S['unseen'])
    for q in Q[QI['Q1 2028']:]:
        put(wsQ, r, QC0 + QI[q], 'n/d', 'Reported', 'no estimate beyond Q4 2027')
    CO[f'conn_{lab}'] = r; r += 1
put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, 'Mgmt connected guide (low-high)'); put(wsQ, r, 3, 'as stated')
put(wsQ, r, QC0 + QI['Q4 2026'], '800-1,000', 'Reported', '800 MW-1 GW connected YE26 (Q2-26 call) ' + S['q226call']); r += 1
# Bloom slip scenario row (Vineland)
vb = ROWS['VIN']['Energized MW base (=midpoint)']
put(wsQ, r, 1, 'Vineland NJ (DataOne)'); put(wsQ, r, 2, 'SCENARIO: Bloom slips 2 quarters - energized IT'); put(wsQ, r, 3, 'IT')
for i in range(len(Q)):
    col = CL(QC0 + i)
    if i < QI['Q1 2027']: f = f'={col}{vb}'
    else: f = f'=MAX(${CL(QC0+QI["Q4 2026"])}${vb},{CL(QC0+i-2)}{vb})'
    put(wsQ, r, QC0 + i, f, 'Derived', 'base path shifted +2 quarters from Q1 2027, floored at Q4 2026 genset level (assumes gensets stay permitted)')
CO['vin_slip'] = r; r += 1
put(wsQ, r, 1, 'Vineland NJ (DataOne)'); put(wsQ, r, 2, 'SCENARIO: Bloom slips 2Q - IT MW lost vs base'); put(wsQ, r, 3, 'IT')
for i in range(len(Q)):
    col = CL(QC0 + i); put(wsQ, r, QC0 + i, f'={col}{vb}-{col}{CO["vin_slip"]}', 'Derived', 'base minus slip scenario')
CO['vin_slip_d'] = r; r += 1
# Billing MW row (IT): Vineland billed tranches (base billing quarter) + other sites' IT-normalized base lagged 1 quarter
TRBILL = {'T1': 'Q4 2025', 'T2': 'Q1 2026', 'T3': 'Q3 2026', 'T4': 'Q3 2026', 'T5': 'Q2 2027', 'T6': 'Q2 2027', 'T7': 'Q3 2027', 'T8': 'Q3 2027', 'T9': 'Q3 2027'}
TRSLIP = {'T1': 'Q4 2025', 'T2': 'Q1 2026', 'T3': 'Q3 2026', 'T4': 'Q3 2026', 'T5': 'Q4 2027', 'T6': 'Q4 2027', 'T7': 'Q1 2028', 'T8': 'Q1 2028', 'T9': 'Q1 2028'}
trIT = f"'Vineland tranche track'!$B${dr+1}"
for lab, sched in [('base', TRBILL), ('Bloom slips 2Q', TRSLIP)]:
    put(wsQ, r, 1, 'COMPANY'); put(wsQ, r, 2, f'Billing MW (IT) - {lab}: Vineland billed tranches + other sites lagged 1Q'); put(wsQ, r, 3, 'IT')
    for i in range(len(Q)):
        nb = sum(1 for t, qq in sched.items() if QI[qq] <= i)
        if i == 0: other = '0'
        else:
            pc = CL(QC0 + i - 1); other = f"({pc}{CO['IT_base']}-N({pc}{vb}))"
        put(wsQ, r, QC0 + i, f'={nb}*{trIT}+{other}', 'Derived', 'Vineland: tranches with billing start <= quarter x 26.7 IT (equal split, Estimated); others: IT-normalized base of prior quarter (>=1Q billing lag, README)')
    CO['bill_' + lab] = r; r += 1
# ---- Vineland tranche track: per-tranche tag columns ----
GL = {'T1': ('Q4 2025', 'Q4 2025', 'Q4 2025', 'Reported'), 'T2': ('Q1 2026', 'Q1 2026', 'Q1 2026', 'Reported'), 'T3': ('Q3 2026', 'Q3 2026', 'Q3 2026', 'Reported'), 'T4': ('Q3 2026', 'Q3 2026', 'Q3 2026', 'Reported'),
      'T5': ('Q2 2027', 'Q1 2027', 'Q4 2026', 'Estimated'), 'T6': ('Q2 2027', 'Q1 2027', 'Q4 2026', 'Estimated'), 'T7': ('Q3 2027', 'Q2 2027', 'Q1 2027', 'Estimated'), 'T8': ('Q3 2027', 'Q2 2027', 'Q1 2027', 'Estimated'), 'T9': ('Q4 2027', 'Q2 2027', 'Q1 2027', 'Estimated')}
hc = 15
for j, h in enumerate(['Owned/leased', 'GPU generation', 'Customer', 'Contracted vs unsold', 'Go-live (energized) low', 'Go-live base', 'Go-live high', 'Go-live tag', 'Billing start if Bloom slips 2Q', 'Billing MW IT (equal split)']):
    c = T.cell(tr, hc + j, h); c.font = BOLD; c.fill = HDR
for j, row in enumerate(TR):
    rr = tr + 1 + j; nm = row[0]; g = GL[nm]
    put(T, rr, hc, 'Leased (build-to-suit; DataOne owns shell)', 'Reported', S['f20'])
    put(T, rr, hc + 1, 'GB300 NVL72 (assumed; not disclosed)', 'Estimated', S['gpumw'])
    put(T, rr, hc + 2, 'Microsoft', 'Reported', S['ms6k'])
    put(T, rr, hc + 3, 'Contracted (take-or-pay)', 'Reported', S['f20'])
    put(T, rr, hc + 4, g[0], g[3], 'T1-T4 delivery dates 20-F / Q2-26 letter; T5-T9 = Bloom energization (Estimated)'); put(T, rr, hc + 5, g[1], g[3], ''); put(T, rr, hc + 6, g[2], g[3], '')
    put(T, rr, hc + 7, g[3]); put(T, rr, hc + 8, TRSLIP[nm], 'Estimated', 'base billing + 2 quarters (Bloom slip scenario)')
    put(T, rr, hc + 9, f'=F{rr}', 'Derived', 'billing MW = tranche IT MW once billing starts')
for c in range(hc, hc + 10): T.column_dimensions[CL(c)].width = 16
# ---- Capacity tranches (company-wide) ----
header(wsC, 1, ['Tranche / block', 'Site', 'Owned/leased', 'GPU generation', 'Customer', 'Contracted vs unsold', 'MW (stated)', 'MW IT (normalized, YE27 base)', 'Go-live low', 'Go-live base', 'Go-live high', 'Go-live tag', 'Evidence', 'Source'])
def first_q(k, block):
    lst = ff_cache[k]; idx = 0 if block == 'low' else 1
    for i, v in enumerate(lst):
        val = v[idx] if block != 'base' else None
        if block == 'base':
            lo, hi = v[0], v[1]
            try: m = (float(lo) + float(hi)) / 2
            except Exception: m = 1
            if m > 0: return Q[i]
        else:
            try:
                if float(val) > 0: return Q[i]
            except Exception: return Q[i]
    return 'after Q4 2030'
ff_cache = {x['k']: ffill(x['k']) for x in SITES if A[x['k']]}
cr = 2
CT = [('Microsoft T1-T9 (see Vineland tranche track)', 'VIN', 'Microsoft', 'Contracted ($17.39B TCV, 9 tranches)', 'Reported', S['f20'])]
for nm, k, cust, ctr, tg, src in CT:
    put(wsC, cr, 1, nm); put(wsC, cr, 2, SK[k]['name']); put(wsC, cr, 3, OWN[k]); put(wsC, cr, 4, 'GB300 NVL72 (assumed)'); put(wsC, cr, 5, cust, tg, src); put(wsC, cr, 6, ctr, tg, src)
    put(wsC, cr, 7, 300, 'Reported', S['kef']); put(wsC, cr, 8, 240, 'Estimated', 'cap 240 IT'); put(wsC, cr, 9, 'Q4 2025'); put(wsC, cr, 10, 'Q4 2025'); put(wsC, cr, 11, 'Q4 2025'); put(wsC, cr, 12, 'Reported (T1); T5-T9 Estimated'); put(wsC, cr, 13, 'T1 Nov-25, T2 Feb-26, T3-T4 Jul/Aug-26; T5-T9 need Bloom'); put(wsC, cr, 14, src); cr += 1
OTHER = [
 ('Meta #1 (2 tranches, $2.88B, 5-yr)', 'site not disclosed', 'n/d', 'n/d', 'Meta', 'Contracted; delivered (T1 Nov/Dec-25, T2 Feb-26)', 'n/d', 'n/d', 'Q4 2025', 'Q4 2025', 'Q4 2025', 'Reported', '20-F: first tranche Nov 2025 (Dec 2025 in FS note), second Feb 2026; MW not disclosed (peer file ~50 IT MW = unsourced, E low)', S['f20']),
 ('Meta #2 dedicated ($12B, 5-yr, Vera Rubin, multiple locations)', 'n/d (Bethune likely per press)', 'n/d', 'Vera Rubin NVL72 (Reported)', 'Meta', 'Contracted', 'n/d', 'n/d', 'Q1 2027', 'Q1 2027', 'Q2 2027', 'Reported (start "early 2027" / 1H27)', '20-F: delivery expected to commence early 2027; plus up to $15B Meta backstop for unsold capacity of certain clusters', S['f20']),
 ('Q2-26 large AI-cloud deals (4 deals, avg >$1B TCV; e.g. Reflection, Cohere)', 'n/d', 'n/d', 'n/d (new-gen)', 'Q2 neolab deals', 'Contracted', 'n/d', 'n/d', 'Q4 2026', 'Q4 2026', 'Q1 2027', 'Reported (from late Q4 2026)', 'ACV $20-25M/MW; "start coming online from late Q4 and onwards" (Q2-26 call); MW = TCV/(term x ACV) but term not disclosed -> n/d', S['q226'] + '; ' + S['q226call']),
]
for row in OTHER:
    for j, v in enumerate(row, 1): put(wsC, cr, j, v, 'Reported' if j in (5, 6) else None, row[-1] if j in (5, 6) else None)
    cr += 1
ye27 = CL(QC0 + QI['Q4 2027'])
for x in SITES:
    k = x['k']
    if k == 'VIN': continue
    cust = x['tag']; contracted = 'Contracted' if any(w in cust for w in ('Meta', 'Microsoft')) else 'Unsold / self-serve or undisclosed customer (not disclosed)'
    put(wsC, cr, 1, f'{x["name"]} - site block'); put(wsC, cr, 2, x['name']); put(wsC, cr, 3, OWN[k], 'Reported', x['src'][:200]); put(wsC, cr, 4, GEN_OF.get(k, x['gen']), 'Estimated', 'Sites tab')
    put(wsC, cr, 5, cust, 'Estimated', x['tagev']); put(wsC, cr, 6, contracted, 'Estimated', x['tagev'])
    put(wsC, cr, 7, x['cmw'] if x['cmw'] is not None else 'n/d', 'Reported', x['src'][:200])
    if A[k]:
        rowb = ROWS[k]['Energized MW base (=midpoint)']
        f = f"='Quarterly MW'!{ye27}{rowb}" + ('' if BASIS[k][2] == 'IT' else f'/{fac(k)}')
        put(wsC, cr, 8, f, 'Derived', 'YE27 base energized, IT-normalized')
        put(wsC, cr, 9, first_q(k, 'low')); put(wsC, cr, 10, first_q(k, 'base')); put(wsC, cr, 11, first_q(k, 'high'))
    else:
        put(wsC, cr, 8, 'n/d'); put(wsC, cr, 9, 'n/d'); put(wsC, cr, 10, 'n/d'); put(wsC, cr, 11, 'n/d')
    put(wsC, cr, 12, 'Estimated (first quarter with energized MW > 0)'); put(wsC, cr, 13, x['my']); put(wsC, cr, 14, x['src'][:300]); cr += 1
for j, w in enumerate([46, 34, 10, 24, 24, 30, 10, 12, 10, 10, 10, 22, 70, 60], 1): wsC.column_dimensions[CL(j)].width = w
# ---- Alt path (mw_path_proposed.csv) + 150 MW split ----
rows_alt = list(_csv.DictReader(open('/workspace/nebius/capacity/mw_path_proposed.csv')))
header(wsA, 1, list(rows_alt[0].keys()))
for i, rw in enumerate(rows_alt, 2):
    for j, (kk, v) in enumerate(rw.items(), 1):
        try: v = float(v)
        except Exception: pass
        put(wsA, i, j, v, 'Estimated' if kk in ('low', 'base', 'high') else None, S['altpath'] if kk in ('low', 'base', 'high') else None)
ar = len(rows_alt) + 4
wsA.cell(ar, 1, "Q4 2027 ACTIVE base: split of difference vs v1 file (nebius_capacity_v1.xlsx)").font = BOLD; ar += 1
header(wsA, ar, ['Component', 'Alt path (CSV) vs v1, MW', 'v2 model vs v1, MW', 'Tag', 'Detail']); ar += 1
import openpyxl as _ox
_v1 = _ox.load_workbook('/workspace/nebius/capacity/vineland_check/cap/nebius_capacity.xlsx', data_only=True)['Quarterly MW']
V1 = {}
for rr_ in range(3, 118):
    a_, b_ = _v1.cell(rr_, 1).value, _v1.cell(rr_, 2).value
    if b_ == 'Energized MW base (=midpoint)': V1[a_] = _v1.cell(rr_, 5 + QI['Q4 2027']).value
V1TOT = _v1.cell(121, 5 + QI['Q4 2027']).value
def v1(k):
    nm0 = {x['k']: x['name'] for x in SITES}[k]
    for a_, v_ in V1.items():
        if a_ and a_.split(' (')[0] == nm0.split(' (')[0]: return v_ if isinstance(v_, (int, float)) else 0
    return 0
def v2ref(k): return f"N('Quarterly MW'!{ye27}{ROWS[k]['Energized MW base (=midpoint)']})"
comps = [
 ('New sites (OK, MN, Newport, UK 3rd)', 153, '=' + '+'.join(f'({v2ref(k)}-{v1(k)})' for k in ('OK', 'MN', 'NEW', 'UK3')), 'OK +100, MN +10, Newport+UK3 +43 in CSV'),
 ('Pulled-forward / slipped sites (CLT1, Israel, KC)', 0, '=' + '+'.join(f'({v2ref(k)}-{v1(k)})' for k in ('AIB', 'MAS', 'BEI', 'KC')), 'CSV: CLT1 +12.5 and Israel -29 occur in Q3 2027 / earlier, net 0 at Q4 2027'),
 ('Basis / cap change (Vineland 240 IT cap)', 0, f"=({v2ref('VIN')}-{v1('VIN')})", 'CSV kept 240 base; v2 caps high at 240 -> base 235'),
]
s0 = ar
for nm, a_, f_, d_ in comps:
    put(wsA, ar, 1, nm); put(wsA, ar, 2, a_, 'Derived', S['altpath'] + ' notes column'); put(wsA, ar, 3, f_, 'Derived', 'v2 Quarterly MW base minus v1 base (hard-coded from nebius_capacity_v1.xlsx)'); put(wsA, ar, 4, 'Derived'); put(wsA, ar, 5, d_); ar += 1
put(wsA, ar, 1, 'Total'); put(wsA, ar, 2, f'=SUM(B{s0}:B{ar-1})', 'Derived', 'sum'); put(wsA, ar, 3, f'=SUM(C{s0}:C{ar-1})', 'Derived', 'sum'); ar += 1
put(wsA, ar, 1, 'Check: v1 total Q4 2027 base'); put(wsA, ar, 2, V1TOT, 'Reported', 'nebius_capacity_v1.xlsx Quarterly MW row 121'); put(wsA, ar, 3, f"='Quarterly MW'!{ye27}{CO['Identified sites - energized base']}", 'Derived', 'v2 total'); ar += 1
put(wsA, ar, 1, 'Memo: IT-normalization effect on v2 Q4 2027 base (IT-normalized minus stated)'); put(wsA, ar, 3, f"='Quarterly MW'!{ye27}{CO['IT_base']}-'Quarterly MW'!{ye27}{CO['Identified sites - energized base']}", 'Derived', 'basis change shown separately (not part of 150 split)'); ar += 1
wsA.column_dimensions['A'].width = 50; wsA.column_dimensions['E'].width = 60; wsA.column_dimensions['M'].width = 80
# ---- Summary ----
U = wsU; U.cell(1, 1, 'SUMMARY (v2) - Nebius energized / connected / leased MW. Base = midpoint; IT/gross via Basis tab. All projections Estimated.').font = BOLD
header(U, 3, ['Metric', 'YE26 low', 'YE26 base', 'YE26 high', 'YE27 low', '', 'YE27 base', 'YE27 high', 'Tag', 'Source / note'])
y26, y27 = CL(QC0 + QI['Q4 2026']), CL(QC0 + QI['Q4 2027'])
def qref(col, row): return f"='Quarterly MW'!{col}{row}"
lines = [
 ('Energized MW - stated basis (mixed)', CO['Identified sites - energized low'], CO['Identified sites - energized base'], CO['Identified sites - energized high']),
 ('Energized MW - IT-normalized', CO['IT_low'], CO['IT_base'], CO['IT_high']),
 ('Energized MW - gross-normalized', CO['gross_low'], CO['gross_base'], CO['gross_high']),
 ('Connected MW (gross/shell, bottom-up)', CO['conn_low'], CO['conn_base'], CO['conn_high']),
]
ur = 4; UR = {}
for nm, lo, ba, hi in lines:
    put(U, ur, 1, nm)
    put(U, ur, 2, qref(y26, lo), 'Derived', 'link'); put(U, ur, 3, qref(y26, ba), 'Derived', 'link'); put(U, ur, 4, qref(y26, hi), 'Derived', 'link')
    put(U, ur, 5, qref(y27, lo), 'Derived', 'link'); put(U, ur, 7, qref(y27, ba), 'Derived', 'link'); put(U, ur, 8, qref(y27, hi), 'Derived', 'link'); put(U, ur, 9, 'Derived/Estimated')
    UR[nm] = ur; ur += 1
put(U, ur, 1, 'Modeling placeholder YE27 energized IT MW'); put(U, ur, 7, 950, 'Reported', 'Modeling team placeholder (internal)'); UR['ph'] = ur; ur += 1
put(U, ur, 1, 'Ours (IT base) minus 950 placeholder'); put(U, ur, 7, f"=G{UR['Energized MW - IT-normalized']}-G{UR['ph']}", 'Derived', 'difference'); ur += 1
put(U, ur, 1, 'Mgmt guide: connected'); put(U, ur, 3, '800-1,000', 'Reported', S['q226call']); put(U, ur, 7, '~1,800-2,000 (YE26 + >1 GW/yr)', 'Derived', S['q226']); ur += 1
put(U, ur, 1, 'Vineland Bloom slips 2Q: IT MW lost at YE27'); put(U, ur, 7, f"='Quarterly MW'!{y27}{CO['vin_slip_d']}", 'Derived', 'scenario row'); ur += 1
put(U, ur, 1, 'Billing MW (IT) at quarter-end: base / Bloom-slip'); put(U, ur, 3, qref(y26, CO['bill_base']), 'Derived', 'link'); put(U, ur, 7, qref(y27, CO['bill_base']), 'Derived', 'link'); put(U, ur, 8, qref(y27, CO['bill_Bloom slips 2Q']), 'Derived', 'col H = Bloom-slip case'); ur += 2
# leased IT MW by year
U.cell(ur, 1, 'LEASED vs OWNED IT MW (for Note 12 $12.05B signed-not-commenced lease check: leased IT MW x $2.0-2.5M/IT MW/yr)').font = BOLD; ur += 1
header(U, ur, ['Site', 'Owned/leased', 'Contracted MW (stated)', 'Contracted IT MW', 'Energized IT YE25 (base)', 'Energized IT YE26 (base)', 'Energized IT YE27 (base)', 'Contracted IT not yet energized at YE26', 'Tag', 'Note']); ur += 1
y25 = CL(QC0 + QI['Q4 2025']); cq = CL(QC0 + QI['Q4 2027'])
l0 = ur
for x in SITES:
    k = x['k']
    if OWN[k] != 'leased': continue
    put(U, ur, 1, x['name']); put(U, ur, 2, 'leased')
    cref = f"'Quarterly MW'!{cq}{ROWS[k]['Contracted MW']}"
    put(U, ur, 3, f'={cref}', 'Reported', 'Quarterly MW contracted row')
    conv = '' if BASIS[k][2] == 'IT' else f'/{fac(k)}'
    if k == 'VIN': put(U, ur, 4, 240, 'Estimated', '300 gross / 1.25 (cap)')
    elif k == 'NEW': put(U, ur, 4, 43, 'Derived', 'Newport + UK 3rd = 65 - 22 (Jun 8 2026 PR); UK3 row left blank to avoid double count')
    elif k in ('UK3', 'MN', 'OK', 'FR2'): put(U, ur, 4, 'n/d', 'Reported', 'contracted MW not disclosed')
    else: put(U, ur, 4, f'=IF(ISNUMBER({cref}),{cref}{conv},"n/d")', 'Derived', 'contracted x IT factor')
    for cc, qq in [(5, y25), (6, y26), (7, y27)]:
        if A[k]:
            put(U, ur, cc, f"=N('Quarterly MW'!{qq}{ROWS[k]['Energized MW base (=midpoint)']}){conv}", 'Derived', 'energized base, IT-normalized (commencement proxy)')
        else: put(U, ur, cc, 'n/d')
    put(U, ur, 8, f'=IF(ISNUMBER(D{ur}),MAX(0,D{ur}-N(F{ur})),"n/d")', 'Derived', 'contracted IT - energized IT YE26')
    put(U, ur, 9, 'Derived/Estimated'); put(U, ur, 10, BASIS[k][5]); ur += 1
put(U, ur, 1, 'TOTAL leased'); 
for cc in (4, 5, 6, 7, 8): put(U, ur, cc, f'=SUM({CL(cc)}{l0}:{CL(cc)}{ur-1})', 'Derived', 'sum (n/d excluded)')
LT = ur; ur += 1
put(U, ur, 1, 'TOTAL owned (IT-normalized energized)')
o_terms = lambda qq: '+'.join(f"N('Quarterly MW'!{qq}{ROWS[k]['Energized MW base (=midpoint)']})/{fac(k)}" for k in OWN if OWN[k] == 'owned')
for cc, qq in [(5, y25), (6, y26), (7, y27)]: put(U, ur, cc, '=' + o_terms(qq), 'Derived', 'owned sites')
ur += 1
put(U, ur, 1, 'Illustrative: rent on leased contracted IT MW not yet energized at YE26, $M/yr @ $2.0M / $2.5M'); put(U, ur, 2, f'=H{LT}*2.0', 'Derived', 'Modeling rate $2.0M/IT MW/yr (assumption, Modeling)'); put(U, ur, 3, f'=H{LT}*2.5', 'Derived', 'Modeling rate $2.5M/IT MW/yr'); ur += 1

# ---- Note 12 lease cross-check (Filings input, Oct 4 2026) ----
ur += 1
U.cell(ur, 1, "LEASES SIGNED NOT COMMENCED (Q2'26 6-K Ex.99.2 lease note) vs our leased IT MW").font = BOLD; ur += 1
q2c = CL(QC0 + QI['Q2 2026'])
LEASED = [k for k in OWN if OWN[k] == 'leased']
def lsum(qq, excl=()):
    t = []
    for k in LEASED:
        if k in excl or not A[k]: continue
        conv = '' if BASIS[k][2] == 'IT' else f'/{fac(k)}'
        t.append(f"N('Quarterly MW'!{qq}{ROWS[k]['Energized MW base (=midpoint)']}){conv}")
    return '=' + '+'.join(t)
LR = {}
def lput(label, val, tag, src, key):
    global ur
    put(U, ur, 1, label); put(U, ur, 2, val, tag, src); put(U, ur, 9, tag); put(U, ur, 10, src[:200]); LR[key] = ur; ur += 1
lput('Leases signed not commenced at 6/30/26, undiscounted $M', 12054.1, 'Reported', "Q2'26 interim FS (6-K Ex.99.2) lease note; primarily DC facilities + equipment; commencing 2026-2027; terms up to 12 yrs; $612.4M prepaid", 'snc')
lput('Same at 12/31/25, $M (terms up to 11 yrs; $271.0M prepaid)', 9759.6, 'Reported', 'FY25 20-F Note 8', 'snc25')
lput('ROU assets 6/30/26 $M (YE25 918.8; H1 additions 996.2)', 1855.1, 'Reported', "Q2'26 interim FS lease note", 'rou')
lput('Term assumption (yrs) - Filings check uses 12 (max disclosed)', 12, 'Estimated', 'Filings team check; disclosed as "up to 12 years"', 'term')
lput('Alt term assumption (yrs)', 10, 'Estimated', 'sensitivity: average term below the maximum', 'term2')
lput('Implied annual payments $M/yr @ 12 yrs', f"=B{LR['snc']}/B{LR['term']}", 'Derived', 'undiscounted / term', 'pay')
lput('Filings check: implied leased IT MW @ $2.0M/IT MW-yr', f"=B{LR['pay']}/2.0", 'Derived', 'Filings rough check (400-500 MW range)', 'f20')
lput('Filings check: implied leased IT MW @ $2.5M/IT MW-yr', f"=B{LR['pay']}/2.5", 'Derived', 'Filings rough check', 'f25')
lput('OURS: leased IT MW energized (commenced proxy) YE25', lsum(y25), 'Derived', 'sum of leased sites, IT-normalized base', 'l25')
lput("OURS: leased IT MW energized at 6/30/26 (already commenced proxy)", lsum(q2c), 'Derived', 'sum of leased sites, IT-normalized base', 'lq2')
lput('OURS: leased IT MW commencing in 2026 (YE26 - YE25)', f"={lsum(y26)[1:]}-B{LR['l25']}", 'Derived', 'energization as commencement proxy', 'c26')
lput('OURS: leased IT MW commencing in 2027 (YE27 - YE26)', f"={lsum(y27)[1:]}-({lsum(y26)[1:]})", 'Derived', 'energization as commencement proxy', 'c27')
lput('OURS: leased IT MW commencing 7/1/26 - 12/31/27 (comparable to the $12.05B pool)', f"={lsum(y27)[1:]}-B{LR['lq2']}", 'Derived', 'YE27 leased energized - 6/30/26 leased energized', 'pool')
lput('  of which excl. Vineland (build-to-suit lease may have commenced Nov-25 with T1)', f"={lsum(y27, ('VIN',))[1:]}-({lsum(q2c, ('VIN',))[1:]})", 'Derived', 'same excl. VIN', 'poolx')
lput('OURS: contracted leased IT MW not energized at 6/30/26 (named sites; n/d excluded)', f"=D{LT}-B{LR['lq2']}", 'Derived', 'contracted leased IT total - energized at 6/30/26', 'cne')
lput('Implied rent $M per IT MW-yr if OUR pool MW is right @ 12 yrs', f"=B{LR['pay']}/B{LR['pool']}", 'Derived', '$12.05B / 12 / our pool MW', 'r12')
lput('Implied rent $M per IT MW-yr if OUR pool MW is right @ 10 yrs', f"=B{LR['snc']}/B{LR['term2']}/B{LR['pool']}", 'Derived', 'sensitivity', 'r10')
lput('Implied rent excl. Vineland @ 12 yrs', f"=B{LR['pay']}/B{LR['poolx']}", 'Derived', 'if Vineland lease already commenced', 'r12x')
POST = ('AIB', 'NEW', 'UK3')
lput('  less leases likely signed after 6/30/26 (AIB CLT1 Sep 30; Newport Aug 13; UK3 lead Sep 1) - YE27 IT MW', '=' + '+'.join(f"N('Quarterly MW'!{y27}{ROWS[k]['Energized MW base (=midpoint)']})" for k in POST), 'Derived', 'excluded: cannot be in the 6/30/26 not-commenced pool', 'post')
lput('OURS: comparable pool excl. Vineland and post-6/30 signings', f"=B{LR['poolx']}-B{LR['post']}", 'Derived', 'best like-for-like vs Filings 400-500', 'pooladj')
lput('Implied rent $M per IT MW-yr on comparable pool @ 12 yrs', f"=B{LR['pay']}/B{LR['pooladj']}", 'Derived', '$12.05B / 12 / comparable pool', 'radj')
U.cell(ur, 1, "VERDICT (Derived/Estimated): Including Vineland, our leased MW commencing 7/1/26-YE27 (~642 IT MW) is ABOVE the 400-500 MW check, i.e. not short. "
       "Like-for-like (Vineland build-to-suit likely commenced with T1 in Nov-25; AIB/Newport/UK3 signed after 6/30) our pool is ~383 IT MW, slightly BELOW 400-500, implying ~$2.6M/IT MW-yr at a 12-yr term. "
       "Leased MW and the $2.0-2.5M rent therefore bracket the disclosure; if anything the rent assumption looks slightly LOW (or more MW than modeled commence by YE27, e.g. Bethune/Beit Shemesh ahead of our base, larger OK). "
       "The pool also includes equipment leases, which would lower implied DC rent. No site-level ladder is disclosed, so this is a range check, not a reconciliation."); LR['verdict'] = ur; ur += 1

U.cell(ur, 1, 'Caveats: Vineland is a build-to-suit lease that has partly commenced (T1-T4); Note 12 composition by site is not disclosed; MN/OK/FR2/UK3 contracted MW n/d (excluded); AIB CLT1 rent ~>=$1.3B/12y/50MW = ~$2.2M/IT MW/yr (Derived, AIB deck) is a direct check on the $2.0-2.5M range.'); ur += 2
# Q2'26 reconciliation
U.cell(ur, 1, "Q2'26 RECONCILIATION (~190 MW implied vs model)").font = BOLD; ur += 1
q2 = CL(QC0 + QI['Q2 2026'])
recon = [
 ('Q2-26 AI cloud revenue $M', 574.9, 'Reported', S['q226']),
 ('2026 base ACV per MW $M/yr (letter chart "ACV per MW... $12M 2026 base"; ACV = annual contract value, NOT capex)', 12, 'Reported', S['q226'] + ' p.1 + footnote 1'),
 ('Implied average billing MW in Q2 = rev x 4 / ACV', '=B{0}*4/B{1}', 'Derived', ''),
 ('June-26 exit: ARR $3.0B / $12M', '=3000/B{1}', 'Derived', 'ARR = last month x 12 (letter footnote)'),
 ('v1 model Q2-26 identified base (stated basis)', 173, 'Reported', 'nebius_capacity_v1.xlsx'),
 ('v2 model Q2-26 identified base (stated basis)', f"='Quarterly MW'!{q2}{CO['Identified sites - energized base']}", 'Derived', 'adds MN 10 + OK 10 (Estimated)'),
 ('v2 model Q2-26 IT-normalized base', f"='Quarterly MW'!{q2}{CO['IT_base']}", 'Derived', ''),
 ('v2 model Q2-26 gross-normalized base', f"='Quarterly MW'!{q2}{CO['gross_base']}", 'Derived', ''),
 ('Sensitivity: Vineland at 2 delivered tranches (2 x 26.7 IT) instead of altdata 39', f"=2*{trIT}-'Quarterly MW'!{q2}{vb}", 'Derived', 'equal-split tranche MW vs thermal-bounded 39'),
]
r0 = ur
for j, (nm, v, tg, src) in enumerate(recon):
    vv = v.format(r0, r0 + 1) if isinstance(v, str) else v
    put(U, ur, 1, nm); put(U, ur, 2, vv, tg, src); put(U, ur, 9, tg); put(U, ur, 10, src); ur += 1
U.cell(ur, 1, "Result: the ~190 MW figure is a REVENUE-implied AVERAGE billing MW on Nebius's own (active = IT-consumed) basis. The gap to 173 closes mainly through MN + OK colos (~20 MW, live by Q2 per on-site jobs) -> v2 Q2-26 base ~193; IT-normalizing pulls it down (owned Mantsala 75 at PUE 1.1 and Longcross gross) while Vineland at tranche-equal-split (53 vs 39 IT) would add ~14. KC (20 in Q2-26) is unchanged in Q2-26; its re-anchoring affects 2025 history."); ur += 2
# GPU anchor
U.cell(ur, 1, "GPU ANCHOR CHECK (Q1'25)").font = BOLD; ur += 1
q1c = CL(QC0 + QI['Q1 2025'])
put(U, ur, 1, 'Company-stated GPUs deployed at Mar 31 2025'); put(U, ur, 2, 30000, 'Reported', S['f24']); ur += 1
put(U, ur, 1, 'Modeled GPUs Q1-25 (v2, GPUs deployed tab)'); put(U, ur, 2, f"='GPUs deployed'!{q1c}{gtot}", 'Derived', 'energized base x GPUs/MW'); ur += 1
put(U, ur, 1, 'Modeled / stated'); put(U, ur, 2, f'=B{ur-1}/B{ur-2}', 'Derived', ''); ur += 1
put(U, ur, 1, 'Energized MW Q1-25 (stated basis, v2)'); put(U, ur, 2, f"='Quarterly MW'!{q1c}{CO['Identified sites - energized base']}", 'Derived', 'Mantsala 25 + Keflavik 10 + Paris 5 + KC 10'); ur += 1
put(U, ur, 1, 'Gross MW needed for 30k Hopper at gross density'); put(U, ur, 2, f"=30000/'GPUs deployed'!$C${GROW['Hopper H100/H200']}", 'Derived', 'GPUs / Hopper GPUs per gross MW'); ur += 1
U.column_dimensions['A'].width = 70
for c in 'BCDEFGH': U.column_dimensions[c].width = 13
U.column_dimensions['J'].width = 70
wb.move_sheet('Summary', offset=-(len(wb.sheetnames) - 2))

# ===== README =====
R = [
 ('Nebius (NBIS) capacity model - built ' + '2026-10-04 (ET) by Capacity & Power analyst. Public information only. Script: /workspace/nebius/capacity/build_capacity.py', True),
 ('', False),
 ('DEFINITIONS (Nebius, Q1-26 letter): contracted = power secured by contracted land and power commitments; connected = power connected into fully built and equipped DCs; active = power consumed by installed operational IT equipment, available for revenue.', False),
 ("'Energized MW' in this workbook = my estimate of power physically live at the site (closest to 'active'). It is NOT 'connected' (built-and-equipped, may not be powered) and NOT 'contracted'.", False),
 ('MW BASIS: Nebius never states IT vs gross. Site rows use the stated basis (mostly unknown/gross). Vineland energized rows are IT MW (alt-data power-bounded); Modi\'in and AIB are IT. Company totals therefore mix bases (1.1-1.45x PUE ambiguity).', False),
 ('', False),
 ('MAPPING FOR MODELING: Energized low -> Bear; Energized base (midpoint formula) -> Base; Energized high -> Bull. Contracted MW is not a revenue driver.', False),
 ('BILLING LAG (Estimated): revenue/billing lags energization by >= 1 quarter for new capacity (commission DC, network, clusters, platform, onboarding: Q2-26 call "takes a few months"; active through 1H27). Exception: Vineland T1-T4 billing at reported delivery dates.', False),
 ('Rates context (Q2-26 letter): ~$12M/MW/yr 2026 installed base; >$20M/MW Q2 new deals ($20-25M ACV/MW); >$40M/MW Q3 short-term deals.', False),
 ('CAPEX (corrected): FY26 capex guide $20-25B (CFO, Q2-26 call Aug 12 2026; not in an SEC filing). On +580-780 MW connected (220 -> 800-1,000) = $25.6-43.1M per connected MW (Derived: 20,000/780 and 25,000/580). Earlier $16-20B figure superseded.', False),
 ('', False),
 ('YE2026 ENERGIZED (identified sites, this model): see Quarterly MW Q4 2026 company rows. Revised from prior ~330/600/800 after alt-data: Vineland capped at the generator ceiling (~107 MW IT) until Bloom is confirmed; Bethune 120 MW at risk; Birmingham/Independence revenue 2027.', False),
 ('PLUG ROW: = mgmt reported/guided figure - identified base. At YE26 it shows the MW that must come from undisclosed sites (MN, OK, France 2nd, UK 3rd) or from counting built-but-unpowered halls as "connected" for the guide to be met. It is a reconciliation, not an estimate.', False),
 ('GPUs: energized base MW x GPUs/MW (gpu_mw_capex.xlsx base rows; gross column for stated-basis sites, IT column for IT-basis sites). Reported GPU counts ran ahead of powered MW (delivered-not-powered): never back MW out of GPU counts.', False),
 ('GPU vintage: generation per site is mostly ASSUMED (Nebius rarely discloses per site). 2027+ owned sites assumed Vera Rubin. Cohort rows = quarterly MW additions for vintage depreciation (servers 5-yr life from Jan 2026 per 20-F).', False),
 ('Colours: blue font = hard-coded input; black = formula. Every input/formula cell has a tag (Reported / Derived / Estimated) and source in Tags_Sources. "n/d" = not disclosed.', False),
 ('V2 (2026-10-05 ET): backup of prior file = nebius_capacity_v1.xlsx (script build_capacity_v1.py). Changes: added Newport (Vantage CWL1), UK 3rd (lead), OK and MN anchors (Estimated); KC starts Q1-25 at 5-15 MW (GPU anchor: FY24 20-F ~30,000 GPUs at Mar 31 2025, replaces ~38k); Vineland capped at 240 IT MW; Israel slipped per Mega Or Q2-26 deck; new Basis tab (IT/gross/not stated + gross/IT factor per site, peer ratios from peers/comparison.xlsx); IT- and gross-normalized energized series, connected-MW line, billing-MW rows, Bloom-slips-2Q scenario (Quarterly MW bottom rows); per-tranche tags (Vineland tranche track cols O-X; Capacity tranches tab); Alt path tab (mw_path_proposed.csv + Q4-27 split); Summary tab.', False),
 ('Verification grades and rationale: /workspace/nebius/capacity/site_verification.md. Alt-data: /workspace/nebius/capacity/altdata/altdata_timelines.xlsx.', False),
]
for i, (t, b) in enumerate(R, 1):
    c = wsR.cell(i, 1, t); c.font = BOLD if b else Font()
wsR.column_dimensions['A'].width = 200

# ===== Tags_Sources =====
header(wsL, 1, ['Sheet', 'Cell', 'Value / formula', 'Tag', 'Source', 'Note'])
for i, row in enumerate(LOG, 2):
    for j, v in enumerate(row, 1):
        wsL.cell(i, j, str(v) if v is not None else '')
wsL.column_dimensions['C'].width = 40; wsL.column_dimensions['E'].width = 100
wb.save(OUT)
print('saved', OUT, 'logged cells', len(LOG))

# ---- recalc via LibreOffice so cached values exist for data_only readers; then error scan ----
import subprocess, shutil, os, tempfile
tmp = tempfile.mkdtemp()
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', tmp, OUT], capture_output=True)
rc = os.path.join(tmp, os.path.basename(OUT))
if os.path.exists(rc):
    chk = openpyxl.load_workbook(rc, data_only=True)
    errs = [(w.title, c.coordinate, c.value) for w in chk for row in w.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('#')]
    print('recalc errors:', len(errs), errs[:10])
    shutil.copy(rc, OUT)
