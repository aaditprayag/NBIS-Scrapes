# Builds private_peers.xlsx: KPI profiles for private neocloud peers (Nscale, Lambda, Crusoe)
# on the SAME row IDs / definitions as /workspace/nebius/peers/coreweave/crwv_kpis.xlsx (I01..I27, K01..K40, D1..D8).
# Private peers have sparse, irregular disclosure, so each company tab is a list of DATED data points
# (as-of date, tag, source URL) instead of quarterly columns. Derived/estimated points are live Excel formulas
# over other data points; the Summary tab links the latest point per company into the CoreWeave row layout.
# Author: Grok Bot for Aadit Prayag, 2026-10-04 ET. Every input is hard-coded below with tag + source.
import re, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

OUT = "/workspace/nebius/peers/private/private_peers.xlsx"
VALS_OUT = "/workspace/nebius/peers/private/private_peers_values.json"
SOFR = 3.87  # NY Fed SOFR, 2026-10-01 (used only for ESTIMATED all-in coupons on floating facilities)

TAGNAME = {"R": "Reported", "P": "Press-reported", "D": "Derived", "E": "Estimated", "N": "n/d"}
SEC = "https://www.sec.gov/Archives/edgar/data/"
S = {
 # Nscale
 "NS-S1": ("Nscale Ltd Form S-1 (filed 2026-09-18; CIK 0002110365). Local copy /workspace/nebius/comps_raw/nscale_s1.htm(.txt)", "2026-09-18", SEC + "2110365/000119312526395475/ck0002110365-20260918.htm"),
 "NS-EDGAR": ("EDGAR filing index, Nscale Ltd (DRS 2026-02-17, DRS/A x3, S-1 2026-09-18; no S-1/A as of 2026-10-04)", "2026-10-04", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0002110365"),
 "NS-IPO": ("Nscale PR: files registration statement for proposed IPO (NYSE: NSCL)", "2026-09-18", "https://www.nscale.com/press-releases/nscale-files-initial-public-offering"),
 "NS-C": ("Nscale PR: $2B Series C at $14.6B valuation", "2026-03-09", "https://www.nscale.tech/press-releases/nscale-series-c"),
 "NS-CV": ("Nscale PR (PRNewswire): $3.36B pre-IPO convertible loan notes", "2026-09-25", "https://www.prnewswire.com/news-releases/nscale-raises-3-36b-in-pre-ipo-convertible-financing-302890199.html"),
 "NS-FT": ("Deep Quarry (Substack) citing FT: S-1 aims at valuation 'as high as $35B'", "2026-09", "https://deepquarry.substack.com/p/what-changed-between-nscales-confidential"),
 # Lambda
 "LM-EDGAR": ("EDGAR: Lambda, Inc./DE (CIK 0001954966) has only Form D filings (2022-11-17, 2024-02-16); EDGAR full-text search finds no S-1/S-1/A/DRS by Lambda", "2026-10-04", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001954966"),
 "LM-FD24": ("Lambda, Inc. Form D (filed 2024-02-16; first sale 2023-10-05)", "2024-02-16", SEC + "1954966/000195496624000001/primary_doc.xml"),
 "LM-IG": ("Lambda PR: closes $1.008B investment-grade fixed-rate DDTL (6.78%)", "2026-10-01", "https://lambda.ai/blog/lambda-closes-1-billion-senior-secured-fixed-rate-financing"),
 "LM-TLB": ("Lambda PR: closes $926M senior secured term loan B (SOFR+3.00%)", "2026-08-27", "https://lambda.ai/blog/lambda-closes-926-million-senior-secured-term-loan-b-facility"),
 "LM-CF": ("Lambda PR: closes $1B senior secured credit facility (upsized from $275M Aug-2025)", "2026-05-07", "https://lambda.ai/blog/lambda-closes-1-billion-senior-secured-credit-facility"),
 "LM-500": ("Lambda PR (BusinessWire): $500M GPU-backed facility (Macquarie)", "2024-04-02", "https://www.businesswire.com/news/home/20240402148086/en/Lambda-Announces-%24500M-GPU-Backed-Facility-to-Expand-Cloud-for-AI"),
 "LM-E": ("Lambda PR: raises over $1.5B Series E led by TWG Global", "2025-11-18", "https://lambda.ai/blog/lambda-raises-over-1.5b-from-twg-global-usit-to-build-superintelligence-cloud-infrastructure"),
 "LM-MSFT": ("Lambda PR: multibillion-dollar agreement with Microsoft (tens of thousands of NVIDIA GPUs incl. GB300 NVL72)", "2025-11-03", "https://lambda.ai/blog/lambda-announces-multibillion-dollar-agreement-with-microsoft-to-deploy-ai-infrastructure-powered-by-tens-of-thousands-of-nvidia-gpus"),
 "LM-KC": ("Lambda PR: Kansas City AI factory (24MW initial, >100MW potential, >10,000 Blackwell Ultra GPUs)", "2025-10-28", "https://lambda.ai/blog/lambda-to-build-a-100mw-ai-factory-in-kansas-city-mo"),
 "LM-GTC": ("Lambda blog at GTC 2026 (AI factory with 10,000+ GB300 GPUs in production)", "2026-03-16", "https://lambda.ai/blog/lambda-at-gtc-2026-building-the-superintelligence-cloud"),
 "LM-TI1": ("DCD citing The Information (documents viewed): Q2'25 revenue ~$140M, H1'25 >$250M, GM 50%/61%, H1 net loss ~$24M", "2025-09-04", "https://www.datacenterdynamics.com/en/news/ai-cloud-lambda-hires-investment-banks-in-preparation-for-ipo-report/"),
 "LM-TI2": ("Deep News / Tiger Brokers relaying The Information: Q2'25 cloud revenue ~$114M, 47MW in Q2'25, ~3GW 2030 target, Amazon+Microsoft ~half of cloud revenue", "2025-09-05", "https://www.itiger.com/news/1116660571"),
 "LM-NV": ("DCD citing The Information: NVIDIA $1.5B / ~18,000-GPU 4-year leaseback; NVIDIA largest customer", "2025-09-08", "https://www.datacenterdynamics.com/en/news/nvidia-signs-15bn-deal-to-lease-its-gpus-back-from-lambda-report/"),
 "LM-TM": ("Technology Magazine interview: CEO 'well over 25,000 GPUs on our cloud platform'; Series D $480M at $2.5B", "2025-02-24", "https://technologymagazine.com/articles/why-lambda-secured-funding-and-investment-from-nvidia"),
 "LM-BBG1": ("Bloomberg: Lambda in talks for up to $3B pre-IPO round at $12B+; expected 2026 revenue >$1.5B", "2026-08-24", "https://www.bloomberg.com/news/articles/2026-08-24/ai-cloud-provider-lambda-in-talks-for-3-billion-pre-ipo-round"),
 "LM-BBG2": ("Bloomberg: Lambda raises ~$1B private short-dated debt (JPMorgan) for GPUs leased to Microsoft", "2026-08-28", "https://www.bloomberg.com/news/articles/2026-08-28/nvidia-backed-lambda-inks-1-billion-private-debt-for-chip-deal"),
 "LM-TC": ("TechCrunch: Series E at $5.43B post-money (PitchBook)", "2026-08-28", "https://techcrunch.com/2026/08/28/neocloud-lambda-secures-1b-in-debt-to-buy-more-chips/"),
 "LM-ANTH": ("TheEnergyMag citing WSJ: Anthropic $35B, six-year, ~350MW at Hut 8 Beacon Point (Nueces County TX), NVIDIA holds lease", "2026-09-01", "https://www.theenergymag.com/news/market-news/anthropic-lambda-nvidia-hut-ai-texas-beacon-point"),
 "LM-ANTH2": ("Proactive Investors citing WSJ: Anthropic-Lambda $35B deal", "2026-09-01", "https://www.proactiveinvestors.com/companies/news/1097930"),
 "LM-WSJ4B": ("TechCrunch citing WSJ (2026-10-06): Lambda raising up to $4B at $14.5B pre-money (Blackstone + Coatue lead), likely last private round before a planned 2027 IPO; investor letter shows backlog $15B (Jun-26) -> $50B (Sep-26), much of it the ~$35B Anthropic deal", "2026-10-06", "https://techcrunch.com/2026/10/06/ai-computing-startup-lambda-to-raise-4b-ahead-of-planned-ipo/"),
 "LM-350": ("DCD citing The Information: Lambda in talks to raise $350M pre-IPO (Mubadala); 15 US data centers per Lambda website", "2026-01-09", "https://www.datacenterdynamics.com/en/news/lambda-in-talks-to-raise-350m-in-pre-ipo-funding-report/"),
 "LM-SACRA": ("Sacra (third-party research) Lambda profile: ~$760M ARR Dec-2025, ~$425M 2024 revenue, Series E $5.9B post, site list (LAX01 21MW, EdgeConneX >30MW)", "2026-09", "https://sacra.com/c/lambda-labs/"),
 # Crusoe
 "CR-F": ("Crusoe PR: $3.9B Series F at $30.9B post-money; $140B+ TCV; 6GW+ gross contracted; ~1GW gross delivered & operational", "2026-09-17", "https://www.crusoe.ai/resources/newsroom/crusoe-announces-series-f-funding"),
 "CR-FDF": ("Crusoe Inc. Form D (filed 2026-09-17; equity; first sale 2026-08-31)", "2026-09-17", SEC + "1924674/000192467426000002/primary_doc.xml"),
 "CR-FDO": ("Crusoe Inc. Form D (filed 2026-09-17; options/rights to acquire securities; first sale 2026-07-17)", "2026-09-17", SEC + "1924674/000192467426000001/primary_doc.xml"),
 "CR-FDE": ("Crusoe Inc. Form D (filed 2025-10-21; equity; first sale 2025-10-08)", "2025-10-21", SEC + "1924674/000192467425000004/primary_doc.xml"),
 "CR-FDD": ("Crusoe Inc. Form D (filed 2024-11-20; equity)", "2024-11-20", SEC + "1924674/000192467424000003/primary_doc.xml"),
 "CR-EDGAR": ("EDGAR: Crusoe Inc. (CIK 0001924674) has only Form D filings; no S-1/DRS public", "2026-10-04", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001924674"),
 "CR-E": ("Crusoe PR: $1.375B Series E at >$10B; bookings 5x 9M'25; power pipeline >45GW", "2025-10-24", "https://www.crusoe.ai/resources/newsroom/crusoe-announces-series-e-funding"),
 "CR-BLOG": ("Crusoe blog: 2025 added TCV ~17x YoY, cloud ARR +150% YoY, new logos +70%", "2026-02-06", "https://www.crusoe.ai/resources/blog/building-the-worlds-favorite-ai-cloud"),
 "CR-MSFT": ("Crusoe PR: new 900MW Abilene AI factory campus for Microsoft (2 x 336MW critical IT; first building mid-2027)", "2026-03-27", "https://www.crusoe.ai/resources/newsroom/crusoe-announces-new-900-mw-ai-factory-campus-in-abilene-texas-to-support-microsoft-ai-infrastructure"),
 "CR-JV": ("Crusoe PR: second phase of $15B Abilene JV with Blue Owl and Primary Digital (1.2GW, 8 buildings, up to 50,000 GB200 per building)", "2025-05-21", "https://www.crusoe.ai/resources/newsroom/crusoe-blue-owl-capital-and-primary-digital-infrastructure-enter-joint-venture"),
 "CR-NMRK": ("Newmark PR: arranges $2.3B construction financing for 206MW build-to-suit data center (Abilene phase 1)", "2024", "https://www.nmrk.com/insights/press-releases/newmark-arranges-2-3-billion-construction-financing-for-206-mw-build-to-suit-data-center"),
 "CR-BBG71": ("Bloomberg: JPMorgan leads $7.1B loan for Blue Owl-tied (Abilene phase 2) data center", "2025-05-22", "https://www.bloomberg.com/news/articles/2025-05-22/jpmorgan-leads-7-1-billion-loan-for-blue-owl-tied-data-center"),
 "CR-BRK": ("Crusoe PR: $750M credit facility from Brookfield", "2025-06-11", "https://www.crusoe.ai/resources/newsroom/crusoe-secures-usd750-million-credit-facility-from-brookfield-to-accelerate"),
 "CR-AMD": ("DCD citing The Information: Goldman $300M loan for AMD chips (Springfield OH), AMD rent-back backstop, ~6% rate", "2026-02-20", "https://www.datacenterdynamics.com/en/news/amd-underwrites-300m-crusoe-loan-so-data-center-firm-can-buy-its-ai-chips-report/"),
 "CR-OCTUS": ("Octus: Crusoe launches ~$509M senior secured GPU DDTL (borrower IS AN03 LLC / Crusoe Iceland III), price talk S+475-500, 98.5 OID, Nov-2030", "2026-09-10", "https://octus.com/resources/articles/primary-crusoe-launches-509m-senior-secured-ddtl-due-november-2030-to-finance-gpu-capital-expenditures/"),
 "CR-HJ": ("HostingJournalist (relaying Crusoe announcement): 4.9GW contracted; >40GW pipeline; Abilene 2 buildings operational / 6 under construction", "2026-06-14", "https://hostingjournalist.com/news/crusoe-nears-5gw-in-contracted-ai-infrastructure-deals"),
 "CR-META": ("DCD citing Bloomberg: Meta contracts ~1.6GW at Crusoe Childress TX + Warrenton MO", "2026-06-19", "https://www.datacenterdynamics.com/en/news/meta-signs-16gw-capacity-agreement-with-crusoe-report/"),
 "CR-JS": ("DCD citing Bloomberg: Jane Street ~$13B five-year Crusoe Cloud contract", "2026-09-04", "https://www.datacenterdynamics.com/en/news/crusoe-signs-13bn-deal-with-jane-street-report/"),
 "CR-NEW": ("Newcomer (financials obtained): Crusoe 'on track for $2 billion in revenue this year'", "2026-09-23", "https://www.newcomer.co/p/crusoe-is-on-track-for-2-billion"),
 "CR-TI": ("The Logic citing The Information: 2024 revenue ~$250M; projections $500M (2025) / $2.2B (2026); Series D $600M at $2.8B", "2025-08-22", "https://thelogic.co/briefing/crusoe-seeks-us1b-at-us10b-valuation-for-more-data-centres/"),
 "CR-SACRA": ("Sacra (third-party research) Crusoe profile: 2024 revenue ~$276M estimate; Upper90 $225M and Victory Park $175M facilities", "2026-09", "https://sacra.com/c/crusoe/"),
 "SOFR": ("Federal Reserve Bank of New York SOFR, 3.87% (2026-10-01) - used only to ESTIMATE all-in coupons on floating loans", "2026-10-01", "https://markets.newyorkfed.org/api/rates/secured/sofr/last/3.json"),
 "NBPX": ("Nebius workspace: /workspace/nebius/parts/price_history_quarterly.csv and price_snapshot_2026-10-04.csv (first-party list pages / Wayback)", "2026-10-04", "local"),
 "NS-SITE": ("Peer site verification: Monarch Compute Campus, Mason County WV (Point Pleasant). /workspace/nebius/peers/sites/site_verification.md §7 + sites.xlsx. Active 0; contracted 460 IT MW; Grade Yellow (460 H1'28) / Yellow-Red (1.37 GW). Off-grid gas microgrid; air permit R14-0042 draft; Anthropic financing uncommitted.", "2026-10-05", "local:/workspace/nebius/peers/sites/"),
 "NS-EDGAR2": ("EDGAR re-check 2026-10-05: Nscale Ltd (CIK 0002110365) — no S-1/A, 424B4 or new 8-K since 2026-10-04. Latest registration remains S-1 filed 2026-09-18.", "2026-10-05", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0002110365"),
 "LM-EDGAR2": ("EDGAR re-check 2026-10-05: Lambda, Inc. (CIK 0001954966) — no S-1/DRS/Form D since 2026-10-04.", "2026-10-05", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001954966"),
 "CR-EDGAR2": ("EDGAR re-check 2026-10-05: Crusoe Inc. (CIK 0001924674) — no new Form D/S-1 since 2026-10-04. Series F Form D still 2026-09-17.", "2026-10-05", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001924674"),
 "CR-SITE": ("Peer site appendix: Abilene TX (Crusoe/Lancium Stargate), as of Oct 8. Active IT L/B/H 421/618/618 [D on R] per Oracle FQ1 FY27 call (2026-09-10: 6 of 8 bldgs, 618 MW, 75% delivered; IT vs gross basis unstated); YE26 618/721/824; Grade Green (B1-6) / Yellow (B7-8). Counted under Crusoe, not Oracle. Oct 5 view 200/421/421 is history.", "2026-10-08", "local:/workspace/nebius/peers/sites/appendix_crusoe_abilene.pdf ; https://www.fool.com/earnings/call-transcripts/2026/09/11/oracle-orcl-q1-2027-earnings-call-transcript/"),
 "CRWV": ("CoreWeave comps: /workspace/nebius/peers/coreweave/crwv_kpis.xlsx / crwv_kpis.md (Q2'26 column)", "2026-10-04", "local"),
}
REPORTED_KEYS = {"NS-S1","NS-EDGAR","NS-EDGAR2","NS-IPO","NS-C","NS-CV","NS-SITE","LM-EDGAR","LM-EDGAR2","LM-FD24","LM-IG","LM-TLB","LM-CF","LM-500","LM-E","LM-MSFT","LM-KC","LM-GTC","LM-TM","CR-F","CR-FDF","CR-FDO","CR-FDE","CR-FDD","CR-EDGAR","CR-EDGAR2","CR-E","CR-BLOG","CR-MSFT","CR-JV","CR-NMRK","CR-BRK","CR-SITE"}

DP = {"Nscale": [], "Lambda": [], "Crusoe": []}
def dp(co, i, row, metric, value, unit, asof, period, tag, src, note="", f=None, intags=""):
    DP[co].append(dict(id=i, row=row, metric=metric, value=value, unit=unit, asof=asof, period=period, tag=tag, src=src, note=note, f=f, intags=intags))

# ======================= NSCALE (S-1 = R) =======================
C = "Nscale"
dp(C,"N01","I01","Revenue, H1'26 (six months)",140.6,"$M","2026-06-30","6M to 2026-06-30","R","NS-S1","Service fee 94.7 / lease revenue (ASC 842) 42.0 / other 3.9")
dp(C,"N02","I01","Revenue, H1'25",10.4,"$M","2025-06-30","6M to 2025-06-30","R","NS-S1","")
dp(C,"N03","I01","Revenue, FY25",33.0,"$M","2025-12-31","FY2025","R","NS-S1","+73% vs FY24")
dp(C,"N04","I01","Revenue, FY24",19.1,"$M","2024-12-31","FY2024","R","NS-S1","93% from one customer")
dp(C,"N05","PX01","Annualized revenue run-rate (H1'26 x 2)",None,"$M/yr","2026-06-30","H1'26 annualized","D","NS-S1","Back-end-loaded ramp: H1 x2 understates exit run-rate", f="{N01}*2")
dp(C,"N06","PX01","Active TCV (contracts already generating revenue)",2600,"$M","2026-08-31","point","R","NS-S1","'approximately $2.6 billion of active' TCV (Dec-31-2025: $0.5B)")
dp(C,"N07","PX01","Active-contract ACV estimate (active TCV / 5.7-yr avg life)",None,"$M/yr","2026-08-31","point","E","NS-S1","ASSUMES active contracts share the 5.7-yr company-wide weighted life; NVIDIA MSA terms are 12-48 months, so true active ACV likely HIGHER", f="{N06}/{N28}")
dp(C,"N08","I02","Adjusted EBITDA (company non-GAAP), H1'26",-199.2,"$M","2026-06-30","6M","R","NS-S1","Margin (142)%; FY25 (81.7), FY24 (9.1)")
dp(C,"N09","I03","Depreciation & amortization (total), H1'26",174.0,"$M","2026-06-30","6M","R","NS-S1","FY25 40.2")
dp(C,"N10","I03","Depreciation of property & equipment, H1'26",77.9,"$M","2026-06-30","6M","R","NS-S1","Rest of D&A = intangible amortization (e.g. acquired power rights)")
dp(C,"N11","I04","Cash capex (purchases of property & equipment), H1'26",3230.1,"$M","2026-06-30","6M","R","NS-S1","Plus intangibles 15.6, Monarch acquisition 131.1, SIN02 96.8. FY25 621.4; FY24 64.2")
dp(C,"N12","I04","Cash capex (purchases of PP&E), FY25",621.4,"$M","2025-12-31","FY2025","R","NS-S1","")
dp(C,"N13","I05","Gross property & equipment (EoP)",5846.9,"$M","2026-06-30","point","R","NS-S1","Dec-25: 981.8")
dp(C,"N14","I05","Gross property & equipment (EoP), YE25",981.8,"$M","2025-12-31","point","R","NS-S1","")
dp(C,"N15","I06","Construction in progress (EoP)",4160.0,"$M","2026-06-30","point","R","NS-S1","Dec-25: 680.3")
dp(C,"N16","I06","Construction in progress (EoP), YE25",680.3,"$M","2025-12-31","point","R","NS-S1","")
dp(C,"N17","I07","Property & equipment, net (EoP)",5727.5,"$M","2026-06-30","point","R","NS-S1","Dec-25: 938.0")
dp(C,"N18","I08","Technology equipment, gross (EoP)",1306.5,"$M","2026-06-30","point","R","NS-S1","Dec-25: 105.6; 'primarily relates to our GPU and related infrastructure assets'")
dp(C,"N19","I05S","In-service gross PP&E (gross - CIP)",None,"$M","2026-06-30","point","D","NS-S1","", f="{N13}-{N15}")
dp(C,"N20","I05S","In-service gross PP&E (gross - CIP), YE25",None,"$M","2025-12-31","point","D","NS-S1","", f="{N14}-{N16}")
dp(C,"N21","I09","Active IT MW (sum of contracted IT MW at the 5 active sites)",None,"IT MW","2026-08-31","point","D","NS-S1","Facilities table: Portugal colo 30 + Norway colo 6 + Norway colo 5 + Iceland colo 15 + Norway owned 25. 'Active capacity' is DEFINED but not quantified in the S-1; this site sum is an UPPER BOUND on live MW", f="30+6+5+15+25")
dp(C,"N22","I10","Active + contracted IT MW",1370,"IT MW","2026-08-31","point","R","NS-S1","'approximately 1.37 GW' = 1 GW owned + 200 MW leased + 165 MW colo; 17 sites (5 active, 12 contracted)")
dp(C,"N23","I10","Contracted, not yet active IT MW",None,"IT MW","2026-08-31","point","D","NS-S1","", f="{N22}-{N21}")
dp(C,"N24","PX02","Potential power capacity (line of sight)",10000,"MW","2026-08-31","point","R","NS-S1","'line of sight to approximately 10 GW of potential power capacity'")
dp(C,"N25","PX02","Monarch Compute Campus (WV) initial phase, IT MW (expected H1'28)",1370,"IT MW","2026-03-31","point","R","NS-S1","Initial 2 GW gross / 1.37 GW IT by H1 2028; ~8 GW gross by 2031. Only 460 MW of it in contracted table; do not add to N22. Site status: see N60-N63 + peers/sites/")
dp(C,"N60","PX02","Monarch WV - contracted IT MW (S-1 facilities table / Anthropic slice)",460,"IT MW","2026-08-31","point","R","NS-S1","US Owned Tier 3 Contracted 460 MW in S-1 facilities table (= Anthropic Monarch slice). Part of N22; not additive.")
dp(C,"N61","I09","Monarch WV - evidenced-active IT MW (Peer KPIs site pack)",0,"IT MW","2026-10-05","point","R","NS-SITE","Sentinel-2 through 2026-09-25: civil/grading only (~85-100 ha bare); no hall roofs. Site manager 2026-09-20: about to go vertical. Active = 0.")
dp(C,"N62","PX10","Monarch WV - Anthropic contract ACV per IT MW (assumes 6-yr press term)",None,"$M/MW-yr","2026-08-25","contract avg","E","NS-S1","= 44,600 / 6 / 460. S-1 states $44.6B 'up to' and four tranches but not the term years; 6 yrs is press (Bloomberg/Reuters via WV MetroNews). Financing NOT committed.", f="{N35}*1000/6/{N60}", intags="P")
dp(C,"N63","PX02","Monarch WV - site grade (Peer KPIs)","Yellow (460 H1'28) / Yellow-Red (1.37 GW)","text","2026-10-05","point","R","NS-SITE","Off-grid gas (864x2.5 MW RICE); WVDEP air R14-0042 DRAFT (not final); High Impact certification 2026-09-03; land option at April permit filing. See peers/sites/site_verification.md §7.")
dp(C,"N26","I11","Active GPUs",25000,"GPUs","2026-08-31","point","R","NS-S1","'approximately 25,000 active GPUs'")
dp(C,"N27","PX04","Active + contracted GPUs",461000,"GPUs","2026-08-31","point","R","NS-S1","")
dp(C,"N28","PX05","Weighted average contract life",5.7,"years","2026-08-31","point","R","NS-S1","Take-or-pay; terms commence on delivery")
dp(C,"N29","PX05","Weighted-average customer prepayment (% of TCV)",23,"%","2026-08-31","point","R","NS-S1","")
dp(C,"N30","I14","RPO (ASC 606, EoP)",56.4,"$B","2026-06-30","point","R","NS-S1","Incl. $7.1B significant financing component; recognized over ~7 years")
dp(C,"N31","I14","RPO (ASC 606), YE25",34.5,"$B","2025-12-31","point","R","NS-S1","Incl. $4.6B SFC")
dp(C,"N32","I16","RPO % expected within 24 months (YE25)",16,"%","2025-12-31","point","R","NS-S1","+40% in years 3-4; not disclosed for 6/30/26")
dp(C,"N33","I15","Total contract value, active + contracted (company metric)",103.4,"$B","2026-08-31","point","R","NS-S1","Excl. extensions/renewals/SFC; Dec-31-2025 $38.0B. NOT ASC 606 RPO")
dp(C,"N34","PX12","Microsoft statements of work (Sep-2025 to Apr-2026)",43.8,"$B","2026-04-30","signed Sep-25..Apr-26","R","NS-S1","'up to approximately $43.8 billion through December 2033'; 5-6 yr GPU terms; Microsoft took OpenAI's Stargate Norway capacity after OpenAI withdrew (Apr-2026)")
dp(C,"N35","PX12","Anthropic GPU Services Agreements (Monarch, Vera Rubin NVL72)",44.6,"$B","2026-08-25","signed 2026-08-25","R","NS-S1","'up to approximately $44.6 billion'; four tranches; financing NOT yet committed; delivery/uptime termination rights")
dp(C,"N36","PX12","NVIDIA master cloud services agreement",1.2,"$B","2025-07-31","signed Jul-2025","R","NS-S1","12-48 month terms. Figure AI (2026-08-24): up to 100k Vera Rubin GPUs, value n/d")
dp(C,"N37","PX12","Microsoft + Anthropic share of active+contracted TCV",None,"%","2026-08-31","point","D","NS-S1","Approximate ('up to' amounts vs TCV)", f="({N34}+{N35})/{N33}*100")
dp(C,"N38","I17","Largest customer, % of revenue (H1'26)",52,"%","2026-06-30","6M","R","NS-S1","FY25 73% (same customer expected <20% of FY26); FY24 93%")
dp(C,"N39","I17","Largest customer, % of revenue (FY25)",73,"%","2025-12-31","FY2025","R","NS-S1","")
dp(C,"N40","I23","Total debt, carrying value (EoP)",137.4,"$M","2026-06-30","point","R","NS-S1","GPU Financing Facility 48.2 + Macquarie Senior 89.2. Post-June: +$329M GPU FF draw (Aug), $830M revolver utilized; see Debt_Facilities")
dp(C,"N41","I23","Total debt, carrying value, YE25",78.9,"$M","2025-12-31","point","R","NS-S1","Macquarie Senior only")
dp(C,"N42","I21","Interest expense on debt, H1'26",12.0,"$M","2026-06-30","6M","R","NS-S1","'Interest expense, net' 95.1 includes SFC interest on prepayments")
dp(C,"N43","I22","Capitalized interest, H1'26",0,"$M","2026-06-30","6M","R","NS-S1","FY25: 11.8")
dp(C,"N44","PX09","GPU Financing Facility effective interest rate",11.30,"%","2026-06-30","point","R","NS-S1","SOFR+5.00% (+0.25% step-up), $1.4B commitments, $1,363.1M unused")
dp(C,"N45","PX09","Macquarie Senior Credit Facility effective interest rate",22.74,"%","2026-06-30","point","R","NS-S1","SOFR+7.50% plus MOIC; 21.48% at YE25")
dp(C,"N46","K16","Useful life - technology equipment (midpoint of 5-6)",5.5,"years","2026-06-30","policy","R","NS-S1","Tech equipment 5-6 yrs; plant 5-15; buildings 25-50; computer equipment 3. Midpoint entered")
dp(C,"N47","PX06","Series B equity",1100,"$M","2025-10-08","round","R","NS-S1","Series B prefs at $6.33")
dp(C,"N48","PX06","Series C equity (S-1: $1.64B first close + $118.3M second close)",None,"$M","2026-04-30","round","D","NS-S1","PR headline $2B incl. Pre-Series C SAFE (N49)", f="1640+118.3")
dp(C,"N49","PX06","Series C headline (company PR)",2000,"$M","2026-03-09","round","R","NS-C","Valuation $14.6B")
dp(C,"N50","PX06","Series financing raised to date (S-1 wording)",3300,"$M","2026-09-18","cumulative","R","NS-S1","'raised over $3.3 billion through our series financing'")
dp(C,"N51","PX06","Pre-IPO convertible loan notes",3360,"$M","2026-09-25","round","R","NS-CV","$2.36B + $1.0B NVIDIA (mid-Nov-2026); S-1 states minimum $3.1B")
dp(C,"N52","PX07","Post-money valuation, Series C",14600,"$M","2026-03-09","round","R","NS-C","")
dp(C,"N53","PX07","IPO target valuation (press)",35000,"$M","2026-09","IPO talk","P","NS-FT","'as high as $35 billion' per FT; no S-1 price range yet; IPO ~Oct-15 per press, unconfirmed")
dp(C,"N54","PX08","Debt & lease-financing commitments signed after 6/30/26 (sum)",None,"$M","2026-09-04","Jul-Sep 2026","D","NS-S1","Kvandal 790 + Macquarie Iceland 331.9 + Ward County 1,850 + North Carolina 1,200 + Dell DFS initial-term rent ~2,540", f="790+331.9+1850+1200+2540")
dp(C,"N55","PX09","Marginal secured GPU debt (NC / Ward County, floating), all-in estimate",None,"%","2026-08-31","Aug-2026 facilities","E","NS-S1","3M Term SOFR + 2.375%; spot SOFR 3.87% (2026-10-01) as proxy", f="3.87+2.375")
dp(C,"N56","I27","Data center sites (active)",5,"#","2026-08-31","point","R","NS-S1","5 active + 12 contracted")
dp(C,"N57","PX03","MW basis","IT","text","2026-08-31","definition","R","NS-S1","Glossary: capacity in IT MW")
dp(C,"N58","I12","GPU list price (on-demand)",None,"$/GPU-hr","2026-10-04","point","N","NBPX","No public list price (pricing console-only)")
dp(C,"N59","PX01","Cash and cash equivalents",1478.4,"$M","2026-06-30","point","R","NS-S1","Plus restricted cash")

# ======================= LAMBDA (no S-1) =======================
C = "Lambda"
dp(C,"L01","I01","Revenue, Q2'25 (total incl. legacy hardware)",140,"$M","2025-06-30","Q2'25","P","LM-TI1","'almost 60%' YoY; documents viewed by The Information")
dp(C,"L02","I01","Revenue, H1'25",250,"$M","2025-06-30","H1'25","P","LM-TI1","'more than $250m', +33% YoY (lower bound)")
dp(C,"L03","I01","Cloud revenue, Q2'25",114,"$M","2025-06-30","Q2'25","P","LM-TI2","~$114M cloud revenue (The Information via Deep News)")
dp(C,"L04","I01","Revenue, FY24 (third-party estimate)",425,"$M","2024-12-31","FY2024","E","LM-SACRA","Sacra; company told investors 'closer to $400M' after guiding ~$600M")
dp(C,"L05","PX01","Annualized revenue run-rate, Dec-2025 (third-party estimate)",760,"$M/yr","2025-12-31","ARR","E","LM-SACRA","Sacra estimate")
dp(C,"L06","PX01","Expected 2026 revenue (press)",1500,"$M","2026-12-31","FY2026E","P","LM-BBG1","'more than $1.5 billion in revenue this year' - people familiar (lower bound)")
dp(C,"L07","PX01","Q2'25 revenue annualized (x4)",None,"$M/yr","2025-06-30","Q2'25 x4","D","LM-TI1","", f="{L01}*4", intags="P")
dp(C,"L08","I02","Gross margin H1'25 (blended; 61% ex non-cloud)",50,"%","2025-06-30","H1'25","P","LM-TI1","H1'25 net loss ~$24M. No EBITDA disclosed")
dp(C,"L09","I09","Active (operating) capacity, Q2'25",47,"MW","2025-06-30","Q2'25","P","LM-TI2","IT vs gross basis NOT stated")
dp(C,"L10","PX02","Capacity target 2030",3000,"MW","2025-09-05","target","P","LM-TI2","'nearly 3 gigawatts by 2030'")
dp(C,"L11","PX02","Kansas City AI factory: initial MW",24,"MW","2025-10-28","announced","R","LM-KC","Early 2026; potential >100 MW; >10,000 Blackwell Ultra GPUs; single customer, multi-year")
dp(C,"L12","PX02","Beacon Point (Hut 8, Nueces County TX) capacity for Anthropic",350,"MW","2026-09-01","contracted (press)","P","LM-ANTH","NVIDIA holds the Hut 8 lease (704 MW, $19.6B, 15 yrs per Hut 8); energization from Q1'27")
dp(C,"L13","PX02","Other leased sites (Prime LAX01 21 MW; EdgeConneX Chicago 23 MW + Atlanta)",51,"MW","2026-09","announced","E","LM-SACRA","21 + >30 (lower bound); third-party compilation")
dp(C,"L14","I10","Contracted / announced MW (company total)",None,"MW","2026-10-04","point","N","LM-EDGAR","No company total. Pieces: KC 24 (->100+), Beacon Point ~350 (press), LAX01 21, EdgeConneX >30")
dp(C,"L15","I11","GPUs on cloud platform",25000,"GPUs","2025-02-24","point","R","LM-TM","CEO: 'well over 25,000 GPUs on our cloud platform' (lower bound)")
dp(C,"L16","I11","GPUs in one GB300 AI factory (production)",10000,"GPUs","2026-03-16","point","R","LM-GTC","'10,000+ NVIDIA GB300 GPUs' (lower bound)")
dp(C,"L17","I27","US data centers",15,"#","2026-01-09","point","P","LM-350","'15 data centers across the US, according to its website'")
dp(C,"L18","PX12","NVIDIA leaseback: $1.3B / 10,000 GPUs / 4 yrs (+$200M / 8,000 GPUs)",1.5,"$B","2025-09-08","signed 2025","P","LM-NV","NVIDIA largest customer")
dp(C,"L19","PX13","Implied price, NVIDIA $1.3B / 10,000-GPU / 4-yr contract",None,"$/GPU-hr","2025-09-08","contract avg","D","LM-NV","= 1,300e6 / (10,000 x 4 x 8,760)", f="1300/(10000*4*8760)*1000000", intags="P")
dp(C,"L20","PX12","Microsoft multi-year agreement",None,"$B","2025-11-03","signed","R","LM-MSFT","'multibillion-dollar'; value/GPUs/term n/d")
dp(C,"L21","PX12","Anthropic cloud contract (six years, ~350 MW)",35,"$B","2026-09-01","signed (press)","P","LM-ANTH","WSJ via TheEnergyMag / Proactive (LM-ANTH2); not confirmed by Lambda")
dp(C,"L22","PX10","Anthropic contract ACV per MW",None,"$M/MW-yr","2026-09-01","contract avg","D","LM-ANTH","= 35,000 / 6 / 350", f="{L21}*1000/6/{L12}", intags="P")
dp(C,"L23","K26","Amazon + Microsoft share of cloud revenue, Q2'25",50,"%","2025-06-30","Q2'25","P","LM-TI2","'nearly half' (approx.)")
dp(C,"L24","PX06","Series C equity (Form D amount sold)",320.0,"$M","2023-10-05","round (first sale)","R","LM-FD24","$319,999,978 sold, 38 investors")
dp(C,"L25","PX06","Series D equity",480,"$M","2025-02-24","round","R","LM-TM","$2.5B post-money; total equity $863M then")
dp(C,"L26","PX06","Series E equity",1500,"$M","2025-11-18","round","R","LM-E","'over $1.5B', led by TWG Global + USIT (lower bound)")
dp(C,"L27","PX07","Post-money valuation, Series E (PitchBook)",5430,"$M","2025-11-18","round","P","LM-TC","Sacra cites $5.9B")
dp(C,"L28","PX07","Pre-IPO round in talks: up to $3B at $12B+",12000,"$M","2026-08-24","talks","P","LM-BBG1","Not closed")
dp(C,"L29","PX09","IG fixed-rate DDTL coupon",6.78,"%","2026-10-01","fixed coupon","R","LM-IG","$1.008B, A(low)/Baa1, matures 2033-05-30, amortizing")
dp(C,"L30","PX09","Term loan B all-in estimate (SOFR+3.00%, 99.5 OID)",None,"%","2026-08-27","floating","E","LM-TLB","$926M, Baa2, to 2030-12-31; = 3.87 + 3.00 + 0.5pt OID / ~4.3 yrs", f="3.87+3.00+0.5/4.3")
dp(C,"L31","PX08","2026 debt facilities announced (sum of commitments)",None,"$M","2026-10-01","2026","D","LM-CF","1,000 (May) + 926 (Aug) + 1,008 (Oct). Excludes Bloomberg ~$1B Aug-28 deal (possible overlap) and 2024 $500M", f="1000+926+1008")
dp(C,"L32","I23","Total debt outstanding",None,"$M","2026-10-04","point","N","LM-EDGAR","No balance sheet public")
dp(C,"L33","I14","RPO / backlog",None,"$B","2026-10-04","point","N","LM-EDGAR","Not disclosed")
dp(C,"L34","K16","Useful life - GPUs",None,"years","2026-10-04","policy","N","LM-TLB","Not disclosed; TLB amortization 'aligned with' GPU useful life (~4.3 yrs from close)")
dp(C,"L35","I12","H100 SXM on-demand list, Q2'25",2.99,"$/GPU-hr","2025-06-16","snapshot","R","NBPX","Now $3.99; B200 $6.69 (2026-10-04)")
dp(C,"L37","PX07","Pre-IPO round in talks: up to $4B at $14.5B PRE-money (Blackstone + Coatue)",14500,"$M","2026-10-06","talks","P","LM-WSJ4B","Supersedes L28 ($3B at $12B+). Not closed; ~$18.5B post if fully raised (E). IPO now targeted 2027 per investor letter")
dp(C,"L38","I15","Backlog (company metric, investor letter), Sep-2026",50,"$B","2026-09-30","point","P","LM-WSJ4B","Not ASC 606 RPO; ~$35B of the jump is Anthropic (L21). Term/schedule n/d")
dp(C,"L39","I15","Backlog (company metric, investor letter), Jun-2026",15,"$B","2026-06-30","point","P","LM-WSJ4B","Pre-Anthropic baseline")
dp(C,"L36","PX01","Lambda S-1 status","none filed","text","2026-10-04","check","R","LM-EDGAR","Only Form D (2022, 2024)")

# ======================= CRUSOE (no S-1) =======================
C = "Crusoe"
dp(C,"C01","I01","Revenue, FY24",250,"$M","2024-12-31","FY2024","P","CR-TI","Sacra estimate $276M")
dp(C,"C02","I01","Revenue, FY25 (projection as of Aug-2025)",500,"$M","2025-12-31","FY2025E","P","CR-TI","Actual FY25 n/d")
dp(C,"C03","PX01","Revenue, FY26 ('on track for $2 billion')",2000,"$M","2026-12-31","FY2026E","P","CR-NEW","Aug-2025 projection $2.2B (CR-TI); Sacra: $3.6B 2027 / $5.5B 2028 projections")
dp(C,"C04","PX01","Managed Inference contracted ARR",100,"$M","2026-09-17","point","R","CR-F","'over $100M' (lower bound)")
dp(C,"C05","PX01","Bookings growth YTD 2026 (YoY)",20,"x","2026-09-17","YTD","R","CR-F","'20x+'; 2025 TCV added ~17x, cloud ARR +150% (CR-BLOG)")
dp(C,"C06","I09","Delivered & operational capacity (gross)",1000,"gross MW","2026-09-17","point","R","CR-F","'1 GW of gross capacity'; includes JV/landlord capacity (Abilene for Oracle/OpenAI)")
dp(C,"C07","PX03","IT/gross ratio (Microsoft Abilene design: 672 IT / 900 MW)",None,"x","2026-03-27","design","D","CR-MSFT","", f="672/900")
dp(C,"C08","I09","Delivered capacity, IT-equivalent",None,"IT MW","2026-09-17","point","E","CR-F","ESTIMATE = gross x 0.747", f="{C06}*{C07}")
dp(C,"C09","I10","Gross contracted capacity",6000,"gross MW","2026-09-17","point","R","CR-F","'6GW+' (lower bound); Jun-2026 4.9 GW (C10)")
dp(C,"C10","I10","Contracted capacity (Jun-2026)",4900,"MW","2026-06-14","point","P","CR-HJ","")
dp(C,"C11","PX02","Power / development pipeline",45000,"MW","2025-10-24","point","R","CR-E","'over 45 gigawatts'; Jun-2026 '>40GW' (CR-HJ)")
dp(C,"C12","PX02","Abilene (Oracle/OpenAI) campus, total",1200,"MW","2025-05-21","8 buildings","R","CR-JV","Phase 1: 2 bldgs 200+ MW (H1'25); phase 2: 6 bldgs, completion end-2026")
dp(C,"C37","I09","Abilene TX - evidenced-active IT MW (Peer KPIs site appendix; count under Crusoe not Oracle)",618,"IT MW","2026-09-10","point","D","CR-SITE","[D on R] Oracle FQ1 FY27 call (Sep 10): 6 of 8 buildings, 618 MW, 75% of capacity delivered to the customer; delivered + accepted read as active. IT vs gross basis unstated. Low/Base/High 421/618/618 (low = Epoch satellite model, P); YE26 618/721/824. Grade Green B1-6 / Yellow B7-8. History: Oct 5 value 421 (L/B/H 200/421/421). GPUs are Oracle/OpenAI's - Crusoe is landlord/builder.", intags="R")
dp(C,"C13","PX02","Abilene Microsoft campus",900,"MW","2026-03-27","announced","R","CR-MSFT","2 x 336 MW critical IT; first building mid-2027")
dp(C,"C14","PX02","Meta capacity (Childress TX + Warrenton MO)",1600,"MW","2026-06-19","contracted (press)","P","CR-META","Value/term n/d")
dp(C,"C15","I15","Total contracted value (company metric)",140,"$B","2026-09-17","point","R","CR-F","'$140B+'; term n/d; not RPO")
dp(C,"C16","PX11","TCV per gross contracted MW",None,"$M/MW","2026-09-17","point","D","CR-F","= 140,000 / 6,000 (both lower bounds)", f="{C15}*1000/{C09}")
dp(C,"C17","PX12","Jane Street Crusoe Cloud contract (five years)",13,"$B","2026-09-04","signed (press)","P","CR-JS","")
dp(C,"C18","PX12","Jane Street contract ACV",None,"$B/yr","2026-09-04","contract avg","D","CR-JS","= 13 / 5", f="{C17}/5", intags="P")
dp(C,"C19","PX12","Anchor customers (named)","Oracle/OpenAI, Microsoft, Meta (press), Jane Street (press)","text","2026-09-17","point","R","CR-MSFT","Concentration % n/d")
dp(C,"C20","I11","GPU count (Crusoe Cloud fleet)",None,"GPUs","2026-10-04","point","N","CR-F","Not disclosed. Abilene GB200s are Oracle/OpenAI's")
dp(C,"C21","PX06","Series F equity (Form D amount sold)",3108.5,"$M","2026-09-17","round","R","CR-FDF","$3,108,543,889 of $3,901,017,922 offered; 108 investors")
dp(C,"C22","PX06","Series E equity (Form D amount sold)",1027.4,"$M","2025-10-21","round","R","CR-FDE","Of $1,374,999,988 offered")
dp(C,"C23","PX06","Series D equity (Form D amount sold)",685.7,"$M","2024-11-20","round","R","CR-FDD","Of $817,744,542 offered; press: $600M at $2.8B")
dp(C,"C24","PX06","Form D: options/rights to acquire securities",595.9,"$M","2026-09-17","offering","R","CR-FDO","Nature not described")
dp(C,"C25","PX06","Equity raised 2024-2026 (Form D D+E+F)",None,"$M","2026-09-17","cumulative","D","CR-FDF","Excl. C24 and pre-2024", f="{C21}+{C22}+{C23}")
dp(C,"C26","PX07","Post-money valuation, Series F",30900,"$M","2026-09-17","round","R","CR-F","Series E >$10B")
dp(C,"C27","PX14","Valuation / FY26E revenue",None,"x","2026-09-17","point","E","CR-F","= 30,900 / 2,000 (press revenue)", f="{C26}/{C03}", intags="P")
dp(C,"C28","PX09","Iceland GPU DDTL all-in estimate (price talk midpoint)",None,"%","2026-09-10","price talk","E","CR-OCTUS","~$509M, Nov-2030; final pricing n/d. = 3.87 + 4.875 + 1.5pt OID / ~4.2 yrs", f="3.87+4.875+1.5/4.2", intags="P")
dp(C,"C29","PX09","AMD-backstopped Goldman GPU loan rate",6,"%","2026-02-20","loan","P","CR-AMD","$300M")
dp(C,"C30","PX08","Corporate facilities (Brookfield 750 + Upper90 225 + Victory Park 175)",None,"$M","2025-06-11","facilities","D","CR-BRK","Upper90/Victory Park per Sacra", f="750+225+175", intags="R/E")
dp(C,"C31","PX08","Abilene JV project financing ($2.3B ph1 + $7.1B ph2)",None,"$M","2025-05-22","JV-level","D","CR-BBG71","JV-level debt", f="2300+7100", intags="R/P")
dp(C,"C32","K14","Abilene JV total cost per MW (shell + power, ex-GPU)",None,"$M/MW","2025-05-21","project","D","CR-JV","= $15B / 1,200 MW", f="15000/{C12}")
dp(C,"C33","K14","Abilene phase-1 construction debt per MW",None,"$M/MW","2024","project","D","CR-NMRK","= $2.3B / 206 MW", f="2300/206")
dp(C,"C34","I23","Total debt outstanding",None,"$M","2026-10-04","point","N","CR-EDGAR","No balance sheet public")
dp(C,"C35","K16","Useful life - GPUs",None,"years","2026-10-04","policy","N","CR-EDGAR","Not disclosed")
dp(C,"C36","I12","H100 SXM on-demand list",3.90,"$/GPU-hr","2026-10-04","snapshot","R","NBPX","H200 $4.29; GB200 contact sales")

# ======================= derived KPI points (CRWV K-rows) =======================
def kp(co, i, row, metric, unit, asof, period, tag, f, note, src, intags=""):
    dp(co, i, row, metric, None, unit, asof, period, tag, src, note, f=f, intags=intags)
C = "Nscale"
kp(C,"NK01","K01","Revenue per active MW (H1'26 x2 / active-site MW)","$M/MW-yr","2026-06-30","H1'26 ann. / Aug-26 MW","D","{N05}/{N21}","MW is an upper bound and later than revenue period -> understates","NS-S1")
kp(C,"NK01b","K01","Active-contract ACV per active MW","$M/MW-yr","2026-08-31","point","E","{N07}/{N21}","Uses N07","NS-S1")
kp(C,"NK03","K03","Revenue per contracted MW (H1'26 x2 / 1,370 MW)","$M/MW-yr","2026-08-31","point","D","{N05}/{N22}","94% of MW not yet live","NS-S1")
kp(C,"NK03b","PX10","Contracted ACV per contracted MW (TCV / 5.7 yrs / 1,370 MW)","$M/MW-yr","2026-08-31","point","E","{N33}*1000/{N28}/{N22}","Assumes TCV spread evenly over 5.7-yr life","NS-S1")
kp(C,"NK03c","PX11","TCV per active+contracted MW","$M/MW","2026-08-31","point","D","{N33}*1000/{N22}","","NS-S1")
kp(C,"NK04","K04","Contracted (active+contracted) / active MW","x","2026-08-31","point","D","{N22}/{N21}","","NS-S1")
kp(C,"NK05","K05","Adj. EBITDA per active MW","$M/MW-yr","2026-06-30","H1'26 ann.","D","{N08}*2/{N21}","Negative: pre-scale","NS-S1")
kp(C,"NK06","K06","Revenue per GPU (H1'26 x2 / Aug-26 active GPUs)","$K/GPU-yr","2026-06-30","H1'26 ann.","E","{N05}*1000/{N26}","Aug-31 GPU count used as H1 average","NS-S1")
kp(C,"NK07","K07","Implied blended $/GPU-hr (H1'26)","$/GPU-hr","2026-06-30","H1'26","E","{N01}*1000000/({N26}*181*24)","181 days; same GPU caveat","NS-S1")
kp(C,"NK07b","K07","Active-contract ACV per active GPU-hour","$/GPU-hr","2026-08-31","point","E","{N07}*1000000/({N26}*8760)","Run-rate view","NS-S1")
kp(C,"NK07c","PX13","Contracted TCV per GPU-hour (TCV / 461k GPUs / 5.7 yrs)","$/GPU-hr","2026-08-31","point","E","{N33}*1e9/{N27}/{N28}/8760","Vera Rubin-era contracts","NS-S1")
kp(C,"NK14","K14","Net PP&E per active MW","$M/MW","2026-06-30","Jun PP&E / Aug MW","D","{N17}/{N21}","Includes $4.16B CIP for contracted sites -> overstated","NS-S1")
kp(C,"NK15","K15","Gross in-service PP&E per active MW","$M/MW","2026-06-30","Jun PP&E / Aug MW","D","{N19}/{N21}","Active base mostly colocation","NS-S1")
kp(C,"NK15b","K15","In-service technology equipment per active GPU","$K/GPU","2026-06-30","Jun / Aug","D","{N18}*1000/{N26}","","NS-S1")
kp(C,"NK17","K17","Depreciation yield (PP&E depreciation x2 / avg in-service gross PP&E)","%","2026-06-30","H1'26","D","{N10}*2/(({N19}+{N20})/2)*100","Total D&A incl. intangible amortization would give ~35%","NS-S1")
kp(C,"NK18","K18","Implied average depreciable life (1/yield)","years","2026-06-30","H1'26","D","100/{NK17}","","NS-S1")
kp(C,"NK19","K19","D&A as % of revenue (H1'26)","%","2026-06-30","H1'26","D","{N09}/{N01}*100","","NS-S1")
kp(C,"NK23","K23","RPO / annualized revenue","years","2026-06-30","point","D","{N30}*1000/{N05}","","NS-S1")
kp(C,"NK28","K28","Weighted effective cost of debt (6/30/26)","%","2026-06-30","point","D","(48.2*{N44}+89.2*{N45})/{N40}","Carrying-value weights","NS-S1")
kp(C,"NK35","K35","GPU-backed share of debt","%","2026-06-30","point","D","48.2/{N40}*100","","NS-S1")
kp(C,"NK36","K36","Interest on debt x2 / avg debt","%","2026-06-30","H1'26","D","{N42}*2/(({N40}+{N41})/2)*100","","NS-S1")
kp(C,"NK38","K38","Total debt (EoP)","$B","2026-06-30","point","D","{N40}/1000","","NS-S1")
kp(C,"NK39","K39","Debt / RPO","x","2026-06-30","point","D","{N40}/1000/{N30}","","NS-S1")
kp(C,"NK40","K40","Capitalized interest % of interest (H1'26)","%","2026-06-30","H1'26","D","{N43}/{N42}*100","FY25: 11.8/12.9 = 91%","NS-S1")
C = "Lambda"
kp(C,"LK01","K01","Cloud revenue per active MW (Q2'25 x4 / 47 MW)","$M/MW-yr","2025-06-30","Q2'25","D","{L03}*4/{L09}","Press inputs; MW basis unknown","LM-TI2","P")
kp(C,"LK01b","K01","Total revenue per active MW (Q2'25 x4 / 47 MW)","$M/MW-yr","2025-06-30","Q2'25","D","{L01}*4/{L09}","Includes legacy hardware","LM-TI1","P")
kp(C,"LK06","K06","Cloud revenue per GPU (Q2'25 x4 / 25,000 GPUs)","$K/GPU-yr","2025-06-30","Q2'25","E","{L03}*4000/{L15}","Feb-2025 lower-bound GPU count used -> overstates","LM-TM","P/R")
kp(C,"LK07","K07","Implied blended cloud $/GPU-hr (Q2'25)","$/GPU-hr","2025-06-30","Q2'25","E","{L03}*1000000/({L15}*91*24)","Same caveat","LM-TM","P/R")
kp(C,"LK08","K08","List price proxy (H100 SXM on-demand, Q2'25)","$/GPU-hr","2025-06-16","snapshot","R","{L35}","Single-SKU proxy","NBPX")
kp(C,"LK09","K09","Implied utilization (realized / list)","%","2025-06-30","Q2'25","E","{LK07}/{LK08}*100","Price-realization x utilization index (D3)","NBPX","P/R")
C = "Crusoe"
kp(C,"CK01","K01","FY26E revenue per delivered gross MW","$M/MW-yr","2026-09-17","FY26E / Sep MW","E","{C03}/{C06}","Mixes landlord and cloud revenue; includes JV capacity","CR-NEW","P/R")
kp(C,"CK01b","K01","FY26E revenue per delivered IT-equivalent MW","$M/MW-yr","2026-09-17","FY26E / Sep MW","E","{C03}/{C08}","Plus IT conversion estimate","CR-NEW","P/E")
kp(C,"CK03","K03","FY26E revenue per gross contracted MW","$M/MW-yr","2026-09-17","point","E","{C03}/{C09}","","CR-NEW","P/R")
kp(C,"CK04","K04","Gross contracted / delivered MW","x","2026-09-17","point","D","{C09}/{C06}","","CR-F")

VAL = {d["id"]: d["value"] for co in DP for d in DP[co] if not d["f"]}
PEND = [d for co in DP for d in DP[co] if d["f"]]
while PEND:  # resolve formulas in dependency order
    nxt = []
    for d in PEND:
        refs = re.findall(r"\{(\w+)\}", d["f"])
        if all(x in VAL for x in refs):
            VAL[d["id"]] = eval(re.sub(r"\{(\w+)\}", lambda m: repr(VAL[m.group(1)]), d["f"]))
        else:
            nxt.append(d)
    assert len(nxt) < len(PEND), "circular formula refs"
    PEND = nxt

FAC = [
 ("Nscale","Macquarie Senior Credit Facility (Glomfjord GPUs)","2025-06",105.0,"SOFR+7.50% + MOIC; eff. 22.74% (6/30/26)",22.74,"R","n/d","Carrying 89.2 at 6/30/26","NS-S1"),
 ("Nscale","GPU Financing Facility (senior secured DDTL)","2026-02",1400.0,"Term SOFR+5.00% (+0.25% step-up); eff. 11.30%",11.30,"R","n/d","Carrying 48.2 at 6/30/26; +$329M drawn Aug-26; +$350M incremental option","NS-S1"),
 ("Nscale","Revolving Credit Facility (upsized from $770M)","2026-05",900.0,"Term SOFR+1.75%",SOFR+1.75,"E (spread R)","n/d","Undrawn 6/30/26; $830M utilized since","NS-S1"),
 ("Nscale","Dell Financial Services equipment lease framework","2026",2540.0,"Financing charges ~$34.8M on ~$2.54B initial-term rent",None,"R","per schedule","As of 2026-09-04","NS-S1"),
 ("Nscale","Kvandal South DC Facility ($725M TL + $65M VAT)","2026-07",790.0,"TL Term SOFR+3.00-3.75%; VAT SOFR+2.00%",SOFR+3.375,"E (spread R)","n/d","DC capex, Norway","NS-S1"),
 ("Nscale","Macquarie Iceland Facility","2026-07",331.9,"SOFR+5.50%",SOFR+5.50,"E (spread R)","46 months","GPU purchase","NS-S1"),
 ("Nscale","Ward County GPU Facility ($630M fixed / $1,220M floating)","2026-08",1850.0,"SOFR+2.375%",SOFR+2.375,"E (spread R)","2033-02","GPU capex","NS-S1"),
 ("Nscale","North Carolina GPU Facility ($370M fixed / $830M floating)","2026-08",1200.0,"SOFR+2.375%",SOFR+2.375,"E (spread R)","2031-12","GPU capex","NS-S1"),
 ("Nscale","Pre-IPO convertible loan notes","2026-09-25",3360.0,"n/d (convert at IPO)",None,"R","converts at IPO","$2.36B + $1.0B NVIDIA mid-Nov","NS-CV"),
 ("Nscale","Series A convertible notes (~$155M + $10M; converted)","2024",165.0,"12% PIK",12.0,"R","converted","Historical","NS-S1"),
 ("Nscale","Sandton term loans (repaid Oct-2025)","2023",None,"13-15%",14.0,"R","repaid","Historical (midpoint)","NS-S1"),
 ("Lambda","GPU-backed SPV facility (Macquarie)","2024-04-02",500.0,"n/d",None,"R","n/d","","LM-500"),
 ("Lambda","Credit facility (Aug-2025)","2025-08",275.0,"n/d",None,"R","n/d","Upsized to $1B May-2026","LM-CF"),
 ("Lambda","Senior secured credit facility (JPM, multi-tranche)","2026-05-07",1000.0,"n/d",None,"R","n/d","","LM-CF"),
 ("Lambda","Senior secured Term Loan B (IG offtaker)","2026-08-27",926.0,"SOFR+3.00%, 99.5 OID; Baa2",SOFR+3.00+0.5/4.3,"E (spread R)","2030-12-31","Fully amortizing; MS lead","LM-TLB"),
 ("Lambda","Private short-dated debt (JPM, Microsoft-linked GPUs)","2026-08-28",1000.0,"n/d",None,"P","short-dated","May overlap Oct-1 DDTL - not summed","LM-BBG2"),
 ("Lambda","IG fixed-rate DDTL (2 IG offtakers)","2026-10-01",1008.0,"6.78% fixed, semi-annual; A(low)/Baa1",6.78,"R","2033-05-30","Fully amortizing; JPM","LM-IG"),
 ("Crusoe","Upper90 credit facility","n/d",225.0,"n/d",None,"E","n/d","Per Sacra","CR-SACRA"),
 ("Crusoe","Victory Park Capital facility","n/d",175.0,"n/d",None,"E","n/d","Per Sacra","CR-SACRA"),
 ("Crusoe","Brookfield credit facility","2025-06-11",750.0,"n/d",None,"R","n/d","","CR-BRK"),
 ("Crusoe","Abilene phase-1 construction loan (JV)","2024",2300.0,"n/d",None,"R","n/d","206 MW; JV-level","CR-NMRK"),
 ("Crusoe","Abilene phase-2 loan (JV; JPM)","2025-05-22",7100.0,"n/d",None,"P","n/d","JV-level","CR-BBG71"),
 ("Crusoe","Goldman GPU loan, AMD backstop (Ohio)","2026-02-20",300.0,"~6%",6.0,"P","n/d","","CR-AMD"),
 ("Crusoe","Iceland GPU DDTL (IS AN03 LLC)","2026-09-10",509.0,"Price talk SOFR+475-500, 0% floor, 98.5 OID",SOFR+4.875+1.5/4.2,"E (talk P)","2030-11","<=75% LTC, 1.30x DSCR; GS; final n/d","CR-OCTUS"),
]

# ======================= Summary mapping (CoreWeave row layout) =======================
CRWV_Q226 = {  # reference only, from crwv_kpis.md (Q2'26)
 "I09": ("1,500", "R"), "I10": ("3,700", "R"), "I14": ("103.7", "R"), "K01": ("6.9", "D"), "K02": ("8.2", "D"), "K03": ("2.8", "D"), "K05": ("4.8", "D"),
 "K06": ("14.5", "E"), "K07": ("1.66", "E"), "K09": ("~18", "E"), "K13": ("20.0", "D"), "K14": ("31.2", "D"), "K16": ("6", "R"), "K17": ("15.5", "D"),
 "K20": ("103.7", "R"), "K22": ("41", "R"), "K23": ("10.1", "D"), "K24": ("36", "R"), "K27": ("3", "R"), "K28": ("8.4", "D"), "K36": ("7.8", "D"),
 "PX10": ("Nebius new deals ~20-25 (ref)", "E"),
}
ROWS = [
 ("Inputs (latest dated point)", None),
 ("I01","Revenue (latest reported period)","$M"), ("PX01","Revenue run-rate / forward revenue","$M/yr"), ("I02","Adj. EBITDA / margin proxy","$M or %"),
 ("I03","D&A (latest period)","$M"), ("I04","Cash capex (latest period)","$M"), ("I05","Gross PP&E (EoP)","$M"), ("I06","Construction in progress","$M"), ("I07","Net PP&E","$M"), ("I08","Technology equipment, gross","$M"),
 ("I09","Active MW","MW"), ("I10","Contracted MW (active + contracted)","MW"), ("PX02","Announced / pipeline MW","MW"), ("PX03","MW basis (IT vs gross)",""),
 ("I11","GPU count (active)","GPUs"), ("PX04","GPUs active + contracted","GPUs"), ("I12","On-demand list price (H100-class)","$/GPU-hr"),
 ("I14","RPO (ASC 606)","$B"), ("I15","Backlog / TCV (company metric)","$B"), ("I16","RPO % within 24 months","%"), ("PX05","Weighted avg contract life","years"),
 ("I17","Largest customer % revenue","%"), ("PX12","Largest disclosed anchor contract","$B"),
 ("I21","Interest expense on debt (latest period)","$M"), ("I22","Capitalized interest","$M"), ("I23","Total debt (EoP)","$M"), ("I27","Data center sites","#"),
 ("PX06","Equity raised (latest round)","$M"), ("PX07","Latest post-money valuation","$M"), ("PX08","Recent debt commitments (sum)","$M"),
 ("PX09","Marginal secured debt cost (latest facility)","%"),
 ("A. Revenue per MW", None),
 ("K01","Revenue per active MW","$M/MW-yr"), ("K02","Revenue per active MW (average MW)","$M/MW-yr"), ("K03","Revenue per contracted MW","$M/MW-yr"), ("PX10","Contract ACV per contracted MW","$M/MW-yr"), ("PX11","TCV per contracted MW","$M/MW"),
 ("K04","Contracted / active MW","x"), ("K05","Adj. EBITDA per active MW","$M/MW-yr"),
 ("B. Revenue per GPU and implied utilization", None),
 ("K06","Revenue per GPU","$K/GPU-yr"), ("K07","Implied blended $/GPU-hr","$/GPU-hr"), ("PX13","Contract-implied $/GPU-hr","$/GPU-hr"), ("K08","List price (fleet-weighted / proxy)","$/GPU-hr"), ("K09","Implied utilization","%"),
 ("C. Capex per MW", None),
 ("K10","Cash capex per new active MW","$M/MW"), ("K11","Cash capex per new MW incl. estimated MW","$M/MW"), ("K12","In-service additions per new active MW","$M/MW"), ("K13","Trailing-4Q cash capex per new active MW","$M/MW"),
 ("K14","Net PP&E (or project cost) per active MW","$M/MW"), ("K15","Gross in-service PP&E per active MW","$M/MW"),
 ("D. D&A useful lives and depreciation yield", None),
 ("K16","Useful life - technology equipment","years"), ("K17","Depreciation yield","%"), ("K18","Implied depreciable life","years"), ("K19","D&A % revenue","%"),
 ("E. Backlog / RPO", None),
 ("K20","RPO (ASC 606)","$B"), ("K21","Backlog / TCV (company metric)","$B"), ("K22","RPO % within 24 months","%"), ("K23","RPO / annualized revenue","years"),
 ("F. Customer concentration", None),
 ("K24","Largest customer % revenue","%"), ("K25","2nd-largest customer %","%"), ("K26","Top-N customers combined %","%"), ("K27","# customers >=10%","#"),
 ("G. Funding cost", None),
 ("K28","Weighted effective cost of debt","%"), ("K29","  DDTL / GPU-backed rate","%"), ("K30","  Senior notes rate","%"), ("K31","  Converts rate","%"), ("K32","  OEM financing rate","%"), ("K33","  Revolver rate","%"), ("K34","  Other term loans rate","%"),
 ("K35","GPU-backed share of debt","%"), ("K36","Interest / avg debt","%"), ("K37","Company-reported weighted rate","%"), ("K38","Total debt","$B"), ("K39","Debt / RPO","x"), ("K40","Capitalized interest % of interest","%"),
 ("Valuation (private-peer extension)", None),
 ("PX14","Valuation / forward revenue","x"),
]
def ND(why): return ("ND", why)
MAP = {
 "Nscale": {"I01":"N01","PX01":"N05","I02":"N08","I03":"N09","I04":"N11","I05":"N13","I06":"N15","I07":"N17","I08":"N18","I09":"N21","I10":"N22","PX02":"N24","PX03":"N57",
   "I11":"N26","PX04":"N27","I12":"N58","I14":"N30","I15":"N33","I16":"N32","PX05":"N28","I17":"N38","PX12":"N35","I21":"N42","I22":"N43","I23":"N40","I27":"N56",
   "PX06":"N51","PX07":"N52","PX08":"N54","PX09":"N55",
   "K01":"NK01","K02":ND("Single MW point (Aug-26)"),"K03":"NK03","PX10":"NK03b","PX11":"NK03c","K04":"NK04","K05":"NK05",
   "K06":"NK06","K07":"NK07","PX13":"NK07c","K08":ND("No public list price"),"K09":ND("No public list price"),
   "K10":ND("No MW history"),"K11":ND("No MW history"),"K12":ND("No MW history"),"K13":ND("No MW history"),"K14":"NK14","K15":"NK15",
   "K16":"N46","K17":"NK17","K18":"NK18","K19":"NK19","K20":"N30","K21":"N33","K22":"N32","K23":"NK23",
   "K24":"N38","K25":ND("Only largest customer disclosed"),"K26":ND("Only largest customer disclosed"),"K27":ND("Not disclosed"),
   "K28":"NK28","K29":"N44","K30":ND("None outstanding"),"K31":ND("Converts post-period"),"K32":ND("Dell DFS leases post-period"),"K33":ND("Revolver undrawn at 6/30"),"K34":"N45","K35":"NK35","K36":"NK36","K37":ND("Not disclosed"),"K38":"NK38","K39":"NK39","K40":"NK40",
   "PX14":ND("No IPO price range yet (press: ~$35B target vs $281M run-rate)")},
 "Lambda": {"I01":"L01","PX01":"L06","I02":"L08","I09":"L09","I10":"L14","PX02":"L10","I11":"L15","I12":"L35","I14":"L33","PX12":"L21","I23":"L32","I27":"L17",
   "PX06":"L26","PX07":"L27","I15":"L38","K21":"L38","PX08":"L31","PX09":"L29","K01":"LK01","PX10":"L22","K06":"LK06","K07":"LK07","PX13":"L19","K08":"LK08","K09":"LK09",
   "K16":"L34","K20":"L33","K26":"L23","K29":"L29","K38":"L32","I17":ND("Not disclosed (press: NVIDIA largest after 2025 leaseback)"),
   "PX03":ND("Basis never stated"),"PX14":ND("Not computed (no reliable current revenue; Sacra ARR $760M vs $5.43B Series E = 7.1x)")},
 "Crusoe": {"I01":"C01","PX01":"C03","I09":"C06","I10":"C09","PX02":"C11","PX03":"C07","I11":"C20","I12":"C36","I15":"C15","PX12":"C17","I23":"C34",
   "PX06":"C21","PX07":"C26","PX08":"C30","PX09":"C28","K01":"CK01","K03":"CK03","PX11":"C16","K04":"CK04","K08":"C36","K14":"C32",
   "K16":"C35","K21":"C15","K29":"C28","K34":"C29","K38":"C34","PX14":"C27","PX10":ND("TCV term not disclosed")},
}

# ======================= workbook =======================
wb = openpyxl.Workbook()
BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="D9E1F2"); SECF = PatternFill("solid", fgColor="F2F2F2")
TF = {"R": Font(color="0000CC"), "P": Font(color="7030A0"), "D": Font(color="000000"), "E": Font(color="C65911"), "N": Font(color="808080", italic=True)}
wrap = Alignment(wrap_text=True, vertical="top")
def header(ws, row, cols, widths):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=h); c.font = BOLD; c.fill = HDR; c.alignment = wrap
    for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w
def tagtxt(d): return ("n/d" if d["tag"] == "N" else d["tag"]) + (f" (inputs {d['intags']})" if d["intags"] else "")

ws = wb.active; ws.title = "README"
README = [
 "Private peer KPI profiles - Nscale, Lambda, Crusoe - built 2026-10-04 (ET) for the Nebius (NBIS) research project. Builder: build_private_peers.py; write-up: private_peers.md (same folder).",
 "TEMPLATE: same row IDs and KPI definitions as /workspace/nebius/peers/coreweave/crwv_kpis.xlsx (Inputs I01-I27, KPIs K01-K40, definitions D1-D8). Private-peer-only rows: PX01-PX14 (Definitions tab).",
 "LAYOUT: company tabs hold DATED data points (as-of date, period, tag, source doc, source date, URL). D/E points are live Excel formulas over other points (Formula column shows DP ids). Summary links the latest point per company into the CoreWeave row layout, with CoreWeave Q2'26 for reference. Debt_Facilities lists every disclosed facility; Sources is the registry.",
 "TAGS: R (blue) = reported by company or filing (S-1, Form D, company PR/blog); P (purple) = press-reported, secondhand (The Information, Bloomberg, WSJ relayed by DCD etc.); D (black) = derived, formula shown; E (orange) = estimated (assumption, or third-party estimate such as Sacra); n/d (grey) = not disclosed. 'inputs P' on a D row = derived from press figures; treat as press quality.",
 "EDGAR CHECK 2026-10-04: Nscale public S-1 filed 2026-09-18 (CIK 0002110365; no S-1/A yet) = only primary financials. Lambda (CIK 0001954966): NO S-1/DRS publicly filed; Form D only (2022, 2024). Crusoe (CIK 0001924674): Form D only (latest 2026-09-17).",
 "UNITS: $M unless stated; per-MW KPIs annualized ($M per MW-yr); $/GPU in $K per GPU-yr; rates % p.a. Floating all-ins ESTIMATED with SOFR 3.87% (NY Fed, 2026-10-01).",
 "MW BASIS: Nscale = IT MW. Crusoe = GROSS MW (IT-equivalent estimated with Microsoft Abilene design ratio 672/900). Lambda = basis not stated.",
]
for i, t in enumerate(README, 1): ws.cell(row=i, column=1, value=t).alignment = wrap
ws.column_dimensions["A"].width = 170

COLS = ["DP id", "Row ID (CRWV template)", "Metric", "Value", "Unit", "As-of date", "Period / basis", "Tag", "Tag name", "Input tags", "Formula (DP ids)", "Source document", "Source date", "URL", "Notes / caveats"]
CELL = {}
for co in ["Nscale", "Lambda", "Crusoe"]:
    for k, d in enumerate(DP[co]): CELL[d["id"]] = (co, f"D{4+k}")
for co in ["Nscale", "Lambda", "Crusoe"]:
    w = wb.create_sheet(co)
    w.cell(row=1, column=1, value=f"{co} - dated data points. Value column is live formulas for D/E rows. Tags: see README.").font = BOLD
    header(w, 3, COLS, [8, 10, 52, 14, 11, 12, 18, 6, 13, 10, 30, 60, 11, 60, 80])
    for k, d in enumerate(DP[co]):
        rr = 4 + k
        doc, sdate, url = S[d["src"]]
        w.cell(row=rr, column=1, value=d["id"]); w.cell(row=rr, column=2, value=d["row"]); w.cell(row=rr, column=3, value=d["metric"]).alignment = wrap
        c = w.cell(row=rr, column=4)
        c.value = ("=" + re.sub(r"\{(\w+)\}", lambda m: CELL[m.group(1)][1], d["f"])) if d["f"] else ("n/d" if d["value"] is None else d["value"])
        c.font = TF[d["tag"]]; c.number_format = "#,##0.00"
        for col, v in [(5, d["unit"]), (6, d["asof"]), (7, d["period"]), (9, TAGNAME[d["tag"]]), (10, d["intags"]), (11, d["f"] or ""), (13, sdate), (14, url)]:
            w.cell(row=rr, column=col, value=v)
        w.cell(row=rr, column=8, value="n/d" if d["tag"] == "N" else d["tag"]).font = TF[d["tag"]]
        w.cell(row=rr, column=12, value=doc).alignment = wrap
        nc = w.cell(row=rr, column=15, value=d["note"]); nc.alignment = wrap
        # Notes such as "= 35,000 / 6 / 350" are human-readable arithmetic, not formulas. openpyxl stores any
        # string starting with "=" as a formula, so force text (otherwise #VALUE!, or a silent number for "= 13 / 5").
        if isinstance(d["note"], str) and d["note"].startswith("="): nc.data_type = "s"
    w.freeze_panes = "D4"

wsS = wb.create_sheet("Summary", 1)
wsS.cell(row=1, column=1, value="Summary - latest dated value per company in the CoreWeave row layout (live links to company tabs). R reported / P press / D derived / E estimated / n/d. CoreWeave Q2'26 for reference (crwv_kpis.md).").font = BOLD
hdr = ["ID", "Item", "Unit"]
for co in ["Nscale", "Lambda", "Crusoe"]: hdr += [f"{co} value", "as-of", "tag", "DP / basis"]
hdr += ["CoreWeave Q2'26 (ref)", "CRWV tag"]
header(wsS, 3, hdr, [7, 42, 11] + [12, 11, 9, 40]*3 + [18, 8])
SUMVAL = {}
r = 4
for row in ROWS:
    if row[1] is None:
        wsS.cell(row=r, column=1, value=row[0]).font = BOLD
        for j in range(1, len(hdr)+1): wsS.cell(row=r, column=j).fill = SECF
        r += 1; continue
    rid, item, unit = row
    wsS.cell(row=r, column=1, value=rid); wsS.cell(row=r, column=2, value=item); wsS.cell(row=r, column=3, value=unit)
    for k, co in enumerate(["Nscale", "Lambda", "Crusoe"]):
        base = 4 + 4*k
        m = MAP[co].get(rid, ND("Not disclosed"))
        if isinstance(m, tuple):
            wsS.cell(row=r, column=base, value="n/d").font = TF["N"]
            wsS.cell(row=r, column=base+2, value="n/d").font = TF["N"]
            wsS.cell(row=r, column=base+3, value=m[1]).alignment = wrap
            SUMVAL[(co, rid)] = (None, None, "n/d", m[1], unit)
        else:
            d = next(x for x in DP[co] if x["id"] == m)
            sh, cell = CELL[m]
            c = wsS.cell(row=r, column=base, value=f"='{sh}'!{cell}"); c.font = TF[d["tag"]]; c.number_format = "#,##0.00"
            wsS.cell(row=r, column=base+1, value=d["asof"])
            wsS.cell(row=r, column=base+2, value=tagtxt(d)).font = TF[d["tag"]]
            wsS.cell(row=r, column=base+3, value=f"{m}: {d['metric']}").alignment = wrap
            SUMVAL[(co, rid)] = (VAL[m], d["asof"], tagtxt(d), d["metric"], d["unit"])
    if rid in CRWV_Q226:
        wsS.cell(row=r, column=16, value=CRWV_Q226[rid][0]); wsS.cell(row=r, column=17, value=CRWV_Q226[rid][1])
    r += 1
wsS.freeze_panes = "D4"

wsD = wb.create_sheet("Debt_Facilities")
wsD.cell(row=1, column=1, value=f"Debt and financing facilities (commitment size unless noted). All-in % on floating facilities = ESTIMATE: SOFR {SOFR}% (2026-10-01) + disclosed spread (+ OID amortized where given).").font = BOLD
header(wsD, 3, ["Company", "Facility", "Date", "Size ($M)", "Pricing (as disclosed)", "All-in % (est. if floating)", "Tag", "Maturity", "Notes", "Source", "URL"], [10, 52, 12, 11, 50, 12, 14, 18, 48, 50, 60])
for j, f in enumerate(FAC, 4):
    co, name, date, size, pricing, allin, tag, mat, note, sk = f
    vals = [co, name, date, size if size is not None else "n/d", pricing, round(allin, 2) if allin is not None else "n/d", tag, mat, note, S[sk][0], S[sk][2]]
    for k, v in enumerate(vals, 1):
        c = wsD.cell(row=j, column=k, value=v); c.alignment = wrap
        if k == 7: c.font = TF.get(tag[0], TF["D"])

wsR = wb.create_sheet("Sources")
header(wsR, 1, ["Key", "Document", "Publication date", "URL", "Class"], [10, 110, 14, 100, 24])
for j, (k, (doc, dte, url)) in enumerate(S.items(), 2):
    cls = "Reported (filing/company)" if k in REPORTED_KEYS else ("Third-party estimate" if "SACRA" in k else ("Data / workspace" if k in ("SOFR", "NBPX", "CRWV") else "Press (secondhand)"))
    for c, v in enumerate([k, doc, dte, url, cls], 1): wsR.cell(row=j, column=c, value=v).alignment = wrap

DEFS = [
 ("D1", "Revenue per MW", "Annualized revenue / MW ($M per MW-yr); denominators active MW (period-end), average MW, contracted MW.", "Periods rarely match: Nscale H1'26 x2 over Aug-26 active-site MW; Lambda Q2'25 press revenue over Q2'25 press MW; Crusoe FY26E press revenue over Sep-26 gross MW (Estimated; mixes landlord and cloud revenue). K02 n/d for all."),
 ("D2", "Revenue per GPU", "Revenue x 4000 / avg GPUs ($K/GPU-yr); $/GPU-hr = revenue / (GPUs x hours).", "Point GPU counts used as averages -> Estimated. PX13 = contract value / (GPUs x term x 8,760)."),
 ("D3", "Implied utilization", "Implied $/GPU-hr / fleet-weighted on-demand list (mixes discounting and utilization).", "Lambda only (H100 SXM $2.99 single-SKU proxy). Nscale: no public list. Crusoe: no GPU count."),
 ("D4", "Capex per MW", "Cash capex / new active MW; stock checks Net PP&E / active MW and in-service gross PP&E / active MW.", "No MW history -> flow versions n/d. Nscale stock ratios from S-1. Crusoe K14 row = Abilene JV project cost ($15B / 1.2 GW, shell + power, ex-GPU) - proxy, not like-for-like."),
 ("D5", "D&A useful lives", "Useful-life policy; depreciation yield = annualized depreciation / avg in-service gross PP&E; implied life = 1/yield.", "Nscale 5-6 yrs (midpoint 5.5); yield uses PP&E depreciation only. Lambda/Crusoe n/d."),
 ("D6", "Backlog / RPO", "ASC 606 RPO; company backlog shown separately.", "TCV is NOT RPO: Nscale RPO $56.4B vs TCV $103.4B; Crusoe TCV $140B+ has no term or schedule."),
 ("D7", "Customer concentration", "Customers >=10% of revenue.", "Nscale largest only; Lambda press only; Crusoe n/d."),
 ("D8", "Funding cost", "Principal-weighted effective rate from the debt table.", "Nscale S-1 effective rates; Lambda/Crusoe no debt table -> PX09 marginal facility cost."),
 ("Gross vs IT MW", "MW basis", "IT MW = critical IT load; gross = utility/total power incl. cooling and overhead.", "Crusoe reports gross; converted at 672/900 = 0.747 (one campus design) -> Estimated."),
 ("Run-rate / press", "Revenue basis", "Run-rate = latest period annualized; forward = company expectation.", "Press-reported (P) and third-party (E) figures kept separate from filings."),
 ("PX01", "Revenue run-rate / forward", "Annualized latest period, ARR, or full-year expectation.", ""),
 ("PX02", "Announced / pipeline MW", "Capacity announced or in pipeline beyond contracted.", "Not additive across sources."),
 ("PX03", "MW basis", "IT vs gross.", ""), ("PX04", "GPUs contracted", "Active + contracted GPUs.", "Nscale only."),
 ("PX05", "Contract life", "Weighted average contract life.", "Nscale only (5.7 yrs; 23% prepaid)."),
 ("PX06-PX08", "Funding raised", "Equity (Form D amount sold where available), valuation, debt commitments.", "Form D amounts can lag headline round sizes."),
 ("PX09", "Marginal secured debt cost", "Coupon or spread on the latest secured GPU/project facility; all-in = SOFR + spread (+ OID amortized).", "CRWV reference: K28 8.4% (Q2'26)."),
 ("PX10-PX11", "Contract value per MW", "ACV/contracted MW = TCV / term / MW; TCV/MW.", "Nebius ~$20-25M ACV/MW on new deals (crwv_kpis.md)."),
 ("PX12-PX13", "Anchor contracts", "Largest disclosed contracts; implied $/GPU-hr.", ""), ("PX14", "Valuation / revenue", "Post-money / forward revenue.", ""),
]
wsF = wb.create_sheet("Definitions")
header(wsF, 1, ["Def / row", "Group", "Definition (as CoreWeave workbook)", "Private-peer application"], [14, 24, 80, 90])
for j, d in enumerate(DEFS, 2):
    for k, v in enumerate(d, 1): wsF.cell(row=j, column=k, value=v).alignment = wrap

NOTES = [
 # --- Structurally unknowable until an S-1 / audited disclosure ---
 ("All", "Structural", "No quarterly history: every KPI is a single dated point. K02 and K10-K13 stay n/d until there is an MW time series.", "Not fixable from press."),
 ("All", "Structural", "Capex per MW (K10-K13) n/d for all three; only stock proxies (Nscale) and project cost (Crusoe Abilene).", "Needs audited MW + capex."),
 ("Lambda", "Structural", "No S-1/DRS on EDGAR (Form D only). Audited revenue, MW total, GPU count since Feb-2025, backlog/RPO, debt outstanding, D&A policy and concentration are unknowable from public sources.", "Largest structural gap."),
 ("Crusoe", "Structural", "No financial statements. Revenue is press-only. MW is gross and includes JV/landlord capacity (Abilene GPUs are Oracle/OpenAI's — see C37 / peers/sites/). TCV has no term; no GPU count, concentration, debt balance or useful lives.", "Rev/MW not comparable to cloud peers."),
 ("Nscale", "Structural", "S-1 defines but does not quantify company-wide active MW. 81 MW = sum of contracted IT MW at the 5 active sites (upper bound). 25k GPUs / 81 MW ≈ 309 GPUs/MW vs ~491 for GB300 → live MW likely lower.", "K01 is a lower bound until S-1/A."),
 ("Nscale", "Structural", "No public list price (console-only) → K08/K09 n/d.", ""),
 # --- Caveats (known, but handle carefully) ---
 ("Nscale", "Caveat", "Revenue base tiny vs book: H1'26 $140.6M vs RPO $56.4B / TCV $103.4B (active TCV $2.6B = 2.5%). ~86% of TCV is Microsoft (up to $43.8B) + Anthropic (up to $44.6B; financing not committed).", "Treat ACV/MW, TCV/MW as contracted economics."),
 ("Nscale", "Caveat", "6/30/26 weighted debt cost 18.7% reflects small legacy facilities; Aug-26 GPU facilities at SOFR+2.375% (~6.2% E), revolver SOFR+1.75%.", "Use PX09 for forward cost."),
 ("Nscale", "Caveat", "Interest expense, net 95.1 (H1'26) includes financing-component interest on prepayments; interest on debt only 12.0.", ""),
 ("Nscale", "Caveat", "Monarch WV (N60-N63): 460 MW contracted / 0 evidenced-active; Grade Yellow for H1'28. Off-grid gas microgrid; air permit R14-0042 still DRAFT; Anthropic financing uncommitted. Do not treat $16.2M ACV/MW (N62) as live capacity. Cite peers/sites/ + this pack for Modeling v0.9a.", "Contract price yes; powered MW no."),
 ("Lambda", "Caveat", "Bloomberg ~$1B Aug-28 private debt may be the Oct-1 $1.008B IG DDTL; not summed into PX08.", ""),
 ("Lambda", "Caveat", "Kansas City (Project Blitz) site status: peers/sites/ — Active L/B/H 0/12/24, Grade Yellow, no go-live confirmation as of 2026-10-04.", ""),
 ("Crusoe", "Caveat", "Abilene evidenced-active IT MW is in peers/sites/ (and C37 here). Count once under Crusoe; Oracle's own key site is Port Washington WI (0 active).", "Avoid double-counting Stargate."),
 ("All", "Method", "D rows from press inputs show 'inputs P' and should be read as press quality.", ""),
 # --- Watch next (could become dated R/P/D points) ---
 ("Nscale", "Watch", "S-1/A with price range; Q3 stub that quantifies active MW; Anthropic financing close; WVDEP final air permit R14-0042; Monarch vertical construction / first roofs on Sentinel-2.", "EDGAR + peers/sites refresh."),
 ("Nscale", "Watch", "IPO pricing talk (~Oct-15, ~$3B raise, ~$35B val) remains P/unconfirmed; no price range in S-1. EDGAR re-check 2026-10-05: no S-1/A yet.", "NS-EDGAR2"),
 ("Lambda", "Watch", "Public DRS/S-1 (IPO now targeted 2027 per WSJ 10/6); close of the up-to-$4B round at $14.5B pre-money (WSJ talks, not closed; supersedes Bloomberg $3B/$12B+); KC Blitz energization confirmation; debt outstanding balance.", "LM-WSJ4B"),
 ("Crusoe", "Watch", "Final pricing on Iceland GPU DDTL (Octus talk S+475-500); any financial statements or Form D updates; Abilene B7-8 delivery (618 of ~824 MW delivered per Oracle Sep 10; peers/sites YE26 618/721/824).", ""),
]
wsN = wb.create_sheet("Notes_Gaps")
header(wsN, 1, ["#", "Company", "Type", "Note / gap", "Impact / fix"], [4, 10, 10, 100, 60])
for j, n in enumerate(NOTES, 2):
    wsN.cell(row=j, column=1, value=j-1)
    for k, v in enumerate(n, 2): wsN.cell(row=j, column=k, value=v).alignment = wrap

wb.save(OUT)
json.dump({"vals": VAL, "summary": {f"{a}|{b}": v for (a, b), v in SUMVAL.items()}}, open(VALS_OUT, "w"), indent=1, default=str)
print("saved", OUT, {co: len(DP[co]) for co in DP}, "facilities", len(FAC))

# ======================= Markdown write-up =======================
MD = "/workspace/nebius/peers/private/private_peers.md"
def fmt(v):
    if v is None: return "n/d"
    if isinstance(v, str): return v
    a = abs(v)
    return f"{v:,.0f}" if a >= 100 else (f"{v:,.1f}" if a >= 10 else f"{v:,.2f}")
KEYROWS = ["PX01","I01","I09","I10","I11","PX04","I14","I15","PX05","K01","K03","PX10","PX11","K04","K05","K06","K07","PX13","K08","K09","K14","K15","K16","K17","K19","K23","I17","K26","K28","PX09","K38","PX06","PX07","PX08","PX14"]
NAMES = {r[0]: (r[1], r[2]) for r in ROWS if r[1] is not None}
lines = ["# Private peer KPI comps: Nscale, Lambda, Crusoe", "",
 "Built Sun 2026-10-04 (ET); refreshed Mon Oct 5 2026 (ET) by `build_private_peers.py`. Workbook: `private_peers.xlsx`, with tabs README, Summary, Nscale, Lambda, Crusoe, Debt_Facilities, Sources, Definitions and Notes_Gaps. Row IDs and definitions match `/workspace/nebius/peers/coreweave/crwv_kpis.xlsx` (I01-I27, K01-K40, D1-D8). Rows that exist only for private peers are PX01-PX14.", "",
 "**Tags:** R = reported by the company or a filing; P = press-reported (secondhand); D = derived (formula in the workbook); E = estimated (rests on an assumption or a third-party estimate such as Sacra); n/d = not disclosed. A D row built on press inputs shows `(inputs P)`. Every figure has its source document, date and URL on the company tabs.", "",
 "## Source base (EDGAR check 2026-10-05; no new filings since Oct 4)",
 "- **Nscale:** public **S-1 filed 2026-09-18** (CIK 0002110365; DRS 2026-02-17 plus 3 DRS/A; **no S-1/A as of 2026-10-05**, no price range). Only audited source among the three. Monarch WV site status folded from `peers/sites/` (N60–N63).",
 "- **Lambda:** **no S-1, F-1 or public DRS** on EDGAR (re-check 2026-10-05). Lambda, Inc. (CIK 0001954966) has only Form D filings (2022, 2024). Financials = press (The Information 2025, Bloomberg 2026) and Sacra estimates.",
 "- **Crusoe:** Form D only (latest 2026-09-17: $3.11B of $3.90B Series F sold; no new Form D as of 2026-10-05). Metrics from company PRs (gross MW, TCV) and press revenue. Abilene site status: `peers/sites/` / C37.", "",
 "## Latest values (Summary tab)", "",
 "| ID | KPI | Unit | Nscale | Lambda | Crusoe | CRWV Q2'26 |", "|---|---|---|---|---|---|---|"]
for rid in KEYROWS:
    nm, un = NAMES[rid]
    cells = []
    for co in ["Nscale", "Lambda", "Crusoe"]:
        v, asof, tag, metric, unit = SUMVAL[(co, rid)]
        cells.append("n/d" if tag == "n/d" else f"{fmt(v)} ({tag}, {asof})")
    cw = CRWV_Q226.get(rid, ("", ""))[0]
    lines.append(f"| {rid} | {nm} | {un} | " + " | ".join(cells) + f" | {cw} |")
lines += ["", "The reference column comes from crwv_kpis.md. PX10 for CRWV is the Nebius ~$20-25M ACV/MW benchmark on new deals.", "",
 "## Findings",
 "**Nscale (S-1).**",
 "- H1'26 revenue was $140.6M (R), a $281M run-rate.",
 "- The contract book dwarfs revenue: RPO $56.4B (R, 6/30/26) and TCV $103.4B (R, 8/31/26). Only $2.6B of TCV is active. Roughly 86% of TCV is Microsoft (up to $43.8B) plus Anthropic (up to $44.6B; financing not yet committed).",
 "- 1,370 IT MW is active + contracted (R). The S-1 does not quantify active MW; the 5 active sites sum to 81 MW (D, upper bound).",
 "- Revenue per active MW is about 3.5 (D). On a contracted basis, ACV is about $13.2M/MW-yr (E: TCV / 5.7-yr life / 1,370 MW), below the ~$20-25M Nebius benchmark.",
 "- Contract-implied pricing is about $4.49 per GPU-hr (E) across 461k GPUs. Realized H1'26 revenue was $1.29 per GPU-hr (E), reflecting the ramp.",
 "- Technology equipment useful life is 5-6 years (R). The PP&E depreciation yield is 15.7%, an implied life of 6.4 years (D), in line with CRWV's 15.5%.",
 "- Largest customer was 52% of H1'26 revenue (R).",
 "- Funding cost: the reported weighted effective rate on 6/30/26 debt was 18.7% (D), but that debt is only $137M of legacy Macquarie and GPU-facility borrowings. The August 2026 GPU facilities price at SOFR+2.375% (about 6.2%, E), with $6.7B of new commitments after June (D). $3.36B of pre-IPO converts and a $14.6B Series C valuation are both R.",
 "- **Monarch WV footnote (for Modeling v0.9a):** contracted **460 IT MW** (N60, R) / evidenced-active **0** (N61, R from peers/sites). Anthropic up to $44.6B (N35); implied ACV **~$16.2M/IT MW-yr** if 6-yr press term (N62, E). Grade **Yellow** for 460 by H1'28 / Yellow-Red for 1.37 GW (N63). Off-grid gas microgrid; air permit still draft; financing uncommitted. **Cite this pack + peers/sites/site_verification.md §7** — contract price stands; do not treat as live capacity.",
 "",
 "**Lambda (press only).**",
 "- Q2'25 revenue was ~$140M, of which ~$114M cloud (P). Sacra estimates Dec-2025 ARR at ~$760M (E). Bloomberg reports 2026 revenue expected above $1.5B (P).",
 "- Q2'25 capacity was 47 MW (P), giving about $9.7M cloud revenue per MW-yr (D on P inputs).",
 "- Revenue per GPU was about $2.09 per GPU-hr (E), against a $2.99 H100 list price, which implies ~70% utilization or price realization (E).",
 "- Contracts: NVIDIA's leaseback implies $3.71 per GPU-hr (P). The press-reported Anthropic deal ($35B over 6 years, ~350 MW) works out to about $16.7M ACV per MW-yr (D on P inputs).",
 "- Backlog (company metric, per a Lambda investor letter reported by WSJ 2026-10-06) rose from $15B (Jun-26) to $50B (Sep-26) (P, L38/L39); ~$35B of the increase is Anthropic. Lambda is raising up to $4B at $14.5B pre-money (Blackstone + Coatue; talks, P, L37) and now targets a 2027 IPO.",
 "- 2026 debt has moved to investment-grade-style structures: a $1.008B DDTL at 6.78% fixed (A(low)/Baa1, R), a $926M TLB at SOFR+300 (R), and a $1B credit facility (R). Total commitments are $2.93B (D). Outstanding debt is n/d.",
 "",
 "**Crusoe (company PR plus press).**",
 "- Series F: $3.9B at $30.9B post-money (R). TCV is $140B+, gross contracted capacity is 6 GW+, and ~1 GW gross is delivered (R, 2026-09-17).",
 "- 2026 revenue is 'on track for $2B' (P, Newcomer), which implies about $2.0M per gross MW-yr (E). That figure mixes landlord and cloud revenue and includes JV capacity, so it is not comparable to cloud peers.",
 "- TCV per gross contracted MW is about $23M (D), but no term is disclosed. Valuation is about 15x 2026E revenue (E).",
 "- The Abilene JV costs about $12.5M per gross MW for shell and power, excluding GPUs (D).",
 "- Debt pricing: the Iceland GPU DDTL price talk is SOFR+475-500 (P), about 9.1% all-in (E). An AMD-backstopped loan is reported at ~6% (P).",
 "", "## Gaps (structural vs watch)",
 "", "### Structurally unknowable (until S-1 / audited disclosure)",
 "- **Lambda:** no S-1 → no audited revenue, MW total, GPU count since Feb-2025, ASC 606 RPO (press backlog only: $50B Sep-26), debt outstanding, useful lives or concentration. Latest hard revenue = Q2'25 press.",
 "- **Crusoe:** no financials; revenue press-only; MW gross incl. JV/landlord (Abilene GPUs = Oracle/OpenAI). No TCV term, GPU count, concentration, debt balance, useful lives.",
 "- **Nscale:** company-wide active MW not quantified (81 MW = site-sum upper bound). No list price → K08/K09 n/d. Half-year only.",
 "- **All three:** K02 and K10–K13 n/d (no MW history). Every KPI is a single dated point.",
 "", "### Known caveats (handle carefully)",
 "- Nscale Anthropic $44.6B / Monarch 460 MW: financing not committed; campus civil-only (Grade Yellow). $16.2M ACV/MW is a contract print, not live capacity (N60–N63; peers/sites §7).",
 "- Crusoe Rev/MW mixes landlord + cloud and includes JV capacity — not comparable to pure cloud peers. Abilene counted under Crusoe (C37 / peers/sites).",
 "- Lambda Aug-28 ~$1B private debt may overlap the Oct-1 IG DDTL — not double-counted in PX08.",
 "", "### Watch next",
 "- Nscale S-1/A (price range; Q3 active MW); Anthropic financing close; WVDEP final R14-0042; Monarch first roofs (peers/sites). IPO talk ~Oct-15 / ~$35B remains P/unconfirmed (EDGAR 2026-10-05: no S-1/A).",
 "- Lambda public DRS/S-1 (IPO now 2027 per WSJ); up-to-$4B / $14.5B pre-money round close (WSJ 10/6 talks); KC Blitz go-live.",
 "- Crusoe Iceland DDTL final pricing; Abilene B5–8 energization (peers/sites YE26).",
]

open(MD, "w").write("\n".join(lines) + "\n")
print("md written")
