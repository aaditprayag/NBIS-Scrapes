import csv, json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
S = dict(
 q2_26="https://assets.nebius.com/assets/4462517b-ce83-41f2-96ed-f2ac1bc06a05/SHLQ226%20%281%29.pdf",
 q2_26sec="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm",
 q2_26call="https://stockanalysis.com/stocks/nbis/transcripts/655490-q2-2026/",
 q1_26="https://assets.nebius.com/assets/aa1bc2e6-df83-40cd-a6a2-95e7cda3d16c/Nebius%20SHL_Q1%202026.pdf",
 q1_26call="https://www.fool.com/earnings/call-transcripts/2026/05/13/nebius-nbis-q1-2026-earnings-transcript/",
 q4_25="https://www.sec.gov/Archives/edgar/data/1513845/000110465926013946/tm266173d1_ex99-2.htm",
 q4_25call="https://stockanalysis.com/stocks/nbis/transcripts/402284-q4-2025/",
 q3_25="https://www.sec.gov/Archives/edgar/data/1513845/000110465925109806/tm2530882d1_ex99-2.htm",
 q3_25call="https://stockanalysis.com/stocks/nbis/transcripts/367767-q3-2025/",
 q2_25="https://www.sec.gov/Archives/edgar/data/1513845/000110465925075028/tm2522866d1_ex99-2.htm",
 q2_25call="https://equibles.com/stocks/nbis/calls/2025-q2",
 q1_25="https://www.sec.gov/Archives/edgar/data/1513845/000110465925050975/tm2515580d1_ex99-2.htm",
 q1_25call="https://www.fool.com/earnings/call-transcripts/2025/05/20/nebius-nbis-q1-2025-earnings-call-transcript/",
 q4_24="https://www.sec.gov/Archives/edgar/data/1513845/000110465925015480/tm257185d1_ex99-1.htm",
 q4_24call="https://equibles.com/stocks/nbis/calls/2024-q4",
 q3_24="https://nebius.com/newsroom/nebius-group-n-v-announces-third-quarter-2024-financial-results",
 q3_24call="https://exa.ai/library/markets/stocks/NBIS/earnings/FY2024/Q3/transcript",
 f20="https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm",
 mant75="https://nebius.com/newsroom/nebius-to-triple-capacity-at-finland-data-center-to-75-mw",
 dcd_kc="https://www.datacenterdynamics.com/en/news/nebius-to-deploy-5mw-nvidia-h200-cluster-at-patmos-data-center-in-kansas-city-missouri/",
 lapp="https://nebius.com/newsroom/nebius-to-construct-310-mw-ai-factory-in-finland",
 lapp_pdf="https://assets.nebius.com/assets/6590d0a0-0fc4-4396-bc3f-f52a9ccb770e/20260331%20Nebius%20to%20construct%20310%20MW%20AI%20factory%20in%20Finland.pdf",
 reut_lapp="https://www.reuters.com/technology/nebius-furthers-european-expansion-with-10-billion-ai-data-centre-finland-2026-03-31/",
 polarnode="https://www.polarnode.fi/historic-investment-in-lappeenranta-finland-nebius-to-operate-polarnodes-310-mw-data-center-campus/",
 caverion_l="https://www.caverion.com/newsroom/releases/2026/caverion-to-deliver-eur-15-million-substation-for-nebius-data-center-in-finland-creating-jobs-for-more-than-80-professionals/",
 caverion_m="https://www.assemblincaverion.com/news-and-media/news/2026/caverion-finland-strengthens-its-position-in-the-data-centre-market-and-will-deliver-a-eur-15-million-substation-for-ai-cloud-company-nebius/",
 dcd_ee="https://www.datacenterdynamics.com/en/news/nebius-expands-european-presence-announces-deployment-in-estonia-and-second-data-center-in-m%C3%A4nts%C3%A4l%C3%A4-finland/",
 err_ee="https://news.err.ee/1610109073/baltics-first-ai-factory-to-open-in-estonia",
 paris="https://nebius.com/blog/posts/nebius-launches-gpu-cluster-in-paris-france",
 reut_beth="https://www.reuters.com/technology/amsterdam-based-ai-firm-nebius-build-240mw-data-centre-near-lille-france-2026-02-12/",
 dcmag_beth="https://dcmag.fr/nebius-annonce-un-mega-datacenter-de-240-mw-a-bethune-pas-de-calais-chez-azur-datacenter/",
 gnt_beth="https://www.generation-nt.com/actualites/nebius-datacenter-bethune-ia-france-bridgestone-2070889",
 nj_ice="https://nebius.com/blog/posts/300-mw-new-jersey-and-iceland-regions",
 nj_bw="https://www.businesswire.com/news/home/20250305405030/en/Nebius-accelerates-US-expansion-adding-up-to-300-MW-capacity-at-new-data-center-in-New-Jersey",
 nj_page="https://nebius.com/vinelandnj",
 msft="https://www.sec.gov/Archives/edgar/data/1513845/000110465925088312/tm2525580d1_6k.htm",
 meta1="https://www.sec.gov/Archives/edgar/data/1513845/000110465925109803/tm2530882d2_6k.htm",
 meta2="https://www.sec.gov/Archives/edgar/data/1513845/000110465926027886/tm268879d1_ex99-1.htm",
 uk_nov25="https://www.datacenterdynamics.com/en/news/nebius-launches-nvidia-gpu-cluster-in-london-uk/",
 uk_ark_jun25="https://www.datacenterdynamics.com/en/news/nebius-to-host-uk-nvidia-blackwell-ultra-gpu-cluster-in-ark-data-centres-facility-in-surrey/",
 uk_ark_jun26="https://www.datacenterdynamics.com/en/news/ark-dc-to-add-new-building-to-longcross-data-center-campus-outside-london-uk/",
 uk_ark_dcn="https://datacentrenews.uk/story/nebius-fills-ark-s-longcross-data-centre-to-capacity",
 uk_pr="https://nebius.com/newsroom/nebius-expands-in-uk-with-more-nvidia-powered-infrastructure-more-customers-and-more-cloud-capabilities-for-agentic-and-enterprise-ai",
 uk_kao="https://www.datacenterdynamics.com/en/news/nebius-signs-22mw-capacity-agreement-with-kao-data-in-the-uk/",
 il_pr="https://nebius.com/newsroom/nebius-brings-nvidia-blackwell-to-israel-with-one-of-the-country-s-first-ai-infrastructure-deployments",
 il_80="https://www.datacenterdynamics.com/en/news/nebius-signs-80mw-data-center-lease-with-mega-or-in-israel/",
 il_ctech="https://www.calcalistech.com/ctechnews/article/p4ds17rzm",
 es_dcd="https://www.datacenterdynamics.com/en/news/nebius-signs-18mw-lease-with-merlin-properties-at-spain-data-center-report/",
 es_merlin="https://www.merlinproperties.com/dia-a-dia/firmamos-un-contrato-de-18-mw-con-nebius-para-nuestro-date-center-madrid-getafe/",
 es_li="https://www.linkedin.com/feed/update/urn:li:activity:7505906816590073857",
 indep="https://nebius.com/independencemo",
 dcd_indep="https://www.datacenterdynamics.com/en/news/nebius-breaks-ground-on-400-acre-data-center-campus-in-independence-missouri/",
 dcd_indep800="https://www.datacenterdynamics.com/en/news/nebius-plans-800mw-data-center-campus-in-kansas-city-missouri/",
 bham="https://nebius.com/birminghamal",
 dcd_bham="https://www.datacenterdynamics.com/en/news/nebius-targets-data-center-development-in-birmingham-alabama/",
 al_com="https://www.al.com/news/2026/02/multibillion-dollar-birmingham-ai-factory-taking-shape-not-coming-to-town-to-build-a-big-grey-box.html",
 dcd_pa="https://www.datacenterdynamics.com/en/news/nebius-details-plans-for-12gw-data-center-campus-in-pennsylvania/",
 pa_pr="https://assets.nebius.com/assets/6fc0ea6c-0884-4a1f-bed8-1f797eb9628f/Financial%20results_Q1%202026.pdf",
 pa_rh="https://www.republicanherald.com/2026/09/01/nebius-gets-green-light-to-expand-its-data-center-plan-in-butler-twp/",
 docs="https://docs.nebius.com/overview/regions",
 aib="https://www.globenewswire.com/news-release/2026/09/30/3371855/0/en/aib-data-centers-signs-contract-with-nebius-for-ai-data-center-capacity.html",
)
cols=["#","Site","City/Region","Country","Lat","Lon","Coord basis","Type","Partner/Landlord","Announced MW – initial phase","Announced MW – full build-out","IT vs gross","Timeline (announced; revisions)","GPU generation(s)","Anchor customer","Status (as of Oct 4 2026)","Sources (date)"]
R=[]
def add(*a): R.append(list(a))
add(1,"Mäntsälä DC #1","Mäntsälä, Uusimaa","Finland",60.633,25.317,"approx, town-level","Owned greenfield (self-built)","— (Nebius-owned)","25 MW pre-expansion (DCD, Nov-2024)","75 MW after tripling (Nebius PR Oct 8 2024)","not disclosed","Expansion announced Oct 8 2024; target 'end-2025 or early 2026' (Q3'24 PR); Q1'25 call: phase 1 late Q3'25, phase 2 ~YE25; Mar 31 2026 PR: expansion to 75 MW 'completed earlier this year'","H100 (installed), H200 (from Nov 2024), GB300 NVL72 ('Europe's first operational', Mar 2026); docs eu-north1: H100/H200/L40S","not disclosed","Operational, 75 MW built",f"{S['mant75']} (2024-10-08); {S['q3_24']} (2024-11); {S['dcd_kc']} (2024-11-19); {S['lapp_pdf']} (2026-03-31); {S['docs']} (accessed 2026-10-04)")
add(2,"Mäntsälä DC #2","Mäntsälä, Uusimaa","Finland",60.633,25.317,"approx, town-level","Owned (implied: Nebius 'building a second data center'; not explicitly labelled)","Caverion (EUR15M 70 MW substation contractor)","not disclosed","up to 70 MW","not disclosed","Announced Aug 13 2026 (LinkedIn, with Q2'26 results; Q2 letter: 'additional sites in … Finland'); capacity 'from 2027'; substation completion autumn 2027. Finland total 455 MW across 3 sites","not disclosed","not disclosed","Planned / pre-construction",f"{S['dcd_ee']} (2026-08-13); {S['caverion_m']} (2026); {S['q2_26']} (2026-08-12)")
add(3,"Lappeenranta AI factory (Pajarila)","Lappeenranta, South Karelia","Finland",61.06,28.19,"approx, city-level","Owned per Nebius ('new owned location', Q1'26 letter) — but developer Polarnode describes Nebius as 'operating Polarnode's campus' (see conflicts)","Polarnode (developer/builder); Caverion (substation)","Substation phase 1: 70 MW (Caverion)","up to 310 MW","not disclosed","Announced Mar 31 2026 (Nebius PR; Reuters: '10th site', >$10B value); first capacity to customers in 2027; substation construction from Jul 2026, completion spring 2027 (Caverion)","'latest Blackwell and Rubin generations' (company-level statement in PR; site-specific not disclosed)","not disclosed","Under construction (pre-construction by Polarnode)",f"{S['lapp_pdf']} (2026-03-31); {S['reut_lapp']} (2026-03-31); {S['polarnode']} (2026-03); {S['caverion_l']} (2026); {S['q1_26']} (2026-05-13)")
add(4,"Paris – Equinix PA10","Saint-Denis, Île-de-France","France",48.92,2.36,"approx, Saint-Denis district","Colocation","Equinix (PA10)","not disclosed (third-party DC Atlas: ~5 MW — unverified)","not disclosed","not disclosed","Agreement Jul 2024 (20-F); announced Sep 25 2024; first client workloads Nov 2024; in operation Nov 2024 (Q4'24 PR)","H200 (docs eu-west1: H200)","not disclosed","Operational",f"{S['paris']} (2024-09-25); {S['q4_24']} (2025-02-20); {S['f20']} (2026-04-30); {S['docs']}")
add(5,"Béthune AI factory (ex-Bridgestone plant)","Béthune, Pas-de-Calais (near Lille)","France",50.53,2.64,"approx, town-level","Colocation per 20-F ('two new co-location sites in France', Feb-2026); Nebius PR says 'building an AI factory'","Azur Datacenter (finances/builds shell; Nebius IT) per DCmag/Reuters","First capacity 'late summer 2026'; ~120 MW (half) by YE2026","240 MW","not disclosed","Announced Feb 12 2026 (with Q4'25 results; Reuters/Les Echos); first capacity late summer 2026; ~half (120 MW) by end-2026; full 240 MW in 2027 (DCmag/GNT: end-2027)","not disclosed","not disclosed","Under construction / phase 1 possibly live (unconfirmed)",f"{S['reut_beth']} (2026-02-12); {S['dcmag_beth']} (2026-02-13); {S['gnt_beth']} (2026-02-17); {S['f20']} (2026-04-30); {S['lapp_pdf']} (2026-03-31)")
add(6,"France – 2nd new colocation site (unnamed)","not disclosed","France","","","not disclosed","Colocation (20-F)","not disclosed","not disclosed","not disclosed","not disclosed","Secured Feb 2026 (20-F). Possibly the private region eu-west2 (docs: France, B300) — INFERENCE","eu-west2 lists B300 (if this is the site — inference)","not disclosed","not disclosed",f"{S['f20']} (2026-04-30); {S['docs']}")
add(7,"Keflavik – Verne","Keflavik (Ásbrú campus), Reykjanes","Iceland",63.97,-22.58,"approx, Verne campus area","Colocation","Verne","10 MW cluster","10 MW (no expansion disclosed)","not disclosed","Agreement Dec 2024 (20-F); announced Mar 5 2025; live Mar 2025 (Q1'25 letter)","H200 (Q4'24 call; docs eu-north2 private region: H200)","not disclosed","Operational (private region)",f"{S['nj_ice']} (2025-03-05); {S['q1_25']} (2025-05-20); {S['q4_24call']} (2025-02-20); {S['docs']}")
add(8,"London – Ark Longcross Park (LP01; LP02 planned)","Chertsey, Surrey","UK",51.38,-0.59,"approx, Longcross Park","Colocation","Ark Data Centres","16 MW; phase 1 4,000 HGX B300 (Nov 2025)","Phase 2 +3,000 B300 (7,000 total). Jun 2026: Nebius leases entire existing facility; Ark building new 36 MW LP02 (Nebius share not disclosed)","not disclosed","Announced Jun 2025 (Ark named host Jun 25 2025); target Q4'25; launched Nov 6 2025. Jun 2026: Nebius fills LP01 (16 MW live + 3 MW in delivery); Ark £807M expansion","B300 / Blackwell Ultra (docs uk-south1: B300)","not disclosed","Operational + expanding",f"{S['uk_ark_jun25']} (2025-06-25); {S['uk_nov25']} (2025-11-06); {S['uk_ark_jun26']} (2026-06-09); {S['uk_ark_dcn']} (2026-06); {S['docs']}")
add(9,"Harlow – Kao Data (KLON-02/03)","Harlow, Essex","UK",51.77,0.10,"approx, Kao Harlow campus","Colocation (10-year)","Kao Data","not disclosed","22 MW","not disclosed","Announced Jun 8 2026 as part of £1.7bn UK plan; three new UK deployments total 65 MW fully ramped in 2027","'latest generations' of NVIDIA (not specified); docs uk-south2: RTX PRO 6000 (site mapping is inference)","not disclosed","Hardware install underway (third-party); capacity ramp to 2027",f"{S['uk_kao']} (2026-06-08); {S['uk_pr']} (2026-06-08)")
add(10,"UK – third new deployment (unnamed)","not disclosed","UK","","","not disclosed","not disclosed (likely colocation — inference)","not disclosed","not disclosed","Residual of 65 MW UK total (65 − 22 Harlow − Longcross expansion share) — not disclosed","not disclosed","Jun 8 2026 PR: 3 new deployments (4 sites incl. existing) at ~£1.7bn; Q2'26 letter: 'additional sites in the UK'","not disclosed","not disclosed","not disclosed",f"{S['uk_pr']} (2026-06-08); {S['q2_26']} (2026-08-12)")
add(11,"Modi'in – Mega Or","Modi'in","Israel",31.90,35.01,"approx, city-level","Colocation","Mega Or / Mega DC","8 MW; 4,000 HGX B200","Mega Or said Nebius may double to 8,000 GPUs subject to demand (not Nebius-confirmed)","not disclosed","Lease signed May 2025 (Mega Or); launched Oct 20 2025","B200 (docs me-west1: B200)","25% of GPUs offered to Israel Innovation Authority (national supercomputer)","Operational",f"{S['il_pr']} (2025-10-20); {S['il_80']} (2026-01); {S['f20']}")
add(12,"Masmiyya – Mega DC","Masmiyya junction area","Israel",31.74,34.78,"approx","Colocation (5-yr + extension option)","Mega Or / Mega DC","22 MW","22 MW (site could reach 64 MW; Nebius share n/d)","not disclosed","Signed ~Jan 2026 (Mega Or TASE disclosure); delivery Q3 2026","not disclosed","not disclosed","Delivery scheduled Q3 2026",f"{S['il_80']} (2026-01); {S['il_ctech']} (2026-01); {S['f20']}")
add(13,"Beit Shemesh – Mega DC","Beit Shemesh","Israel",31.75,34.99,"approx, city-level","Colocation (5-yr + extension option)","Mega Or / Mega DC","Staged from Q3 2026","58 MW (site could reach ~222 MW; Nebius share n/d)","not disclosed","Signed ~Jan 2026; staged delivery Q3 2026 – Q1 2027","not disclosed","not disclosed","Under construction",f"{S['il_80']} (2026-01); {S['il_ctech']} (2026-01)")
add(14,"Madrid-Getafe – Merlin Edged","Getafe, Madrid","Spain",40.30,-3.73,"approx, city-level","Colocation","Merlin Properties / Edged","Phase 1 live Sep 2026 (MW n/d); phases 2-3 'coming weeks'","18 MW","not disclosed","Q1'26 letter (May 13 2026): 'new colocation site in Spain'; 18 MW reported Jul 3 2026 (El Mundo/DCD; Merlin confirmed); live mid-Sep 2026 (employee LinkedIn post)","RTX PRO 6000 (phase 1) and B300 planned (LinkedIn); docs eu-south1: RTX PRO 6000","not disclosed","Operational (phase 1)",f"{S['q1_26']} (2026-05-13); {S['es_dcd']} (2026-07-03); {S['es_merlin']} (2026); {S['es_li']} (2026-09-16); {S['docs']}")
add(15,"Tallinn – Greenergy Hüüru","Hüüru, Saue Parish (W of Tallinn)","Estonia",59.37,24.58,"approx","Colocation","Greenergy Data Centers (MCF Group)","22 MW first phase later in 2026","22 MW (facility total 31.5 MW)","not disclosed","Signed Q2 2026; announced Aug 12-13 2026 with Q2 results; first phase operational later in 2026","not disclosed","not disclosed","Under construction",f"{S['dcd_ee']} (2026-08-13); {S['err_ee']} (2026-08); {S['q2_26']} (2026-08-12)")
add(16,"Kansas City – Patmos","Kansas City, Missouri (downtown, former KC Star press building)","USA",39.093,-94.580,"approx, downtown KC","Colocation (20-F); Nov-24 PR called it 'first U.S. cluster'","Patmos","5 MW initial phase (thousands of Hopper GPUs)","Expandable to 40 MW (~35,000 GPUs)","not disclosed","Announced Nov 2024; go-live target end-Q1 2025 → launched Apr 2025; phase 2 agreed Mar 2025 (target end-Q2'25) → added early Jul 2025; first B200 cluster launched Q3'25","H200 (phase 1), B200 (phase 2); docs us-central1: B200/H200/RTX PRO 6000","not disclosed","Operational",f"{S['dcd_kc']} (2024-11-19); {S['q4_24']} (2025-02-20); {S['nj_ice']} (2025-03-05); {S['q1_25']} (2025-05-20); {S['q2_25']} (2025-08-07); {S['q3_25']} (2025-11-11)")
add(17,"Vineland – DataOne","Vineland, New Jersey","USA",39.49,-75.03,"approx, city-level","Build-to-suit (DataOne owns/operates shell; Nebius tenant)","DataOne; Northeast Precast (builder); Bloom Energy (on-site fuel cells)","Phase 1 'summer 2025'; 100 MW available in 2025 + 100 MW in 2026 (200 MW contracted, Q2'25)","up to 300 MW (Mar 2025); option for 'hundreds of MW' more (Q3'25)","not disclosed","Announced Mar 5 2025 (agreement Feb 2025 per 20-F); first phase target summer 2025 → 'end of summer', instances fall 2025; MSFT tranche 1 Nov 2025, tranche 2 Feb 2026; Q1'26: majority of remaining capacity 2H26; Q2'26: building finished summer 2026, power switched to Bloom, site-plan amendment hearing adjourned; still in 2026 connected power","not disclosed by Nebius for site","Microsoft – $17.4B (up to $19.4B) 5-yr, served exclusively from NJ; 9 tranches 2025-26 (+2 via Jan-2026 addendum)","Partially operational, ramping",f"{S['nj_ice']} (2025-03-05); {S['q2_25']} (2025-08-07); {S['msft']} (2025-09-08); {S['q3_25']} (2025-11-11); {S['f20']} (2026-04-30); {S['q1_26']} (2026-05-13); {S['q2_26call']} (2026-08-12); {S['nj_page']}")
add(18,"Independence AI factory (Eastgate Commerce Center)","Independence, Missouri (Rte 78 & Little Blue Pkwy)","USA",39.04,-94.33,"approx, from stated intersection","Owned greenfield","— (land from NorthPoint; power from Independence Power & Light / new Blue Valley plant by IPP)","IPL initial power 2H 2026 (MW n/d); first buildings operational 2027","up to 1.2 GW (originally reported 800 MW)","not disclosed","Land Dec 2025; announced as owned site Feb 12 2026; Ch.100 approval Mar 2026; groundbreaking May 12 2026; operational 2027 (Q1'26 letter); additional power study for early 2029; build-out over 5-7 yrs","Vera Rubin planned in US (company-level); site-specific n/d","not disclosed (Meta #2 'multiple locations' — not site-attributed)","Under construction",f"{S['indep']} (accessed 2026-10-04); {S['dcd_indep']} (2026-05-13); {S['dcd_indep800']}; {S['q4_25']} (2026-02-12); {S['q1_26']} (2026-05-13)")
add(19,"Birmingham BHM01 (Oxmoor; former Regions Lakeshore ops center)","Birmingham, Alabama (201 Milan Pkwy)","USA",33.43,-86.86,"approx, from street address","Owned brownfield (via Alabama ADC Holdings LLC)","— (Alabama Power supplies; Hoar Construction GC; Raeden affiliate involved in 201 Milan transaction)","Phase 1 ~100 MW (press-reported)","~300 MW (press-reported; not in Nebius filings)","'electrical capacity' per press","Land Sep 2025; announced Feb 12 2026; Nebius page: construction Q2 2026, initial phase commissioning/energization Q4 2026, full build-out through 2028; Q1'26 letter: operational in 2027","not disclosed","not disclosed","Under construction",f"{S['bham']} (accessed 2026-10-04); {S['dcd_bham']} (2026-02-10); {S['al_com']} (2026-02); {S['q1_26']} (2026-05-13)")
add(20,"Pennsylvania AI factory (Highridge Business Park)","Butler Twp, Schuylkill County, Pennsylvania (I-81 & Gordon Mountain Rd)","USA",40.78,-76.25,"approx, from stated intersection","Owned greenfield","— (PPL 1.2 GW power agreement; land from SEDCO)","~250-300 MW by end-2027 (Q1'26 call); 260 MW by Oct 2027 (DCD/site plans)","up to 1.2 GW; ~660 MW by 2029 (DCD); +~300 MW/yr to 1.2 GW by ~2030 (call)","not disclosed","Announced May 13 2026; land closed May 1 2026 ($187.5M, local press); expansion overlay approved Sep 1 2026; first two buildings construction late 2026 (~3 yrs)","not disclosed","not disclosed","Site prep",f"{S['pa_pr']} (2026-05-13); {S['q1_26']} (2026-05-13); {S['q1_26call']} (2026-05-13); {S['dcd_pa']} (2026-07-27); {S['pa_rh']} (2026-09-01)")
add(21,"Oklahoma colocation (unnamed)","Oklahoma City (third-party only)","USA",35.47,-97.52,"approx, city-level (location unverified)","Colocation (20-F)","not disclosed","not disclosed","not disclosed","not disclosed","Secured Feb 2026 (20-F / Q4'25 letter map)","not disclosed","not disclosed","not disclosed",f"{S['f20']} (2026-04-30); {S['q4_25']} (2026-02-12)")
add(22,"Minnesota colocation (us-north1)","Woodbury, Minnesota (per Nebius docs)","USA",44.92,-92.96,"approx, city-level","Colocation (20-F)","not disclosed","not disclosed","not disclosed","not disclosed","Secured Feb 2026 (20-F); region us-north1 live (docs)","B300 (docs us-north1)","not disclosed","Operational (region listed)",f"{S['f20']} (2026-04-30); {S['docs']}")
add(23,"AIB Data Centers – southeastern US (unnamed)","not disclosed (southeastern US)","USA","","","not disclosed","Colocation / leased (12-yr initial term)","AIB Data Centers Inc. (NYSE American: AIB)","not disclosed","50 MW critical IT (65 MW utility ESA)","IT MW stated (50 MW critical IT); 65 MW utility load","Announced by AIB Sep 30 2026 (not yet in Nebius release); two data halls; dates n/d","not disclosed","not disclosed","Contracted",S['aib']+" (2026-09-30)")

# ---- Quarterly KPI long table ----
Q=["Q1'24","Q2'24","Q3'24","Q4'24","Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26"]
K=[]
def k(entity,q,metric,val,tag,note,src): K.append([entity,q,metric,val,tag,note,src])
ND="not disclosed"
# Company-level
for q in ["Q1'24","Q2'24"]:
    k("Company",q,"Active MW",ND,"Reported (none)","Pre-relisting (Nasdaq trading resumed Oct 2024); no capacity KPI published","—")
    k("Company",q,"GPUs deployed",ND,"Reported (none)","Third-party aggregator claims (~10k GPUs mid-2024) are not company disclosures — excluded","—")
k("Company","Q3'24","GPUs deployed","~10,000 (est.)","Estimated","Q4'24 PR: Hopper base 'nearly doubled QoQ' in Q4; Q3'24 call guided '>20,000 GPUs deployed by YE24' → Q3'24 ≈ 20,000/2 ≈ 10,000","%s; %s"%(S['q3_24call'],S['q4_24']))
k("Company","Q3'24","Active MW",ND,"—","Only Mäntsälä (25 MW site capacity) + Paris (live Nov) existed","—")
k("Company","Q4'24","GPUs deployed",">20,000 (est.)","Estimated","Q3'24 call target >20,000 by YE24; Q4'24 PR: base nearly doubled QoQ (actual count not given)","%s; %s"%(S['q3_24call'],S['q4_24']))
k("Company","Q4'24","Connected MW","25 MW","Reported","Q3'25 letter chart 'Connected power CYE 2024: 25 MW'",S['q3_25'])
k("Company","Q1'25","GPUs deployed","~38,000 (mostly H200, remainder H100)","Reported","Feb 20 2025 call: 'total current capacity ~38,000 GPUs' incl. Iceland & KC, 'by end of March'; Q1'25 letter: mostly H200",'%s; %s'%(S['q4_24call'],S['q1_25']))
k("Company","Q1'25","Active MW",ND,"—","Locations: Finland, Paris, Iceland (KC online Apr)",S['q1_25'])
k("Company","Q2'25","Active MW",ND,"—","5 active locations at end Q2 (Finland, Paris, Iceland, KC, +Israel announced)",S['q2_25'])
k("Company","Q2'25","GPUs deployed",ND,"—","B200 rollout began Q2'25; GB200 GA in Europe Jun 2025",S['q2_25'])
k("Company","Q3'25","Active MW",ND,"—","Guide: 220 MW connected / 100 MW active by YE25",S['q3_25'])
k("Company","Q4'25","Active MW","~170 MW","Reported","Q4'25 letter / 20-F: ~170 MW active at Dec 31 2025 (vs 100 MW target)","%s; %s"%(S['q4_25'],S['f20']))
k("Company","Q4'25","Connected MW","~220 MW (target)","Reported target, actual n/d","YE25 connected target 220 MW (Aug & Nov 2025); actual YE25 connected not explicitly reported",S['q3_25'])
k("Company","Q4'25","Contracted MW","n/d at YE; >2,000 MW by Feb 12 2026","Reported","",S['q4_25'])
k("Company","Q1'26","Active MW",ND,"—","Letter: delivered all commitments; 'significantly increase capacity in place in 3Q'",S['q1_26'])
k("Company","Q1'26","Contracted MW",">3,500 MW (as of May 13 2026)","Reported","owned >75%",S['q1_26'])
k("Company","Q2'26","Active MW",ND,"—","Letter: added capacity in US & Europe late Q2; no MW figure",S['q2_26'])
k("Company","Q2'26","Contracted MW","n/d (YE26 target raised to 5 GW)","Reported target","",S['q2_26'])
# Site-level
def site_rows(name, d):
    for q in Q:
        v=d.get(q,(ND,"—","",""))
        k(name,q,"Live MW / GPUs / GPU gen",v[0],v[1],v[2],v[3])
site_rows("Mäntsälä #1",{
 "Q1'24":("25 MW site capacity (live MW n/d); H100","Reported (capacity)","Pre-expansion capacity 25 MW per DCD",S['dcd_kc']),
 "Q2'24":("25 MW site capacity; H100","Reported (capacity)","",S['dcd_kc']),
 "Q3'24":("25 MW site capacity; H100 (+H200 from Nov)","Reported (capacity)","Expansion announced Oct 8 2024",S['mant75']),
 "Q4'24":("25 MW site capacity; H100/H200","Reported (capacity)","Q4'24 call: first part of expansion due Q3'25",S['q4_24call']),
 "Q1'25":("25 MW; H100/H200","Derived","No change disclosed until Q3'25 phase 1",S['q1_25call']),
 "Q2'25":("25 MW; H100/H200 + GB200 (Europe GA Jun-25, site assumed Finland — inference)","Estimated","GB200 location not stated",S['q2_25']),
 "Q3'25":("Expansion phase 1 targeted late Q3'25 (MW n/d)","Reported (plan)","",S['q1_25call']),
 "Q4'25":("Expansion 'should be completed by end of year' (MW n/d); B300/GB300 deployment began 2025","Reported (plan)","",S['q3_25']),
 "Q1'26":("75 MW built ('completed earlier this year'); GB300 NVL72 operational","Reported","Live/active share n/d",S['lapp_pdf']),
 "Q2'26":("75 MW built; GB300 NVL72","Reported","",S['lapp_pdf'])})
site_rows("Paris (Equinix PA10)",{
 "Q3'24":("Deployment underway; H200","Reported","",S['paris']),
 "Q4'24":("Live Nov 2024; H200; MW n/d","Reported","",S['q4_24']),
 "Q1'25":("Live; H200; MW n/d","Reported","",S['q1_25']),"Q2'25":("Live; H200","Reported","",S['docs']),"Q3'25":("Live; H200","Reported","",S['docs']),"Q4'25":("Live; H200","Reported","",S['docs']),"Q1'26":("Live; H200","Reported","",S['docs']),"Q2'26":("Live; H200","Reported","",S['docs'])})
site_rows("Keflavik (Verne)",{
 "Q4'24":("Agreement signed Dec 2024; not live","Reported","",S['f20']),
 "Q1'25":("10 MW cluster live Mar 2025; H200","Reported","Assumes full 10 MW at go-live ('fully operational' by end-Mar)",S['nj_ice']),
 **{q:("10 MW; H200","Derived","No change disclosed",S['docs']) for q in ["Q2'25","Q3'25","Q4'25","Q1'26","Q2'26"]}})
site_rows("Kansas City (Patmos)",{
 "Q4'24":("Announced; not live (5 MW phase 1 planned)","Reported","",S['dcd_kc']),
 "Q1'25":("Deploying; live Apr 2025 (i.e., not live at Q1 end)","Reported","",S['q1_25']),
 "Q2'25":("5 MW phase 1 live (H200); phase 2 B200 being installed","Reported (MW = initial-phase capacity)","",S['q1_25call']),
 "Q3'25":("Phase 2 added early Jul 2025; first B200 cluster launched; MW n/d (≤40 MW cap)","Reported","",S['q3_25']),
 "Q4'25":("H200 + B200; MW n/d","Reported","",S['q3_25']),"Q1'26":("H200/B200/RTX PRO 6000; MW n/d","Reported","",S['docs']),"Q2'26":("H200/B200/RTX PRO 6000; MW n/d","Reported","",S['docs'])})
site_rows("Vineland (DataOne)",{
 "Q1'25":("Announced (up to 300 MW); under construction","Reported","",S['nj_ice']),
 "Q2'25":("Under construction; 200 MW contracted","Reported","",S['q2_25']),
 "Q3'25":("Opening end of summer 2025 targeted; 100 MW to be available in 2025","Reported (plan)","",S['q2_25']),
 "Q4'25":("MSFT tranche 1 delivered Nov 2025; MW n/d (plan: 100 MW in 2025)","Reported","",S['f20']),
 "Q1'26":("MSFT tranche 2 delivered Feb 2026; MW n/d","Reported","",S['f20']),
 "Q2'26":("Additional MSFT tranches delivered (all to date); building finished summer 2026; MW n/d","Reported","",S['q2_26'])})
site_rows("Modi'in (Mega Or)",{
 "Q2'25":("Lease signed May 2025 (8 MW)","Reported","",S['il_80']),
 "Q3'25":("Facility began operations Sep 2025 (Mega Or); Nebius launch Oct","Reported","",S['il_ctech']),
 **{q:("8 MW; 4,000 HGX B200","Reported","",S['il_pr']) for q in ["Q4'25","Q1'26","Q2'26"]}})
site_rows("London Longcross (Ark)",{
 "Q2'25":("Announced Jun 2025; not live","Reported","",S['uk_ark_jun25']),
 "Q3'25":("Not live (target Q4'25)","Reported","",S['uk_ark_jun25']),
 "Q4'25":("16 MW; 4,000 HGX B300 (phase 1)","Reported","",S['uk_nov25']),
 "Q1'26":("16 MW; B300; phase 2 (+3,000 B300) timing n/d","Reported","Q3'25 letter: next UK phase Q1 2026",S['q3_25']),
 "Q2'26":("16 MW live + 3 MW in delivery; B300 (up to 7,000 GPUs if phase 2 done — unconfirmed)","Reported","",S['uk_ark_dcn'])})
for nm,src in [("Mäntsälä #2",S['dcd_ee']),("Lappeenranta",S['lapp_pdf']),("Béthune",S['reut_beth']),("France 2nd colo",S['f20']),("Harlow (Kao)",S['uk_kao']),("UK 3rd new site",S['uk_pr']),("Masmiyya",S['il_80']),("Beit Shemesh",S['il_80']),("Madrid-Getafe",S['es_dcd']),("Tallinn Hüüru",S['dcd_ee']),("Independence MO",S['indep']),("Birmingham AL",S['bham']),("Pennsylvania",S['dcd_pa']),("Oklahoma colo",S['f20']),("Minnesota colo",S['f20']),("AIB SE-US",S['aib'])]:
    site_rows(nm,{q:("0 MW live (not yet announced/energized in history window)" ,"Derived","Site announced or first capacity dated after the quarter / not live by Q2'26; Minnesota & Oklahoma live status by Q2'26 n/d",src) for q in Q})
# fix Minnesota/Oklahoma: unknown rather than 0 for Q1'26-Q2'26
for r in K:
    if r[0] in ("Minnesota colo","Oklahoma colo") and r[1] in ("Q1'26","Q2'26"):
        r[3]="not disclosed (secured Feb 2026; Minnesota region us-north1 B300 live per docs as of Oct 2026)"; r[4]="—"
    if r[0]=="Béthune" and r[1] in ("Q1'26","Q2'26"): r[3]="0 MW live (first capacity targeted late summer 2026)"

# targets history
T=[
["2024-10-08","Mäntsälä tripled to 75 MW; >$1B Europe AI infra by mid-2025",S['mant75']],
["2024-10-31","Expect >20,000 GPUs deployed by YE2024",S['q3_24call']],
["2025-02-20","~38,000 GPUs by end-Mar 2025; 100 MW 'in operations' by YE2025; new facilities scalable to >300 MW",S['q4_24call']],
["2025-05-20","~100 MW contracted capacity by YE2025; >1 GW pipeline 'in the midterm'",S['q1_25']],
["2025-08-07","220 MW connected by YE2025 (of which 100 MW active); >1 GW contracted by YE2026",S['q2_25']],
["2025-11-11",">2.5 GW contracted by YE2026; 800 MW–1 GW connected by YE2026; 220 MW connected/100 MW active YE25 on track",S['q3_25']],
["2026-02-12","Actual ~170 MW active YE2025; >2 GW contracted already; YE26 target >3 GW; 16 sites; 800 MW–1 GW connected reaffirmed",S['q4_25']],
["2026-03-31",">750 MW contracted in EMEA (own + colo)",S['lapp_pdf']],
["2026-05-13",">3.5 GW contracted; YE26 target >4 GW; owned >75% of contracted; owned = 3 GW across 5 sites; 800 MW–1 GW connected reaffirmed",S['q1_26']],
["2026-08-12","YE26 contracted target 5 GW; deploy >1 GW/yr starting 2027; 800 MW–1 GW connected YE26 reaffirmed on call (active through 1H27); asset-light partner model",f"{S['q2_26']}; {S['q2_26call']}"],
]
import pickle
pickle.dump((cols,R,K,T,S),open('/workspace/nebius/capacity/data.pkl','wb'))
print(len(R),len(K))
