from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
BLUE=Font(color="0000FF"); BOLD=Font(bold=True); HDR=PatternFill("solid",fgColor="DDEBF7"); WR=Alignment(wrap_text=True,vertical="top")
wb=Workbook()
def hdr(ws,row,cols):
    for i,c in enumerate(cols,1):
        x=ws.cell(row=row,column=i,value=c); x.font=BOLD; x.fill=HDR; x.alignment=WR
def widths(ws,w):
    for i,v in enumerate(w,1): ws.column_dimensions[L(i)].width=v

# ---------------- Sources ----------------
S=[
("S1","NVIDIA DGX H100/H200 User Guide – system power 10.2 kW max","https://docs.nvidia.com/dgx/dgxh100-user-guide/introduction-to-dgxh100.html","Reported"),
("S2","NVIDIA DGX SuperPOD H100 data center design – 127 DGX = 1,295 kW + ~116 kW IB/mgmt/storage switches (EAP)","https://docs.nvidia.com/dgx-superpod/design-guides/dgx-superpod-data-center-design-h100/latest/planning.html","Reported"),
("S3","NVIDIA DGX B200 User Guide – 14.3 kW max","https://docs.nvidia.com/dgx/dgxb200-user-guide/introduction-to-dgxb200.html","Reported"),
("S4","NVIDIA DGX B300 User Guide / DC best practices – 14.5 kW (busbar) / 15 kW (AC) est., 19-19.7 kW peak; nvidia.com ~14 kW","https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html ; https://docs.nvidia.com/dgx-pdf/data-center-best-practices-with-dgx-b300-v1.pdf","Reported"),
("S5","NVIDIA NVL72 AI Factory ref. arch – GB300 NVL72 'up to 142 kW'; Lenovo 135 kW TDP (155 peak); HPE 132-140 kW","https://docs.nvidia.com/enterprise-reference-architectures/nvl72-ai-factory/latest/components.html ; https://lenovopress.lenovo.com/lp2357-lenovo-nvidia-gb300-nvl72-rack-scale-ai ; https://buy.hpe.com/emea_europe/en/compute/rack-scale-system/nvidia-nvl-system/nvidia-gb300-nvl72-by-hpe/p/1014890105","Reported"),
("S6","SemiAnalysis InferenceX – GB200 NVL72: 1,200 W TDP, 1.87 kW all-in per GPU; SemiAnalysis H100 vs GB200 (Aug-2025): H100 server ~$190k, all-in ~$250k/server; GB200 NVL72 rack $3.1M, all-in ~$3.9M","https://inferencex.semianalysis.com/chips/gb200-nvl72 ; https://newsletter.semianalysis.com/p/h100-vs-gb200-nvl72-training-benchmarks","Reported (third-party estimate)"),
("S7","Vera Rubin NVL72 rack power – NVIDIA does not publish; Ming-Chi Kuo est. ~190 kW Max-Q / ~230 kW Max-P (via Spheron); Supermicro sizes VR NVL72 SU at 227 kW/rack; 1,152 GPUs (16 racks) per 5 MW envelope","https://www.spheron.network/blog/nvidia-vera-rubin-nvl72-guide/ ; https://www.securities.io/supermicro-ships-nvidia-vera-rubin-nvl72-racks-with-dlc-2-cooling/","Estimated (third-party)"),
("S8","Tom's Hardware (Mar 24 2026) – GB200 NVL72 $2.8-3.4M; GB300 NVL72 $6-6.5M; VR200 NVL72 $5-7M incl ~$1M NAND; VR300 NVL144 $7-8.8M","https://www.tomshardware.com/tech-industry/artificial-intelligence/price-of-nvidias-vera-rubin-nvl72-racks-skyrockets-to-as-much-as-usd8-8-million-apiece-but-server-makers-margins-will-be-tight-nvidia-is-moving-closer-to-shipping-entire-full-scale-systems","Reported (press/source)"),
("S9","Data Gravity (Aug 31 2026) – GB300 NVL72 POs just under $5.0M; ~$5.7M deployed (excl. DC, core switching, storage); 132-140 kW","https://www.datagravity.dev/p/how-much-does-an-nvidia-nvl72-cost","Reported (third-party)"),
("S10","HGX B200 server $360-560k (typical $420-500k); HGX B300 $420-750k (planning $550-650k) – vendor/blog price surveys","https://www.mercatus-ai.com/blog/b200-server-price ; https://gpusmith.com/articles/en/nvidia-hgx-b300-price-specs","Estimated (third-party, low confidence)"),
("S11","Cushman & Wakefield 2026 DC Development Cost Guide (Sep 3 2026) – all-in greenfield US/Canada avg $17.6M/MW excl. chips/GPUs; costs +21%/MW vs Q4-24; powered land $584k/MW","https://ir.cushmanwakefield.com/news/press-release-details/2026/Cushman--Wakefield-Releases-2026-Data-Center-Development-Cost-Guide-Citing-21-Rise-in-Per-MW-Construction-Costs/default.aspx","Reported"),
("S12","JLL 2026 Global Data Center Outlook – ~$11.3M/MW shell & core global; ~10% liquid-cooling premium (from search summary; PDF not opened)","https://www.jll.com/content/dam/jllcom/en/global/documents/reports/research-reports/26-research-global-data-center-outlook-new.pdf","Reported (secondary – verify)"),
("S13","Wholesale colocation asking rents N.America 2026: ~$165/kW-mo (10 MW+), ~$175 (3-10 MW), ~$205 (250-500 kW); Encor ~$196/kW-mo H2-25 (via Data Gravity)","https://datacenterscouts.com/data-center-lease-rates/ ; https://coloprice.com/colocation/wholesale/","Reported (secondary)"),
("S14","Nebius Q4'25 call (Feb 12 2026) – 2026 capex $16-20B; data-center build ≈20% of total capex","https://www.roic.ai/quote/NBIS/transcripts/2025-year/4-quarter","Reported"),
("S15","Nebius Q2'26 letter + call (Aug 12 2026) – deals $20-25M ACV/MW; prepayments cover 50-60% of capex; payback 1y10m on 'associated capex and related operating costs'; Q2 capex ~$5.7B; '2026 base' ACV/MW ~$12M; AI cloud adj. EBITDA margin 50%","/workspace/nebius/src/q2_26.txt ; /workspace/nebius/src/q2t.txt ; https://assets.nebius.com/assets/4462517b-ce83-41f2-96ed-f2ac1bc06a05/SHLQ226%20%281%29.pdf","Reported"),
("S16","Nebius 20-F FY2025 – PP&E gross: servers & network $794.9M (2024) / $3,123.6M (2025); assets not yet in use $163.8M / $2,417.4M; land & buildings $381.4M; 2025 capex (cash) $4,066.0M; Finland 'one of the world's leading PUE'; owned DCs 'up to four times lower overheads' vs industry average","/workspace/nebius/src/20f.txt ; https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm","Reported"),
("S17","Nebius Q1'26 letter definitions – Active = 'power being consumed by installed, operational IT equipment'; Connected = 'power connected into fully built and equipped data centers'; Contracted = land + power commitments","https://assets.nebius.com/assets/aa1bc2e6-df83-40cd-a6a2-95e7cda3d16c/Nebius%20SHL_Q1%202026.pdf","Reported"),
("S18","Nebius site list (this project) – Modi'in 8 MW / 4,000 B200; Longcross 16 MW / 4,000 B300 ph1 (+3,000 ph2), 16 MW live + 3 MW in delivery; Patmos KC 40 MW ~35,000 GPUs (landlord); ~25 MW connected YE24; ~38k GPUs end-Q1'25; ~170 MW active YE25; 220 MW connected YE25 target; 800 MW-1 GW connected YE26","/workspace/nebius/capacity/site_list.md","Reported (as compiled)"),
("S19","Nebius blog Mar 5 2025 – NJ site 'expandable up to a total of 300 MW' (no IT/gross label)","https://nebius.com/blog/posts/300-mw-new-jersey-and-iceland-regions","Reported"),
("S20","Nebius/Bloom PR May 20 2026 – first project 328 MW installed fuel-cell capacity, operational 2026","https://nebius.com/newsroom/nebius-and-bloom-energy-partner-to-power-ai-infrastructure-build-out","Reported"),
("S21","Courier-Post (Aug 20 2026) – Bloom fuel cells 'capable of providing 374 MW, though less is expected to be needed'; nj.com: facility 'expected to draw about 300 MW of power'","https://www.courierpostonline.com/story/money/business/2026/08/20/vineland-planning-board-hearing-vote-dataone-lincoln-avenue-data-center-new-site-plan/91338935007/ ; https://www.nj.com/business/2026/08/you-sold-us-out-nj-town-approves-next-phase-of-massive-ai-data-center-as-residents-erupt.html","Reported (press)"),
("S22","DataOne Q&A site – Vineland 'maximum capacity of 350MW'; 14 data halls; 'data halls and associated mechanical loads will be fed from the fuel cells'; grid as backup for critical network racks","https://askvineland.dataone.eu/","Reported"),
("S23","Measured AI (Sep 9 2026) – claims 315 MW IT; Bloom 328 MW installed / ~250 MW guaranteed; 5 MW grid for mechanical load; >150,000 GB300 at Vineland; $290M upfront to DataOne; Bloom fees up to $2.6B (secondary, partially unsourced)","https://measuredai.substack.com/p/microsoft-new-jersey-ai-data-center","Reported (secondary – unverified)"),
("S24","Vineland Planning Board agendas Aug 5/17 2026 – Phase 1 129,622 sf + Phase 2 587,980 sf; nine Bloom power areas; chiller bldg; LNG (no MW stated)","https://www.vinelandcity.org/Archive/Planning%20Board/Agendas/2026/Agenda%208-17-26%20%28Special%20Meeting%29.pdf","Reported"),
("S25","IREN Microsoft PR (Nov 3 2025) – $9.7B 5-yr; GB300; Dell GPUs+ancillary ~$5.8B; Horizon 1-4 = ~200 MW critical IT","https://www.globenewswire.com/news-release/2025/11/03/3178993/0/en/IREN-Secures-9-7bn-AI-Cloud-Contract-with-Microsoft.html","Reported"),
("S26","IREN Horizon 1 – $300-350M for 75 MW design / 50 MW IT liquid-cooled DC (Feb/Mar 2025)","https://www.datacenterdynamics.com/en/news/iren-plans-75mw-liquid-cooled-ai-data-center-in-texas/ ; https://www.sec.gov/Archives/edgar/data/1878848/000187884825000031/pressreleasemarch2025mon.htm","Reported"),
("S27","IREN FY26 results (Aug 27 2026) – 3-yr contracts >$20M rev/MW (IT), ~2-yr payback (GPU+ancillaries capex / (revenue - direct costs)); prepayments 45-55% of GPU capex","/workspace/nebius/comps_raw/iren_fy26pr__irenreportsfy26results.htm.txt","Reported"),
("S28","CoreWeave Q4'25 & Q2'26 PRs – active >850 MW YE25, 1.5 GW Q2'26; contracted ~3.7 GW; PP&E net $30,557M (YE25) / $46,736M (Q2'26); Sep-26 convert deck: ~4.2 GW contracted","/workspace/nebius/comps_raw/crwv_q2pr__coreweave2q26earningspress.htm.txt ; crwv_q4pr__coreweave4q25earningspress.htm.txt ; crwv_0917__crwv2026convertpresentat.htm.txt","Reported"),
("S29","CoreWeave 2026 capex guide $35-39B; YE26 active >1.85 GW (secondary summary of Q2'26 call)","https://www.computelabs.ai/research/2026-08-11/coreweave-104-billion-backlog-forecast-year-end-active-power-above-1-85","Reported (secondary – verify vs transcript)"),
("S30","Nscale S-1 – ~461,000 active+contracted GPUs; ~1.37 GW active+contracted (1 GW owned, 200 MW leased, 165 MW colo); SIN02 200 MW for >66,000 VR NVL72 GPUs; Monarch 6.5 GW IT / 8 GW gross","/workspace/nebius/comps_raw/nscale_s1.htm.txt","Reported"),
("S31","Crusoe/Blue Owl/Primary Digital – $15B JV for 1.2 GW Abilene; phase 1 (200+ MW) $3.4B","https://www.globenewswire.com/news-release/2025/05/21/3085765/0/en/Crusoe-Blue-Owl-Capital-and-Primary-Digital-Infrastructure-Enter-Second-Phase-of-15-Billion-Joint-Venture-to-Fund-AI-Data-Center-in-Abilene-Texas.html","Reported"),
("S32","DCD – OpenAI/Oracle to deploy >450,000 GB200 at Stargate Abilene (1.2 GW)","https://www.datacenterdynamics.com/en/news/openai-and-oracle-to-deploy-450000-gb200-gpus-at-stargate-abilene-data-center/","Reported (press)"),
("S33","Lambda (Oct 28 2025) – Kansas City: 24 MW at launch, >10,000 Blackwell Ultra GPUs, scalable to >100 MW","https://lambda.ai/blog/lambda-to-build-a-100mw-ai-factory-in-kansas-city-mo","Reported"),
("S34","Jensen Huang – 1 GW AI factory ≈ $50-60B all-in (G20, Sep 2 2026; recurring since 2025); Bernstein ~$35B/GW (GPUs ~39%, land+buildings ~10%); Barclays 65-70% of $50-60B = compute & networking","https://cryptobriefing.com/nvidia-huang-ai-monetization-g20/ ; https://www.businessinsider.com/why-nvidia-worth-5-trillion-inside-35-billion-ai-datacenter-2025-10 ; https://www.benzinga.com/markets/equities/25/09/47935027/nvidias-jensen-math-of-50-billion-revenue-per-gw-gets-a-reality-check-heres-what-barclays-analysts-project-for-nvda","Reported (press)"),
("S35","Locsic (Morgan Stanley BOM teardown, Howard Kao, May 21 2026) – VR200 NVL72 ~$7.8M ODM rack vs GB300 NVL72 ~$4.0M (GPU ~$2.52M, memory ~$374k, liquid cooling ~$65k, ODM value-add ~$108k); ~$6.7M if hyperscaler buys SOCAMM direct","https://locsic.com/thinking/morgan-stanley-rubin-bom-teardown/","Reported (secondary – sell-side via blog)"),
("S36","Wolfe Research via Investing.com/Yahoo (Jan 30 2026) – 'Reported prices are about $3 million for GB200 NVL72 racks and $4.3 million for GB300 NVL72'; Rubin racks $5-6M; boards ~75% of final rack price","https://finance.yahoo.com/news/wolfe-lifts-nvidia-target-25-145033078.html","Reported (sell-side via press)"),
("S37","Loop Capital (Ananda Baruah) via Quartz/heise/TMTPost (Mar 24-25 2025) – Apple ~$1.0B of GB300 NVL72 = ~250 systems at $3.7-4.0M each (Dell + Supermicro)","https://qz.com/apple-nvidia-servers-1-billion-1851772546 ; https://wccftech.com/loop-capital-apple-is-placing-1-billion-worth-of-orders-for-nvidias-gb300nvl72-with-super-micro-computer-smci-and-dell-tapped-as-key-server-partners/","Reported (sell-side via press)"),
("S38","Bernstein via Wccftech (Jun 8 2026) – VR NVL72 rack up to $9.1M on HBM4 $53/GB in 2027 (MS $7.8M uses 'stale memory prices'); memory+storage ~$3.2M/rack","https://wccftech.com/bernstein-warns-nvidias-vera-rubin-racks-will-hit-9-1-million-as-hbm4-prices-triple-to-53-per-gigabyte/","Reported (sell-side via press)"),
("S39","IREN PR (Aug 28 2025) – 1.2k B300 + 1.2k GB300 for ~$168M, footnote: 'Includes servers, InfiniBand, cabling, software and licensing costs' => ~$70k per GPU all-in","https://iren.gcs-web.com/news-releases/news-release-details/iren-expands-ai-cloud-109k-gpus-nvidia-preferred-partner-status","Reported (company)"),
("S40","IREN (Mar 4 2026) – >50,000 B300 purchased; ~$3.5B capex 'including GPUs, servers, storage, networking, labor and ancillary equipment' => ≤~$70k per GPU (HGX, air-cooled)","https://iren.com/resources/news/iren-expands-ai-cloud-capacity-to-150000-gpus","Reported (company)"),
("S41","IREN 8-K/PR (May 19/26 2026) – Dell ~$1.6B 'inclusive of GPUs, servers, storage, networking, ancillary equipment, integration services, and warranties' (air-cooled Blackwell, count n/d) for $3.4B NVIDIA contract in ~60 MW existing Childress capacity","https://www.stocktitan.net/sec-filings/IREN/8-k-iren-ltd-reports-material-event-675eb20880e1.html","Reported (company filing via aggregator)"),
("S42","IREN Microsoft GB300 count '76,000+ GB300s across Horizons 1-4' – third-party IREN thesis site (not IREN); IREN's own releases give 200 MW IT and no GPU count. IREN Aug 13 2026: Horizon 1 delivered/accepted, 50 MW IT","https://irenlimited.com/ ; https://www.sec.gov/Archives/edgar/data/1878848/000114036126032638/ef20080141_ex99-1.htm","Reported (secondary – unverified count)"),
("S43","DCD / Tom's Hardware (Nov 2025) – Microsoft to get >100,000 GB300 under the $17.4B Nebius deal (no Nebius capex disclosed)","https://www.datacenterdynamics.com/en/news/microsoft-to-get-100000-nvidia-gb300s-under-nebius-deal-has-signed-more-than-33bn-in-capacity-agreements-with-neoclouds/","Reported (press)"),
("S44","Nscale PR (Oct 15 2025) – ~200k GB300 for Microsoft with Dell; 104k GB300 in ~240 MW Texas (433/MW, basis n/s); Aker-Nscale Narvik ~$6.2B 5-yr for ~52k GB300 (revenue, not capex)","https://www.nscale.tech/press-releases/nscale-microsoft-2025 ; https://news.cision.com/aker-asa/r/aker-nscale-jv-signs-multi-billion-dollar-ai-infrastructure-agreement-with-microsoft,c4233986","Reported (company)"),
("S45","Hedgehog (Jun 8 2026) – 1,024-GPU 800G leaf-spine on Spectrum-X SN5600 ~$1,500/port: switches $1.54M + cabling $74k-707k (excl. NICs, firewall) = $1.6-2.2k per GPU, ~3% of GB300 GPU BOM; 5.8-6.9% incl. a $1.65M NGFW","https://hedgehog.cloud/blog/network-capex-is-the-smallest-line-on-your-ai-cluster-bom-and-the-one-that-determines-what-every-other-dollar-can-earn","Reported (vendor blog – scale-out fabric only)"),
("S46","Morgan Stanley via Guru3D/Tom's (Nov 2025) – GB300 NVL72 in-rack liquid-cooling content ~$49,860 (18 compute trays × $2,260 + 9 NVLink switch trays × $1,020) – confirms NVLink switch trays sit inside the rack price","https://www.guru3d.com/story/3f5782f43f98056f6165e6938b8fa0f8e06adaeb/","Reported (sell-side via press)"),
("S47","chrislabs.ai (Sep 2026) – capex per GPU incl. share of networking+storage: GB300 NVL72 $61-90k, GB200 $52-66k, VR NVL72 $69-153k; SemiAnalysis TCO owner cost GB300 $2.31/GPU-hr","https://www.chrislabs.ai/tokenomics/gpu-cost","Estimated (third-party)"),
("S48","Nebius Q2'26 6-K Ex-99.2 Note 7 (Aug 12 2026) – server & network equipment gross $3,123.6M (Dec-25) / $5,197.7M (Jun-26); assets not yet in use $7,836.4M; active MW ~183 (team input, Q2'26 letter)","/workspace/nebius/comps_raw/nb_q2_fs.txt ; https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm","Reported (MW = team input)"),
]
ws=wb.active; ws.title="Sources"
hdr(ws,1,["ID","Description / datapoints used","URL / path","Tag"])
for r in S: ws.append(list(r))
widths(ws,[6,110,90,26])
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment=WR

# ---------------- GPUsPerMW ----------------
g=wb.create_sheet("GPUsPerMW",0)
g["A1"]="GPUs per MW by NVIDIA generation – IT-MW and gross (facility) MW basis. Blue = input; black = formula. GPU = one GPU package (VR NVL72 = 72 packages; NVIDIA's 'NVL144' naming counts 144 dies)."; g["A1"].font=BOLD
g["A2"]="Method: kW/GPU at system = unit kW / GPUs per unit; All-in IT kW/GPU = × (1 + overhead for scale-out network, storage, mgmt/CPU nodes outside the unit); GPUs/IT-MW = 1000 / all-in kW; Gross kW/GPU = IT kW × PUE; GPUs/gross-MW = 1000 / gross kW. Low case = high power + high overhead + high PUE; High case = low power + low overhead + low PUE."
cols=["Generation","Scenario","Unit","Unit power kW (input)","GPUs per unit (input)","kW/GPU at unit","Overhead % (input)","All-in IT kW/GPU","GPUs per IT MW","PUE used (input)","Gross kW/GPU","GPUs per gross MW","Tag","Basis / sources"]
hdr(g,4,cols)
gens=[
("Hopper H100/H200","HGX/DGX 8-GPU",8,[(10.2,0.18,1.45,"Low: 10.2 kW max (S1) + 18% overhead (heavier storage/mgmt) + PUE 1.45 (older air-cooled colo)"),
 (10.2,0.10,1.30,"Base: 10.2 kW (S1); +10% ≈ SuperPOD switches/mgmt ~9% of server load (S2) and SemiAnalysis H100 ~1.4 kW all-in (S6); PUE 1.30 air-cooled colo (Estimated)"),
 (9.2,0.07,1.10,"High: provisioned at ~90% of nameplate (Estimated, common power-capping practice) + 7% + PUE 1.10 (Nebius Finland 'leading PUE', S16)")]),
("Blackwell B200 (HGX)","HGX/DGX 8-GPU",8,[(14.3,0.18,1.45,"Low: 14.3 kW max (S3) +18% + PUE 1.45"),
 (14.3,0.10,1.30,"Base: 14.3 kW (S3) +10% + PUE 1.30 (air-cooled colo e.g. Modi'in/KC)"),
 (13.0,0.07,1.10,"High: ~91% of nameplate (Estimated) +7% + PUE 1.10")]),
("GB200 NVL72","NVL72 rack",72,[(132,0.12,1.35,"Low: 132 kW rack (top of 120-132 kW range, S6/HPE) +12% + PUE 1.35"),
 (125,0.07,1.20,"Base: 125 kW (mid of ~120-132 kW) +7% (NVLink switches already in rack) → 1.86 kW ≈ SemiAnalysis 1.87 kW all-in (S6); PUE 1.20 liquid-cooled"),
 (120,0.05,1.10,"High: 120 kW +5% + PUE 1.10")]),
("Blackwell Ultra B300 (HGX)","HGX/DGX 8-GPU",8,[(15.0,0.18,1.45,"Low: 15 kW (DGX B300 AC est., S4) +18% + PUE 1.45"),
 (14.5,0.10,1.30,"Base: 14.5 kW (DGX B300, S4) +10% + PUE 1.30"),
 (14.0,0.07,1.10,"High: ~14 kW (nvidia.com, S4) +7% + PUE 1.10")]),
("GB300 NVL72","NVL72 rack",72,[(142,0.12,1.35,"Low: 'up to 142 kW' (NVIDIA, S5) +12% + PUE 1.35"),
 (137,0.07,1.20,"Base: 137 kW (mid of HPE 132-140 / Lenovo 135 / NVIDIA 142, S5) +7% + PUE 1.20"),
 (132,0.05,1.10,"High: 132 kW (HPE low end, S5) +5% + PUE 1.10")]),
("Vera Rubin NVL72 (VR200; 'NVL144' dies)","NVL72 rack",72,[(230,0.12,1.35,"Low: ~230 kW Max-P (Kuo est., S7; Supermicro sizes 227 kW) +12% + PUE 1.35"),
 (210,0.07,1.20,"Base: 210 kW (midpoint of 190-230 kW est., S7) +7% + PUE 1.20. Estimated – NVIDIA has not published rack power"),
 (190,0.05,1.10,"High: ~190 kW Max-Q (S7) +5% + PUE 1.10")]),
]
r=5; base_rows={}; rows_by_gen={}
for gen,unit,n,sc in gens:
    rows_by_gen[gen]=[]
    for (kw,oh,pue,note),name in zip(sc,["Low","Base","High"]):
        g.cell(r,1,gen); g.cell(r,2,name); g.cell(r,3,unit)
        for col,val in [(4,kw),(5,n),(7,oh),(10,pue)]:
            c=g.cell(r,col,val); c.font=BLUE
        g.cell(r,6,f"=D{r}/E{r}"); g.cell(r,8,f"=F{r}*(1+G{r})"); g.cell(r,9,f"=1000/H{r}")
        g.cell(r,11,f"=H{r}*J{r}"); g.cell(r,12,f"=1000/K{r}")
        g.cell(r,13,"Estimated (spec Reported; overhead/PUE Estimated)"); g.cell(r,14,note)
        for col,f in [(6,"0.000"),(8,"0.000"),(11,"0.000"),(9,"0"),(12,"0"),(7,"0%")]: g.cell(r,col).number_format=f
        rows_by_gen[gen].append(r)
        if name=="Base": base_rows[gen]=r
        r+=1
r+=1; g.cell(r,1,"Summary – GPUs per MW (rounded in display)").font=BOLD; r+=1
hdr(g,r,["Generation","IT-MW Low","IT-MW Base","IT-MW High","Gross-MW Low","Gross-MW Base","Gross-MW High","PUE Low/Base/High"]); r+=1
summ_start=r
for gen,_,_,_ in gens:
    lo,ba,hi=rows_by_gen[gen]
    g.cell(r,1,gen)
    for j,rr in enumerate([lo,ba,hi]): g.cell(r,2+j,f"=I{rr}").number_format="0"; g.cell(r,5+j,f"=L{rr}").number_format="0"
    g.cell(r,8,f'=TEXT(J{lo},"0.00")&" / "&TEXT(J{ba},"0.00")&" / "&TEXT(J{hi},"0.00")')
    r+=1
r+=1
g.cell(r,1,"Notes: (1) Spec power is max/nameplate; actual average draw is typically lower, so operators that provision on measured draw can fit more GPUs per MW (that is the High case). (2) Overhead covers scale-out fabric, storage, CPU/management nodes outside the GPU unit; in-rack NVLink switches are already inside NVL72 rack power. (3) PUE converts IT MW to gross/facility MW; Nebius does not disclose IT vs gross for its MW KPIs (see Nebius checks).").alignment=WR
widths(g,[34,9,16,12,10,10,10,11,11,10,11,12,22,90])

# ---------------- CapexPerMW ----------------
c=wb.create_sheet("CapexPerMW",1)
c["A1"]="Capex per MW by generation – split GPU/server, networking+storage, DC shell+power (owned) vs colocation (lease instead of shell). US$ millions per IT MW unless noted. Blue = input."; c["A1"].font=BOLD
c["A2"]="Density uses the Base GPUs/IT-MW from GPUsPerMW for all three price cases (isolates price uncertainty). Owned total = GPU/server + network/storage + DC shell/power. Colo capex = GPU/server + network/storage + Nebius in-hall fit-out; lease shown separately and as PV (lease-equivalent capex)."
c["A3"]="Global inputs:"; c["A3"].font=BOLD
c["B3"]="Colo lease discount rate"; c["C3"]=0.08; c["C3"].font=BLUE
c["D3"]="Lease term (yrs)"; c["E3"]=10; c["E3"].font=BLUE
c["F3"]="Annuity factor"; c["G3"]="=(1-(1+C3)^-E3)/C3"
cc=["Generation","Scenario","GPUs/IT MW (base, link)","Unit price $M (server/rack, input)","Network+storage+other $M per unit (input)","GPUs per unit","All-in $k/GPU","GPU/server $M/MW","Network+storage $M/MW","DC shell+power $M/MW IT (owned, input)","Owned total $M/IT MW","Owned total $M/gross MW (base PUE)","Implied $B per GW (IT)","Colo in-hall fit-out $M/MW (input)","Colo capex $M/IT MW","Colo lease $/kW-mo (input)","Lease $M/MW-yr","PV lease $M/MW","Colo lease-equivalent $M/IT MW","Tag","Price / DC-cost basis"]
hdr(c,5,cc)
prices={
"Hopper H100/H200":[(0.19,0.05,7,0.5,150,"Server $190k (SemiAnalysis H100 2025, S6); +$50k net/storage; DC $7M/MW (IREN Horizon-type brownfield $6-7M/MW IT, S26); lease $150/kW-mo"),
                    (0.23,0.06,10,1.0,175,"Server $230k (Estimated: between H100 $190k and higher-HBM H200; no primary H200 price found – low confidence); +$60k net/storage (=SemiAnalysis $250k all-in less $190k server, S6); DC $10M/MW air-cooled (Estimated between JLL $11.3M shell-core and IREN $6-7M, S12/S26); lease $175 (S13)"),
                    (0.28,0.08,14,2.0,205,"Server $280k (Estimated, 2024-vintage H200 pricing; low confidence); +$80k; DC $14M/MW; lease $205 (S13)")],
"Blackwell B200 (HGX)":[(0.36,0.065,7,0.5,150,"Server $360k (low end S10); net $65k (Estimated: H100 $60k scaled for 400/800G); DC $7M"),
                    (0.43,0.08,10,1.0,175,"Server $430k (typical $420-500k, S10); net $80k (Estimated); DC $10M/MW air (Estimated)"),
                    (0.50,0.11,14,2.0,205,"Server $500k (S10); net $110k; DC $14M")],
"GB200 NVL72":[(2.8,0.6,8,0.5,150,"Rack $2.8M (Tom's low, S8); net/storage $0.6M; DC $8M liquid (IREN-type, S26 +liquid)"),
               (3.1,0.8,12.5,1.0,175,"Rack $3.1M, all-in $3.9M (SemiAnalysis, S6); DC $12.5M/MW (Crusoe Abilene $15B/1.2GW, S31; ≈JLL $11.3M +10% liquid, S12)"),
               (3.4,1.0,17.6,2.0,205,"Rack $3.4M (Tom's high, S8); net $1.0M; DC $17.6M (Cushman all-in avg excl. GPUs, S11)")],
"Blackwell Ultra B300 (HGX)":[(0.42,0.07,7,0.5,150,"Server $420k (Supermicro listed config, S10); net $70k; DC $7M"),
               (0.55,0.09,10,1.0,175,"Server $550k (planning range $550-650k, S10); net $90k (Estimated, 800G CX-8); DC $10M"),
               (0.65,0.12,14,2.0,205,"Server $650k (S10); net $120k; DC $14M")],
"GB300 NVL72":[("='GB300 rack price'!C21","='GB300 rack price'!C25",8,0.5,150,"LINKED to 'GB300 rack price' sheet. Rack $4.0M (Morgan Stanley BOM ~$4.0M S35; Loop $3.7-4.0M S37); net+storage+install $0.6M; DC $8M"),
               ("='GB300 rack price'!D21","='GB300 rack price'!D25",12.5,1.0,175,"LINKED. Rack $4.5M (between hyperscale-volume $4.0-4.3M S35/S36 and small-order PO ~$5.0M S9; IREN all-in ~$70k/GPU incl IB S39); +$0.9M ≈ SemiAnalysis GB200 $3.9M-$3.1M = $0.8M (S6) scaled for 800G CX-8; DC $12.5M"),
               ("='GB300 rack price'!E21","='GB300 rack price'!E25",17.6,2.0,205,"LINKED. Rack $5.0M (Data Gravity small-order POs, S9); +$1.2M; DC $17.6M (S11). Stress (not used): Tom's $6-6.5M inference config (S8)")],
"Vera Rubin NVL72 (VR200; 'NVL144' dies)":[(6.7,0.9,8,0.5,150,"v2: Rack $6.7M (MS BOM if hyperscaler buys SOCAMM direct, S35; Tom's $5-7M S8, Wolfe $5-6M S36 are older/lower); +$0.9M; DC $8M"),
               (7.8,1.2,12.5,1.0,175,"v2: Rack $7.8M (Morgan Stanley ODM BOM May-26, S35; was $6.0M in v1); +$1.2M (v1 $0.8M too low: CX-9 1.6T doubles per-GPU scale-out bandwidth vs GB300's $0.9M; in-rack NAND offsets only storage); DC $12.5M. Colo capex ≈ team base $41M"),
               (9.1,1.5,17.6,2.0,205,"v2: Rack $9.1M (Bernstein, HBM4 $53/GB, S38); +$1.5M; DC $17.6M. Colo capex ≈ team bear $48.5M")]}
r=6; capex_base={}
for gen,unit,n,_ in gens:
    br=base_rows[gen]
    for (p,net,dc,fit,lease,note),name in zip(prices[gen],["Low","Base","High"]):
        c.cell(r,1,gen); c.cell(r,2,name)
        c.cell(r,3,f"=GPUsPerMW!I{br}").number_format="0"
        for col,val in [(4,p),(5,net),(10,dc),(14,fit),(16,lease)]:
            x=c.cell(r,col,val); x.font=BLUE
        c.cell(r,6,n)
        c.cell(r,7,f"=(D{r}+E{r})/F{r}*1000")
        c.cell(r,8,f"=C{r}*D{r}/F{r}")
        c.cell(r,9,f"=C{r}*E{r}/F{r}")
        c.cell(r,11,f"=H{r}+I{r}+J{r}")
        c.cell(r,12,f"=K{r}/GPUsPerMW!J{br}")
        c.cell(r,13,f"=K{r}")
        c.cell(r,15,f"=H{r}+I{r}+N{r}")
        c.cell(r,17,f"=P{r}*1000*12/1000000")
        c.cell(r,18,f"=Q{r}*$G$3")
        c.cell(r,19,f"=O{r}+R{r}")
        c.cell(r,20,"Estimated (prices = third-party Reported/Estimated; see basis)")
        c.cell(r,21,note)
        for col in [7,8,9,11,12,13,15,17,18,19]: c.cell(r,col).number_format="0.0"
        if name=="Base": capex_base[gen]=r
        r+=1
r+=1; c.cell(r,1,"Summary – Owned all-in capex $M per IT MW (Low / Base / High) and colo capex (excl. lease)").font=BOLD; r+=1
hdr(c,r,["Generation","Owned Low","Owned Base","Owned High","Owned Base $M/gross MW","Colo capex Base","Colo lease-equiv Base","GPU+net share of owned Base"]); r+=1
for gen,_,_,_ in gens:
    b=capex_base[gen]; lo=b-1; hi=b+1
    c.cell(r,1,gen)
    for j,rr in enumerate([lo,b,hi]): c.cell(r,2+j,f"=K{rr}").number_format="0.0"
    c.cell(r,5,f"=L{b}").number_format="0.0"; c.cell(r,6,f"=O{b}").number_format="0.0"; c.cell(r,7,f"=S{b}").number_format="0.0"
    c.cell(r,8,f"=(H{b}+I{b})/K{b}").number_format="0%"
    r+=1
r+=1
c.cell(r,1,"Cross-checks: Jensen Huang $50-60B per GW all-in (S34) ≈ $50-60M/MW; Bernstein $35B/GW (S34); Barclays 65-70% compute+networking (S34) vs our GPU+net share column; IREN $5.8B GPUs+ancillary / 200 MW IT = $29M/MW (Peers). Basis of external $/GW (IT vs gross) is generally NOT stated.").alignment=WR
widths(c,[34,8,10,12,12,8,10,10,10,12,11,12,10,10,10,10,9,9,12,24,90])

# ---------------- Nebius checks ----------------
n=wb.create_sheet("Nebius checks",2)
n["A1"]="Nebius-implied cross-checks (inputs blue = Reported from disclosures/sources; formulas = Derived)"; n["A1"].font=BOLD
hdr(n,3,["Check","Input A","Value A","Input B","Value B","Derived","Unit","Model comparison","Model value","Tag","Interpretation","Source"])
GB=lambda gen,col: f"=GPUsPerMW!{col}{base_rows[gen]}"
checks=[
("Modi'in B200 density","GPUs",4000,"MW",8,"=C{r}/E{r}","GPUs/MW","B200 base GPUs/IT MW",GB("Blackwell B200 (HGX)","I"),"Derived","500/MW ≈ model 508/IT-MW (vs ~391/gross MW) → the 8 MW reads as IT/critical load, or near-IT","S18"),
("Longcross B300 – phase 1","GPUs",4000,"MW",16,"=C{r}/E{r}","GPUs/MW","B300 base GPUs/IT MW",GB("Blackwell Ultra B300 (HGX)","I"),"Derived","250/MW is ~half of model density → phase 1 did not fill the 16 MW","S18"),
("Longcross B300 – ph1+ph2 on 16 MW","GPUs",7000,"MW",16,"=C{r}/E{r}","GPUs/MW","B300 base GPUs/gross MW",GB("Blackwell Ultra B300 (HGX)","L"),"Derived","438/MW sits between gross (386) and IT (502) model density","S18"),
("Longcross B300 – ph1+ph2 on 19 MW","GPUs",7000,"MW",19,"=C{r}/E{r}","GPUs/MW","B300 base GPUs/gross MW",GB("Blackwell Ultra B300 (HGX)","L"),"Derived","368/MW ≈ gross-MW basis","S18"),
("Patmos KC full build (landlord claim)","GPUs",35000,"MW",40,"=C{r}/E{r}","GPUs/MW","H200 HIGH-case GPUs/IT MW","=GPUsPerMW!I{0}".format(rows_by_gen["Hopper H100/H200"][2]),"Derived","875/MW is above even the high-case Hopper IT density (~813) → the landlord's marketing claim isn't physically consistent; don't use it","S18"),
("YE2024 GPUs vs connected MW",">20k GPUs (Estimated from target)",20000,"MW connected YE24",25,"=C{r}/E{r}","GPUs/MW","H100/H200 base GPUs/IT MW",GB("Hopper H100/H200","I"),"Derived","~800/MW is above base IT density (713) → GPU count likely includes GPUs not yet energized, or MW is understated","S18"),
("End-Q1'25 GPUs vs est. MW","GPUs (Reported ~38k)",38000,"MW est. = 25 YE24 + 10 Keflavik (+0 KC live Apr-25)",35,"=C{r}/E{r}","GPUs/MW","H100/H200 high GPUs/IT MW","=GPUsPerMW!I{0}".format(rows_by_gen["Hopper H100/H200"][2]),"Derived (MW Estimated)","~1,086/MW is physically implausible for Hopper → '38k' must include GPUs delivered or allocated but not yet powered. Don't back out MW from GPU counts","S18"),
("YE2025 in-service servers+network per active MW","Gross servers & network $M",3123.6,"Active MW YE25",170,"=C{r}/E{r}","$M/MW","Hopper base GPU+net $M/IT MW","=CapexPerMW!H{0}+CapexPerMW!I{0}".format(capex_base["Hopper H100/H200"]),"Derived","$18.4M/MW is below the model's ~$25.9M/IT MW for Hopper. Possible reasons: active MW is close to gross, GPUs were bought cheaply, or part of the active hardware sits in 'assets not yet in use'","S16/S18"),
("YE2025 upper bound incl. assets not in use","Servers+network + AnYU $M","=3123.6+2417.4","Active MW YE25",170,"=C{r}/E{r}","$M/MW","B200 base GPU+net $M/IT MW","=CapexPerMW!H{0}+CapexPerMW!I{0}".format(capex_base["Blackwell B200 (HGX)"]),"Derived","Upper bound. Assets not yet in use (AnYU) include prepayments and DC infrastructure for future MW, so this overstates $/active MW","S16"),
("Q2'26 in-service servers+network per active MW","Gross server & network equipment Jun-26 $M",5197.7,"Active MW Q2'26 (team input)",183,"=C{r}/E{r}","$M/MW","GB300 base GPU+net $M/IT MW","=CapexPerMW!H{0}+CapexPerMW!I{0}".format(capex_base["GB300 NVL72"]),"Derived","$28.4M/active MW: above Hopper model GPU+net (~$25.9M/IT MW @713/MW) and below GB300 base (~$36.8M) → consistent with active MW = IT and a Hopper-heavy in-service fleet plus a Blackwell tail (~25-30% Blackwell by MW at model prices). If active MW were gross (÷1.2 → 152 IT MW) it would be $34M/IT MW, i.e. a near-all-Blackwell fleet, which the GPU mix doesn't support. Caveat: vintage pricing and partially-energized MW","S48"),
("2026 capex guide / connected MW added – low","Capex guide low $B",16,"MW added (800-220)",580,"=C{r}*1000/E{r}","$M/MW","GB300 base owned $M/IT MW","=CapexPerMW!K{0}".format(capex_base["GB300 NVL72"]),"Derived","$27.6M per connected MW added. Understated because GPUs for capacity connected late in 2026 are bought in 1H27","S14/S18"),
("2026 capex guide / connected MW added – high","Capex guide high $B",20,"MW added (1000-220)",780,"=C{r}*1000/E{r}","$M/MW","—","","Derived","$25.6M/MW; range across both ends ≈ $20.5-34.5M/MW (16/0.78 to 20/0.58)","S14/S18"),
("DC build share of 2026 capex","Capex mid $B",18,"DC share",0.20,"=C{r}*E{r}*1000/680","$M/MW (on 680 MW mid added)","DC shell base liquid $M/MW","=CapexPerMW!J{0}".format(capex_base["GB300 NVL72"]),"Derived","~$5.3M/MW of DC capex per connected MW added. Well below $12.5M owned-build cost, which fits a colo-heavy 2026 (Vineland, Béthune, Israel and UK are leased) plus spend on land/power for 2027+ sites","S14"),
("Payback-implied capex/MW (base)","ACV $M/MW (Q2 deals mid)",22.5,"Payback yrs (1y10m)","=1+10/12","=C{r}*E{r}*0.6","$M/MW","GB300 base GPU+net $M/IT MW","=CapexPerMW!H{0}+CapexPerMW!I{0}".format(capex_base["GB300 NVL72"]),"Estimated","capex ≈ payback × ACV × contribution margin (60% assumed; AI-cloud adj. EBITDA 50%, S15) = ~$24.8M/MW; range 18.3 (20 × 0.5) to 32.1 (25 × 0.7). Sits below the GB300 GPU+net model → per-MW denominator likely gross/contracted MW and/or the payback excludes DC shell","S15"),
("Microsoft ACV per MW (Vineland)","TCV $B / 5 yrs",17.4,"MW (Nebius 'up to 300')",300,"=C{r}/5*1000/E{r}","$M ACV/MW","Nebius '2026 base' ACV/MW (Q2 letter)",12,"Derived","$11.6M/MW on 300 MW, close to the ~$12M '2026 base' ACV/MW in the Q2 letter. On ~240 MW IT it would be $14.5M/MW","S15/S19"),
("Vineland: max facility supply (installed)","Bloom installed MW",328,"Grid MW (mechanical)",5,"=C{r}+E{r}","MW supply","300 MW as IT × base PUE 1.20","=300*1.2","Derived","333 MW installed supply < 360 MW needed if 300 MW were IT at PUE 1.2 → 300 MW reads as facility/gross","S20/S23"),
("Vineland: firm supply (guaranteed)","Bloom guaranteed MW",250,"Grid MW",5,"=C{r}+E{r}","MW supply","Implied IT if 300 = gross, PUE 1.2-1.3","=300/1.25","Estimated","~255 MW firm supply; implied IT ≈ 231-250 MW (300/1.3 to 300/1.2)","S23"),
("Vineland: 150k GB300 claim vs power","GPUs (secondary claim)",150000,"GB300 base all-in IT kW/GPU",GB("GB300 NVL72","H"),"=C{r}*E{r}/1000","MW IT needed","× base PUE → gross MW","=F{r}*GPUsPerMW!J"+str(base_rows["GB300 NVL72"]),"Derived","Needs ~305 MW IT / ~366 MW gross, more than the 333 MW installed supply → the '>150k GB300' claim (S23, unverified) is probably overstated, or a mix with lower-power parts",""),
]
r=4
for ch in checks:
    name,ia,va,ib,vb,der,unit,mc,mv,tag,interp,src=ch
    n.cell(r,1,name); n.cell(r,2,ia); x=n.cell(r,3,va); x.font=BLUE; n.cell(r,4,ib); y=n.cell(r,5,vb); y.font=BLUE
    n.cell(r,6,der.format(r=r) if "{r}" in der else der); n.cell(r,7,unit); n.cell(r,8,mc)
    n.cell(r,9,mv.replace("{r}",str(r)) if isinstance(mv,str) else mv)
    if isinstance(mv,(int,float)): n.cell(r,9).font=BLUE
    n.cell(r,10,tag); n.cell(r,11,interp); n.cell(r,12,src)
    for col in [6,9]: n.cell(r,col).number_format="#,##0.0"
    r+=1
r+=1
n.cell(r,1,"MW basis – what Nebius says").font=BOLD; r+=1
for t in ["Active MW: defined as 'power being consumed by installed, operational IT equipment' (Q1'26 letter, S17). Read literally, that is IT-equipment draw, so IT basis (Reported wording; Nebius never uses the terms 'IT load' or 'critical' explicitly).",
          "Connected MW: 'power connected into fully built and equipped data centers' (S17); Q2'25: includes active power plus power 'fully provisioned… can be activated immediately upon GPU installation'. The basis is ambiguous; it reads as facility/provisioned capacity (Estimated).",
          "Contracted MW: land + power commitments (S17), i.e. utility/grid or generation capacity, so gross basis (Estimated).",
          "Site MW (Vineland 300, Lappeenranta 310, Independence 1.2 GW, PA 1.2 GW, etc.) are never labelled IT or gross. AIB (50 MW critical IT / 65 MW utility, ratio 1.30) and IREN/Nscale (Monarch 6.5 GW IT / 8 GW gross, ratio 1.23) show the typical gap.",
          "Vineland verdict (Estimated, medium confidence): '300 MW' is most consistent with FACILITY/gross power, implying ~230-250 MW IT. Evidence: Nebius 'expandable up to a total of 300 MW' (no label, S19); nj.com 'draw about 300 MW of power' (S21); Bloom 328 MW installed (Nebius PR, S20) / ~250 MW guaranteed (S23) + 5 MW grid for mechanical load (S23). DataOne says fuel cells feed 'data halls and associated mechanical loads' (S22), so the fuel-cell output must cover IT + cooling, and 300 MW IT × PUE ≥1.2 = ≥360 MW exceeds the 333 MW installed. Counter-evidence: Measured AI claims '315 MW IT capacity' (secondary, no primary cite, S23); DataOne cites a 'maximum capacity of 350MW' (S22); Courier-Post says the Bloom plant is 'capable of 374 MW' (S21). The planning-board agendas (S24) give square footage only, no MW."]:
    n.cell(r,1,t).alignment=WR; n.merge_cells(start_row=r,start_column=1,end_row=r,end_column=12); n.row_dimensions[r].height=48; r+=1
widths(n,[38,24,11,26,10,11,14,30,11,16,80,12])

# ---------------- Peers ----------------
p=wb.create_sheet("Peers",3)
p["A1"]="Peer disclosures – GPUs per MW and $ per MW (inputs blue; Derived formulas)"; p["A1"].font=BOLD
hdr(p,3,["Peer","Datapoint","Input A","Value A","Input B","Value B","Derived","Unit","MW basis stated?","Generation","Tag","Comment","Source"])
peers=[
("IREN","Microsoft GB300 GPU+ancillary capex per IT MW","Dell purchase $B",5.8,"IT MW (Horizon 1-4)",200,"=D{r}*1000/F{r}","$M/MW IT","IT (critical)","GB300 NVL72","Derived","$29M/MW IT is GPU+ancillary only. Below our GB300 base GPU+net of ~$40M/IT MW → either IREN provisions fewer GPUs per IT MW (~360-410/MW at $70-80k/GPU) or Dell pricing is ~$4M/rack all-in","S25"),
("IREN","Horizon 1 DC build (low) per IT MW","Capex $M",300,"IT MW",50,"=D{r}/F{r}","$M/MW IT","IT","DC shell (liquid)","Derived","$6.0-7.0M/MW IT ($4.0-4.7M on the 75 MW design figure). Texas brownfield with an existing substation, so a low-end anchor","S26"),
("IREN","Horizon 1 DC build (high) per IT MW","Capex $M",350,"IT MW",50,"=D{r}/F{r}","$M/MW IT","IT","DC shell (liquid)","Derived","","S26"),
("IREN","Microsoft ACV per IT MW","TCV $B/5y",9.7,"IT MW",200,"=D{r}/5*1000/F{r}","$M ACV/MW IT","IT","GB300","Derived","$9.7M/MW IT vs Nebius MSFT $11.6M/MW on 300 MW","S25"),
("IREN","New contracts revenue per IT MW (Reported)","Rev $M/MW IT",20,"Payback yrs",2,"=D{r}*F{r}","$M/MW (rev × payback; upper bound on GPU capex)","IT","Blackwell/GB300","Derived","Upper bound. GPU capex = payback × (rev − direct costs), so it is < $40M/MW","S27"),
("CoreWeave","PP&E net per active MW (YE25)","PP&E net $M",30557,"Active MW",850,"=D{r}/F{r}","$M/MW","not stated","Mixed H100-GB300","Derived","Mostly leased DCs, so this is ≈ GPU+network on a net (depreciated) basis; includes CIP for future MW","S28"),
("CoreWeave","PP&E net per active MW (Q2'26)","PP&E net $M",46736,"Active MW",1500,"=D{r}/F{r}","$M/MW","not stated","Mixed","Derived","$31.2M/MW","S28"),
("CoreWeave","2026 capex guide mid / active MW added","Capex mid $B",37,"MW added (1850-850)",1000,"=D{r}*1000/F{r}","$M/MW added","not stated","GB300/VR","Derived","~$37M per incremental active MW (the guide is from a secondary summary, so verify it)","S29/S28"),
("Nscale","GPUs per MW (portfolio)","GPUs active+contracted",461000,"MW active+contracted",1365,"=D{r}/F{r}","GPUs/MW","not stated","GB300/VR mix","Derived","338/MW. Between our GB300 gross (~409) and VR gross (~267)","S30"),
("Nscale","SIN02 Portugal – VR NVL72","GPUs",66000,"MW",200,"=D{r}/F{r}","GPUs/MW","not stated ('200 MW powered land')","Vera Rubin NVL72","Derived","330/MW ≈ our VR base IT (320) → if 200 MW is gross, this implies higher density or a lower-power VR profile","S30"),
("Nscale","Monarch IT/gross ratio","Gross GW",8,"IT GW",6.5,"=D{r}/F{r}","gross/IT","both","—","Derived","1.23 PUE-equivalent (includes behind-the-meter generation losses)","S30"),
("AIB/Nebius","SE US lease IT vs utility","Utility MW",65,"Critical IT MW",50,"=D{r}/F{r}","gross/IT","both","n/d","Derived","1.30 ratio on a Nebius-contracted site","S18"),
("Oracle/OpenAI","Stargate Abilene GB200","GPUs",450000,"MW",1200,"=D{r}/F{r}","GPUs/MW","not stated","GB200 NVL72","Derived","375/MW, below our GB200 gross base (~448) → 1.2 GW is likely utility/gross with headroom","S32"),
("Crusoe","Abilene JV $ per MW","JV $B",15,"MW",1200,"=D{r}*1000/F{r}","$M/MW","not stated","DC shell+power","Derived","$12.5M/MW DC (no GPUs)","S31"),
("Crusoe","Abilene phase 1 $ per MW","Financing $B",3.4,"MW (200+)",200,"=D{r}*1000/F{r}","$M/MW","not stated","DC shell+power","Derived","≤$17M/MW ('200+ MW')","S31"),
("Lambda","Kansas City launch density","GPUs (>)",10000,"MW at launch",24,"=D{r}/F{r}","GPUs/MW","not stated","B300 (Blackwell Ultra)","Derived","≥417/MW. Between B300 gross (386) and IT (502)","S33"),
("Supermicro","VR NVL72 blueprint","GPUs per SU",1152,"MW envelope",5,"=D{r}/F{r}","GPUs/MW (envelope)","envelope (gross-like)","Vera Rubin NVL72","Derived","230/MW envelope; 16 racks × 227 kW = 3.63 MW IT → envelope/IT 1.38","S7"),
("NVIDIA (Huang)","All-in $ per GW (low)","$B per GW",50,"—",1,"=D{r}/F{r}","$M/MW","not stated","Blackwell/Rubin","Reported","$50-60M/MW all-in","S34"),
("NVIDIA (Huang)","All-in $ per GW (high)","$B per GW",60,"—",1,"=D{r}/F{r}","$M/MW","not stated","Blackwell/Rubin","Reported","","S34"),
("Bernstein","All-in $ per GW","$B per GW",35,"GPU share",0.39,"=D{r}*F{r}","$B GPUs per GW","not stated","GB200/Rubin","Derived","GPUs ~$13.7B/GW; land+buildings ~10%","S34"),
("Cushman & Wakefield","Greenfield DC all-in excl. GPUs","$M/MW avg",17.6,"Range low/high","8.9-23.3","=D{r}","$M/MW","per MW (basis n/s)","DC","Reported","+21% vs Q4-24","S11"),
("Colocation market","Wholesale rent 10 MW+","$/kW-mo",165,"3-10 MW",175,"=D{r}*12*1000/1000000","$M/MW-yr","per kW leased (IT)","—","Derived","≈$2.0-2.5M per MW-yr ($165-205)","S13"),
]
r=4
for pr in peers:
    for i,v in enumerate(pr,1):
        if i==7: v=v.format(r=r)
        x=p.cell(r,i,v)
        if i in (4,6) and isinstance(v,(int,float)): x.font=BLUE
    p.cell(r,7).number_format="#,##0.0"
    r+=1
widths(p,[18,40,22,10,22,10,10,20,20,18,10,70,10])
for ws_ in [p,n,c,g]:
    for row in ws_.iter_rows(min_row=4):
        for x in row: x.alignment=WR
# ---------------- GB300 rack price (v2, Oct 4 2026) ----------------
q=wb.create_sheet("GB300 rack price",2)
q["A1"]="GB300 NVL72 rack price – sourced input (v2, Oct 4 2026). Rack price INCLUDES 18 compute trays (72 B300 + 36 Grace, ConnectX-8 NICs) and 9 in-rack NVLink switch trays + in-rack liquid cooling (S35/S46/S9). EXCLUDES scale-out fabric (leaf/spine switches, optics), front-end/storage network, storage, install/integration, DC. Blue = input."; q["A1"].font=BOLD
hdr(q,3,["#","Source","Date","Price $M / rack","Per GPU $k","What's included","Buyer/volume","Tag","URL / ID"])
obs=[("Loop Capital (Apple order ~$1B / ~250 systems)","2025-03-24",3.85,"System (Dell/SMCI), range $3.7-4.0M; scope n/s","Large (Apple)","Reported (sell-side via press)","S37"),
("Morgan Stanley BOM (Howard Kao)","2026-05-21",4.0,"ODM rack price, in-rack NVLink + cooling; no scale-out","ODM sell-in","Reported (secondary)","S35"),
("Wolfe Research supply-chain checks","2026-01-30",4.3,"'Reported prices' per rack","n/s","Reported (sell-side via press)","S36"),
("Data Gravity purchase orders","2026-08-31",4.97,"EXW rack, no CDU ('just under $5.0M'); +$32k with CDU; deployed ~$5.7M excl core switching, storage, DC","Small orders","Reported (third-party)","S9"),
("Tom's Hardware source (inference-optimized config)","2026-03-24",6.25,"Range $6-6.5M; config/warranty differ; many online quotes = ODM price w/o warranty","n/s","Reported (press)","S8"),
("IREN Aug-2025 order: 2.4k B300+GB300 for ~$168M","2025-08-28","=168/2400*72","ALL-IN incl servers, InfiniBand, cabling, software (blended HGX B300 + GB300)","Small (2.4k GPUs)","Derived (company $ & count)","S39"),
("IREN-Dell Microsoft GB300 $5.8B / 76k GPUs","2025-11-03","=5800/76000*72","ALL-IN 'GPUs and ancillary'; GPU count from secondary site","Large (Dell)","Derived (count unverified)","S25/S42"),
("GB200 ref: SemiAnalysis hyperscaler rack / all-in","2025-08-20",3.1,"Rack $3.1M; all-in incl networking, storage, other $3.9M (+$0.8M, +26%)","Hyperscaler","Reported (third-party)","S6"),
("GB200 ref: Wolfe / Tom's","2026",3.0,"Wolfe ~$3.0M; Tom's $2.8-3.4M","n/s","Reported (press)","S36/S8")]
r=4
for i,(src,dt,pr,inc,buy,tag,sid) in enumerate(obs,1):
    q.cell(r,1,i); q.cell(r,2,src); q.cell(r,3,dt); x=q.cell(r,4,pr); x.font=BLUE; x.number_format="0.00"
    q.cell(r,5,f"=D{r}/72*1000").number_format="0.0"; q.cell(r,6,inc); q.cell(r,7,buy); q.cell(r,8,tag); q.cell(r,9,sid); r+=1
r=19; q.cell(r,1,"Selected inputs ($M)").font=BOLD
hdr(q,20,["Item","Unit","Low","Base","High","Stress","Basis"])
rows=[("GB300 NVL72 rack price","$M/rack",4.0,4.5,5.0,6.25,"Low = MS BOM / Loop (hyperscale-type volume); Base = midpoint of volume ($4.0-4.3M) and small-order POs ($5.0M) – Nebius buys at >100k-GPU scale (S43) but via OEM/ODM in a memory-inflation year; High = Data Gravity POs; Stress = Tom's"),
("Racks per IT MW (base density, link)","racks","=GPUsPerMW!I{0}/72".format(base_rows["GB300 NVL72"]),"=GPUsPerMW!I{0}/72".format(base_rows["GB300 NVL72"]),"=GPUsPerMW!I{0}/72".format(base_rows["GB300 NVL72"]),"=GPUsPerMW!I{0}/72".format(base_rows["GB300 NVL72"]),"491 GPUs/IT MW (137 kW + 7%)"),
("Compute $M/IT MW","$M/MW","=C21*C22","=D21*D22","=E21*E22","=F21*F22","Derived"),
("  memo: scale-out fabric only $k/GPU (2-tier 800G)","$k/GPU",1.6,2.5,4.0,4.0,"Hedgehog Spectrum-X 1,024 GPUs $1.6-2.2k (S45); larger 3-tier clusters + optics higher (Estimated)"),
("Network+storage+install/other $M/rack","$M/rack",0.6,0.9,1.2,1.2,"Bucket = scale-out + front-end + storage + install/integration. SemiAnalysis GB200 all-in +$0.8M/rack (S6); Data Gravity +$0.7M deploy (S9) before core switching/storage; GB300 800G CX-8. NICs + NVLink already in rack price"),
("Network etc. $M/IT MW","$M/MW","=C25*C22","=D25*D22","=E25*E22","=F25*F22","Derived ($6.1M base unchanged)"),
("Compute + network $M/IT MW","$M/MW","=C23+C26","=D23+D26","=E23+E26","=F23+F26","Compare to cutoff"),
("Colo fit-out $M/IT MW","$M/MW",0.5,1.0,2.0,2.0,"as CapexPerMW"),
("Leased capex (no shell) $M/IT MW","$M/MW","=C27+C28","=D27+D28","=E27+E28","=F27+F28","Equals CapexPerMW colo capex"),
("IT-reading cutoff (team)","$M/MW",33,33,33,33,"Team breakeven for IT reading of $20-25M ACV/MW & 22-mo payback"),
("Headroom vs cutoff (+ = IT reading works)","$M/MW","=C30-C27","=D30-D27","=E30-E27","=F30-F27","Derived"),
("Breakeven rack price at this network cost","$M/rack","=C30/C22-C25","=D30/D22-D25","=E30/E22-E25","=F30/F22-F25","Derived"),
("Breakeven GPUs/IT MW at this rack+network","GPU/MW","=C30/(C21+C25)*72","=D30/(D21+D25)*72","=E30/(E21+E25)*72","=F30/(F21+F25)*72","Derived – IREN's $5.8B/200 MW IT/76k = 380/IT MW passes")]
r=21
for it,u,*v,basis in rows:
    q.cell(r,1,it); q.cell(r,2,u)
    for j,val in enumerate(v):
        x=q.cell(r,3+j,val); x.number_format="0.00"
        if not (isinstance(val,str) and val.startswith("=")): x.font=BLUE
    q.cell(r,7,basis); r+=1
q.cell(r+1,1,"Density sensitivity – compute+network $M/IT MW at base prices").font=BOLD
hdr(q,r+2,["GPUs/IT MW","Racks/MW","Compute+net $M/MW","vs $33M"])
rr=r+3
for d in [380,453,491,519]:
    q.cell(rr,1,d).font=BLUE; q.cell(rr,2,f"=A{rr}/72"); q.cell(rr,3,f"=B{rr}*($D$21+$D$25)"); q.cell(rr,4,f"=$D$30-C{rr}")
    for c_ in (2,3,4): q.cell(rr,c_).number_format="0.0"
    rr+=1
q.cell(rr+1,1,"Peer cross-checks: IREN-Dell $5.8B / 200 MW IT = $29.0M/IT MW all-in GPU+ancillary (S25) – below the cutoff but at lower density (~380/MW if 76k GPUs; at 491/MW it would imply ~$4.25M per rack incl network). IREN all-in ~$70k/GPU on small 2025 order incl InfiniBand (S39) = $5.0M per 72 GPUs all-in → $34.4M/IT MW at 491/MW. Rubin cross-check (GPU Market team): MS VR $7.8M ÷ 1.95 = $4.0M; Bernstein $9.1M ÷ 1.68 = $5.4M mixes Bernstein's price with MS's ratio (ratio is internal to MS's BOM) – treat as upper bound only.").alignment=WR
widths(q,[44,38,12,12,12,40,22,26,14])
for row in q.iter_rows(min_row=4):
    for x in row: x.alignment=WR

wb.save("/workspace/nebius/capacity/gpu_mw_capex.xlsx")
print("ok")
