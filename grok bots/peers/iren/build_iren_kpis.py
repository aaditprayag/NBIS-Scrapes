# Builds iren_kpis.xlsx (IREN Ltd peer KPI comps, calendar Q1'24-Q2'26 = IREN FQ3 FY24 - FQ4 FY26) on the CoreWeave/Nebius template.
# Engine: /workspace/nebius/peers/peer_kpi_template.py (same row IDs / definitions / tagging as peers/coreweave/build_crwv_kpis.py).
# Author: Grok Bot for Aadit Prayag, 2026-10-04/05 ET. Every hard-coded figure carries a tag + source key + section/formula.
import sys; sys.path.insert(0, "/workspace/nebius/peers")
from peer_kpi_template import Peer, build

OUT = "/workspace/nebius/peers/iren/iren_kpis.xlsx"
QS = ["Q4'23*", "Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26"]
QEND = ["2023-12-31", "2024-03-31", "2024-06-30", "2024-09-30", "2024-12-31", "2025-03-31", "2025-06-30", "2025-09-30", "2025-12-31", "2026-03-31", "2026-06-30"]
FQ = ["FQ2 FY24 (Dec-23)", "FQ3 FY24 (Mar-24)", "FQ4 FY24 (Jun-24)", "FQ1 FY25 (Sep-24)", "FQ2 FY25 (Dec-24)", "FQ3 FY25 (Mar-25)", "FQ4 FY25 (Jun-25)", "FQ1 FY26 (Sep-25)", "FQ2 FY26 (Dec-25)", "FQ3 FY26 (Mar-26)", "FQ4 FY26 (Jun-26)"]
DAYS = [None, 91, 91, 92, 92, 90, 91, 92, 92, 90, 91]
E = "https://www.sec.gov/Archives/edgar/data/1878848/"
S = {
 "Q24Q3": ("6-K FQ3 FY24 (Mar-24 qtr; IFRS) - Ex.99.1 results PR / Ex.99.3 interim financial statements, filed 2024-05-15", E + "000114036124026268/ef20027345_ex99-3.htm"),
 "K24": ("Form 20-F FY24 (FYE 30-Jun-2024; IFRS), filed 2024-08-28", E + "000162828024038677/iren-20240630.htm"),
 "PR24": ("FY24 results PR (6-K Ex.99.1, 2024-08-28)", E + "000162828024038679/fy24resultspressreleasev.htm"),
 "Q25Q1": ("6-K FQ1 FY25 (Sep-24 qtr; IFRS) - interim FS + results PR, filed 2024-11", E + "000187884824000036/iren-20240930_d2.htm"),
 "Q25Q2": ("6-K FQ2 FY25 (Dec-24 qtr; IFRS) - interim FS / MD&A Ex.99.4 / PR, filed 2025-02", E + "000187884825000009/iren-20241231_d2.htm"),
 "Q25Q3": ("6-K FQ3 FY25 (Mar-25 qtr; IFRS) - interim FS / MD&A Ex.99.4 / PR, filed 2025-05", E + "000187884825000043/iren-20250331_d2.htm"),
 "K25": ("10-K FY25 (FYE 30-Jun-2025; first US GAAP annual; incl. unaudited quarterly table), filed 2025-08-28", E + "000187884825000063/iren-20250630.htm"),
 "PR25": ("FY25 results PR (8-K Ex.99.1, 2025-08-28)", E + "000187884825000062/ex991fy25pr.htm"),
 "Q26Q1": ("10-Q FQ1 FY26 (Sep-25 qtr), filed 2025-11-06", E + "000187884825000081/iren-20250930.htm"),
 "PR26Q1": ("Q1 FY26 results PR (8-K Ex.99.1, 2025-11-06)", E + "000187884825000079/q1fy26resultspressreleas.htm"),
 "Q26Q2": ("10-Q FQ2 FY26 (Dec-25 qtr), filed 2026-02-05", E + "000187884826000015/iren-20251231.htm"),
 "PR26Q2": ("Q2 FY26 results PR (8-K Ex.99.1, 2026-02-05)", E + "000187884826000014/a991irenreportsq2fy26res.htm"),
 "Q26Q3": ("10-Q FQ3 FY26 (Mar-26 qtr), filed 2026-05-07", E + "000187884826000026/iren-20260331.htm"),
 "PR26Q3": ("Q3 FY26 results PR (8-K Ex.99.1, 2026-05-07)", E + "000187884826000025/irenreportsq3fy26results.htm"),
 "K26": ("10-K FY26 (FYE 30-Jun-2026), filed 2026-08-27", E + "000187884826000052/iren-20260630.htm"),
 "PR26": ("FY26 results PR (8-K Ex.99.1, 2026-08-27)", E + "000187884826000051/irenreportsfy26results.htm"),
 "XBRL": ("SEC companyfacts XBRL CIK 0001878848 (local copy comps_raw/cf_0001878848.json)", "https://data.sec.gov/api/xbrl/companyfacts/CIK0001878848.json"),
 "NBGPU": ("Nebius workspace: /workspace/nebius/capacity/gpu_mw_capex.xlsx (GPUsPerMW base IT-MW rows) / gpu_mw_capex.md s.1", "local"),
 "CAL": ("calendar", ""),
}
P = Peer("IREN (IREN) - AI Cloud perimeter", OUT, QS, QEND, FQ, DAYS, S)
nd = None
add = P.add

# ---------------- A. P&L (AI Cloud perimeter = I01) ----------------
add("I01","AI Cloud Services revenue (segment)","$M",[nd,0.567,2.538,3.189,2.661,3.581,6.963,7.347,17.298,33.635,70.515],
    [None,"R","D","R","R","R","R","R","R","R","D"],
    ["","Q25Q3","K26","Q26Q1","Q26Q2","Q26Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","FQ3 FY25 6-K Ex.99.4 MD&A table, 3m to 31-Mar-24 comparative 'AI cloud service revenue 567' (IFRS)",
     "= FY24 3.105 (FY26 10-K segment note, US GAAP recast) - 9M FY24 0.567 (IFRS, FQ3 FY25 MD&A) -> mixed basis",
     "Q1 FY26 10-Q segment note comparative (3m Sep-24)","Q2 FY26 10-Q comparative (3m Dec-24)","Q3 FY26 10-Q comparative (3m Mar-25)",
     "FY25 10-K unaudited quarterly table (Jun-25 qtr 6,963)","Q1 FY26 10-Q segment note","Q2 FY26 10-Q segment note","Q3 FY26 10-Q segment note",
     "= FY26 128.795 (FY26 10-K segment note) - 9M 58.280 (Q3 FY26 10-Q)"],
    "KPI perimeter = AI Cloud Services segment only (all AI revenue generated in Canada through FY26 Q3 per 10-Q; Childress Texas from H2 CY26). 2024 quarters (FQ3-FQ4 FY24) are IFRS; from FY25 10-K onward US GAAP (FY24/FY25 recast).")
add("I02","Adj. EBITDA of KPI perimeter (AI segment)","$M",[nd]*11,[None]*11,"","",
    "n/d: IREN reports Adj. EBITDA only company-wide (memo I47). Segment measure is gross profit excl. D&A (see K45).")
add("I03","Depreciation & amortization (company-wide, P&L)","$M",[nd,8.692,26.780,33.931,36.077,47.311,63.817,85.226,99.176,121.245,112.082],
    [None,"R","D","R","D","R","D","R","D","R","D"],
    ["","Q24Q3","K24","Q26Q1","Q26Q3","Q26Q3","K25","Q26Q1","Q26Q3","Q26Q3","K26"],
    ["","FQ3 FY24 6-K Ex.99.3 P&L 'Depreciation (8,692)' (IFRS)","= FY24 50.650 (20-F Note 6) - 9M 23.870 (6-K) (IFRS)",
     "Q1 FY26 10-Q PP&E note comparative ($33,931,000)","= 9M FY25 117.319 - Q3 47.311 - Q1 33.931 (Q3 FY26 10-Q comparatives)","Q3 FY26 10-Q segment reconciliation comparative",
     "= FY25 181.136 - 9M 117.319","Q1 FY26 10-Q","= 9M FY26 305.647 - Q3 121.245 - Q1 85.226","Q3 FY26 10-Q","= FY26 417.729 - 9M 305.647"],
    "Company-wide (mining + AI + buildings); IREN does not allocate D&A to segments.")
add("I04","Cash capex (company-wide: PP&E net of hardware prepayments + computer/mining hardware incl. prepayments)","$M",
    [nd,nd,296.6,383.4,187.6,443.8,357.9,280.6,719.1,1355.3,1978.2],
    [None,None,"D","R","D","D","D","R","D","D","D"],
    ["","","K24","Q26Q1","Q26Q2","Q26Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","n/d (6M FY24 cash flow not in local set)","= FY24 (141.855 + 338.054 mining-hardware prepayments) - 9M (112.664 + 70.638) (IFRS cash flows)",
     "Q1 FY26 10-Q CF comparative","YTD difference (10-Q CF)","YTD difference","YTD difference (FY25 10-K CF)","Q1 FY26 10-Q CF","YTD difference","YTD difference","YTD difference (FY26 10-K CF)"],
    "Company-wide; not split AI vs mining. FY26 capex is predominantly AI (Childress Horizon GPU halls, BC ASIC->GPU conversions, GPUs) but no split is disclosed.")
add("I05","Gross property, plant & equipment (EoP)","$M",[nd,427.510,538.026,1010.401,1288.150,1699.480,2112.200,2396.859,3553.7,4987.0,7663.316],
    [None,"D","D","D","D","D","R","R","R","R","R"],
    ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","= sum of at-cost lines, IFRS Note 10 (land 1.911 + buildings 169.765 + P&E 4.704 + mining hw 144.623 + HPC hw 29.603 + development assets 76.904)",
     "= sum of at-cost lines, 20-F Note 14","= sum of at-cost lines, Note 12","= sum of at-cost lines, Note 12","= sum of at-cost lines, Note 13",
     "PP&E note 'Property, plant and equipment, gross'","PP&E note","PP&E note","PP&E note","PP&E note"],
    "IFRS (to Mar-25) shows accumulated impairment separately from gross; GAAP (Jun-25+) likewise ('Less: Impairment'). Basis break between Mar-25 and Jun-25.")
add("I06","Construction in progress / development assets (EoP; not in service)","$M",[nd,76.904,102.946,153.625,177.639,213.763,237.734,464.832,1038.321,2089.371,3657.072],
    [None]+["R"]*10, ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","IFRS 'Development assets - at cost'"]+["IFRS 'Development assets - at cost'"]*4+["'Construction in progress'"]*5)
add("I07","Property, plant & equipment, net (EoP)","$M",[nd,357.081,441.371,965.994,1209.339,1575.107,1930.567,2115.436,3170.5,4369.9,6753.183],
    [None]+["R"]*10, ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"], "Balance sheet / PP&E note")
add("I08","Technology equipment, gross (CRWV row; n/d - see I31 GPU hardware, I48 mining hardware)","$M",[nd]*11,[None]*11,"","")
add("I09","Operating AI Cloud capacity, reported (EoP)","MW",[nd]*10+[40],[None]*10+["R"],[""]*10+["K26"],[""]*10+["10-K FY26 Item 1: 'As of June 30, 2026, our operating AI Cloud Services capacity represented approximately 40MW'"],
    "Only AI MW figure IREN has disclosed. Basis (IT vs gross) not stated. Earlier quarters estimated from GPUs (I09E). Dec-25/Mar-26: n/d (GPU counts no longer disclosed; Horizon 1 at Childress (50MW IT) delivered to Microsoft only after FY26 per FY26 PR).")
add("I10","Contracted AI MW (EoP)","MW",[nd]*11,[None]*11,"","",
    "n/d as a metric. FY26 PR: Microsoft contract = four 50MW (IT) liquid-cooled 'Horizon' deployments at Childress (=200MW IT, D) - not entered because other contracts' MW are not disclosed.")
add("I11","AI GPUs operating, reported (EoP)","GPUs",[nd,816,816,nd,1896,1896,1900,2100,nd,nd,nd],
    [None,"R","R",None,"R","R","R","R",None,None,None],
    ["","Q24Q3","PR24","","Q25Q3","Q25Q3","K25","Q26Q1","","",""],
    ["","FQ3 FY24 PR: '816 NVIDIA H100 GPUs'","FY24 PR (Aug-24): '816 NVIDIA H100 GPUs'","",
     "FQ3 FY25 MD&A: GPUs increased 'from 816 as of March 31, 2024 to 1,896 as of March 31, 2025'; Q1 FY25 PR (Nov-24) already cites 1,896 H100 & H200 -> Dec-24 1,896",
     "FQ3 FY25 MD&A (1,896 as of 31-Mar-25)","10-K FY25: 'approximately 1,900 NVIDIA GPUs operating'","10-Q Q1 FY26: 'As of September 30, 2025, we had approximately 2,100 GPUs operating'","","",""],
    "Dec-25 onward IREN discloses only GPUs 'installed or contracted/on order' (e.g., 140k expansion target, Q1 FY26 PR) - not operating counts. Sep-24 estimated in I11E.")
add("I12","Hopper price proxy: IREN's own revenue assumption ($/GPU-hr at 100% uptime)","$/GPU-hr",[nd,2.00,2.00,2.00,2.00,2.00,2.00,2.00,nd,nd,nd],
    [None]+["E"]*7+[None]*3, "Q25Q1", ["","Q1 FY25 PR economics footnote: '$2.00 per GPU hour revenue assumption' (applied to all Hopper-era quarters)"]+["same"]*6+["","",""],
    "IREN publishes no on-demand list price ('contact sales'). The $2.00 is IREN's own assumed realized rate, so K09 for IREN is closer to a utilization ratio than CRWV/Nebius (list-based). Not comparable 1:1.")
add("I13","Blackwell list price","$/GPU-hr",[nd]*11,[None]*11,"","","No public list price.")
add("I14","RPO (ASC 606, EoP)","$B",[nd]*7+[0.195027,0.289411,0.710272,5.1],[None]*7+["R"]*4,[""]*7+["Q26Q1","Q26Q2","Q26Q3","K26"],
    [""]*7+["10-Q Note: '$195,027,000 of unsatisfied RPO'","'$289,411,000'","'$710,272,000'","10-K FY26 RPO note: $5.1B (XBRL RevenueRemainingPerformanceObligation)"],
    "Pre-Sep-25 n/d (no RPO disclosure under IFRS 6-Ks / FY25 10-K). Microsoft $9.7B TCV enters RPO only as tranches are delivered and accepted (nil at Dec-25/Mar-26).")
add("I15","Contracted backlog incl. non-RPO (RPO + undelivered Microsoft TCV / ASC 842 lease value)","$B",[nd]*8+[9.989,10.410,16.6],[None]*8+["D","D","R"],
    [""]*8+["Q26Q2","Q26Q3","K26"],[""]*8+["= RPO 0.289 + Microsoft TCV 9.7 (undelivered; 10-Q 'Strategic Customer Agreement')","= RPO 0.710 + Microsoft 9.7","10-K FY26: $5.1B RPO + $11.4B ASC 842 lease contract value = $16.6B"],
    "Dec-25/Mar-26 are lower bounds (other signed-not-delivered contracts not quantified in filings).")
add("I16","RPO % expected to be recognized within 24 months","%",[nd]*7+[95.8,98.7,80.7,43.1],[None]*7+["D"]*4,[""]*7+["Q26Q1","Q26Q2","Q26Q3","K26"],
    [""]*7+["= (93.363 + 93.410)/195.027","= (150.072 + 135.580)/289.411","= (307.950 + 265.198)/710.272","= (0.9 + 1.3)/5.1 (10-K FY26 RPO maturity)"])
add("I17","Largest customer, % of quarter revenue","%",[nd]*11,[None]*11,"","",
    "n/d for AI Cloud. Mining: 97%/98%/100% of FY25/FY24/FY23 revenue via 3/3/2 mining pool operators (FY25 10-K) - pools are not end-customers in the CRWV sense. Microsoft ($9.7B) and NVIDIA ($3.4B) are the 'substantial majority' of contracted value at Jun-26 (backlog, not revenue).")
add("I18","2nd-largest customer %","%",[nd]*11,[None]*11,"","")
add("I19","3rd-largest customer %","%",[nd]*11,[None]*11,"","")
add("I20","Customers >=10% of revenue (count)","#",[nd]*11,[None]*11,"","")
add("I21","Finance expense (P&L; incl. issuance-cost amortization & lease interest)","$M",[nd,nd,nd,0.022,6.253,4.119,5.182,9.280,10.668,14.837,24.466],
    [None,None,None,"R","R","R","D","R","D","R","D"],
    ["","","","Q26Q1","Q25Q2","Q26Q3","K25","Q26Q1","Q26Q3","Q26Q3","K26"],
    ["","","","Q1 FY26 10-Q comparative (GAAP)","FQ2 FY25 6-K P&L 'Finance expense (6,253)' (IFRS)","Q3 FY26 10-Q comparative","= FY25 11.045 - 9M 5.863","Q1 FY26 10-Q","= 9M 34.785 - Q1 9.280 - Q3 14.837","Q3 FY26 10-Q","= FY26 59.251 - 9M 34.785"],
    "Not contractual coupon interest (IREN gives contractual interest only annually: FY26 $37.2M, FY25 $9.1M per 10-K debt note). Zero debt Mar-24-Sep-24.")
add("I22","Capitalized interest","$M",[nd]*11,[None]*11,"","","Not disclosed quarterly.")
add("I23","Total debt principal (EoP; excl. finance leases)","$M",[nd,0,0,0,440,440,990,990,3745.699,3745.699,7705.649],
    [None]+["R"]*10, ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"], "Debt note principal (see Debt_Facilities). Excludes GPU finance leases ($243.8M liability at Jun-26).")
add("I24","Company-reported weighted average interest rate","%",[nd]*11,[None]*11,"","")
add("I25","Company-definition capex","$M",[nd]*11,[None]*11,"","")
add("I26","Days in quarter","days",DAYS,[None]+["R"]*10,"CAL","calendar")
add("I27","Data center sites (operating)","#",[nd]*11,[None]*11,"","")
add("I28","Useful life - HPC / GPU hardware","years",[nd]+[5]*10,[None]+["R"]*10,["","K24","K24","K25","K25","K25","K25","K26","K26","K26","K26"],
    ["","20-F FY24: 'HPC hardware is depreciated over 5 years'"]+["same"]+["10-K FY25 useful-life table 'HPC hardware 5'"]*4+["10-K FY26 table 'HPC hardware 5'"]*4,
    "Mining hardware 4 yrs (S19j Pro/S19 Pro shortened to 2 yrs in FY24); buildings 20; other 3-10. No change to HPC life disclosed FY24-FY26.")
add("I29","Company-reported GPU utilization","%",[nd]*11,[None]*11,"","")
add("I30","RPO % expected to be recognized within 12 months","%",[nd]*7+[47.9,51.9,43.4,17.6],[None]*7+["D"]*4,[""]*7+["Q26Q1","Q26Q2","Q26Q3","K26"],
    [""]*7+["= 93.363/195.027","= 150.072/289.411","= 307.950/710.272","= 0.9/5.1"])
add("I31","GPU / HPC hardware, gross at cost (EoP; owned, excl. finance-lease ROU)","$M",[nd,29.603,33.315,33.678,68.190,69.406,76.001,106.047,550.636,811.390,2034.811],
    [None]+["R"]*10, ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","IFRS Note 10 'HPC Hardware' at cost","20-F Note 14 'HPC hardware - at cost'","'HPC hardware - at cost'","same","same","10-K FY25 'HPC Hardware'","10-Q 'HPC hardware'","10-Q 'GPU hardware'","10-Q 'GPU hardware'","10-K FY26 'GPU Hardware'"],
    "Excludes GPU finance-lease right-of-use assets (I49) and GPUs in CIP/prepayments. In-service AI hardware, so the cleanest AI-only asset base.")
add("I32","AI Cloud segment cost of revenue (excl. D&A)","$M",[nd,nd,nd,0.232,0.275,0.336,0.476,0.728,2.395,4.597,9.212],
    [None]*3+["R","R","R","D","R","R","R","D"], ["","","","Q26Q1","Q26Q2","Q26Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","","","segment note comparative","segment note comparative","segment note comparative","= FY25 1.319 - 9M 0.843","segment note","segment note","segment note","= FY26 16.932 - 9M 7.720"],
    "Mostly electricity. IFRS 2024 quarters: no segment cost split.")
# ---------------- memo (IREN-specific, I40+) ----------------
add("I40","Bitcoin mining revenue (segment)","$M",[nd,53.383,54.260,49.575,113.483,141.242,180.330,232.948,167.394,111.160,66.710],
    [None,"R","D","R","R","R","R","R","R","R","D"],
    ["","Q25Q3","K26","Q26Q1","Q26Q2","Q26Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],
    ["","FQ3 FY25 MD&A comparative 3m Mar-24 (IFRS)","= FY24 184.087 (FY26 10-K segment note) - 9M FY24 129.827 (IFRS)","segment note comparative","comparative","comparative","FY25 10-K quarterly table","segment note","segment note","segment note","= FY26 578.212 - 9M 511.502"])
add("I42","Accumulated impairment of PP&E (EoP)","$M",[nd,25.714,25.605,6.941,6.934,6.942,0.385,16.177,36.3,171.5,484.491],
    [None]+["R"]*10, ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"], "PP&E note 'Less: accumulated impairment' (mining hardware)")
add("I43","Total operating data-center capacity, company (all uses; gross MW)","MW",[nd,260,260,360,510,660,810,810,810,nd,420],
    [None,"R","R","R","R","R","R","R","R",None,"D"],
    ["","Q24Q3","K24","PR24","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","","K26"],
    ["","FQ3 FY24 PR: '260MW of operating data centers'","20-F FY24 (260MW at 30-Jun-24; FY24 PR of Aug-24 says 300MW)","Q1 FY25 PR (Nov-24): '360MW of operating data centers'","FQ2 FY25 PR: '510MW'","FQ3 FY25 PR: 660MW","10-K FY25: 810MW","Q1 FY26 10-Q/PR: 810MW","Q2 FY26: 810MW","n/d","= Bitcoin ~380MW (10-K FY26 'approximately 23.2 EH/s, representing approximately 380MW') + AI ~40MW"],
    "Several values are boilerplate 'as of PR date' (5-7 weeks after quarter-end). Gross/grid basis; AI 40MW basis unstated -> sum is approximate.")
add("I44","Bitcoin mining capacity (MW), reported","MW",[nd]*10+[380],[None]*10+["R"],[""]*10+["K26"],[""]*10+["10-K FY26: '23.2 EH/s, representing approximately 380MW'"])
add("I45","Installed hashrate (EoP)","EH/s",[nd,10,10,21,31,37,50,50,46,38,23.2],[None]+["R"]*10,
    ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],"Results PR / 10-K operating highlights (installed EH/s)")
add("I46","Secured grid-connected power (company)","MW",[nd,2160,2310,2310,2310,2910,2910,3000,4500,4510,5000],[None]+["R"]*10,
    ["","Q24Q3","PR24","Q25Q1","Q25Q2","Q25Q3","K25","PR26Q1","PR26Q2","PR26Q3","PR26"],
    "PR 'grid-connected power secured' (Q1 FY26 '3GW', Q2 FY26 '>4.5GW', FY26 '~5GW' = rounded / lower bounds)")
add("I47","Adj. EBITDA, company-wide (non-GAAP)","$M",[nd,21.8,12.2,2.5,62.6,83.3,121.9,91.7,75.3,59.5,19.2],
    [None,"R","D","R","R","R","R","R","R","R","R"],
    ["","Q25Q3","K24","PR26Q1","Q25Q3","Q25Q3","PR25","PR26Q1","PR26Q2","PR26Q3","PR26"],
    ["","FQ3 FY25 MD&A comparative (IFRS, 3m Mar-24 21.840)","= FY24 54.7 (IFRS) - 9M 42.5 (IFRS)","results PR","FQ3 FY25 PR '($62.6 million in Q2 FY25)' (revised)","FQ3 FY25 PR $83.3M","FY25 PR quarterly","Q1 FY26 PR","Q2 FY26 PR","Q3 FY26 PR","FY26 PR (Q4)"],
    "Company-wide incl. mining, SG&A, unallocated. Not used in K05.")
add("I48","Mining hardware, gross at cost (EoP)","$M",[nd,144.623,177.766,521.755,643.475,877.934,1135.584,1134.032,1079.663,1041.767,596.988],[None]+["R"]*10,
    ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"],"PP&E note 'Mining hardware'")
add("I49","GPU finance-lease right-of-use assets, gross (EoP)","$M",[nd]*8+[160.042,301.776,301.848],[None]*8+["R"]*3,[""]*8+["Q26Q2","Q26Q3","K26"],"PP&E note 'Right-of-use assets - Finance lease'")

P.param("P01","GB200 NVL72 GPUs per IT-MW (base)","GPUs/MW",538,"E","NBGPU","GPUsPerMW base IT-MW row (487/538/571)")
P.param("P03","Hopper HGX GPUs per IT-MW (base)","GPUs/MW",713,"E","NBGPU","GPUsPerMW base IT-MW row (665/713/813); IREN's own economics footnote assumes 1.25kW/GPU (= 800 GPUs/MW)")

# ---------------- derived ----------------
N = 11
P.derive("I11E","AI GPUs incl. estimates (EoP)","GPUs",
    lambda i: {3: "[I11@2]"}.get(i, "[I11@%d]" % i) if i <= 7 else "[I09@10]*[P01@0]",
    lambda i: [("I11",2)] if i == 3 else ([("I11",i)] if i <= 7 else ([("I09",10),("P01",0)] if i == 10 else [("I11",i)])),
    "Sep-24 ESTIMATED = Jun-24 816 (1,896 expansion went live ~Nov/Dec-24). Jun-26 ESTIMATED = 40 MW x GB200-class 538 GPUs/IT-MW (Nebius base) = ~21.5k; range 19.6k (GB300 491) to 32k (IREN 1.25kW/GPU). Dec-25/Mar-26 n/d.",
    force=lambda i: "E" if i in (3, 10) else None)
P.derive("I11A","Average AI GPUs in quarter","GPUs", lambda i: "([I11E@%d]+[I11E@%d])/2" % (i-1, i), lambda i: [("I11E",i-1),("I11E",i)], "Average of I11E.", idx=range(1, N))
P.derive("I09E","AI MW incl. estimates (EoP)","MW", lambda i: "[I09@10]" if i == 10 else "[I11E@%d]/[P03@0]" % i,
    lambda i: [("I09",10)] if i == 10 else [("I11E",i),("P03",0)],
    "ESTIMATED to Sep-25 = operating GPUs / 713 Hopper GPUs per IT-MW (Nebius base); Jun-26 = reported 40MW. (IREN 1.25kW/GPU would give ~12% higher MW.)",
    force=lambda i: None if i == 10 else "E")
P.derive("I09A","Average AI MW in quarter","MW", lambda i: "([I09E@%d]+[I09E@%d])/2" % (i-1, i), lambda i: [("I09E",i-1),("I09E",i)], "", idx=range(1, N))
P.derive("I09T4","AI MW added over trailing 4 quarters","MW", lambda i: "[I09E@%d]-[I09E@%d]" % (i, i-4), lambda i: [("I09E",i),("I09E",i-4)],
    "Jun-26: 40 (reported) - Jun-25 est. 2.7 = ~37.3 MW (Estimated).", idx=range(4, N))
P.derive("I12W","Fleet-weighted price proxy","$/GPU-hr", lambda i: "[I12@%d]" % i, lambda i: [("I12",i)], "= I12 (all Hopper fleet through Sep-25).", idx=range(1, N), force=lambda i: "E")
P.derive("I05S","In-service gross PP&E (gross - CIP - accumulated impairment, EoP)","$M", lambda i: "[I05@%d]-[I06@%d]-[I42@%d]" % (i, i, i),
    lambda i: [("I05",i),("I06",i),("I42",i)], "Company-wide. Impairment deducted because IREN carries impaired mining hardware gross.")
P.derive("I23A","Average debt principal","$M", lambda i: "([I23@%d]+[I23@%d])/2" % (i-1, i), lambda i: [("I23",i-1),("I23",i)], "", idx=range(1, N))
P.derive("I41","Total revenue (AI + Bitcoin)","$M", lambda i: "[I01@%d]+[I40@%d]" % (i, i), lambda i: [("I01",i),("I40",i)], "Segment sum = consolidated revenue.")
P.derive("I44E","Bitcoin mining MW incl. estimates","MW", lambda i: "[I44@10]" if i == 10 else "[I43@%d]-[I09E@%d]" % (i, i),
    lambda i: [("I44",10)] if i == 10 else [("I43",i),("I09E",i)], "ESTIMATED = total operating DC MW - est. AI MW (mixes gross and IT bases; AI share was <2% so error is small). Jun-26 reported 380MW.",
    force=lambda i: None if i == 10 else "E")

# ---------------- debt ----------------
SNAP = ["Mar-24","Jun-24","Sep-24","Dec-24","Mar-25","Jun-25","Sep-25","Dec-25","Mar-26","Jun-26"]
SNAPQ = list(range(1, 11))
SNAPSRC = ["Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26"]
FAC = [
 ("3.25% convertible senior notes due 2030 (issued Dec-24)","Converts",[None]*3+[(440,3.87),(440,3.87),(440,3.87),(440,3.87),(212.31,3.87),(212.31,3.87),(212.31,3.87)]),
 ("3.50% convertible senior notes due 2029 (issued Jun-25)","Converts",[None]*5+[(550,4.22),(550,4.22),(233.389,4.22),(233.389,4.22),(233.389,4.22)]),
 ("0.00% convertible senior notes due 2031 (issued Dec-25)","Converts",[None]*7+[(1000,0.37),(1000,0.37),(1000,0.37)]),
 ("0.25% convertible senior notes due 2032 (issued Dec-25)","Converts",[None]*7+[(1150,0.45),(1150,0.45),(1150,0.45)]),
 ("1.00% convertible senior notes due Jun-2033 (issued Dec-25)","Converts",[None]*7+[(1150,1.18),(1150,1.18),(1150,1.18)]),
 ("1.00% convertible senior notes due Dec-2033 (issued FY26 H2)","Converts",[None]*9+[(3000,1.18)]),
 ("GPU delayed-draw term loan (non-recourse, SOFR+2.25%; $1,545M facility)","DDTL",[None]*9+[(413,7.12)]),
 ("USPP senior secured notes (5.96% coupon; $2,100M programme)","Senior notes",[None]*9+[(525,7.05)]),
 ("Other debt (various)","Other term loans",[None]*9+[(21.95,None)]),
]
CATS = ["DDTL","Senior notes","Converts","OEM / vendor financing","Revolver","Other term loans"]
debt = dict(SNAP=SNAP, SNAPQ=SNAPQ, SNAPSRC=SNAPSRC, FAC=FAC, CATS=CATS, SECT="Debt / convertible notes note (principal; effective interest rate)",
  RATE_NOTE="Effective rates as printed in the FY26 10-K debt note (Note 23) and FY26 10-Qs; for the 2029/2030 converts the same instrument-level effective rate is applied to earlier snapshots (Dec-24-Sep-25). Converts were partly repurchased in Dec-25 (2030: 440->212.3; 2029: 550->233.4). 'Other debt' has no disclosed rate and is excluded from weighted-rate denominators. Mar-24-Sep-24: no debt outstanding (FY24 contractual interest nil).",
  POST=["POST-PERIOD (not in Jun-26): Aug-2026 $2.4B GPU financing (Blue Owl / PIMCO) at 9.0% fixed (FY26 PR / 8-K). Marginal secured GPU cost ~7-9% vs 1.9% blended (converts-heavy).",
        "Finance leases (GPU leases, liability $243.8M at Jun-26) excluded, consistent with CRWV/Nebius template."])

ovr = {
 "K01": dict(note="AI Cloud revenue x4 / operating AI MW. Only Jun-26 MW reported (~40MW, basis unstated). IREN FY26 PR: recent 3-yr contracts '>$20m revenue per MW (IT)', active discussions at ~$25m/MW -> fleet 7.1 reflects legacy Hopper/on-demand mix and capacity live <1 quarter."),
 "K02": dict(note="Average-MW variant. MW ESTIMATED from GPUs / 713 per IT-MW through Sep-25; Dec-25/Mar-26 MW n/d so Q4'25-Q2'26 n/d."),
 "K03": dict(note="n/d - no contracted-MW metric (Microsoft = 4 x 50MW IT Horizons; other contracts' MW not disclosed)."),
 "K04": dict(note="n/d (no contracted MW)."),
 "K05": dict(note="n/d - IREN reports Adj. EBITDA only company-wide (memo I47; FY26 $245.7M incl. mining). See K45 for AI segment gross margin."),
 "K06": dict(note="Avg GPUs (I11A) reported except Sep-24 (Estimated). Last computable Q3'25 (GPU counts stop). Q2'26 memo M07 uses estimated period-end GPUs."),
 "K07": dict(note="All-in AI revenue per GPU-hour on operating GPUs. Note: IREN H100/H200 fleet; GPUs listed as 'operating' may include unrented capacity."),
 "K08": dict(note="ESTIMATED: IREN publishes no list price; uses IREN's own $2.00/GPU-hr 100%-uptime revenue assumption (Q1 FY25 PR). Different basis from CRWV/Nebius (on-demand list)."),
 "K09": dict(note="Realized $/GPU-hr / IREN's own $2.00 assumption -> closer to a physical-utilization x price-realization index than CRWV/Nebius (list-based, which embeds contract discounts). Not comparable 1:1."),
 "K10": dict(note="n/d - capex is company-wide (mining + AI; no split) and AI MW is n/d quarterly. Context: FY26 cash capex $4.33B vs +~37MW operating AI MW = $116M/MW, meaningless because $3.66B sits in CIP (Childress Horizons, Sweetwater) not yet operating."),
 "K11": dict(note="n/d (see K10)."),
 "K12": dict(note="n/d (no AI new-MW series)."),
 "K13": dict(idx=[], note="Suppressed: trailing-4Q company capex / AI MW added would be ~116 $M/MW (timing + perimeter mismatch). See memo M06 (GPU hardware per AI MW)."),
 "K14": dict(idx=[], note="n/d on AI perimeter (net PP&E is company-wide). Company-wide analogue in memo M05 (net PP&E / total operating DC MW)."),
 "K15": dict(idx=[], note="n/d on AI perimeter. AI-only analogue: memo M06 = GPU hardware gross / AI MW."),
 "K16": dict(note="10-K FY26 / FY25 / 20-F FY24 useful-life table: HPC (GPU) hardware 5 yrs; mining hardware 4 yrs (2 yrs for S19j Pro/S19 Pro from FY24); buildings 20. No change to GPU life FY24-FY26. Nebius: 4->5 yrs from 1-Jan-2026; CRWV 6 yrs."),
 "K17": dict(note="Company-wide D&A x4 / avg in-service gross PP&E (gross - CIP - impairment). Q3'24-Q1'25 IFRS, Q2'25 straddles IFRS->GAAP. High yields reflect 2-4 yr mining hardware lives; falling to ~10% as GPU (5y) and buildings (20y) dominate."),
 "K18": dict(note=""),
 "K19": dict(expr=lambda I, i: f"{I('I03',i)}/{I('I41',i)}*100", deps=lambda i: [("I03",i),("I41",i)], name="D&A as % of revenue (company D&A / total revenue)", note="Overridden to company-wide denominator (D&A is not segmented)."),
 "K20": dict(note="AI Cloud RPO only (mining has none). Microsoft $9.7B enters RPO tranche-by-tranche on delivery; Jun-26 RPO $5.1B excludes $11.4B ASC 842 lease value (K21)."),
 "K21": dict(note="RPO + contracted value outside RPO: Dec-25/Mar-26 = RPO + undelivered Microsoft TCV (Derived, lower bound); Jun-26 = $16.6B reported (RPO 5.1 + ASC 842 11.4). Microsoft ~$9.7B (~58%) + NVIDIA $3.4B (~20%) = 'substantial majority' (Derived concentration)."),
 "K22": dict(note="Derived from RPO maturity buckets (0-12m, 13-24m). Jun-26 43% (CRWV 41%, Nebius 36%)."),
 "K23": dict(note="RPO / (AI Cloud quarterly revenue x4)."),
 "K24": dict(note="n/d - IREN does not disclose AI customer revenue shares. Mining revenue via 3 pools (97% of FY25 revenue). Backlog concentration (K21 note): Microsoft ~58%, NVIDIA ~20% of $16.6B."),
 "K25": dict(note="n/d"), "K26": dict(note="n/d"), "K27": dict(note="n/d"),
 "K28": dict(note="Principal-weighted effective rates (converts at 0.4-4.2% effective dominate). Jun-26 incl. first secured GPU debt (DDTL 7.12%, USPP 7.05%). comps.csv had 1.85% (consistent)."),
 "K29": dict(note="GPU DDTL (non-recourse, SOFR+2.25%), first drawn FY26 Q4."),
 "K30": dict(note="USPP senior secured notes (5.96% coupon; 7.05% effective)."),
 "K31": dict(note="Convertible notes (zero/low coupon; effective incl. issuance costs)."),
 "K36": dict(note="Finance expense (incl. issuance-cost amortization, finance-lease interest) x4 / avg principal. Annual contractual interest: FY26 $37.2M / avg debt (990; 7,706) = 0.85% (contractual coupons only)."),
 "K37": dict(note="n/d"),
 "K38": dict(note="Excl. GPU finance leases ($243.8M at Jun-26)."),
 "K40": dict(note="n/d (capitalized interest not disclosed quarterly)."),
 "K41": dict(note="n/d - IREN does not publish GPU utilization."),
 "K42": dict(note="Derived from RPO maturity buckets."),
 "K43": dict(note="(AI rev Q2'26 - Q2'25) x4 / (40 MW - est. 2.7 MW at Jun-25). Estimated (Jun-25 MW from GPUs/713). Contract-level IREN claim: >$20M/MW (IT) on recent 3-yr contracts."),
 "K44": dict(note="AI revenue x4 / avg owned GPU hardware gross (excl. finance-lease ROU and CIP). Falls in Q4'25-Q2'26 because large GPU batches were capitalized ahead of revenue ramp (Horizon/Microsoft)."),
 "K45": dict(note="AI segment (revenue - segment cost of revenue excl. D&A) / revenue; costs mostly power."),
}

memo = [
 dict(kid="M01", name="AI Cloud share of total revenue", unit="%", expr=lambda I, i: f"{I('I01',i)}/{I('I41',i)}*100", deps=lambda i: [("I01",i),("I41",i)], note="Mining-to-AI transition: <2% (2024) -> 51% (Q2'26).", fmt="0.0"),
 dict(kid="M02", name="Bitcoin mining revenue per mining MW (annualized)", unit="$M/MW-yr", expr=lambda I, i: f"{I('I40',i)}*4/(({I('I44E',i-1)}+{I('I44E',i)})/2)", deps=lambda i: [("I40",i),("I44E",i-1),("I44E",i)], note="Mining MW ESTIMATED (total DC MW - AI MW) except Jun-26; avg of open/close. Contrast with AI rev/MW (K01/K02).", fmt="0.00"),
 dict(kid="M03", name="AI MW share of operating data-center MW", unit="%", expr=lambda I, i: f"{I('I09E',i)}/{I('I43',i)}*100", deps=lambda i: [("I09E",i),("I43",i)], note="Estimated (AI MW from GPUs) except Jun-26 (40/420).", fmt="0.0"),
 dict(kid="M04", name="Installed hashrate", unit="EH/s", expr=lambda I, i: f"{I('I45',i)}", deps=lambda i: [("I45",i)], note="Peak 50 EH/s (Jun/Sep-25) -> 23.2 EH/s as ASICs replaced by GPUs.", fmt="0.0", passthru=True),
 dict(kid="M05", name="Net PP&E per total operating DC MW (company-wide)", unit="$M/MW", expr=lambda I, i: f"{I('I07',i)}/{I('I43',i)}", deps=lambda i: [("I07",i),("I43",i)], note="Company-wide; includes CIP for unbuilt capacity (Jun-26 CIP $3.7B).", fmt="0.0"),
 dict(kid="M06", name="Owned GPU hardware (gross) per AI MW", unit="$M/MW", expr=lambda I, i: f"{I('I31',i)}/{I('I09E',i)}", deps=lambda i: [("I31",i),("I09E",i)], note="Hopper era ~$11-13M/MW (est. MW). Jun-26 $51M/MW suggests GPU hardware in service well ahead of 'operating' AI MW (40MW) - or MW stated on a different basis.", fmt="0.0"),
 dict(kid="M07", name="Revenue per GPU on period-end est. GPUs ($K/GPU-yr) - Q2'26 bridge", unit="$K/GPU-yr", expr=lambda I, i: f"{I('I01',i)}*4000/{I('I11E',i)}", deps=lambda i: [("I01",i),("I11E",i)], idx=[10], note="Estimated GPUs (40MW x 538). Range on 19.6k-32k GPUs: $8.8K-14.4K.", fmt="0.0"),
 dict(kid="M08", name="Company Adj. EBITDA margin", unit="%", expr=lambda I, i: f"{I('I47',i)}/{I('I41',i)}*100", deps=lambda i: [("I47",i),("I41",i)], note="Company-wide; Q2'26 compressed by AI ramp costs and shrinking mining.", fmt="0.0"),
 dict(kid="M09", name="Secured grid power / operating DC MW", unit="x", expr=lambda I, i: f"{I('I46',i)}/{I('I43',i)}", deps=lambda i: [("I46",i),("I43",i)], note="Pipeline multiple (company-wide).", fmt="0.0"),
]

readme = [
 "Peer KPI comps - IREN Ltd (NASDAQ: IREN), AI Cloud perimeter, calendar Q1'24-Q2'26. Built 2026-10-05 (ET) for the Nebius (NBIS) project. File: /workspace/nebius/peers/iren/iren_kpis.xlsx; builder: build_iren_kpis.py (same folder) on shared engine /workspace/nebius/peers/peer_kpi_template.py.",
 "TEMPLATE: identical to peers/coreweave/crwv_kpis.xlsx - same row IDs (I01-I27, K01-K40), definitions D1-D8, tag colours and sourcing rules. Phase-2 extensions: I28-I32 / K41-K45 (useful life, reported GPU utilization, RPO<12m, incremental rev per incremental MW, revenue / GPU hardware, segment gross margin); I40+ / M01+ are IREN-specific memo rows (mining split).",
 "FISCAL MAPPING: IREN FY ends 30-Jun, so fiscal quarters end on calendar quarter-ends -> exact mapping (Q1'24 = FQ3 FY24 Mar-24 ... Q2'26 = FQ4 FY26 Jun-26). See Fiscal_Map tab and row 2 of KPIs/Inputs.",
 "PERIMETER: I01 = AI Cloud Services segment revenue; MW/GPUs = AI only. Bitcoin mining shown separately (I40, I44-I48, M01-M04). Capex, D&A, PP&E, debt are company-wide (not segmented) - KPIs mixing perimeters are suppressed (n/d with explanation) or labelled company-wide.",
 "ACCOUNTING BASIS: IFRS for FQ3 FY24-FQ3 FY25 (6-K / 20-F); US GAAP from FY25 10-K (FY24/FY25 recast). Basis breaks are flagged on each input.",
 "TAGS / COLOURS: blue = Reported; black = Derived (arithmetic; formula in Sources); orange = Estimated (assumption); grey n/d. KPI tag = weakest input.",
 "UNITS: $M unless stated; $/MW KPIs annualized ($M per MW-year); $/GPU in $K per GPU-year.",
 "Multi-peer Comparison tab: /workspace/nebius/peers/comparison.xlsx (built by peers/build_comparison.py).",
]
defs_app = {
 "D1": "I01 = AI Cloud Services revenue. MW = 'operating AI Cloud Services capacity' (~40MW at Jun-26, only disclosure; basis unstated). Earlier AI MW ESTIMATED from operating GPUs / 713 Hopper GPUs per IT-MW. IREN quotes contract economics per MW (IT): '>$20m revenue per MW' on recent 3-yr contracts (FY26 PR).",
 "D2": "Operating GPUs disclosed to Sep-25 (816 -> ~2,100); thereafter only installed/contracted targets. Jun-26 GPUs estimated from 40MW x 538.",
 "D3": "No public list price; IREN's own $2.00/GPU-hr revenue assumption used as price proxy (Estimated). No utilization disclosure.",
 "D4": "Capex/PP&E company-wide (mining+AI). AI-only asset base = 'HPC/GPU hardware' at cost (I31) -> memo M06 and K44.",
 "D5": "HPC hardware 5 yrs (unchanged FY24-FY26); mining 4 yrs (S19j Pro 2 yrs); buildings 20. Yield on company in-service PP&E (net of impairment).",
 "D6": "RPO disclosed from Sep-25 (US GAAP 10-Qs). Microsoft $9.7B contract enters RPO tranche-by-tranche; Jun-26 lease (ASC 842) value $11.4B shown in K21.",
 "D7": "No AI customer % disclosed. Mining revenue via 3 pools.",
 "D8": "Converts-heavy capital structure until FY26 Q4; first secured GPU debt (DDTL, USPP) at Jun-26. Finance expense used for K36 (no quarterly contractual interest).",
}
mapping = [(QS[i], QEND[i], FQ[i].split(" (")[0], FQ[i].split("(")[1].rstrip(")"), "0 (exact)", "FY ends 30-Jun" + ("; opening-balance column" if i == 0 else "")) for i in range(11)]
notes = [
 ("Gap","AI MW disclosed only once (~40MW at Jun-26); operating GPU counts stop after Sep-25 (~2,100). Dec-25 and Mar-26 AI MW/GPUs n/d.","K02, K06-K09 cannot be computed for Q4'25-Q2'26; K01/K43 rely on one MW datapoint.","Earnings-call transcripts (Horizon commissioning MW), Microsoft tranche acceptance dates; IR request."),
 ("Gap","No AI-vs-mining split of capex, D&A, PP&E (except GPU vs mining hardware at cost), or EBITDA.","Capex/MW (K10-K15) and EBITDA/MW (K05) n/d on the AI perimeter.","Use GPU hardware (I31) as AI asset proxy; segment note changes in FY27."),
 ("Gap","No customer revenue concentration for AI Cloud; backlog concentration only derivable (Microsoft ~58%, NVIDIA ~20% of $16.6B).","K24-K27 n/d.",""),
 ("Gap","No public GPU list price; IREN's $2.00/GPU-hr revenue assumption used (Estimated).","K08/K09 not comparable 1:1 with CRWV/Nebius list-based utilization.",""),
 ("Gap","Contractual interest only annual; quarterly finance expense used.","K36 includes issuance-cost amortization / lease interest.",""),
 ("Caveat","IFRS (to Mar-25) vs US GAAP (from Jun-25) basis break; Q2'24 values derived as FY (GAAP recast) - 9M (IFRS).","Small (<$0.1M on AI revenue; D&A/capex may differ more).",""),
 ("Caveat","MW figures from PR boilerplate are 'as of PR date' (5-7 weeks after quarter end); total DC MW is gross basis while AI 40MW basis unstated.","Mining MW estimates (I44E) approximate.",""),
 ("Note","Jun-26 transition: AI 51% of revenue; Bitcoin capacity 23.2 EH/s (~380MW) vs 50 EH/s peak; $16.6B contracted backlog; $4bn contracted ARR / $1bn operating ARR (FY26 PR).","",""),
 ("Note","Post-period: Aug-2026 $2.4B GPU financing at 9.0% fixed (Blue Owl / PIMCO).","Marginal secured cost far above the 1.9% blended rate.",""),
]
CMP = ["K01","K02","K03","K05","K06","K07","K09","K10","K12","K13","K14","K15","K16","K17","K20","K21","K22","K24","K26","K27","K28","K30","K31","K36","K41","K42","K43","K44","K45"]
if __name__ == "__main__":
    build(P, readme, defs_app, notes, mapping, ovr, CMP, debt=debt, memo_kpis=memo)
