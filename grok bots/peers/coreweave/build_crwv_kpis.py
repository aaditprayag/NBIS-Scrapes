# POLISH 2026-10-05 (originally applied in-place on the xlsx) is MERGED INTO THIS BUILDER as of 2026-10-08 ET:
# - Notes_Gaps A/B/C, K02b≡K02 (kpi call after K02), I09B alias (frow after I23A), README/Definitions/Sources/Comparison polish text.
# - FIX 2026-10-08: K18 'Implied average depreciable life' = 100 / K17 (depreciation yield), referenced via KR['K17'].
#   Root cause: the in-place K02b row insert shifted rows by 2 but not K18's relative ref (100/kr-2 -> landed on K16 = 6 yrs -> 16.67).
#   A full rebuild now reproduces the polished pack; no in-place edits needed.
# Builds crwv_kpis.xlsx (CoreWeave peer KPI comps, Q1'24-Q2'26) mirroring Nebius KPI definitions.
# Author: Grok Bot for Aadit Prayag, 2026-10-04 ET. All inputs hard-coded below with tags + sources.
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

OUT = "/workspace/nebius/peers/coreweave/crwv_kpis.xlsx"
PEER = "CoreWeave (CRWV)"
QS = ["Q4'23*", "Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26"]
QEND = ["2023-12-31", "2024-03-31", "2024-06-30", "2024-09-30", "2024-12-31", "2025-03-31", "2025-06-30", "2025-09-30", "2025-12-31", "2026-03-31", "2026-06-30"]
DAYS = [None, 91, 91, 92, 92, 90, 91, 92, 92, 90, 91]
C0 = 5  # column E = Q4'23 opening (stock items only); F..O = Q1'24..Q2'26
def qc(i): return L(C0 + i)  # i = 0..10

# ---------- source registry ----------
SEC = "https://www.sec.gov/Archives/edgar/data/1769628/"
S = {
 "S1":  ("424B4 prospectus (IPO), filed 2025-03-31", SEC + "000119312525067651/d899798d424b4.htm"),
 "Q125": ("10-Q Q1'25, filed 2025-05-15", SEC + "000176962825000014/crwv-20250331.htm"),
 "Q225": ("10-Q Q2'25, filed 2025-08-13", SEC + "000176962825000041/crwv-20250630.htm"),
 "Q325": ("10-Q Q3'25, filed 2025-11-13", SEC + "000176962825000062/crwv-20250930.htm"),
 "K25":  ("10-K FY25, filed 2026-03-02", SEC + "000176962826000104/crwv-20251231.htm"),
 "Q126": ("10-Q Q1'26, filed 2026-05-08", SEC + "000176962826000222/crwv-20260331.htm"),
 "Q226": ("10-Q Q2'26, filed 2026-08-12", SEC + "000176962826000366/crwv-20260630.htm"),
 "PR125": ("Q1'25 earnings PR (8-K Ex.99.1, 2025-05-14)", SEC + "000176962825000010/coreweave1q25earningspress.htm"),
 "PR225": ("Q2'25 earnings PR (8-K Ex.99.1, 2025-08-12)", SEC + "000176962825000039/coreweave2q25earningspress.htm"),
 "PR325": ("Q3'25 earnings PR (8-K Ex.99.1, 2025-11-10)", SEC + "000176962825000059/coreweave3q25earningspress.htm"),
 "PR425": ("Q4'25 earnings PR (8-K Ex.99.1, 2026-02-26)", SEC + "000176962826000094/coreweave4q25earningspress.htm"),
 "PR126": ("Q1'26 earnings PR (8-K Ex.99.1, 2026-05-07)", SEC + "000176962826000220/coreweave1q26earningspress.htm"),
 "PR226": ("Q2'26 earnings PR (8-K Ex.99.1, 2026-08-11)", SEC + "000176962826000362/coreweave2q26earningspress.htm"),
 "DECK0917": ("Convertible-notes investor presentation (8-K Ex., 2026-09-17)", SEC + "000176962826000429/crwv2026convertpresentat.htm"),
 "DECK126": ("Q1'26 earnings presentation (IR site PDF)", "https://s205.q4cdn.com/133937190/files/doc_financials/2026/q1/CoreWeave-1Q26-Earnings-Presentation.pdf"),
 "XBRL": ("SEC companyfacts XBRL, CIK 0001769628 (local copy comps_raw/cf_0001769628.json)", "https://data.sec.gov/api/xbrl/companyfacts/CIK0001769628.json"),
 "DDTL55": ("8-K 2026-08-10 DDTL 5.5 (Item 1.01 + Ex.99.1)", SEC + "000176962826000357/ex991pr.htm"),
 "CV33": ("8-K 2026-09-22 2.875% converts due 2033 (Item 1.01) + pricing PR", SEC + "000176962826000432/crwv-20260917.htm"),
 "NBGPU": ("Nebius workspace: /workspace/nebius/capacity/gpu_mw_capex.xlsx (GPUsPerMW base IT-MW rows) / gpu_mw_capex.md s.1", "local"),
 "NBPX": ("Nebius workspace: /workspace/nebius/parts/price_history_quarterly.csv (CoreWeave on-demand list, GPU Market agent; Wayback snapshots)", "local"),
}
SRC_ROWS = []  # (input id, item, quarter, value, tag, source key, section, note)

# ---------- inputs ----------
# Each input: id, item, unit, values list(11) (None = n/d), tags list(11) or single tag, source key per quarter (list or single), section, notes
nd = None
INPUTS = []
def add(iid, item, unit, vals, tags, srcs, sect, note=""):
    if isinstance(tags, str): tags = [tags]*11
    if isinstance(srcs, str): srcs = [srcs]*11
    if isinstance(sect, str): sect = [sect]*11
    INPUTS.append(dict(id=iid, item=item, unit=unit, vals=vals, tags=tags, srcs=srcs, sect=sect, note=note))

PRQ = ["S1","S1","S1","S1","S1","PR125","PR225","PR325","PR425","PR126","PR226"]
TQ  = ["S1","Q125","Q225","Q325","S1","Q125","Q225","Q325","K25","Q126","Q226"]

add("I01","Revenue","$M",[nd,188.684,395.371,583.941,747.430,981.632,1212.788,1364.676,1572,2078,2575],
    [None,"R","R","R","R","R","R","R","R","R","R"], PRQ,
    ["","S-1 Quarterly Results of Operations p.117","p.117","p.117","p.117","PR financial highlights","PR highlights","PR highlights","PR highlights ($M)","PR highlights ($M)","PR highlights ($M)"])
add("I02","Adjusted EBITDA (company non-GAAP)","$M",[nd,104.546,249.841,378.757,486.114,606.131,753.169,838.124,898,1157,1510],
    [None]+["R"]*10, PRQ, ["","S-1 Quarterly Adj. EBITDA reconciliation p.120"]*1+["p.120"]*3+["PR non-GAAP table"]*6)
add("I03","Depreciation & amortization (total, cash-flow basis)","$M",[nd,79.510,164.460,254.024,365.419,443.5,559.5,630.5,820.5,1147,1393],
    [None,"R","R","R","R","R","D","D","D","R","D"],
    ["","S1","S1","S1","S1","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL"],
    ["","S-1 Adj. EBITDA reconciliation p.120","p.120","p.120","p.120","Q1'25 10-Q CF stmt (XBRL DepreciationDepletionAndAmortization)",
     "= H1'25 1,003.0 - Q1 443.5 (YTD diff)","= 9M 1,633.5 - H1 1,003.0","= FY 2,454 - 9M 1,633.5","Q1'26 10-Q CF stmt","= H1'26 2,540 - Q1 1,147"],
    "Nebius depreciation-yield uses PP&E depreciation; CRWV PP&E depreciation is disclosed rounded to $0.1B in 2026, so total D&A (incl. small intangible amortization) is used.")
add("I04","Cash capex (purchases of property & equipment)","$M",[nd,1741.9,2247.2,1215.2,3497.7,1407.4,2453.0,2388.8,4059.8,7695,6422],
    [None,"R","D","D","D","R","D","D","D","R","D"], "XBRL",
    ["","Q1'25 10-Q comparative (YTD Q1'24)","= H1'24 3,989.1 - Q1 1,741.9","= 9M'24 5,204.3 - H1 3,989.1","= FY24 8,702 - 9M 5,204.3","Q1'25 10-Q CF",
     "= H1 3,860.4 - Q1 1,407.4","= 9M 6,249.2 - H1 3,860.4","= FY 10,309 - 9M 6,249.2","Q1'26 10-Q CF","= H1'26 14,117 - Q1 7,695"],
    "Mirrors Nebius Hist 'Capex (purchases of PP&E + intangibles, cash)'. CRWV also finances GPUs via OEM/vendor financing, so cash capex understates economic capex in some quarters.")
add("I05","Gross property & equipment (EoP)","$M",[3600.786,nd,nd,nd,12880.005,15610.549,18585.117,23233.759,33941,40933,52622],
    ["R",None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 5 Property and Equipment, Net")
add("I06","Construction in progress (EoP; not yet in service)","$M",[2256.673,nd,nd,nd,3200.866,4073.810,4111.256,6910.066,9376,9581,11918],
    ["R",None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 5")
add("I07","Property & equipment, net (EoP)","$M",[3483.990,nd,nd,nd,11914.774,14210.992,16631.510,20659.181,30557,36424,46736],
    ["R",None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 5")
add("I08","Technology equipment, gross (EoP)","$M",[1298.127,nd,nd,nd,9146.575,10746.445,13170.295,14628.790,20903,26627,33823],
    ["R",None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 5")
add("I09","Active power, reported (EoP)","MW",[70,nd,nd,nd,360,420,470,590,850,1000,1500],
    ["R",None,None,None,"R","R","R","R","R","R","R"],
    ["S1","","","","S1","PR125","PR225","PR325","PR425","PR126","PR226"],
    ["S-1 Business ('more than 70 MW'; 10-K FY25 says 'approximately 70 MW')","","","","S-1 ('more than 360 MW'); 10-K FY25 'approximately 360 MW'","'approximately 420 MW of active power'","'approximately 470 MW of active power'","'bringing the total to approximately 590 MW'","'more than 850 MW' (lower bound)","'Surpassed 1 GW' (lower bound; Q2 PR '+nearly 500 MW to reach 1.5 GW' implies ~1.0 GW)","'Expanded active power by nearly 500 MWs to reach 1.5 GW'"],
    "MW basis (IT vs gross) never stated by CRWV. YE24 cross-check: >250k GPUs / ~360 MW = ~694 GPUs/MW, close to Nebius Hopper base 713/IT-MW (vs 548/gross-MW) -> CRWV 'active power' most consistent with IT (critical) load. Several values are 'more than'/'approximately' -> treat as lower bounds/rounded. | POLISH: active power ≈ billed/revenue-generating for K02b≡K02 (no separate billing MW).")
add("I10","Contracted power, reported (EoP)","MW",[nd,nd,nd,nd,1300,1600,2200,2900,3100,3500,3700],
    [None,None,None,None,"R","R","R","R","R","R","R"],
    ["","","","","S1","PR125","PR225","PR325","PR425","PR126","PR226"],
    ["","","","","S-1 'approximately 1.3 GW'","~1.6 GW","2.2 GW","~2.9 GW","~3.1 GW","'over 3.5 GW'","~3.7 GW (deck: ~4.2 GW as of Aug-11-2026, post-quarter)"])
add("I11","GPUs deployed, reported (EoP)","GPUs",[53000,nd,nd,nd,250000,nd,nd,nd,nd,nd,nd],
    ["R",None,None,None,"R",None,None,None,None,None,None], ["S1","","","","S1","","","","","",""],
    ["S-1 'more than 53,000 GPUs' (lower bound)","","","","S-1 'more than 250,000 GPUs' across 32 DCs (lower bound)"]+[""]*6,
    "CRWV has not disclosed a fleet GPU count since the S-1 (checked 10-Qs Q1'25-Q2'26, 10-K FY25, all earnings PRs, Sep-26 convert deck).")
add("I12","Hopper-class on-demand list price (H100 SXM 8-GPU instance / 8)","$/GPU-hr",[nd,6.155,6.155,6.155,6.155,6.155,6.155,6.155,6.155,6.155,6.155],
    [None,"E","E","D","D","D","D","D","D","D","D"], "NBPX",
    ["","carried back from first observation Q3'24 (no instance-price snapshot)","carried back","$49.24/8 (Wayback snapshot)"]+["$49.24/8 per GPU Market csv"]*7,
    "H200 list $6.305 is ~2% higher; A100 ($2.70) and older GPUs are cheaper, so using H100 for all pre-Blackwell GPUs overstates fleet list price.")
add("I13","Blackwell on-demand list price (GB200 NVL72 4-GPU instance / 4)","$/GPU-hr",[nd,nd,nd,nd,nd,10.5,10.5,10.5,10.5,10.5,10.5],
    [None,None,None,None,None,"D","D","D","D","D","D"], "NBPX",
    ["","","","","","$42.00/4 (first listed 2025)"]+["$42.00/4"]*5,
    "B200 lists at $8.60 (2026); GB300 is contact-sales (no public list). GB200 used for all Blackwell-era additions.")
add("I14","RPO (ASC 606 remaining performance obligations, EoP)","$B",[9.9,nd,nd,nd,15.1,14.7,30.1,50.0,60.7,98.8,103.7],
    ["R",None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 2 Revenue - RPO paragraph (also XBRL RevenueRemainingPerformanceObligation)")
add("I15","Revenue backlog (company metric: RPO + other committed amounts)","$B",[9.9,nd,nd,nd,15.1,25.9,30.1,55.6,66.8,99.4,104.2],
    ["R",None,None,None,"R","R","R","R","R","R","R"],
    ["DECK0917","","","","DECK0917","PR125","PR225","PR325","PR425","PR126","PR226"],
    ["Sep-26 deck slide 27 note 5 ($9.9B YE23)","","","","deck slide 27 ($15.1B YE24)","PR ($14.7B RPO + $11.2B other)","PR","PR","PR","PR","PR '~$104B' / deck $104.2B; excludes >$25B signed early Q3"])
add("I16","RPO % expected to be recognized within 24 months","%",[nd,nd,nd,nd,54,58,50,42,43,36,41],
    [None,None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 2 Revenue - RPO paragraph")
add("I17","Largest customer, % of quarter revenue","%",[nd,33,59,65,68.2,72,71,67,60.2,45,36],
    [None,"R","R","R","D","R","R","R","D","R","R"],
    ["","Q125","Q225","Q325","S1","Q125","Q225","Q325","K25","Q126","Q226"],
    ["","Q1'25 10-Q Note 2 (Q1'24 comparative, Customer A)","Q2'25 10-Q Note 2 (Q2'24 comparative)","Q3'25 10-Q Note 2 (Q3'24 comparative)",
     "= (FY24 62% x 1,915.0 - 9M'24 58% x 1,168.0) / Q4 747.4; +/-~2pp from rounding of disclosed %",
     "Q1'25 10-Q Note 2","Q2'25 10-Q Note 2","Q3'25 10-Q Note 2","= (FY25 67% x 5,131 - 9M'25 70% x 3,559.1) / Q4 1,571.9; +/-~2pp","Q1'26 10-Q Note 2","Q2'26 10-Q Note 2"],
    "Customer letters are per-filing labels; Customer A is historically Microsoft (S-1/10-K name Microsoft as FY23-FY25 top customer). Q1'24: A 33% vs C 32%.")
add("I18","2nd-largest customer, % of quarter revenue (<10% = not disclosed)","%",[nd,32,20,12,10.3,nd,nd,nd,nd,20,26],
    [None,"R","R","R","D",None,None,None,None,"R","R"],
    ["","Q125","Q225","Q325","S1","Q125","Q225","Q325","K25","Q126","Q226"],
    ["","Customer C 32%","Customer B 20%","Customer B 12%","= (FY24 15% x 1,915 - 9M 18% x 1,168)/747.4; +/-2pp -> ~10% (ambiguous)","'*' (<10%)","<10%","<10%","<10% (no other customer >=10% of FY25)","Customer B 20%","Customer B 26%"])
add("I19","3rd-largest customer, % of quarter revenue (<10% = not disclosed)","%",[nd,18,nd,nd,nd,nd,nd,nd,nd,nd,10],
    [None,"R"]+[None]*8+["R"], ["","Q125"]+[""]*8+["Q226"], ["","Customer B 18%"]+[""]*8+["Customer C 10%"])
add("I20","Customers >=10% of quarter revenue (count)","#",[nd,3,2,2,2,1,1,1,1,2,3],
    [None,"R","R","R","D","R","R","R","D","R","R"],
    ["","Q125","Q225","Q325","S1","Q125","Q225","Q325","K25","Q126","Q226"],
    ["","Note 2","Note 2","Note 2","derived (B ~10% borderline)","Note 2","Note 2","Note 2","FY25 10-K: only Microsoft >=10% for FY","Note 2","Note 2"])
add("I21","Contractual interest expense (before capitalization)","$M",[nd,64.622,100.237,130.003,179.982,232.772,250.312,324.131,412.785,483,592],
    [None,"R","R","R","D","R","R","R","D","R","R"],
    ["","Q125","Q225","Q325","S1","Q125","Q225","Q325","K25","Q126","Q226"],
    ["","Note 10 interest table (2024 comparative)","Note 10","Note 10","= FY24 474.844 - 9M 294.862","Note 10","Note 10","Note 10","= FY25 1,220 - 9M 807.215","Note 10","Note 10"])
add("I22","Capitalized interest","$M",[nd,32.177,40.687,35.072,51.081,13.092,23.186,61.764,83.958,97,79],
    [None,"R","R","R","D","R","R","R","D","R","R"],
    ["","Q125","Q225","Q325","S1","Q125","Q225","Q325","K25","Q126","Q226"],
    ["","Note 10","Note 10","Note 10","= FY24 159.017 - 9M 107.936","Note 10","Note 10","Note 10","= FY25 182 - 9M 98.042","Note 10","Note 10"])
add("I23","Total debt principal (EoP; excl. leases)","$M",[1555.049,nd,nd,nd,8033.477,8807.921,11171.485,14205.268,21615,25149,35551],
    ["R",None,None,None,"R","R","R","R","R","R","R"], TQ, "Note 10 Debt table / maturity table")
add("I24","Company-reported weighted average interest rate on debt","%",[14.9,nd,nd,nd,12.2,nd,nd,nd,9.0,nd,8.3],
    ["R",None,None,None,"R",None,None,None,"R",None,"R"], ["DECK0917","","","","DECK0917","","","","DECK0917","","DECK0917"],
    ["deck slide 27 'FY-23' (uses FY-24 eff. rates on FY-23 principal)","","","","slide 27 FY-24","","","","slide 27 FY-25","","slide 27 Q2-26 (PF incl. $3.5B convert @2.625%: 7.8%)"],
    "Company method: effective rates x EoP principal, OEM/software financing assumed 10%.")
add("I25","Company-definition capex (additions to PP&E + finance-lease assets - change in CIP), reported","$M",[nd,nd,nd,nd,nd,1858,nd,nd,nd,6786,9400],
    [None]*5+["R",None,None,None,"R","R"], [""]*5+["DECK126","","","","DECK126","DECK0917"],
    [""]*5+["Q1'26 deck capex reconciliation (Q1'25 comparative)","","","","Q1'26 deck: $6,991 gross PP&E increase - $205 CIP change","Sep-26 deck slide 26 '$9.4B'"])
add("I26","Days in quarter","days",[nd,91,91,92,92,90,91,92,92,90,91],[None]+["R"]*10,"","calendar")
add("I27","Active data centers (EoP)","#",[10,nd,nd,nd,32,nd,nd,nd,43,49,51],["R",None,None,None,"R",None,None,None,"R","R","R"],
    ["S1","","","","S1","","","","K25","DECK126","DECK0917"],"Business / deck 'Active Data Centers'")

# Estimation parameters (Nebius gpu_mw_capex.xlsx base IT-MW densities)
PARAMS = [
 ("P01","GB200 NVL72 GPUs per IT-MW (base)",538,"E","NBGPU","GPUsPerMW base IT-MW row (gpu_mw_capex.md s.1: 487/538/571)"),
 ("P02","GB300 NVL72 GPUs per IT-MW (base)",491,"E","NBGPU","GPUsPerMW base IT-MW row (453/491/519)"),
 ("P03","Hopper HGX GPUs per IT-MW (base; cross-check only)",713,"E","NBGPU","GPUsPerMW base IT-MW row (665/713/813)"),
]

# ---------- debt facility table ----------
SNAP = ["Dec-23","Dec-24","Mar-25","Jun-25","Sep-25","Dec-25","Mar-26","Jun-26"]
SNAPQ = [0,4,5,6,7,8,9,10]  # index in QS
SNAPSRC = ["S1","S1","Q125","Q225","Q325","K25","Q126","Q226"]
# facility, category, (principal $M, eff rate %) per snapshot (None = not outstanding)
FAC = [
 ("2021 Convertible Senior Secured Notes","Converts",[(55.125,17)]+[None]*7),
 ("2022 Senior Secured Notes","Senior notes",[(125.0,10)]+[None]*7),
 ("DDTL 1.0 Facility","DDTL",[(1374.924,15),(2012.5,15),(1897.5,15),(1782.5,15),(1667.5,15),(1553,15),(1438,15),(1300,15)]),
 ("DDTL 2.0 Facility","DDTL",[None,(3843.819,11),(4386.303,11),(4953.967,11),(5037.372,10),(5037,10),(4425,11),(3190,11)]),
 ("DDTL 2.1 Facility","DDTL",[None]*4+[(1241.325,9),(2741,9),(3000,9),(3000,9)]),
 ("DDTL 3.0 Facility","DDTL",[None]*4+[(358.696,9),(340,9),(1700,9),(2215,9)]),
 ("DDTL 4.0 Facility (non-recourse; IG)","DDTL",[None]*6+[(1260,7),(2837,7)]),
 ("DDTL 5.0 Facility","DDTL",[None]*7+[(1101,9)]),
 ("2024 Term Loan Facility","Other term loans",[None,(1000,12),(1000,12)]+[None]*5),
 ("Magnetar Loan","Other term loans",[None,None,(248.917,12),(256.630,12),(264.666,12),(273,12),(281,12),(189,12)]),
 ("2030 Senior Notes (9.25%)","Senior notes",[None,None,None,(2000,10),(2000,10),(2000,10),(2000,10),(2000,10)]),
 ("2031 9.00% Senior Notes","Senior notes",[None]*4+[(1750,10),(1750,10),(1750,10),(1750,10)]),
 ("2031 9.75% Senior Notes","Senior notes",[None]*7+[(2750,10)]),
 ("2032 9.625% Senior Notes","Senior notes",[None]*7+[(1250,10)]),
 ("2032 EUR Senior Notes (8.50%)","Senior notes",[None]*7+[(2279,9)]),
 ("2031 Convertible Senior Notes (1.75%)","Converts",[None]*5+[(2588,2),(2588,2),(2588,2)]),
 ("2032 Convertible Senior Notes (1.75%)","Converts",[None]*7+[(4000,2)]),
 ("Convertible Promissory Notes","Converts",[None]*5+[(168,7),(171,7),None]),
 ("Revolving Credit Facility","Revolver",[None,None,(50,6),(450,7),(700,7),(1000,6),(1500,6),(0,7)]),
 ("OEM & software-license financing (recourse)","OEM / vendor financing",[None,(1177.158,10),(1225.201,9.5),(1728.388,9.5),(1185.709,9),(4165,10),(5036,10),(4220,11)]),
 ("OEM & software-license financing (non-recourse)","OEM / vendor financing",[None]*7+[(882,9)]),
]
CATS = ["DDTL","Senior notes","Converts","OEM / vendor financing","Revolver","Other term loans"]
RATE_NOTE = "Effective rates are as printed in each filing's Note 10 debt table (rounded to whole %). Ranges are entered at midpoint: OEM 'Mar-25/Jun-25 9-10%' -> 9.5; 'Sep-25 8-10%' -> 9; S-1 Dec-24 OEM '9%-11%' -> 10. Dec-25 OEM recourse/non-recourse combined (10-K shows a single 10% line of $4,165M)."

# ---------- workbook ----------
wb = openpyxl.Workbook()
BLUE = Font(color="0000CC"); ORANGE = Font(color="C65911"); GREY = Font(color="808080", italic=True); BLACK = Font(color="000000")
BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="D9E1F2"); SEC_F = PatternFill("solid", fgColor="F2F2F2")
TAGFONT = {"R": BLUE, "D": BLACK, "E": ORANGE, None: GREY}
TAGNAME = {"R": "Reported", "D": "Derived", "E": "Estimated", None: "n/d"}
wrap = Alignment(wrap_text=True, vertical="top")

def header(ws, row, cols, widths=None):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=h); c.font = BOLD; c.fill = HDR; c.alignment = wrap
    if widths:
        for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w

# README
ws = wb.active; ws.title = "README"
readme = [
 "Peer KPI comps - CoreWeave (CRWV) - quarterly Q1'24 to Q2'26. Built 2026-10-04 (ET); polished 2026-10-05 (ET) — Notes_Gaps reorganized, K02b≡K02 alignment, EDGAR scan (no new ops data). File: /workspace/nebius/peers/coreweave/crwv_kpis.xlsx; builder: build_crwv_kpis.py.",
 "PURPOSE: CoreWeave KPI history on the SAME definitions as the Nebius work (see Definitions tab: every KPI cites the Nebius file/row it mirrors).",
 "TABS: KPIs (quarterly, live formulas off Inputs/Debt_Facilities; each KPI = value row + tag row) | Inputs (raw figures, colour-coded by tag) | Debt_Facilities (facility-level principal x effective rate per quarter-end) | Sources (one row per input cell: document, section, URL, formula for derived cells) | Definitions | Notes_Gaps | Comparison (template vs Nebius; CRWV linked live).",
 "TAGS / COLOURS: blue = Reported (as disclosed); black = Derived (arithmetic on reported figures; formula in Sources/KPIs); orange = Estimated (uses an assumption, e.g. interpolated MW, GPUs from MW x density); grey 'n/d' = not disclosed. KPI cells inherit the WEAKEST tag of their inputs (shown on the tag row under each KPI).",
 "COLUMN E (Q4'23*) is an opening-balance column for stock items (MW, PP&E, debt, GPUs) used for averages/deltas in Q1'24; flow KPIs start in column F (Q1'24).",
 "TEMPLATE USE (IREN, Oracle OCI, private peers): copy this workbook, keep every row ID (I01..I27, P01..P03, K01..K40) and column layout, overwrite Inputs values/tags/sources and the Debt_Facilities rows for the new peer. KPI formulas reference fixed Inputs rows, so they recompute unchanged. Leave unavailable inputs blank or 'n/d' - KPIs return 'n/d'. Then link the peer's latest-quarter KPI column into the Comparison tab.",
 "UNITS: $M unless stated; $/MW KPIs are annualized ($M per MW-year = quarterly revenue x 4 / MW); $/GPU in $K per GPU-year; rates in % p.a.",
 "Caveat summary: CoreWeave does not state IT vs gross MW, has not disclosed GPU counts since the S-1, and did not publish 2024 interim balance sheets or MW (pre-IPO). See Notes_Gaps.",
 "POLISH 2026-10-05: (1) Notes_Gaps rebuilt as A=n/d inventory / B=caveats / C=post-quarter+cross-pack. (2) K02b added (=K02); I09B alias = I09E. (3) EDGAR through Oct 5: latest ops still Q2'26 10-Q; Oct 2–3 = Form 4 only. No quarterly number changes vs Oct-4 pack.",
 "FIX 2026-10-08 (ET): KPI K18 'Implied average depreciable life' now = 100 / K17 (depreciation yield) in every quarter (Q2'26 ~6.47 yrs). Previously 100 / K16 (6-yr useful-life row) = 16.67 for all quarters after the in-place K02b row insert; polish merged into build_crwv_kpis.py.",
]
for i, t in enumerate(readme, 1):
    c = ws.cell(row=i, column=1, value=t); c.alignment = wrap
ws.column_dimensions["A"].width = 160

# Inputs
wsI = wb.create_sheet("Inputs")
wsI.cell(row=1, column=1, value=f"Inputs - {PEER}. Raw reported figures (blue), derived (black), estimated (orange), n/d (grey). Per-cell source detail on Sources tab.").font = BOLD
header(wsI, 3, ["ID", "Item", "Unit", "Tag mix"] + QS + ["Source summary / notes"], [7, 52, 10, 12] + [10]*11 + [90])
ROW = {}
r = 4
for inp in INPUTS:
    ROW[inp["id"]] = r
    wsI.cell(row=r, column=1, value=inp["id"]); wsI.cell(row=r, column=2, value=inp["item"]); wsI.cell(row=r, column=3, value=inp["unit"])
    tagset = sorted(set(TAGNAME[t] for t, v in zip(inp["tags"], inp["vals"]) if v is not None))
    wsI.cell(row=r, column=4, value="/".join(tagset))
    for i, (v, t) in enumerate(zip(inp["vals"], inp["tags"])):
        c = wsI.cell(row=r, column=C0 + i)
        if v is None:
            c.value = "n/d"; c.font = GREY
        else:
            c.value = v; c.font = TAGFONT[t]
            sk = inp["srcs"][i]
            doc, url = S.get(sk, ("", ""))
            SRC_ROWS.append((inp["id"], inp["item"], QS[i], v, TAGNAME[t], doc, inp["sect"][i], url))
            if t in ("D", "E") and inp["sect"][i]:
                c.comment = Comment(inp["sect"][i], "GrokBot")
    srcset = []
    for sk in inp["srcs"]:
        if sk and sk in S and S[sk][0] not in srcset: srcset.append(S[sk][0])
    wsI.cell(row=r, column=C0 + 11, value=("; ".join(srcset) + (" | NOTE: " + inp["note"] if inp["note"] else ""))).alignment = wrap
    r += 1
# derived / estimated input rows (formulas)
r += 1
wsI.cell(row=r, column=1, value="Derived & estimated input rows (formulas)").font = BOLD; r += 1
def frow(iid, item, unit, tagmix, fn, note, tags):
    global r
    ROW[iid] = r
    wsI.cell(row=r, column=1, value=iid); wsI.cell(row=r, column=2, value=item); wsI.cell(row=r, column=3, value=unit); wsI.cell(row=r, column=4, value=tagmix)
    for i in range(11):
        f = fn(i)
        c = wsI.cell(row=r, column=C0 + i, value=f if f is not None else "n/d")
        c.font = TAGFONT.get(tags[i], GREY) if f is not None else GREY
    wsI.cell(row=r, column=C0 + 11, value=note).alignment = wrap
    r += 1
def ref(iid, i): return f"{qc(i)}{ROW[iid]}"
# params
PROW = {}
for pid, name, val, tag, sk, sect in PARAMS:
    PROW[pid] = r
    wsI.cell(row=r, column=1, value=pid); wsI.cell(row=r, column=2, value=name); wsI.cell(row=r, column=3, value="GPUs/MW"); wsI.cell(row=r, column=4, value="Estimated")
    c = wsI.cell(row=r, column=C0, value=val); c.font = ORANGE
    wsI.cell(row=r, column=C0 + 11, value=S[sk][0] + " - " + sect).alignment = wrap
    SRC_ROWS.append((pid, name, "param", val, "Estimated", S[sk][0], sect, S[sk][1]))
    r += 1
P = lambda pid: f"$E${PROW[pid]}"
# I09E active MW incl. estimates (linear interpolation YE23->YE24 for Q1-Q3'24)
frow("I09E", "Active power incl. estimates (EoP) - 2024 interim interpolated", "MW", "Reported/Estimated",
     lambda i: (f"={ref('I09',i)}" if i in (0,4,5,6,7,8,9,10) else f"=$E${ROW['I09']}+($I${ROW['I09']}-$E${ROW['I09']})*{i}/4"),
     "ESTIMATED for Q1-Q3'24: straight-line between ~70 MW (YE23) and ~360 MW (YE24). CRWV did not disclose interim 2024 MW. Other quarters = reported.",
     ["R","E","E","E","R","R","R","R","R","R","R"])
frow("I09A", "Average active power in quarter ((open+close)/2)", "MW", "Derived/Estimated",
     lambda i: None if i == 0 else f"=({ref('I09E',i-1)}+{ref('I09E',i)})/2",
     "Average of opening/closing active MW (I09E). Q1'24-Q4'24 use the interpolated estimates -> Estimated.",
     [None,"E","E","E","E","D","D","D","D","D","D"])
frow("I09N", "New active MW added in quarter (reported MW only)", "MW", "Derived",
     lambda i: None if i < 5 else f"={ref('I09',i)}-{ref('I09',i-1)}",
     "= closing - opening reported active MW. Q4'24 n/d (Q3'24 MW not disclosed). Lower-bound wording ('more than') makes single-quarter deltas noisy (+/-10-50 MW).",
     [None]*5+["D"]*6)
frow("I09NE", "New active MW added in quarter (incl. estimated 2024 MW)", "MW", "Derived/Estimated",
     lambda i: None if i == 0 else f"={ref('I09E',i)}-{ref('I09E',i-1)}",
     "Uses I09E; 2024 quarters each +72.5 MW by construction (Estimated).",
     [None,"E","E","E","E","D","D","D","D","D","D"])
# GPUs estimated
def gpu_f(i):
    if i == 0: return f"={ref('I11',0)}"
    if i in (1,2,3): return f"=$E${ROW['I11']}+($I${ROW['I11']}-$E${ROW['I11']})*{i}/4"
    if i == 4: return f"={ref('I11',4)}"
    if i in (5,6,7,8): return f"=$I${ROW['I11']}+({ref('I09',i)}-$I${ROW['I09']})*{P('P01')}"
    return f"=$M${r}+({ref('I09',i)}-$M${ROW['I09']})*{P('P02')}"
frow("I11E", "GPUs incl. estimates (EoP)", "GPUs", "Reported/Estimated", gpu_f,
     "ESTIMATED except YE23/YE24 (reported lower bounds). 2024 interim: linear 53k->250k. 2025: 250k + (active MW - 360) x GB200 base 538 GPUs/IT-MW. 2026: YE25 est + (MW - 850) x GB300 base 491. Mirrors Nebius capacity method 'GPUs = MW x GPUs/MW (gpu_mw_capex.xlsx base rows)'. Assumes no retirements and all post-YE24 MW is Blackwell.",
     ["R","E","E","E","R","E","E","E","E","E","E"])
frow("I11A", "Average GPUs in quarter ((open+close)/2)", "GPUs", "Estimated",
     lambda i: None if i == 0 else f"=({ref('I11E',i-1)}+{ref('I11E',i)})/2", "Average of I11E.", [None]+["E"]*10)
frow("I12W", "Fleet-weighted on-demand list price (EoP GPU weights)", "$/GPU-hr", "Estimated",
     lambda i: None if i == 0 else (f"={ref('I12',i)}" if i <= 4 else f"=({ref('I12',i)}*$I${ROW['I11']}+{ref('I13',i)}*({ref('I11E',i)}-$I${ROW['I11']}))/{ref('I11E',i)}"),
     "Mirrors Nebius KPIs row 'Fleet-weighted on-demand list price' (GPU_End_H x P_OD_H + GPU_End_B x P_OD_B ...)/GPUs. Legacy 250k GPUs (YE24) priced at Hopper H100 list; all later GPUs at GB200 list.",
     [None]+["E"]*10)
frow("I05S", "In-service gross PP&E (gross - CIP, EoP)", "$M", "Derived",
     lambda i: None if i in (1,2,3) else f"={ref('I05',i)}-{ref('I06',i)}",
     "Mirrors Nebius Capex_DA 'In-service gross PP&E' = gross PP&E - assets not yet in use.", ["D",None,None,None]+["D"]*7)
frow("I25D", "Company-definition capex, derived = change in in-service gross PP&E", "$M", "Derived",
     lambda i: None if i < 5 else f"={ref('I05S',i)}-{ref('I05S',i-1)}",
     "= delta(gross PP&E) - delta(CIP). Matches CRWV's own capex definition (Q1'26 deck: 'Increase in total gross P&E less change in CIP'): Q1'25 1,858 / Q1'26 6,786 / Q2'26 ~9.4B reported. Also = Nebius 'Transfer into service'.",
     [None]*5+["D"]*6)
frow("I23A", "Average debt principal ((open+close)/2)", "$M", "Derived",
     lambda i: None if i < 5 else f"=({ref('I23',i-1)}+{ref('I23',i)})/2", "Q1'24-Q4'24 n/d (no 2024 interim debt tables).", [None]*5+["D"]*6)
frow("I09B", "Billed / billing MW alias — EQUALS I09E (documentation only; CRWV has no separate billed series)", "MW", "Align",
     lambda i: f"={ref('I09E',i)}", "POLISH 2026-10-05: alias for cross-pack K02b alignment. Do not treat as a new estimate. Notes_Gaps B1.",
     ["R","E","E","E","R","R","R","R","R","R","R"])

# Debt_Facilities
wsD = wb.create_sheet("Debt_Facilities")
wsD.cell(row=1, column=1, value=f"Debt by facility - {PEER}. Principal ($M) and effective interest rate (%) as printed in each quarter-end debt table (Note 10 / S-1 Note 10). Excludes operating/finance leases.").font = BOLD
wsD.cell(row=2, column=1, value=RATE_NOTE).alignment = wrap
cols = ["Facility", "Category"]
for s in SNAP: cols += [f"{s} principal", f"{s} eff. rate %"]
header(wsD, 4, cols, [44, 22] + [11]*16)
dr = 5
for name, cat, vals in FAC:
    wsD.cell(row=dr, column=1, value=name); wsD.cell(row=dr, column=2, value=cat)
    for k, v in enumerate(vals):
        if v is not None:
            a = wsD.cell(row=dr, column=3 + 2*k, value=v[0]); a.font = BLUE
            b = wsD.cell(row=dr, column=4 + 2*k, value=v[1]); b.font = BLUE
            SRC_ROWS.append(("DEBT", name + " (principal / eff. rate)", SNAP[k], f"{v[0]} / {v[1]}%", "Reported", S[SNAPSRC[k]][0], "Note 10 Debt table", S[SNAPSRC[k]][1]))
    dr += 1
last = dr - 1
dr += 1
DROW = {}
wsD.cell(row=dr, column=1, value="Total principal").font = BOLD
for k in range(8):
    pc = L(3 + 2*k); wsD.cell(row=dr, column=3 + 2*k, value=f"=SUM({pc}5:{pc}{last})")
DROW["tot"] = dr; dr += 1
wsD.cell(row=dr, column=1, value="Check vs reported total principal (Inputs I23)")
for k in range(8):
    wsD.cell(row=dr, column=3 + 2*k, value=f"={L(3+2*k)}{DROW['tot']}-Inputs!{qc(SNAPQ[k])}{ROW['I23']}")
DROW["chk"] = dr; dr += 1
wsD.cell(row=dr, column=1, value="WEIGHTED EFFECTIVE COST OF DEBT (principal-weighted, all facilities)").font = BOLD
for k in range(8):
    pc, rc = L(3 + 2*k), L(4 + 2*k)
    wsD.cell(row=dr, column=4 + 2*k, value=f"=IFERROR(SUMPRODUCT({pc}5:{pc}{last},{rc}5:{rc}{last})/SUM({pc}5:{pc}{last}),\"n/d\")")
DROW["wavg"] = dr; dr += 1
for cat in CATS:
    wsD.cell(row=dr, column=1, value=f"  {cat}: principal / weighted eff. rate")
    for k in range(8):
        pc, rc = L(3 + 2*k), L(4 + 2*k)
        wsD.cell(row=dr, column=3 + 2*k, value=f"=SUMIF($B$5:$B${last},\"{cat}\",{pc}5:{pc}{last})")
        wsD.cell(row=dr, column=4 + 2*k, value=f"=IFERROR(SUMPRODUCT(($B$5:$B${last}=\"{cat}\")*{pc}5:{pc}{last}*{rc}5:{rc}{last})/{pc}{dr},\"n/d\")")
    DROW[cat] = dr; dr += 1
wsD.cell(row=dr, column=1, value="  Share of principal: secured asset-level (DDTL) %")
for k in range(8):
    wsD.cell(row=dr, column=4 + 2*k, value=f"=IFERROR({L(3+2*k)}{DROW['DDTL']}/{L(3+2*k)}{DROW['tot']}*100,\"n/d\")")
DROW["ddtlsh"] = dr; dr += 2
post = [
 "POST-QUARTER (not in Q2'26 figures): Aug-10-2026 DDTL 5.5 $2.6B delayed-draw at SOFR+5.50% (Ba2/BB+), undrawn fee 0.50% [8-K 2026-08-10]. Sep-22-2026 2.875% convertible senior notes due 2033, $4.2B incl. $0.5B option [8-K 2026-09-22]. Company Q2'26 PF weighted rate 7.8% (incl. $3.5B convert at 2.625% illustrative, excl. DDTL 5.5) [Sep-26 deck slide 27].",
 "Q1'26 PR: DDTL 4.0 $8.5B non-recourse IG facility, floating tranche SOFR+2.25%, fixed tranche ~5.9% (as of initial draw).",
]
for t in post:
    wsD.cell(row=dr, column=1, value=t).alignment = wrap; dr += 1

def drefq(key, i, rate=True):
    """reference to Debt_Facilities for quarter index i (returns None if no snapshot)"""
    if i not in SNAPQ: return None
    k = SNAPQ.index(i)
    return f"Debt_Facilities!{L(4+2*k) if rate else L(3+2*k)}{DROW[key]}"

# KPIs
wsK = wb.create_sheet("KPIs", 1)
wsK.cell(row=1, column=1, value=f"KPIs - {PEER} - quarterly, Nebius definitions (Definitions tab). Value row = live formula; row below = tag (R/D/E/n/d = weakest input). Annualized where $/MW or $/GPU.").font = BOLD
header(wsK, 3, ["ID", "KPI", "Unit", "Def"] + QS + ["Method / caveat"], [7, 56, 12, 6] + [10]*11 + [80])
for i, q in enumerate(QEND): wsK.cell(row=2, column=C0 + i, value=q).font = GREY
KR = {}
kr = 4
def section(title):
    global kr
    c = wsK.cell(row=kr, column=1, value=title); c.font = BOLD
    for j in range(1, 17): wsK.cell(row=kr, column=j).fill = SEC_F
    kr += 1
def kpi(kid, name, unit, did, fn, tags, note, fmt="0.00"):
    global kr
    KR[kid] = kr
    wsK.cell(row=kr, column=1, value=kid); wsK.cell(row=kr, column=2, value=name); wsK.cell(row=kr, column=3, value=unit); wsK.cell(row=kr, column=4, value=did)
    for i in range(11):
        f = fn(i)
        t = tags[i] if isinstance(tags, list) else tags
        c = wsK.cell(row=kr, column=C0 + i, value=(f if f is not None else "n/d"))
        c.number_format = fmt
        c.font = TAGFONT.get(t, GREY) if f is not None else GREY
        tc = wsK.cell(row=kr + 1, column=C0 + i, value=(TAGNAME.get(t, "n/d") if f is not None else "n/d")); tc.font = GREY
    wsK.cell(row=kr + 1, column=2, value="   tag").font = GREY
    wsK.cell(row=kr, column=C0 + 11, value=note).alignment = wrap
    kr += 2
I = lambda iid, i: f"Inputs!{qc(i)}{ROW[iid]}"
def guard(expr): return f"=IFERROR({expr},\"n/d\")"
Q10 = list(range(1, 11))
def only(idx, f):
    return lambda i: f(i) if i in idx else None

section("A. Revenue per MW ($M per MW-year, annualized: quarterly revenue x 4 / MW)")
kpi("K01","Revenue per active MW - period-end reported MW","$M/MW-yr","D1", only([4,5,6,7,8,9,10], lambda i: guard(f"{I('I01',i)}*4/{I('I09',i)}")),
    [None,None,None,None,"D","D","D","D","D","D","D"],
    "Nebius KPIs!F17 'Revenue per energized/active MW' (= AI rev x4 / MW) and parts/comps_rev_per_mw.csv (CRWV Q2'26 6.87 on period-end MW). Understates when MW ramps late in quarter.")
kpi("K02","Revenue per active MW - average MW (2024 uses interpolated MW)","$M/MW-yr","D1", only(Q10, lambda i: guard(f"{I('I01',i)}*4/{I('I09A',i)}")),
    [None,"E","E","E","E","D","D","D","D","D","D"],
    "Same numerator on (open+close)/2 MW - the comps file's 'avg Q2 MW est.' variant (8.24 on 1,250 MW). Preferred trend line.")
kpi("K02b","Revenue per BILLING IT MW - average MW (= K02; CRWV has no separate billed series)","$M/MW-yr","D1", only(Q10, lambda i: guard(f"{I('I01',i)}*4/{I('I09A',i)}")),
    [None,"E","E","E","E","D","D","D","D","D","D"],
    'ALIGN: K02b ≡ K02. CRWV discloses active power only (≈ billed/revenue-generating). No Capacity-style billing MW path. See Notes_Gaps B1 / Definitions D1. Cross-pack: comparison.xlsx K02b CRWV = K02; price_or_mw treats active≈billed.')
kpi("K03","Revenue per contracted MW","$M/MW-yr","D1", only([4,5,6,7,8,9,10], lambda i: guard(f"{I('I01',i)}*4/{I('I10',i)}")),
    [None,None,None,None,"D","D","D","D","D","D","D"], "Nebius KPIs 'Revenue per contracted MW'. Contracted = power under contract incl. not yet built.")
kpi("K04","Contracted / active MW (build-out multiple)","x","D1", only([4,5,6,7,8,9,10], lambda i: guard(f"{I('I10',i)}/{I('I09',i)}")),
    [None,None,None,None,"D","D","D","D","D","D","D"], "Pipeline still to be energized relative to the live base.")
kpi("K05","Adj. EBITDA per active MW (average MW)","$M/MW-yr","D1", only(Q10, lambda i: guard(f"{I('I02',i)}*4/{I('I09A',i)}")),
    [None,"E","E","E","E","D","D","D","D","D","D"], "Nebius KPIs 'Adj. EBITDA per energized MW'. CRWV leases all DCs, so EBITDA is after DC rent (Nebius 'leased cohort' analogue).")

section("B. Revenue per GPU and implied utilization (GPU counts ESTIMATED after YE24)")
kpi("K06","Revenue per GPU ($K per GPU-year, avg GPUs)","$K/GPU-yr","D2", only(Q10, lambda i: guard(f"{I('I01',i)}*4000/{I('I11A',i)}")),
    [None]+["E"]*10, "Nebius KPIs 'Revenue per GPU' = AI rev x 4000 / GPU_Avg. GPUs estimated (I11E). FY24 on reported counts: 1,915 / ((53k+250k)/2) = $12.6K/GPU-yr (Derived, lower-bound counts).", "0.0")
kpi("K07","Implied blended realized price ($/GPU-hr)","$/GPU-hr","D2", only(Q10, lambda i: guard(f"{I('I01',i)}*1000000/({I('I11A',i)}*{I('I26',i)}*24)")),
    [None]+["E"]*10, "Nebius KPIs 'Implied blended $/GPU-hr' = revenue / (avg GPUs x hours). Revenue includes storage/networking/software, so this is all-in revenue per GPU-hour.")
kpi("K08","Fleet-weighted on-demand list price ($/GPU-hr)","$/GPU-hr","D3", only(Q10, lambda i: f"={I('I12W',i)}"),
    [None]+["E"]*10, "Nebius KPIs 'Fleet-weighted on-demand list price'. CRWV list from GPU Market agent's Wayback series (H100 $6.155, GB200 $10.50; flat over the period).")
kpi("K09","Implied utilization (realized $/GPU-hr / weighted list)","%","D3", only(Q10, lambda i: guard(f"{I('I01',i)}*1000000/({I('I11A',i)}*{I('I26',i)}*24)/{I('I12W',i)}*100")),
    [None]+["E"]*10, "Nebius KPIs row 'Implied utilization (blended $/hr / weighted list)' - 'mixes realized-price discount with utilization'. CRWV discloses contract discounts up to ~60% vs on-demand, so a low % is mainly discount, not idle GPUs. NOT a physical utilization rate.", "0.0")

section("C. Capex per MW ($M per MW)")
kpi("K10","Cash capex per new active MW (reported MW)","$M/MW","D4", only([5,6,7,8,9,10], lambda i: guard(f"{I('I04',i)}/{I('I09N',i)}")),
    [None]*5+["D"]*6, "Nebius KPIs 'Capex per new active MW' = Capex / MW_New. Timing mismatch: capex leads energization by 1-3 qtrs, so single quarters are noisy; see K12 (in-service basis) and K13 (trailing 4Q).", "0.0")
kpi("K11","Cash capex per new active MW (incl. estimated 2024 MW)","$M/MW","D4", only(Q10, lambda i: guard(f"{I('I04',i)}/{I('I09NE',i)}")),
    [None,"E","E","E","E","D","D","D","D","D","D"], "2024 quarters use interpolated MW (+72.5 MW/qtr) -> Estimated. FY24 total: 8,702 / 290 MW = $30.0M/MW (Derived).", "0.0")
kpi("K12","In-service additions per new active MW (CRWV capex definition)","$M/MW","D4", only([5,6,7,8,9,10], lambda i: guard(f"{I('I25D',i)}/{I('I09N',i)}")),
    [None]*5+["D"]*6, "= delta(gross PP&E - CIP) / delta active MW. Better matched in time (assets placed in service vs MW going live). FY24: (9,679-1,344)/290 = $28.7M/MW.", "0.0")
kpi("K13","Trailing-4Q cash capex per trailing-4Q new active MW","$M/MW","D4", only([8,9,10], lambda i: guard(f"SUM({qc(i-3)}{ROW['I04']}:{qc(i)}{ROW['I04']})/({I('I09',i)}-{I('I09',i-4)})".replace(f"SUM({qc(i-3)}", f"SUM(Inputs!{qc(i-3)}"))),
    [None]*8+["D"]*3, "Smooths timing. Uses reported MW only (needs MW 4 quarters back).", "0.0")
kpi("K14","Net PP&E per active MW (EoP)","$M/MW","D4", only([0,4,5,6,7,8,9,10], lambda i: guard(f"{I('I07',i)}/{I('I09',i)}")),
    ["D",None,None,None,"D","D","D","D","D","D","D"], "Cross-check used in Nebius capacity/gpu_mw_capex.xlsx Peers tab (CRWV $36M YE25, $31M Q2'26). Includes CIP for not-yet-live capacity -> overstates.", "0.0")
kpi("K15","Gross in-service PP&E per active MW (EoP)","$M/MW","D4", only([0,4,5,6,7,8,9,10], lambda i: guard(f"{I('I05S',i)}/{I('I09',i)}")),
    ["D",None,None,None,"D","D","D","D","D","D","D"], "Nebius 'Nebius checks' row 'in-service servers+network per active MW' analogue (Nebius YE25 $18.4M/MW). CRWV leases DCs, so mostly GPU/network + DC fit-out.", "0.0")

section("D. D&A useful lives and depreciation yield")
kpi("K16","Useful life - technology equipment (GPUs/servers/network)","years","D5", only(Q10, lambda i: "=6"),
    [None]+["R"]*10, "S-1 & 10-K FY25 Note 1: 'Technology equipment 6 years'. Last change: effective 1-Jan-2023 computing equipment 5 -> 6 yrs (FY23 expense -$20M). No change disclosed in any 10-Q through Q2'26. Nebius: 4 -> 5 yrs from 1-Jan-2026.", "0")
kpi("K17","Depreciation yield (D&A x4 / avg in-service gross PP&E)","%","D5", only([5,6,7,8,9,10], lambda i: guard(f"{I('I03',i)}*4/(({I('I05S',i-1)}+{I('I05S',i)})/2)*100")),
    [None]*5+["D"]*6, "Mirrors Nebius Capex_DA 'Depreciation yield - reported basis' (D&A / avg in-service gross PP&E). A 6-yr straight-line fleet would yield ~16.7% on tech equipment; lower values = lives >6 yrs on DC equipment (up to 12 yrs) and new assets placed in service late in quarter.", "0.0")
kpi("K18","Implied average depreciable life (1 / yield)","years","D5", only([5,6,7,8,9,10], lambda i: guard(f"100/{qc(i)}{KR['K17']}")),
    [None]*5+["D"]*6, "Blended across tech equipment (6y), software (3-6y), DC equipment/leaseholds (up to 12y).", "0.0")
kpi("K19","D&A as % of revenue","%","D5", only(Q10, lambda i: guard(f"{I('I03',i)}/{I('I01',i)}*100")),
    [None]+["D"]*10, "Context row.", "0.0")

section("E. Backlog / RPO")
kpi("K20","RPO (ASC 606, EoP)","$B","D6", only([0,4,5,6,7,8,9,10], lambda i: f"={I('I14',i)}"),
    ["R",None,None,None,"R","R","R","R","R","R","R"], "Nebius Hist 'Remaining performance obligations' (same ASC 606 basis).", "0.0")
kpi("K21","Revenue backlog (company metric, EoP)","$B","D6", only([0,4,5,6,7,8,9,10], lambda i: f"={I('I15',i)}"),
    ["R",None,None,None,"R","R","R","R","R","R","R"], "= RPO + other committed amounts subject to delivery (e.g., Q1'25 OpenAI $11.2B before contract effectiveness). Not GAAP.", "0.0")
kpi("K22","RPO % recognized within 24 months","%","D6", only([4,5,6,7,8,9,10], lambda i: f"={I('I16',i)}"),
    [None,None,None,None,"R","R","R","R","R","R","R"], "Nebius tracks the same split (36% <24m at 6/30/26).", "0")
kpi("K23","RPO / annualized quarterly revenue","years","D6", only([4,5,6,7,8,9,10], lambda i: guard(f"{I('I14',i)}*1000/({I('I01',i)}*4)")),
    [None,None,None,None,"D","D","D","D","D","D","D"], "Coverage of current run-rate.", "0.0")

section("F. Customer concentration (% of quarter revenue; customers <10% not disclosed)")
kpi("K24","Largest customer % of revenue","%","D7", only(Q10, lambda i: f"={I('I17',i)}"),
    [None,"R","R","R","D","R","R","R","D","R","R"], "Note 2 'customers >=10% of revenue'. Q4 = FY less 9M (Derived, +/-2pp rounding).", "0")
kpi("K25","2nd-largest customer %","%","D7", only([1,2,3,4,9,10], lambda i: f"={I('I18',i)}"),
    [None,"R","R","R","D",None,None,None,None,"R","R"], "n/d = below 10% threshold (not disclosed).", "0")
kpi("K26","Top-N >=10% customers combined %","%","D7", only(Q10, lambda i: f"=SUM({I('I17',i)},{I('I18',i)},{I('I19',i)})"),
    [None,"R","R","R","D","R","R","R","D","R","R"], "Sum of disclosed >=10% customers (Q1'24 top-3, Q2'26 top-3). FY24 top-2 = 77% (10-K FY25).", "0")
kpi("K27","Number of customers >=10% of revenue","#","D7", only(Q10, lambda i: f"={I('I20',i)}"),
    [None,"R","R","R","D","R","R","R","D","R","R"], "", "0")

section("G. Funding cost (% p.a.)")
kpi("K28","Weighted effective cost of debt (principal-weighted, EoP) - all facilities","%","D8", lambda i: (f"={drefq('wavg',i)}" if drefq('wavg',i) else None),
    ["R" if False else "D",None,None,None,"D","D","D","D","D","D","D"], "Mirrors comps.csv 'Wtd effective cost of debt (%) = principal-weighted effective rates, Note 10 debt table' (CRWV 8.41 at 6/30/26). Rates rounded to whole % in filings -> +/-0.5pp.")
for cat, kid in zip(CATS, ["K29","K30","K31","K32","K33","K34"]):
    kpi(kid, f"  of which {cat} - weighted eff. rate", "%", "D8", (lambda c: (lambda i: (f"={drefq(c,i)}" if drefq(c,i) else None)))(cat),
        ["D",None,None,None,"D","D","D","D","D","D","D"], "n/d where facility type not outstanding.")
kpi("K35","DDTL (secured asset-level) share of principal","%","D8", lambda i: (f"={drefq('ddtlsh',i)}" if drefq('ddtlsh',i) else None),
    ["D",None,None,None,"D","D","D","D","D","D","D"], "", "0")
kpi("K36","Implied cash cost: contractual interest x4 / avg principal","%","D8", only([5,6,7,8,9,10], lambda i: guard(f"{I('I21',i)}*4/{I('I23A',i)}*100")),
    [None]*5+["D"]*6, "Before capitalization; excludes discount amortization. FY24: 474.8 / avg(1,555; 8,033) = 9.9% (Derived; average understates mid-year balances).")
kpi("K37","Company-reported weighted average interest rate","%","D8", only([0,4,8,10], lambda i: f"={I('I24',i)}"),
    ["R",None,None,None,"R",None,None,None,"R",None,"R"], "Sep-26 convert deck slide 27 (OEM assumed 10%). Q2'26 PF 7.8%.", "0.0")
kpi("K38","Total debt principal (EoP)","$B","D8", only([0,4,5,6,7,8,9,10], lambda i: guard(f"{I('I23',i)}/1000")),
    ["R",None,None,None,"R","R","R","R","R","R","R"], "Excl. leases (operating lease liabilities $10.1B at 3/31/26 per Q1'26 10-Q risk factors).", "0.0")
kpi("K39","Debt principal / RPO","x","D8", only([0,4,5,6,7,8,9,10], lambda i: guard(f"{I('I23',i)}/1000/{I('I14',i)}")),
    ["D",None,None,None,"D","D","D","D","D","D","D"], "Company shows debt / backlog 0.2x-0.5x.", "0.00")
kpi("K40","Capitalized interest % of contractual interest","%","D8", only(Q10, lambda i: guard(f"{I('I22',i)}/{I('I21',i)}*100")),
    [None,"R","R","R","D","R","R","R","D","R","R"], "Context: share of interest deferred into PP&E (later depreciated).", "0")

wsK.freeze_panes = "E4"

# Sources
wsS = wb.create_sheet("Sources")
header(wsS, 1, ["Input ID", "Item", "Period", "Value", "Tag", "Document", "Section / page / formula", "URL"], [9, 48, 9, 14, 11, 52, 70, 90])
for j, row in enumerate(SRC_ROWS, 2):
    for k, v in enumerate(row, 1):
        wsS.cell(row=j, column=k, value=v)
j = len(SRC_ROWS) + 3
wsS.cell(row=j, column=1, value="Document register").font = BOLD; j += 1
for k, (doc, url) in S.items():
    wsS.cell(row=j, column=1, value=k); wsS.cell(row=j, column=6, value=doc); wsS.cell(row=j, column=8, value=url); j += 1
wsS.cell(row=j+1, column=1, value="Local copies: /workspace/nebius/peers/coreweave/raw/ (S-1, 10-Qs Q1-Q3'25 & Q1'26, 10-K FY25, earnings PRs, .txt extracts) and /workspace/nebius/comps_raw/ (Q2'26 10-Q, Q2'26/Q4'25 PRs, Aug/Sep-26 8-Ks, convert deck, XBRL).")
wsS.cell(row=j+3, column=1, value='POLISH'); wsS.cell(row=j+3, column=2, value='EDGAR scan 2026-10-05')
wsS.cell(row=j+3, column=3, value="No new operating filing after Q2'26 10-Q (acc 0001769628-26-000366, filed 2026-08-12). Oct 2–3 2026 Form 4 / Form 144 RSU sell-to-cover only. Convert package Sep 17–22 already in pack (C1).")
wsS.cell(row=j+3, column=7, value='https://www.stocktitan.net/sec-filings/CRWV/ ; comps_raw/cf_0001769628.json')

# Definitions
wsF = wb.create_sheet("Definitions")
header(wsF, 1, ["Def", "KPI group", "Definition used (formula, units)", "Mirrored from (Nebius workspace)", "CoreWeave application / basis notes"], [6, 24, 70, 70, 80])
DEFS = [
 ("D1","Revenue per MW","Quarterly revenue x 4 / MW, $M per MW-year. Denominators: (a) active MW at period end; (b) average of opening/closing active MW; (c) contracted MW. EBITDA/MW same construction.",
  "/workspace/nebius/model/nebius_model.xlsx KPIs rows 'Revenue per BILLING / contracted / energized MW' (=AIRev*4/MW); /workspace/nebius/parts/comps_rev_per_mw.csv (period-end and avg-MW variants); capacity/gpu_mw_capex.md s.4 MW basis.",
  "CRWV 'active power' = power supporting live DCs ('data centers... for which we are receiving active power', S-1 glossary). IT vs gross never stated; YE24 GPU density (~694/MW vs Nebius Hopper 713/IT-MW, 548/gross-MW) suggests IT basis - same reading as Nebius 'active' (IT). CRWV has no 'billing MW' or 'connected MW' disclosure, so Nebius's billing-MW variant is n/d. Whole-company revenue used (CRWV is ~100% AI cloud). POLISH 2026-10-05: K02b ≡ K02 for CRWV — no separate billed-IT MW series; active power ≈ billed (S-1 glossary). Nebius billing-MW variant remains n/d on CRWV."),
 ("D2","Revenue per GPU","Quarterly revenue x 4000 / average GPUs ($K per GPU-year); implied $/GPU-hr = revenue / (avg GPUs x days x 24).",
  "nebius_model.xlsx KPIs 'Revenue per GPU ($K per GPU-yr)' (=AIRev*4000/GPU_Avg) and 'Implied blended $/GPU-hr' (=AIRev*1e6/(GPU_Avg*Hours)); Timeline Hours = days x 24.",
  "GPU counts reported only at YE22/YE23/YE24 (S-1, 'more than'). Post-YE24 GPUs ESTIMATED as MW x GPUs/MW using Nebius gpu_mw_capex.xlsx base IT-MW densities (GB200 538 for 2025 adds, GB300 491 for 2026 adds) - the Nebius capacity workbook's own method ('GPUs: energized base MW x GPUs/MW'). Nebius found its own GPU counts include delivered-not-powered GPUs; CRWV's S-1 count is 'running' GPUs."),
 ("D3","Implied utilization","Implied blended $/GPU-hr / fleet-weighted on-demand list price (period-end GPU mix weights).",
  "nebius_model.xlsx KPIs 'Fleet-weighted on-demand list price' and 'Implied utilization (blended $/hr / weighted list)' - flagged there as 'Mixes realized-price discount with utilization'. Prices: /workspace/nebius/parts/price_history_quarterly.csv.",
  "List prices: H100 SXM $49.24/8 = $6.155 (Hopper proxy), GB200 NVL72 $42/4 = $10.50 (Blackwell proxy). Committed contracts (the vast majority of CRWV revenue) are priced below list (CRWV advertises up to ~60% discounts), so the ratio is a price-realization x utilization index, not physical utilization. CRWV discloses no utilization metric."),
 ("D4","Capex per MW","Capex / new active MW in the quarter ($M per MW). Primary numerator = cash purchases of PP&E (Nebius Hist definition). Variants: in-service additions (delta gross PP&E - delta CIP = CRWV's own capex definition) per new MW; trailing-4Q; stock cross-checks net PP&E / active MW and in-service gross PP&E / active MW.",
  "nebius_model.xlsx KPIs 'Capex per new active MW ($M/MW)' (=Capex/MW_New); Hist 'Capex (purchases of PP&E + intangibles, cash)'; Capex_DA 'Transfer into service'; capacity/gpu_mw_capex.xlsx 'Nebius checks' (servers+network per active MW) and 'Peers' (CRWV net PP&E per active MW).",
  "CRWV leases all data centers, so capex/MW is mostly GPU + network + DC fit-out (compare to Nebius 'leased/colo' case, not owned-shell case). OEM/vendor-financed equipment can appear in PP&E without cash capex (cash basis understates)."),
 ("D5","D&A useful lives","Disclosed useful-life table and any change in estimate; depreciation yield = D&A x 4 / average in-service gross PP&E; implied life = 1/yield.",
  "/workspace/nebius/filings/nebius_filings.xlsx D&A_Policy tab; nebius_model.xlsx Capex_DA 'Depreciation yield - REPORTED basis' and 'In-service gross PP&E'.",
  "CRWV: technology equipment 6 yrs; software 3-6; DC equipment & leaseholds up to 12 (shorter of lease); last change 1-Jan-2023 (5->6). Total D&A used (PP&E-only depreciation is rounded to $0.1B in 2026 filings)."),
 ("D6","Backlog / RPO","RPO = ASC 606 remaining performance obligations at period end ($B); % within 24 months; company 'revenue backlog' shown separately; RPO / (quarterly revenue x 4).",
  "nebius_model.xlsx Hist 'Remaining performance obligations'; nebius_filings.xlsx Quarterly 'RPO (end of period, $M)'.",
  "RPO is the like-for-like metric vs Nebius. CRWV 'revenue backlog' adds committed amounts not yet in RPO (e.g., $11.2B OpenAI at Q1'25)."),
 ("D7","Customer concentration","Each customer >=10% of quarter revenue (% of revenue), top-N combined, count of >=10% customers.",
  "nebius_filings.xlsx Quarterly 'Customer concentration (>10% of revenue, quarter)'; shared/findings_log.xlsx concentration row.",
  "Q4 values derived as FY less 9M (rounded % -> +/-2pp)."),
 ("D8","Funding cost","Weighted effective cost of debt = sum(principal x effective interest rate) / sum(principal) at period end, all debt facilities excl. leases; split by facility type (DDTL, senior notes, converts, OEM/vendor, revolver, other). Cross-checks: contractual interest x4 / avg principal; company-reported weighted rate.",
  "/workspace/nebius/parts/comps_sources.csv row 'Wtd effective cost of debt (%) - Principal-weighted effective rates, Note 10 debt table' (CRWV 8.41%) and findings_log cost-of-debt row (Nebius 5.1%).",
  "Effective rates as printed in Note 10 (whole %). Converts carry low effective rates (2%) -> blended rate falls as converts grow, while secured DDTL/HY costs stay ~9-11%."),
]
for j, d in enumerate(DEFS, 2):
    for k, v in enumerate(d, 1):
        wsF.cell(row=j, column=k, value=v).alignment = wrap

# Notes / gaps (POLISH 2026-10-05 layout: A = remaining n/d, B = caveats, C = post-quarter / cross-pack)
wsN = wb.create_sheet("Notes_Gaps")
NG_ROWS = [['Notes & Gaps — CoreWeave KPI pack (polished 2026-10-05 ET). Section A = remaining n/d (do not fill with guesses). Section B = caveats on Derived/Estimated cells. Section C = post-quarter / cross-pack notes.', None, None, None, None], [None, None, None, None, None], ['#', 'Type', 'Item (keep tagged; do not invent)', 'Impact on KPIs', 'Where it shows / possible fix'], ["A. REMAINING n/d (as of polish date; last operating filing = Q2'26 10-Q filed 2026-08-12; EDGAR scan through 2026-10-05 found only Form 4/144 sell-to-cover on Oct 2-3 — no new MW/revenue/GPU data)", None, None, None, None], ['A1', 'n/d', "Q1–Q3'24 active MW, contracted MW, balance sheet, and debt tables (pre-IPO; S-1 gives YE23 & YE24 only).", 'K01/K03/K04/K10–K15/K28–K36 n/d or Estimated for those quarters. I09E interpolates MW YE23→YE24 (E) so K02/K05 exist but are Estimated.', 'Inputs I09/I10/I05–I08/Debt_Facilities cols F–H = n/d. Unlikely to appear in primary filings.'], ['A2', 'n/d', "GPU counts after YE24. S-1: '>250,000' at YE24; no subsequent disclosure in 10-Q/10-K/PR through Q2'26.", 'K06–K09 Estimated from MW × density (P01/P02). Density choice alone ≈ ±15–20%.', 'Inputs I11 = n/d after col I; I11E = Estimated. Fix only if company discloses fleet GPUs.'], ['A3', 'n/d', "MW basis (IT vs gross / critical IT load). Never stated. Several MW figures are 'more than' / 'approximately' (lower-bound wording).", 'All $/MW ratios carry PUE ambiguity (~1.2–1.4×) and possible lower-bound bias on MW (→ higher $/MW). Single-quarter ΔMW noisy (±10–50).', 'Definitions D1; Inputs I09 note. YE24 density ~694 GPUs/MW ≈ Nebius Hopper 713/IT-MW → lean IT (E), not proven. Ask IR / site-level leases.'], ['A4', 'n/d', 'Company-reported physical GPU utilization. Not disclosed.', 'K09 is NOT physical utilization — it is realized $/GPU-hr ÷ list (price-realization index).', "Definitions D3; KPI K09 caveat. Do not compare K09 to peers' physical util %."], ['A5', 'n/d', "Billed / billing MW as a separate metric. CRWV discloses only 'active power' and 'contracted power' — no billing-MW or connected-MW series.", 'Nebius K02b (billed IT) has no CRWV twin series. Pack sets K02b = K02 (see B1 / KPI K02b).', "price_or_mw.xlsx MW_Paths: 'active power ≈ billed'. comparison.xlsx K02b CRWV cell = K02."], ['A6', 'n/d', "Q1–Q3'24 interim RPO / revenue backlog / customer concentration (pre-IPO).", 'K20–K27 n/d for those quarters.', 'Inputs I14–I20 cols F–H.'], ['A7', 'n/d', 'Adj. EBITDA bridge detail / segment EBITDA (CRWV is ~100% AI cloud; company Adj. EBITDA is the perimeter).', 'None for K05 (uses company Adj. EBITDA). Not a gap for perimeter choice.', '—'], ['B. CAVEATS on Derived / Estimated cells (values exist but read carefully)', None, None, None, None], ['B1', 'Align', "K02b = K02 for CoreWeave. Reason: no separate billed-IT MW path. Company 'active power' = power at live DCs receiving power (S-1 glossary) ≈ revenue-generating / billed capacity. price_or_mw and comparison treat CRWV active = billed.", 'K02b is a documentation/alignment row (same formula as K02 on I09A). Not a second estimate.', 'KPI K02b; Definitions D1; Inputs I09 note; peers/price_or_mw.xlsx; peers/comparison.xlsx K02b.'], ['B2', 'Caveat', 'Implied utilization (K09) mixes contract discounts with idle capacity (Nebius-defined caveat).', 'Read K09 as price-realization, not physical utilization (see A4).', 'KPI K09 tag row + Definitions D3.'], ['B3', 'Caveat', "Quarterly capex/MW is timing-distorted (GPUs bought ahead of energization; e.g. Q1'26 cash capex $7.7B).", 'Prefer K12 (in-service additions / new MW) and K13 (trailing-4Q) over single-quarter K10.', 'KPI K10–K13.'], ['B4', 'Caveat', 'Customer letters (A/B/C) re-assigned per filing; Q4 concentration = FY − 9M (±2pp rounding).', 'K24–K27 Q4 values Derived with rounding noise.', 'Inputs I17–I20 notes.'], ['B5', 'Caveat', 'Note 10 effective rates rounded to whole %; ranges entered at midpoint.', '±0.5pp on K28–K34.', 'Debt_Facilities.'], ['B6', 'Caveat', '2024 interim MW (I09E) = straight-line YE23~70 → YE24~360. Capex/BS still n/d for those quarters.', "K02/K05 for Q1–Q3'24 are Estimated on interpolated MW only.", 'Inputs I09E tag = Reported/Estimated.'], ["C. POST-QUARTER / CROSS-PACK notes (not in Q2'26 columns; do not backfill into quarterly grid)", None, None, None, None], ['C1', 'Note', "Post-Q2 financing (already in Debt_Facilities memo / Notes): DDTL 5.5 $2.6B SOFR+5.50% (Aug-26); $4.2B 2.875% converts due 2031 (Sep-26). Not in Q2'26 EoP debt stock.", 'Blended cost keeps falling via converts; marginal secured still ~9–10%.', 'Debt_Facilities; comps_raw/crwv_0917_*; crwv_0922_*.'], ['C2', 'Note', 'Sep-26 convert deck: ~70% price increase across SKUs Jul-26; ~$40M/MW marginal short-term contracts (slide 7).', "Supports rising rev/MW into H2'26; $40M/MW is marginal, not fleet average (fleet K02 = $8.24 Q2'26).", 'comps_raw/crwv_0917__crwv2026convertpresentat.htm.txt.'], ['C3', 'Note', "Q2'26 PR backlog ~$104B excludes >$25B net new commitments early Q3'26; contracted power ~4.2 GW as of Aug-11-2026 (vs 3.7 GW at Jun-30).", 'K10/K21 as-of Jun-30 understate post-quarter book.', 'Q2 PR; convert deck.'], ['C4', 'Note', 'Cross-pack: peers/price_or_mw.xlsx uses CRWV active MW path for FY27 $/billed test (active≈billed). peers/payback_dilution.xlsx uses CRWV K02 series for dilution fit. peers/comparison.xlsx K02b CRWV = K02.', 'Keep I09/I09A as the single MW series; do not invent a parallel billed path.', 'Those workbooks — do not edit from this polish.'], ['C5', 'Scan', "EDGAR / comps_raw scan 2026-10-05: no Q3'26 10-Q or earnings PR yet. Latest ops data still Q2'26 (10-Q acc 0001769628-26-000366). Oct 2–3 filings = Form 4 / Form 144 RSU sell-to-cover only — no KPI updates.", 'Quarterly grid unchanged vs Oct-4 pack on numbers.', 'stocktitan CRWV filings; comps_raw/cf_0001769628.json latest ends 2026-06-30.']]
_NG_TITLE = PatternFill("solid", fgColor="D9E1F2"); _NG_HDR = PatternFill("solid", fgColor="DDEBF7"); _NG_A = PatternFill("solid", fgColor="FFF2CC"); _NG_ND = PatternFill("solid", fgColor="F2F2F2")
for j, rowv in enumerate(NG_ROWS, 1):
    for k, v in enumerate(rowv, 1):
        if v is not None:
            c = wsN.cell(row=j, column=k, value=v); c.alignment = wrap
    a0 = rowv[0]
    if j == 1:
        wsN.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5); wsN.cell(row=1, column=1).font = BOLD; wsN.cell(row=1, column=1).fill = _NG_TITLE
    elif j == 3:
        for k in range(1, 6): wsN.cell(row=3, column=k).font = BOLD; wsN.cell(row=3, column=k).fill = _NG_HDR
    elif isinstance(a0, str) and a0[:2] in ("A.", "B.", "C.") and all(v is None for v in rowv[1:]):
        wsN.merge_cells(start_row=j, start_column=1, end_row=j, end_column=5)
        c = wsN.cell(row=j, column=1); c.font = BOLD; c.fill = _NG_A if a0.startswith("A.") else _NG_HDR; c.alignment = Alignment(wrap_text=False)
    elif rowv[1] == "n/d":
        wsN.cell(row=j, column=2).fill = _NG_ND
for col, w in zip("ABCDE", [5, 14, 95, 55, 55]): wsN.column_dimensions[col].width = w

# Comparison template
wsC = wb.create_sheet("Comparison")
wsC.cell(row=1, column=1, value="Comparison (template) - latest quarter, same KPI IDs. CoreWeave column is LIVE-linked to KPIs!O (Q2'26). Nebius column = values from existing Nebius workspace files (cited); blanks = to fill. Add peers by linking their KPIs latest column.").font = BOLD
header(wsC, 3, ["ID", "KPI", "Unit", "Nebius (NBIS) Q2'26", "Nebius source / basis", "CoreWeave Q2'26", "IREN", "Oracle OCI", "Private peer 1", "Private peer 2", "Comment"], [7, 50, 12, 14, 60, 14, 10, 10, 12, 12, 50])
NB = {
 "K01": (13.7, "parts/comps_rev_per_mw.csv: Q2'26 rev x4 / ~170 MW YE25 active (stale MW -> overstated; Q2 active MW n/d)"),
 "K20": (37.4906, "filings/nebius_filings.xlsx Quarterly: RPO $37,490.6M at 6/30/26"),
 "K22": (36, "model Hist RPO note: 36% <24 months at 6/30/26"),
 "K24": (24, "filings/nebius_filings.xlsx Quarterly: Customer C 24% (Q2'26)"),
 "K26": (59, "C 24% + D 21% + E 14% (Q2'26)"),
 "K27": (3, "Q2'26 Quarterly tab"),
 "K16": (5, "filings D&A_Policy: servers & network 4 -> 5 yrs from 1-Jan-2026"),
 "K28": (5.1, "shared/findings_log.xlsx (GPU Market & Comps): wtd effective 5.1% on converts + secured"),
}
cr = 4
for kid in ["K01","K02","K02b","K03","K05","K06","K07","K09","K10","K12","K13","K14","K15","K16","K17","K20","K21","K22","K24","K26","K27","K28","K30","K31","K36"]:
    rr = KR[kid]
    wsC.cell(row=cr, column=1, value=kid)
    wsC.cell(row=cr, column=2, value=f"=KPIs!B{rr}"); wsC.cell(row=cr, column=3, value=f"=KPIs!C{rr}")
    if kid == "K02b":  # POLISH 2026-10-05 alignment row
        wsC.cell(row=cr, column=4, value="see peers/comparison.xlsx K02b (13.06 E)").font = GREY
        wsC.cell(row=cr, column=5, value="Capacity billing path").alignment = wrap
        wsC.cell(row=cr, column=8, value="ALIGN K02b≡K02")
        wsC.cell(row=cr, column=11, value="ALIGN: K02b≡K02 (active≈billed). See Notes_Gaps B1. Canonical peer table: peers/comparison.xlsx").alignment = wrap
    elif kid in NB:
        c = wsC.cell(row=cr, column=4, value=NB[kid][0]); c.font = BLUE
        wsC.cell(row=cr, column=5, value=NB[kid][1]).alignment = wrap
    else:
        wsC.cell(row=cr, column=4, value="to fill").font = GREY
    c = wsC.cell(row=cr, column=6, value=f"=KPIs!O{rr}"); c.number_format = "0.0"
    cr += 1
wsC.cell(row=cr+1, column=1, value="Basis warnings: Nebius rev/MW uses stale YE25 MW; Nebius MW 'active' is IT basis per its own definition, CRWV basis unstated (likely IT). Nebius capex/MW not on same quarterly basis (2026 guide implies $20.5-34.5M per connected MW, capacity/gpu_mw_capex.md s.3).").alignment = wrap

for _ws in wb:
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and _c.value.startswith("= "):
                _c.value = "calc: " + _c.value[2:]
wb.save(OUT)
print("saved", OUT, "inputs", len(INPUTS), "src rows", len(SRC_ROWS))
