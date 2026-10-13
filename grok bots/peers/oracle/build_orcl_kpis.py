# Builds oracle_kpis.xlsx (Oracle OCI peer KPI comps) on the CoreWeave/Nebius template.
# Calendar columns Q1'24..Q2'26 mapped to Oracle FQ3 FY24 (Feb-24) .. FQ4 FY26 (May-26) (Oracle FY ends 31-May; fiscal quarters end ~1 month
# before calendar quarter-ends) + extra column Q3'26-dagger = FQ1 FY27 (Aug-26, latest filed).
# Engine: /workspace/nebius/peers/peer_kpi_template.py. Author: Grok Bot for Aadit Prayag, 2026-10-05 ET.
import sys, json, re; sys.path.insert(0, "/workspace/nebius/peers")
from peer_kpi_template import Peer, build

OUT = "/workspace/nebius/peers/oracle/oracle_kpis.xlsx"
QS = ["Q4'23*", "Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26†"]
QEND = ["2023-11-30", "2024-02-29", "2024-05-31", "2024-08-31", "2024-11-30", "2025-02-28", "2025-05-31", "2025-08-31", "2025-11-30", "2026-02-28", "2026-05-31", "2026-08-31"]
FQ = ["FQ2 FY24 (Nov-23)", "FQ3 FY24 (Feb-24)", "FQ4 FY24 (May-24)", "FQ1 FY25 (Aug-24)", "FQ2 FY25 (Nov-24)", "FQ3 FY25 (Feb-25)", "FQ4 FY25 (May-25)",
      "FQ1 FY26 (Aug-25)", "FQ2 FY26 (Nov-25)", "FQ3 FY26 (Feb-26)", "FQ4 FY26 (May-26)", "FQ1 FY27 (Aug-26)"]
DAYS = [None, 91, 92, 92, 91, 90, 92, 92, 91, 90, 92, 92]
E = "https://www.sec.gov/Archives/edgar/data/1341439/"
PRACC = {"PR24Q3": ("FQ3 FY24", "000095017024029250", "2024-03-11"), "PR24Q4": ("FQ4 FY24", "000095017024071916", "2024-06-11"),
         "PR25Q1": ("FQ1 FY25", "000095017024104753", "2024-09-09"), "PR25Q2": ("FQ2 FY25", "000095017024134537", "2024-12-09"),
         "PR25Q3": ("FQ3 FY25", "000095017025036295", "2025-03-10"), "PR25Q4": ("FQ4 FY25", "000095017025084831", "2025-06-11"),
         "PR26Q1": ("FQ1 FY26", "000119312525199175", "2025-09-09"), "PR26Q2": ("FQ2 FY26", "000119312525314207", "2025-12-10"),
         "PR26Q3": ("FQ3 FY26", "000119312526100148", "2026-03-10"), "PR26Q4": ("FQ4 FY26", "000119312526265848", "2026-06"),
         "PR27Q1": ("FQ1 FY27", "000119312526387905", "2026-09-10")}
S = {k: (f"{v[0]} earnings PR (8-K Ex.99.1, {v[2]})", E + v[1] + "/orcl-ex99_1.htm") for k, v in PRACC.items()}
S.update({
 "Q24Q3": ("10-Q FQ3 FY24 (Feb-24)", E + "000095017024029904/orcl-20240229.htm"),
 "Q25Q1": ("10-Q FQ1 FY25 (Aug-24)", E + "000095017024104905/orcl-20240831.htm"),
 "Q25Q2": ("10-Q FQ2 FY25 (Nov-24)", E + "000095017024134973/orcl-20241130.htm"),
 "Q25Q3": ("10-Q FQ3 FY25 (Feb-25)", E + "000095017025037143/orcl-20250228.htm"),
 "Q26Q1": ("10-Q FQ1 FY26 (Aug-25)", E + "000119312525200095/orcl-20250831.htm"),
 "Q26Q2": ("10-Q FQ2 FY26 (Nov-25)", E + "000119312525315925/orcl-20251130.htm"),
 "Q26Q3": ("10-Q FQ3 FY26 (Feb-26)", E + "000119312526101045/orcl-20260228.htm"),
 "Q27Q1": ("10-Q FQ1 FY27 (Aug-26), filed 2026-09", E + "000119312526389274/orcl-20260831.htm"),
 "K24": ("10-K FY24 (FYE 31-May-2024)", E + "000095017024075605/orcl-20240531.htm"),
 "K25": ("10-K FY25 (FYE 31-May-2025)", E + "000095017025087926/orcl-20250531.htm"),
 "K26": ("10-K FY26 (FYE 31-May-2026)", E + "000119312526277521/orcl-20260531.htm"),
 "XBRL": ("SEC companyfacts XBRL CIK 0001341439 (local copy comps_raw/cf_0001341439.json)", "https://data.sec.gov/api/xbrl/companyfacts/CIK0001341439.json"),
 "CALL26Q2": ("Oracle FQ2 FY26 earnings call transcript (Motley Fool, 2025-12-10)", "https://www.fool.com/earnings/call-transcripts/2025/12/10/oracle-orcl-q2-2026-earnings-call-transcript/"),
 "CALL26Q3": ("Oracle FQ3 FY26 earnings call - CEO blog 'From the Q3 earnings call' (+ transcript 2026-03-10)", "https://blogs.oracle.com/ceo/from-the-q3-earnings-call"),
 "CALL26Q4": ("Oracle FQ4 FY26 earnings call - CEO blog 'From the Q4 earnings call'", "https://blogs.oracle.com/ceo/from-the-q4-earnings-call"),
 "CALL27Q1": ("Oracle FQ1 FY27 earnings call transcript (Motley Fool, 2026-09-11)", "https://www.fool.com/earnings/call-transcripts/2026/09/11/oracle-orcl-q1-2027-earnings-call-transcript/"),
 "NBPX": ("Nebius workspace: /workspace/nebius/parts/price_history_raw.csv + price_snapshot_2026-10-04.csv (OCI list, GPU Market agent; observations Aug-Oct 2026 only)", "local"),
 "ASSUMP": ("Analyst assumption (this workbook) - segment allocation", "n/a"),
 "CAL": ("calendar", ""),
})
P = Peer("Oracle OCI (ORCL) - Cloud Infrastructure perimeter", OUT, QS, QEND, FQ, DAYS, S)
nd = None; add = P.add; N = 12
PRQ = ["", "PR25Q3", "PR25Q4", "PR26Q4", "PR26Q4", "PR26Q4", "PR26Q4", "PR27Q1", "PR27Q1", "PR27Q1", "PR27Q1", "PR27Q1"]
add("I01","OCI revenue = 'Cloud infrastructure' revenue (IaaS)","$M",[nd,1779.9,1970.4,2154,2434,2652,2995,3347,4079,4888,5787,7388],
    [None,"D","D"]+["R"]*9, PRQ,
    ["","= FQ3 FY25 2,652 / 1.49 (PR: cloud infrastructure +49% USD); headline '$1.8 billion' (FQ3 FY24 PR)","= FQ4 FY25 2,995 / 1.52 (+52% USD); headline '$2.0 billion' (FQ4 FY24 PR)",
     "FQ4 FY26 PR revenue-by-offering table (FY25 quarters)","same","same","same","FQ1 FY27 PR table (FY26 quarters)","same","same","same","FQ1 FY27 PR 'Cloud infrastructure 7,388'"],
    "OCI is not a reported segment: 'Cloud infrastructure' is a revenue line inside the Cloud & license segment. Includes non-AI OCI (general compute, database on OCI, multicloud) - AI share not disclosed (FQ3 FY26 call: AI infrastructure revenue +243% YoY, no level).")
add("I02","Adj. EBITDA of OCI","$M",[nd]*N,[None]*N,"","","n/d - no segment EBITDA/margin for OCI. FQ3 FY26 call: gross margin on delivered AI capacity ~32% (company statement, not a filed metric).")
add("I03","Depreciation (company-wide, cash-flow statement; excl. intangible amortization)","$M",[nd,808,811,804,908,1003,1152,1351,1704,2153,2415,3156],
    [None,"D","D","R","D","D","D","R","D","D","D","R"],
    ["","PR24Q3","PR24Q4","PR25Q1","PR25Q2","PR25Q3","PR25Q4","PR26Q1","PR26Q2","PR26Q3","PR26Q4","PR27Q1"],
    ["","= 9M FY24 2,318 - 6M 1,510 (XBRL Depreciation YTD)","= FY24 3,129 - 9M 2,318","FQ1 FY25 PR CF","= 6M 1,712 - 804","= 9M 2,715 - 1,712","= FY25 3,867 - 2,715","FQ1 FY26 PR CF","= 6M 3,055 - 1,351","= 9M 5,208 - 3,055","= FY26 7,623 - 5,208","FQ1 FY27 PR CF"],
    "Company-wide PP&E depreciation (OCI not separable).")
add("I50","Cash capex, company-wide (capital expenditures, CF statement)","$M",[nd,1674,2798,2303,3970,5862,9080,8502,12033,18635,16493,28499],
    [None,"D","D","R","D","D","D","R","D","D","D","R"],
    ["","PR24Q3","PR24Q4","PR25Q1","PR25Q2","PR25Q3","PR25Q4","PR26Q1","PR26Q2","PR26Q3","PR26Q4","PR27Q1"],
    ["","= 9M FY24 4,068 - 6M 2,394","= FY24 6,866 - 9M 4,068","FQ1 FY25 PR CF 2,303","= 6M 6,273 - 2,303","= 9M 12,135 - 6,273","= FY25 21,215 - 12,135","FQ1 FY26 PR CF 8,502","= 6M 20,535 - 8,502","= 9M 39,170 - 20,535","= FY26 55,663 - 39,170","FQ1 FY27 PR CF 28,499"],
    "Excludes leased data-center shells (operating leases; see I43/I51) and GPUs obtained via unconditional purchase commitments / leases.")
add("I05","Gross PP&E (EoP)","$M",[nd,nd,34818,nd,nd,nd,59554,nd,nd,nd,122651,153596],[None,None,"R",None,None,None,"R",None,None,None,"R","R"],
    ["","","K24","","","","K25","","","","K26","Q27Q1"],"PP&E note (10-K annual; FQ1 FY27 10-Q first quarterly disclosure in set)")
add("I06","Construction in progress (EoP)","$M",[nd,nd,5634,nd,nd,nd,16510,nd,nd,nd,39973,48546],[None,None,"R",None,None,None,"R",None,None,None,"R","R"],
    ["","","K24","","","","K25","","","","K26","Q27Q1"],"PP&E note")
add("I07","PP&E, net (EoP)","$M",[18009,19117,21536,23094,26432,31970,43522,53194,67875,83617,99957,127845],["R"]*N,
    ["Q24Q3","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26","Q27Q1"],"Balance sheet (XBRL PropertyPlantAndEquipmentNet)")
add("I08","Computer, network, machinery & equipment, gross (EoP)","$M",[nd,nd,20989,nd,nd,nd,30345,nd,nd,nd,59634,77755],[None,None,"R",None,None,None,"R",None,None,None,"R","R"],
    ["","","K24","","","","K25","","","","K26","Q27Q1"],"PP&E note line (useful life 1-5 yrs FY24; 1-6 yrs FY25+)")
add("I09","Installed / active OCI (AI) MW (EoP)","MW",[nd]*N,[None]*N,"","",
    "n/d - Oracle discloses MW DELIVERED per quarter (I09N), never an installed base. Cumulative deliveries since Jun-25 in I45E (Estimated).")
add("I09N","AI capacity delivered to customers in quarter","MW",[nd]*7+[110,400,400,290,850],[None]*7+["E","R","R","E","R"],
    [""]*7+["ASSUMP","CALL26Q2","CALL26Q3","CALL27Q1","PR27Q1"],
    [""]*7+["ESTIMATED residual = FY26 '>1.2GW' - 400 - 400 - 290 (range ~75-110)","'close to 400 MW' delivered in FQ2 (call)","'more than 400 MW' (lower bound; CEO blog / call)",
            "ESTIMATED: FQ1 FY27 850MW described as 'almost 3x' Q4 -> ~290 (850/2.93)","FQ1 FY27 PR / call: 850MW delivered containing >300k GPUs"],
    "MW basis (IT vs gross) not stated; >300k GPUs in 850MW (~353 GPUs/MW) vs Nebius GB200 538/IT-MW suggests gross/facility MW. Abilene: 618MW = ~75% of site delivered by Aug-26.")
add("I10","Contracted MW (EoP)","MW",[nd]*N,[None]*N,"","","n/d as a filed metric. Calls: '>10 GW secured through partners'. Stargate/Abilene ~1.2GW site (618MW delivered = 75%).")
add("I11","GPUs deployed (fleet, EoP)","GPUs",[nd]*N,[None]*N,"","","n/d fleet count. FQ1 FY27: '>300k GPUs' in the 850MW delivered that quarter (I46); Abilene 131k GPUs delivered in FQ1 FY27.")
add("I12","OCI H100 on-demand list price (BM.GPU.H100.8 / 8)","$/GPU-hr",[nd]*11+[10.0],[None]*11+["R"],[""]*11+["NBPX"],[""]*11+["OCI price list $10.00/GPU-hr (observed Aug-Oct 2026; H200 also $10)"],
    "No earlier snapshots in the price set -> n/d before Q3'26.")
add("I13","OCI GB200 list price","$/GPU-hr",[nd]*11+[16.0],[None]*11+["R"],[""]*11+["NBPX"],[""]*11+["OCI list GB200 $16 (B200 $14, B300 $15, GB300 $18, A100 $4, MI300X $6)"])
add("I14","RPO (company-wide, ASC 606, EoP)","$B",[65.5,80.2,97.9,99.1,97.3,130.2,137.8,455.3,523.3,552.6,638.0,664.0],["R"]*N,
    ["XBRL","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26","Q27Q1"],"Revenue recognition note RPO paragraph (XBRL RevenueRemainingPerformanceObligation)",
    "Company-wide (incl. SaaS/license support). Growth from FQ1 FY26 'primarily attributable to certain significant cloud contracts' (10-K FY26) - i.e., OCI AI contracts (OpenAI/Stargate, Meta, xAI etc.).")
add("I15","Backlog incl. non-RPO","$B",[nd]*N,[None]*N,"","","Oracle reports RPO only.")
add("I16","RPO % within 24 months (ESTIMATED: 12m % + half of 13-36m %)","%",[nd,60.5,57,57,58.5,51,53.5,22.5,25,27.5,29,31.5],[None]+["E"]*11,
    ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26","Q27Q1"],
    ["","43 + 35/2","39 + 36/2","38 + 38/2","39 + 39/2","31 + 40/2","33 + 41/2","10 + 25/2","10 + 30/2","12 + 31/2","12 + 34/2","13 + 37/2"],
    "Oracle discloses <=12m and 13-36m buckets; 24m share interpolated linearly within the 13-36m bucket (assumption).")
add("I17","Largest customer, % of revenue","%",[nd]*N,[None]*N,"","","No customer >=10% of revenue in FY24-FY26 (10-K concentration note) -> largest <10% (not disclosed).")
add("I18","2nd-largest customer %","%",[nd]*N,[None]*N,"","")
add("I19","3rd-largest customer %","%",[nd]*N,[None]*N,"","")
add("I20","Customers >=10% of revenue (count; fiscal-year basis)","#",[nd,nd,0,nd,nd,nd,0,nd,nd,nd,0,nd],[None,None,"R",None,None,None,"R",None,None,None,"R",None],
    ["","","K24","","","","K25","","","","K26",""],"10-K: no single customer accounted for 10% or more of revenues (fiscal year)")
add("I21","Interest expense (P&L, as reported)","$M",[nd,876,878,842,866,892,978,923,1057,1180,1438,1428],[None]+["R"]*11,
    ["","PR24Q3","PR24Q4","PR25Q1","PR25Q2","PR25Q3","PR25Q4","PR26Q1","PR26Q2","PR26Q3","PR26Q4","PR27Q1"],"PR income statement 'Interest expense'",
    "As reported (after any capitalized interest; Oracle does not disclose capitalized interest in PRs/10-Qs in this set).")
add("I22","Capitalized interest","$M",[nd]*N,[None]*N,"","")
add("I23","Total borrowings, carrying value (notes payable & other borrowings, current + non-current)","$M",[nd,87980,86869,84515,88624,96276,92568,91315,108075,134605,129541,125337],[None]+["R"]*11,
    ["","PR24Q3","PR24Q4","PR25Q1","PR25Q2","PR25Q3","PR25Q4","PR26Q1","PR26Q2","PR26Q3","PR26Q4","PR27Q1"],
    ["","5,510 + 82,470","10,605 + 76,264","9,201 + 75,314","8,162 + 80,462","8,167 + 88,109","7,271 + 85,297","9,079 + 82,236","8,091 + 99,984","9,887 + 124,718","7,199 + 122,342","7,625 + 117,712"],
    "Balance-sheet carrying value (net of unamortized discount/issuance costs); 10-K debt tables (Debt_Facilities) are at principal -> check row shows the difference.")
add("I24","Company-reported weighted average interest rate","%",[nd]*N,[None]*N,"","")
add("I25","Company-definition capex","$M",[nd]*N,[None]*N,"","")
add("I26","Days in quarter","days",DAYS,[None]+["R"]*11,"CAL","calendar")
add("I27","Data center regions","#",[nd]*N,[None]*N,"","")
add("I28","Useful life - servers & networking equipment","years",[nd,5,5]+[6]*9,[None]+["R"]*11,["","K24","K24"]+["Q25Q1","Q25Q2","Q25Q3","K25","K25","K25","K25","K26","K26"],
    ["","10-K FY24: servers 4 -> 5 yrs effective FY23","same","FQ1 FY25 10-Q: servers & networking 5 -> 6 yrs effective FY25 (+$733M FY25 opex benefit, 10-K FY25)"]+["same"]*8,
    "PP&E table: computer, network, machinery & equipment 1-5 yrs (FY24) -> 1-6 yrs (FY25-FY26, Aug-26 10-Q).")
add("I29","Company-reported GPU utilization","%",[nd]*10+[97.5,97.9],[None]*10+["R","R"],[""]*10+["CALL26Q4","CALL27Q1"],[""]*10+["FQ4 FY26 call: 97.5%","FQ1 FY27 call: 97.9%"],
    "Physical GPU utilization of delivered fleet (management statement on call, not in filings).")
add("I30","RPO % expected to be recognized within 12 months","%",[nd,43,39,38,39,31,33,10,10,12,12,13],[None]+["R"]*11,
    ["","Q24Q3","K24","Q25Q1","Q25Q2","Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26","Q27Q1"],"RPO paragraph 'approximately X% ... over the next twelve months'")
add("I31","GPU-compute hardware, gross","$M",[nd]*N,[None]*N,"","","n/d (I08 covers all computer/network equipment, not GPU-only).")
add("I32","OCI cost of revenue","$M",[nd]*N,[None]*N,"","","n/d (Oracle reports cloud services & license support expenses combined).")
# memo
add("I40","Total revenue (company)","$M",[nd,13280,14287,13307,14059,14130,15903,14926,16058,17190,19184,19345],
    [None,"R","D","R","R","R","D","R","R","R","D","R"],
    ["","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL","XBRL"],
    ["","10-Q P&L","= FY24 52,961 - 9M 38,674","10-Q","10-Q","10-Q","= FY25 57,399 - 9M 41,496","10-Q","10-Q","10-Q","= FY26 67,357 - 9M 48,173","10-Q"])
add("I43","Operating lease liabilities (EoP)","$M",[nd,nd,7545,nd,nd,nd,13450,16290,18679,21310,30190,34621],[None,None]+["R"]+[None]*3+["R"]*6,
    ["","","K24","","","","K25","Q26Q1","Q26Q2","Q26Q3","K26","Q27Q1"],"Leases note (XBRL OperatingLeaseLiability)")
add("I46","GPUs delivered in quarter","GPUs",[nd]*11+[300000],[None]*11+["R"],[""]*11+["PR27Q1"],[""]*11+["'850MW ... containing more than 300,000 GPUs' (lower bound)"])
add("I51","Lease commitments not yet commenced (substantially all data centers)","$B",[nd]*5+[48.4,43.4,99.8,248,261,260,288],[None]*5+["R"]*7,
    [""]*5+["Q25Q3","K25","Q26Q1","Q26Q2","Q26Q3","K26","Q27Q1"],
    [""]*5+["'$48.4 billion of additional lease commitments, primarily for data centers'","'$43.4 billion'","'$99.8 billion'","'$248 billion'","'$261 billion'","'$260 billion' (commence FQ1 FY27-FY29, 15-19 yr terms)","'$288 billion' (commence FQ2 FY27-FY29)"],
    "Off-balance-sheet DC lease commitments = Oracle's contracted capacity pipeline (capacity built by partners e.g. Crusoe/Blue Owl at Abilene). Not in capex.")
P.param("P05","Share of company capex attributed to OCI / AI infrastructure","%",100,"E","ASSUMP",
        "UPPER-BOUND ASSUMPTION: all capex attributed to OCI (Oracle does not split capex; management describes the build as overwhelmingly revenue-generating cloud infrastructure equipment). Lower share scales K10/K13/M06 proportionally.")
P.derive("I04","Cash capex attributed to OCI (= company capex x P05) - SEGMENT-ALLOCATION ESTIMATE","$M", lambda i: "[I50@%d]*[P05@0]/100" % i, lambda i: [("I50",i),("P05",0)],
    "ESTIMATED allocation (see P05).", force=lambda i: "E")
P.derive("I09T4","AI MW delivered over trailing 4 quarters","MW", lambda i: "1200" if i == 10 else "SUM([I09N@%d],[I09N@%d],[I09N@%d],[I09N@%d])" % (i-3, i-2, i-1, i),
    lambda i: [("I09N",i-3),("I09N",i-2),("I09N",i-1),("I09N",i)], "Q2'26 (FY26) = '>1.2GW' delivered in FY26 (CEO blog, FQ4 call; lower bound, Reported). Q3'26 = sum of quarterly deliveries (incl. estimates).",
    idx=[10, 11], force=lambda i: "R" if i == 10 else "E")
P.derive("I45E","Cumulative AI MW delivered since 1-Jun-2025","MW", lambda i: "SUM(%s)" % ",".join("[I09N@%d]" % j for j in range(7, i+1)), lambda i: [("I09N",j) for j in range(7, i+1)],
    "ESTIMATED (includes estimated quarters). Not an installed base - excludes pre-FY26 OCI capacity.", idx=range(7, N), force=lambda i: "E")
P.derive("I05S","In-service gross PP&E (gross - CIP, EoP)","$M", lambda i: "[I05@%d]-[I06@%d]" % (i, i), lambda i: [("I05",i),("I06",i)], "Company-wide.")
P.derive("I23A","Average borrowings (carrying)","$M", lambda i: "([I23@%d]+[I23@%d])/2" % (i-1, i), lambda i: [("I23",i-1),("I23",i)], "", idx=range(1, N))

# ---- debt facilities from 10-K Notes Payable tables (parsed by work/parse_orcl_debt.py) ----
d = json.load(open("/workspace/nebius/peers/oracle/work/orcl_debt_tables.json"))
norm = lambda s: re.sub(r"(\(\d\))+", "", s).strip()
CATMAP = {"Senior notes (fixed)": "Senior notes", "Floating-rate notes": "Senior notes", "Term loans": "Other term loans", "Commercial paper": "Revolver / commercial paper", "Other": "Other borrowings"}
fac = {}
for f, snap in [("k_fy24__orcl-20240531.htm.txt", "May-24"), ("k_fy25__orcl-20250531.htm.txt", "May-25"), ("k_fy26__orcl-20260531.htm.txt", "May-26")]:
    for r in d[f]["rows"]:
        n = norm(r["name"]); fac.setdefault(n, {"cat": CATMAP[r["cat"]] + (" (floating)" if r["cat"] == "Floating-rate notes" and False else "")})
        v = r.get(snap)
        if v and v[0]: fac[n][snap] = (v[0], v[1])
FAC = [(n + ("  [floating]" if "SOFR" in n and v["cat"] == "Senior notes" else ""), v["cat"], [v.get(s) for s in ("May-24", "May-25", "May-26")]) for n, v in fac.items() if any(v.get(s) for s in ("May-24", "May-25", "May-26"))]
debt = dict(SNAP=["May-24", "May-25", "May-26"], SNAPQ=[2, 6, 10], SNAPSRC=["K24", "K25", "K26"], FAC=FAC,
  CATS=["DDTL", "Senior notes", "Converts", "OEM / vendor financing", "Revolver / commercial paper", "Other term loans"], SECT="Notes to consolidated financial statements - Notes payable and other borrowings table (amount; effective interest rate)",
  RATE_NOTE="Annual snapshots only (10-K tables list every note with effective rate; 10-Qs list only new issuances). Principal = 'Amount' column; check row vs Inputs I23 = unamortized discount/issuance costs & fair-value adjustments (carrying basis). Floating-rate senior notes included in 'Senior notes' (flagged [floating]); term loans (SOFR+1.35-1.70%) in 'Other term loans'; commercial paper in 'Revolver / commercial paper'. Quarterly funding cost: K36 (interest expense x4 / avg carrying debt).",
  POST=["Feb-2026: $25B senior notes issuance (8-K Feb-2026; in May-26 table). Aug-2026 (FQ1 FY27): ~$19.9B equity raised via at-the-market program (PR CF) - deleveraging borrowings to $125.3B.",
        "Off-balance-sheet: $288B of not-yet-commenced data-center lease commitments at Aug-26 (I51) - economically debt-like, excluded here (template excludes leases)."])

ovr = {
 "K01": dict(note="n/d - no installed OCI MW. See K43 (incremental OCI revenue per delivered MW)."),
 "K02": dict(note="n/d (no installed MW)."), "K03": dict(note="n/d (no contracted MW; >10GW 'secured through partners' per calls)."), "K04": dict(note="n/d"),
 "K05": dict(note="n/d - no OCI EBITDA. FQ3 FY26 call: ~32% gross margin on delivered AI capacity (management, not filed)."),
 "K06": dict(note="n/d (no GPU fleet count)."), "K07": dict(note="n/d"),
 "K08": dict(note="n/d (fleet mix unknown). OCI list Aug-Oct 2026: H100/H200 $10, B200 $14, B300 $15, GB200 $16, GB300 $18 per GPU-hr (inputs I12/I13) - ~1.6x CRWV's H100 list."),
 "K09": dict(note="n/d. Oracle states physical GPU utilization instead (K41: 97.5% / 97.9%)."),
 "K10": dict(note="SEGMENT-ALLOCATION ESTIMATE: (company capex x 100% OCI share, P05) / AI MW delivered in the quarter. Upper bound (100% allocation); capex leads delivery by 1-3 quarters; excludes leased shells (I51) -> compare with CRWV's leased-DC capex/MW, not owned-shell builds.",
            tagf=lambda i: "E"),
 "K12": dict(note="n/d (in-service additions only annual)."),
 "K13": dict(note="Trailing-4Q OCI-attributed capex / trailing-4Q MW delivered. Q2'26 = FY26 capex 55.7B / '>1.2GW' (upper bound on $/MW since MW is a lower bound and capex allocation 100%). ESTIMATED.", tagf=lambda i: "E"),
 "K14": dict(note="n/d (no installed MW)."), "K15": dict(note="n/d (no installed MW)."),
 "K16": dict(note="Servers & networking: 4->5 yrs (FY23), 5->6 yrs effective FY25 (Jun-2024; +$733M FY25 opex benefit). No change through Aug-26 10-Q. CRWV 6, Nebius 5 (from 2026), IREN 5."),
 "K17": dict(note="Quarterly yield computable only where gross & CIP exist at both ends (May-26 -> Aug-26). Annual version in memo M05."),
 "K18": dict(note=""),
 "K19": dict(expr=lambda I, i: f"{I('I03',i)}/{I('I40',i)}*100", deps=lambda i: [("I03",i),("I40",i)], name="Depreciation as % of revenue (company / total revenue)", note="Company-wide denominator."),
 "K20": dict(note="Company-wide RPO (OCI share not disclosed; the FQ1 FY26 +$317B jump and later growth attributed by Oracle to 'certain significant cloud contracts'). Media reports (not filings) attribute ~$300B to OpenAI."),
 "K21": dict(note="n/d (RPO only)."),
 "K22": dict(note="ESTIMATED (linear interpolation of 13-36m bucket). Collapse from ~55% to ~25-30% = long-dated AI contracts (34% in months 37-60)."),
 "K23": dict(expr=lambda I, i: f"{I('I14',i)}*1000/({I('I40',i)}*4)", deps=lambda i: [("I14",i),("I40",i)], name="RPO / annualized quarterly total revenue", note="Company RPO / company revenue run-rate (perimeter-consistent)."),
 "K24": dict(note="n/d - no customer >=10% of FY revenue (FY24-FY26). RPO concentration not disclosed."),
 "K25": dict(note="n/d"), "K26": dict(note="n/d"),
 "K27": dict(note="0 on fiscal-year basis (10-K)."),
 "K28": dict(note="Principal-weighted effective rate across all notes/term loans/CP (10-K table, annual). Rising with new issuance at 5-6%+ (Feb-26 $25B)."),
 "K30": dict(note="Fixed + floating senior notes."), "K33": dict(note="Commercial paper."), "K34": dict(note="Term loans (SOFR+1.35-1.70%)."),
 "K35": dict(note="0 - no asset-level DDTL (Oracle funds on corporate balance sheet + partner-financed leases)."),
 "K36": dict(note="Interest expense (as reported) x4 / average carrying borrowings. comps.csv Oracle cost of debt 4.48% (interest / avg debt) consistent."),
 "K38": dict(note="Carrying value (I23)."), "K39": dict(note="Company debt / company RPO."),
 "K40": dict(note="n/d"),
 "K41": dict(note="Management call statements: 97.5% (FQ4 FY26), 97.9% (FQ1 FY27). Physical utilization, not price-based."),
 "K42": dict(note="Reported (10-Q/10-K)."),
 "K43": dict(note="(OCI rev - OCI rev 4Q earlier) x4 / AI MW delivered over the same 4 quarters. SEGMENT-ALLOCATION ESTIMATE: attributes ALL incremental OCI revenue (incl. non-AI OCI growth) to delivered AI MW; MW delivered late in a quarter earns little in-quarter revenue (Q3'26: 850MW delivered in FQ1 FY27).",
            tagf=lambda i: "E"),
 "K44": dict(note="n/d"), "K45": dict(note="n/d (no OCI cost of revenue)."),
}
memo = [
 dict(kid="M01", name="OCI share of total revenue", unit="%", expr=lambda I, i: f"{I('I01',i)}/{I('I40',i)}*100", deps=lambda i: [("I01",i),("I40",i)], note="", fmt="0.0"),
 dict(kid="M02", name="OCI revenue growth YoY", unit="%", expr=lambda I, i: f"({I('I01',i)}/{I('I01',i-4)}-1)*100", deps=lambda i: [("I01",i),("I01",i-4)], note="Matches PR USD growth (55/68/84/93/121%).", fmt="0"),
 dict(kid="M03", name="Company capex / total revenue", unit="%", expr=lambda I, i: f"{I('I50',i)}/{I('I40',i)}*100", deps=lambda i: [("I50",i),("I40",i)], note="Reported basis (no allocation).", fmt="0"),
 dict(kid="M04", name="GPUs per delivered MW (FQ1 FY27)", unit="GPUs/MW", expr=lambda I, i: f"{I('I46',i)}/{I('I09N',i)}", deps=lambda i: [("I46",i),("I09N",i)], note=">300k GPUs / 850MW = ~353/MW (lower bound) vs Nebius base 538 (GB200) / 491 (GB300) per IT-MW -> Oracle MW likely gross/facility basis (PUE ~1.4-1.5x) or includes not-yet-populated halls.", fmt="0"),
 dict(kid="M05", name="Annual depreciation yield (trailing-4Q depreciation / avg in-service gross PP&E, FY)", unit="%", expr=lambda I, i: f"SUM(Inputs!{chr(69+i-3)}{{rI03}}:{chr(69+i)}{{rI03}})/(({I('I05S',i-4)}+{I('I05S',i)})/2)*100", deps=lambda i: [("I03",i-3),("I03",i),("I05S",i-4),("I05S",i)], note="FY basis (May columns). Low vs 6-yr life (16.7%) because buildings/land and long-lived assets are in the base and FY26 additions came in late.", fmt="0.0"),
 dict(kid="M06", name="Trailing-4Q OCI-attributed capex per $ of incremental annualized OCI revenue", unit="x", expr=lambda I, i: f"SUM(Inputs!{chr(69+i-3)}{{rI04}}:{chr(69+i)}{{rI04}})/(({I('I01',i)}-{I('I01',i-4)})*4)", deps=lambda i: [("I04",i-3),("I04",i),("I01",i),("I01",i-4)], note="ESTIMATED (100% capex allocation). Capex dollars spent per $1 of new OCI run-rate revenue added YoY.", fmt="0.00"),
 dict(kid="M07", name="Cumulative AI MW delivered since Jun-25", unit="MW", expr=lambda I, i: f"{I('I45E',i)}", deps=lambda i: [("I45E",i)], note="ESTIMATED; not an installed base.", fmt="0", passthru=True),
 dict(kid="M08", name="Not-yet-commenced DC lease commitments", unit="$B", expr=lambda I, i: f"{I('I51',i)}", deps=lambda i: [("I51",i)], note="Contracted capacity pipeline (partner-built shells).", fmt="0", passthru=True),
 dict(kid="M09", name="RPO increase QoQ", unit="$B", expr=lambda I, i: f"{I('I14',i)}-{I('I14',i-1)}", deps=lambda i: [("I14",i),("I14",i-1)], note="Contract-announcement proxy: +$317B FQ1 FY26 (OpenAI/Stargate era), +$68B FQ2 (Meta etc.), +$85B FQ4 FY26.", fmt="0.0"),
]
readme = [
 "Peer KPI comps - Oracle Corporation (ORCL), Oracle Cloud Infrastructure (OCI) perimeter. Built 2026-10-05 (ET) for the Nebius (NBIS) project. File: /workspace/nebius/peers/oracle/oracle_kpis.xlsx; builder: build_orcl_kpis.py on shared engine /workspace/nebius/peers/peer_kpi_template.py.",
 "TEMPLATE: identical row IDs/definitions/tags to peers/coreweave/crwv_kpis.xlsx (I01-I27, K01-K40) + phase-2 extensions I28-I32 / K41-K45; Oracle memo rows I40+ / M01+.",
 "FISCAL MAPPING: Oracle FY ends 31-May. Fiscal quarters end Aug/Nov/Feb/May, one month BEFORE calendar quarter-ends. Each calendar column uses the Oracle fiscal quarter ending one month earlier (Q1'24 = FQ3 FY24 ended 29-Feb-24 ... Q2'26 = FQ4 FY26 ended 31-May-26). Extra column Q3'26† = FQ1 FY27 (ended 31-Aug-26; latest filed) shown for currency. See Fiscal_Map.",
 "PERIMETER: OCI is NOT a reported segment. I01 = 'Cloud infrastructure' revenue line (IaaS, includes non-AI OCI). Capex, depreciation, PP&E, debt, RPO are company-wide. Segment-allocation estimates are tagged ESTIMATED (orange) and labelled: capex attributed to OCI at 100% (P05, upper bound), incremental OCI revenue attributed to delivered AI MW (K43), MW delivered from earnings calls.",
 "MW: Oracle discloses MW DELIVERED per quarter (calls/PR), not installed MW -> K01-K04/K14/K15 n/d; capex/MW uses MW delivered.",
 "TAGS / COLOURS: blue = Reported; black = Derived; orange = Estimated; grey n/d. KPI tag = weakest input.",
 "Multi-peer Comparison tab: /workspace/nebius/peers/comparison.xlsx (built by peers/build_comparison.py).",
]
defs_app = {
 "D1": "No installed MW -> revenue/MW n/d. K43 = YoY incremental OCI revenue x4 / AI MW delivered over the trailing 4 quarters (ESTIMATED allocation).",
 "D2": "No GPU fleet count. Only '>300k GPUs' in the 850MW delivered in FQ1 FY27.",
 "D3": "OCI list prices observed Aug-Oct 2026 only. Reported physical GPU utilization 97.5-97.9% (K41).",
 "D4": "Company capex attributed 100% to OCI (assumption P05) / MW delivered in quarter (calls). Large DC shells are partner-built and leased (I51 $288B not-yet-commenced), so capex is mostly IT equipment.",
 "D5": "Servers & networking 6 yrs since FY25 (5 before, 4 before FY23).",
 "D6": "Company-wide RPO; 12m and 13-36m buckets reported -> 24m share Estimated.",
 "D7": "No >=10% customer at fiscal-year level.",
 "D8": "Investment-grade unsecured notes; annual note-level table. Quarterly cost via interest expense / average carrying debt.",
}
mapping = [(QS[i], QEND[i], FQ[i].split(" (")[0], QEND[i], ("-1 month (fiscal quarter ends one month before " + ["Dec-23","Mar-24","Jun-24","Sep-24","Dec-24","Mar-25","Jun-25","Sep-25","Dec-25","Mar-26","Jun-26","Sep-26"][i] + ")"),
            "opening-balance column" if i == 0 else ("EXTRA column: latest filed quarter (FQ1 FY27)" if i == 11 else "")) for i in range(N)]
notes = [
 ("Gap","OCI is not a segment: no OCI capex, D&A, PP&E, EBITDA, cost of revenue, or customer data. All OCI-perimeter capex/MW figures are allocation ESTIMATES.","K05, K10, K13, K43, M06 Estimated; K44/K45 n/d.","Oracle analyst-day disclosures; NVIDIA/AMD allocation data."),
 ("Gap","No installed MW or GPU fleet count - only MW delivered per quarter (calls, partly lower bounds/approximations) and '>300k GPUs' for FQ1 FY27.","Revenue/MW, revenue/GPU, implied utilization n/d. K43 relies on call datapoints.","Site-level (Abilene 1.2GW, Stargate sites) press + partner filings (Crusoe, Blue Owl, Vantage)."),
 ("Gap","Debt tables only annual (10-K); quarterly via carrying balances.","K28-K35 annual snapshots.",""),
 ("Gap","RPO is company-wide; no OCI split, no customer concentration. OpenAI ~$300B is media-reported only.","Backlog concentration n/d.",""),
 ("Gap","List prices only Aug-Oct 2026.","K08/K09 n/d historically.","Wayback snapshots of oracle.com/cloud/price-list."),
 ("Caveat","Calendar mapping: Oracle quarters end 1 month before calendar quarters (Q2'26 column = Mar-May 2026).","Comparisons vs CRWV/IREN/Nebius Q2'26 are 1 month earlier.",""),
 ("Caveat","FQ3/FQ4 FY24 OCI revenue derived from FY25 growth rates (rounding +/-0.5%).","",""),
 ("Note","Contract announcements (sourced): FQ3 FY25 PR - multi-billion contracts with OpenAI, xAI, Meta, NVIDIA, AMD; Stargate first contract expected (Mar-25); FQ1 FY26 PR - four multi-billion contracts with three customers, RPO +$317B; FQ2 FY26 - Meta and others (RPO +$68B); Abilene 618MW (75%) delivered by Aug-26 incl. 131k GPUs in FQ1 FY27.","",""),
 ("Note","Off-balance-sheet: $288B not-yet-commenced DC lease commitments (Aug-26) + $19B unconditional purchase commitments for cloud infrastructure assets entered after May-26 (10-K FY26).","Economic leverage far above $125B borrowings.",""),
 ("Note","FQ1 FY27: capex $28.5B in one quarter; ~$19.9B ATM equity raised; OCI $7.4B (+121%); FY27 revenue guide >= $90B (call).","",""),
]
CMP = ["K01","K02","K03","K05","K06","K07","K09","K10","K12","K13","K14","K15","K16","K17","K20","K21","K22","K24","K26","K27","K28","K30","K31","K36","K41","K42","K43","K44","K45"]
if __name__ == "__main__":
    # placeholder tokens for memo SUM ranges
    import peer_kpi_template as T
    _orig = T.build
    build(P, readme, defs_app, notes, mapping, ovr, CMP, debt=debt, memo_kpis=memo)
