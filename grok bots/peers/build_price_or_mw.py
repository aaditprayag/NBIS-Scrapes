#!/usr/bin/env python3
"""Build peers/price_or_mw.xlsx - PRICE-OR-MW TEST (Nebius Peer KPIs, 2026-10-05).
Tabs: README, Implied_Price, MW_Paths, Revenue_Sources, Nebius_Check, Sources.
Also patches comparison.xlsx with K02b (billed-IT) and $14.5/$17.1 alignment note.
Run: python3 build_price_or_mw.py && python3 ../model/recalc.py price_or_mw.xlsx && cp /tmp/rc/price_or_mw.xlsx .
Then: python3 build_price_or_mw.py --charts
"""
import os as _os, sys as _sys
if not _os.environ.get("FORCE_STALE_PRICE_OR_MW"):
    _sys.exit("build_price_or_mw.py is STALE (2026-10-08): price_or_mw.xlsx is now produced by peers/patch_price_or_mw.py "
              "(tranche FY27 locks + ex-Microsoft benchmark). Running this would overwrite it and may patch comparison.xlsx.")

import sys, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.chart import BarChart, Reference

ROOT = "/workspace/nebius"
OUT = f"{ROOT}/peers/price_or_mw.xlsx"
BOLD = Font(bold=True); BLUE = Font(color="0000CC"); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); WARN = PatternFill("solid", fgColor="FCE4D6")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# ---------- Nebius Capacity billing / energized (copied, not edited) ----------
cap = openpyxl.load_workbook(f"{ROOT}/capacity/nebius_capacity.xlsx", data_only=True)["Quarterly MW"]
def caprow(label):
    for r in range(1, cap.max_row + 1):
        if cap.cell(r, 1).value == "COMPANY" and cap.cell(r, 2).value == label:
            return r
    raise KeyError(label)
# Cols: M=Q1'26 ... T=Q4'27 (5=Q1'24 so M=13)
def qvals(r, cols):
    return [cap.cell(r, c).value for c in cols]

R_BILL = caprow("Billing MW (IT) - base: Vineland billed tranches + other sites lagged 1Q")
R_EN_L = caprow("Energized MW low - IT-normalized")
R_EN_B = caprow("Energized MW base - IT-normalized")
R_EN_H = caprow("Energized MW high - IT-normalized")
# FY27 = Q1'27..Q4'27 = cols 17-20
NB_BILL_FY27 = qvals(R_BILL, range(17, 21))  # quarter-end
NB_BILL_Q226 = cap.cell(R_BILL, 14).value  # Q2'26
NB_BILL_Q126 = cap.cell(R_BILL, 13).value
NB_BILL_Q426 = cap.cell(R_BILL, 16).value  # YE26 billing
# Model's FY27 avg billing = quarter averages of energized lagged 1q (Judge_Slide 474.96)
# Reconstruct: energized base IT cols O..T = Q3'26..Q4'27; billing lag 1q from prior quarter end
EN_B = qvals(R_EN_B, range(13, 21))  # Q1'26 .. Q4'27
EN_H = qvals(R_EN_H, range(13, 21))
EN_L = qvals(R_EN_L, range(13, 21))
# Model billing = energized[t-1] as quarter-end billing; avg within quarter = avg of open/close
def lag_avg(en):
    # en[0]=Q1'26 ... en[7]=Q4'27; FY27 billing QoQ ends = en[Q4'26..Q3'27] as open, en[Q1'27..Q4'27] as close? 
    # Judge: "energized lagged 1 qtr, quarter averages"
    # Q1'27 billing avg = avg(en[Q4'26], en[Q1'27]) where en[Q4'26] is YE26 energized = billing open
    # Indices: 0=Q1'26,1=Q2'26,2=Q3'26,3=Q4'26,4=Q1'27,5=Q2'27,6=Q3'27,7=Q4'27
    out = []
    for i in range(4, 8):  # FY27 quarters
        open_ = en[i - 1]; close = en[i]
        out.append((open_ + close) / 2)
    return out
NB_MODEL_BILL_AVG_Q = lag_avg(EN_B)
NB_MODEL_BILL_AVG = sum(NB_MODEL_BILL_AVG_Q) / 4
NB_MODEL_BILL_AVG_H = sum(lag_avg(EN_H)) / 4
NB_MODEL_BILL_AVG_L = sum(lag_avg(EN_L)) / 4
NB_CAP_BILL_QEND_AVG = sum(NB_BILL_FY27) / 4

# Nebius AI cloud Q2'26 revenue
NB_REV_Q226 = 574.9  # from filings history R

wb = openpyxl.Workbook()

# ===================== Sources =====================
src = wb.active; src.title = "Sources"
hdrs = ["ID", "What", "URL / local path"]
for i, h in enumerate(hdrs, 1):
    c = src.cell(1, i, h); c.font = BOLD; c.fill = HDR
SRC = [
("S1", "Nebius FY27 consensus revenue $12.30B (21 analysts, range $8.81-17.99); MarketScreener 12.302 agrees. As of 2026-10-04.",
 "https://finance.yahoo.com/quote/NBIS/analysis/ ; https://www.marketscreener.com/quote/stock/NEBIUS-GROUP-8037501/finances/ ; skeptic/01_priced_in_v3.xlsx"),
("S2", "Nebius FY27 avg billing IT MW ~475 = model: Capacity v2 energized IT base lagged 1 qtr, quarter averages (Judge_Slide). Capacity 'Billing MW (IT) - base' quarter-end FY27 avg = 593.9 (memo).",
 "model/nebius_model.xlsx Judge_Slide row 16/25; capacity/nebius_capacity.xlsx Quarterly MW rows 136/147"),
("S3", "Nebius YE26 ARR guide $7-9B (Dec AI cloud x 12, reaffirmed Aug 12 2026); YE26 connected 800MW-1GW.",
 "Q2'26 letter https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm"),
("S4", "CoreWeave FY27 consensus revenue $26.26B / EBITDA $16.64B (Yahoo + MarketScreener, 40 analysts rev). As of 2026-10-04. (Later aggregators ~$26.8-28.0B - we use team's Oct-4 pack.)",
 "skeptic/01_priced_in_v3.xlsx ; https://finance.yahoo.com/quote/CRWV/analysis/ ; https://ca.marketscreener.com/quote/stock/COREWEAVE-INC-185380761/calendar/"),
("S5", "CoreWeave active power: YE25 >850 MW (R); Q1'26 >1 GW (R); Q2'26 1.5 GW (R); contracted ~3.7 GW Q2 / ~4.2 GW by Aug 11 (R). Q1'26 call: plan 1.7 GW YE26 (R, CNBC). Q2'26 call coverage: exit 2026 >1.85 GW active, $18-19B ARR; YE27 ARR >$30B reiterated (P, secondary).",
 "Q2'26 PR https://www.sec.gov/Archives/edgar/data/1769628/000176962826000366/ ; convert deck comps_raw/crwv_0917__crwv2026convertpresentat.htm.txt ; CNBC 2026-05-07; Tickeron Q2 summary"),
("S6", "CoreWeave NVIDIA partnership: 'Build and deliver more than 5GW of cloud infrastructure'; CEO: 'well on our way to more than 8 GW by 2030'.",
 "comps_raw/crwv_0917__crwv2026convertpresentat.htm.txt ; Q1'26 PR"),
("S7", "IREN: cumulative IT delivery target 0.3 GW (2026) / 0.8 GW (2027) (R). Also 480 MW gross YE26 / 1.2 GW gross YE27. YE26 ARR target >$4bn (85% under contract Jul-20); operating ARR $1bn as of Aug 26 2026.",
 "FY26 PR https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm ; Jul-20 8-K https://www.sec.gov/Archives/edgar/data/1878848/000114036126028871/ef20078253_ex99-1.htm"),
("S8", "IREN 3-yr contracts >$20M revenue per MW (IT); discussions ~$25M per MW (IT). Microsoft deal $9.7B / 5 yr / 200 IT MW = $9.7M/IT MW-yr.",
 "FY26 PR (S7); Nov-3 2025 PR"),
("S9", "IREN AI Cloud GAAP revenue: no FY27 Street AI-only consensus found on box. Redburn initiates Neutral, forecasts $2.8B fiscal FY27 total company revenue (P, secondary; fiscal YE Jun-30).",
 "https://blockspace.media/insight/iren-neutral-40-target-revenue-forecast-2029/"),
("S10", "Oracle FQ1 FY27 (qtr ended Aug 31 2026): Cloud Infrastructure revenue $7.388B (+121%); delivered 850 MW additional datacenter capacity; >300k GPUs to AI Cloud customers. FY27 total revenue guide ≥$90B. RPO $664B.",
 "https://www.sec.gov/Archives/edgar/data/1341439/... orcl-ex99_1.htm ; comps_raw/orcl_q1pr__orcl-ex99_1.htm.txt"),
("S11", "Oracle cumulative AI MW delivered since 1-Jun-2025: ~2,050 MW by Aug-26 (E; peers/oracle/oracle_kpis.xlsx I45E). MW basis unlabeled; 353 GPUs/MW vs GB200 538/IT suggests gross/facility. Includes non-AI OCI in Cloud Infrastructure revenue.",
 "peers/oracle/oracle_kpis.xlsx Inputs I09N/I45E"),
("S12", "Nebius-Microsoft ACV: $17.4B / 5 yr = $3.48B/yr. On stated 300 MW = $11.60/MW-yr; on GPU-derived 203.7 IT MW (100k/491) = $17.08; on Capacity gross-read ~240 IT MW = ~$14.5.",
 "parts/contract_economics.xlsx ; peers/build_comparison.py Contracts note; capacity Vineland"),
("S13", "Nebius Q2'26 AI cloud revenue $574.9M (R); energized MW 171.5/173 (Capacity); billing IT MW Q1/Q2'26 = 165.5/186.8 (Capacity row 147).",
 "filings/nebius_kpi_history.xlsx ; capacity/nebius_capacity.xlsx",
("S14", "YE26 HIGH billing IT MW = 284.2 (Judge_Slide B35; Capacity v2 high energized IT lagged 1q, Vineland billed tranches).",
 "model/nebius_model.xlsx Judge_Slide; capacity/owned_campus_power_path.md", "Derived/Estimated")),
]
for i, (a, b, c) in enumerate(SRC, 2):
    src.cell(i, 1, a); src.cell(i, 2, b).alignment = WRAP; src.cell(i, 3, c).alignment = WRAP
src.column_dimensions['A'].width = 6; src.column_dimensions['B'].width = 110; src.column_dimensions['C'].width = 80

# ===================== Revenue_Sources =====================
rs = wb.create_sheet("Revenue_Sources", 0)
rs.column_dimensions['A'].width = 28
for c in range(2, 10): rs.column_dimensions[CL(c)].width = 14
rs['A1'] = "FY27 (or nearest run-rate) revenue inputs. Blue = hard-coded. Tags: R reported, P press/secondary, D derived, E estimated, n-d not disclosed."
rs['A1'].font = BOLD
for i, h in enumerate(["Company", "Metric", "Value $M", "Tag", "Period / basis", "Source", "Notes"], 1):
    c = rs.cell(3, i, h); c.font = BOLD; c.fill = HDR; c.alignment = WRAP
ROWS_REV = [
("Nebius", "Street FY27 consensus revenue", 12300, "E", "CY2027", "S1", "Yahoo/MarketScreener mean, 21 analysts, as-of 2026-10-04"),
("Nebius", "YE26 ARR guide - low", 7000, "R", "YE26 run-rate", "S3", "Dec AI cloud x 12"),
("Nebius", "YE26 ARR guide - high", 9000, "R", "YE26 run-rate", "S3", ""),
("CoreWeave", "Street FY27 consensus revenue", 26260, "E", "CY2027", "S4", "Team pack Oct-4; later aggregators 26.8-28.0"),
("CoreWeave", "YE26 ARR guide (call coverage)", 18500, "P", "YE26 run-rate mid", "S5", "Q2'26 call coverage $18-19B ARR exit 2026"),
("CoreWeave", "YE27 ARR guide (mgmt reiterated)", 30000, "R", "YE27 run-rate floor", "S5", "'annualized revenue should exceed $30B by end of 2027' (Q1'26, reiterated)"),
("IREN", "YE26 AI Cloud ARR target (operating)", 4000, "R", "YE26 run-rate", "S7", ">$4bn ARR targeted operational by Dec 31 2026; 85% contracted Jul-20"),
("IREN", "Operating ARR (as of Aug 26 2026)", 1000, "R", "Aug-26 run-rate", "S7", ""),
("IREN", "FY27 AI GAAP revenue (Street)", "n-d", "n-d", "FY27", "S9", "No AI-only consensus on box. Redburn $2.8B = total company fiscal FY27 (YE Jun-30), P"),
("IREN", "Memo: YE27 ARR if 800 IT MW at $20M (own 3-yr pricing)", 16000, "E", "YE27 run-rate", "S7/S8", "800 x 20; not a company guide"),
("IREN", "Memo: YE27 ARR if 800 IT MW at $13.3M (same as YE26 ARR/MW)", 10667, "E", "YE27 run-rate", "S7", "800 x (4000/300)"),
("Oracle", "FQ1 FY27 Cloud Infrastructure (IaaS) revenue", 7388, "R", "Qtr ended Aug-31-26", "S10", "Includes non-AI OCI"),
("Oracle", "FQ1 FY27 Cloud Infra annualized (x4)", 29552, "D", "Run-rate", "S10", "Not a FY27 guide; growth still high"),
("Oracle", "FY27 total company revenue guide", 90000, "R", "FY27 (YE May-31)", "S10", "'at least $90 billion' - whole company, not OCI/AI"),
("Oracle", "Cloud Infra FY26 (full year)", 18101, "R", "FY26", "S10", ""),
]
for j, row in enumerate(ROWS_REV, 4):
    for k, v in enumerate(row, 1):
        cell = rs.cell(j, k, v)
        if k == 3 and isinstance(v, (int, float)): cell.font = BLUE; cell.number_format = "#,##0"
        cell.alignment = WRAP
# Named refs for formulas - store row numbers
REV = {("Nebius", "Street"): 4, ("Nebius", "ARR_L"): 5, ("Nebius", "ARR_H"): 6,
       ("CRWV", "Street"): 7, ("CRWV", "ARR26"): 8, ("CRWV", "ARR27"): 9,
       ("IREN", "ARR26"): 10, ("IREN", "ARR_now"): 11, ("IREN", "ARR27_20"): 13, ("IREN", "ARR27_133"): 14,
       ("ORCL", "Q1"): 15, ("ORCL", "Q1x4"): 16, ("ORCL", "FY26"): 18}

# ===================== MW_Paths =====================
mp = wb.create_sheet("MW_Paths", 0)
mp.column_dimensions['A'].width = 55
for c in range(2, 14): mp.column_dimensions[CL(c)].width = 11
mp['A1'] = "Billed / active IT MW paths. Prefer billed/billing over installed. Blue = input. CoreWeave 'active power' ≈ billed. Oracle = cumulative AI capacity delivered (not a clean billed-IT stock)."
mp['A1'].font = BOLD; mp['A1'].alignment = WRAP; mp.row_dimensions[1].height = 35

# --- Nebius ---
r = 3
mp.cell(r, 1, "NEBIUS - Capacity Billing MW (IT) base row + model lag method").font = BOLD
r = 4
for i, h in enumerate(["Item", "Q1'26", "Q2'26", "Q3'26", "Q4'26", "Q1'27", "Q2'27", "Q3'27", "Q4'27", "FY27 avg", "Tag", "Source"], 1):
    c = mp.cell(r, i, h); c.font = BOLD; c.fill = HDR
r = 5
mp.cell(r, 1, "Capacity 'Billing MW (IT) - base' quarter-end")
for k, v in enumerate(qvals(R_BILL, range(13, 21)), 2):
    c = mp.cell(r, k, round(v, 1) if isinstance(v, float) else v); c.font = BLUE; c.number_format = "#,##0.0"
mp.cell(r, 10, "=AVERAGE(F5:I5)"); mp.cell(r, 10).number_format = "#,##0.0"
mp.cell(r, 11, "E"); mp.cell(r, 12, f"S2 row {R_BILL}")
r = 6
mp.cell(r, 1, "Model billing (energized IT base lagged 1q, quarter avg) - STREET DENOMINATOR")
# Put FY27 quarterly avgs in F-I; leave B-E blank for hist
for k, v in enumerate(NB_MODEL_BILL_AVG_Q, 6):
    c = mp.cell(r, k, round(v, 2)); c.font = BLUE; c.number_format = "#,##0.0"
mp.cell(r, 10, "=AVERAGE(F6:I6)"); mp.cell(r, 10).number_format = "#,##0.0"; mp.cell(r, 10).fill = KEY
mp.cell(r, 11, "D/E"); mp.cell(r, 12, "S2 Judge_Slide / Capacity row 136 lagged")
NB_AVG_CELL = "$J$6"
r = 7
mp.cell(r, 1, "Same method - LOW energized IT path")
for k, v in enumerate(lag_avg(EN_L), 6):
    c = mp.cell(r, k, round(v, 2)); c.font = BLUE; c.number_format = "#,##0.0"
mp.cell(r, 10, "=AVERAGE(F7:I7)"); mp.cell(r, 10).number_format = "#,##0.0"; mp.cell(r, 11, "D/E")
r = 8
mp.cell(r, 1, "Same method - HIGH energized IT path")
for k, v in enumerate(lag_avg(EN_H), 6):
    c = mp.cell(r, k, round(v, 2)); c.font = BLUE; c.number_format = "#,##0.0"
mp.cell(r, 10, "=AVERAGE(F8:I8)"); mp.cell(r, 10).number_format = "#,##0.0"; mp.cell(r, 11, "D/E")
r = 9
mp.cell(r, 1, "YE26 billing IT MW (Capacity base Q4'26)")
mp.cell(r, 5, round(NB_BILL_Q426, 1)); mp.cell(r, 5).font = BLUE; mp.cell(r, 11, "E"); mp.cell(r, 12, "S2")
YE26_BILL = "$E$9"
r = 10
mp.cell(r, 1, "YE26 billing IT MW (Capacity HIGH / Judge_Slide B35)")
# 284.2 = Judge_Slide B35: Capacity v2 high energized IT lagged 1q to billing (Vineland billed tranches)
YE26_BILL_HIGH = 284.2
mp.cell(r, 5, YE26_BILL_HIGH); mp.cell(r, 5).font = BLUE; mp.cell(r, 5).number_format = "#,##0.0"
mp.cell(r, 5).fill = KEY
mp.cell(r, 11, "D/E"); mp.cell(r, 12, "S14 / Judge_Slide B35; Capacity HIGH billing YE26")
YE26_BILL_H = "$E$10"

# --- CoreWeave ---
r = 11
mp.cell(r, 1, "COREWEAVE - active power (company metric; ≈ billed / revenue-generating)").font = BOLD
r = 12
for i, h in enumerate(["Item", "YE25", "Q1'26", "Q2'26", "YE26E", "YE27E", "FY27 avg (linear)", "Tag", "Source / method"], 1):
    c = mp.cell(r, i, h); c.font = BOLD; c.fill = HDR
r = 13
mp.cell(r, 1, "BASE path")
for k, v in enumerate([850, 1000, 1500, 1850, 2800], 2):
    c = mp.cell(r, k, v); c.font = BLUE; c.number_format = "#,##0"
mp.cell(r, 7, "=(E13+F13)/2"); mp.cell(r, 7).number_format = "#,##0.0"; mp.cell(r, 7).fill = KEY
mp.cell(r, 8, "R/P/E"); mp.cell(r, 9, "YE25/Q1/Q2 R (S5); YE26 1850 P (Q2 call coverage; Q1 guide was 1700 R); YE27 2800 E (interp: contracted 4.2 GW x ~67% active by YE27, between 1.85 and NVIDIA 5GW multi-year)")
r = 14
mp.cell(r, 1, "LOW path (slower YE27 conversion)")
for k, v in enumerate([850, 1000, 1500, 1700, 2200], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 7, "=(E14+F14)/2"); mp.cell(r, 7).number_format = "#,##0.0"
mp.cell(r, 8, "R/E"); mp.cell(r, 9, "YE26 = Q1 guide 1.7 GW R; YE27 2.2 GW E (modest add)")
r = 15
mp.cell(r, 1, "HIGH path")
for k, v in enumerate([850, 1000, 1500, 2000, 3500], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 7, "=(E15+F15)/2"); mp.cell(r, 7).number_format = "#,##0.0"
mp.cell(r, 8, "E"); mp.cell(r, 9, "YE26 2.0 GW E; YE27 3.5 GW E (faster toward 5GW NVIDIA buildout)")
r = 16
mp.cell(r, 1, "Memo: YE27 MW if YE27 ARR $30B at Q2'26 rev/avg MW $8.24")
mp.cell(r, 6, f"=Revenue_Sources!C{REV[('CRWV','ARR27')]}/8.24"); mp.cell(r, 6).number_format = "#,##0"
mp.cell(r, 8, "E"); mp.cell(r, 9, "Implies ~3,640 MW if unit revenue stays at Q2'26 diluted level - upper bound on MW / lower bound on $/MW")
CRWV_AVG = "$G$13"; CRWV_AVG_L = "$G$14"; CRWV_AVG_H = "$G$15"

# --- IREN ---
r = 18
mp.cell(r, 1, "IREN AI Cloud - IT MW cumulative delivery targets (company; billed ≈ delivered after acceptance lag)").font = BOLD
r = 19
for i, h in enumerate(["Item", "Jun-26 reported", "YE26 target", "YE27 target", "FY27 avg (linear YE26→YE27)", "Tag", "Source"], 1):
    c = mp.cell(r, i, h); c.font = BOLD; c.fill = HDR
r = 20
mp.cell(r, 1, "IT MW (company target)")
for k, v in enumerate([40, 300, 800], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 5, "=(C20+D20)/2"); mp.cell(r, 5).number_format = "#,##0.0"; mp.cell(r, 5).fill = KEY
mp.cell(r, 6, "R/D"); mp.cell(r, 7, "S7: 0.3 GW IT 2026 / 0.8 GW IT 2027; Jun-26 ~40 MW operating (R, basis unstated)")
r = 21
mp.cell(r, 1, "Gross MW (company target, memo)")
for k, v in enumerate(["n-d", 480, 1200], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 5, "=(C21+D21)/2"); mp.cell(r, 6, "R/D"); mp.cell(r, 7, "S7 Jul-20: 480 MW gross YE26 / 1.2 GW gross YE27")
IREN_AVG = "$E$20"; IREN_YE26 = "$C$20"

# --- Oracle ---
r = 23
mp.cell(r, 1, "ORACLE - cumulative AI capacity DELIVERED (not billed-IT stock; MW basis unlabeled, likely gross)").font = BOLD
r = 24
for i, h in enumerate(["Item", "FQ4'26 (May)", "FQ1'27 (Aug)", "FY27E EoP (May'27)", "FY27 avg (linear)", "Tag", "Source"], 1):
    c = mp.cell(r, i, h); c.font = BOLD; c.fill = HDR
r = 25
mp.cell(r, 1, "Cumulative AI MW delivered since Jun-2025 (BASE)")
for k, v in enumerate([1200, 2050, 4500], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 5, "=(C25+D25)/2"); mp.cell(r, 5).number_format = "#,##0.0"; mp.cell(r, 5).fill = KEY
mp.cell(r, 6, "E"); mp.cell(r, 7, "S11: 1200 at May-26, 2050 at Aug-26 (incl. 850 delivered FQ1). YE27 4500 E = +~2.5 GW over remaining 3 qtrs of FY27 at ~similar pace to FQ1's 850 (decelerating)")
r = 26
mp.cell(r, 1, "LOW (slower delivery)")
for k, v in enumerate([1200, 2050, 3200], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 5, "=(C26+D26)/2"); mp.cell(r, 5).number_format = "#,##0.0"; mp.cell(r, 6, "E")
r = 27
mp.cell(r, 1, "HIGH")
for k, v in enumerate([1200, 2050, 6000], 2):
    c = mp.cell(r, k, v); c.font = BLUE
mp.cell(r, 5, "=(C27+D27)/2"); mp.cell(r, 5).number_format = "#,##0.0"; mp.cell(r, 6, "E")
r = 28
mp.cell(r, 1, "Memo: if 850 MW/qtr continues x4 through FY27 from May-26 base")
mp.cell(r, 4, "=1200+850*4"); mp.cell(r, 5, "=(1200+(1200+850*4))/2"); mp.cell(r, 6, "E")
ORCL_AVG = "$E$25"; ORCL_AVG_L = "$E$26"; ORCL_AVG_H = "$E$27"
ORCL_AUG = "$C$25"

# ===================== Implied_Price =====================
ip = wb.create_sheet("Implied_Price", 0)
ip.column_dimensions['A'].width = 42
for c in range(2, 12): ip.column_dimensions[CL(c)].width = 13
ip['A1'] = "PRICE-OR-MW TEST: FY27 revenue (or run-rate guide) ÷ average billed IT MW → implied $M per billed IT MW-yr. Compare to Nebius Street ~$26M."
ip['A1'].font = BOLD; ip['A1'].alignment = WRAP; ip.row_dimensions[1].height = 30
ip['A2'] = "Question: do peers land at ~$13-18M (Nebius alone at an unsupported premium) or are peers just as rich?"
ip['A2'].font = Font(bold=True, italic=True)

r = 4
for i, h in enumerate(["Company", "Revenue $M", "Rev tag", "Avg billed IT MW", "MW tag", "Implied $M/MW", "vs Nebius Street $26M", "Premium/(discount)", "Verdict vs $13-18 band", "Notes"], 1):
    c = ip.cell(r, i, h); c.font = BOLD; c.fill = HDR; c.alignment = WRAP
ip.row_dimensions[4].height = 40

# Row helpers: revenue cell, mw cell, tags
def add_row(r, co, rev_ref, rev_tag, mw_ref, mw_tag, notes, key=False):
    ip.cell(r, 1, co)
    ip.cell(r, 2, f"=Revenue_Sources!C{rev_ref}" if isinstance(rev_ref, int) else rev_ref)
    ip.cell(r, 2).number_format = "#,##0"
    ip.cell(r, 3, rev_tag)
    ip.cell(r, 4, f"=MW_Paths!{mw_ref}" if not str(mw_ref).startswith("=") else mw_ref)
    ip.cell(r, 4).number_format = "#,##0.0"
    ip.cell(r, 5, mw_tag)
    ip.cell(r, 6, f"=B{r}/D{r}"); ip.cell(r, 6).number_format = "0.00"
    ip.cell(r, 7, f"=F{r}-Nebius_Check!$B$5"); ip.cell(r, 7).number_format = "0.00"
    ip.cell(r, 8, f"=F{r}/Nebius_Check!$B$5-1"); ip.cell(r, 8).number_format = "0.0%"
    ip.cell(r, 9, f'=IF(AND(F{r}>=13,F{r}<=18),"IN $13-18 band",IF(F{r}>18,"ABOVE (rich like/above Nebius Street)","BELOW $13-18"))')
    ip.cell(r, 10, notes).alignment = WRAP
    if key:
        for c in range(1, 10): ip.cell(r, c).fill = KEY

# Primary headline rows
add_row(5, "Nebius STREET (consensus / model avg billing)", REV[("Nebius","Street")], "E", "J6", "D/E",
        "THE BENCHMARK: $12.3B / ~475 = ~$26M. Verified vs Judge_Slide.", key=True)
add_row(6, "Nebius YE26 ARR guide / YE26 billing (low)", REV[("Nebius","ARR_L")], "R", "E9", "E",
        "Guide test: $7B / YE26 BASE billing ~257.7 (Capacity Q4'26)")
add_row(7, "Nebius YE26 ARR guide / YE26 billing (high)", REV[("Nebius","ARR_H")], "R", "E10", "E",
        "CORRECTION: $9B / YE26 HIGH billing 284.2 (Judge_Slide B35) → ~$31.7M in $25-32 band (was wrongly on E9=257.7)")
add_row(8, "CoreWeave STREET / BASE MW path", REV[("CRWV","Street")], "E", "G13", "E",
        "Active power path YE26 1.85 GW P → YE27 2.8 GW E; avg 2.325 GW", key=True)
add_row(9, "CoreWeave STREET / LOW MW path", REV[("CRWV","Street")], "E", "G14", "E",
        "YE26 1.7 GW R (Q1 guide) → YE27 2.2 GW E")
add_row(10, "CoreWeave STREET / HIGH MW path", REV[("CRWV","Street")], "E", "G15", "E",
        "Faster MW → lower $/MW")
add_row(11, "CoreWeave YE27 ARR guide $30B / BASE YE27 EoP MW", REV[("CRWV","ARR27")], "R", "F13", "E",
        "Exit run-rate test (not FY27 average revenue)")
add_row(12, "IREN YE26 ARR $4B / YE26 300 IT MW", REV[("IREN","ARR26")], "R", "C20", "R",
        "Cleanest IREN like-for-like: mgmt ARR / mgmt IT MW target", key=True)
add_row(13, "IREN YE27 ARR at $20M pricing / YE27 800 IT MW", REV[("IREN","ARR27_20")], "E", "D20", "R",
        "Uses IREN's own >$20M/IT MW 3-yr pricing on full YE27 capacity")
add_row(14, "IREN YE27 ARR at YE26 $/MW / YE27 800 IT MW", REV[("IREN","ARR27_133")], "E", "D20", "R",
        "Same unit price as YE26 ARR/MW")
add_row(15, "IREN FY27 recognized (avg of YE26 ARR & YE27@$20M) / avg MW",
        f"=(Revenue_Sources!C{REV[('IREN','ARR26')]}+Revenue_Sources!C{REV[('IREN','ARR27_20')]})/2", "E",
        "E20", "D",
        "Proxy for calendar-2027 recognized AI revenue; no Street AI-only consensus on box", key=True)
add_row(16, "Oracle FQ1'27 Cloud Infra annualized / cum. AI MW at Aug-26",
        REV[("ORCL","Q1x4")], "D", "C25", "E",
        "Numerator includes non-AI OCI; denominator = cumulative delivered since Jun-25 (not stock), likely gross MW", key=True)
add_row(17, "Oracle FQ1'27 annualized / FY27 avg cum. delivered (BASE)",
        REV[("ORCL","Q1x4")], "D", "E25", "E",
        "Same numerator on growing delivered base → lower $/MW; still conflates AI vs non-AI")

# Sensitivity block
r = 19
ip.cell(r, 1, "SENSITIVITY - CoreWeave $/MW across MW and revenue").font = BOLD
r = 20
for i, h in enumerate(["", "MW LOW 1,950", "MW BASE 2,325", "MW HIGH 2,750", "MW at $8.24 ARR ($30B→3,640)"], 1):
    c = ip.cell(r, i, h); c.font = BOLD; c.fill = HDR
for j, (lab, rev) in enumerate([("Rev $26.26B (Street)", f"Revenue_Sources!C{REV[('CRWV','Street')]}"),
                                 ("Rev $28.0B (later aggregator)", 28000),
                                 ("YE27 ARR $30B as if FY27 rev", f"Revenue_Sources!C{REV[('CRWV','ARR27')]}")], 21):
    ip.cell(j, 1, lab)
    for k, mw in enumerate([1950, 2325, 2750, f"Revenue_Sources!C{REV[('CRWV','ARR27')]}/8.24"], 2):
        if isinstance(rev, str) and rev.startswith("Revenue"):
            ip.cell(j, k, f"={rev}/{mw}" if isinstance(mw, (int, float)) else f"={rev}/({mw})")
        else:
            ip.cell(j, k, f"={rev}/{mw}" if isinstance(mw, (int, float)) else f"={rev}/({mw})")
        ip.cell(j, k).number_format = "0.00"

r = 25
ip.cell(r, 1, "HEADLINE TABLE (ready to DM)").font = BOLD
r = 26
for i, h in enumerate(["Company", "FY27 rev $B", "Avg billed IT MW", "Implied $/MW", "Tag", "vs Nebius Street $26M"], 1):
    c = ip.cell(r, i, h); c.font = BOLD; c.fill = KEY
HEAD = [
(27, "Nebius Street", "F5", "D5", "F5", "E (cons) / D/E (MW)", "benchmark"),
(28, "CoreWeave", "B8/1000", "D8", "F8", "E / E", None),
(29, "IREN (ARR path)", "B15/1000", "D15", "F15", "E / D", None),
(30, "IREN (YE26 ARR/MW)", "B12/1000", "D12", "F12", "R / R", None),
(31, "Oracle OCI (run-rate / cum MW)", "B16/1000", "D16", "F16", "D / E", None),
]
for r, co, rev, mw, imp, tag, _ in HEAD:
    ip.cell(r, 1, co)
    ip.cell(r, 2, f"={rev}" if "B" in str(rev) or "F" in str(rev) else rev)
    if "B" in str(rev) or rev.startswith("F"): ip.cell(r, 2, f"={rev.replace('/1000','')}/1000")
    ip.cell(r, 2).number_format = "0.00"
    ip.cell(r, 3, f"={mw}"); ip.cell(r, 3).number_format = "#,##0"
    ip.cell(r, 4, f"={imp}"); ip.cell(r, 4).number_format = "0.0"
    ip.cell(r, 5, tag)
    ip.cell(r, 6, f"=D{r}-Nebius_Check!$B$5"); ip.cell(r, 6).number_format = "+0.0;-0.0"

r = 33
ip.cell(r, 1, "VERDICT").font = BOLD
ip.cell(r + 1, 1, (
    "Peers come out in the ~$11-16M per billed IT MW band on base paths - NOT as rich as Nebius Street's ~$26M. "
    "CoreWeave Street / base MW ≈ $11-13M; IREN YE26 ARR/IT MW = $13.3M (R/R) and FY27 ARR-path proxy ≈ $14-18M; "
    "Oracle FQ1 Cloud Infra annualized / cum. delivered MW ≈ $14M (with AI/non-AI and gross-MW caveats). "
    "Nebius Street at ~$26M is a ~60-130% premium to this peer band. "
    "Confidence MEDIUM: CoreWeave YE27 MW is Estimated (largest swing); IREN lacks AI GAAP consensus; Oracle MW/revenue perimeter is impure. "
    "Even on CoreWeave LOW MW (higher $/MW), Street still lands ~$13-14M, well below $26M. "
    "YE26 Nebius guide ($7-9B / ~258-284 billing MW) also implies $25-35M - same rich print as Street FY27.")
).alignment = WRAP
ip.merge_cells(start_row=r+1, start_column=1, end_row=r+1, end_column=10)
ip.row_dimensions[r+1].height = 95

# ===================== Nebius_Check =====================
nc = wb.create_sheet("Nebius_Check", 0)
nc.column_dimensions['A'].width = 70; nc.column_dimensions['B'].width = 14; nc.column_dimensions['C'].width = 10; nc.column_dimensions['D'].width = 80
nc['A1'] = "Verify Nebius Street $12.3B / ~475 avg billing MW → ~$26M, and YE26 guide $25-32M. Align $14.5M vs $17.1M Microsoft readings."
nc['A1'].font = BOLD; nc['A1'].alignment = WRAP
for i, h in enumerate(["Check", "Value", "Tag", "Source / note"], 1):
    c = nc.cell(3, i, h); c.font = BOLD; c.fill = HDR
checks = [
("FY27 consensus revenue ($M)", f"=Revenue_Sources!C{REV[('Nebius','Street')]}", "E", "S1"),
("FY27 avg billing IT MW - model method (STREET DENOMINATOR)", "=MW_Paths!J6", "D/E", "S2 - Judge_Slide 474.96"),
("Implied $M per billed IT MW (Street)", "=B4/B5", "D", "Should ≈ 25.9"),
("FY27 avg billing - Capacity quarter-end row (memo)", "=MW_Paths!J5", "E", "593.9; NOT the Street denominator"),
("Implied $M/MW if used Capacity qtr-end avg", "=B4/B7", "D", "Lower (~$20.7) - do not use for Street compare"),
("Judge_Slide printed Street $/MW", 25.8968215362389, "D", "model/nebius_model.xlsx Judge_Slide r16"),
("Delta vs Judge_Slide (should be ~0)", "=B6-B9", "D", ""),
("YE26 ARR low ($7B) / YE26 BASE billing 257.7", f"=Revenue_Sources!C{REV[('Nebius','ARR_L')]}/MW_Paths!E9", "D", "S3 / MW_Paths!E9"),
("YE26 ARR high ($9B) / YE26 BASE billing 257.7 (memo)", f"=Revenue_Sources!C{REV[('Nebius','ARR_H')]}/MW_Paths!E9", "D", "Memo only — headline high uses E10"),
("YE26 ARR high ($9B) / YE26 HIGH billing 284.2 (HEADLINE)", f"=Revenue_Sources!C{REV[('Nebius','ARR_H')]}/MW_Paths!E10", "D", "Judge_Slide B35/B41 ~31.7"),
("YE26 ARR low ($7B) / YE26 HIGH billing 284.2", f"=Revenue_Sources!C{REV[('Nebius','ARR_L')]}/MW_Paths!E10", "D", "Judge_Slide ~24.6 → band $25-32 with headline high"),
]
for j, (a, b, c, d) in enumerate(checks, 4):
    nc.cell(j, 1, a); cell = nc.cell(j, 2, b); cell.number_format = "0.00" if j != 4 and j != 5 and j != 7 else "#,##0.00"
    if j == 4: cell.number_format = "#,##0"
    if j in (5, 7): cell.number_format = "#,##0.0"
    nc.cell(j, 3, c); nc.cell(j, 4, d).alignment = WRAP
nc['B6'].fill = KEY

nc['A16'] = "$14.5M vs $17.1M ALIGNMENT (Microsoft-Nebius deal)".font if False else "$14.5M vs $17.1M ALIGNMENT (Microsoft-Nebius deal)"
nc['A16'].font = BOLD
nc['A17'] = "These are the SAME contract under three MW readings - not three prices. Use all three; do not pick one silently."
align = [
("TCV committed ($M)", 17392.9, "R", "S12"),
("Term (years)", 5, "R", ""),
("ACV = TCV/term ($M/yr)", "=B18/B19", "D", ""),
("Reading A: stated site MW 300 (unlabeled)", 300, "R", "6-K 'up to 300 MW'"),
("→ $M per stated MW-yr", "=B20/B21", "D", "$11.60"),
("Reading B: Capacity gross→IT ~240 MW (300/1.25)", 240, "E", "Capacity Vineland gross read"),
("→ $M per IT MW-yr (gross read)", "=B20/B23", "D", "~$14.5 - memo/slide 'gross read'"),
("Reading C: GPU-derived IT MW = 100,000 / 491", "=100000/491.17", "D", "Capacity GB300 density"),
("→ $M per IT MW-yr (GPU-derived)", "=B20/B25", "D", "$17.08 - contract_economics Normalized; payback_dilution P_new base"),
("Q2'26 AI cloud revenue per avg ENERGIZED MW (K02)", f"={NB_REV_Q226}*4/AVERAGE(171.5,173)", "D/E", "Comparison K02 = 13.35 E"),
("Q2'26 AI cloud revenue per avg BILLING IT MW (K02b)", f"={NB_REV_Q226}*4/AVERAGE({NB_BILL_Q126},{NB_BILL_Q226})", "D/E", "New Comparison K02b"),
]
for j, (a, b, c, d) in enumerate(align, 18):
    nc.cell(j, 1, a); cell = nc.cell(j, 2, b)
    if isinstance(b, float) or (isinstance(b, str) and b.startswith("=")): cell.number_format = "0.00"
    if isinstance(b, (int, float)) and not isinstance(b, bool): cell.font = BLUE
    nc.cell(j, 3, c); nc.cell(j, 4, d).alignment = WRAP
for rr in (22, 24, 26):
    nc.cell(rr, 2).fill = KEY

nc['A30'] = ("Alignment rule for the team: quote Microsoft-Nebius as '$11.6 on stated 300 MW | ~$14.5 on Capacity gross→IT read | $17.1 on GPU-derived 203.7 IT MW'. "
            "K02 (energized) = 13.35; K02b (billed IT) ≈ 13.0. Street FY27 at $26/MW is on BILLED IT MW (~475), not energized.")
nc['A30'].alignment = WRAP; nc.merge_cells('A30:D30'); nc.row_dimensions[30].height = 45

# ===================== README =====================
rd = wb.create_sheet("README", 0)
rd.column_dimensions['A'].width = 140
TXT = [
"price_or_mw.xlsx - Nebius Peer KPIs - PRICE-OR-MW TEST - built 2026-10-05 ET; CORRECTED midday ET (YE26 high ARR uses HIGH billing 284.2). Recalc: 0 errors.",
"",
"QUESTION: FY27 consensus (or mgmt run-rate) ÷ average billed IT MW → implied $/billed IT MW. Do peers land at ~$13-18M (Nebius Street ~$26M = unsupported premium) or are peers just as rich?",
"",
"NEBIUS BENCHMARK (verified): Street FY27 $12.3B (S1) / ~475 avg billing IT MW (S2, model: Capacity energized IT base lagged 1q, quarter averages) = $25.9M. YE26 ARR guide $7-9B / YE26 billing 258-284 MW = $25-35M (Judge_Slide $25-32 on high billing).",
"",
"HEADLINE RESULT: Peers land in ~$11-16M on base paths - NOT as rich as Nebius Street. Nebius Street is a large premium to the peer band. Confidence MEDIUM.",
"  CoreWeave: $26.26B / ~2,325 avg active MW (E) ≈ $11.3M (range ~$9.5-13.5 across MW paths).",
"  IREN: YE26 ARR $4B / 300 IT MW = $13.3M (R/R); FY27 ARR-path proxy ≈ $14-18M (E).",
"  Oracle: FQ1 Cloud Infra annualized $29.6B / ~2,050 cum. delivered MW ≈ $14.4M (D/E; AI/non-AI and gross-MW caveats).",
"",
"$14.5 vs $17.1 ALIGNMENT: same Microsoft contract, three MW readings ($11.6 stated 300 | ~$14.5 gross→IT | $17.1 GPU-derived). See Nebius_Check.",
"K02b: Comparison.xlsx now has K02 (energized avg MW) AND K02b (billed IT avg MW) so the slide can show both.",
"",
"TABS: Implied_Price (headline + verdict) | MW_Paths | Revenue_Sources | Nebius_Check | Sources | README.",
"TAGS: R reported; P press/secondary; D derived; E estimated; n-d not disclosed. Blue = hard-coded input.",
]
for i, t in enumerate(TXT, 1):
    c = rd.cell(i, 1, t); c.alignment = WRAP
    if t.startswith(("QUESTION", "HEADLINE", "NEBIUS", "price_or", "$14.5")): c.font = BOLD

# Order sheets
order = ["README", "Implied_Price", "MW_Paths", "Revenue_Sources", "Nebius_Check", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(OUT)
print("saved", OUT)
print("Nebius model avg billing FY27", round(NB_MODEL_BILL_AVG, 2), "qtr-end avg", round(NB_CAP_BILL_QEND_AVG, 2))
print("K02b", round(NB_REV_Q226 * 4 / ((NB_BILL_Q126 + NB_BILL_Q226) / 2), 2))

# ---------- Patch comparison.xlsx: add K02b + alignment note ----------
cmp_path = f"{ROOT}/peers/comparison.xlsx"
cwb = openpyxl.load_workbook(cmp_path)
# Nebius_Inputs: add billing MW and K02b formula support
wsN = cwb["Nebius_Inputs"]
# Find next empty row
nr = wsN.max_row + 1
# Check if already patched
already = any(str(wsN.cell(r, 1).value or "").startswith("N_BILL") or str(wsN.cell(r, 2).value or "").startswith("Billing IT") for r in range(1, wsN.max_row + 1))
if not already:
    # Find N01 revenue row and MW rows for structure
    n01_row = None; last_n = 0
    for r in range(1, wsN.max_row + 1):
        v = str(wsN.cell(r, 1).value or "")
        if v == "N01": n01_row = r
        if v.startswith("N") and v[1:].isdigit(): last_n = max(last_n, int(v[1:]))
    # Append billing MW inputs
    r = wsN.max_row + 1
    wsN.cell(r, 1, "N_BILL_Q1"); wsN.cell(r, 2, "Billing IT MW Q1'26 (Capacity row 147)"); wsN.cell(r, 3, round(NB_BILL_Q126, 2)); wsN.cell(r, 3).font = BLUE
    wsN.cell(r, 4, "E"); wsN.cell(r, 5, "capacity/nebius_capacity.xlsx"); BILL_Q1_R = r
    r += 1
    wsN.cell(r, 1, "N_BILL_Q2"); wsN.cell(r, 2, "Billing IT MW Q2'26 (Capacity row 147)"); wsN.cell(r, 3, round(NB_BILL_Q226, 2)); wsN.cell(r, 3).font = BLUE
    wsN.cell(r, 4, "E"); wsN.cell(r, 5, "capacity/nebius_capacity.xlsx"); BILL_Q2_R = r
    r += 1
    wsN.cell(r, 1, "NOTE_1417"); wsN.cell(r, 2, "Microsoft ACV alignment: $11.60 on stated 300 MW | ~$14.5 on Capacity gross→IT (~240) | $17.08 on GPU-derived 203.7 IT MW. Same contract, three readings.")
    wsN.cell(r, 2).alignment = WRAP
    # Insert K02b into Comparison after K02
    wsC = cwb["Comparison"]
    # Find K02 row
    k02r = None
    for r in range(1, wsC.max_row + 1):
        if wsC.cell(r, 1).value == "K02": k02r = r; break
    if k02r:
        wsC.insert_rows(k02r + 1)
        r = k02r + 1
        wsC.cell(r, 1, "K02b")
        wsC.cell(r, 2, "Revenue per BILLING IT MW - average MW (Capacity billing path)")
        wsC.cell(r, 3, "$M/MW-yr")
        # Nebius formula
        wsC.cell(r, 4, f"=Nebius_Inputs!$C${n01_row}*4/AVERAGE(Nebius_Inputs!$C${BILL_Q1_R},Nebius_Inputs!$C${BILL_Q2_R})")
        wsC.cell(r, 5, "E")
        wsC.cell(r, 6, "Jun-26")
        # Peers: CoreWeave active≈billed so same as K02; IREN/Oracle n/d for billed
        # Copy CoreWeave K02 value ref - hardcode from K02 row peers
        # Cols: Nebius D/E/F=4/5/6; CRWV G/H/I=7/8/9; IREN J/K/L; Oracle M/N/O; ...
        wsC.cell(r, 7, f"=G{k02r}")  # CRWV same as K02 (active≈billed)
        wsC.cell(r, 8, "Derived")
        wsC.cell(r, 9, "Jun-26")
        for col, tagcol in [(10, 11), (13, 14), (16, 17), (19, 20), (22, 23)]:
            wsC.cell(r, col, "n/d"); wsC.cell(r, tagcol, "n/d")
        wsC.cell(r, 25, f"=G{r}")  # median of available = just CRWV
        wsC.cell(r, 26, "Better")
        wsC.cell(r, 27, "Nebius billed-IT basis (Capacity) vs CoreWeave active (≈billed); other peers n/d billed MW")
        wsC.cell(r, 28, "NEW: billed-IT twin of K02 for slide alignment with Street $/billed MW test")
    # README note
    wsR = cwb["README"]
    rr = wsR.max_row + 2
    wsR.cell(rr, 1, "K02b ADDED 2026-10-05 (Peer KPIs): Revenue per average BILLING IT MW. Nebius uses Capacity billing path (Q1/Q2'26). CoreWeave active≈billed so K02b=K02. Microsoft ACV: quote $11.6 / ~$14.5 / $17.1 (stated / gross-read / GPU-derived) — see Nebius_Inputs NOTE_1417 and peers/price_or_mw.xlsx Nebius_Check.")
    wsR.cell(rr, 1).alignment = WRAP
    cwb.save(cmp_path)
    print("patched comparison.xlsx with K02b")
else:
    print("comparison.xlsx already has billing patch; skipped")

print("done build")
