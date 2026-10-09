# Builds /workspace/nebius/peers/comparison.xlsx (step 5, final): Comparison (Nebius + 6 peers, tags, as-of, better/worse flag),
# Ranking, Nebius_Inputs (live formulas feed Comparison), Contracts_$MWyr, Sites, Chart_Data + native Excel charts, Peer_memos, README.
# Inputs (read-only): peers/coreweave|iren|oracle/*_kpis.xlsx (cached), peers/private/private_peers_values.json,
#   parts/contract_economics.xlsx (GPU Market agent), filings/nebius_kpi_history.xlsx + nebius_filings.xlsx, shared/findings_log.xlsx,
#   capacity/nebius_capacity.xlsx + site_list.xlsx, peers/sites/sites.xlsx.
# Author: Grok Bot for Aadit Prayag, 2026-10-04 ET. v0.9d refresh 2026-10-05 ET: K02b (billed-IT) folded into the build (was hand-patched
#   into the 08:31 pack), Nebius_Inputs locks/cites model v0.9d (FD 358.3M; billing IT MW YE26 258/284, FY27 avg 475/645) read live from
#   model/nebius_model.xlsx Judge_Slide/Checks (+ recalc preview v09d for formula cells) and capacity/nebius_capacity.xlsx row 147.
# 2026-10-08 ET: (1) CRWV K18 read straight from crwv_kpis.xlsx again (pack fixed: K18 = 100/K17); the 10-05 re-derivation guard is replaced by
#   an assert-only consistency check. (2) PRIMARY share basis N_FD = Filings pro forma FD after the Aug-26 exchange/new converts, parsed from
#   filings/post630_capital_update.md (376,648,956); v0.9d model FD 358.3M kept as memo row N_FD_V09D (+5.1% share-base sensitivity).
#   FD lock assertion relaxed (model FD may be 358.3 v0.9d or the Filings PF). Billing MW locks unchanged. (3) Nebius Q3'26 results date
#   labelled "~Nov 10 (est.)" (not announced by the company) in README/Comparison notes.
# 2026-10-08 ~08:45 ET (coordinator follow-up): FY27 avg billed-MW locks moved to the TRANCHE basis (Vineland billed at delivered Microsoft
#   tranches = revenue-engine basis): BASE 536.48 / HIGH 663.8, read live by label from capacity/final_billed_vs_connected.xlsx Final_Table
#   (peers/nebius_locks.py); v0.9d 474.96 / 645.4 kept as superseded memo rows. New forward memo: Street FY27 ex-Microsoft $/billed MW
#   (peer comparator). Share rows restructured to the price-threshold convention from filings/fd_share_bridge.md (debt below the new notes'
#   conversion prices -> 358.56M; all if-converted above -> 376.65M); 358.34M = v0.9d historical memo. "376.6M primary" dropped.
# 2026-10-08 ~10:00 ET (v0.9e): team model v0.9e is the SOURCE OF TRUTH. All Nebius locks are read from model/nebius_model.xlsx via
#   peers/nebius_locks.py model_locks() (Drivers / Judge_Slide / Inputs, by label): FY27 billed MW engine-linked 531.84 / 659.16, Microsoft
#   FY27 revenue $2,643.8M on Vineland 182.33 billed MW-yr, Street $12.3B -> ex-MSFT 27.63 / blended 23.13; FD 358.56 / 376.65 (Inputs r473/474,
#   cross-checked with filings/fd_share_bridge.md). Capacity tranche locks (536.48 / 663.8) and the v0.9d 475 / 645 superseded rows are
#   retired from Nebius_Inputs (one README history line only). YE26 257.7 / 284.2 unchanged (Judge_Slide B34/B35).
# 2026-10-09 ~11:10 ET: N1 (Microsoft) shown only as $14.49M per IT MW [E, house] = model Contracts!E12 (17,392.9 / (5 x 240 IT est.)); 11.60 / 17.08 kept as HISTORY notes only.
# 2026-10-09 ~11:05 ET: N1o (Microsoft incl. option, 12.93 = 19,400/5/300) retired to HISTORY: 300 MW is gross site size, not stated contract MW.
# 2026-10-08 ~23:15 ET: Abilene 618 MW [D on R] propagated via peers/sites/sites.xlsx; Sites rows asserted against peers/sites/site_data.py appendices.
import json, statistics, openpyxl, re
import nebius_locks as NL
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

W = "/workspace/nebius"
OUT = f"{W}/peers/comparison.xlsx"
BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="D9E1F2"); SEC = PatternFill("solid", fgColor="F2F2F2")
GREY = Font(color="808080", italic=True); BLUE = Font(color="0000CC"); ORANGE = Font(color="C65911"); BLACK = Font(color="000000"); PURPLE = Font(color="7030A0")
GRN = PatternFill("solid", fgColor="E2EFDA"); RED = PatternFill("solid", fgColor="FCE4D6"); YEL = PatternFill("solid", fgColor="FFF2CC")
wrap = Alignment(wrap_text=True, vertical="top")
def tagfont(t):
    t = (t or "").strip()
    if t.startswith(("R", "Reported")): return BLUE
    if t.startswith(("P", "Press")): return PURPLE
    if t.startswith(("E", "Estimated")): return ORANGE
    if t.startswith(("D", "Derived")): return BLACK
    return GREY

# ---------------------------------------------------------------- public peers (cached KPI workbooks)
PUB = {"CoreWeave": (f"{W}/peers/coreweave/crwv_kpis.xlsx", "O", "Jun-26"),
       "IREN": (f"{W}/peers/iren/iren_kpis.xlsx", "O", "Jun-26"),
       "Oracle OCI": (f"{W}/peers/oracle/oracle_kpis.xlsx", "O", "May-26")}
DATA = {}
for lab, (p, col, per) in PUB.items():
    wb = openpyxl.load_workbook(p, data_only=True); ws = wb["KPIs"]
    rows = {ws.cell(r, 1).value: r for r in range(4, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value[:1] in "KM" and len(ws.cell(r, 1).value) <= 4}
    wi = wb["Inputs"]; inp = {wi.cell(r, 1).value: r for r in range(4, wi.max_row + 1) if wi.cell(r, 1).value}
    DATA[lab] = (ws, rows, [c.value for c in ws[3]], wi, inp)
QLAB = {"E": "Dec-23", "F": "Mar-24", "G": "Jun-24", "H": "Sep-24", "I": "Dec-24", "J": "Mar-25", "K": "Jun-25", "L": "Sep-25", "M": "Dec-25", "N": "Mar-26", "O": "Jun-26", "P": "Sep-26"}
OQ = {"E": "Feb-24", "F": "May-24", "G": "Aug-24", "H": "Nov-24", "I": "Feb-25", "J": "May-25", "K": "Aug-25", "L": "Nov-25", "M": "Feb-26", "N": "Feb-26", "O": "May-26", "P": "Aug-26"}
def pub(lab, kid):
    ws, rows, hdr, wi, inp = DATA[lab]; p, col, per = PUB[lab]
    if kid not in rows: return None, "n/d", "", ""
    r = rows[kid]; ci = CI(col); nm = ws.cell(r, 2).value
    order = [ci] + ([ci + 1] if lab == "Oracle OCI" else []) + list(range(ci - 1, 5, -1))
    for k, cj in enumerate(order):
        v = ws.cell(r, cj).value
        if isinstance(v, (int, float)):
            if k > 0 and kid in ("K29", "K30", "K31", "K32", "K33", "K34", "K35"): break
            t = str(ws.cell(r + 1, cj).value or "")
            asof = (OQ if lab == "Oracle OCI" else QLAB).get(L(cj), hdr[cj - 1])
            if lab == "Oracle OCI" and L(cj) == "P": asof = "Aug-26 (FQ1 FY27)"
            elif lab == "Oracle OCI" and L(cj) == "O": asof = "May-26 (FQ4 FY26)"
            return v, t.split(" ")[0] if t else "", asof, nm
    return None, "n/d", "", nm
# CRWV K43/K44 computed from CRWV Inputs (CRWV workbook predates K41-K45)
ws_, rows_, hdr_, wi_, inp_ = DATA["CoreWeave"]; g = lambda i, c: wi_[f"{c}{inp_[i]}"].value
CX = {"K43": ((g("I01", "O") - g("I01", "K")) * 4 / (g("I09", "O") - g("I09", "K")), "D", "Jun-26"),
      "K44": (g("I01", "O") * 4 / ((g("I08", "O") + g("I08", "N")) / 2), "D", "Jun-26")}

# ---------------------------------------------------------------- private peers (user's parallel workstream; read-only)
PJ = json.load(open(f"{W}/peers/private/private_peers_values.json"))["summary"]
def priv(peer, kid):
    v = PJ.get(f"{peer}|{kid}")
    if not v or not isinstance(v[0], (int, float)): return None, "n/d", "", ""
    return v[0], v[2], v[1], v[3]

# ---------------------------------------------------------------- model v0.9e LOCK (read-only; source of truth from 2026-10-08 ~10:00 ET)
ML = NL.model_locks()                                       # asserts Cover A2 = v0.9e and that the model's own formulas close
MV = lambda k: ML[k]["value"]; MC = lambda k: ML[k]["cell"]
LOCK = {k: MV(k) for k in ("YE26_BASE", "YE26_HIGH", "FY27_BASE", "FY27_HIGH", "STREET_FY27", "GUIDE_LO", "GUIDE_HI", "FD_DEBT", "FD_IFCONV")}
MSFT_REV, MSFT_MW = MV("MSFT_REV"), MV("VIN_MW")
HURDLE, BLENDED = MV("HURDLE_XMSFT"), MV("BLENDED")
MODELF = "model/nebius_model.xlsx"
# FD share conventions: model Inputs r473/474 (v0.9e) must equal filings/fd_share_bridge.md (parsed - not typed); thresholds + mid tier from the md.
FDC = NL.fd_conventions()
FD_DEBT, FD_IFCONV, FD_MID = FDC["FD_DEBT"] / 1e6, FDC["FD_IFCONV"] / 1e6, FDC["FD_MID"] / 1e6
FDMD = "filings/fd_share_bridge.md"
assert abs(FD_DEBT - MV("FD_DEBT")) < 1e-6 and abs(FD_IFCONV - MV("FD_IFCONV")) < 1e-6, ("model FD vs fd_share_bridge", FD_DEBT, FD_IFCONV, ML["FD_DEBT"], ML["FD_IFCONV"])
assert abs(FD_DEBT - 358.56) < 0.01 and abs(FD_IFCONV - 376.65) < 0.01 and abs(FDC["PX_2030"] - 313.46) < 0.01 and abs(FDC["PX_2034"] - 324.65) < 0.01, FDC
assert abs(MV("FD_USED") - FD_DEBT) < 1e-6, ("model PT share basis should be the debt treatment at spot", ML["FD_USED"])
FD_V09D = FDC["FD_630"] / 1e6                     # 358,337,958 = v0.9d historical (Filings FD at 6/30/26)
FD_CONV_SENS = FD_IFCONV / FD_DEBT - 1
print(f"FD conventions: debt {FD_DEBT:.3f}M (< ${FDC['PX_2030']}) | mid {FD_MID:.3f}M | if-converted {FD_IFCONV:.3f}M (>= ${FDC['PX_2034']}) | v0.9d memo {FD_V09D:.3f}M")
print(f"v0.9e locks: FY27 base {LOCK['FY27_BASE']:.3f} ({MC('FY27_BASE')}), high {LOCK['FY27_HIGH']:.3f} ({MC('FY27_HIGH')}); MSFT FY27 rev {MSFT_REV:.2f} on {MSFT_MW:.3f} MW-yr; ex-MSFT {HURDLE:.4f} / blended {BLENDED:.4f}")
for k, (tgt, tol) in {"YE26_BASE": (257.7, 0.05), "YE26_HIGH": (284.2, 0.05), "FY27_BASE": (531.8, 0.1), "FY27_HIGH": (659.2, 0.1),
                      "STREET_FY27": (12300, 0.5), "GUIDE_LO": (7000, 0.5), "GUIDE_HI": (9000, 0.5)}.items():
    assert abs(LOCK[k] - tgt) <= tol, (k, LOCK[k], tgt)
assert abs(MSFT_REV - 2643.8) < 0.05 and abs(HURDLE - 27.6) < 0.05 and abs(BLENDED - 23.1) < 0.05, (MSFT_REV, HURDLE, BLENDED)
# Capacity billing IT MW path (row 147 'Billing MW (IT) - base'): Q1'26 = M, Q2'26 = N, Q4'26 = P
_qm = openpyxl.load_workbook(f"{W}/capacity/nebius_capacity.xlsx", data_only=True)["Quarterly MW"]
assert str(_qm["B147"].value).startswith("Billing MW (IT) - base") and _qm["M1"].value == "Q1 2026" and _qm["N1"].value == "Q2 2026" and _qm["P1"].value == "Q4 2026"
CAPB = {"Q1": _qm["M147"].value, "Q2": _qm["N147"].value, "Q4": _qm["P147"].value}
# 2026-10-08 ~23:30 ET (coordinator decision): Nebius ENERGIZED MW read live from the current Capacity lock (nebius_capacity.xlsx, NOT the outdated
#   nebius_capacity_v1.xlsx that gave 45 / 171.5 / 173): Quarterly MW row 130 'Total incl. plug' J130 (Q2'25) / M130 (Q1'26) / N130 (Q2'26);
#   identified-site low/high rows 124/125 col N. Same cells as filings/nebius_kpi_history.xlsx KPI row 30 and Checks row 22. Modeled GPUs follow the
#   same lock (Capacity 'GPUs deployed' row 34 M/N = Filings KPI row 34); fleet-weighted list price = Filings KPI row 52 col L (recalculated).
assert str(_qm["B130"].value).startswith("Total incl. plug") and str(_qm["B124"].value).startswith("Identified sites - energized low") and str(_qm["B125"].value).startswith("Identified sites - energized high") and _qm["J1"].value == "Q2 2025"
CAPE = {"Q2_25": _qm["J130"].value, "Q1_26": _qm["M130"].value, "Q2_26": _qm["N130"].value, "ID_LO": _qm["N124"].value, "ID_HI": _qm["N125"].value}
assert (CAPE["Q2_25"], CAPE["Q1_26"], CAPE["Q2_26"], CAPE["ID_LO"], CAPE["ID_HI"]) == (50, 181.5, 193, 157, 229), ("Capacity energized lock moved", CAPE)
_qg = openpyxl.load_workbook(f"{W}/capacity/nebius_capacity.xlsx", data_only=True)["GPUs deployed"]
assert str(_qg["A34"].value).startswith("TOTAL modeled GPUs")
CAPG = (_qg["M34"].value + _qg["N34"].value) / 2
assert abs(CAPG - 86738) < 1, CAPG
FLEET_LIST_Q2 = 5.6063348027  # filings/nebius_kpi_history.xlsx KPI!L52 (formula; value from a LibreOffice recalc of a /tmp copy, 2026-10-08 23:25 ET)
assert abs(CAPB["Q4"] - LOCK["YE26_BASE"]) < 0.05, CAPB
# price_or_mw.xlsx (this folder): YE26 rows + Street row must sit on the same v0.9e locks (run patch_price_or_mw.py first)
_pm = openpyxl.load_workbook(f"{W}/peers/price_or_mw.xlsx", data_only=True)["Implied_Price"]
PM = {"hi_mw": _pm["D7"].value, "hi_per_mw": _pm["F7"].value, "lo_mw": _pm["D6"].value, "lo_per_mw": _pm["F6"].value, "street_mw": _pm["D5"].value, "street_per_mw": _pm["F5"].value}
assert abs(PM["hi_mw"] - LOCK["YE26_HIGH"]) < 0.05 and abs(PM["lo_mw"] - LOCK["YE26_BASE"]) < 0.05, PM
assert abs(PM["street_mw"] - LOCK["FY27_BASE"]) < 0.01, ("price_or_mw Street row not on the v0.9e base - run patch_price_or_mw.py", PM)
JS = "model/nebius_model.xlsx Judge_Slide"; DR = "model/nebius_model.xlsx Drivers"

# ---------------------------------------------------------------- Nebius inputs (all from the Nebius workspace)
QS = "filings/nebius_filings.xlsx Quarterly"; KH = "filings/nebius_kpi_history.xlsx KPI"
NI = [  # id, item, value, unit, tag, source
 ("N01", "AI cloud segment revenue Q2'26", 574.9, "$M", "R", f"{QS} 'Segment revenue - Nebius AI cloud'; 6-K 2026-08-12"),
 ("N02", "AI cloud segment revenue Q2'25", 93.7, "$M", "R", f"{QS}"),
 ("N03", "Group revenue Q2'26 (memo)", 582.3, "$M", "R", f"{QS}"),
 ("N04", "AI cloud adj. EBITDA Q2'26", 285.7, "$M", "R", f"{QS}"),
 ("N05", "D&A Q2'26 (group)", 259.7, "$M", "R", f"{QS}"),
 ("N06", "Cash capex Q3'25", 955.5, "$M", "R", f"{QS}"),
 ("N07", "Cash capex Q4'25", 2056, "$M", "R", f"{QS}"),
 ("N08", "Cash capex Q1'26", 2472.9, "$M", "R", f"{QS}"),
 ("N09", "Cash capex Q2'26", 5657.4, "$M", "R", f"{QS}"),
 ("N10", "PP&E net 30-Jun-26", 13045.2, "$M", "R", f"{QS}"),
 ("N11", "Gross in-service PP&E 31-Mar-26 (gross 7,976.6 - not yet in use 2,427.9)", 5548.7, "$M", "D", f"{KH} row 26 (Q1'26 interim FS Note 7)"),
 ("N12", "Gross in-service PP&E 30-Jun-26 (gross 14,149.5 - not yet in use 7,836.4)", 6313.1, "$M", "D", f"{KH} row 26 (Q2'26 interim FS Note 7)"),
 ("N13", "Server & network equipment gross 31-Mar-26", 4692.7, "$M", "R", f"{KH} row 23"),
 ("N14", "Server & network equipment gross 30-Jun-26", 5197.7, "$M", "R", f"{KH} row 23"),
 ("N15", "Energized MW 30-Jun-25 (Capacity base case incl. plug)", CAPE["Q2_25"], "MW", "E", "capacity/nebius_capacity.xlsx Quarterly MW!J130 (= filings KPI row 30; mixed IT/gross basis; was 45 from outdated nebius_capacity_v1.xlsx)"),
 ("N16", "Energized MW 31-Mar-26 (Capacity base case incl. plug)", CAPE["Q1_26"], "MW", "E", "capacity/nebius_capacity.xlsx Quarterly MW!M130 (= filings KPI K30; was 171.5 from v1)"),
 ("N17", "Energized MW 30-Jun-26 (Capacity base case incl. plug)", CAPE["Q2_26"], "MW", "E", f"capacity/nebius_capacity.xlsx Quarterly MW!N130 (= filings KPI L30 / Checks r22; identified sites low {CAPE['ID_LO']:.0f} / high {CAPE['ID_HI']:.0f}, rows 124/125; was 173, v1)"),
 ("N18", "Active MW as stated by mgmt, YE25 (memo; old K01 basis)", 170, "MW", "R", "filings MW_Disclosures: Q4'25 letter '~170 MW of active power'"),
 ("N19", "Contracted power (>3.5 GW, May-26; YE26 target 5 GW)", 3500, "MW", "R", "filings MW_Disclosures: Q1'26 letter p.6 (lower bound)"),
 ("N20", "Average GPUs Q2'26 (modeled)", CAPG, "GPUs", "E", "capacity/nebius_capacity.xlsx GPUs deployed!M34:N34 average (= filings KPI row 55; MW x GPUs/MW, not a company count; was 80,047.5 on v1 MW)"),
 ("N21", "Days in Q2'26", 91, "days", "R", "calendar"),
 ("N22", "Fleet-weighted on-demand list price Q2'26 avg", FLEET_LIST_Q2, "$/GPU-hr", "E", "filings/nebius_kpi_history.xlsx KPI!L52 (GPU Market list prices x current Capacity GPU mix; was 5.571 on v1 mix)"),
 ("N23", "RPO 30-Jun-26", 37490.6, "$M", "R", f"{QS}; 100% AI cloud"),
 ("N24", "RPO % within 24 months", 36, "%", "R", "shared/findings_log.xlsx row 32 (Q2'26 interim FS)"),
 ("N25", "Customer C % of Q2'26 revenue", 24, "%", "R", f"{QS} concentration; Q2'26 interim FS Note 1 p.F-17"),
 ("N26", "Customer D %", 21, "%", "R", "as above"),
 ("N27", "Customer E %", 14, "%", "R", "as above"),
 ("N28", "Customers >=10% of revenue", 3, "#", "R", "as above"),
 ("N29", "Useful life servers & network (from 1-Jan-26)", 5, "years", "R", "filings D&A_Policy (20-F FY25 p.56/F-15)"),
 ("N30", "Weighted effective cost of debt (converts + secured)", 5.1, "%", "D", "shared/findings_log.xlsx row 51 (GPU Market & Comps)"),
 ("N31", "Interest expense Q2'26 (net of capitalized; incl. accretion)", 119.1, "$M", "R", "filings FactBase row 21 (6-K 2026-08-12)"),
 ("N32", "Capitalized interest Q2'26", 40.2, "$M", "R", "shared/findings_log.xlsx row 137 / D&A_Policy (Q2FS Note 12)"),
 ("N33", "Total debt carrying 31-Mar-26", 8450.4, "$M", "R", f"{QS}"),
 ("N34", "Total debt carrying 30-Jun-26", 8545.7, "$M", "R", f"{QS} (accreted principal at maturity 10,041.8)"),
 ("N35", "Gross in-service PP&E 31-Dec-24", 919.2, "$M", "D", f"{KH} row 26"),
 ("N36", "Gross in-service PP&E 31-Dec-25", 3772.9, "$M", "D", f"{KH} row 26"),
 ("N37", "Energized MW 31-Dec-24", 25, "MW", "E", f"{KH} row 28"),
 ("N38", "Energized MW 31-Dec-25", 170, "MW", "E", f"{KH} row 28 (ties to mgmt ~170)"),
 # --- billed-IT basis (K02b) and v0.9d LOCKS ---
 ("N_BILL_Q1", "Billing IT MW 31-Mar-26 (Capacity 'Billing MW (IT) - base': Vineland billed tranches + other sites lagged 1Q)", CAPB["Q1"], "IT MW", "E", "capacity/nebius_capacity.xlsx Quarterly MW row 147 col M (Q1 2026); K02b denominator"),
 ("N_BILL_Q2", "Billing IT MW 30-Jun-26 (same row)", CAPB["Q2"], "IT MW", "E", "capacity/nebius_capacity.xlsx Quarterly MW row 147 col N (Q2 2026); K02b denominator"),
 ("N_FD_DEBT", f"Fully diluted shares - new Feb-2030/Feb-2034 converts as DEBT (share price < ~${FDC['PX_2030']:.2f}; both series out of the money)", FD_DEBT, "M shares", "D",
  f"{FDMD} line {FDC['FD_DEBT_LINE']} 'Alternative convention' FD = {FDC['FD_DEBT']:,} (FD_Bridge C30); pair with the $5.75B face (accreting to $6.67B) kept in DEBT. Applies at the {FDMD} reference price ${FDC['REF_PX']} ({FDC['REF_PX_NOTE']})."),
 ("N_FD_IFCONV", f"Fully diluted shares - new converts IF-CONVERTED (share price >= ~${FDC['PX_2034']:.2f}; both series in the money)", FD_IFCONV, "M shares", "D",
  f"{FDMD} line {FDC['FD_IFCONV_LINE']} FD1 = {FDC['FD_IFCONV']:,} (= post630_capital_update.md 'FD pro forma after Aug-26'); pair with ALL convert debt REMOVED from EV. Between ${FDC['PX_2030']:.2f} and ${FDC['PX_2034']:.2f} only the 2030s convert: {FDC['FD_MID']:,} (see NOTE_FD)."),
 ("N_FD_V09D", "Fully diluted shares - v0.9d HISTORICAL memo: Filings FD at 6/30/26 (FDSel=1)", FD_V09D, "M shares", "D",
  f"{FDMD} line {FDC['FD_630_LINE']} FD0 {FDC['FD_630']:,} = model/nebius_model.xlsx Checks r7 (v0.9d) = filings/dilution_scenarios.xlsx Base_ShareCount B22; pre-Aug-26 exchange/new converts - historical only"),
 ("N_BILL_YE26_BASE", "Billing IT MW YE26 - BASE (Capacity Q4'26 'Billing MW (IT) - base') [LOCK v0.9e ~258, unchanged]", LOCK["YE26_BASE"], "IT MW", "D",
  f"{JS} {MC('YE26_BASE').split('!')[1]} (257.7; model v0.9e); = capacity/nebius_capacity.xlsx Quarterly MW row 147 col P {CAPB['Q4']:.2f}"),
 ("N_BILL_YE26_HIGH", "Billing IT MW YE26 - HIGH (Vineland at billed tranches) [LOCK v0.9e ~284, unchanged]", LOCK["YE26_HIGH"], "IT MW", "D",
  f"{JS} {MC('YE26_HIGH').split('!')[1]} (284.2; model v0.9e); same MW as peers/price_or_mw.xlsx Implied_Price D7"),
 ("N_BILL_FY27_BASE", "FY27 average billed IT MW - BASE, engine-linked (non-Vineland path + Vineland billed MW-yr) [LOCK v0.9e ~531.8]", LOCK["FY27_BASE"], "IT MW", "D",
  f"{MODELF} {MC('FY27_BASE')} = {MC('NONVIN_BASE')} {MV('NONVIN_BASE'):.2f} non-Vineland + {MC('VIN_MW')} {MSFT_MW:.2f} Vineland billed MW-yr (model v0.9e, source of truth)."),
 ("N_BILL_FY27_HIGH", "FY27 average billed IT MW - HIGH, engine-linked (high non-Vineland path + Vineland billed MW-yr) [LOCK v0.9e ~659.2]", LOCK["FY27_HIGH"], "IT MW", "D",
  f"{MODELF} {MC('FY27_HIGH')} = {MC('NONVIN_HIGH')} {MV('NONVIN_HIGH'):.2f} + {MSFT_MW:.2f} (model v0.9e)."),
 ("N_MSFT_FY27_REV", "Microsoft (Vineland) FY27 revenue in the model engine ($M; Contracts MS_Rev)", MSFT_REV, "$M", "D",
  f"{MODELF} {MC('MSFT_REV')} = SUMIF(Q_Yr,2027,MS_Rev) (Contracts row 145; T1-T4 + ramp + Q1'27 genset cap + T5-T9 Bloom starts) - model v0.9e"),
 ("N_MSFT_FY27_MW", "Vineland / Microsoft FY27 billed MW-yr on the revenue basis (= MS revenue / $14.5M)", MSFT_MW, "IT MW", "D",
  f"{MODELF} {MC('VIN_MW')} = {MC('MSFT_REV').split('!')[1]} / Inputs F111 ($14.5M); alt Contracts row 141 billing-MW avg {MV('MS_BILLMW_ALT'):.1f} (memo)"),
 ("N_STREET_FY27", "Street FY27 revenue consensus (memo for forward $/billed MW)", LOCK["STREET_FY27"], "$M", "E", f"{JS} {MC('STREET_FY27').split('!')[1]} / Cons_Rev (model v0.9e); peers/price_or_mw.xlsx Revenue_Sources C4 (21 analysts, as of 2026-10-04)"),
 ("N_GUIDE_LO", "YE26 run-rate guide - low (Dec AI cloud x 12)", LOCK["GUIDE_LO"], "$M", "R", f"{JS} {MC('GUIDE_LO').split('!')[1]}; Q2'26 letter / Aug 12 call (reaffirmed)"),
 ("N_GUIDE_HI", "YE26 run-rate guide - high", LOCK["GUIDE_HI"], "$M", "R", f"{JS} {MC('GUIDE_HI').split('!')[1]}"),
]
# derived forward memo rows (live formulas in Nebius_Inputs; context only, not scored, not in Comparison medians)
NMEMO = [  # id, item, formula builder, unit, tag, source
 ("D_YE26_LO_BASE", "YE26 guide low / YE26 BASE billing ($7B / ~258)", lambda n: f"={n('N_GUIDE_LO')}/{n('N_BILL_YE26_BASE')}", "$M/billed IT MW-yr", "D", f"ties to {JS} B40 (27.16); price_or_mw Implied_Price F6"),
 ("D_YE26_HI_HIGH", "YE26 guide high / YE26 HIGH billing ($9B / ~284)", lambda n: f"={n('N_GUIDE_HI')}/{n('N_BILL_YE26_HIGH')}", "$M/billed IT MW-yr", "D", f"ties to {JS} C41 (31.67); price_or_mw Implied_Price F7"),
 ("D_STREET_FY27_BASE", "Street FY27 / FY27 avg BASE billed, engine-linked ($12.3B / ~531.8) - BLENDED memo (incl. $14.5M Microsoft anchor)", lambda n: f"={n('N_STREET_FY27')}/{n('N_BILL_FY27_BASE')}", "$M/billed IT MW-yr", "D", f"~23.1; ties to {DR} {MC('BLENDED').split('!')[1]} ({BLENDED:.2f}) / {JS} B16; price_or_mw Implied_Price F5"),
 ("D_STREET_FY27_NONMSFT", "Street FY27 EX-MICROSOFT / FY27 BASE billed ex-Vineland (($12.3B - MSFT) / (531.8 - 182.3)) - USE AGAINST PEERS", lambda n: f"=({n('N_STREET_FY27')}-{n('N_MSFT_FY27_REV')})/({n('N_BILL_FY27_BASE')}-{n('N_MSFT_FY27_MW')})", "$M/billed IT MW-yr", "D", f"(12,300 - {MSFT_REV:,.1f}) / ({LOCK['FY27_BASE']:.2f} - {MSFT_MW:.2f}) = {HURDLE:.2f}; ties to {DR} {MC('HURDLE_XMSFT').split('!')[1]} / {JS} B20 (~$27.6M, Grok Bot). Peer $/MW has no $14.5M Microsoft anchor."),
 ("D_STREET_FY27_HIGH", "Street FY27 / FY27 avg HIGH billed, engine-linked ($12.3B / ~659.2) - blended", lambda n: f"={n('N_STREET_FY27')}/{n('N_BILL_FY27_HIGH')}", "$M/billed IT MW-yr", "D", "high path; price_or_mw Implied_Price F24"),
 ("D_STREET_FY27_NONMSFT_HIGH", "Street FY27 EX-MICROSOFT / FY27 HIGH billed ex-Vineland - high path", lambda n: f"=({n('N_STREET_FY27')}-{n('N_MSFT_FY27_REV')})/({n('N_BILL_FY27_HIGH')}-{n('N_MSFT_FY27_MW')})", "$M/billed IT MW-yr", "D", f"Vineland billed MW-yr identical on both paths ({DR} B33 = C33)"),
 ("D_COV16_BASE", "MW coverage at $16M comps: (Street / $16M) / FY27 BASE billed (x)", lambda n: f"=({n('N_STREET_FY27')}/16)/{n('N_BILL_FY27_BASE')}", "x", "D", f"ties to {DR} {MC('COV16_BASE').split('!')[1]} ({MV('COV16_BASE'):.3f}x; 768.75 MW needed)"),
 ("D_COV16_HIGH", "MW coverage at $16M comps: (Street / $16M) / FY27 HIGH billed (x)", lambda n: f"=({n('N_STREET_FY27')}/16)/{n('N_BILL_FY27_HIGH')}", "x", "D", f"ties to {DR} {MC('COV16_HIGH').split('!')[1]} ({MV('COV16_HIGH'):.3f}x)"),
 ("D_FD_CONV_SENS", "Share convention spread: if-converted FD vs debt-convention FD (%)", lambda n: f"=({n('N_FD_IFCONV')}/{n('N_FD_DEBT')}-1)*100", "%", "D", "~+5.0%. No Nebius per-share or EV-per-MW figure exists in this workbook/charts, so no KPI value depends on the convention."),
 ("D_K02B_BASE_MW", "Q2'26 average billing IT MW (K02b denominator)", lambda n: f"=AVERAGE({n('N_BILL_Q1')},{n('N_BILL_Q2')})", "IT MW", "E", "capacity row 147 M:N average"),
 ("D_YE26_VS_Q2", "YE26 BASE / HIGH billing vs Q2'26 average billing (x)", lambda n: f"={n('N_BILL_YE26_BASE')}/AVERAGE({n('N_BILL_Q1')},{n('N_BILL_Q2')})", "x", "D", "base-path billing build multiple to YE26 (HIGH: next row)"),
 ("D_YE26H_VS_Q2", "YE26 HIGH billing vs Q2'26 average billing (x)", lambda n: f"={n('N_BILL_YE26_HIGH')}/AVERAGE({n('N_BILL_Q1')},{n('N_BILL_Q2')})", "x", "D", ""),
]
NOTE_1417 = "Microsoft ACV: $14.49M per IT MW [E, house] = model v0.9e Contracts!E12 = $17,392.9M / (5 x 240 IT MW est.); agrees with the Drivers 14.5 input (Inputs!F111). Neither company states the contract MW. HISTORY (superseded Oct 9 2026): 11.60 on 300 MW (site size, not contract MW) and 17.08 GPU-derived (retired)."
NR = {r[0]: 4 + i for i, r in enumerate(NI)}
MEMO_HDR = 4 + len(NI) + 1
for j, r_ in enumerate(NMEMO): NR[r_[0]] = MEMO_HDR + 1 + j
n = lambda i: f"Nebius_Inputs!$C${NR[i]}"
nl = lambda i: f"$C${NR[i]}"  # same-sheet reference
NBF = {  # KPI -> (formula, tag, as-of, note)
 "K01": (f"={n('N01')}*4/{n('N17')}", "E", "Jun-26", "Q2'26 AI cloud rev x4 / Jun-26 energized MW (193, Capacity base case, E)"),
 "K02": (f"={n('N01')}*4/AVERAGE({n('N16')},{n('N17')})", "E", "Jun-26", "average energized MW (181.5+193)/2 = 187.25 (Capacity base case, E); replaces old 13.7 on stale YE25 170 MW"),
 "K02b": (f"={n('N01')}*4/AVERAGE({n('N_BILL_Q1')},{n('N_BILL_Q2')})", "E", "Jun-26",
          "BILLED-IT basis: Q2'26 AI cloud rev x4 / AVERAGE(Capacity billing IT MW Q1'26, Q2'26) (row 147; Vineland billed tranches + other sites lagged 1Q). "
          "Not energized/gross MW. Same billed-IT basis as the forward $/billed MW memo (YE26 258/284; FY27 avg 531.8/659.2 engine-linked; model v0.9e - see Nebius_Inputs)."),
 "K03": (f"={n('N01')}*4/{n('N19')}", "D", "Jun-26", ">3.5 GW is a lower bound -> value is an upper bound"),
 "K04": (f"={n('N19')}/{n('N17')}", "E", "Jun-26", ""),
 "K05": (f"={n('N04')}*4/AVERAGE({n('N16')},{n('N17')})", "E", "Jun-26", "AI cloud segment adj. EBITDA"),
 "K06": (f"={n('N01')}*4000/{n('N20')}", "E", "Jun-26", "modeled GPUs"),
 "K07": (f"={n('N01')}*1000000/({n('N20')}*{n('N21')}*24)", "E", "Jun-26", "incl. non-GPU-hour revenue"),
 "K08": (f"={n('N22')}", "E", "Jun-26", ""),
 "K09": (f"={n('N01')}*1000000/({n('N20')}*{n('N21')}*24)/{n('N22')}*100", "E", "Jun-26", ""),
 "K10": ("n.m.", "n.m.", "Jun-26", "MW added in Q2'26 only +1.5 (Capacity est.) -> not meaningful"),
 "K12": (f"=({n('N36')}-{n('N35')})/({n('N38')}-{n('N37')})", "E", "Dec-25 (FY25)", "FY25 basis; quarterly n.m."),
 "K13": (f"=SUM({n('N06')},{n('N07')},{n('N08')},{n('N09')})/({n('N17')}-{n('N15')})", "E", "Jun-26", "capex funds future (not-yet-in-use) capacity -> inflated"),
 "K14": (f"={n('N10')}/{n('N17')}", "E", "Jun-26", "incl. $7.8B not yet in use"),
 "K15": (f"={n('N12')}/{n('N17')}", "E", "Jun-26", ""),
 "K16": (f"={n('N29')}", "R", "Jun-26", ""),
 "K17": (f"={n('N05')}*4/AVERAGE({n('N11')},{n('N12')})*100", "D", "Jun-26", "group D&A incl. amortization"),
 "K18": (f"=100/({n('N05')}*4/AVERAGE({n('N11')},{n('N12')})*100)", "D", "Jun-26", ""),
 "K19": (f"={n('N05')}/{n('N01')}*100", "D", "Jun-26", "group D&A / AI cloud revenue"),
 "K20": (f"={n('N23')}/1000", "R", "Jun-26", ""),
 "K22": (f"={n('N24')}", "R", "Jun-26", ""),
 "K23": (f"={n('N23')}/({n('N01')}*4)", "D", "Jun-26", ""),
 "K24": (f"={n('N25')}", "R", "Jun-26", "Customer C (unnamed)"),
 "K25": (f"={n('N26')}", "R", "Jun-26", ""),
 "K26": (f"={n('N25')}+{n('N26')}+{n('N27')}", "R", "Jun-26", ""),
 "K27": (f"={n('N28')}", "R", "Jun-26", ""),
 "K28": (f"={n('N30')}", "D", "Jun-26", ""),
 "K36": (f"={n('N31')}*4/AVERAGE({n('N33')},{n('N34')})*100", "D", "Jun-26", "P&L interest (incl. convert accretion) / avg carrying debt"),
 "K38": (f"={n('N34')}/1000", "R", "Jun-26", "carrying value (not principal)"),
 "K39": (f"={n('N34')}/{n('N23')}", "D", "Jun-26", ""),
 "K40": (f"={n('N32')}/({n('N31')}+{n('N32')})*100", "D", "Jun-26", ""),
 "K43": (f"=({n('N01')}-{n('N02')})*4/({n('N17')}-{n('N15')})", "E", "Jun-26", ""),
 "K44": (f"={n('N01')}*4/AVERAGE({n('N13')},{n('N14')})", "D", "Jun-26", "server & network equipment (GPU hardware) basis"),
}
# python mirror of Nebius values (for flags/ranking; Excel cells stay live formulas)
V = {r[0]: r[2] for r in NI}
avg = lambda a, b: (a + b) / 2
NBV = {"K01": V["N01"]*4/V["N17"], "K02": V["N01"]*4/avg(V["N16"], V["N17"]), "K02b": V["N01"]*4/avg(V["N_BILL_Q1"], V["N_BILL_Q2"]), "K03": V["N01"]*4/V["N19"], "K04": V["N19"]/V["N17"],
 "K05": V["N04"]*4/avg(V["N16"], V["N17"]), "K06": V["N01"]*4000/V["N20"], "K07": V["N01"]*1e6/(V["N20"]*91*24), "K08": V["N22"],
 "K09": V["N01"]*1e6/(V["N20"]*91*24)/V["N22"]*100, "K12": (V["N36"]-V["N35"])/(V["N38"]-V["N37"]),
 "K13": (V["N06"]+V["N07"]+V["N08"]+V["N09"])/(V["N17"]-V["N15"]), "K14": V["N10"]/V["N17"], "K15": V["N12"]/V["N17"], "K16": 5,
 "K17": V["N05"]*4/avg(V["N11"], V["N12"])*100, "K19": V["N05"]/V["N01"]*100, "K20": V["N23"]/1000, "K22": 36, "K23": V["N23"]/(V["N01"]*4),
 "K24": 24, "K25": 21, "K26": 59, "K27": 3, "K28": 5.1, "K36": V["N31"]*4/avg(V["N33"], V["N34"])*100, "K38": V["N34"]/1000, "K39": V["N34"]/V["N23"],
 "K40": V["N32"]/(V["N31"]+V["N32"])*100, "K43": (V["N01"]-V["N02"])*4/(V["N17"]-V["N15"]), "K44": V["N01"]*4/avg(V["N13"], V["N14"])}
NBV["K18"] = 100 / NBV["K17"]

# ---------------------------------------------------------------- KPI grid
ORDER = ["K01","K02","K02b","K03","K04","K05","K06","K07","K08","K09","K10","K12","K13","K14","K15","K16","K17","K18","K19","K20","K21","K22","K23",
         "K24","K25","K26","K27","K28","K29","K34","K35","K36","K37","K38","K39","K40","K41","K42","K43","K44","K45"]
PEERS = ["CoreWeave", "IREN", "Oracle OCI", "Nscale", "Lambda", "Crusoe"]
NONCOMP = {("Crusoe", "K14"): "shell+power only (ex-GPU)", ("IREN", "K09"): "vs IREN's own $2.00 assumption, not list",
           ("Crusoe", "K21"): "company TCV metric", ("Nscale", "K28"): "legacy $137M debt; marginal secured ~6.25% E (PX09)"}
# direction: +1 higher is better for Nebius, -1 lower is better, 0 context only
DIR = {"K01":1,"K02":1,"K02b":1,"K03":1,"K04":0,"K05":1,"K06":1,"K07":1,"K08":0,"K09":1,"K10":-1,"K12":-1,"K13":-1,"K14":-1,"K15":-1,"K16":-1,"K17":1,"K18":-1,
       "K19":0,"K20":0,"K21":0,"K22":1,"K23":1,"K24":-1,"K25":-1,"K26":-1,"K27":-1,"K28":-1,"K29":0,"K34":0,"K35":0,"K36":-1,"K37":-1,"K38":0,"K39":-1,
       "K40":-1,"K41":1,"K42":1,"K43":1,"K44":1,"K45":1}
WHY = {  # brief interpretation appended to the auto numbers
 "K01":"higher revenue density per energized MW (Nebius MW is a Capacity estimate; mixed IT/gross basis)",
 "K02":"same on average MW - the fixed, comparable figure; Blackwell-heavy, mostly new fleet lifts density",
 "K02b":"same numerator on average BILLING IT MW (billed-IT basis; Nebius Capacity billing path, CoreWeave active~billed) - the basis the v0.9e Judge slide / Drivers use for $/billed MW",
 "K03":"monetises contracted power faster? Nebius has far more unbuilt contracted MW than CRWV",
 "K04":"build-out multiple: big growth runway but heavy execution/funding need (context, not scored)",
 "K05":"segment adj. EBITDA per MW well above CRWV-style peers",
 "K06":"revenue per (modeled) GPU", "K07":"realized $/GPU-hr above peers (modeled GPU base; includes non-GPU revenue)",
 "K08":"list-price context", "K09":"price-based 'utilization' (mixes discount with use); not physical utilization",
 "K10":"single-quarter capex/MW not meaningful for Nebius (+1.5 MW)", "K12":"in-service $ per MW added (FY25 basis for Nebius)",
 "K13":"capex per MW added is high because capex funds 2027 capacity (not-yet-in-use $7.8B)",
 "K14":"net PP&E per active MW inflated by construction-in-progress", "K15":"gross in-service PP&E per MW",
 "K16":"shorter life = more conservative accounting (Nebius moved 4->5 yrs in 2026; CRWV/Oracle 6)",
 "K17":"higher depreciation yield = more conservative", "K18":"lower implied life = more conservative",
 "K19":"D&A burden context", "K20":"absolute size context", "K21":"company backlog metric (context)",
 "K22":"share of RPO converting within 24m", "K23":"backlog coverage of current run-rate",
 "K24":"lower concentration is better", "K25":"lower is better", "K26":"lower combined concentration is better", "K27":"fewer >=10% customers is better",
 "K28":"cheaper debt (converts-heavy) is better", "K29":"facility-level rate (context)", "K34":"facility-level rate (context)", "K35":"structure context",
 "K36":"lower cash cost of debt is better", "K37":"company-reported rate", "K38":"size context", "K39":"lower debt per $ of backlog is better",
 "K40":"high capitalized share flatters P&L interest", "K41":"physical utilization (Oracle only)", "K42":"n/d for Nebius (24m bucket only)",
 "K43":"marginal revenue per new MW", "K44":"asset turn on GPU hardware", "K45":"segment gross margin (n/d for Nebius)",
}
cells = {}  # (kid, peer) -> (value, tag, asof, label)
names = {}
for kid in ORDER:
    for p in PEERS:
        if p in PUB:
            v, t, a, nm = pub(p, kid)
            if p == "CoreWeave" and v is None and kid in CX: v, t, a = CX[kid]
            if nm and kid not in names: names[kid] = nm
        else:
            v, t, a, nm = priv(p, kid)
        cells[(kid, p)] = (v, t, a)
# K18 consistency check (2026-10-08): packs are read as-is (no re-derivation). crwv_kpis.xlsx K18 fixed at source (= 100/K17).
K18CHK = {}
for p in PUB:
    v17, v18 = cells[("K17", p)][0], cells[("K18", p)][0]
    if isinstance(v17, (int, float)) and v17 and isinstance(v18, (int, float)):
        K18CHK[p] = round(v18, 2)
        assert abs(v18 - 100 / v17) <= 0.01, (p, "K18 != 100/K17 in pack - fix the pack, do not patch here", v18, 100 / v17)
print("K18 check (pack values, = 100/K17):", K18CHK)
UNIT = {}
ws0, rows0, _, _, _ = DATA["IREN"]
for kid in ORDER:
    if kid in rows0:
        names.setdefault(kid, ws0.cell(rows0[kid], 2).value); UNIT[kid] = ws0.cell(rows0[kid], 3).value
for kid in ORDER:
    if kid not in UNIT:
        wsx, rx, _, _, _ = DATA["CoreWeave"]
        if kid in rx: UNIT[kid] = wsx.cell(rx[kid], 3).value; names.setdefault(kid, wsx.cell(rx[kid], 2).value)

names["K02b"] = "Revenue per BILLING IT MW - average MW (billed-IT basis; Nebius = Capacity billing path)"; UNIT.setdefault("K02b", "$M/MW-yr")

def flag(kid):
    d = DIR.get(kid, 0); nv = NBV.get(kid)
    peers = [(p, cells[(kid, p)][0]) for p in PEERS if isinstance(cells[(kid, p)][0], (int, float)) and (p, kid) not in NONCOMP]
    if nv is None: return "n/a", None, None, f"Nebius n/d. {WHY.get(kid,'')}", peers
    if not peers: return "n/a", None, None, f"No comparable peer value. {WHY.get(kid,'')}", peers
    med = statistics.median([x for _, x in peers])
    den = max(abs(med), statistics.mean([abs(x) for _, x in peers]))
    gap = (nv - med) / den * 100 if den else None
    hi = max(peers, key=lambda x: x[1]); lo = min(peers, key=lambda x: x[1])
    base = f"Nebius {nv:,.2f} vs peer median {med:,.2f} (range {lo[0]} {lo[1]:,.2f} - {hi[0]} {hi[1]:,.2f}; n={len(peers)})"
    if d == 0: return "Context", med, gap, f"{base}; {WHY.get(kid,'')}", peers
    sg = gap * d if gap is not None else 0
    f = "Better" if sg > 10 else ("Worse" if sg < -10 else "In line")
    return f, med, sg, f"{base}; {WHY.get(kid,'')}", peers

# ---------------------------------------------------------------- workbook
wb = openpyxl.Workbook()
rd = wb.active; rd.title = "README"
wsC = wb.create_sheet("Comparison"); wsR = wb.create_sheet("Ranking"); wsN = wb.create_sheet("Nebius_Inputs")
wsK = wb.create_sheet("Contracts_$MWyr"); wsS = wb.create_sheet("Sites"); wsD = wb.create_sheet("Chart_Data"); wsG = wb.create_sheet("Charts")

# Nebius_Inputs
wsN["A1"] = "Nebius inputs for the Comparison tab (Q2'26 = quarter ended 30-Jun-2026 unless stated). Comparison Nebius cells are live formulas on column C."; wsN["A1"].font = BOLD
for j, h in enumerate(["ID", "Item", "Value", "Unit", "Tag", "Source (Nebius workspace file / row)"], 1):
    c = wsN.cell(row=3, column=j, value=h); c.font = BOLD; c.fill = HDR
for i, (iid, it, v, u, t, s) in enumerate(NI):
    r = 4 + i
    for j, x in enumerate([iid, it, v, u, t, s], 1): wsN.cell(row=r, column=j, value=x)
    wsN.cell(row=r, column=3).font = tagfont(t)
    if iid.startswith(("N_FD", "N_BILL_YE26", "N_BILL_FY27")) and not iid.endswith("_V09D"):  # yellow = current locks / share conventions
        for j in (1, 2, 3): wsN.cell(row=r, column=j).fill = YEL
        wsN.cell(row=r, column=3).number_format = "0.0"
    elif iid.startswith("N_BILL_Q") or iid.endswith("_V09D") or iid.startswith("N_MSFT"): wsN.cell(row=r, column=3).number_format = "0.00"
wsN.cell(row=MEMO_HDR, column=1, value="Forward memo - $ per BILLED IT MW (live formulas on the model v0.9e lock rows above; context only, NOT scored and NOT in Comparison medians)").font = BOLD
for j, (iid, it, fb, u, t, s_) in enumerate(NMEMO):
    r = MEMO_HDR + 1 + j
    for k, x in enumerate([iid, it, fb(lambda i: nl(i)), u, t, s_], 1): wsN.cell(row=r, column=k, value=x)
    wsN.cell(row=r, column=3).number_format = "0.00"; wsN.cell(row=r, column=3).font = tagfont(t)
rn = MEMO_HDR + len(NMEMO) + 2
for t_ in [("NOTE_1417", NOTE_1417),
           ("NOTE_FD", f"SHARE-COUNT CONVENTION (2026-10-08, {FDMD}): the new Feb-2030 ($3.45B, conv. ${FDC['PX_2030']:.2f}) and Feb-2034 ($2.30B, conv. ${FDC['PX_2034']:.2f}) notes count as DEBT while the share price is below their conversion prices -> N_FD_DEBT {FD_DEBT:.2f}M with the $5.75B face in debt; "
                       f"IF-CONVERTED above -> N_FD_IFCONV {FD_IFCONV:.2f}M with ALL convert debt removed from EV. Between ${FDC['PX_2030']:.2f} and ${FDC['PX_2034']:.2f} only the 2030s convert: {FD_MID:.2f}M (D). At the {FDMD} reference ${FDC['REF_PX']} ({FDC['REF_PX_NOTE']}) both are out of the money -> debt convention. "
                       f"Never mix share and debt sides. N_FD_V09D {FD_V09D:.2f}M is the v0.9d historical (6/30, pre-exchange) value only. No per-share or EV/MW metric in this workbook."),
           ("NOTE_FY27", f"FY27 BILLED-MW LOCKS (model v0.9e = source of truth, read via peers/nebius_locks.py model_locks()): BASE {LOCK['FY27_BASE']:.2f} ({MODELF} {MC('FY27_BASE')}) / HIGH {LOCK['FY27_HIGH']:.2f} ({MC('FY27_HIGH')}) = non-Vineland path {MV('NONVIN_BASE'):.2f} / {MV('NONVIN_HIGH'):.2f} + Vineland billed MW-yr {MSFT_MW:.2f} "
                       f"(engine Microsoft FY27 revenue ${MSFT_REV:,.1f}M / $14.5M). Street blended $12.3B / {LOCK['FY27_BASE']:.2f} = {BLENDED:.2f} (memo); PEER COMPARATOR = D_STREET_FY27_NONMSFT = (12,300 - {MSFT_REV:,.1f}) / {MV('NONVIN_BASE'):.2f} = {HURDLE:.2f} (= {MC('HURDLE_XMSFT')}). "
                       f"At $16M comps the Street needs {MV('MW_NEEDED_16'):.2f} MW = {MV('COV16_BASE'):.2f}x base / {MV('COV16_HIGH'):.2f}x high (D_COV16_*). YE26 unchanged 257.7 / 284.2."),
           ("NOTE_LOCKSRC", f"LOCK SOURCE (v0.9e): YE26 = {JS} B34/B35 = capacity row 147 col P {CAPB['Q4']:.2f}; FY27 / Microsoft / hurdle = {DR} rows 9-10, 30-39; FD = {MODELF} Inputs F473/F474 (= {FDMD}). "
                         f"Build asserts model version v0.9e, FD {FD_DEBT:.2f}/{FD_IFCONV:.2f} (model = Filings), YE26 257.7/284.2, FY27 531.8/659.2, MSFT $2,643.8M, hurdle ~27.6 / blended ~23.1, and that "
                         f"price_or_mw.xlsx Implied_Price D7/F7 = {PM['hi_mw']:.1f} MW / ${PM['hi_per_mw']:.2f}M and Street row D5 = {PM['street_mw']:.2f} MW / ${PM['street_per_mw']:.2f}M re-read and matched. Yellow = current lock rows."),
           ("NOTE_K02B", "K02b keeps the BILLED-IT basis: denominators N_BILL_Q1/Q2 are Capacity billing IT MW (not energized N16/N17). Comparison K02b and every forward $/billed MW memo above share that basis.")]:
    wsN.cell(row=rn, column=1, value=t_[0]).font = BOLD; wsN.cell(row=rn, column=2, value=t_[1]).alignment = wrap
    wsN.merge_cells(start_row=rn, start_column=2, end_row=rn, end_column=6); wsN.row_dimensions[rn].height = 30; rn += 1
for col, w in zip("ABCDEF", [20, 62, 12, 16, 6, 90]): wsN.column_dimensions[col].width = w

# Comparison
wsC["A1"] = ("Peer Comparison - Nebius vs CoreWeave, IREN (AI Cloud), Oracle OCI, Nscale, Lambda, Crusoe. Template KPI IDs/definitions (peers/coreweave/crwv_kpis.xlsx 'Definitions'; "
             "K41-K45 in peers/peer_kpi_template.py). Latest available period per peer; as-of shown per cell. Nebius cells are formulas on Nebius_Inputs. "
             "Private-peer values from peers/private/private_peers_values.json (Summary) with their tags. Built 2026-10-04 ET; v0.9e refresh 2026-10-08 ET by peers/build_comparison.py "
             f"(Nebius_Inputs locks from model v0.9e: billed IT MW YE26 258 base / 284 high; FY27 avg {LOCK['FY27_BASE']:.1f} base / {LOCK['FY27_HIGH']:.1f} high, engine-linked. Shares: {FD_DEBT:.2f}M with new converts as debt (< ${FDC['PX_2030']:.0f}) / {FD_IFCONV:.2f}M if-converted (> ${FDC['PX_2034']:.0f}). "
             "Next Nebius print: Q3'26 results ~Nov 10 (est.) - date not announced by the company.")
wsC["A1"].alignment = wrap; wsC.merge_cells("A1:AC1"); wsC.row_dimensions[1].height = 48
hdr = ["ID", "KPI", "Unit", "Nebius", "Tag", "As of"]
for p in PEERS: hdr += [p, "Tag", "As of"]
hdr += ["Peer median (formula; comparable peers)", "Nebius better / worse", "Reason (auto numbers + interpretation)", "Nebius basis note", "Peer basis comment"]
for j, h in enumerate(hdr, 1):
    c = wsC.cell(row=3, column=j, value=h); c.font = BOLD; c.fill = HDR; c.alignment = wrap
wsC.row_dimensions[3].height = 32
widths = [5, 44, 9, 9, 5, 9] + [9, 6, 9] * 6 + [10, 9, 70, 40, 60]
for j, w in enumerate(widths, 1): wsC.column_dimensions[L(j)].width = w
PCOM = {
 "K01": "Basis differs: CRWV IT-like active MW; IREN 40 MW AI (basis unstated); Oracle no installed MW; Nscale IT MW at 5 active sites (H1'26 x2); Lambda Q2'25 (press MW); Crusoe FY26E rev / ~1 GW GROSS.",
 "K02b": "CoreWeave K02b = K02 per crwv_kpis.xlsx K02b row (active power ~ billed; no separate billing series). IREN/Oracle/Nscale/Lambda/Crusoe disclose no billed-MW series -> n/d (not inferred). Median n=1: treat as directional.",
 "K06": "IREN latest Q3'25. Lambda/Nscale ESTIMATED from GPU counts.",
 "K07": "Nebius realized $3.29 includes non-GPU-hour revenue and uses modeled GPUs.",
 "K09": "Not like-for-like (see NONCOMP list in README).",
 "K13": "Oracle FY26 capex / delivered MW (E). Nscale/Lambda/Crusoe n/d.",
 "K14": "Crusoe 12.5 = Abilene JV shell+power only (excluded from median).",
 "K16": "Lives: CRWV 6, Oracle 6, Nscale 5-6, IREN 5, Nebius 5 (from 2026).",
 "K17": "IREN incl. mining hardware; Oracle Aug-26 quarter; Nscale H1'26.",
 "K18": "CoreWeave K18 read straight from crwv_kpis.xlsx K18 (= 100/K17; pack fixed 2026-10-08 - was 100/K16 = 16.67 after its K02b row insert). Build asserts K18 = 100/K17 in every public pack.",
 "K20": "Oracle RPO is company-wide incl. SaaS.", "K21": "Crusoe/Nscale company TCV metrics (not ASC 606).",
 "K23": "Nscale 201 yrs: tiny H1'26 revenue vs $56.4B RPO (pre-ramp).",
 "K24": "Nscale H1'26 largest customer 52%. Lambda: Amazon+Microsoft 50% (K26, press).",
 "K28": "Nscale 18.7% on legacy $137M (excluded from median); marginal secured: Nscale ~6.25% E, Lambda 6.78% R, Crusoe ~9.1% E (PX09).",
 "K36": "IREN finance expense incl. amortization; Oracle interest/avg carrying.",
 "K43": "Oracle allocates all incremental OCI revenue to delivered AI MW (E).",
 "K44": "CRWV denominator = technology equipment (broader than GPUs).",
}
r = 4; ROWOF = {}; FLAGS = {}
for kid in ORDER:
    ROWOF[kid] = r
    wsC.cell(row=r, column=1, value=kid); wsC.cell(row=r, column=2, value=names.get(kid, "")); wsC.cell(row=r, column=3, value=UNIT.get(kid, ""))
    if kid in NBF:
        fml, t, a, note = NBF[kid]
        c = wsC.cell(row=r, column=4, value=fml); c.font = tagfont(t) if t != "n.m." else GREY; c.number_format = "0.00"
        wsC.cell(row=r, column=5, value=t).font = GREY; wsC.cell(row=r, column=6, value=a).font = GREY
        wsC.cell(row=r, column=28, value=note).alignment = wrap
    else:
        wsC.cell(row=r, column=4, value="n/d").font = GREY; wsC.cell(row=r, column=5, value="n/d").font = GREY
        wsC.cell(row=r, column=28, value="Not disclosed in Nebius filings / not computable on the same definition").alignment = wrap
    numcols = []
    for k, p in enumerate(PEERS):
        v, t, a = cells[(kid, p)]; cc = 7 + 3 * k
        c = wsC.cell(row=r, column=cc, value=v if v is not None else "n/d"); c.number_format = "0.00"
        c.font = tagfont(t) if v is not None else GREY
        wsC.cell(row=r, column=cc + 1, value=t if v is not None else "n/d").font = GREY
        wsC.cell(row=r, column=cc + 2, value=a).font = GREY
        if v is not None and (p, kid) not in NONCOMP: numcols.append(f"{L(cc)}{r}")
        if (p, kid) in NONCOMP: c.fill = YEL
    if numcols:
        wsC.cell(row=r, column=25, value=f"=MEDIAN({','.join(numcols)})").number_format = "0.00"
    else:
        wsC.cell(row=r, column=25, value="n/d").font = GREY
    f, med, gap, why, pp = flag(kid); FLAGS[kid] = (f, med, gap, why)
    c = wsC.cell(row=r, column=26, value=f); c.font = BOLD
    c.fill = GRN if f == "Better" else (RED if f == "Worse" else (YEL if f == "In line" else SEC))
    wsC.cell(row=r, column=27, value=why).alignment = wrap
    pc = PCOM.get(kid, "")
    nc = "; ".join(f"{p}: {NONCOMP[(p, kid)]}" for p in PEERS if (p, kid) in NONCOMP)
    wsC.cell(row=r, column=29, value="; ".join(x for x in [pc, ("Excluded from median: " + nc) if nc else ""] if x)).alignment = wrap
    r += 1
r += 1
for t in [
 "Colours: blue = Reported, purple = Press-reported (P, new for private peers), black = Derived, orange = Estimated, grey = n/d / n.m.; yellow fill = shown but excluded from the peer median (not same definition). Composite tags such as 'D (inputs P)' are kept verbatim from private_peers.",
 "Flag rule: Nebius vs MEDIAN of comparable peer values (all 6 peers, latest available period). Better/Worse if > 10% away in the favourable/unfavourable direction, In line otherwise; 'Context' rows are not scored (size, build-out multiple, list price, structure).",
 "Periods: CRWV/IREN Jun-26; Oracle May-26 (FQ4 FY26) or Aug-26 where May n/d; Nscale H1'26 / Aug-26; Lambda mostly Q2'25 (press, no filings); Crusoe Sep-26 (FY26E revenue, GROSS MW). Mixed vintages: treat Lambda/Crusoe as directional.",
 "Nebius MW: Capacity & Power base-case energized MW (Estimated, mixed IT/gross; 181.5 / 193 = current Capacity lock Quarterly MW!M130/N130, as selected in model v0.9e Sites tab (capacity v2); was 171.5 / 173 from outdated nebius_capacity_v1.xlsx). K02 uses AVERAGE energized MW (187.25) = 12.28 E; the old Comparison used group revenue on YE25 170 MW (13.7).",
 "K02b BILLED-IT BASIS: revenue / AVERAGE Capacity BILLING IT MW (Q1'26 165.5, Q2'26 186.8 -> 176.1 IT MW; capacity/nebius_capacity.xlsx row 147) - NOT energized or gross MW. CoreWeave K02b = K02 (active ~ billed); other peers n/d. This is the basis of every v0.9d Judge_Slide '$ per billed MW' figure.",
 f"SHARE COUNT: N_FD_DEBT {FD_DEBT:.2f}M = new Feb-2030/Feb-2034 converts as debt (share price < ${FDC['PX_2030']:.2f}); N_FD_IFCONV {FD_IFCONV:.2f}M = if-converted (>= ${FDC['PX_2034']:.2f}; {FD_MID:.2f}M in between, 2030s only) - fd_share_bridge.md. N_FD_V09D {FD_V09D:.2f}M = v0.9d historical. Never mix share and debt sides. No per-share or EV/MW metric in this workbook.",
 "Nebius Q3'26 results date: ~Nov 10 (est.) - not announced by the company; Q3'26 columns elsewhere (Sites) are capacity quarters, not the results date.",
 f"BILLED-MW LOCKS (Nebius_Inputs N_BILL_*, model v0.9e): YE26 {LOCK['YE26_BASE']:.1f} base / {LOCK['YE26_HIGH']:.1f} high; FY27 average {LOCK['FY27_BASE']:.2f} base / {LOCK['FY27_HIGH']:.2f} high, engine-linked ({DR} B34/C34). Forward $/billed MW memo (not scored): $7B/258 = 27.2, $9B/284 = 31.7, Street $12.3B/{LOCK['FY27_BASE']:.1f} = {BLENDED:.1f} blended, Street EX-MICROSOFT = {HURDLE:.1f} (USE AGAINST PEERS), /{LOCK['FY27_HIGH']:.1f} = {LOCK['STREET_FY27']/LOCK['FY27_HIGH']:.1f}; $16M coverage {MV('COV16_BASE'):.2f}x / {MV('COV16_HIGH'):.2f}x.",
 "Sites: see 'Sites' tab (peers/sites/sites.xlsx + capacity/site_verification.md) for announced vs contracted vs evidenced-active MW.",
]:
    wsC.cell(row=r, column=1, value=t).alignment = wrap; wsC.merge_cells(start_row=r, start_column=1, end_row=r, end_column=29); wsC.row_dimensions[r].height = 30; r += 1
wsC.freeze_panes = "D4"

# Ranking
wsR["A1"] = "Where Nebius is better / worse - ranked by signed gap vs comparable-peer median (positive = favourable to Nebius). Context and n/a rows listed last. Recomputed 2026-10-05 ET after v0.9d Nebius_Inputs refresh (K02b now scored in rank order); rebuilt 2026-10-08 ET (CRWV K18 read from fixed pack; share-count convention 358.56M debt / 376.65M if-converted; ~10:00 ET Nebius locks re-pointed to model v0.9e - no ranked KPI uses shares or forward MW, so ranks unchanged)."; wsR["A1"].font = BOLD
for j, h in enumerate(["Rank", "ID", "KPI", "Flag", "Nebius", "Peer median", "Signed gap % (favourable +; vs max(|median|, mean |peer|), computed at build)", "Reason"], 1):
    c = wsR.cell(row=3, column=j, value=h); c.font = BOLD; c.fill = HDR
scored = sorted([k for k in ORDER if FLAGS[k][0] in ("Better", "Worse", "In line")], key=lambda k: -FLAGS[k][2])
rest = [k for k in ORDER if k not in scored]
rr = 4
for i, kid in enumerate(scored + rest, 1):
    f, med, gap, why = FLAGS[kid]; rc = ROWOF[kid]
    vals = [i if kid in scored else "", kid, names.get(kid, ""), f"=Comparison!Z{rc}", f"=Comparison!D{rc}", f"=Comparison!Y{rc}",
            (round(gap, 1) if kid in scored else ""), why]
    for j, x in enumerate(vals, 1):
        c = wsR.cell(row=rr, column=j, value=x)
        if j in (5, 6, 7): c.number_format = "0.0"
    wsR.cell(row=rr, column=8).alignment = wrap
    wsR.cell(row=rr, column=4).fill = GRN if f == "Better" else (RED if f == "Worse" else (YEL if f == "In line" else SEC))
    rr += 1
for col, w in zip("ABCDEFGH", [5, 5, 48, 9, 9, 10, 12, 110]): wsR.column_dimensions[col].width = w

# ---------------------------------------------------------------- Contracts_$MWyr
CE = "parts/contract_economics.xlsx (GPU Market agent)"
#  id, status, provider, customer, category, signed, TCV $M, TCV tag, term, MW stated, MW basis, IT MW, IT MW method, chip, GPUs, GPU tag, mkt $M/MW-yr stated, mkt $M/IT MW-yr, mkt $/GPU-hr, primary source, notes
K = [
 ("N1","Reused","Nebius","Microsoft","Neocloud-hyperscaler","2025-09-08",17392.9,"R",5,"n/d (not stated)","not stated by Nebius/Microsoft (300 MW = Vineland site size)",240,"E (house): model v0.9e Inputs!F216 MSFT_MW = 240 IT MW est.; $/IT MW = Contracts!E12","GB300 NVL72 (P)",100000,"P (floor)",None,None,3.971,"6-K 2025-09-08 (SEC) | model/nebius_model.xlsx v0.9e Contracts!E12 | "+CE+" Deals/Normalized N1","$14.49M per IT MW [E, house] (model Contracts!E12 = 17,392.9 / (5 x 240 IT MW est.)); agrees with the Drivers 14.5 input (Inputs!F111). HISTORY (superseded Oct 9 2026, not current values): 11.60 on 300 MW (site size, not stated contract MW); 17.08 on GPU-derived 203.7 IT MW (100k GB300 / 491; retired by GPU Market); market prepay-adjusted 18.56 rests on the same GPU-derived MW."),
 ("N1o","HISTORY (superseded)","Nebius","Microsoft (incl. option) - HISTORY","Neocloud-hyperscaler","2025-09-08",19400,"R (up to)",5,300,"unlabeled",203.666,"as N1","GB300 (P)",100000,"P (floor)",12.933,19.051,4.429,CE+" N1o","HISTORY - superseded: 300 MW is gross site size, not stated contract MW (Vineland 'scalable up to 300 MW', 2024 20-F; Nebius/Microsoft state no contract MW). 12.93 = $19,400M / 5 / 300 is retired (Oct 9 2026 ET); not a current value."),
 ("N2","Reused","Nebius","Meta #1","Neocloud-hyperscaler","2025-11-01",2880.7,"R",5,"n/d","not disclosed",50,"E: ~50 IT MW team estimate (model Contracts!D23:D24 = 48.0 MW; Capacity tranches list n/d)",
  "n/d",None,"",None,11.523,None,"6-K 2025-11-12 | "+CE+" N2","IT MW is a low-confidence estimate (market); Nebius filings tie no MW to the contract."),
 ("N3","Reused","Nebius","Meta #2 (dedicated)","Neocloud-hyperscaler","2026-03-13",12000,"R",5,"n/d","not disclosed",None,"","Vera Rubin NVL72",None,"",None,None,None,"6-K 2026-03-16 | "+CE+" N3","Break-even ~149 IT MW to match Anthropic-Nscale $16.16M (market)."),
 ("N4","Reused","Nebius","Meta #2 backstop (contingent)","Neocloud-hyperscaler","2026-03-13",15000,"R (up to, contingent)",5,"n/d","",None,"","n/d",None,"",None,None,None,"6-K 2026-03-16 | "+CE+" N4","Contingent; excluded from normalized comps."),
 ("NAGG1","Reused (KPI)","Nebius","4 'landmark' Q2'26 customers (unnamed)","Neocloud-mixed","Q2'26","n/d","",None,"n/d","",None,"","n/d",None,"","20-25 (ACV/MW, R)",None,None,"Q2'26 letter p.2-4 (filings MW_Disclosures) | "+CE+" Notes_Method","Company aggregate: deals averaged >$1B TCV, ACV $20-25M/MW; Q2 deals >$20M/MW vs 2026 base $12M/MW; Q3 short-term >$40M/MW. Basis not stated (Filings: likely IT/active)."),
 ("P1","Reused","IREN","Microsoft","Neocloud-hyperscaler","2025-11-03",9700,"R",5,200,"IT",200,"R","GB300 NVL72",98200,"E (200 x 491)",9.7,9.7,2.255,"8-K 2025-11-03 | "+CE+" P1","Prepay-adjusted 10.28 (market)."),
 ("NEW1","New","IREN","NVIDIA","Neocloud-chip vendor","2026-05 (8-K)",3400,"R",5,60,"unlabeled (existing Childress halls; air-cooled)",None,"","NVIDIA Blackwell (air-cooled)",None,"n/d",None,None,None,"https://iren.gcs-web.com/news-releases/news-release-details/iren-secures-34bn-ai-cloud-contract-nvidia ; peers/iren (FY26 10-K backlog)","~60 MW of existing Childress capacity; ramp from early 2027 (peers/sites site_verification.md). Below IREN's 'recent 3-yr >$20M/IT MW' claim."),
 ("PAGG","Reused (KPI)","IREN","recent 3-yr contracts (unnamed)","Neocloud-mixed","FY26","n/d","",3,"n/d","IT",None,"","n/d",None,"",">20 (R)",None,None,CE+" Notes_Method","Company statement only."),
 ("P2","Reused","Nscale","Microsoft (all SOWs)","Neocloud-hyperscaler","2025-09..2026-04",43800,"R (up to)",5.5,"n/a","",596.474,"D: 191.6k/491 + 66k/320","GB300 + Vera Rubin",257600,"R (sum)",None,13.351,3.529,"Nscale S-1 2026-09-18 | "+CE+" P2",""),
 ("P2a","Reused","Nscale (Aker JV)","Microsoft - Narvik","Neocloud-hyperscaler","2025-09-17",6200,"R",5,"n/a","",105.906,"D: 52k/491","GB300",52000,"R",None,11.708,2.722,"akernscale.com PR | "+CE+" P2a","Within P2."),
 ("P2b","Reused","Nscale","Microsoft ex-Narvik","Neocloud-hyperscaler","",37600,"D",5.5,"n/a","",490.568,"D","GB300 + Vera Rubin",205600,"R (sum)",None,13.936,3.796,CE+" P2b","Residual."),
 ("O1","Reused","Nscale","Anthropic","Neocloud-AI lab","2026-08-25",44600,"R (up to)",6,460,"IT",460,"R (S-1 site #17)","Vera Rubin NVL72",147200,"E (460 x 320)",16.159,16.159,5.765,"Nscale S-1 | "+CE+" O1",""),
 ("P3","Reused","Lambda","Microsoft","Neocloud-hyperscaler","2025-11-03","n/d","'multibillion'",None,"n/d","",None,"","GB300 NVL72 (partly)",None,"",None,None,None,"lambda.ai blog 2025-11-03 | "+CE+" P3","No economics public."),
 ("NEW2","New","Lambda","Anthropic","Neocloud-AI lab","2026-09-01 (press)",35000,"P",6,350,"unlabeled (Hut 8 Beacon Point TX)",None,"","n/d",None,"",None,None,None,"WSJ via theenergymag.com (peers/private private_peers.xlsx L21/L12)","Not confirmed by Lambda; NVIDIA holds the Hut 8 lease (row F6). TOP OF PEER COMP BAND $9.7-16.7M/MW-yr (2026-10-08): 35,000 / 6 / 350 = 16.67; Medium confidence: press-only TCV (WSJ via theenergymag; private_peers PX10 16.7 D on P inputs), 6-yr term secondary, IT vs gross MW unstated. Band low = IREN-Microsoft 9.7 (P1)."),
 ("NEW3","New","Lambda","NVIDIA (GPU leaseback, main tranche)","Neocloud-chip vendor","2025-09-08 (press)",1300,"P",4,"n/d","",None,"","NVIDIA GPUs (type n/d)",10000,"P",None,None,None,"DCD citing The Information (private_peers L18/L19)","Total reported $1.5B / ~18k GPUs (incl. +$200M / 8k GPUs) -> $2.38/GPU-hr on the total; private_peers uses the $1.3B/10k tranche = $3.71."),
 ("P4","Reused","CoreWeave","Meta #1","Neocloud-hyperscaler","2025-09-30",14200,"R (up to)",6.22,"n/d","",None,"","GB300 (P)",None,"",None,None,None,"8-K 2025-09-30 | "+CE+" P4",""),
 ("P5","Reused","CoreWeave","Meta #2","Neocloud-hyperscaler","2026-04-09",21000,"R (up to)",6.73,"n/d","",None,"","incl. Vera Rubin",None,"",None,None,None,"8-K | "+CE+" P5",""),
 ("P6","Reused","CoreWeave","OpenAI T1","Neocloud-AI lab","2025-03-10",11900,"R (up to)",5,"n/d","",None,"","n/d",None,"",None,None,None,CE+" P6","$350M CRWV equity to OpenAI."),
 ("P7","Reused","CoreWeave","OpenAI T2","Neocloud-AI lab","2025-05-15",4000,"R (up to)",3.9,"n/d","",None,"","n/d",None,"",None,None,None,CE+" P7",""),
 ("P8","Reused","CoreWeave","OpenAI T3","Neocloud-AI lab","2025-09-25",6500,"R (up to)",5.68,"n/d","",None,"","n/d",None,"",None,None,None,CE+" P8","OpenAI total ~$22.4B."),
 ("NEW4","New","CoreWeave","Jane Street","Neocloud-other","2026-04-15",6000,"R",None,"n/d","",None,"","Vera Rubin (part)",None,"",None,None,None,"https://www.coreweave.com/news/jane-street-signs-6-billion-ai-cloud-agreement-with-coreweave","Term/MW not disclosed; plus $1B equity investment by Jane Street."),
 ("NEW5","New","CoreWeave","Anthropic","Neocloud-AI lab","2026-04-10","n/d","R (no value)",None,"n/d","",None,"","n/d",None,"",None,None,None,"https://coreweave.com/news/coreweave-announces-multi-year-agreement-with-anthropic","Multi-year; no value."),
 ("NEW6","New","Crusoe","Jane Street","Neocloud-other","2026-09-04 (press)",13000,"P",5,"n/d","",None,"","n/d",None,"",None,None,None,"DCD citing Bloomberg (private_peers C17)",""),
 ("NEW7","New","Crusoe","Meta","Neocloud-hyperscaler","2026-06-19 (press)","n/d","",None,1600,"unlabeled (Childress TX + Warrenton MO)",None,"","n/d",None,"",None,None,None,"DCD citing Bloomberg (private_peers C14)","Value/term n/d."),
 ("NEW8","New","Crusoe","Microsoft (Abilene 900 MW campus)","Neocloud-hyperscaler","2026-03-27","n/d","",None,900,"gross (2 x 336 MW IT = 672 IT)",672,"R","n/d",None,"",None,None,None,"Crusoe PR 2026-03-27 (private_peers C13)","First building mid-2027; value n/d."),
 ("NEW9","New","Crusoe (JV) / Oracle","Oracle for OpenAI (Abilene Stargate)","Developer-hyperscaler","2025-05-21","n/d","",None,1200,"gross (~824 IT implied = 618 / 0.75 per peers/sites appendix; Epoch ~843)",None,"","GB200 (up to 50k per bldg)",None,"",None,None,None,"Crusoe JV PR 2025-05-21 (private_peers C12); peers/sites site_verification.md","All 8 buildings leased to Oracle; contract value n/d (the $15B is JV project cost, not contract value). Economics flow through O2."),
 ("O2","Reused","Oracle","OpenAI (Stargate)","Hyperscaler-AI lab","2025-07/09 (press)",300000,"P (WSJ)",5,4500,"unlabeled",3750,"E: /1.2 PUE if gross","Mixed (GB200 etc.)",None,"",13.333,16.0,None,"WSJ | "+CE+" O2","Not named in Oracle filings."),
 ("NEW10","New","Oracle","Meta","Hyperscaler-hyperscaler","2025-10-16 (confirmed)",20000,"P (talks, Bloomberg)",None,"n/d","",None,"","n/d",None,"",None,None,None,"Bloomberg/Reuters 2025-09-19; CNBC 2025-10-16 (Oracle confirmed Meta among $65B new commitments)","Value ~ $20B is press (in talks); term 'multi-year'."),
 ("O3","Reused","SpaceX (xAI)","Anthropic","AI lab-AI lab","2026 (424B4)","$1.25B/month","R",2.9,"n/d","",None,"","unspecified NVIDIA",325000,"R",None,None,5.269,"SpaceX 424B4 | "+CE+" O3","90-day termination; $/GPU-hr = 1,250e6/(325k x 730)."),
 ("NEW11","New","Microsoft (Azure)","OpenAI","Hyperscaler-AI lab","2025-10-28",250000,"R (incremental commitment)",None,"n/d","",None,"","n/d",None,"",None,None,None,"https://openai.com/index/next-chapter-of-microsoft-openai-partnership/","Term not disclosed."),
 ("NEW12","New","AWS","OpenAI","Hyperscaler-AI lab","2025-11-03",38000,"R",7,"n/d","",None,"","GB200 + GB300",None,"'hundreds of thousands'",None,None,None,"https://openai.com/index/aws-and-openai-partnership/","All capacity targeted by end-2026."),
 ("NEW13","New","Microsoft (Azure)","Anthropic","Hyperscaler-AI lab","2025-11-18",30000,"R",None,"n/d (plus up to 1 GW additional)","",None,"","Grace Blackwell / Vera Rubin",None,"",None,None,None,"https://www.anthropic.com/news/microsoft-nvidia-anthropic-announce-strategic-partnerships","1 GW is additional capacity, not tied to the $30B -> no $/MW computed."),
 ("NEW14","New","Google Cloud","Anthropic","Hyperscaler-AI lab","2025-10-23","n/d ('tens of billions')","P",None,">1,000 (2026)","unlabeled",None,"","TPU (up to 1M)",None,"",None,None,None,"https://www.anthropic.com/news/expanding-our-use-of-google-cloud-tpus-and-services ; Reuters 2025-10-23","Expanded Apr-2026 (multi-GW from 2027, Google/Broadcom)."),
 ("NEW15","New","Google Cloud","Meta","Hyperscaler-hyperscaler","2025-08-21 (press)",10000,"P (>$10B)",6,"n/d","",None,"","n/d",None,"",None,None,None,"Reuters/CNBC 2025-08-21",""),
 ("NEW16","New","Cerebras","OpenAI","Chip vendor cloud-AI lab","2026-01-14",10000,"P (>$10B)",3,750,"unlabeled (ramp 250/500/750 by YE26/27/28)",None,"","Cerebras WSE",None,"",None,None,None,"Reuters/CNBC 2026-01-14; MRA on SEC (Cerebras filing)","NAIVE: ramp means average MW << 750 -> true $/MW-yr higher; lower bound."),
 # facility / colocation leases (not compute prices)
 ("REF","Reused","TeraWulf","Fluidstack (Google backstop)","Facility lease","2025-08-14",3700,"R",10,200,"IT",200,"R","n/a",None,"",1.85,1.85,None,CE+" REF","Facility only."),
 ("F1","New","Cipher Mining","AWS","Facility lease","2025-11-03",5500,"R",15,300,"gross (Davis Polk)",None,"","n/a",None,"",None,None,None,"https://www.sec.gov/Archives/edgar/data/1819989/000181998925000110/nov3_earningspressreleasex.htm","Turnkey space+power; rent from Aug-2026."),
 ("F2","New","Cipher Mining","Fluidstack (Google backstop $1.4B)","Facility lease","2025-09-25",3000,"R",10,168,"IT",168,"R","n/a",None,"",None,None,None,"https://www.globenewswire.com/news-release/2025/09/25/3156107/0/en/cipher-mining-signs-168-mw-10-year-ai-hosting-agreement-with-fluidstack.html","+39 MW IT add-on Nov-2025 (total ~$3.83B)."),
 ("F3","New","Hut 8","Fluidstack (Anthropic; Google backstop)","Facility lease","2025-12-17",7000,"R",15,245,"IT",245,"R","n/a",None,"",None,None,None,"https://www.prnewswire.com/news-releases/hut-8-signs-15-year-245-mw-ai-data-center-lease-at-river-bend-campus-with-total-contract-value-of-7-0-billion-302644600.html","Triple-net, 3% escalator."),
 ("F4","New","Applied Digital","CoreWeave (Polaris Forge 1)","Facility lease","2025-06..08",11000,"R",15,400,"IT",400,"R","n/a",None,"",None,None,None,"https://ir.applieddigital.com/news-events/press-releases/detail/128/applied-digital-finalizes-additional-150mw-lease-with","3 leases ~15 yrs."),
 ("F5","New","Galaxy","CoreWeave (Helios)","Facility lease","2025","~15,000 (>$1B/yr avg)","D (R: >$1B/yr)",15,526,"IT (800 gross)",526,"R","n/a",None,"",None,None,None,"https://www.galaxy.com/newsroom/galaxy-completes-phase-i-of-its-helios-data-center-campus","TCV derived = >$1B/yr x 15 (lower bound)."),
 ("F6","New","Hut 8","NVIDIA (Beacon Point; for Lambda/Anthropic)","Facility lease","2026-09 (press)",19600,"P",15,704,"unlabeled",None,"","n/a",None,"",None,None,None,"theenergymag.com (private_peers L12 note)",""),
 ("F7","New","AIB Data Centers","Nebius (CLT1)","Facility lease (Nebius tenant)","2026-09-30","n/d","",12,50,"IT",50,"R","n/a",None,"",None,None,None,"https://www.globenewswire.com/news-release/2026/09/30/3371855/0/en/aib-data-centers-signs-contract-with-nebius-for-ai-data-center-capacity.html ; capacity/site_list.xlsx row 23","Value n/d; Red grade (capacity)."),
]
wsK["A1"] = ("Hyperscaler / neocloud / AI-lab contract economics: implied $M per MW-yr = TCV / term / MW (8,760 h/yr for $/GPU-hr). 'Reused' rows are the GPU Market agent's entries in "
             "parts/contract_economics.xlsx (Deals/Normalized; nothing contract-related exists in market/ or nebius_market.xlsx) - cited, not re-researched; their values shown in cols R-T and checked vs ours in col U. "
             "'New' rows added here with sources. Facility leases at the bottom are NOT compute prices (shell+power only). Never fabricated: n/d where undisclosed.")
wsK["A1"].alignment = wrap; wsK.merge_cells("A1:X1"); wsK.row_dimensions[1].height = 60
KH_ = ["ID", "Status", "Provider", "Customer", "Category", "Signed", "TCV ($M)", "TCV tag", "Term (yrs)", "MW (as stated)", "MW basis", "IT MW (normalized)", "IT MW method",
       "Chip", "GPUs", "GPU tag", "TCV / yr ($M)", "Market agent $M/MW-yr (stated MW)", "Market agent $M/IT MW-yr", "Market agent $/GPU-hr", "Check: ours - market ($M/MW-yr)",
       "$M per MW-yr (stated MW)", "$M per IT MW-yr", "$ per GPU-hr", "Source", "Notes / conflicts"]
for j, h in enumerate(KH_, 1):
    c = wsK.cell(row=3, column=j, value=h); c.font = BOLD; c.fill = HDR; c.alignment = wrap
wsK.row_dimensions[3].height = 45
rk = 4; KROW = {}
for row in K:
    (iid, st, prov, cust, cat, sig, tcv, ttag, term, mw, mwb, itmw, itm, chip, gpus, gtag, m1, m2, m3, src, note) = row
    KROW[iid] = rk
    if iid == "F5": tcv = 15000
    vals = [iid, st, prov, cust, cat, sig, tcv, ttag, term, mw, mwb, itmw, itm, chip, gpus, gtag]
    for j, x in enumerate(vals, 1): wsK.cell(row=rk, column=j, value=x)
    if iid == "O3":
        wsK.cell(row=rk, column=17, value="=1250*12")
        wsK.cell(row=rk, column=24, value="=1250*1000000/(O%d*730)" % rk)
    else:
        wsK.cell(row=rk, column=17, value=f'=IF(AND(ISNUMBER(G{rk}),ISNUMBER(I{rk})),G{rk}/I{rk},"n/d")')
        wsK.cell(row=rk, column=24, value=f'=IF(AND(ISNUMBER(G{rk}),ISNUMBER(I{rk}),ISNUMBER(O{rk})),G{rk}*1000000/(O{rk}*I{rk}*8760),"n/d")')
    for j, x in zip((18, 19, 20), (m1, m2, m3)): wsK.cell(row=rk, column=j, value=x if x is not None else "")
    wsK.cell(row=rk, column=22, value=f'=IF(AND(ISNUMBER(Q{rk}),ISNUMBER(J{rk})),Q{rk}/J{rk},"n/d")')
    wsK.cell(row=rk, column=23, value=f'=IF(AND(ISNUMBER(Q{rk}),ISNUMBER(L{rk})),Q{rk}/L{rk},"n/d")')
    wsK.cell(row=rk, column=21, value=f'=IF(AND(ISNUMBER(S{rk}),ISNUMBER(W{rk})),ROUND(W{rk}-S{rk},2),IF(AND(ISNUMBER(R{rk}),ISNUMBER(V{rk})),ROUND(V{rk}-R{rk},2),""))')
    wsK.cell(row=rk, column=25, value=src); wsK.cell(row=rk, column=26, value=note).alignment = wrap
    for j in (17, 18, 19, 20, 21, 22, 23, 24): wsK.cell(row=rk, column=j).number_format = "0.00"
    wsK.cell(row=rk, column=7).font = tagfont(ttag); wsK.cell(row=rk, column=2).fill = SEC if st.startswith("Reused") else GRN
    if cat.startswith("Facility"): wsK.cell(row=rk, column=5).fill = YEL
    rk += 1
rk += 1
CONF = [
 "CONFLICTS / DIFFERENCES vs the GPU Market agent's data (parts/contract_economics.xlsx):",
 "1) Location: the request pointed to /workspace/nebius/market/ and nebius_market.xlsx; neither holds contract data (price history/scrapers only). The contract table is parts/contract_economics.xlsx - used as the reuse source.",
 "2) Microsoft-Nebius MW basis (ruling Oct 9 2026): shown only as $14.49M per IT MW [E, house] = model v0.9e Contracts!E12 = $17,392.9M / (5 x 240 IT MW est.). Nebius filings tie no MW to the contract. HISTORY (superseded): the market headline 17.08 (GPU-derived 203.7 IT MW, retired by GPU Market) and 11.60 (on 300 MW, which is Vineland's site size, not stated contract MW).",
 "3) Meta-Nebius #1: 11.52 [E] on a ~50 IT MW team estimate (model Contracts!D23:D24 = 48.0 MW; Capacity tranches list n/d; GPU Market Print_Table #3). Nebius filings (MW_Disclosures) state no MW. (v3 wording, Oct 9.)",
 "4) Lambda-NVIDIA leaseback (new row): reported $1.5B / ~18k GPUs / 4 yrs -> $2.38/GPU-hr overall vs $3.71 on the $1.3B / 10k tranche used in private_peers (PX13). Same source, different slicing.",
 "5) IREN: company claims recent 3-yr contracts >$20M/IT MW (market) but the NVIDIA contract (new) is ~$11.3M/MW-yr on ~60 MW (unlabeled) and Microsoft $9.7M -> the >$20M deals are not the two named anchors.",
 "6) Oracle-OpenAI $300B (market O2) is press-only; Oracle filings/peers/oracle do not name it; RPO $638B (May-26) is consistent in scale. Oracle MW basis: ~353 GPUs/MW in peers/oracle implies GROSS, matching market's 16.0 'if gross' variant.",
 "7) Nscale: market P2 13.35 $M/IT MW-yr (Microsoft) vs private_peers PX10 13.24 (all contracted TCV / 5.7 yrs / 1,370 MW) - consistent (<1%).",
 "8) Cerebras-OpenAI 4.44 is a naive lower bound (750 MW is the end-2028 ramp level).",
 "9) COMP BAND (2026-10-08 ET, coordinator correction): peer/hyperscaler compute-contract band = $9.7-16.7M per MW-yr. Low = IREN-Microsoft 9.7 (P1, R, IT MW). Top = Lambda-Anthropic ~16.7 (NEW2: $35B / 6 yrs / ~350 MW = 16.67; Medium confidence: press-only TCV (WSJ via theenergymag; private_peers PX10 16.7 D on P inputs), 6-yr term secondary, IT vs gross MW unstated). Was $9.7-16.2M with Nscale-Anthropic 16.16 (O1) on top. Band = memo set P1/NEW1/P2/O1/NEW2/O2 (build asserts min/max). Not a Nebius KPI; Microsoft-Nebius rows and facility leases excluded.",
]
_E12 = openpyxl.load_workbook(f"{W}/model/nebius_model.xlsx", data_only=True)["Contracts"]["E12"].value
_n1 = next(x for x in K if x[0] == "N1"); assert abs(_n1[6] / _n1[8] / _n1[11] - _E12) < 1e-9 and abs(_E12 - 14.494) < 0.001, ("N1 vs model Contracts!E12", _E12)
_num = lambda x: isinstance(x, (int, float))
_BAND = {x[0]: x[6] / x[8] / (x[11] if _num(x[11]) else x[9]) for x in K if x[0] in ("P1", "NEW1", "P2", "O1", "NEW2", "O2") and _num(x[6]) and _num(x[8]) and (_num(x[11]) or _num(x[9]))}
_blo, _bhi = min(_BAND, key=_BAND.get), max(_BAND, key=_BAND.get)
assert _blo == "P1" and _bhi == "NEW2" and abs(_BAND["P1"] - 9.7) < 0.05 and abs(_BAND["NEW2"] - 16.67) < 0.01, ("comp band drift", _BAND)
print("comp band", f"{_BAND[_blo]:.2f}-{_BAND[_bhi]:.2f}", _blo, _bhi)
for t in CONF:
    c = wsK.cell(row=rk, column=1, value=t); c.alignment = wrap
    if t.startswith("CONFLICTS"): c.font = BOLD
    wsK.merge_cells(start_row=rk, start_column=1, end_row=rk, end_column=26); wsK.row_dimensions[rk].height = 30; rk += 1
for j, w in enumerate([6, 9, 14, 22, 16, 13, 10, 12, 7, 9, 18, 9, 18, 16, 9, 10, 9, 10, 10, 9, 10, 10, 10, 9, 50, 60], 1): wsK.column_dimensions[L(j)].width = w
wsK.freeze_panes = "E4"

# ---------------------------------------------------------------- Sites
wsS["A1"] = ("Sites: announced vs contracted vs EVIDENCED-ACTIVE MW (low/base/high), YE26, grade and confidence. Peers: peers/sites/sites.xlsx 'Sites' + peers/sites/site_verification.md; the four one-page appendices (peers/sites/appendix_*.pdf, site_data.py; as of Oct 8) are cross-checked in cols Q-R (build asserts active and announced match). Abilene active 618 [D on R] per Oracle Sep 10 (IT vs gross basis unstated). "
             "Nebius: capacity/nebius_capacity.xlsx ('Sites' grades, 'Quarterly MW' Q3'26 = current evidenced energized; YE26 = Q4'26) + capacity/site_verification.md. Same method/grade scale for both (Green/Yellow/Red).")
wsS["A1"].alignment = wrap; wsS.merge_cells("A1:P1"); wsS.row_dimensions[1].height = 48
SH = ["Company", "Site", "MW basis", "Announced MW (text)", "Announced MW (numeric, for chart)", "Contracted MW", "Active low", "Active base", "Active high", "YE26 low", "YE26 base", "YE26 high", "Grade", "Confidence", "Evidence / key images", "Source", "Billed MW (site appendix)", "Appendix check / tags"]
def sh(r):
    for j, h in enumerate(SH, 1):
        c = wsS.cell(row=r, column=j, value=h); c.font = BOLD; c.fill = HDR; c.alignment = wrap
import importlib.util as _ilu
_sp = _ilu.spec_from_file_location("site_data", f"{W}/peers/sites/site_data.py"); _sd = _ilu.module_from_spec(_sp); _sp.loader.exec_module(_sd)
APPX = _sd.SITES; APPX_KEY = {"Denton": "coreweave_denton", "Childress": "iren_childress", "Abilene": "crusoe_abilene", "Port Washington": "oracle_port_washington"}
APPX_NOTE = {"Denton": "Matches appendix_coreweave_denton (as of Oct 8): active 250/260/262 (D), billed ~260 (R, landlord CORZ billing CoreWeave), YE26 260/262/262 (E). IT MW.",
             "Childress": "Matches appendix_iren_childress (as of Oct 8): active 50/75/100 AI IT (D), billed 50 (D on R, H1 accepted Aug 13), YE26 100/150/200 (E). Site grid 750 incl. mining.",
             "Abilene": "Matches appendix_crusoe_abilene (as of Oct 8): active 421/618/618 [D on R] - Oracle FQ1 FY27 call (Sep 10): 6 of 8 bldgs, 618 MW, 75% delivered; low 421 = Epoch (P). Billed ~618 (D). YE26 618/721/824 (E). IT vs gross basis unstated. History: Oct 5 view 200/421/421, YE26 421/632/843.",
             "Port Washington": "Matches appendix_oracle_port_washington (as of Oct 8): active 0/0/0 (R), billed 0 (R), YE26 0 (E); YE27 0/~310/438 = grid MW. Announced ~902 IT / 1.3 GW electrical."}
ANNNUM = {"Denton": 262, "Childress": 260, "Kvandal": 230, "Abilene": 824, "Kansas City": 24, "Port Washington": 902, "Monarch": 460}
wsS.cell(row=3, column=1, value="A. Peer key sites (peers/sites/sites.xlsx)").font = BOLD; sh(4)
sw = openpyxl.load_workbook(f"{W}/peers/sites/sites.xlsx", data_only=True)["Sites"]
# header-based column lookup (sites.xlsx gained H1'28 / permits / imagery / thermal columns at its 09:04 ET rebuild)
SX = {str(c.value).strip(): j for j, c in enumerate(sw[1]) if c.value}
_g = lambda r0, h: r0[SX[h]]
for h in ("Peer", "Site", "MW basis", "Announced MW", "Contracted MW", "Active low (MW)", "Active base (MW)", "Active high (MW)", "Active basis (evidence)", "YE26 low", "YE26 base", "YE26 high", "Grade", "Confidence", "Key images"):
    assert h in SX, ("sites.xlsx header missing", h)
rs = 5; SITECHART = []
for r0 in sw.iter_rows(min_row=2, values_only=True):
    if not r0[1]: continue
    ann = next((v for k, v in ANNNUM.items() if k in r0[1]), None)
    vals = [_g(r0, "Peer"), _g(r0, "Site"), _g(r0, "MW basis"), _g(r0, "Announced MW"), ann, _g(r0, "Contracted MW"), _g(r0, "Active low (MW)"), _g(r0, "Active base (MW)"), _g(r0, "Active high (MW)"),
            _g(r0, "YE26 low"), _g(r0, "YE26 base"), _g(r0, "YE26 high"), _g(r0, "Grade"), _g(r0, "Confidence"), f"{_g(r0, 'Active basis (evidence)')} | {_g(r0, 'Key images')}", "peers/sites/sites.xlsx Sites; site_verification.md"]
    for j, x in enumerate(vals, 1):
        try: x = float(x) if j in (7, 8, 9, 10, 11, 12) and x not in (None, "") else x
        except ValueError: pass
        c = wsS.cell(row=rs, column=j, value=x); c.alignment = wrap
    g = str(_g(r0, "Grade") or ""); wsS.cell(row=rs, column=13).fill = GRN if g.startswith("Green") else (RED if g.startswith("Red") else YEL)
    _ak = next((k for k in APPX_KEY if k in r0[1]), None)
    if _ak:  # 2026-10-08 ~23:15 ET: one-page site appendices (peers/sites/site_data.py) are the current site view; sites.xlsx must agree
        cd = APPX[APPX_KEY[_ak]]["chart"]
        assert (float(_g(r0, "Active low (MW)")), float(_g(r0, "Active base (MW)")), float(_g(r0, "Active high (MW)"))) == tuple(float(x) for x in cd["act"]), ("sites.xlsx vs appendix active drift", _ak)
        assert ann == cd["ann"] + (60 if _ak == "Childress" else 0), ("announced numeric vs appendix drift", _ak, ann, cd["ann"])  # Childress: appendix 200 IT (H1-4) + ~60 NVIDIA = 260
        wsS.cell(row=rs, column=17, value=float(cd["bill"])).alignment = wrap
        wsS.cell(row=rs, column=18, value=APPX_NOTE[_ak]).alignment = wrap
    SITECHART.append((f"{r0[0].split(' ')[0]}: {r0[1].split('(')[0].split(' - ')[0].split(',')[0].strip()}", rs)); rs += 1
wsS.cell(row=rs, column=1, value="Announced numeric = IT MW where available (Denton 262 IT; Childress AI 4x50 + 60 = 260 IT; Kvandal 230 grid; Abilene ~824 IT implied of 1.2 GW gross (618 MW delivered = 75%, Oracle Sep 10; was ~843 per Epoch, Oct 5); Lambda KC 24; Port Washington ~902 IT; Nscale Monarch 460 IT = S-1 contracted (Phase 1 announced ~1.35-1.37 GW IT)), read from the text column - Derived.").font = GREY
rs += 2
wsS.cell(row=rs, column=1, value="B. Nebius sites (capacity/nebius_capacity.xlsx + capacity/site_verification.md)").font = BOLD; rs += 1; sh(rs); rs += 1
cap = openpyxl.load_workbook(f"{W}/capacity/nebius_capacity.xlsx", data_only=True)
cs = cap["Sites"]; qm = cap["Quarterly MW"]
QM = {}
for r0 in qm.iter_rows(min_row=3, values_only=True):
    if r0[0] and r0[1]: QM.setdefault(r0[0], {})[r0[1][:18]] = r0
sl = openpyxl.load_workbook(f"{W}/capacity/site_list.xlsx", data_only=True)["Sites"]
SL = [r0 for r0 in sl.iter_rows(min_row=2, values_only=True) if r0[1]]
def fnum(x):
    try: return float(x)
    except (TypeError, ValueError): return x
nb_first = rs
for r0 in cs.iter_rows(min_row=2, values_only=True):
    if not r0[0]: continue
    site = r0[0]; q = QM.get(site, {})
    lo = q.get("Energized MW low", [None] * 20); hi = q.get("Energized MW high", [None] * 20); ba = next((v for k, v in q.items() if k.startswith("Energized MW base")), [None] * 20)
    key = site.split(" (")[0].split(" ")[0].replace("Mantsala", "Mäntsälä").replace("Bethune", "Béthune")
    ann = next((s[10] for s in SL if key.lower()[:6] in str(s[1]).lower()[:30]), "n/d")
    vals = ["Nebius", site, r0[6], ann, fnum(r0[5]), fnum(r0[5]), fnum(lo[14]), fnum(ba[14]), fnum(hi[14]), fnum(r0[12]), fnum(r0[14]), fnum(r0[13]), r0[10],
            "see site_verification.md (MW tags in nebius_capacity Tags_Sources)", f"{r0[9]}", "capacity/nebius_capacity.xlsx Sites + Quarterly MW (Q3'26/Q4'26)"]
    for j, x in enumerate(vals, 1): wsS.cell(row=rs, column=j, value=x).alignment = wrap
    g = str(r0[10] or ""); wsS.cell(row=rs, column=13).fill = GRN if g.startswith("Green") else (RED if g.startswith("Red") else YEL)
    if site.startswith(("Vineland", "Mantsala FI #1", "Kansas City", "London", "Bethune", "Independence")): SITECHART.append((f"Nebius: {site.split(' (')[0]}", rs))
    rs += 1
nb_last = rs - 1
wsS.cell(row=rs, column=2, value="Nebius identified-site total (formula)").font = BOLD
for j in (6, 7, 8, 9, 10, 11, 12):
    wsS.cell(row=rs, column=j, value=f"=SUM({L(j)}{nb_first}:{L(j)}{nb_last})").font = BOLD
wsS.cell(row=rs, column=16, value="Capacity check: Q3'26 identified 182/207.5/233; YE26 185/304.5/424 ENERGIZED (site_verification.md). Company base incl. plug 207.5 (Q3'26); mgmt guide 800 MW-1 GW CONNECTED YE26. "
    "Not the same as BILLING: billed IT MW locks YE26 257.7 base / 284.2 high; FY27 avg 531.8 / 659.2 engine-linked (model v0.9e; Nebius_Inputs N_BILL_*).")
rs += 2
for t in ["Nebius 'Announced MW' column = site_list.xlsx 'Announced MW - full build-out' text; numeric column = contracted MW (capacity Sites tab), basis mostly gross/unknown.",
          "Nebius active = Capacity & Power 'Energized MW' low/base/high for Q3'26 (current, as of Oct-4 evidence); YE26 = Q4'26. Peers active = evidence-supported energized and in service as of 2026-10-04.",
          "Read-across: only CoreWeave Denton (~260) and Crusoe Abilene (618 delivered, Oracle Sep 10; D on R; IT vs gross basis unstated) show hundreds of MW evidenced-active on grid; Nebius's largest site (Vineland) is Yellow/Red on unpermitted gensets; Nscale Kvandal and Oracle Port Washington are 0 active - like most new Nebius sites."]:
    wsS.cell(row=rs, column=1, value=t).alignment = wrap; wsS.merge_cells(start_row=rs, start_column=1, end_row=rs, end_column=16); wsS.row_dimensions[rs].height = 30; rs += 1
for j, w in enumerate([10, 34, 16, 34, 10, 10, 8, 8, 8, 8, 8, 8, 18, 22, 60, 36, 10, 60], 1): wsS.column_dimensions[L(j)].width = w

# ---------------------------------------------------------------- Chart_Data (formulas -> Comparison / Contracts / Sites) + native charts
COLS = ["Nebius"] + PEERS
def ccol(name): return "D" if name == "Nebius" else L(7 + 3 * PEERS.index(name))
wsD["A1"] = "Chart data (live formulas to Comparison / Contracts_$MWyr / Sites). Blank = n/d. PNG versions: peers/charts/*.png (peers/build_charts.py)."; wsD["A1"].font = BOLD
CH = [("rev_per_active_mw", "Revenue per active / billing MW ($M/MW-yr)", ["K01", "K02", "K02b"]),
      ("rev_per_gpu_hr", "Realized $/GPU-hr vs list", ["K07", "K08"]),
      ("capex_per_mw", "Capex / PP&E per MW ($M/MW)", ["K13", "K14", "K15"]),
      ("cost_of_debt", "Cost of debt (%)", ["K28", "K36"]),
      ("top_customer", "Customer concentration (%)", ["K24", "K26"]),
      ("rpo_coverage", "RPO / annualized revenue (years)", ["K23"]),
      ("depreciation", "Depreciation yield (%) and useful life (yrs)", ["K17", "K16"])]
rdd = 3; CHROWS = {}
for key, title, kids in CH:
    wsD.cell(row=rdd, column=1, value=title).font = BOLD
    for j, nm in enumerate(COLS, 2): wsD.cell(row=rdd, column=j, value=nm).font = BOLD
    start = rdd
    for kid in kids:
        rdd += 1; rc = ROWOF[kid]
        wsD.cell(row=rdd, column=1, value=f"{kid} {names.get(kid,'')[:40]}")
        for j, nm in enumerate(COLS, 2):
            ref = f"Comparison!{ccol(nm)}{rc}"
            noncomp = (nm, kid) in NONCOMP
            wsD.cell(row=rdd, column=j, value=("" if noncomp else f'=IF(ISNUMBER({ref}),{ref},"")')).number_format = "0.00"
    CHROWS[key] = (start, rdd, title); rdd += 2
# active vs contracted MW
wsD.cell(row=rdd, column=1, value="Active vs contracted MW (latest; mixed IT/gross)").font = BOLD
for j, nm in enumerate(COLS, 2): wsD.cell(row=rdd, column=j, value=nm).font = BOLD
mwstart = rdd
# v2 (2026-10-09 coordinator ruling): Oracle live MW = ~2,050 cumulative AI MW delivered since Jun-25 at Aug-26 (FQ1 FY27), tag E, read from
# peers/oracle/oracle_kpis.xlsx Inputs!P48 (I45E) - same source/tag as peer_benchmark (OR-CALL) and price_or_mw Implied_Price r16. Was 1,200 (May-26, Inputs!O48).
_oi = openpyxl.load_workbook(f"{W}/peers/oracle/oracle_kpis.xlsx", data_only=True)["Inputs"]
assert str(_oi["A48"].value) == "I45E" and "FQ1 FY27" in str(_oi["P2"].value or _oi["P1"].value or "FQ1 FY27") and abs(_oi["P48"].value - 2050) < 1e-9, (_oi["A48"].value, _oi["P48"].value)
ORCL_LIVE = float(_oi["P48"].value)
ACT = {"Nebius": "=Nebius_Inputs!C%d" % NR["N17"], "CoreWeave": 1500, "IREN": 40, "Oracle OCI": ORCL_LIVE, "Nscale": 81, "Lambda": 47, "Crusoe": 1000}
CON = {"Nebius": "=Nebius_Inputs!C%d" % NR["N19"], "CoreWeave": 3700, "IREN": 260, "Oracle OCI": 10000, "Nscale": 1370, "Lambda": "", "Crusoe": 6000}
wsD.cell(row=rdd + 1, column=1, value="Active / energized MW"); wsD.cell(row=rdd + 2, column=1, value="Contracted MW")
for j, nm in enumerate(COLS, 2):
    wsD.cell(row=rdd + 1, column=j, value=ACT[nm]); wsD.cell(row=rdd + 2, column=j, value=CON[nm])
wsD.cell(row=rdd + 3, column=1, value=("Basis/tags: Nebius 193 E (Capacity base case) / >3.5 GW R; CRWV 1,500 / 3,700 R; IREN ~40 MW [R, Jun 30, 2026, FY26 10-K] company-wide, excludes Horizon 1 (50 IT MW, accepted Aug 13); Childress ~75 MW active (Oct, D) is one site, not used / 260 D (Microsoft 200 IT + NVIDIA ~60); IREN contracted/live <=6.5x [D] (upper bound); Oracle ~2,050 = cumulative AI MW "
    "delivered since Jun-25 at Aug-26 (E; FQ1 FY27 call + peers/oracle I45E, same as peer_benchmark / price_or_mw; was 1,200 May-26) / contracted '>10 GW secured' [R, FQ1 FY27 call Sep 10] (not a filed contract figure; chart bar drawn at 10,000 as a floor); installed base n/d; Oracle contracted/live n/d; Nscale 81 D / 1,370 R IT; Lambda 47 P / n/d (memo: Anthropic ~350 MW P); Crusoe 1,000 / 6,000 R GROSS.")).alignment = wrap
CHROWS["mw"] = (mwstart, mwstart + 2, "Active vs contracted MW (log scale in PNG)"); rdd += 5
# contracts chart data
wsD.cell(row=rdd, column=1, value="Contracts: $M per MW-yr (IT MW where available, else stated MW)").font = BOLD; wsD.cell(row=rdd, column=2, value="$M/MW-yr").font = BOLD
cstart = rdd
for iid in ["N1", "N2", "P1", "NEW1", "P2", "P2a", "O1", "NEW2", "O2", "NEW16", "REF", "F1", "F2", "F3", "F4", "F5", "F6"]:
    rdd += 1; kr = KROW[iid]; lab = f"{wsK.cell(row=kr, column=3).value} - {wsK.cell(row=kr, column=4).value}"[:40]
    wsD.cell(row=rdd, column=1, value=f"{iid} {lab}")
    wsD.cell(row=rdd, column=2, value=f"=IF(ISNUMBER('Contracts_$MWyr'!W{kr}),'Contracts_$MWyr'!W{kr},IF(ISNUMBER('Contracts_$MWyr'!V{kr}),'Contracts_$MWyr'!V{kr},\"\"))").number_format = "0.00"
CHROWS["contracts"] = (cstart, rdd, "Implied $M per MW-yr by contract"); rdd += 2
# sites chart data
wsD.cell(row=rdd, column=1, value="Sites: evidenced-active (base) vs announced/contracted MW").font = BOLD
for j, h in enumerate(["Active low", "Active base", "Active high", "Announced / contracted", "YE26 base"], 2): wsD.cell(row=rdd, column=j, value=h).font = BOLD
sstart = rdd
for lab, srow in SITECHART:
    rdd += 1; wsD.cell(row=rdd, column=1, value=lab[:40])
    for j, col in zip(range(2, 7), ["G", "H", "I", "E", "K"]):
        wsD.cell(row=rdd, column=j, value=f'=IF(ISNUMBER(Sites!{col}{srow}),Sites!{col}{srow},"")')
CHROWS["sites"] = (sstart, rdd, "Sites: evidenced-active vs announced MW")
rdd += 2
# Nebius billed IT MW path - locks (model v0.9e) (MW + implied $ per billed IT MW)
wsD.cell(row=rdd, column=1, value="Nebius billed IT MW path - locks (model v0.9e)").font = BOLD
for j, h in enumerate(["Billing IT MW", "Revenue / run-rate $M", "$M per billed IT MW-yr", "Basis"], 2): wsD.cell(row=rdd, column=j, value=h).font = BOLD
bstart = rdd
NI_ = lambda i: "Nebius_Inputs!" + nl(i)
for lab, mw, rev, per, basis in [
    ("Q2'26 average (K02b)", f"=AVERAGE({NI_('N_BILL_Q1')},{NI_('N_BILL_Q2')})", f"={NI_('N01')}*4", f"=Comparison!D{ROWOF['K02b']}", "Q2'26 AI cloud rev x4 (R) / Capacity billing E"),
    ("YE26 BASE (~258) @ $7B guide", f"={NI_('N_BILL_YE26_BASE')}", f"={NI_('N_GUIDE_LO')}", f"={NI_('D_YE26_LO_BASE')}", "guide low / base billing"),
    ("YE26 HIGH (~284) @ $9B guide", f"={NI_('N_BILL_YE26_HIGH')}", f"={NI_('N_GUIDE_HI')}", f"={NI_('D_YE26_HI_HIGH')}", "guide high / high billing (price_or_mw fix)"),
    ("FY27 avg BASE (~531.8) @ Street (blended)", f"={NI_('N_BILL_FY27_BASE')}", f"={NI_('N_STREET_FY27')}", f"={NI_('D_STREET_FY27_BASE')}", "Street $12.3B / v0.9e base (incl. MSFT $14.5M anchor)"),
    ("FY27 BASE ex-Vineland @ Street ex-MSFT", f"={NI_('N_BILL_FY27_BASE')}-{NI_('N_MSFT_FY27_MW')}", f"={NI_('N_STREET_FY27')}-{NI_('N_MSFT_FY27_REV')}", f"={NI_('D_STREET_FY27_NONMSFT')}", "PEER COMPARATOR: (Street - Microsoft) / (base - Vineland MW)"),
    ("FY27 avg HIGH (~659.2) @ Street (blended)", f"={NI_('N_BILL_FY27_HIGH')}", f"={NI_('N_STREET_FY27')}", f"={NI_('D_STREET_FY27_HIGH')}", "Street $12.3B / v0.9e high")]:
    rdd += 1
    for j, x in enumerate([lab, mw, rev, per, basis], 1):
        c = wsD.cell(row=rdd, column=j, value=x)
        if j in (2, 3, 4): c.number_format = "0.0"
CHROWS["billing"] = (bstart, rdd, "Nebius billed IT MW - locks (model v0.9e)")
wsD.column_dimensions["A"].width = 52

def bar(title, r0, r1, ncols, anchor, by_row=True, ytitle=None, horiz=False):
    ch = BarChart(); ch.type = "bar" if horiz else "col"; ch.title = title; ch.height = 8; ch.width = 16
    data = Reference(wsD, min_col=1, max_col=1 + ncols, min_row=r0 + 1, max_row=r1) if by_row else Reference(wsD, min_col=2, max_col=1 + ncols, min_row=r0, max_row=r1)
    cats = Reference(wsD, min_col=2, max_col=1 + ncols, min_row=r0) if by_row else Reference(wsD, min_col=1, min_row=r0 + 1, max_row=r1)
    ch.add_data(data, from_rows=by_row, titles_from_data=True); ch.set_categories(cats)
    if ytitle: ch.y_axis.title = ytitle
    wsG.add_chart(ch, anchor)
wsG["A1"] = "Native Excel charts (linked to Chart_Data). PNG versions with notes: /workspace/nebius/peers/charts/."; wsG["A1"].font = BOLD
anchors = ["A3", "K3", "A20", "K20", "A37", "K37", "A54"]
for (key, title, kids), an in zip(CH, anchors):
    s, e, t = CHROWS[key]; bar(t, s, e, len(COLS), an)
s, e, t = CHROWS["mw"]; bar(t, s, e, len(COLS), "K54")
s, e, t = CHROWS["contracts"]; bar(t, s, e, 1, "A71", by_row=False, horiz=True)
s, e, t = CHROWS["sites"]; bar(t, s, e, 5, "K71", by_row=False, horiz=True)
s, e, t = CHROWS["billing"]; bar(t, s, e, 1, "A88", by_row=False, ytitle="IT MW")

# ---------------------------------------------------------------- Peer_memos (unchanged logic)
wm = wb.create_sheet("Peer_memos")
wm.cell(row=1, column=1, value="Peer-specific memo KPIs (M-rows; not template-standard) - full quarterly history").font = BOLD
rr = 3
for lab in ("IREN", "Oracle OCI"):
    wsp, rows, hdr0, _, _ = DATA[lab]; ncol = 12 if lab.startswith("Oracle") else 11
    wm.cell(row=rr, column=1, value=lab).font = BOLD; rr += 1
    for j, h in enumerate(["ID", "Memo KPI", "Unit"] + hdr0[4:4 + ncol], 1):
        c = wm.cell(row=rr, column=j, value=h); c.font = BOLD; c.fill = HDR
    rr += 1
    for kid, r0 in rows.items():
        if not kid.startswith("M"): continue
        for j in range(1, 4): wm.cell(row=rr, column=j, value=wsp.cell(row=r0, column=j).value)
        for j in range(5, 5 + ncol):
            v = wsp.cell(row=r0, column=j).value; c = wm.cell(row=rr, column=j - 1, value=v)
            if isinstance(v, float): c.number_format = "0.0"
        rr += 1
    rr += 1
wm.column_dimensions["B"].width = 60

# ---------------------------------------------------------------- README
RD = [
 ("Peer comparison workbook (step 5, final) - built 2026-10-04 ET; v0.9d REFRESH 2026-10-05 ET; share-basis / K18 refresh 2026-10-08 ET; MODEL v0.9e LOCK REFRESH 2026-10-08 ~10:00 ET (prior pack kept as peers/comparison_pre_v09e_backup.xlsx) by /workspace/nebius/peers/build_comparison.py (prior pack 08:31 ET kept as peers/comparison_0831_prev.xlsx; 10-05 pack as peers/comparison_pre_oct8_backup.xlsx)", BOLD),
 ("Tabs: Comparison (7 companies x 40 template KPIs, value/tag/as-of, peer median, Nebius better/worse + reason) | Ranking | Nebius_Inputs (sources; Comparison Nebius cells are live formulas) | Contracts_$MWyr | Sites | Chart_Data + Charts (native Excel) | Peer_memos.", None),
 ("Tags: R Reported, P Press-reported (new, private peers), D Derived, E Estimated, n/d not disclosed, n.m. not meaningful. Derived tag = weakest input. Private-peer composite tags kept verbatim.", None),
 ("Nebius rev/MW FIX: old Comparison K01 13.7 = Q2'26 group revenue x4 / YE25 170 MW (stale). Now K01 = AI cloud revenue x4 / Jun-26 energized 193 MW = 11.92 E and K02 = x4 / AVERAGE(181.5, 193) = 12.28 E (switched Oct 8 ~23:30 ET from v1 173 / 171.5) (Capacity & Power base case incl. plug; identified sites 152-194 at Jun-26). Same definitions as peers' K01/K02.", None),
 ("Inputs read (not edited): peers/coreweave|iren|oracle/*_kpis.xlsx; peers/private/private_peers_values.json; parts/contract_economics.xlsx (GPU Market agent's contract table - market/ and nebius_market.xlsx hold no contract data); filings/nebius_filings.xlsx + nebius_kpi_history.xlsx; shared/findings_log.xlsx; capacity/nebius_capacity.xlsx + site_list.xlsx + site_verification.md; peers/sites/sites.xlsx + site_verification.md.", None),
 ("Charts: PNGs in /workspace/nebius/peers/charts/ (peers/build_charts.py reads this workbook's recalculated values).", None),
 ("Recalc: python3 /workspace/nebius/model/recalc.py peers/comparison.xlsx (LibreOffice) -> copy /tmp/rc/comparison.xlsx back; 0 errors required.", None),
 ("2026-10-08 ~23:15 ET: Abilene active base 421 -> 618 MW [D on R] (Oracle FQ1 FY27 call Sep 10: 6 of 8 bldgs, 618 MW, 75% delivered; IT vs gross basis unstated); low/base/high 200/421/421 -> 421/618/618, YE26 421/632/843 -> 618/721/824, announced numeric ~843 -> ~824 IT implied. Sites rows cross-checked vs the four site appendices (cols Q-R). No KPI, median or ranking moves (Abilene is a Sites/context row only).", None),
 ("", None),
 ("K02b (added 2026-10-05; now generated by this script, previously hand-patched into the 08:31 pack): Revenue per average BILLING IT MW = Q2'26 AI cloud rev x4 / AVERAGE(Capacity billing IT MW Q1'26, Q2'26) (Nebius_Inputs N_BILL_Q1/Q2, capacity row 147). BILLED-IT basis, not energized/gross. CoreWeave K02b = K02 (crwv_kpis.xlsx K02b row; active ~ billed). Other peers n/d. Microsoft ACV: $14.49M per IT MW [E, house] = model Contracts!E12 (Nebius_Inputs NOTE_1417; earlier 11.6 / 17.1 readings superseded Oct 9).", None),
 (f"SHARE-COUNT CONVENTION (2026-10-08 ~08:45 ET; {FDMD}, parsed): N_FD_DEBT {FD_DEBT:.2f}M ({FDC['FD_DEBT']:,}, line {FDC['FD_DEBT_LINE']}) = new Feb-2030/Feb-2034 converts as debt below their conversion prices (${FDC['PX_2030']:.2f} / ${FDC['PX_2034']:.2f}, lines {FDC['LINE_2030']}-{FDC['LINE_2034']}); "
  f"N_FD_IFCONV {FD_IFCONV:.2f}M ({FDC['FD_IFCONV']:,}, line {FDC['FD_IFCONV_LINE']}) = all if-converted above ${FDC['PX_2034']:.2f}; between the two prices only the 2030s convert ({FDC['FD_MID']:,}). N_FD_V09D {FD_V09D:.2f}M = v0.9d historical memo. "
  "'376.6M primary' dropped. No Nebius per-share or EV-per-MW figure exists in this workbook or peers/charts, so no KPI, ranking or chart value depends on it; v0.9e: build asserts model Inputs F473/F474 = these two values.", BOLD),
 (f"FY27 BILLED-MW LOCKS - MODEL v0.9e (2026-10-08 ~10:00 ET; source of truth, read live by label via peers/nebius_locks.py model_locks()): BASE {LOCK['FY27_BASE']:.2f} = {MODELF} {MC('FY27_BASE')} / HIGH {LOCK['FY27_HIGH']:.2f} = {MC('FY27_HIGH')} "
  f"(non-Vineland {MV('NONVIN_BASE'):.2f} / {MV('NONVIN_HIGH'):.2f} + Vineland {MSFT_MW:.2f} billed MW-yr = engine Microsoft FY27 revenue ${MSFT_REV:,.1f}M / $14.5M). Forward memo: Street $12.3B / {LOCK['FY27_BASE']:.2f} = {BLENDED:.2f} blended; "
  f"EX-MICROSOFT (12,300 - {MSFT_REV:,.1f}) / ({LOCK['FY27_BASE']:.2f} - {MSFT_MW:.2f}) = {HURDLE:.2f} = the figure to use against peers (peer $/MW has no $14.5M Microsoft anchor); high path {LOCK['STREET_FY27']/LOCK['FY27_HIGH']:.2f} blended. "
  f"$16M coverage: {MV('MW_NEEDED_16'):.2f} MW needed = {MV('COV16_BASE'):.2f}x base / {MV('COV16_HIGH'):.2f}x high. Chart: peers/charts/nebius_billing_lock.png.", BOLD),
 ("HISTORY (superseded, not inputs): Oct 8 ~08:45 ET the FY27 locks were Capacity's tranche pair 536.48 / 663.8 with Microsoft $2,711M on 186.97 MW (ex-MSFT 27.44, blended 22.93); before that model v0.9d 474.96 / 645.4 (energized-lagged). Both are retired by v0.9e (the Judge_Slide 475/645 flag is closed). Prior pack: comparison_pre_v09e_backup.xlsx.", None),
 ("NEBIUS Q3'26 RESULTS DATE: ~Nov 10 (est.) - the company has not announced it; use this label wherever the Q3 print date is cited.", None),
 (f"MODEL v0.9e LOCK (Nebius_Inputs yellow rows; read live from {MODELF} Drivers / Judge_Slide / Inputs; build asserts the values): FD {FD_DEBT:.2f}M debt / {FD_IFCONV:.2f}M if-converted ({FD_V09D:.2f}M v0.9d historical memo) | billing IT MW YE26 BASE {LOCK['YE26_BASE']:.1f} / HIGH {LOCK['YE26_HIGH']:.1f} | FY27 avg {LOCK['FY27_BASE']:.2f} / {LOCK['FY27_HIGH']:.2f}. "
  f"None of these feed the Q2'26 trailing KPIs (K01/K02/K02b/K03...) - they are forward/lock references; the forward $/billed MW memo (27.2 / 31.7 / {BLENDED:.1f} blended / {HURDLE:.1f} ex-MSFT / {LOCK['STREET_FY27']/LOCK['FY27_HIGH']:.1f} high) is context only, not scored.", BOLD),
 (f"price_or_mw consistency: Implied_Price YE26 high row {PM['hi_mw']:.1f} MW -> ${PM['hi_per_mw']:.2f}M and low row {PM['lo_mw']:.1f} MW -> ${PM['lo_per_mw']:.2f}M match Nebius_Inputs D_YE26_* memo rows. Street row there now uses the v0.9e base {PM['street_mw']:.2f} MW (${PM['street_per_mw']:.2f}M blended; build asserts it equals N_BILL_FY27_BASE) and its peer 'vs' columns benchmark the ex-MSFT figure (price_or_mw Nebius_Check B37; patched by peers/patch_price_or_mw.py).", None),
 (f"K18 (2026-10-08): read straight from each pack; build asserts K18 = 100/K17 (values {K18CHK}). CoreWeave pack fixed at source in peers/coreweave/build_crwv_kpis.py (was 100/K16 = 16.67 after the in-place K02b insert); the 10-05 re-derivation guard here is removed.", None),
 ("COMP BAND (2026-10-08 ET): compute-contract $/MW-yr band = $9.7-16.7M. Low IREN-Microsoft $9.7M (R). Top Lambda-Anthropic ~$16.7M ($35B / 6-yr term / ~350 MW; Contracts_$MWyr NEW2; Medium confidence: press-only TCV (WSJ via theenergymag; private_peers PX10 16.7 D on P inputs), 6-yr term secondary, IT vs gross MW unstated). Replaces $9.7-16.2M (Nscale-Anthropic $16.16M top). Nebius rows and facility leases excluded. Memo/summary/figures and price_or_mw cite the same band.", BOLD),
 ("2026-10-08 ~23:30 ET (coordinator decision): Nebius energized MW switched from the outdated capacity/nebius_capacity_v1.xlsx 45 / 171.5 / 173 to the current lock 50 / 181.5 / 193 (Quarterly MW J130/M130/N130 = filings KPI row 30, Checks r22); identified sites 157-229; modeled GPUs 80,047.5 -> 86,738 and fleet list $5.571 -> $5.606 follow the same lock. Peer values re-read from the existing peer packs only (no new peer KPIs).", None),
]
for i, (t, f) in enumerate(RD, 1):
    c = rd.cell(row=i, column=1, value=t); c.alignment = wrap
    if f: c.font = f
rd.column_dimensions["A"].width = 160
wb.move_sheet("Comparison", offset=-1)
# v3 (Oct 9 coordinator): IREN / Oracle contracted-vs-live wording on the K04 cells (values stay n/d: a bound / n/d is not a point value, so the K04 median set is unchanged)
from openpyxl.comments import Comment
_k04 = next(r for r in range(4, wsC.max_row + 1) if wsC.cell(r, 1).value == "K04")
assert wsC.cell(3, 10).value == "IREN" and wsC.cell(3, 13).value == "Oracle OCI" and wsC.cell(_k04, 10).value == "n/d" and wsC.cell(_k04, 13).value == "n/d"
wsC.cell(_k04, 10).comment = Comment("IREN contracted/live <=6.5x [D] (upper bound; not in the median). Live = ~40 MW AI Cloud operating capacity at Jun 30, 2026 (FY26 10-K, R), company-wide; excludes Horizon 1 (50 IT MW, accepted Aug 13). Childress ~75 MW active (Oct, D) is one site and not used. Contracted ~260 IT = Microsoft 200 + NVIDIA ~60 (D).", "Peers v3")
wsC.cell(_k04, 13).comment = Comment("Oracle contracted/live n/d. Contracted '>10 GW secured' [R, FQ1 FY27 call Sep 10] (not a filed contract figure). Live ~2,050 = AI MW delivered since Jun-25 (E); installed base n/d.", "Peers v3")
wb.save(OUT); print("saved", OUT)
