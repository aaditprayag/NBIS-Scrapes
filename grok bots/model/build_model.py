import sys, datetime as dt
sys.path.insert(0, '/workspace/nebius/model')
from build_core import *

wb = Workbook()
wb.remove(wb.active)
RN = {}   # registry: row names
def reg(sheet, nm, r):
    name(wb, nm, sheet, f"$F${r}:$AG${r}")
    name(wb, nm + "_E", sheet, f"$E${r}")
    RN[nm] = (sheet, r)
def X(nm, n):
    return f"INDEX({nm},1,{n})" if n >= 1 else f"{nm}_E"
def Xc(nm, ncell):  # index by a cell holding n
    return f"INDEX({nm},1,{ncell})"

# ------------------------------------------------------------------ README
rd = wb.create_sheet("README")
# filled at end

# ------------------------------------------------------------------ INPUTS
ws = wb.create_sheet("Inputs")
ws["A1"] = "Inputs & Assumptions (scenario-driven)"; ws["A1"].font = TITLE
ws["A2"] = ("Blue = input. Fill: orange = PLACEHOLDER (no source yet), yellow = Estimated/Consensus, none = Reported/Derived. "
            "Live column = CHOOSE(Scenario, Bull, Base, Bear). Owners fill/replace their cells; never overwrite formulas.")
ws["A2"].font = ITAL
setw(ws, {"A": 50, "B": 12, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 16, "K": 14, "L": 90})
r = 4
ws[f"A{r}"] = "Scenario selector (1 = Bull, 2 = Base, 3 = Bear)"; ws[f"A{r}"].font = BOLD
inp(ws[f"D{r}"], 2, "Input"); name(wb, "Scen", "Inputs", f"$D${r}")
ws[f"E{r}"] = '=CHOOSE(Scen,"BULL","BASE","BEAR")'; ws[f"E{r}"].font = BOLD
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="3"); ws.add_data_validation(dv); dv.add(f"D{r}")
r = 6
for c, t in zip("ABCDEFGHIJKL", ["Scalar input", "Unit", "Bull", "Base", "Bear", "LIVE", "", "", "", "Owner", "Status", "Source / method"]):
    ws[f"{c}{r}"] = t; ws[f"{c}{r}"].font = HDR; ws[f"{c}{r}"].fill = HDRFILL
r += 1
SC = {}
def sec(title):
    global r
    ws[f"A{r}"] = title; ws[f"A{r}"].font = BOLD
    for c in range(1, 13): ws.cell(r, c).fill = SECFILL
    r += 1
def S(nm, label, unit, base, owner, status, src, bull=None, bear=None, fmt="#,##0.00"):
    """scenario scalar; if bull/bear None -> single value (non-scenario)"""
    global r
    ws[f"A{r}"] = label; ws[f"B{r}"] = unit; ws[f"B{r}"].font = ITAL
    if bull is None and bear is None:
        inp(ws[f"D{r}"], base, status, fmt)
        if isinstance(base, str) and base.startswith("="): ws[f"D{r}"].font = BLACK
        ws[f"C{r}"] = "—"; ws[f"E{r}"] = "—"
        fml(ws[f"F{r}"], f"=D{r}", fmt, True)
    else:
        inp(ws[f"C{r}"], bull if bull is not None else base, status, fmt)
        inp(ws[f"D{r}"], base, status, fmt)
        inp(ws[f"E{r}"], bear if bear is not None else base, status, fmt)
        fml(ws[f"F{r}"], f"=CHOOSE(Scen,C{r},D{r},E{r})", fmt, True)
    ws[f"J{r}"] = owner; ws[f"K{r}"] = status; ws[f"L{r}"] = src
    for c in "JKL": ws[f"{c}{r}"].font = ITAL
    name(wb, nm, "Inputs", f"$F${r}")
    SC[nm] = r
    r += 1

sec("General")
S("ValDate", "Valuation date", "date", dt.date(2026, 10, 4), "Modeling", "Input", "Pitch valuation date; drives DCF discounting and Outputs horizons", fmt="yyyy-mm-dd")
S("LastHist", "Last actual quarter index (10 = Q2'26)", "#", 10, "Modeling", "Input", "Roll forward to 11 when Q3'26 is reported", fmt="0")
S("FYBase", "Fiscal-year base (year before first model year)", "yr", 2023, "Modeling", "Input", "Structural: annual-input column = FY - FYBase", fmt="0")
S("SharePx", "Current share price (NBIS)", "$/sh", 242.81, "Skeptic", "Reported", "Oct 2 2026 close per Skeptic (/workspace/nebius/skeptic/01_priced_in.xlsx), medium confidence; used for convert ITM test, TSM, equity raises")
S("MWSel", "MW block that drives revenue (1 contracted, 2 energized est. [Capacity, default], 3 mgmt-connected guide, 4 Capacity proposed path [scenario, not base until Tuesday file])", "#", 2, "Modeling", "Input", "Default = Capacity energized base (estimated live IT load); guided path kept as comparison only", fmt="0")
S("EnLag", "Energized (live IT load) -> billing lag, MW source 2", "qtrs", 1, "Capacity", "Estimated", "Base 1 qtr. 2 qtrs = STRESS case (Sensitivity snapshot), not the bear default", bull=1, bear=1, fmt="0")
S("CapexMethod", "IT capex method (1 = GPUs x $/GPU, 2 = flat $M per IT MW, 3 = site path: new energized IT MW x per-site $/MW by ownership & generation [default])", "#", 3, "Modeling", "Input", "Method 3 = path-consistent per Capacity gpu_mw_capex.xlsx", fmt="0")
S("CPM_Own_GB200", "Owned site all-in capex per IT MW - GB200", "$M/MW", 42, "Capacity", "Estimated", "/workspace/nebius/capacity/gpu_mw_capex.xlsx CapexPerMW (incl ~$10-12.5M shell+power)", fmt="0.0")
S("RackOn", "GB300 capex from rack price (1 = racks/IT MW x GB_RackPx + network; owned adds CPM_SP; 0 = v0.7 flat $41M leased / $53M owned)", "0/1", 1, "Capacity/GPU Market", "Input", "", fmt="0")
S("GB_RackPx", "GB300 NVL72 rack price (servers incl. NVLink; ex network)", "$M/rack", 4.5, "Capacity", "Reported", "gpu_mw_capex.xlsx v2 'GB300 rack price' tab (sourced): base $4.5M, range $4.0-5.0M (MS BOM $4.0M, Loop $3.85M, Wolfe $4.3M, Data Gravity $4.97M; stress $6.25M)", fmt="0.00")
S("GB_Racks", "GB300 NVL72 racks per IT MW", "racks/IT MW", "=491/72", "Capacity", "Derived", "491 GPUs/IT MW / 72 = 6.82", fmt="0.00")
S("GB_NetMW", "GB300 network + storage per IT MW", "$M/IT MW", 6.1, "Capacity", "Derived", "gpu_mw_capex.xlsx CapexPerMW GB300 base", fmt="0.00")
S("CPM_Own_GB300", "Owned site all-in capex per IT MW - GB300 (RackOn: rack-based servers+network + CPM_SP hall)", "$M/MW", "=IF(RackOn=1,GB_Racks*GB_RackPx+GB_NetMW+CPM_SP,53)", "Capacity", "Derived", "v0.7 53 = 34.1 servers + 6.1 network + 12.5 hall", fmt="0.0")
S("CPM_Own_VR", "Owned site all-in capex per IT MW - Vera Rubin (LOW-MEDIUM confidence)", "$M/MW", 52.6, "Capacity", "Estimated", "Capacity gpu_mw_capex v2: base $52.6M (= leased 41.1 - 1.0 fit-out + 12.5 hall); bull 45.8 / bear 60.7 same build-up; v0.7 56 (53/59)", bull=45.8, bear=60.7, fmt="0.0")
S("CPM_Own_Other", "Owned site capex per IT MW - Hopper/B200/B300 (GB200 proxy)", "$M/MW", 42, "Capacity", "Estimated", "proxy", fmt="0.0")
S("CPM_Lease", "Leased/colo capex per IT MW - GB200 (and Hopper/B200/B300 stand-ins); IT only, rent in opex", "$M/MW", 30, "Capacity", "Estimated", "gpu_mw_capex.xlsx (Capacity), medium confidence", fmt="0.0")
S("CPM_Lease_GB300", "Leased/colo capex per IT MW - GB300 (RackOn: racks x rack price + network; no hall slice - rent covers the building)", "$M/MW", "=IF(RackOn=1,GB_Racks*GB_RackPx+GB_NetMW,41)", "Capacity", "Derived", "6.82 x $4.5M + $6.1M = $36.8M (range $31.4-42.3M, Capacity v2); v0.7 41. NOTE: Capacity's leased Rubin $41.1M includes ~$1M colo fit-out; GB300 here excludes it (per team spec)", fmt="0.0")
S("CPM_Lease_VR", "Leased/colo capex per IT MW - Vera Rubin, EX $1M colo fit-out (rent is the only building cost; same as GB300) (LOW-MEDIUM confidence)", "$M/MW", 40.1, "Capacity/Skeptic", "Estimated", "v0.7: base $41M flat to GB300 (SemiAnalysis TCO 1.51-1.56x/GPU-hr at 320 vs 491 GPUs/IT MW -> ~0.98-1.02x GB300 per MW; Capacity at Morgan Stanley rack cost: leased $42-48M); bull $41M; bear $48.5M (Skeptic $45-52M midpoint: 0.65x racks/MW x MS 1.68-1.95x rack cost). $32M = OUTER-BULL sensitivity only (soft: ~$6.0M rack ~1.2x GB300; networking $3.6M looks low vs GB300 $6.1M - Capacity reconciles Tuesday)", bull=33.3, bear=48.2, fmt="0.0")
S("FixMgmt", "v0.2 fix: GPU fleet MW-ratio only in forecast + capex follows DRIVER MW (1 = on)", "0/1", 1, "Modeling", "Input", "Off = v0.1 behaviour (history fleet scaled by selected MW block; capex followed energized path only)", fmt="0")
S("TermBase", "Terminal basis: 2 = EXCLUDE BOTH Q4'30-onward spend for MW billing after Q4'30 and that MW (terminal = Q4'30 exit run-rate x 4) [default, conservative]; 1 = INCLUDE BOTH (memo only: run-rate scaled x(MW Q4'30 + MW billing Q1'31)/MW Q4'30, unsourced ~1.2x step-up); 0 = v0.1 (FY2030 x (1+TG), Q4'30 spend charged)", "0/1/2", 2, "Modeling", "Input", "v0.2 decision (Oct 4): do not credit unbilled capacity on an unsourced step-up; 'include both' EV shown as memo row in Valuation", fmt="0")
S("TVFloor", "Terminal value floor: max(Gordon/blend, run-off value of existing fleet with no reinvestment) (1 = on)", "0/1", 1, "Modeling", "Input", "Operator would stop replacing GPUs at negative NPV; TV cannot be below run-off (>= 0)", fmt="0")
S("OM_MW", "Owned-site O&M (facility staff, maintenance, insurance, property tax)", "$M/IT MW-yr", 0.75, "Capacity", "Placeholder", "Leased sites carry this inside rent; owned sites had none in v0.1. Placeholder pending Capacity", fmt="0.00")
S("LeasePrepayOn", "Lease prepayments as upfront cost on NEW leased IT MW (1 = on)", "0/1", 1, "Filings", "Input", "", fmt="0")
S("LeasePrepay", "Lease prepayment per new leased IT MW (prepaid rent at commencement; capitalised, amortised over Life_Pow)", "$M/IT MW", 1.6, "Filings", "Derived", "6/30/26 prepayments on not-yet-commenced leases $612.4M / ~383 leased IT MW starting by YE27 (Capacity v2)", fmt="0.00")
S("PrepayMode", "New customer prepayment inflow: 1 = % of new-build capex (v0.7); 2 = Q2'26 pace $1.2B/qtr [base]; 3 = guidance case: $4.5B in H2'26 (>$9B FY26 guide), then Q2 pace", "#", 2, "Filings", "Derived", "Q2'26 revenue note: deferred revenue +~$3.2B Q1, +~$1.2B Q2", fmt="0")
S("PrepayQ", "Q2'26-pace new prepayments per quarter", "$M/qtr", 1200, "Filings", "Derived", "Deferred revenue added ~$1.2B in Q2'26", fmt="#,##0")
S("PrepayH2", "Guidance case: new prepayments H2'26 total", "$M", 4500, "Filings", "Derived", "FY26 guide >$9B less H1 ~$4.4B", fmt="#,##0")
S("PrepayEndIdx", "Last quarter of the fixed-pace prepayment path (capex-linked % after)", "idx", 20, "Modeling", "Estimated", "Q4'28; 2029+ reverts to capex-linked % (PrepayMode 1 logic)", fmt="0")
S("DRRefPct", "PLACEHOLDER: share of deferred-revenue balance treated as debt-like (refundable) at valuation", "%", 0.0, "Filings", "Placeholder", "NOT DISCLOSED: filing says balance includes 'amounts that may be refundable'; bear placeholder 25%", bull=0.0, bear=0.25, fmt="0%")
S("VinIdleOn", "Bloom slip / Vineland loss: slipped or lost Microsoft MW sit IDLE (not re-sold to new-deal books) (1 = v0.7c fix)", "0/1", 1, "Capacity", "Input", "v0.7b: slipped T5-T9 GPUs were re-priced as new 2027+ deals, offsetting most of the slip", fmt="0")
S("VinT14Loss", "BEAR SENS: Vineland T1-T4 permanently lost if generators must stop before fuel cells run (1 = on)", "0/1", 0, "Capacity", "Estimated", "4 x ~26.7 = ~107 IT MW, ~$1.55B/yr at $14.49M/IT MW; trigger ~Nov 6-9 2026 (NJDEP 45-day deadline)", fmt="0")
S("VinSLAOn", "Vineland T1-T4 SLA-credit step on (1)", "0/1", 1, "Capacity", "Input", "Step (i); step (ii) = permanent loss (VinLossOn)", fmt="0")
S("VinLossOn", "Vineland T1-T4 permanent-loss step on (1)", "0/1", 1, "Capacity", "Input", "", fmt="0")
S("VinSLA", "PLACEHOLDER: SLA credits as % of T1-T4 revenue (Q4'26 + Q1'27)", "%", 0.10, "Capacity", "Placeholder", "NOT DISCLOSED - placeholder 10% for 2 quarters", fmt="0%")
S("VinLossIdx", "First quarter T1-T4 revenue permanently lost", "idx", 14, "Capacity", "Estimated", "Q2'27", fmt="0")
S("VinRedGap", "Quarters lost T1-T4 GPUs sit idle before redeployment to new-deal books", "qtrs", 2, "Capacity", "Estimated", "Redeployed at the live scenario's GB300 new-deal price (bear $14.5M) after 2 qtrs", fmt="0")
S("ThreeYrOn", "3-yr GB300 new deals re-price at renewal to the then-prevailing curve (1 = on)", "0/1", 0, "Contracts", "Input", "Option (v0.7c); bear sensitivity", fmt="0")
S("ThreeYrSh", "Share of new GB300 MW on 3-yr deals (re-price at year 3)", "%", 0.0, "Contracts", "Estimated", "Bear 40%", bull=0.0, bear=0.40, fmt="0%")
S("FP_Resize", "Funding plan re-sized to the live gap: ATM (remaining 12.27M sh) -> 2027 converts (same terms) -> equity top-up each quarter to FP_MinBuf (1 = v0.7c)", "0/1", 1, "Modeling", "Estimated", "", fmt="0")
S("FP_MinBuf", "Minimum liquidity buffer every quarter (re-sized plan)", "$M", 1000, "Modeling", "Input", "", fmt="#,##0")
S("FP_TopEnd", "Last quarter of equity top-up (and post-funding count date when FP_Resize)", "idx", 28, "Modeling", "Input", "Q4'30 = every quarter of the horizon", fmt="0")
S("FP_EqDisc", "Equity top-up issue price discount to SharePx (scenario issuance price path, flat)", "%", 0.03, "Modeling", "Estimated", "Base 3% (EqDisc), bull 0%, bear 40% (as bear ATM)", bull=0.0, bear=0.40, fmt="0%")
S("StopBuild", "v0.8 SCENARIO 'contracted only / stop building': after StopIdx only signed capacity is built (contract tranches incl. Meta backstop if on); non-contract fleet frozen at its StopIdx level; no new-deal prepayments or lease prepayments", "0/1", 0, "Modeling", "Input", "Uncontracted 2027+ slice per Capacity uncontracted_split.csv (Birmingham, Lappeenranta, Butler, Kapuli owned Rubin; Beit Shemesh, UK3, CLT1 leased) and all later growth is not built", fmt="0")
S("StopIdx", "Stop-building: last quarter whose ACTIVE (billing) MW is kept - Q1'27 = sites energized by YE26 (EnLag 1 qtr); ATM plan + 2027 plan converts switched off, equity top-ups only to FP_MinBuf at the self-consistent PT", "idx", 13, "Capacity", "Estimated", "Q1'27 billing = sites live by YE26 (Oklahoma, Newport, Harlow, Masmiyya, Tallinn) are built / capex largely spent", fmt="0")
S("StopCommitWt", "Stop-building: share of signed leases-not-commenced + energy commitments ($17.4B), pro rata to the cut MW, still paid with no modeled rent / power (net of sublease / resale)", "%", 0.5, "Modeling", "Estimated", "Only when StopBuild = 1; 0% / 100% as range", fmt="0%")
S("NDPrepayOn", "v0.8b FUNDING STACK: new-deal prepayment toggle (1 = prepayments = NDPrepayPct x new-build capex on NON-signed MW; replaces PrepayMode pace / PrepayPct27/28 / HSPre; signed tranches 0%)", "0/1", 0, "Modeling", "Input", "0 = v0.8 logic", fmt="0")
S("NDPrepayPct", "v0.8b: new-deal prepayment share of new-build capex (non-signed MW) when NDPrepayOn = 1: 0% / 25% / 55%", "%", 0.55, "Filings", "Estimated", "Q2'26 letter: prepayments fund 50-60% of capex on deals that carry them", fmt="0%")
S("ABS_LTC", "v0.8b: asset-backed debt on CONTRACTED MW only (signed Microsoft / Meta incl. Meta backstop): loan-to-cost on their new-build capex 0 / 25 / 50%", "%", 0.0, "Modeling", "Input", "Drawn with capex; none on uncontracted MW; existing 6/30 fleet not refinanced", fmt="0%")
S("ABS_Spread", "v0.8b: ABS spread over SOFR", "%", 0.025, "Filings", "Estimated", "SOFR + 2.50%", fmt="0.00%")
S("ABS_Term", "v0.8b: ABS straight-line amortisation over contract term from billing start", "qtrs", 20, "Contracts", "Estimated", "5-yr Microsoft / Meta / backstop terms", fmt="0")
S("UncPxOn", "v0.8b: UNSOLD new 2027+ non-contract MW priced at UncPx (both generations, flat then MYDecl path) on 3-yr terms; contracted / RPO-sold MW unchanged", "0/1", 0, "Skeptic", "Input", "Legacy GPU-hour fleet (147 MW at Q1'27) stays on books (b)/(c)", fmt="0")
S("UncPx", "v0.8b: unsold new-MW price when UncPxOn = 1 ($M per IT MW-yr): 16 / 20 / 22.5", "$M/IT MW-yr", 20, "Skeptic", "Estimated", "", fmt="0.0")
S("UncRenew", "v0.8b: 3-yr terms - unsold vintages re-price at year 3 (2027 vintages -> 2030) to the prevailing new-deal curve (ND_Px GB300 / RubinMW Rubin path)", "0/1", 1, "Skeptic", "Input", "", fmt="0")
S("UncSplitOn", "v0.8d: SPLIT uncontracted pricing (replaces single UncPx in the grid): unsold GB300/B300 vintages = LEASED slice at UncGBPx on UncGBTermY terms; unsold Rubin vintages = OWNED Rubin halls at UncRubPx (5-yr)", "0/1", 0, "Skeptic", "Input", "UncPxOn takes precedence if both on", fmt="0")
S("UncGBPx", "v0.8d: LEASED GB300/B300 uncontracted price ($M/IT MW-yr): bear 16.5 (5-yr) / base 20 (3-yr) / bull 22.5 (3-yr)", "$M/IT MW-yr", 20, "Skeptic", "Estimated", "", bull=22.5, bear=16.5, fmt="0.0")
S("UncGBTermY", "v0.8d: LEASED slice contract term (yrs): bear 5 / base 3 / bull 3", "yrs", 3, "Skeptic", "Estimated", "", bull=3, bear=5, fmt="0")
S("UncGBRen", "v0.8d: LEASED slice renewal rate = 5-yr rate (base 16.5, bear 14.5; bull keeps the 3-yr 22.5); applied at age 4 x term, incl. terminal value", "$M/IT MW-yr", 16.5, "Skeptic", "Estimated", "", bull=22.5, bear=14.5, fmt="0.0")
S("UncRubPx", "v0.8d: OWNED Rubin halls uncontracted price ($M/IT MW-yr, 5-yr; renews flat): bear 14.5 / base 16.2 / bull 20", "$M/IT MW-yr", 16.2, "Skeptic", "Estimated", "", bull=20, bear=14.5, fmt="0.0")
S("UncCohort", "v0.8d bridge: vintages the UncPx / split toggle applies to (0 all; 1 = vintages billing by Q4'27 [YE27 cohort]; 2 = 2028-30 vintages only)", "#", 0, "Modeling", "Input", "bridge diagnostic", fmt="0")
S("UncPxX", "v0.8d bridge: price for vintages OUTSIDE UncCohort (UncPx single toggle: both gens; split: leased GB300 only)", "$M/IT MW-yr", 16, "Modeling", "Input", "bridge diagnostic; irrelevant when UncCohort = 0", fmt="0.0")
S("UncRenAge", "v0.8d: unsold GB300 renewal age (qtrs) = 12 under UncPxOn (v0.8b), else 4 x UncGBTermY", "qtrs", "=IF(UncPxOn=1,12,4*UncGBTermY)", "Modeling", "Derived", "", fmt="0")
S("TVRenOn", "v0.8d: terminal-year cash flow reflects renewed pricing of unsold vintages not yet renewed by Q4'30 (1 = on)", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("CpnChg", "v0.8d: charge convert cash coupons paid after V_Q through FP_PostQ (after tax when taxable) against equity value (they are funded by top-ups inside V_FPCash; Q4'30 claim = carrying, which excludes window coupons -> no double count)", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("SC_NoPrem", "v0.8d: R5 OFF - new shares issued at today's value v (SC_Disc = Ke/(1+Ke)), not at the 12-month PT v x (1+Ke) [headline]; 0 = v0.8c premium memo", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("BackstopMW", "Meta backstop: covered IT MW (0 / 150 / 190), applied first to the owned Rubin halls (Birmingham 56, Lappeenranta 46, Butler 48, Kapuli 15 = 165 MW)", "IT MW", 0, "Filings", "Derived", "FY25 20-F p.75 (Reported terms): Meta buys unsold capacity on certain clusters for 5 yrs from deployment, up to $15B (~$3B/yr = ~190 MW at $16M or ~150 MW at $20M). MW mapping Derived / low confidence. Off-RPO (not in signed backlog / floor)", fmt="0")
S("BackstopPx", "Meta backstop price ($M per IT MW-yr) - PLACEHOLDER", "$M/IT MW-yr", 16, "Filings", "Placeholder", "Bear $14.5M (below $16M)", bear=14.5, fmt="0.0")
S("BackstopIdx", "Meta backstop billing start (MW-weighted owned Rubin hall go-live: Q1'27 56 / Q2'27 46 / Q4'27 63)", "idx", 15, "Capacity", "Estimated", "Q3'27", fmt="0")
S("PTBasis", "HEADLINE PT issuance-price basis: 1 = SELF-CONSISTENT (ATM, top-ups and convert terms priced at the solved PT itself; fixed point) [v0.7d]; 2 = MEMO spot SharePx x (1 - disc) (v0.7c; depends on market access at today's price); 3 = MEMO path from spot to the path's own PT over FP_PathQ (depends on market access at today's price)", "#", 1, "Modeling", "Input", "v0.7d: headline no longer rests on issuing at ~$236 against ~$0 of value", fmt="0")
S("SC_Disc", "Self-consistent basis: issue discount to the solved 12-month PT (v0.8d: = SC_NoPrem x Ke/(1+Ke) -> issue at today's value)", "%", "=SC_NoPrem*Ke/(1+Ke)", "Modeling", "Derived", "v0.8c and earlier: 0 = issue at the 12-month PT itself", fmt="0.0%")
S("FP_PathQ", "Path memo: quarters over which the issue price moves linearly from SharePx to the path PT", "qtrs", 8, "Modeling", "Estimated", "8 = Q3'28; flat after", fmt="0")
S("AuthAvail", "Share-issuance authority available now (v0.8f): ~94M = 64.4M pool + 50.2M treasury - 21.1M NVIDIA pre-funded warrant; ~114.6M without", "M sh", 93.5, "Filings", "Estimated", "v0.8c per parent: 64.4M pool + ~68.1M treasury; above it a solved PT is a constraint, not an answer; more needs an EGM", fmt="#,##0.0")
S("TVROResOn", "Run-off terminal value includes the 15% after-tax GPU residual as each exit-fleet cohort retires (1 = v0.7c fix)", "0/1", 1, "Modeling", "Input", "v0.7b run-off TV (the binding TV basis) dropped hardware to zero at end of life; going-concern TV already netted V_GResid", fmt="0")
S("BridgeFix", "Asset floor on the IDENTICAL EV-to-equity bridge as the DCF (1 = v0.7c)", "0/1", 1, "Modeling", "Input", "Same OffBSWt x commitments, converts (carrying), cash, refundable-DR placeholder, SOTP, funding cash & shares", fmt="0")
S("OwnOpexOn", "v0.7b fix: owned-site opex = OwnOpex all-in (power + O&M, NO rent); O&M = OwnOpex - power per IT MW (1 = on; 0 = OM_MW)", "0/1", 1, "Capacity", "Input", "Leased sites pay RentMW x PUE (shell+power) + power; owned sites carry the building in capex", fmt="0")
S("OwnOpex", "Owned-site opex per IT MW-yr, all-in power + ops, no rent", "$M/IT MW-yr", 2.25, "Capacity", "Estimated", "Capacity: $2.0-2.5M per IT MW-yr (use 2.25)", fmt="0.00")
S("GPU_Resid", "Residual / resale credit for retired GPUs+servers at refresh (% of original cohort cost)", "%", 0.15, "GPU Market", "Estimated", "gpu_price_history.xlsx: base 15% (range 10-25%, LOW confidence); bear 0. H100 median $4.09 (Mar-24) -> $3.99; A100 neocloud +~12%/yr", bull=0.15, bear=0.0, fmt="0%")
S("HallLife", "Data hall + power ECONOMIC life (residual value in payback/IRR/NPV only; book lives Life_Shell / Life_Pow)", "yrs", 18, "Capacity", "Estimated", "Economic-life estimate only (15-20+ yrs); not a book life", fmt="0")
S("CapIntFix", "v0.7b fix: UFCF tax and terminal D&A exclude capitalized-interest depreciation (financing must not change unlevered EV) (1 = on)", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("CapV", "Capacity file version: 1 = v1 (YE27 ~950 IT MW energized), 2 = v2 (Oct 4 22:33 ET; YE26 ~302 / YE27 ~1,082 IT MW, Basis-tab IT factors)", "#", 2, "Capacity", "Input", "v2 adds Newport UK, Oklahoma, Minnesota, UK 3rd, CLT1 from file; new MW carry own path capex", fmt="0")
S("AuthBase", "Issued share capital incl. treasury at AGM (220.4 Class A + 68.1 treasury + 33.5 Class B)", "M sh", 322.0, "Filings", "Reported", "AGM notice items 13-15 (Aug-26): 5-yr authority to issue up to 20% of issued capital", fmt="#,##0.0")
S("AuthPct", "Board issuance authority (% of issued capital, rises as shares are issued)", "%", 0.20, "Filings", "Reported", "Base authority ~64M + treasury (medium confidence); board can seek more at an EGM", fmt="0%")
S("RepriceOn", "H2'26 repricing: Q2 list avg (H100 $2.95->$3.85 Jun-26), legacy reserved book renews +30% at 50%/12 months, new GPUs at current list (1 = on)", "0/1", 1, "GPU Market", "Input", "Off = v0.1 (Q2 at Jun list; reserved = lagged list)", fmt="0")
S("RenewUp", "Reserved renewal repricing vs expiring contract", "%", 0.30, "Filings", "Reported", "Q2'26 letter: >30% repricing on renewals", fmt="0%")
S("ODq_H100", "H100 on-demand list - Q2'26 average", "$/GPU-hr", "=(2*2.95+3.85)/3", "GPU Market", "Derived", "Archived nebius.com/prices: $2.95 Apr-May -> $3.85 from Jun 2026 (Reported pages); Q2 avg derived", fmt="0.00")
S("ARR_RRMW_Lo", "Team June run-rate per MW - low", "$M/MW-yr", 13.3, "Modeling", "Estimated", "June list + renewals signed by 6/30", fmt="0.0")
S("ARR_RRMW_Hi", "Team June run-rate per MW - high", "$M/MW-yr", 13.5, "Modeling", "Estimated", "", fmt="0.0")
S("ARR_ResTeamLo", "Team estimate: unexplained ARR residual - low", "$M", 650, "Modeling", "Estimated", "~45-50 MW", fmt="#,##0")
S("ARR_ResTeamHi", "Team estimate: unexplained ARR residual - high", "$M", 700, "Modeling", "Estimated", "", fmt="#,##0")
S("MarginCons28", "Consensus FY28 adj. EBITDA margin (13.98 / 23.0)", "%", "=13.98/23", "Skeptic", "Estimated", "01_priced_in.xlsx", fmt="0.0%")
S("MarginPeer", "Peer adj. EBITDA margin reference (CoreWeave ~62%)", "%", 0.62, "GPU Market", "Placeholder", "pending GPU Market comps", fmt="0%")
S("VinCapOn", "Cap Microsoft GPU consumption at Vineland energized IT MW (Capacity) (1 = on)", "0/1", 1, "Modeling", "Input", "Fix for Q2'26 utilization: MSFT delivered-tranche MW (53 MW) exceeded Vineland energized (39 MW / 19,157 GPUs)", fmt="0")
S("ReplOn", "GPU end-of-life replacement capex in forecast (1 = on)", "0/1", 1, "Modeling", "Input", "Cohort reaching GPU life is replaced at original cost (fleet persists in Capacity file)", fmt="0")
S("HistMWBasis", "History MW basis: 1 = Capacity named-site (173 MW Q2'26) [default]; 2 = company-implied (~190 MW, ~$12M rev/installed MW) applied to history AND forward path (+uplift); 3 = MISMATCH test: history at 190, forward on 173 path", "#", 1, "Capacity/Skeptic", "Input", "Revenue/MW in history and the forward MW path must use the same basis (2); 3 only quantifies Skeptic's ~11% claim", fmt="0")
S("CoMWQ2", "Company-implied installed MW at Q2'26 (Q2 letter ~$12M revenue per installed MW)", "MW", 190, "Filings", "Derived", "~$2.3B annualized / ~$12M per MW", fmt="#,##0")
S("CLT1Rent", "AIB CLT1 colocation rent (12-yr), per IT MW-yr", "$M/IT MW-yr", 2.2, "Capacity", "Estimated", "CLT1-sourced ~$2.2M per MW-yr (Capacity); applied via differential vs generic leased rent (RentMW x PUE)", fmt="0.00")
S("CLT1Prepay", "Nebius prepayment to AIB to fund CLT1 build (outflow Q4'26-Q2'27, equal thirds)", "$M", 0, "Filings", "Placeholder", "Amount undisclosed (AIB 8-K Sep 30 2026); sensitivity $250M / $500M; liquidity only", fmt="#,##0")
S("InferizeOn", "Inferize acquisition cash outflow Q4'26 (1 = on)", "0/1", 0, "Filings", "Placeholder", "6-K Oct 1 2026: terms undisclosed; PRESS-REPORTED $100-150M (unconfirmed)", fmt="0")
S("InferizeAmt", "Inferize purchase price (press-reported, unconfirmed)", "$M", 130, "Filings", "Placeholder", "Press $100-150M; liquidity only", fmt="#,##0")
S("JuneExitOn", "Q3'26 starts from June exit run-rate: AI cloud revenue >= reported ARR / 4 before new MW (overlay held flat in $)", "0/1", 0, "Filings", "Input", "Filings: ARR = June recognized AI cloud revenue x 12 (no signed-not-billing). Overlay not tied to MW; power/rent not added (EBITDA margin on overlay = 1 - other COR - PD - SG&A %)", fmt="0")
S("MarMo", "March 2026 AI cloud revenue (monthly)", "$M/mo", 162, "Filings", "Estimated", "~$160M (Filings); June $250M = $3.0B / 12", fmt="#,##0")
S("PeerMarginOn", "Sensitivity: cap adj. EBITDA margin at MarginPeer from PeerMarginYr (cost overlay)", "0/1", 0, "GPU Market", "Input", "", fmt="0")
S("PeerMarginYr", "First year of peer-margin cap", "yr", 2028, "Modeling", "Input", "", fmt="0")
S("VinShutSel", "Vineland shutdown scenario (Capacity vineland_shutdown.xlsx): 0 off; 1 no shutdown; 2 partial 50% 2m; 3 partial minor-source 2m; 4 full 1m; 5 full 2m; 6 full 3m; 7 probability-weighted", "#", 0, "Capacity", "Estimated", "Outage from Nov 7 2026 (NJDEP order Sep 22 + 45 days); ~102.5 MW IT genset-fed; scenarios & weights = Capacity judgment. SLA credits not disclosed (note only); genset power saving not credited", fmt="0")
S("VinRevMW", "Microsoft revenue per IT MW-yr for outage loss", "$M/IT MW-yr", 14.5, "Capacity", "Derived", "$17.4B TCV / 5 yrs / 240 MW IT; probability-weighted Q4'26 loss ~$67-68M", fmt="0.0")
S("BloomSlipOn", "Bloom slips 2 quarters: Vineland T5-T9 (2027) billing and Bloom go-live +2 qtrs (1 = on)", "0/1", 0, "Capacity", "Input", "", fmt="0")
S("NewDealOn", "New-deal pricing for 2027+ new non-contract MW: multi-year $/IT MW-yr tranche + short-term tranche (1 = on)", "0/1", 1, "Filings", "Input", "Q2'26 letter: new deals >$20M/MW, landmark $20-25M ACV/MW (basis likely IT, medium-low); <=6-month deals $40-50M/MW (one signed). Legacy (Q4'26) non-contract fleet stays on GPU-hour pricing", fmt="0")
S("NewDealIdx", "First quarter of new-deal pricing", "idx", 13, "Filings", "Input", "Q1'27", fmt="0")
S("ND_Px", "NEW-DEAL GB300 $/GPU-hr (multi-year, IT reading) - the ONE linked new-deal price input", "$/GPU-hr", 16e6 / (491 * 8760), "GPU Market/Filings", "Estimated", "v0.7b from CONTRACTS: base $16M/IT MW = $3.72/GPU-hr (5-yr deals $14.5-16M: Microsoft ~$14.5M = $3.37, Astra ~$15.9M = $3.70); bull = mix-shift $18M ($4.18; ~60% of MW on 3-yr deals at ~$20M like IREN $4.65); bear $14.5M ($3.37). NOT calibrated to the 22-month payback. v0.7 was $22M ($5.12)", bull=18e6 / (491 * 8760), bear=14.5e6 / (491 * 8760), fmt="0.000")
S("ND_GPUMW_GB", "GB300 NVL72 GPUs per IT MW (new-deal calibration)", "GPU/MW", 491, "Capacity", "Estimated", "gpu_mw_capex.xlsx base; 409 per gross MW", fmt="#,##0")
S("ND_Occ", "New-deal occupancy (multi-year take-or-pay)", "%", 1.0, "Filings", "Estimated", "Take-or-pay = 100%", fmt="0%")
S("MYRate", "Multi-year GB300 ACV per IT MW-yr = ND_Px x ND_GPUMW_GB x 8,760 x ND_Occ (derived; do not overwrite - change ND_Px)", "$M/IT MW-yr", "=ND_Px*ND_GPUMW_GB*8760*ND_Occ/1000000", "Filings", "Derived", "IT reading $22M <-> $5.12/GPU-hr; gross reading (NDBasis 2) x 491/409 = $26.4M <-> $6.14", fmt="0.0")
S("MYPathMode", "Multi-year $/MW path: 1 = v0.6 flat nominal within generation, then MYDecl/yr from MYDeclIdx (no per-chip ageing stacked); 0 = v0.5 Blackwell per-chip list path (double count)", "0/1", 1, "Modeling", "Estimated", "Team consensus v0.6: -6%/yr was per-chip on-demand ageing; Rubin switch handled by Rubin_Prem x GPUs/MW only", fmt="0")
S("MYDecl", "Multi-year new-deal $/MW change per yr (within generation)", "%/yr", -0.03, "Modeling", "Estimated", "Base flat through 2028 then -3%/yr; bear -6%/yr from 2027; bull -1.25%/yr from 2029 (base ratio -3/-6 applied to bull's -2.5% chip ageing)", bull=-0.0125, bear=-0.06, fmt="0.00%")
S("MYDeclIdx", "First quarter of multi-year $/MW decline", "idx", 21, "Modeling", "Estimated", "Base/bull Q1'29; bear Q1'27", bull=21, bear=13, fmt="0")
S("MYStressOn", "STRESS: one-off 31% cut to new-deal $/MW once Rubin ships (never stacked with -6%/yr; sensitivity only)", "0/1", 0, "Skeptic", "Estimated", "", fmt="0")
S("MYStressCut", "Stress cut", "%", 0.31, "Skeptic", "Estimated", "", fmt="0%")
S("ND_GenOn", "Rubin new-deal $/MW = GB300 $/MW x Rubin_Prem x GPUMW_N / ND_GPUMW_GB (1 = v0.6; 0 = v0.5 single $/MW for all generations)", "0/1", 1, "Modeling", "Input", "Rubin at 1.2x and 320 GPUs/IT MW: $17.2M/IT MW (IT reading), -22% vs GB300", fmt="0")
S("RubinMode", "Rubin new-deal revenue: 1 = direct $/IT MW-yr (RubinMW; v0.7, no IT/gross toggle); 0 = v0.6 GB300 x Rubin_Prem x 320/491", "#", 1, "Modeling", "Input", "IT/gross toggle applies to GB300-era deals only (avoids Filings-flagged double count)", fmt="0")
S("RubinMW", "Rubin multi-year ACV per IT MW-yr (new deals, signing at Rubin availability; then MYDecl path within generation)", "$M/IT MW-yr", 16.2, "GPU Market", "Estimated", "Two contracts: Anthropic-Nscale ~$16.2M/MW (6-yr, MW basis unknown); WhiteFiber-Prime Intellect 576 VR200 ~$108.2M/3 yrs = ~$7.15/GPU-hr ~ $20M/IT MW if packages counted (medium). v0.7c base $16.2M (Anthropic-Nscale) / bull $20M / bear $14.5M (v0.7b 18/20/16). Implied $/GPU-hr at 320: 5.78 / 7.13 / 5.17; ~1.55x GB300 $16M per GPU-hr at base", bull=20, bear=14.5, fmt="0.0")
S("RubinGPUMW_D", "Rubin density for $/GPU-hr reference only (sensitivity; not used in revenue)", "GPU/MW", 400, "Capacity", "Estimated", "320 vs 400 packages/IT MW: reference only", fmt="0")
S("STShare", "Short-term (<=6-month) tranche share of TOTAL average active IT MW (from Q1'27)", "%", 0.05, "GPU Market", "Estimated", "v0.5: 5% of total capacity (on-demand/self-serve ~4% of Q2'26 revenue; one short-term deal signed); tests 2% / 10%; 0% floor row", fmt="0%")
S("STMode", "Short-term tranche mode: 1 = v0.5 share of TOTAL active MW x STOcc; 0 = v0.4 share of new 2027+ MW, no occupancy haircut", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("DRSchedMode", "Deferred revenue release: 1 = v0.5 (6/30 balance on Reported split; new prepayments straight-line over DR_AmortQ); 0 = v0.4 (balance / DR_AmortQ geometric)", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("STOcc", "Short-term tranche occupancy between deals", "%", 0.85, "GPU Market", "Estimated", "$45M x 85% ~ $38M per IT MW-yr effective yield", fmt="0%")
S("MYRateU", "Multi-year GB300-equivalent ACV for UNSOLD new MW ($M/IT MW-yr; contracted backlog stays at MYRate; Rubin scales by the same ratio)", "$M/IT MW-yr", "=MYRate", "Skeptic", "Estimated", "Headline sens $18 / $22 / $25 (Skeptic); default = MYRate; contracted vs unsold from RPO until Capacity site tags (Tuesday)", fmt="0.0")
S("RPO_Tot", "Remaining performance obligations at 6/30/26", "$M", 37490, "Filings", "Reported", "Q2'26 6-K Note 2", fmt="#,##0")
S("RPO24Pct", "Share of RPO recognised Jul'26-Jun'28", "%", 0.36, "Filings", "Reported", "Q2'26 Note 2 (~$13.5B)", fmt="0%")
S("DR_Cur", "6/30/26 deferred revenue - current (recognised within 12 months)", "$M", 979.4, "Filings", "Reported", "Q2'26 6-K balance sheet; non-current $4,995.8M; total $5,975.2M; YE25 $1,577.5M (only $154.3M recognised in H1'26)", fmt="#,##0.0")
S("DR_24", "6/30/26 deferred revenue recognised within 24 months", "$M", 2250, "Filings", "Estimated", "$2.0-2.5B: months 13-24 at 1.0-1.5x the $979.4M current portion", fmt="#,##0")
S("DR_TailQ", "Quarters to release the remainder of the 6/30 balance after month 24", "qtrs", 12, "Filings", "Estimated", "Strategic contracts recognised over 1-5 yrs; tail Q3'28-Q2'31 (2 qtrs after horizon fall into the DR run-off deduction)", fmt="0")
S("NDBasis", "New-deal $/MW basis: 1 = per IT MW [base; Filings: same axis as 2026 base $12M/MW ~ 250 MW active path, medium]; 2 = slide is per GROSS MW -> MYRate x ND_GrossX per IT MW", "#", 1, "Filings", "Estimated", "", fmt="0")
S("ND_GrossX", "Gross-to-IT conversion for new-deal $/MW (GPUs per IT MW / GPUs per gross MW = 491/409)", "x", 491 / 409, "Capacity", "Estimated", "Capacity: 409 GPUs/gross MW vs 491/IT MW -> $22M per gross MW ~ $26.4M per IT MW", fmt="0.000")
S("STRate", "Short-term tranche $/IT MW-yr (near on-demand)", "$M/IT MW-yr", 45, "Filings", "Reported", "Q2'26 letter: Q3 short-term deals >$40M; $40-50M range", bull=50, bear=40, fmt="0.0")
S("STGBOnly", "Short-term tranche ($40-50M) on GB300-era capacity only (1 = v0.7; 0 = v0.6 any generation)", "0/1", 1, "Filings", "Input", "$40-50M is a GB300 rate, not a Rubin rate", fmt="0")
S("STRecov", "Short-term tranche recovers to list after post-Rubin cut (cut PxCyc for 3 qtrs after Rubin ships, then linear recovery over 2 qtrs)", "0/1", 1, "GPU Market", "Estimated", "Bull 6.25% with recovery; base 22% with recovery; bear 31% no recovery. Reprices every 6 months", bull=1, bear=0, fmt="0")
S("PrepayOn", "New-deal customer prepayments in liquidity (1 = on)", "0/1", 1, "Filings", "Reported", "70% of deals x 55% of capex = 38.5% of new-build capex, received at capex time, released to revenue over DR_AmortQ", fmt="0")
S("PrepayDecayOn", "Prepayment decay by signing period (1 = on; 0 = flat PrepayPct for all new non-hyperscaler deals)", "0/1", 1, "Skeptic", "Estimated", "v0.5: Neolab/other new deals 38.5% before Rubin ships, 22.5% for 2028+ signings; bull keeps 38.5%; bear 0% beyond already signed", fmt="0")
S("PrepayPct27", "Prepayment % of new-build capex - other new deals, capex in 2027 (Rubin ships Q1'27)", "%", 0.385, "Skeptic", "Estimated", "Capex leads signing by ~2-4 qtrs, so 2027 capex treated as pre-Rubin signings (input); bear 0% beyond already-signed", bull=0.385, bear=0.0, fmt="0.0%")
S("PrepayPct28", "Prepayment % of new-build capex - other new deals, 2028+ signings (capex from PrepayIdx28)", "%", 0.225, "Skeptic", "Estimated", "Bull keeps 38.5%; bear 0%", bull=0.385, bear=0.0, fmt="0.0%")
S("PrepayIdx28", "First capex quarter at the 2028+ prepayment rate", "idx", 17, "Skeptic", "Estimated", "Q1'28", fmt="0")
S("HSPre", "Prepayment % of related capex - NEW hyperscaler capacity (future IG / hyperscaler tranches)", "%", 0.30, "GPU Market", "Estimated", "Range 20-40%: Microsoft-Nebius ~40% (~$7.0B of $17.39B), Microsoft-IREN ~20% (GPU Market, medium); bull 40%, bear 0%. Existing signed Microsoft/Meta tranches: 0% new prepayment (advances already in 6/30 DR $5,975M)", bull=0.40, bear=0.0, fmt="0%")
S("DRTermOn", "Closing deferred revenue at Q4'30 runs off as non-cash revenue after the horizon: deduct PV from EV (1 = on)", "0/1", 1, "Modeling", "Input", "v0.5 cash double-count fix: terminal UFCF is EBITDA-based, so the unreleased prepaid balance would otherwise be counted as cash twice", fmt="0")
S("PrepayByCust", "Prepayment by customer type (1 = v0.5: signed Microsoft/Meta 0%, new hyperscaler HSPre, other deals at rate path; 0 = v0.4 flat PrepayPct on all new-build capex)", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("PrepayBase", "Prepayment base: 1 = new-build capex (GPU + shell + power) [v0.4]; 0 = total capex (v0.3)", "0/1", 1, "Filings", "Input", "", fmt="0")
S("GPUAnchorOn", "GPU history anchor: ~30,000 GPUs at 3/31/2025 (single quarter-end actual) - add gap vs Capacity file to fleet from Q1'25 (1 = on)", "0/1", 1, "Filings", "Reported", "FY24 20-F: ~30,000 GPUs at 3/31/2025 (the ~38k was a target). Kansas City likely live Q1'25 at 5-15 MW (Capacity, Estimated) - gap carried as Hopper-era fleet", fmt="0")
S("GPU_Q125", "GPUs at 3/31/2025 (actual)", "GPUs", 30000, "Filings", "Reported", "FY24 20-F", fmt="#,##0")
S("ReplPxOn", "Replacement capex at scenario GPU price path: cohort cost x (1 + annual list-price change of its gen)^(life yrs) (1 = on)", "0/1", 0, "Modeling", "Input", "Bear on for internal consistency (bear rental prices fall 20-30%/yr; constant-$ replacement overstated capex vs revenue). Base/bull off (constant $, conservative); see Sensitivity", bull=0, bear=1, fmt="0")
S("OpexHold", "Hold product development + SG&A at FY28 % of revenue for FY29-30 (1 = on; sensitivity)", "0/1", 0, "Modeling", "Input", "Default 0 = placeholder step-down (PD 8%->6%, SG&A 10%->7%)", fmt="0")
S("TermMode", "Terminal capex: 1 = max(D&A, GPU replacement + non-GPU + growth reinvestment); 0 = D&A x MaintRatio (old)", "0/1", 1, "Modeling", "Input", "", fmt="0")
S("RenewRate", "Reserved book renewal pace (share of book repricing to current list per quarter from RenewIdx)", "%/qtr", 0.125, "GPU Market", "Estimated", "Nebius reserved terms max 24 months (3/6/12 mo 2024; 6/12/24 on H200); Q2 letter >30% repricing on renewals. 50% of the reserved book renews within FY27 in ALL scenarios (12.5%/qtr from Q1'27, 100% by Q4'28). GPU Market archived pricing pages, Estimated", fmt="0.0%")
S("RenewIdx", "Renewal schedule start quarter (first quarter after 6/30)", "idx", 11, "GPU Market", "Estimated", "Q1'27", fmt="0")
S("Sk3_ZeroQ", "Skeptic v3.2 path-tied capex base: first quarter liquidity < 0", "qtr", "Q4'27", "Skeptic", "Estimated", "/workspace/nebius/skeptic/skeptic_03_depreciation.xlsx; Skeptic path run uses CONSENSUS EBITDA", fmt="@")
S("Sk3_YE27", "Skeptic v3.2 base YE27 liquidity", "$B", -0.15, "Skeptic", "Estimated", "", fmt="0.00")
S("Sk3_YE28", "Skeptic v3.2 base YE28 liquidity", "$B", -6.8, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3L_ZeroQ", "Skeptic v3.2 low: first quarter < 0", "qtr", "never thru 2030", "Skeptic", "Estimated", "", fmt="@")
S("Sk3L_YE27", "Skeptic v3.2 low YE27 liquidity", "$B", 7.1, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3L_YE28", "Skeptic v3.2 low YE28 liquidity", "$B", 1.5, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3H_ZeroQ", "Skeptic v3.2 high: first quarter < 0", "qtr", "Q3'27", "Skeptic", "Estimated", "", fmt="@")
S("Sk3H_YE27", "Skeptic v3.2 high YE27 liquidity", "$B", -7.4, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3H_YE28", "Skeptic v3.2 high YE28 liquidity", "$B", -15.6, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3C_ZeroQ", "Skeptic v3.2 consensus-capex run: first quarter < 0", "qtr", "Q2'27", "Skeptic", "Estimated", "", fmt="@")
S("Sk3C_YE27", "Skeptic v3.2 consensus-capex YE27 liquidity", "$B", -9.9, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3_DCF4", "Skeptic v3.2 DCF EV at 4-yr life", "$B", "negative", "Skeptic", "Estimated", "", fmt="@")
S("Sk3_DCF5", "Skeptic v3.2 DCF EV at 5-yr life", "$B", 2.3, "Skeptic", "Estimated", "vs $74.7B market EV", fmt="0.0")
S("Sk3_DCF6", "Skeptic v3.2 DCF EV at 6-yr life", "$B", 16.3, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk3_PC27", "Skeptic price-cycle leg FY27 revenue (50% renewal)", "$M", -340, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk3_PC27Bear", "Skeptic price-cycle leg FY27 revenue (100% renewal)", "$M", -782, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk3_PC28", "Skeptic price-cycle leg FY28 revenue", "$M", -1600, "Skeptic", "Estimated", "", fmt="#,##0")
S("Chk_SrvLo", "Capacity server+network cost per GPU - low", "$/GPU", 45000, "Capacity", "Estimated", "servers + network", fmt="#,##0")
S("Chk_SrvHi", "Capacity server+network cost per GPU - high", "$/GPU", 60000, "Capacity", "Estimated", "", fmt="#,##0")
S("CPM_SP", "Data hall + power/cooling portion of owned capex per IT MW (not replaced on the GPU cycle)", "$M/MW", 12.5, "Capacity", "Derived", "gpu_mw_capex.xlsx CapexPerMW: GB300/VR owned base DC shell+power $12.5M/IT MW (range 8-17.6); owned GB300 53 = 34.1 servers + 6.1 network + 12.5 hall; owned VR 56 = 43.5 + 12.5. v0.7 was 11.25", fmt="0.00")
S("PxCycleOn", "Price-cycle leg: one-off post-launch cut on prior gen (Rubin 1H27 hits Blackwell)", "0/1", 1, "GPU Market", "Input", "v0.4: on in all scenarios (bull 6.25% / base 22% / bear 31%). On-demand immediately; reserved at renewal; take-or-pay unaffected", bull=1, bear=1, fmt="0")
S("PxCyc", "Price-cycle cut in $/GPU-hr over 12 months after new-gen launch (Blackwell, from Rubin)", "%", 0.22, "GPU Market", "Estimated", "GPU Market gpu_price_history.xlsx, on-demand list, Estimated. Base 22% = H200 median $5.15 (Dec-24) -> $4.04 (Sep-25); bear 31% = Nebius B200 $8.50 premium closes to $7.49 median then 22% swing (~$5.84); bull 6.25% = midpoint 0-12.5% (H100 median flat $3.44 Jun-25..Mar-26 after B200 GA; B200 -2% since CoreWeave Rubin launch Sep-30-26)", bull=0.0625, bear=0.31, fmt="0.00%")
S("PxCycIdx", "Price-cycle start quarter (Rubin launch)", "idx", 13, "GPU Market", "Estimated", "Q1'27 (1H27)", fmt="0")

sec("Capacity & power")
S("MW_Haircut", "Forecast MW haircut (x driver MW)", "x", 1.00, "Skeptic", "Placeholder", "Bear case: execution slippage on site ramps", bull=1.00, bear=0.85)
S("MW_Delay", "Additional forecast MW slip (Skeptic overlay)", "qtrs", 0, "Skeptic", "Placeholder", "0 in all scenarios; energized low already embeds bear slippage; stress = EnLag 2", bull=0, bear=0, fmt="0")
S("PUE", "PUE (facility MW / IT MW)", "x", 1.25, "Capacity", "Estimated", "Converts facility-basis site MW to IT MW and IT load to grid power; modern liquid-cooled AI DC ~1.15-1.35", bull=1.20, bear=1.35)
S("LoadFactor", "Average IT draw / IT capacity", "%", 0.80, "Capacity", "Placeholder", "Power-cost driver", fmt="0%")
S("GPUMW_H", "GPUs per IT MW - Hopper (H100/H200)", "GPU/MW", 713, "Capacity", "Estimated", "~1.47 kW per GPU all-in IT (HGX H100 ~10.2 kW/8 GPUs + ~15% network/storage); not disclosed by Nebius", fmt="#,##0")
S("GPUMW_B", "GPUs per IT MW - Blackwell (B200/B300/GB200/GB300)", "GPU/MW", 500, "Capacity", "Estimated", "~2.1 kW per GPU all-in IT (GB200 NVL72 ~132 kW/72 + overhead); not disclosed", fmt="#,##0")
S("GPUMW_N", "GPUs per IT MW - Next-gen (Vera Rubin)", "GPU/MW", 320, "Capacity", "Placeholder", "~3.0 kW per GPU all-in IT assumed; awaiting NVIDIA VR NVL144 rack power data", fmt="#,##0")

sec("Pricing (chip-level list prices on Pricing tab)")
S("ResDisc", "Reserved/commit discount vs on-demand", "%", 0.35, "GPU Market", "Derived", "nebius.com/prices: 'Pay up to 35% less than on-demand' -> reserved floor = OD x (1-35%)", bull=0.25, bear=0.35, fmt="0%")
S("PxMult", "Forecast list-price multiplier (on-demand, reserved, spot)", "x", 1.00, "GPU Market", "Placeholder", "Scenario lever on books (b) and (c); book (a) contracts are fixed $/MW-yr", bull=1.05, bear=0.90)
S("PxPathMode", "Price path: 1 = v0.4 (one-off post-Rubin Blackwell cut, then annual decline from year 2: Blackwell from PxCycIdx+4, Next-gen from NextLaunchIdx+4); 0 = v0.3 (annual change from PxBaseIdx)", "0/1", 1, "GPU Market", "Input", "", fmt="0")
S("PwrLooseOn", "Bear stress: power loosens 2028 - steeper Blackwell/Next-gen decline from PwrLooseIdx (1 = on)", "0/1", 0, "Skeptic", "Placeholder", "Stress toggle; pending Skeptic's source", fmt="0")
S("PwrLooseDecl", "Annual list-price change when power loosens", "%/yr", -0.15, "Skeptic", "Placeholder", "Placeholder pending Skeptic's source", fmt="0%")
S("PwrLooseIdx", "Power-loosening start quarter", "idx", 17, "Skeptic", "Placeholder", "Q1'28", fmt="0")
S("PxChg_H", "Annual list-price change - Hopper", "%/yr", -0.15, "GPU Market", "Placeholder", "Applied from PxBaseIdx; bear owned by Skeptic", bull=-0.05, bear=-0.30, fmt="0%")
S("PxChg_B", "Annual list-price change - Blackwell (year 2 onward, after post-Rubin cut)", "%/yr", -0.06, "GPU Market", "Estimated", "GPU Market, from archived list prices: bull -2.5% / base -6% / bear -9.5% per yr from year 2, plus one-off post-Rubin cut (PxCyc 6.25/22/31%). v0.3 was -10% / 0% / -25% from Q4'26", bull=-0.025, bear=-0.095, fmt="0.0%")
S("PxChg_N", "Annual list-price change - Next-gen (year 2 onward after launch)", "%/yr", -0.06, "GPU Market", "Estimated", "GPU Market, from archived list prices: bull -2.5% / base -6% / bear -9.5%. v0.3 was -1.7% / -5% / -20% from Q4'26", bull=-0.025, bear=-0.095, fmt="0.0%")
S("PxBaseIdx", "First quarter on new (Oct 1 2026) list prices", "idx", 12, "GPU Market", "Reported", "nebius.com/prices 'Effective October 1, 2026' column -> Q4'26 (idx 12)", fmt="0")
S("GB_Prem", "GB200/GB300 NVL72 price premium vs HGX B300", "x", 1.22, "GPU Market", "Derived", "Nebius = 'Contact us'. Proxy: CoreWeave GB200 $10.50 vs B200 $8.60 per GPU-hr (Oct 4 2026 snapshot)")
S("Rubin_Prem", "Vera Rubin $/GPU-hr multiple vs GB300 - DERIVED (v0.7c): Rubin priced directly per IT MW (RubinMW); multiple = (RubinMW/GPUMW_N)/(MYRate/ND_GPUMW_GB)", "x", "=(RubinMW/GPUMW_N)/(MYRate/ND_GPUMW_GB)", "GPU Market", "Derived", "v0.7b input 1.2x removed (it no longer drove any output under RubinMode 1); now an output, used only by the on-demand list-price placeholder for Rubin chips", fmt="0.00")
S("LaunchOn", "Launch-premium driver on (0/1)", "0/1", 1, "GPU Market", "Input", "Toggle kept; default premium 0% / head-start 0 months", fmt="0")
S("LaunchPrem", "Nebius launch premium on new generation vs peers", "%", 0.0, "GPU Market", "Derived", "GPU Market & Comps: Nebius trails CoreWeave by 2-6 months each generation since H200 and launched at/below peer prices -> 0%", fmt="0%")
S("HeadStart", "Head-start window for launch premium", "months", 0, "GPU Market", "Derived", "As above -> 0 months", fmt="0")
S("NextLaunchIdx", "Next-gen (Rubin) launch quarter", "idx", 13, "GPU Market", "Estimated", "Meta #2 Vera Rubin from early 2027", fmt="0")
S("RampMS", "Microsoft tranche billing ramp: quarters from delivery to full run-rate", "qtrs", 4, "Capacity/Modeling", "Derived", "Calibrated: Q2'26 MS revenue ~$121M = T1 (3/4) + T2 (2/4) x $96.6M/qtr (equal split); ramp defers, not recovered (conservative)", bull=3, bear=5, fmt="0")
S("RealOD", "Price realization vs list - on-demand (book c)", "x", 0.95, "GPU Market", "Estimated", "bear 0.85 / base 0.95 / bull 1.00", bull=1.00, bear=0.85, fmt="0.00")
S("RealRes", "Price realization vs on-demand list - reserved/commit (book b)", "x", 0.72, "GPU Market", "Estimated", "bear 0.65 / base 0.72 / bull 0.85 (list floor = OD x 0.65)", bull=0.85, bear=0.65, fmt="0.00")
for _n in ["RealOD", "RealRes", "RampMS", "PUE"]:  # base-case value: Q2'26 history/calibration is a fact, scenario factors apply to forecast only
    name(wb, _n + "_B", "Inputs", wb.defined_names[_n].attr_text.split("!")[1].replace("$F$", "$D$"))
S("Util_OD", "Forecast utilization - on-demand book (c)", "%", "=IF(UtilQ2>1,1,IF(UtilQ2<0.4,0.4,UtilQ2))", "GPU Market", "Placeholder", "Base linked to Q2-26 calibrated utilization, bounded to 40-100% for the FORECAST only (calibration itself is flagged, not forced)", bull=0.90, bear=0.65, fmt="0%")
S("Util_Res", "Forecast utilization - reserved/commit book (b)", "%", "=IF(UtilQ2>1,1,IF(UtilQ2<0.4,0.4,UtilQ2))", "GPU Market", "Placeholder", "Base linked to Q2'26 calibrated utilization", bull=0.95, bear=0.75, fmt="0%")
S("Util_Spot", "Billable utilization - spot/preemptible", "%", 0.40, "GPU Market", "Placeholder", "", bull=0.50, bear=0.30, fmt="0%")
S("ResidUtil", "Post-term merchant utilization (residual value of contract GPUs)", "%", 0.65, "Skeptic/GPU Market", "Placeholder", "Bear = A100 decay proxy at ~60% utilization (Skeptic)", bull=0.70, bear=0.60, fmt="0%")

sec("Hyperscaler take-or-pay contracts (tranche detail on Contracts tab)")
S("MSFT_TCV", "Microsoft committed TCV", "$M", 17392.9, "Filings", "Reported", "6-K Sep 8 2025 + 20-F FY25 Note 1: committed fees up to $17,392.9M, 5-yr, Vineland NJ, tranches 2025-26; up to ~$19.4B with extra capacity")
S("MSFT_MW", "Microsoft dedicated IT MW (total across 9 tranches)", "MW", 240, "Capacity", "Estimated", "Equal-size ASSUMPTION ~26.7 IT MW per tranche (not disclosed); Vineland 300 MW gross", fmt="#,##0")
S("MSFT_N", "Microsoft tranche count (base contract)", "#", 9, "Filings", "Reported", "20-F FY25: nine tranches during 2025-2026", fmt="0")
S("ACV_2026", "ACV per MW - 2026 base capacity", "$M/MW-yr", 12.0, "Filings", "Reported", "Q2'26 shareholder letter p.1 chart: '2026 base ~$12M' ACV per MW (approximate)")
S("ACV_Q2", "ACV per MW - Q2'26 deals / new-gen capacity", "$M/MW-yr", 22.5, "Filings", "Derived", "Q2'26 letter: $20-25M ACV per MW (midpoint)")
S("T59Align", "Vineland T5-T9 billing aligned to Capacity v2 (1 = full run-rate from the Capacity billing-start quarter, no extra ramp / mid-quarter haircut; 0 = v0.7c RampMS + DelivFactor)", "0/1", 1, "Capacity", "Estimated", "Capacity v2 billing start = Bloom energization + >=1 qtr (lag already in the date): T5-T6 Q2'27, T7-T9 Q3'27 base -> FY27 $1,159.5M at $96.6M/tranche-qtr. v0.7c applied a further 4-qtr ramp + 50% start quarter -> $447M", fmt="0")
S("DelivFactor", "Revenue factor in delivery quarter", "x", 0.50, "Modeling", "Estimated", "Mid-quarter convention")
S("GensetMW", "Microsoft/Vineland power ceiling before Bloom (gross MW)", "MW", 122.9, "Capacity", "Estimated", "62 rented Caterpillar gensets x 1.98 MW = 122.8 MW gross (Capacity & Power)", fmt="#,##0.0")
S("BloomOn", "Bloom fuel cells live at Vineland (0/1 toggle)", "0/1", 1, "Capacity", "Input", "1 = ceiling lifted from BloomIdx; 0 = genset cap forever", fmt="0")
S("BloomIdx", "Quarter Bloom fuel cells go live", "idx", 14, "Capacity", "Estimated", "Bull Q1'27 / base Q2'27 / bear Q3'27 (tranches 5-9 need Bloom)", bull=13, bear=15, fmt="0")
S("MSFT_Ceil", "Microsoft revenue ceiling check level", "$M/qtr", 350, "Capacity/Skeptic", "Estimated", "~$350M/qtr at genset-limited power (check row only)", fmt="#,##0")
S("TopCust", "Q2'26 24% customer = 1 Meta (default) / 2 Microsoft", "#", 1, "Filings", "Input", "Q2'26: customers C/D/E = 24%/21%/14% of revenue (names not disclosed)", fmt="0")
S("MS_Q2RR", "Microsoft Q2'26 revenue run rate (bridge start)", "$M/qtr", "=IF(TopCust=1,0.21,0.24)*574.9", "Filings", "Derived", "21% (if Meta = 24%) or 24% of Q2 AI cloud revenue $574.9M -> ~$121M / ~$138M", fmt="#,##0.0")
S("Meta1Q", "Meta #1 quarterly revenue (flat)", "$M/qtr", "=2880.7/20", "Filings", "Derived", "$2,880.7M total (Reported) / 20 quarters", fmt="#,##0.0")
S("PxLine_Lo", "Q4 bridge: Oct-1 list-price uplift - low", "$M/qtr", 40, "GPU Market", "Derived", "GPU Market & Comps: $40-125M/qtr", fmt="#,##0")
S("PxLine_Hi", "Q4 bridge: Oct-1 list-price uplift - high", "$M/qtr", 125, "GPU Market", "Derived", "", fmt="#,##0")
S("Unseen_OK", "Oklahoma (7 DC roles, live-looking) - IT MW from Q4'26 [Bull col = high, Base, Bear col = low]", "MW", 0, "Capacity", "Placeholder", "PLACEHOLDER - Capacity & Power low/base/high, due Tue Oct 6", bull=0, bear=0, fmt="#,##0")
S("Unseen_MN", "Minnesota (RMA roles, likely colo live) - IT MW from Q4'26", "MW", 0, "Capacity", "Placeholder", "PLACEHOLDER - due Tue Oct 6", bull=0, bear=0, fmt="#,##0")
S("Unseen_OT", "Other unidentified colo - IT MW from Q4'26", "MW", 0, "Capacity", "Placeholder", "PLACEHOLDER - due Tue Oct 6", bull=0, bear=0, fmt="#,##0")
S("VinPUE", "Vineland PUE used for genset ceiling", "x", 1.2, "Capacity", "Estimated", "Capacity file 'Vineland tranche track'")
S("ResLag", "Reserved/commit repricing lag (renewal) vs on-demand list", "qtrs", 4, "GPU Market", "Estimated", "List changes hit on-demand immediately; reserved reprices on renewal", bull=2, bear=6, fmt="0")
S("MixA_Q2", "Q2'26 revenue mix - hyperscaler take-or-pay (calibration target)", "%", 0.45, "GPU Market", "Estimated", "Range 40-46% (medium/low confidence)", fmt="0%")
S("MixB_Q2", "Q2'26 revenue mix - reserved/committed (calibration target)", "%", 0.50, "GPU Market", "Estimated", "Range 44-57%", fmt="0%")
S("MixC_Q2", "Q2'26 revenue mix - on-demand + Token Factory (calibration target)", "%", 0.04, "GPU Market", "Estimated", "Range 1.3-9%", fmt="0%")
S("UnseenChkLo", "Skeptic/Capacity estimate of unseen billing MW - low (check only)", "MW", 83, "Skeptic/Capacity", "Estimated", "Shown as a check against the bridge-implied MW, not used in calculations", fmt="#,##0")
S("UnseenChkHi", "Skeptic/Capacity estimate of unseen billing MW - high (check only)", "MW", 200, "Skeptic/Capacity", "Estimated", "", fmt="#,##0")
S("ResidBearCut", "Bear: GPU price at age 5 vs today's list (cut)", "%", 0.35, "Skeptic", "Estimated", "Skeptic: 22-48% below today's list (A100 2022-peak to 2025-low path); midpoint 35%", fmt="0%")
S("LD_Pct", "Late-delivery liquidated damages (% of scheduled qtr revenue per late quarter)", "%", 0.00, "Skeptic/Capacity", "Placeholder", "Contract LD terms not disclosed", fmt="0%")

sec("Other revenue")
S("OthRev_g", "Other businesses revenue growth (Avride, TripleTen, net of elims)", "%/qtr", 0.02, "Filings", "Placeholder", "Q2'26: Avride $1.0M + TripleTen $10.0M - elim $3.6M = $7.4M", bull=0.03, bear=0.0, fmt="0.0%")

sec("Operating costs (ex-SBC)")
S("PowerPx", "All-in power price", "$/MWh", 75, "Capacity", "Placeholder", "Blended across sites; not disclosed", bull=65, bear=90, fmt="#,##0")
S("RentMW", "Leased/colo/build-to-suit rent (shell + power infrastructure)", "$M/gross MW-yr", 2.25, "Capacity", "Estimated", "Capacity & Power: $2.0-2.5M per gross MW-yr (market $165-205/kW-month); applied to leased IT MW x PUE; deducted from EBITDA", bull=2.0, bear=2.5)
S("VinPower", "Vineland power cost (Bloom fuel-cell PPA / gensets) - opex", "$M/gross MW-yr", 1.1, "Capacity", "Derived", "Capacity & Power: ~$1.1M per gross MW-yr, range $93-185/MWh (low confidence); PPA = opex not capex", bull=0.9, bear=1.6)
S("OtherCOR", "Other cost of revenue (network, staff, support)", "% rev", 0.06, "Modeling", "Placeholder", "Calibrate to reported COR (Q2'26 23% of revenue incl. rent/power)", bull=0.05, bear=0.08, fmt="0.0%")
S("SBCpct", "Share-based compensation", "% rev", 0.04, "Filings", "Estimated", "Excluded from adj. EBITDA; below-EBITDA charge", fmt="0.0%")

sec("Capex & depreciation")
S("GPUCost_H", "All-in cost per GPU - Hopper (server, network, storage)", "$/GPU", 35000, "GPU Market", "Placeholder", "x GPUs/MW = ~$24M per IT MW", fmt="#,##0")
S("GPUCost_B", "All-in cost per GPU - Blackwell", "$/GPU", 60000, "GPU Market", "Placeholder", "x 475/MW = ~$28.5M per IT MW (vs Capacity colo-case ~$30M/IT MW)", fmt="#,##0")
S("GPUCost_N", "All-in cost per GPU - Next-gen", "$/GPU", 90000, "GPU Market", "Placeholder", "", fmt="#,##0")
S("ITCapexMW", "IT capex per IT MW (method 2, colo case)", "$M/MW", 30.0, "Capacity", "Estimated", "Capacity: ~$30M per IT MW (~$25M per gross MW)")
S("ShellMW", "Owned shell/building capex per IT MW (30% of data hall; book life Life_Shell 20y; no cash replacement in horizon)", "$M/MW", 3.75, "Capacity", "Placeholder", "PLACEHOLDER 30/70 split of $12.5M pending Capacity sourced split (Tue); v0.7 4.0. Land not depreciated")
S("PowerMW", "Owned power/cooling/electrical capex per IT MW (70% of data hall; book life Life_Pow 10y; cash-replaced at year 10)", "$M/MW", 8.75, "Capacity", "Placeholder", "PLACEHOLDER 30/70 split (20-F infrastructure 3-10y line, Estimated); v0.7 6.0")
S("GPULead", "IT/GPU (and rest) capex lead vs MW BILLING (qtrs; 1 or 2)", "qtrs", 2, "Capacity", "Estimated", "v0.4 split lead (gpu_mw_capex.xlsx): GPUs/rest 1-2 qtrs ahead of billing (default 2); shell/power via DCLead 3. v0.3 used uniform 3 (sensitivity). Evidence: 6/30 assets-not-in-use $7.84B; FY26 capex guide $20-25B. GPU depreciation lag = PISLag + GPULead - 1 for forecast cohorts (starts at billing). First-negative quarter Q2'27 vs Q3'27 PROVISIONAL pending June MW", fmt="0")
S("PISLag", "Placed-in-service lag: capex quarter -> depreciation start", "qtrs", 1, "Filings/Skeptic", "Estimated", "Separate from billing lag. Bear 0 = depreciation starts at energization (unremediated material weakness over depreciation start dates, FY24 & FY25 20-F; adverse ICFR opinion, Reanda)", bull=1, bear=0, fmt="0")
S("BillLag", "Mgmt CONNECTED power -> billing lag (MW source 3 only)", "qtrs", 1, "Capacity", "Estimated", "Call transcript: 'a few months' from CONNECTED power to revenue -> 1-2 qtrs", bull=1, bear=2, fmt="0")
S("LifeChgFY26", "FY26 depreciation benefit from 4->5 yr life change (guided)", "$M", 167.6, "Filings", "Reported", "20-F p.56 / Q2'26 Note 1", fmt="#,##0.0")
S("AvailIdx_H200", "H200 self-serve availability (Feb 4 2025)", "idx", 5, "GPU Market", "Reported", "", fmt="0")
S("AvailIdx_B200", "B200 GA (Jun 11 2025)", "idx", 6, "GPU Market", "Reported", "", fmt="0")
S("AvailIdx_GB300", "GB300/B300 in customer use (Dec 17 2025)", "idx", 8, "GPU Market", "Reported", "", fmt="0")
S("AvailIdx_VR", "Vera Rubin availability (not yet; placeholder)", "idx", 14, "GPU Market", "Placeholder", "Not yet available as of Oct 4 2026", fmt="0")
S("Cons_Q4", "Consensus Q4'26 revenue", "$M", 1450, "Skeptic", "Estimated", "Consensus, /workspace/nebius/skeptic/01_priced_in.xlsx (Oct 2 2026, medium confidence)", fmt="#,##0")
S("Q2GroupRev", "Q2'26 group revenue (bridge start)", "$M", 582.3, "Filings", "Reported", "Q2'26 6-K", fmt="#,##0.0")
S("NewDealLo", "New-deal rate low (bridge MW implied)", "$M/MW-yr", 20, "GPU Market", "Derived", "", fmt="0.0")
S("NewDealHi", "New-deal rate high", "$M/MW-yr", 25, "GPU Market", "Derived", "", fmt="0.0")
S("BaseRateMW", "Base revenue rate", "$M/MW-yr", 12, "GPU Market", "Derived", "", fmt="0.0")
S("CapexMode", "Capex toggle: 1 = consensus (FY27 = $29.1B via IT overlay cohort), 2 = path-consistent", "#", 2, "Modeling", "Input", "Path-consistent = Capacity energized path x capex per MW; GPU capex lands GPULead (1) qtr before billing (GPUs ship once site has power)", fmt="0")
S("CapIntPct", "Capitalized interest as % of gross convert + secured interest (forecast)", "%", 0.374, "Filings", "Derived", "Q2'26: 40.2 / 107.5 (Note 12); non-cash PP&E addition, reduces P&L interest, depreciated over CapIntLife", bull=0.374, bear=0.374, fmt="0.0%")
S("CapIntLife", "Depreciation life of capitalized interest", "yrs", 5, "Modeling", "Estimated", "Follows server life", fmt="0")
S("CIPRel1", "Release of Q2'26 assets-not-in-use into service in Q3'26", "%", 0.6, "Capacity/Filings", "Estimated", "Jul-Aug Vineland tranches T3-T4 + Q2 deployments", fmt="0%")
S("CIPRel2", "Release of Q2'26 assets-not-in-use into service in Q4'26", "%", 0.4, "Capacity/Filings", "Estimated", "", fmt="0%")
S("Q3TransAct", "Q3'26 transfer into service - actual (enter when Q3 reported)", "$M", None, "Filings", "Placeholder", "From Q3'26 Note 7", fmt="#,##0")
S("Q3TransMin", "Q3'26 transfer threshold for 'large transfer' check", "$M", 3000, "Filings", "Estimated", "Expect large transfer after Jul-Aug Vineland tranches", fmt="#,##0")
S("DCLead", "Shell/power capex lead vs MW BILLING (qtrs)", "qtrs", 3, "Capacity", "Estimated", "v0.4 split lead (gpu_mw_capex.xlsx): shell + power ~$10-12.5M/MW spent ~3 qtrs ahead of billing; depreciation from energization (PISLag + DCLead - 2 for forecast cohorts)", fmt="0")
S("MaintPct", "Maintenance/other capex", "% rev", 0.02, "Modeling", "Placeholder", "", fmt="0.0%")
S("HistGPUsh", "Historical capex split - IT/GPU share", "%", 0.80, "Filings", "Placeholder", "Split not disclosed; drives historical depreciation cohorts", fmt="0%")
S("HistShellSh", "Historical capex split - shell/building share", "%", 0.08, "Filings", "Placeholder", "", fmt="0%")
S("HistPowSh", "Historical capex split - power/electrical share", "%", 0.12, "Filings", "Placeholder", "", fmt="0%")
S("Life_H", "Useful life - Hopper vintage", "yrs", 5, "Filings/Skeptic", "Reported", "Servers & network 5 yrs from 1/1/2026 (was 4) - 20-F FY25 p.F-15; bull 6 / bear 4 (Skeptic) | GPU Market archived A100 list data: neocloud A100 composite $1.23 -> $2.00 (Nov 2020 -> Oct 2026); hyperscaler composite -1.8%/yr (Estimated support for 5-6 yr economic life)", bull=6, bear=4, fmt="0.0")
S("Life_B", "Useful life - Blackwell vintage", "yrs", 5, "Filings/Skeptic", "Reported", "Company policy 5 yrs; bull 6 / bear 4 | GPU Market archived A100 list data: neocloud A100 composite $1.23 -> $2.00 (Nov 2020 -> Oct 2026); hyperscaler composite -1.8%/yr (Estimated support for 5-6 yr economic life)", bull=6, bear=4, fmt="0.0")
S("Life_N", "Useful life - Next-gen vintage", "yrs", 5, "Filings/Skeptic", "Reported", "Company policy 5 yrs; bull 6 / bear 4", bull=6, bear=4, fmt="0.0")
S("LifeOverride", "GPU life override for ALL vintages (0 = use per-vintage)", "yrs", 0, "Modeling", "Input", "Set 4/5/6 to flex all vintages; Sensitivity tab shows 4/5/6 regardless", fmt="0")
S("Life_Shell", "Useful life - buildings/shell", "yrs", 20, "Filings", "Reported", "20-F FY25: buildings 20 yrs", fmt="0")
S("Life_Pow", "Useful life - power/electrical/infrastructure systems", "yrs", 10, "Filings", "Estimated", "20-F FY25: infrastructure systems 3-10 yrs (upper end)", fmt="0")
S("Life_Maint", "Useful life - maintenance/other capex", "yrs", 5, "Modeling", "Estimated", "", fmt="0")
S("ImpQ_H", "Impairment quarter index - Hopper (0 = none)", "idx", 0, "Skeptic", "Placeholder", "e.g. 14 = Q2'27; write-down of remaining NBV of vintage cohorts placed up to that quarter", fmt="0")
S("ImpPct_H", "Impairment % of NBV - Hopper", "%", 0.0, "Skeptic", "Placeholder", "", fmt="0%")
S("ImpQ_B", "Impairment quarter index - Blackwell (0 = none)", "idx", 0, "Skeptic", "Placeholder", "", fmt="0")
S("ImpPct_B", "Impairment % of NBV - Blackwell", "%", 0.0, "Skeptic", "Placeholder", "", fmt="0%")
S("ImpQ_N", "Impairment quarter index - Next-gen (0 = none)", "idx", 0, "Skeptic", "Placeholder", "", fmt="0")
S("ImpPct_N", "Impairment % of NBV - Next-gen", "%", 0.0, "Skeptic", "Placeholder", "", fmt="0%")
S("OpenGPU_H", "Opening (pre-2024) Hopper gross cost", "$M", 142.4, "Filings", "Derived", "Proxy = Q1'24 reported D&A 8.9 x 16 quarters (4-yr life then in force)")
S("OtherDA", "Other D&A (intangibles, software) per quarter", "$M/qtr", 5.0, "Filings", "Placeholder", "Acquired intangibles (Tavily, Clarifai, Eigen AI) not split out")

sec("Customer prepayments / deferred revenue")
S("DR_Open", "Deferred revenue at Q2'26 (current + non-current)", "$M", 5975.2, "Filings", "Reported", "Q2'26 interim FS: 979.4 current + 4,995.8 non-current")
S("DR_YE25", "Deferred revenue at Q4'25", "$M", 1577.5, "Filings", "Reported", "275.5 + 1,302.0")
S("Prepay26", "FY2026 customer prepayments guided", "$M", 9000, "Filings", "Reported", "Q2'26 letter: 'over $9 billion in customer prepayments in 2026' (lower bound)")
S("PrepayPct", "New prepayments as % of capex (post 6/30)", "%", 0.385, "Filings/Skeptic", "Derived", "70% of deals x 55% of capex (Q2'26 letter: 50-60% of capex self-financed). Only NEW prepayments are cash; $5.975B opening DR is recognition only", bull=0.45, bear=0.25, fmt="0.0%")
S("DR_AmortQ", "NEW prepayment release period (straight-line from billing start)", "qtrs", 16, "Modeling", "Estimated", "v0.5: multi-year tranches 3-5 yrs -> 4 yrs; slow burn-off (6/30 balance has its own Reported-split schedule)", fmt="0")

sec("Financing & capital structure (convert tranches on Financing tab)")
S("LoanOn", "$775M asset-backed secured facility (SOFR+2.50%) drawn? 0 = $0 drawn (v0.7 base; drawn amount undisclosed), 1 = $775M (sensitivity)", "0/1", 0, "Filings", "Input", "Drives cash, debt, interest AND runway consistently. Skeptic rerun basis includes it ($14.6B start)", fmt="0")
S("ATMOn", "Runway: include remaining ATM (~12.27M sh at SharePx)", "0/1", 0, "Filings", "Input", "Excluded by default", fmt="0")
S("AugGross", "Aug-2026 converts gross ($3.45B 2030s + $2.30B 2034s)", "$M", 5750, "Filings", "Reported", "CapStructure", fmt="#,##0")
S("CS_Cash", "Unrestricted cash 6/30/26", "$M", 8042.1, "Filings", "Reported", "Excludes $1,056M restricted", fmt="#,##0.0")
S("CS_Restr", "Restricted cash 6/30/26 (excluded)", "$M", 1056, "Filings", "Reported", "", fmt="#,##0")
S("CS_ConvCarry", "Older converts carrying, pro forma for Aug exchange", "$M", 7697.1, "Filings", "Derived", "Sum of 6/30 carrying (8,499.0) less 87.5% of 2029/2031 carrying exchanged (801.9)", fmt="#,##0.0")
S("CS_CurDebt", "Debt, current 6/30/26", "$M", 46.7, "Filings", "Reported", "", fmt="#,##0.0")
S("Sk_Start", "Skeptic rerun: starting liquidity", "$B", 14.6, "Skeptic", "Estimated", "Same basis (cash + Aug gross + loan)", fmt="0.0")
S("Sk_Q127", "Skeptic rerun: liquidity end Q1'27", "$B", 1.1, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk_Q227", "Skeptic rerun: liquidity end Q2'27", "$B", -2.7, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk_YE27", "Skeptic rerun: liquidity YE27", "$B", -9.9, "Skeptic", "Estimated", "", fmt="0.0")
S("Sk_Rev27Lo", "Skeptic v2 FY27 revenue - low (1-qtr lag)", "$M", 7000, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk_Rev27Base", "Skeptic v2 FY27 revenue - base", "$M", 9900, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk_Rev27Hi", "Skeptic v2 FY27 revenue - high", "$M", 12900, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk_Rev27St", "Skeptic v2 FY27 revenue - stress (2-qtr lag)", "$M", 7600, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk_Rev28", "Skeptic v2 FY28 revenue", "$M", 25300, "Skeptic", "Estimated", "", fmt="#,##0")
S("Sk_EVEB27", "Skeptic v2 EV / FY27 EBITDA - Nebius", "x", 11.1, "Skeptic", "Estimated", "", fmt="0.0")
S("CRWV_EVEB27", "CoreWeave EV / FY27 EBITDA", "x", 4.8, "Skeptic", "Estimated", "", fmt="0.0")
S("Chk_RevGPUhr", "Q2'26 revenue per GPU-hour of capacity (check)", "$/GPU-hr", 3.27, "GPU Market", "Derived", "$2.30B / 80,400 GPUs / 8,760", fmt="0.00")
S("Chk_RevSrv", "Q2'26 annualized AI revenue per $ avg in-service servers (check)", "x", 0.46, "Filings", "Derived", "2.30 / avg(4.69, 5.20)", fmt="0.00")
S("Chk_SrvCap", "Gross servers & network in service 6/30/26", "$M", 5197.7, "Filings", "Reported", "Note 7", fmt="#,##0")
S("Chk_RevEnMW", "Q2'26 revenue per energized MW (check)", "$M/MW-yr", 13, "Capacity/GPU Market", "Derived", "2.30B / ~173 MW", fmt="0.0")
S("MinCash", "Minimum cash balance", "$M", 2000, "Modeling", "Placeholder", "Funding plug trigger", fmt="#,##0")
S("CashYield", "Interest income yield on cash", "%", 0.035, "Modeling", "Placeholder", "", fmt="0.00%")
S("SOFR", "SOFR (forward avg)", "%", 0.040, "Filings", "Placeholder", "", fmt="0.00%")
S("SecSpread", "Secured term loan spread", "%", 0.025, "Filings", "Reported", "Jul-2026 $775M secured term loan at SOFR+2.50% (6-K/A Jul 17 2026)", fmt="0.00%")
S("SecAmt", "Secured term loan amount", "$M", 775, "Filings", "Reported", "Signed 7/10/26; Nebius Compute II; MUFG lead", fmt="#,##0")
S("SecIdx", "Secured loan draw quarter", "idx", 11, "Filings", "Reported", "Q3'26", fmt="0")
S("SecMatIdx", "Secured loan maturity quarter", "idx", 28, "Filings", "Reported", "Matures 10/31/2030 (Q4'30); amortization not disclosed -> bullet assumed", fmt="0")
S("PlugSpread", "New debt (funding plug) spread over SOFR", "%", 0.030, "Filings", "Placeholder", "", fmt="0.00%")
S("PlugDebtSh", "Funding gap filled with debt (rest = equity)", "%", 0.50, "Modeling", "Placeholder", "", bull=0.60, bear=0.30, fmt="0%")
S("EqDisc", "Equity issuance discount to share price", "%", 0.03, "Modeling", "Placeholder", "", fmt="0%")
S("AugNet", "Aug-2026 convert net proceeds", "$M", 5680, "Filings", "Reported", "Pricing PR Aug 19 2026: ~$5.68B if options fully exercised (they were)", fmt="#,##0")
S("ExchShares", "Shares issued in Aug-2026 exchange of 2029/2031 notes", "M sh", 15.77, "Filings", "Reported", "Filings & Capital: 15.77M (PR: ~15.8M) Class A shares for $400M 2029s + $400M 2031s")
S("ExchIdx", "Exchange settlement quarter", "idx", 11, "Filings", "Reported", "Q3'26 (Aug 24 2026)", fmt="0")
S("ATM_Q3", "ATM shares sold Q3'26", "M sh", 0.0, "Filings", "Not disclosed", "No post-6/30 ATM sales disclosed; ~0 in July (AGM record-date count); 12.27M remaining")
S("DilMethod", "Convert dilution method (1 if-converted all, 2 ITM only, 3 net-share)", "#", 2, "Modeling", "Input", "Notes settle in cash/shares/combination at Nebius's election", fmt="0")
S("NVDA_W", "NVIDIA pre-funded warrant shares", "M sh", 21.065936, "Filings", "Reported", "SPA 3/11/26, $0.0001 exercise; in basic EPS count but not in issued/outstanding", fmt="0.000")
S("NVDA_Idx", "Warrant issue quarter", "idx", 9, "Filings", "Reported", "Q1'26", fmt="0")
S("Options", "Share options outstanding", "M sh", 6.7406, "Filings", "Reported", "Q2'26 FS Note 13: 6,740,600, WAEP $88.65", fmt="0.000")
S("OptK", "Options weighted-average exercise price", "$/sh", 88.65, "Filings", "Reported", "")
S("RSU", "Unvested RSUs", "M sh", 6.234091, "Filings", "Reported", "Q2'26 FS Note 13: 6,234,091", fmt="0.000")
S("Eigen", "Unvested Eigen AI consideration shares", "M sh", 2.743655, "Filings", "Reported", "Q2'26 Note 13; excluded from FD by default (not in Filings CapStructure tie-out)", fmt="0.000")
S("EigenIn", "Include Eigen AI shares in diluted count (0/1)", "0/1", 0, "Modeling", "Input", "", fmt="0")
S("ATM_Rem", "ATM shares remaining (funding source, NOT in today's count)", "M sh", 12.270507, "Filings", "Reported", "12,270,507 at 6/30/26; used only in FD tie-out check", fmt="0.000")
S("FD_Filings", "Fully diluted shares at $242.81 EXCL. unsold ATM (Filings CapStructure)", "M sh", 369.2, "Filings", "Reported", "/workspace/nebius/filings/nebius_filings.xlsx CapStructure (check only)", fmt="0.0")
S("FD_FilingsATM", "Fully diluted incl. 12.27M unsold ATM (Filings CapStructure)", "M sh", 381.5, "Filings", "Reported", "check only", fmt="0.0")
S("ITM_Sh_Chk", "ITM convert shares at $242.81 (Filings)", "M sh", 50.03, "Filings", "Reported", "Six ITM converts", fmt="0.00")
S("ITM_Prin_Chk", "ITM convert principal at $242.81 (Filings)", "$M", 7680, "Filings", "Reported", "", fmt="#,##0")
S("CRWV_EV", "CoreWeave EV at 6/30/26", "$M", 79500, "GPU Market", "Derived", "GPU Market & Comps", fmt="#,##0")
S("CRWV_MW", "CoreWeave active MW", "MW", 1500, "GPU Market", "Derived", "", fmt="#,##0")
S("CRWV_Contr", "CoreWeave contracted MW", "MW", 3700, "GPU Market", "Derived", "", fmt="#,##0")
S("CRWV_Lease", "CoreWeave commenced lease liabilities", "$M", 0, "GPU Market", "Placeholder", "", fmt="#,##0")
S("CRWV_Commit", "CoreWeave other commitments (leases not commenced, power)", "$M", 0, "GPU Market", "Placeholder", "", fmt="#,##0")
S("DilPxMode", "Price used for quarterly convert/option dilution (1 = SharePx, 2 = solved valuation price)", "#", 2, "Modeling", "Input", "Valuation solves price and dilution jointly (unrolled fixed-point, no circularity)", fmt="0")
S("SBCsh", "Net share issuance from new equity comp grants", "%/qtr", 0.0, "Filings", "Placeholder", "Default 0 (existing RSUs/options counted separately); set >0 for future grants", fmt="0.00%")
S("TaxRate", "Tax rate on positive EBT", "%", 0.258, "Filings", "Estimated", "NL statutory 25.8%; NOLs not modeled", fmt="0.0%")
S("CVDebtBasis", "Converts left in debt (OTM at solved price): 1 = original principal (v0.7); 2 = linear accretion from face (v0.6); 3 = CARRYING value: Note 12 amortized cost 6/30/26 ($8,499.0M pre-Aug series) / Aug-26 notes at $5.75B face, accreted to redemption amounts, rolled to the valuation qtr [v0.7b]", "#", 3, "Filings", "Input", "Filings CapStructure: $13.425B remaining original principal; NSS cash H = $1,000 per $1k original", fmt="0")
sec("Funding plan v0.7 (ESTIMATED, not guidance): remaining ATM + two ~$5B convert deals in 2027 (Aug-26 terms); bear = converts shut, ATM at a discount")
S("FP_On", "Funding plan on (1) / off (0 = v0.6, no 2027 raises)", "0/1", 1, "Modeling", "Estimated", "Precedent (Filings): ~$21B raised externally in 15 months - $14.25B converts in 14 months, $1.15B follow-on Sep-25, NVIDIA $2.0B pre-funded warrant Mar-26, $2.81B net ATM Q2-26", fmt="0")
S("FP_ATMMode", "ATM sizing: 1 = remaining 12,270,507 sh; 2 = $ amount FP_ATMAmt (needs new program)", "#", 1, "Modeling", "Estimated", "Bear: converts shut -> ATM sized to fill YE27 gap", bull=1, bear=2, fmt="0")
S("FP_ATMIdx", "ATM sale quarter", "idx", 12, "Modeling", "Estimated", "Q4'26", fmt="0")
S("FP_ATMDisc", "ATM price discount to SharePx", "%", 0.0, "Modeling", "Estimated", "Base/bull at $242.81 (~$2.98B); bear 40% below (test 30% / 50%)", bull=0.0, bear=0.40, fmt="0%")
S("FP_ATMAmt", "ATM $ amount if FP_ATMMode = 2", "$M", 0, "Modeling", "Estimated", "Bear: sized so YE27 liquidity >= MinCash", bull=0, bear=9000, fmt="#,##0")
S("FP_ATMPx", "ATM issue price", "$/sh", "=SharePx*(1-FP_ATMDisc)", "Modeling", "Derived", "", fmt="0.00")
S("FP_ATMSh", "ATM shares issued (M)", "M sh", "=FP_On*(1-StopBuild)*IF(OR(FP_ATMMode=1,FP_Resize=1),ATM_Rem,FP_ATMAmt/FP_ATMPx)", "Modeling", "Derived", "", fmt="0.000")
S("FP_ATMProc", "ATM proceeds ($M)", "$M", "=FP_ATMSh*FP_ATMPx", "Modeling", "Derived", "Fees ignored", fmt="#,##0")
S("FP_CVOpen", "Convert market open (1) / shut (0)", "0/1", 1, "Modeling", "Estimated", "Bear: shut", bull=1, bear=0, fmt="0")
S("FP_CV1Idx", "Convert deal 1 quarter", "idx", 13, "Modeling", "Estimated", "Q1'27 (ahead of first negative quarter Q2'27)", fmt="0")
S("FP_CV2Idx", "Convert deal 2 quarter", "idx", 15, "Modeling", "Estimated", "Q3'27", fmt="0")
S("FP_CV1Amt", "Convert deal 1 gross ($M)", "$M", 6400, "Modeling", "Estimated", "Sized so base YE27 liquidity ~ MinCash $2B: pre-plan YE27 gap -$13.5B in v0.7 (loan $0, Rubin capex $41M) vs ~-$9.5B in v0.6; $5B each leaves -$0.7B. Aug-26 deal $5.75B", fmt="#,##0")
S("FP_CV2Amt", "Convert deal 2 gross ($M)", "$M", 6400, "Modeling", "Estimated", "", fmt="#,##0")
S("ABS_CVOff", "v0.8b: ABS draws displace the 2027 plan converts first (1 = deal 1 / deal 2 cut by cumulative ABS drawn by their quarter)", "0/1", 1, "Modeling", "Estimated", "So ABS reduces convert / equity need instead of sitting as idle cash", fmt="0")
S("FP_CV1Net", "Convert deal 1 gross after ABS displacement ($M)", "$M", "=MAX(0,FP_CV1Amt-ABS_CVOff*SUMPRODUCT((Q_Idx<=FP_CV1Idx)*ABS_Draw))", "Modeling", "Derived", "", fmt="#,##0")
S("FP_CV2Net", "Convert deal 2 gross after ABS displacement ($M)", "$M", "=MAX(0,FP_CV2Amt-MAX(0,ABS_CVOff*SUMPRODUCT((Q_Idx<=FP_CV2Idx)*ABS_Draw)-(FP_CV1Amt-FP_CV1Net)))", "Modeling", "Derived", "", fmt="#,##0")
S("FP_ShortSh", "Share of each deal in the short tranche", "%", 0.6, "Filings", "Derived", "Aug-26: $3.45B 2030s of $5.75B", fmt="0%")
S("FP_CpnS", "Short tranche coupon", "%", 0.005, "Filings", "Reported", "Aug-26 0.50% Feb-2030 (Mar-26: 1.25% Mar-2031)", fmt="0.00%")
S("FP_CpnL", "Long tranche coupon", "%", 0.045, "Filings", "Reported", "Aug-26 4.50% Feb-2034 (Mar-26: 2.625% Mar-2033)", fmt="0.00%")
S("FP_AccS", "Short tranche accretion at maturity (x principal)", "x", 1.10, "Filings", "Reported", "Aug-26 2030s 110%", fmt="0.00")
S("FP_AccL", "Long tranche accretion at maturity", "x", 1.25, "Filings", "Reported", "Aug-26 2034s 125%", fmt="0.00")
S("FP_PremS", "Short tranche conversion premium over reference price", "%", 0.40, "Filings", "Derived", "Aug-26: $313.46 / $223.90 ref close = 1.40", fmt="0%")
S("FP_PremL", "Long tranche conversion premium", "%", 0.45, "Filings", "Derived", "Aug-26: $324.65 / $223.90 = 1.45", fmt="0%")
S("FP_TenS", "Short tranche tenor (years)", "yrs", 3.5, "Filings", "Derived", "Aug-26 2030s: 8/24/26 -> 2/15/30", fmt="0.0")
S("FP_TenL", "Long tranche tenor (years)", "yrs", 7.5, "Filings", "Derived", "Aug-26 2034s: 8/24/26 -> 2/15/34", fmt="0.0")
S("FP_RefPx", "Reference share price at issue", "$/sh", "=SharePx", "Modeling", "Estimated", "Market price (deals priced off market, not model value)", fmt="0.00")
S("FP_Fee", "Convert fees (% of gross)", "%", "=1-AugNet/AugGross", "Filings", "Derived", "Aug-26: 5,680 net / 5,750 gross", fmt="0.0%")
S("FP_PostQ", "Post-funding share count quarter", "idx", "=IF(FP_Resize=1,FP_TopEnd,16)", "Modeling", "Derived", "v0.7b YE27 (16); re-sized plan: FP_TopEnd", fmt="0")
S("AVBurnQ", "Avride adj. EBITDA per quarter (held flat)", "$M", -37.1, "Filings", "Reported", "H1'26 -$74.2M / 2; inside consolidated EBITDA", fmt="#,##0.0")
S("AVBurnOn", "Add back PV of Avride burn when Avride is valued separately (V_AV > 0) - avoids double count", "0/1", 1, "Modeling", "Input", "v0.7 fix; 0 = v0.6 (bull double count)", fmt="0")

sec("Valuation")
S("WACC", "WACC", "%", 0.11, "Modeling", "Estimated", "", bull=0.10, bear=0.13, fmt="0.0%")
S("Ke", "Cost of equity (12-month PT roll-forward)", "%", 0.13, "Modeling", "Estimated", "", bull=0.12, bear=0.15, fmt="0.0%")
S("TG", "Terminal growth (TGMode 0 only)", "%", 0.03, "Modeling", "Estimated", "", bull=0.035, bear=0.02, fmt="0.0%")
S("TGMode", "Terminal growth basis: 1 = fundamentals g = TermRR x TermRONC with matching reinvestment (TermRR x NOPAT... = g/RONC x NOPAT) deducted [v0.4 default]; 0 = fixed TG with TG x net PP&E growth capex", "0/1", 1, "Skeptic", "Input", "Skeptic: terminal g must be consistent with reinvestment", fmt="0")
S("TermRONC", "Terminal return on NEW capital (after tax)", "%", 0.12, "Skeptic", "Estimated", "Placeholder ~WACC + 1pp (new capital earns little excess return in a competitive GPU market); pending Skeptic", bull=0.14, bear=0.10, fmt="0.0%")
S("TermRR", "Terminal reinvestment rate (net reinvestment / NOPAT)", "%", 0.25, "Skeptic", "Estimated", "Placeholder; g = 25% x 12% = 3.0% base; pending Skeptic", bull=0.25, bear=0.20, fmt="0.0%")
S("TGCap", "Cap on fundamental terminal g (nominal GDP-type ceiling)", "%", 0.04, "Modeling", "Estimated", "", fmt="0.0%")
S("ExitMult", "Exit / horizon EV / NTM adj. EBITDA", "x", 14.0, "GPU Market", "Placeholder", "Peer comps pending (CRWV ~5.7x FY27 per Skeptic)", bull=18.0, bear=9.0, fmt="0.0x")
S("DCFWeight", "Weight on Gordon-growth TV (rest = exit multiple)", "%", 1.00, "Modeling", "Input", "Was 50%: the 14x exit multiple is an unsourced placeholder (CRWV ~4.8x FY27); Gordon is base, exit multiple shown as cross-check", fmt="0%")
S("MaintRatio", "Terminal capex / D&A", "x", 1.00, "Modeling", "Estimated", "")
S("PeerEVMW", "Peer EV per contracted MW", "$M/MW", 0.0, "GPU Market", "Placeholder", "Cross-check only")
S("CH_Lo", "ClickHouse stake value - low (carrying value)", "$M", 1517.7, "Filings", "Reported", "Q2'26 6-K Ex.99.2 Note 5 carrying value (Filings SOTP tab)", fmt="#,##0.0")
S("CH_Hi", "ClickHouse stake value - mark (~28% fully diluted x ~$15B Series D)", "$M", 4200, "Filings", "Estimated", "Filings SOTP tab", fmt="#,##0")
S("TO_Lo", "Toloka stake value - low (fair value)", "$M", 130.6, "Filings", "Reported", "Fair value $130.6M (carrying $86.5M); Note 5", fmt="#,##0.0")
S("TO_Hi", "Toloka stake value - high (fair value)", "$M", 130.6, "Filings", "Reported", "Fair value $130.6M", fmt="#,##0.0")
S("AV_Lo", "Avride value - low", "$M", 0, "Filings", "Reported", "Not disclosed (consolidated sub) -> floor $0", fmt="#,##0")
S("AV_Hi", "Avride value - high (Caplight third-party estimate, LOW confidence)", "$M", 2750, "Filings", "Estimated", "Avride H1'26 adj. EBITDA -$74.2M (Reported) is inside consolidated opex calibration -> adding a stand-alone value double-counts its burn; high case only", fmt="#,##0")
S("AVPt", "Avride point in low-high range (0 = $0 floor)", "x", 0, "Modeling", "Input", "v0.7: all scenarios 0 = keep Avride burn in consolidated EBITDA, credit $0 (v0.6 bull 0.5 credited value AND kept burn = double count). Alternative: AVPt 0.5 + AVBurnOn 1 (credit value, add back PV of burn) - sensitivity", fmt="0.0")
S("Other_Val", "VC fund + other non-marketable investments (TripleTen not disclosed - blank)", "$M", 16.3, "Filings", "Reported", "VC fund $13.8M + other non-marketable $2.5M (Note 5)", fmt="#,##0.0")
S("SOTPPt", "SOTP point in low-high range (0 = low, 1 = high)", "x", 0.5, "Modeling", "Input", "Bull = high, base = mid, bear = low", bull=1.0, bear=0.0, fmt="0.0")
S("OffBS_Energy", "Energy purchase commitments (off-balance-sheet)", "$M", 5320, "Filings", "Reported", "Q2'26 interim FS commitments note (Filings & Capital)", fmt="#,##0")
S("OffBS_Lease", "Signed leases not yet commenced (off-balance-sheet)", "$M", 12050, "Filings", "Reported", "Q2'26 interim FS (Filings & Capital)", fmt="#,##0")
S("OpLease", "Operating lease liabilities on balance sheet", "$M", 1625, "Filings", "Reported", "Q2'26 balance sheet (Filings & Capital); in lease toggle block (forecast rent already in EBITDA -> partial weight)", fmt="#,##0")
S("OffBSWt", "Weight of off-B/S commitments in EV bridge", "%", 0.0, "Modeling", "Input", "v0.7d 0% (was 50%): the $19.0B energy + lease commitments are the contractual form of the rent (RentMW, CLT1Rent) and power (VinPower, PowerPx) already deducted in EBITDA over the DCF horizon and terminal -> deducting them again double counts. Not in capex (leased capex excludes the building). Valuation still shows 0% / chosen / 100%", fmt="0%")
S("SkEBIT4", "Skeptic FY27 EBIT at 4-yr GPU life", "$M", 820, "Skeptic", "Estimated", "Skeptic sensitivity figures for cross-check", fmt="#,##0")
S("SkEBIT5", "Skeptic FY27 EBIT at 5-yr GPU life", "$M", 1980, "Skeptic", "Estimated", "", fmt="#,##0")
S("SkEBIT6", "Skeptic FY27 EBIT at 6-yr GPU life", "$M", 2750, "Skeptic", "Estimated", "", fmt="#,##0")
S("SOTP_Disc", "Holdco discount on stakes", "%", 0.25, "Modeling", "Estimated", "", bull=0.15, bear=0.40, fmt="0%")

# ---- annual inputs
r += 1
ws[f"A{r}"] = "Annual time-series inputs (looked up by fiscal year)"; ws[f"A{r}"].font = BOLD
for c in range(1, 13): ws.cell(r, c).fill = SECFILL
r += 1
for i, y in enumerate(YEARS):
    c = CL(3 + i); ws[f"{c}{r}"] = y; ws[f"{c}{r}"].font = HDR; ws[f"{c}{r}"].fill = HDRFILL
for c, t in [("A", "Annual input"), ("B", "Unit"), ("J", "Owner"), ("K", "Status"), ("L", "Source / method")]:
    ws[f"{c}{r}"] = t; ws[f"{c}{r}"].font = HDR; ws[f"{c}{r}"].fill = HDRFILL
r += 1
def A(nm, label, unit, vals, owner, status, src, fmt="0%"):
    global r
    ws[f"A{r}"] = label; ws[f"B{r}"] = unit
    for i, v in enumerate(vals):
        inp(ws[f"{CL(3+i)}{r}"], v, status, fmt)
    ws[f"J{r}"] = owner; ws[f"K{r}"] = status; ws[f"L{r}"] = src
    for c in "BJKL": ws[f"{c}{r}"].font = ITAL
    name(wb, nm, "Inputs", f"$C${r}:$I${r}")
    r += 1
A("OwnedSh", "Owned share of unallocated (portfolio-level) MW", "%", [0.3, 0.3, 0.3, 0.5, 0.65, 0.7, 0.75], "Capacity", "Estimated", "Q1'26 letter: >75% of contracted power owned; energized mix shifts to owned as MO/PA/AL ramp 2027+")
A("Mix_H", "New-deployment generation mix - Hopper", "%", [1.0, 0.6, 0.1, 0.0, 0.0, 0.0, 0.0], "GPU Market", "Placeholder", "")
A("Mix_B", "New-deployment generation mix - Blackwell", "%", [0.0, 0.4, 0.9, 0.6, 0.3, 0.1, 0.0], "GPU Market", "Placeholder", "")
A("Mix_N", "New-deployment generation mix - Next-gen (Rubin)", "%", [0.0, 0.0, 0.0, 0.4, 0.7, 0.9, 1.0], "GPU Market", "Placeholder", "Meta #2 = Vera Rubin from early 2027")
A("ResSh", "Book (b) reserved/commit share of non-hyperscaler GPU capacity", "%", [0.85, 0.88, 0.90, 0.90, 0.90, 0.90, 0.90], "GPU Market", "Estimated", "Set so Q2'26 b:c revenue ~50:4 (GPU Market mix) at realization 0.72 / 0.95")
A("SpotSh", "Spot share within book (c) on-demand capacity", "%", [0.1] * 7, "GPU Market", "Placeholder", "")
A("StudioUp", "AI Studio / Token Factory uplift on book (c)", "%", [0.05, 0.05, 0.10, 0.15, 0.20, 0.20, 0.20], "GPU Market", "Placeholder", "Token Factory production inference >3x QoQ in Q2'26 (no $ disclosed)")
A("AssetLight", "Asset-light partner revenue", "$M/yr", [0, 0, 0, 0, 0, 0, 0], "GPU Market", "Placeholder", "Model launched Q2'26; no economics disclosed", fmt="#,##0")
A("Cons_Rev", "Consensus revenue (FY24..FY30; FY26-28 only)", "$M", [None, None, 3340, 12300, 23000, None, None], "Skeptic", "Estimated", "Consensus, /workspace/nebius/skeptic/01_priced_in.xlsx, Oct 2 2026, medium confidence", fmt="#,##0")
A("Cons_EB", "Consensus adj. EBITDA", "$M", [None, None, 1390, 6730, 13980, None, None], "Skeptic", "Estimated", "as above", fmt="#,##0")
A("Cons_Cx", "Consensus capex", "$M", [None, None, None, 29100, None, None, None], "Skeptic", "Estimated", "as above (FY27 only)", fmt="#,##0")
A("PDpct", "Product development ex-SBC", "% rev", [0.11, 0.11, 0.11, 0.09, 0.08, 0.07, 0.06], "Modeling", "Placeholder", "Q2'26 PD $191.0M incl. $115.9M one-off Eigen AI SBC/post-combination -> ~11-13% underlying")
A("SGApct", "SG&A ex-SBC", "% rev", [0.22, 0.22, 0.22, 0.14, 0.10, 0.08, 0.07], "Modeling", "Placeholder", "Calibrated so Q2'26 group adj. EBITDA margin ~41%")
r += 1
ws[f"A{r}"] = "Residual (post-term) price curve: % of launch on-demand price by GPU age"; ws[f"A{r}"].font = BOLD
r += 1
for i in range(8):
    c = CL(3 + i); ws[f"{c}{r}"] = f"Yr {i+1}"; ws[f"{c}{r}"].font = HDR; ws[f"{c}{r}"].fill = HDRFILL
r += 1
ws[f"A{r}"] = "H100 SXM on-demand neocloud median, annual avg ($/GPU-hr): 2024 | 2025 | 2026 YTD"
for i, v in enumerate([3.5275, 3.1796, 3.717]): inp(ws[f"{CL(3+i)}{r}"], v, "Derived", "0.00")
ws[f"J{r}"] = "GPU Market"; ws[f"K{r}"] = "Derived"; ws[f"L{r}"] = "Loaded from /workspace/nebius/market/gpu_price_history.xlsx sheet H100_Nebius_vs_Medians (monthly fwd-filled medians, Jan-24..Oct-26); re-run build_model.py to refresh"
hm = r
r += 1
ws[f"A{r}"] = "Base curve: yrs 1-3 = H100 median relative to 2024 (Derived); yrs 4-8 placeholder pending longer history"
fml(ws[f"C{r}"], "=1", "0%"); fml(ws[f"D{r}"], f"=D{hm}/C{hm}", "0%"); fml(ws[f"E{r}"], f"=E{hm}/C{hm}", "0%")
for i, v in enumerate([0.85, 0.70, 0.55, 0.45, 0.35]): inp(ws[f"{CL(6+i)}{r}"], v, "Placeholder", "0%")
ws[f"J{r}"] = "GPU Market"; ws[f"K{r}"] = "Placeholder"; name(wb, "Curve_Base", "Inputs", f"$C${r}:$J${r}".replace("$J", "$J")); cb = r
r += 1
ws[f"A{r}"] = "Bear curve (A100 proxy): linear to (1 - ResidBearCut) at age 5, then same pace compounding; never zero; bear ResidUtil 60%"
for i in range(8):
    k = i + 1
    f = f"=1-ResidBearCut*{k-1}/4" if k <= 5 else f"=(1-ResidBearCut)*(1-ResidBearCut/4)^{k-5}"
    fml(ws[f"{CL(3+i)}{r}"], f, "0%"); ws[f"{CL(3+i)}{r}"].fill = F_EST
ws[f"J{r}"] = "Skeptic"; ws[f"K{r}"] = "Placeholder"; cbe = r
r += 1
ws[f"A{r}"] = "LIVE curve"
for i in range(8):
    c = CL(3 + i); fml(ws[f"{c}{r}"], f"=IF(Scen=3,{c}{cbe},{c}{cb})", "0%", True)
name(wb, "Curve", "Inputs", f"$C${r}:$J${r}")
# fix Curve_Base name range to C..J
name(wb, "Curve_Base", "Inputs", f"$C${cb}:$J${cb}")
ws.freeze_panes = "B7"

# ------------------------------------------------------------------ TIMELINE
tl = wb.create_sheet("Timeline")
tl["A1"] = "Timeline (one column per quarter, Q1'24-Q4'30)"; tl["A1"].font = TITLE
for rr, lab in [(3, "Quarter"), (4, "Period end"), (5, "Index"), (6, "A/F"), (7, "Fiscal year"), (8, "Days"), (9, "Hours")]:
    tl[f"A{rr}"] = lab; tl[f"A{rr}"].font = BOLD
tl["E4"] = dt.date(2023, 12, 31); tl["E4"].number_format = "yyyy-mm-dd"; tl["E5"] = 0
for n in range(1, NQ + 1):
    c = qc(n); p = qc(n - 1)
    tl[f"{c}3"] = QL[n - 1]
    tl[f"{c}4"] = QEND[n - 1]; tl[f"{c}4"].number_format = "yyyy-mm-dd"
    tl[f"{c}5"] = n
    tl[f"{c}6"] = f'=IF({c}5<=LastHist,"A","F")'
    tl[f"{c}7"] = f"=YEAR({c}4)"
    tl[f"{c}8"] = f"={c}4-{p}4"
    tl[f"{c}9"] = f"={c}8*24"
    tl.column_dimensions[c].width = 11
tl.column_dimensions["A"].width = 16
name(wb, "Q_End", "Timeline", "$F$4:$AG$4"); name(wb, "Q_Idx", "Timeline", "$F$5:$AG$5")
name(wb, "Q_Lab", "Timeline", "$F$3:$AG$3"); name(wb, "Q_Yr", "Timeline", "$F$7:$AG$7"); name(wb, "Q_Hrs", "Timeline", "$F$9:$AG$9")
def HRS(n): return f"Timeline!{qc(n)}$9"
def YR(n): return f"{qc(n)}$7"
def ANN(nm, n): return f"INDEX({nm},1,{YR(n)}-FYBase)"

# ------------------------------------------------------------------ HIST
hs = wb.create_sheet("Hist")
hs["A1"] = "Historical inputs Q1'24-Q2'26 with Reported / Derived / Estimated / Not disclosed tags"; hs["A1"].font = TITLE
hs["A2"] = ("Each metric = 5 rows: value as disclosed | tag | best estimate (numeric) | method/source | VALUE USED (= value if disclosed, else best estimate). "
            "No silent gaps: every quarter carries a tag. Other tabs link to the VALUE USED row only."); hs["A2"].font = ITAL
for c, t in zip("ABCDE", ["Metric", "Unit", "Owner", "Notes", ""]):
    hs[f"{c}3"] = t; hs[f"{c}3"].font = HDR; hs[f"{c}3"].fill = HDRFILL
for n in range(1, NHIST + 1):
    c = qc(n); hs[f"{c}3"] = QL[n - 1]; hs[f"{c}3"].font = HDR; hs[f"{c}3"].fill = HDRFILL
    hs[f"{c}4"] = QEND[n - 1]; hs[f"{c}4"].number_format = "mmm-yy"
    hs.column_dimensions[c].width = 13
setw(hs, {"A": 34, "B": 8, "C": 12, "D": 50, "E": 3})
tagdv = DataValidation(type="list", formula1='"Reported,Derived,Estimated,Not disclosed"', allow_blank=False)
hs.add_data_validation(tagdv)
hr = 6
ND = None
def H(nm, label, unit, owner, notes, vals, tags, est, method, fmt="#,##0.0;(#,##0.0);-"):
    """vals: list of 10 (None = not disclosed); tags list of 10; est list of 10 (best estimate numeric or None); method list/str"""
    global hr
    if isinstance(method, str): method = [method] * NHIST
    hs[f"A{hr}"] = label; hs[f"A{hr}"].font = BOLD; hs[f"B{hr}"] = unit; hs[f"C{hr}"] = owner; hs[f"D{hr}"] = notes
    hs[f"D{hr}"].font = ITAL; hs[f"D{hr}"].alignment = Alignment(wrap_text=True)
    labs = ["value as disclosed", "tag", "best estimate", "method / source", "VALUE USED"]
    for k, lab in enumerate(labs):
        hs[f"A{hr+k}"] = ("   " + lab) if k else label
        if k: hs[f"A{hr+k}"].font = ITAL
    hs[f"A{hr+4}"].font = BOLD
    for i in range(NHIST):
        c = qc(i + 1)
        if vals[i] is not None: inp(hs[f"{c}{hr}"], vals[i], tags[i], fmt)
        else: hs[f"{c}{hr}"] = "n.d."; hs[f"{c}{hr}"].font = ITAL
        inp(hs[f"{c}{hr+1}"], tags[i], tags[i]); tagdv.add(f"{c}{hr+1}")
        if est[i] is not None: inp(hs[f"{c}{hr+2}"], est[i], "Estimated", fmt)
        hs[f"{c}{hr+3}"] = method[i]; hs[f"{c}{hr+3}"].font = ITAL; hs[f"{c}{hr+3}"].alignment = Alignment(wrap_text=False)
        fml(hs[f"{c}{hr+4}"], f'=IF({c}{hr+1}="Not disclosed",{c}{hr+2},{c}{hr})', fmt, True)
    # register used row (F..O); pad to AG not needed
    name(wb, nm, "Hist", f"$F${hr+4}:$O${hr+4}")
    RN[nm] = ("Hist", hr + 4)
    hr += 6
R_, D_, E_, N_ = "Reported", "Derived", "Estimated", "Not disclosed"
SRC_Q = {1: "Q1'25 6-K PR comparatives (pre-Toloka-deconsolidation basis)", 2: "Q2'25 6-K PR comparative", 3: "Q3'25 6-K PR comparative (restated)",
         4: "Q4'25 6-K PR comparative", 5: "Q1'26 6-K PR comparative (restated)", 6: "Q2'26 6-K PR comparative", 7: "Q3'25 6-K PR",
         8: "Q4'25 6-K PR", 9: "Q1'26 6-K PR", 10: "Q2'26 6-K PR (acc. 0001104659-26-094568)"}
srcs = [SRC_Q[i] for i in range(1, 11)]
H("H_Rev", "Group revenue", "$M", "Filings", "Continuing ops. Q1'24 figure is on the pre-restatement basis (Q1'25 PR) - flag.",
  [11.4, 14.5, 32.1, 35.2, 50.9, 105.1, 146.1, 227.7, 399.0, 582.3], [R_] * 10, [None] * 10, srcs)
aic_est = [round(v * 0.9, 1) for v in [11.4, 14.5, 32.1, 35.2]] + [None, None, round(146.1 * 0.95, 1), None, None, None]
H("H_AIRev", "Nebius AI cloud revenue", "$M", "Filings", "Segment disclosure; earlier quarters derived from YoY growth or estimated as % of group",
  [None, None, None, None, 41.4, 93.6, None, 214.2, 389.7, 574.9], [N_, N_, N_, N_, D_, D_, N_, R_, R_, R_],
  aic_est[:4] + [None, None, aic_est[6], None, None, None],
  ["Est = 90% of group revenue (pre-2025 mix; Avride/TripleTen remainder)"] * 4 +
  ["Derived: Q1'26 $389.7M / (1+841%)", "Derived: Q2'26 $574.9M / (1+514%)", "Est = 95% of group revenue", "Q4'25 letter", "Q1'26 letter", "Q2'26 letter"])
H("H_ARR", "ARR (last-month AI cloud revenue x 12)", "$M", "Filings", "Mgmt definition; Q2'25/Q1'25 derived from stated YoY growth",
  [None, None, None, 90, 248, 430, 551, 1250, 1920, 3000], [N_, N_, N_, R_, D_, D_, R_, R_, R_, R_],
  [None, None, None, None, None, None, None, None, None, None],
  ["Not disclosed (no est.)"] * 3 + ["Q4'25 letter ARR chart (Dec-24 $90M)", "Derived: $1.92B / (1+674%)", "Derived: $3.0B / (1+598%)",
   "Q4'25 letter: 'more than doubled from $551M'", "Q4'25 letter", "Q1'26 letter ($1.92B)", "Q2'26 letter"])
# ARR n.d. rows need numeric best estimate to avoid silent gaps
for i, c in enumerate(["F", "G", "H"]):
    inp(hs[f"{c}{RN['H_ARR'][1]-2}"], [11.4, 14.5, 32.1][i] * 0.9 * 4, "Estimated", "#,##0.0")
    hs[f"{c}{RN['H_ARR'][1]-1}"] = "Est = AI cloud revenue est x 4 (run-rate proxy)"
opx_note = "GAAP incl. SBC. Q1'24/Q2'24 estimated by splitting (revenue - adj. EBITDA) on Q3'24 mix"
cash_opex = [11.4 + 70.9, 14.5 + 58.1]
mix = [9.9, 31.4, 47.9]; tot = sum(mix)
def splitest(k): return [round(cash_opex[0] * mix[k] / tot, 1), round(cash_opex[1] * mix[k] / tot, 1)]
H("H_COR", "Cost of revenues", "$M", "Filings", opx_note, [None, None, 9.9, 21.1, 24.7, 30.1, 42.9, 68.5, 103.8, 133.6],
  [N_, N_] + [R_] * 8, splitest(0) + [None] * 8, ["Est: (rev - adj EBITDA) x Q3'24 COR share"] * 2 + srcs[2:])
H("H_PD", "Product development", "$M", "Filings", opx_note + "; Q2'26 incl. $115.9M one-off Eigen AI", [None, None, 31.4, 31.9, 36.5, 42.8, 44.9, 53.1, 67.4, 191.0],
  [N_, N_] + [R_] * 8, splitest(1) + [None] * 8, ["Est: (rev - adj EBITDA) x Q3'24 PD share"] * 2 + srcs[2:])
H("H_SGA", "Sales, general & administrative", "$M", "Filings", opx_note, [None, None, 47.9, 85.4, 60.9, 68.2, 89.5, 159.9, 143.8, 173.9],
  [N_, N_] + [R_] * 8, splitest(2) + [None] * 8, ["Est: (rev - adj EBITDA) x Q3'24 SG&A share"] * 2 + srcs[2:])
H("H_DA", "Depreciation & amortization", "$M", "Filings", "", [8.9, 11.4, 23.5, 33.3, 49.1, 75.2, 99.0, 180.7, 212.0, 259.7], [R_] * 10, [None] * 10, srcs)
H("H_DAex", "D&A excluding 4->5 yr life change", "$M", "Filings", "Q1'26 benefit $43.1M, Q2'26 $43.0M (FY26 guided $167.6M); = reported + benefit",
  [8.9, 11.4, 23.5, 33.3, 49.1, 75.2, 99.0, 180.7, 255.1, 302.7], [R_] * 8 + [D_, D_], [None] * 10, srcs[:8] + ["Derived: 212.0 + 43.1 (Q1'26 6-K Note 1)", "Derived: 259.7 + 43.0 (Q2'26 6-K Note 1)"])
H("H_PPE", "PP&E, net (end of period)", "$M", "Filings", "Q1'24 not comparable (pre-divestiture); /workspace/nebius/filings/nebius_filings.xlsx Quarterly",
  [None, 328.3, 475.3, 846.7, 1334.1, 1789.4, 3314.4, 5553.3, 7131.7, 13045.2], [N_] + [R_] * 9, [250.0] + [None] * 9, ["Est: Q2'24 less Q2 capex net of D&A"] + srcs[1:])
H("H_CIP", "Assets not yet in use (CIP / not placed in service)", "$M", "Filings", "Note 7 of Q1/Q2'26 statements (Filings & Capital); pre-Q4'25 not disclosed",
  [None] * 7 + [2417.4, 2427.9, 7836.4], [N_] * 7 + [R_] * 3, [None] * 10, ["Not disclosed (no best estimate; yield check starts Q1'26)"] * 7 + ["Q1/Q2'26 6-K Note 7"] * 3)
H("H_ISG", "In-service gross PP&E", "$M", "Filings", "Derived = gross PP&E - assets not yet in use (Note 7)",
  [None] * 7 + [3772.9, 5548.7, 6313.1], [N_] * 7 + [D_] * 3, [None] * 10, ["Not disclosed"] * 7 + ["Derived, Note 7"] * 3)
H("H_SRV", "  of which servers & network equipment (gross, in service)", "$M", "Filings", "Note 7",
  [None] * 7 + [3123.6, 4692.7, 5197.7], [N_] * 7 + [R_] * 3, [None] * 10, ["Not disclosed"] * 7 + ["Note 7"] * 3)
H("H_PPDep", "PP&E depreciation (excl. intangible amortization)", "$M", "Filings", "Note 7; pre-Q1'26 best estimate = total D&A",
  [None] * 8 + [208.8, 253.8], [N_] * 8 + [R_] * 2, [8.9, 11.4, 23.5, 33.3, 49.1, 75.2, 99.0, 180.7, None, None], ["Est = total D&A (amortization small)"] * 8 + ["Note 7"] * 2)
H("H_CapInt", "Capitalized interest", "$M", "Filings", "Note 12: Q2'26 $40.2M vs $107.5M gross convert interest",
  [None] * 8 + [15.7, 40.2], [N_] * 8 + [R_] * 2, [0.0] * 8 + [None, None], ["Not disclosed; est 0"] * 8 + ["Note 12"] * 2)
H("H_EBITDA", "Adjusted EBITDA (group)", "$M", "Filings", "Q1'24 on pre-restatement basis", [-70.9, -58.1, -45.9, -63.9, -53.7, -21.0, -5.2, 15.0, 129.5, 236.2], [R_] * 10, [None] * 10, srcs)
H("H_Capex", "Capex (purchases of PP&E + intangibles, cash)", "$M", "Filings", "FY24 cash capex $807.5M; FY25 $4,066.0M (20-F FY25 cash-flow)",
  [None, None, None, 417.6, 544.0, 510.6, 955.5, 2056.0, 2472.9, 5657.4], [N_, N_, N_, R_, R_, R_, R_, R_, D_, R_],
  [130.0, 130.0, 129.9] + [None] * 7,
  ["Est: (FY24 807.5 - Q4'24 417.6) / 3 - split not disclosed"] * 3 + ["Q4'24 PR", "Q1'25 letter", "Q2'25 letter", "Q3'25 letter", "20-F FY25 cash flow (FY 4,066.0 less 9M)",
   "Derived: 1H26 8,130.3 - Q2 5,657.4", "Q2'26 PR (Filings)"])
H("H_Cash", "Cash & equivalents (ex-restricted)", "$M", "Filings", "", [None, None, None, 2434.7, 1447.0, 1679.3, 4794.8, 3678.1, 9298.2, 8042.1],
  [N_, N_, N_, R_, R_, R_, R_, R_, R_, R_], [None, None, 2288.2] + [None] * 7,
  ["Not loaded - Filings to add from Q1/Q2'24 6-Ks (best est. n/a)", "Not loaded - Filings to add", "Q3'24 PR (pre-restatement) as best estimate"] + srcs[3:])
for c in ["F", "G"]:
    inp(hs[f"{c}{RN['H_Cash'][1]-2}"], 2288.2, "Estimated", "#,##0.0"); hs[f"{c}{RN['H_Cash'][1]-1}"] = "Est: carry Q3'24 balance (not loaded)"
H("H_Shares", "Shares issued & outstanding (Class A + B, ex-treasury)", "M sh", "Filings", "",
  [None, None, 199.340118, 235.7536, 238.108831, 238.705092, 251.807222, 253.016971, 253.898194, 271.855218],
  [N_, N_] + [R_] * 8, [199.34, 199.34] + [None] * 8, ["Est: carry Q3'24 (pre-restructuring counts not comparable)"] * 2 + srcs[2:], fmt="#,##0.000")
H("H_Active", "Active power (IT MW consumed)", "MW", "Capacity", "Only YE25 disclosed (~170 MW). Est = 170 x group revenue / Q4'25 revenue (pre-2026) or linear path to 900 MW by Q2'27 (2026)",
  [None] * 7 + [170] + [None, None], [N_] * 7 + [R_] + [N_, N_],
  [round(170 * v / 227.7, 1) for v in [11.4, 14.5, 32.1, 35.2, 50.9, 105.1, 146.1]] + [None, 291.7, 413.3],
  ["Est: revenue-proportional to YE25 170 MW (YE24 ~25 MW connected per Q3'25 letter = sanity check)"] * 7 + ["Q4'25 letter: ~170 MW active"] +
  ["Est: linear 170 MW (Q4'25) -> 900 MW (Q2'27; 800MW-1GW connected YE26, active through 1H27 per Q2'26 call)"] * 2)
H("H_Contracted", "Contracted power (portfolio)", "MW", "Capacity/Filings", "Contracted = land + power commitments, not built",
  [None] * 7 + [2000, None, 3500], [N_] * 7 + [R_, N_, R_], [None] * 8 + [2750, None],
  ["n.d. pre-Feb-26 (best est. n/a - not used for revenue)"] * 7 + ["Q4'25 letter (Feb-26): >2 GW", "Est: midpoint Feb->May", "Q1'26 letter (May-26): >3.5 GW contracted"])
for k in range(7):
    c = qc(k + 1); inp(hs[f"{c}{RN['H_Contracted'][1]-2}"], 0, "Estimated", "#,##0")
H("H_DR", "Deferred revenue (current + non-current)", "$M", "Filings", "Customer prepayments; Filings & Capital",
  [None] * 7 + [1578, 4778, 5975], [N_] * 7 + [R_] * 3, [0] * 7 + [None] * 3, ["n.d. in collected files; best est. 0 (pre-prepayment era) - Filings to verify"] * 7 + ["Q4'25 20-F", "Q1'26 6-K FS", "Q2'26 6-K FS"], fmt="#,##0")
H("H_RPO", "Remaining performance obligations", "$M", "Filings", "100% AI cloud at 6/30/26 (36% <24 months)",
  [None] * 7 + [21300, 33600, 37490.6], [N_] * 7 + [R_] * 3, [None] * 10, ["n.d. in collected files (not used in model)"] * 7 + ["20-F FY25", "Q1'26 6-K FS", "Q2'26 6-K FS Note"], fmt="#,##0")
hs.freeze_panes = "F5"

# ------------------------------------------------------------------ SITES
st = TS(wb, "Sites", "Site-by-site MW: three blocks (contracted / energized estimate [Capacity] / mgmt-guided). Selector Inputs!MWSel picks the revenue driver. All MW converted to IT basis via PUE where basis = Facility.")
sws = st.ws
import openpyxl as _ox
CAPF1 = '/workspace/nebius/capacity/nebius_capacity_v1.xlsx'; CAPF2 = '/workspace/nebius/capacity/nebius_capacity.xlsx'
def _loadcap(f):
    d = {}
    for _r in _ox.load_workbook(f, data_only=True)['Quarterly MW'].iter_rows(min_row=3, values_only=True):
        if not _r[0] or not _r[1]: continue
        d.setdefault(_r[0], {})[_r[1]] = (_r[2], list(_r[4:32]))
    return d
CAP = _loadcap(CAPF1); CAP2 = _loadcap(CAPF2)
BAS2 = {}
for _r in _ox.load_workbook(CAPF2, data_only=True)['Basis'].iter_rows(min_row=2, values_only=True):
    if _r[0] and _r[5] in ("IT", "gross") and isinstance(_r[6], (int, float)): BAS2[_r[0]] = (1.0 / _r[6]) if _r[5] == "gross" else 1.0
SKEYS = ["Vineland", "Independence", "Butler", "Birmingham", "Bethune", "Lappeenranta", "Mantsala FI #1", "Mantsala FI #2", "Kansas City", "Paris", "Keflavik", "Longcross",
         "Harlow", "Newport", "UK 3rd", "Modi'in", "Masmiyya", "Beit Shemesh", "Madrid", "Tallinn", "Minnesota", "Oklahoma", "France 2nd", "AIB"]
def _sk(nm): return next((k for k in SKEYS if k in nm), nm)
CAP2K = {_sk(k): k for k in CAP2 if k != "COMPANY"}
CAP_ASOF = "Capacity & Power 'Quarterly MW' (copied values): v1 = nebius_capacity_v1.xlsx, v2 = nebius_capacity.xlsx (Oct 4 2026 22:33 ET); selector Inputs!CapV"
META = {  # keyword: (ownership, basis, grade)
 "Vineland": ("Build-to-suit lease", "IT", "Red"), "Independence": ("Owned", "Facility", "Green"), "Butler": ("Owned", "Facility", "Red"),
 "Birmingham": ("Owned", "Facility", "Yellow"), "Bethune": ("Leased-colo", "Facility", "Yellow"), "Lappeenranta": ("Owned", "Facility", "Green"),
 "Mantsala FI #1": ("Owned", "IT", "Green"), "Mantsala FI #2": ("Owned", "Facility", "Yellow"), "Kansas City": ("Leased-colo", "IT", "Green"),
 "Paris": ("Leased-colo", "IT", "Green"), "Keflavik": ("Leased-colo", "IT", "Green"), "Longcross": ("Leased-colo", "IT", "Green"),
 "Harlow": ("Leased-colo", "IT", "Yellow"), "UK 3rd": ("Leased-colo", "IT", "Red"), "Modi'in": ("Leased-colo", "IT", "Green"),
 "Masmiyya": ("Leased-colo", "Facility", "Yellow"), "Beit Shemesh": ("Leased-colo", "Facility", "Yellow"), "Madrid": ("Leased-colo", "IT", "Green"),
 "Tallinn": ("Leased-colo", "IT", "Yellow"), "Minnesota": ("Leased-colo", "IT", "Yellow"), "Oklahoma": ("Leased-colo", "IT", "Yellow"),
 "AIB": ("Leased-colo", "IT", "Red"), "Newport": ("Leased-colo", "IT", "Yellow")}
UNSEEN = {"Oklahoma": "Oklahoma (7 DC roles, live-looking)", "Minnesota": "Minnesota (RMA roles, likely colo live)"}
SITES = []
CAPCO = CAP.pop("COMPANY", {})
for nm_ in CAP:
    m = next((v for k, v in META.items() if k in nm_), ("Leased-colo", "IT", "Yellow"))
    key = next((k for k in UNSEEN if k in nm_), None)
    k2 = CAP2K.get(_sk(nm_))
    SITES.append({"name": UNSEEN[key] if key else nm_, "own": m[0], "basis": m[1], "grade": m[2], "cap": CAP[nm_], "unseen": key, "cap2": CAP2.get(k2) if k2 else None, "f2": BAS2.get(k2, 1.0) if k2 else 1.0})
_v1k = {_sk(n) for n in CAP if n != "COMPANY"}
for k_, n2 in CAP2K.items():
    if k_ not in _v1k:  # v2-only sites (zero under CapV = 1)
        m = next((v for k, v in META.items() if k in n2), ("Leased-colo", "IT", "Yellow"))
        SITES.append({"name": n2, "own": m[0], "basis": m[1], "grade": m[2], "cap": None, "unseen": None, "cap2": CAP2[n2], "f2": BAS2.get(n2, 1.0), "v2only": True})
CAPCO2 = CAP2.get("COMPANY", {})
SITES.append({"name": "Other unidentified (colo not yet visible)", "own": "Leased-colo", "basis": "IT", "grade": "Red", "cap": None, "unseen": "Other"})
def capvals(site, block, v=1):
    c_ = site.get("cap2") if (v == 2 or site.get("cap") is None) else site["cap"]
    if c_ is None or block not in c_: return None
    return c_[block][1]
st.section("Site metadata (ownership switches shell+power between capex [Owned] and rent [Leased-colo / Build-to-suit lease]; MW basis flag converts Facility MW to IT MW via PUE)")
own_dv = DataValidation(type="list", formula1='"Owned,Leased-colo,Build-to-suit lease"'); sws.add_data_validation(own_dv)
bas_dv = DataValidation(type="list", formula1='"IT,Facility"'); sws.add_data_validation(bas_dv)
grd_dv = DataValidation(type="list", formula1='"Green,Yellow,Red"'); sws.add_data_validation(grd_dv)
gen_dv = DataValidation(type="list", formula1='"Hopper H100/H200,B200 (HGX),B300 (HGX),GB200 NVL72,GB300 NVL72,Vera Rubin NVL72,n/d"'); sws.add_data_validation(gen_dv)
SITEGEN = {}
for _r in _ox.load_workbook(CAPF1, data_only=True)['GPUs deployed'].iter_rows(values_only=True):
    if _r[0] and _r[1] and _r[0] in CAP:
        g_ = str(_r[1]); SITEGEN[_r[0]] = ("GB300 NVL72" if "GB300" in g_ else "B200 (HGX)" if "B200" in g_ else "Hopper H100/H200" if ("Hopper" in g_ or "H100" in g_) else g_)
for _r in _ox.load_workbook(CAPF2, data_only=True)['GPUs deployed'].iter_rows(values_only=True):
    if _r[0] and _r[1] and _r[0] in CAP2 and _r[0] not in SITEGEN:
        g_ = str(_r[1]); SITEGEN[_r[0]] = ("GB300 NVL72" if "GB300" in g_ else "B200 (HGX)" if "B200" in g_ else "B300 (HGX)" if "B300" in g_ else "Vera Rubin NVL72" if "Rubin" in g_ else "Hopper H100/H200" if ("Hopper" in g_ or "H100" in g_) else g_)
ev_dv = DataValidation(type="list", formula1='"Interconnect,Permits,Job postings,Shipments,Company disclosure,None yet"'); sws.add_data_validation(ev_dv)
for c, t in [("A", "Site"), ("B", "Ownership"), ("C", "MW basis"), ("D", "Capacity-file basis / notes"), ("E", "Risk grade"), ("F", "IT factor"), ("G", "Owned?"),
             ("H", "Go-live idx - early (bull)"), ("I", "Go-live idx - base"), ("J", "Go-live idx - late (bear)"), ("K", "Go-live (live)"), ("L", "Evidence tag"), ("M", "Owner / due"), ("N", "GPU generation"), ("Q", "IT factor - Capacity v2 Basis tab (1/factor if gross)"), ("O", "Capex $M/IT MW"), ("P", "of which shell+power")]:
    sws[f"{c}{st.r}"] = t; sws[f"{c}{st.r}"].font = BOLD
st.r += 1
meta0 = st.r
for i, s_ in enumerate(SITES):
    rr = meta0 + i
    sws[f"A{rr}"] = s_["name"]
    inp(sws[f"B{rr}"], s_["own"], "Estimated"); own_dv.add(f"B{rr}")
    inp(sws[f"C{rr}"], s_["basis"], "Estimated"); bas_dv.add(f"C{rr}")
    cb_ = s_["cap"]["Energized MW base (=midpoint)"][0] if s_["cap"] and "Energized MW base (=midpoint)" in s_["cap"] else "n/a"
    sws[f"D{rr}"] = f"file basis: {cb_}" + ("; Vineland leased from DataOne; Bloom PPA power" if "Vineland" in s_["name"] else "") + (
        "; PLACEHOLDER low/base/high on Inputs (Capacity & Power)" if s_["unseen"] else "") + ("; Bethune 0/30/60 MW YE26 (Capacity)" if "Bethune" in s_["name"] else "")
    sws[f"D{rr}"].font = ITAL
    inp(sws[f"E{rr}"], s_["grade"], "Estimated"); grd_dv.add(f"E{rr}")
    fml(sws[f"F{rr}"], f'=IF(CapV=2,Q{rr},IF(C{rr}="Facility",1/PUE,1))', "0.00")
    inp(sws[f"Q{rr}"], round(s_.get("f2", 1.0), 4), "Estimated", "0.000")
    fml(sws[f"G{rr}"], f'=IF(B{rr}="Owned",1,0)', "0")
    gl = []
    for blk in ["Energized MW high", "Energized MW base (=midpoint)", "Energized MW low"]:
        v = capvals(s_, blk)
        k_ = next((n for n in range(1, NQ + 1) if v and isinstance(v[n - 1], (int, float)) and v[n - 1] > 0), None) if v else None
        gl.append(k_)
    for c_, k_ in zip("HIJ", gl): inp(sws[f"{c_}{rr}"], k_, "Placeholder", "0")
    fml(sws[f"K{rr}"], f'=IF(N(CHOOSE(Scen,H{rr},I{rr},J{rr}))=0,"n/a",INDEX(Q_Lab,1,CHOOSE(Scen,H{rr},I{rr},J{rr})))', "@")
    inp(sws[f"L{rr}"], None, "Placeholder"); ev_dv.add(f"L{rr}")
    sws[f"M{rr}"] = "Capacity & Power - due Tue Oct 6 (prefill = first non-zero energized qtr in Capacity file)"; sws[f"M{rr}"].font = ITAL
    _g = SITEGEN.get(s_["name"]) or SITEGEN.get(CAP2K.get(_sk(s_["name"]), "?")) or "n/d"
    inp(sws[f"N{rr}"], _g, "Estimated"); gen_dv.add(f"N{rr}")
    fml(sws[f"O{rr}"], f'=IF(B{rr}="Owned",IF(N{rr}="GB300 NVL72",CPM_Own_GB300,IF(N{rr}="Vera Rubin NVL72",CPM_Own_VR,IF(N{rr}="GB200 NVL72",CPM_Own_GB200,CPM_Own_Other))),IF(N{rr}="GB300 NVL72",CPM_Lease_GB300,IF(N{rr}="Vera Rubin NVL72",CPM_Lease_VR,CPM_Lease)))', "0.0")
    fml(sws[f"P{rr}"], f'=IF(B{rr}="Owned",CPM_SP,0)', "0.00")
NS = len(SITES)
rrO = meta0 + NS
sws[f"A{rrO}"] = "Other / unallocated portfolio MW (plug to portfolio total)"; inp(sws[f"B{rrO}"], "Mixed", "Estimated"); inp(sws[f"C{rrO}"], "IT", "Estimated"); bas_dv.add(f"C{rrO}")
sws[f"D{rrO}"] = "Owned share from Inputs!OwnedSh"; inp(sws[f"E{rrO}"], "Yellow", "Estimated"); grd_dv.add(f"E{rrO}")
fml(sws[f"F{rrO}"], f'=IF(C{rrO}="Facility",1/PUE,1)', "0.00")
st.r = rrO + 2
# ---- import data (energized low / base / high) from Capacity file
st.section("Capacity-file import: energized MW low / base / high by site (as stated basis) - " + CAP_ASOF)
IMP = {}
for i, s_ in enumerate(SITES):
    for blk, lab in [("Energized MW low", "low"), ("Energized MW base (=midpoint)", "base"), ("Energized MW high", "high")]:
        v = capvals(s_, blk)
        if s_.get("v2only"):
            r1 = None
        elif s_["unseen"]:
            nm_in = {"Oklahoma": "Unseen_OK", "Minnesota": "Unseen_MN", "Other": "Unseen_OT"}[s_["unseen"]]
            col = {"low": "E", "base": "D", "high": "C"}[lab]
            r1 = st.row(f"{s_['name']} - {lab} [v1]", "MW", "Capacity", "PLACEHOLDER (Inputs)", kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False,
                         fn=lambda n, nm_in=nm_in, col=col: f"=IF({qc(n)}$5>=12,Inputs!${col}${SC[nm_in]},0)")
            for n in range(1, NQ + 1): sws[f"{qc(n)}{r1}"].fill = F_PLACE
        else:
            vals = {}
            for n in range(1, NQ + 1):
                x = v[n - 1] if v else None
                if isinstance(x, (int, float)): vals[n] = (x, "Estimated")
            r1 = st.row(f"{s_['name']} - {lab} [v1]", "MW", "Capacity", "Capacity file v1", values=vals, kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False)
        v2_ = capvals(s_, blk, 2) if s_.get("cap2") else None
        if v2_ is not None and any(isinstance(x, (int, float)) for x in v2_):
            r2 = st.row(f"{s_['name']} - {lab} [v2]", "MW", "Capacity", "Capacity file v2", values={n: (v2_[n - 1], "Estimated") for n in range(1, NQ + 1) if isinstance(v2_[n - 1], (int, float))},
                        kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False)
        else:
            r2 = r1
        if r1 is None:
            rr_ = st.row(f"{s_['name']} - {lab}", "MW", "Modeling", "selected (CapV)", kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False, fn=lambda n, r2=r2: f"=IF(CapV=2,N({qc(n)}{r2}),0)")
        elif r2 == r1:
            rr_ = r1
        else:
            rr_ = st.row(f"{s_['name']} - {lab}", "MW", "Modeling", "selected (CapV)", kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False, fn=lambda n, r1=r1, r2=r2: f"=IF(CapV=2,N({qc(n)}{r2}),N({qc(n)}{r1}))")
        IMP[(i, lab)] = rr_
cb_ = CAPCO.get("Identified sites - energized base", (None, [None] * NQ))[1]
cb2_ = CAPCO2.get("Identified sites - energized base", (None, [None] * NQ))[1]
tf1 = st.row("Capacity file v1: identified sites energized base (as stated, mixed basis)", "MW", "Capacity", "COMPANY row, copied", kind="stock", fmt="#,##0", hist_fill=False,
               values={n: (cb_[n - 1], "Estimated") for n in range(1, NQ + 1) if isinstance(cb_[n - 1], (int, float))})
tf2 = st.row("Capacity file v2: identified sites energized base (as stated, mixed basis)", "MW", "Capacity", "COMPANY row, copied", kind="stock", fmt="#,##0", hist_fill=False,
               values={n: (cb2_[n - 1], "Estimated") for n in range(1, NQ + 1) if isinstance(cb2_[n - 1], (int, float))})
tie_f = st.row("Capacity file (selected CapV): identified sites energized base", "MW", "Modeling", "", kind="stock", fmt="#,##0", hist_fill=False,
               fn=lambda n: f"=IF(CapV=2,N({qc(n)}{tf2}),N({qc(n)}{tf1}))")
tie_m = st.row("Model: sum of named-site base imports (as stated)", "MW", "Modeling", "", kind="stock", fmt="#,##0",
               fn=lambda n: "=" + "+".join(f"N({qc(n)}{IMP[(i, 'base')]})" for i, s_ in enumerate(SITES) if not s_["unseen"]) + "+(CapV=2)*(" + "+".join(f"N({qc(n)}{IMP[(i, 'base')]})" for i, s_ in enumerate(SITES) if s_["unseen"] and s_.get("cap2")) + ")")
st.row("Tie check (model - file; should be 0)", "MW", "Modeling", "", kind="stock", fmt="#,##0;(#,##0);-", fn=lambda n: f"={qc(n)}{tie_m}-{qc(n)}{tie_f}")
reg("Sites", "CapTie", st.r - 1)
BLK = {}
def block(key, title, portfolio_vals, portfolio_basis, site_fn, owner, note):
    st.section(title)
    first = st.r
    for i, s_ in enumerate(SITES):
        st.row(s_["name"], "MW", owner, "", fn=(lambda n, i=i: site_fn(i, n)), kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False)
    last = st.r - 1
    pr = st.row(f"Portfolio total ({portfolio_basis} basis)", "MW", owner, note, values=portfolio_vals.get("vals"),
                fn=portfolio_vals.get("fn"), kind="stock", fmt="#,##0;(#,##0);-", hist_fill=False)
    sws[f"E{pr}"] = portfolio_basis
    pf = "1" if portfolio_basis == "IT" else "(1/PUE)"
    orr = st.row("Other / unallocated (IT MW)", "MW", "Modeling", "= MAX(0, portfolio (IT) - named sites (IT))",
                 fn=lambda n: f"=MAX(0,N({qc(n)}{pr})*{pf}-SUMPRODUCT({qc(n)}{first}:{qc(n)}{last},$F${meta0}:$F${meta0+NS-1}))",
                 kind="stock", fmt="#,##0;(#,##0);-")
    tot = st.row(f"TOTAL IT MW - {key}", "MW", "Modeling", "every site normalized to IT basis via PUE", bold=True,
                 fn=lambda n: f"=SUMPRODUCT({qc(n)}{first}:{qc(n)}{last},$F${meta0}:$F${meta0+NS-1})+{qc(n)}{orr}",
                 kind="stock", fmt="#,##0;(#,##0);-")
    own = st.row(f"  of which owned IT MW - {key}", "MW", "Modeling", "",
                 fn=lambda n: f"=SUMPRODUCT({qc(n)}{first}:{qc(n)}{last},$F${meta0}:$F${meta0+NS-1},$G${meta0}:$G${meta0+NS-1})+{qc(n)}{orr}*{ANN('OwnedSh', n)}",
                 kind="stock", fmt="#,##0;(#,##0);-")
    BLK[key] = (tot, own, first, last, pr)
def contracted_site(i, n):
    v = capvals(SITES[i], "Contracted MW")
    x = v[n - 1] if v else None
    return (x, "Estimated") if isinstance(x, (int, float)) else None
cvals = {8: (2000, "Reported"), 9: (2750, "Estimated"), 10: (3500, "Reported"), 11: (4250, "Estimated")}
for n in range(12, NQ + 1): cvals[n] = (5000, "Reported" if n == 12 else "Estimated")
block("Contracted", "Block 1 - CONTRACTED MW (land + power commitments; pipeline, not revenue capacity)", {"vals": cvals}, "Facility", contracted_site, "Capacity",
      "Feb-26 >2 GW; May-26 >3.5 GW; YE26 target 5 GW (Q2'26 letter); flat after (no guidance)")
def energized_site(i, n):
    return f"=IF({qc(n)}$5<=LastHist,{qc(n)}{IMP[(i,'base')]},CHOOSE(Scen,{qc(n)}{IMP[(i,'high')]},{qc(n)}{IMP[(i,'base')]},{qc(n)}{IMP[(i,'low')]}))"
block("Energized", "Block 2 - ENERGIZED MW (Capacity estimate; bull = high, base = midpoint, bear = low). Named sites YE26 185/305/424 MW, YE27 764/1,069/1,374 MW",
      {"vals": {}}, "IT", energized_site, "Capacity", "No portfolio plug: energized = named + unseen sites only")
_ef, _el = BLK["Energized"][2], BLK["Energized"][3]
_F = f"$F${meta0}:$F${meta0+NS-1}"; _O = f"$O${meta0}:$O${meta0+NS-1}"; _P = f"$P${meta0}:$P${meta0+NS-1}"; _G = f"$G${meta0}:$G${meta0+NS-1}"
def _dlt(n): return f"(({qc(n)}{_ef}:{qc(n)}{_el}-{qc(n-1)}{_ef}:{qc(n-1)}{_el})>0)*({qc(n)}{_ef}:{qc(n)}{_el}-{qc(n-1)}{_ef}:{qc(n-1)}{_el})"
st.section("Path-consistent capex at energization (new energized IT MW x per-site $/IT MW by ownership & generation); Capex_DA shifts it to 1 qtr before billing")
pci = st.row("Path IT/GPU capex at energization quarter", "$M", "Capacity", "SUMPRODUCT(new energized MW x IT factor x (rate - shell+power))", fmt="#,##0",
             fn=lambda n: f"=SUMPRODUCT({_dlt(n)},{_F},{_O}-{_P})")
reg("Sites", "PathIT", pci)
pcs = st.row("Path shell+power capex (owned sites) at energization quarter", "$M", "Capacity", "", fmt="#,##0", fn=lambda n: f"=SUMPRODUCT({_dlt(n)},{_F},{_P})")
reg("Sites", "PathSP", pcs)
enn = st.row("New energized IT MW (sum of site increases x IT factor)", "MW", "Capacity", "", fmt="#,##0.0", fn=lambda n: f"=SUMPRODUCT({_dlt(n)},{_F})")
reg("Sites", "EnNewIT", enn)
enl = st.row("New energized LEASED IT MW (sum of site increases x IT factor x (1 - owned))", "MW", "Capacity", "lease-prepayment base", fmt="#,##0.0", fn=lambda n: f"=SUMPRODUCT({_dlt(n)},{_F},1-{_G})")
reg("Sites", "EnNewL", enl)
_avgIT = f"IFERROR(SUMPRODUCT((Q_Idx>LastHist)*PathIT)/SUMPRODUCT((Q_Idx>LastHist)*EnNewIT),CPM_Lease_GB300)"
_avgSP = f"IFERROR(SUMPRODUCT((Q_Idx>LastHist)*PathSP)/SUMPRODUCT((Q_Idx>LastHist)*EnNewIT),0)"
prt = st.row("Path IT capex rate per new energized IT MW (site mix)", "$M/MW", "Capacity", "fallback = forecast average when no new energized MW", fmt="0.0",
             fn=lambda n: f"=IF({qc(n)}{enn}>0.01,{qc(n)}{pci}/{qc(n)}{enn},{_avgIT})")
reg("Sites", "PathRateIT", prt)
prs = st.row("Path shell+power capex rate per new energized IT MW", "$M/MW", "Capacity", "", fmt="0.0",
             fn=lambda n: f"=IF({qc(n)}{enn}>0.01,{qc(n)}{pcs}/{qc(n)}{enn},{_avgSP})")
reg("Sites", "PathRateSP", prs)
vin_ = st.row("Vineland energized IT MW (Capacity, live scenario)", "MW", "Capacity", "caps Microsoft GPU consumption (VinCapOn)", kind="stock", fmt="#,##0.0", fn=lambda n: f"={qc(n)}{_ef}*$F${meta0}")
reg("Sites", "VinEnIT", vin_)
_r31 = st.r; st.ws[f"A{_r31}"] = "MW billing from Q1'31 not yet in Q4'30 active MW (energized in Q4'30 at EnLag 1; capex in Q4'30)"; st.ws[f"A{_r31}"].font = BOLD
fml(st.ws[f"E{_r31}"], "=MAX(0,IF(MWSel=2,INDEX(MW_En,1,MIN(28,29-EnLag))*MW_Haircut,INDEX(MW_Act,1,28)+INDEX(MW_New,1,28))-INDEX(MW_Act,1,28))", "#,##0", True)
name(wb, "MW_Add31", "Sites", f"$E${_r31}"); st.r += 1
gvals = {}
hact = [round(170 * v / 227.7, 1) for v in [11.4, 14.5, 32.1, 35.2, 50.9, 105.1, 146.1]]
for n in range(1, 8): gvals[n] = (hact[n - 1], "Estimated")
gvals[8] = (170, "Reported")
path = {9: 291.7, 10: 413.3, 11: 650.0, 12: 900.0, 13: 1125.0, 14: 1350.0, 15: 1575.0, 16: 1800.0}
for n, v in path.items(): gvals[n] = (v, "Estimated")
for n in range(17, NQ + 1): gvals[n] = (1800 + 250 * (n - 16), "Estimated")
block("Guided", "Block 3 - GUIDED COMPARISON ONLY (413 MW Q2'26 best-estimate path) - MANAGEMENT-GUIDED CONNECTED IT MW path (YE25 170 MW active reported; 800MW-1GW connected YE26 = 900 mid, restated Aug 12 call, not in Q2 letter; YE27 1.8 GW connected; billing = connected lagged BillLag)",
      {"vals": gvals}, "IT", lambda i, n: None, "Capacity/Filings", "Est. interpolation of guidance; pre-Q4'25 revenue-proportional (Hist tab)")
st.section("Capacity file 'GPUs deployed' (copied values, by generation) and named-site energized base in IT MW (fleet scaling base)")
_GN = {"H": "Hopper", "B": "Blackwell", "N": "Next-gen"}
def _gpuf(f):
    G_ = {g: [0.0] * NQ for g in "HBN"}; _hdr = False
    for _r in _ox.load_workbook(f, data_only=True)['GPUs deployed'].iter_rows(values_only=True):
        if _r[0] == "Site": _hdr = True; continue
        if not _hdr or not _r[0] or str(_r[0]).startswith(("TOTAL", "Reported")): continue
        gen = str(_r[1]); vals = [v if isinstance(v, (int, float)) else 0 for v in _r[4:32]]
        for n in range(NQ):
            v = vals[n]
            if "first" in gen:  # mixed rows: Hopper/H200 tranche capped at first non-zero value
                cap_ = next(x for x in vals if x > 0)
                G_["H"][n] += min(v, cap_); G_["B"][n] += max(0, v - cap_)
            elif "Hopper" in gen or "H200" in gen or "H100" in gen: G_["H"][n] += v
            elif "Rubin" in gen: G_["N"][n] += v
            else: G_["B"][n] += v
    return G_
GPUF = _gpuf(CAPF1); GPUF2 = _gpuf(CAPF2)
GFR = {}
for g in "HBN":
    g1_ = st.row(f"Capacity file v1 GPUs deployed - {_GN[g]}", "GPUs", "Capacity", "nebius_capacity_v1.xlsx 'GPUs deployed' (Estimated)",
                    values={n: (GPUF[g][n - 1], "Estimated") for n in range(1, NQ + 1)}, kind="stock", fmt="#,##0", hist_fill=False)
    g2_ = st.row(f"Capacity file v2 GPUs deployed - {_GN[g]}", "GPUs", "Capacity", "nebius_capacity.xlsx (v2) 'GPUs deployed' (Estimated)",
                    values={n: (GPUF2[g][n - 1], "Estimated") for n in range(1, NQ + 1)}, kind="stock", fmt="#,##0", hist_fill=False)
    GFR[g] = st.row(f"Capacity file GPUs deployed - {_GN[g]} (selected CapV)", "GPUs", "Modeling", "", kind="stock", fmt="#,##0", hist_fill=False,
                    fn=lambda n, g1_=g1_, g2_=g2_: f"=IF(CapV=2,{qc(n)}{g2_},{qc(n)}{g1_})")
    reg("Sites", f"FileGPU_{g}", GFR[g])
_ga = st.r
st.row("GPU history anchor gap: GPU_Q125 actual (~30,000 at 3/31/25, 20-F) - Capacity file total at Q1'25 (carried as Hopper-era fleet from Q1'25 when GPUAnchorOn)", "GPUs", "Filings/Capacity", "Kansas City likely live Q1'25 at 5-15 MW (Capacity, Estimated)", kind="none", fmt="#,##0",
       fn=lambda n: (f"=GPUAnchorOn*(GPU_Q125-(INDEX(FileGPU_H,1,5)+INDEX(FileGPU_B,1,5)+INDEX(FileGPU_N,1,5)))" if n == 1 else None))
name(wb, "GPUAnchorD", "Sites", f"$F${_ga}")
_named = [i for i, s_ in enumerate(SITES) if not s_["unseen"]]
_nv2 = [i for i, s_ in enumerate(SITES) if s_["unseen"] and s_.get("cap2")]  # MN/OK: file GPUs exist in v2 -> include in denominator under CapV 2
ebi = st.row("Named-site energized BASE, IT MW (PUE-normalized; CapV 2 adds MN/OK which carry v2 file GPUs)", "MW", "Capacity", "Denominator for GPU-fleet scaling", kind="stock", fmt="#,##0",
             fn=lambda n: "=" + "+".join(f"N({qc(n)}{IMP[(i, 'base')]})*$F${meta0+i}" for i in _named) + ("+(CapV=2)*(" + "+".join(f"N({qc(n)}{IMP[(i, 'base')]})*$F${meta0+i}" for i in _nv2) + ")" if _nv2 else ""))
reg("Sites", "EnBaseIT", ebi)
st.section("Revenue-driving MW (selector, delay, haircut)")
st.section("Capacity proposed path (mw_path_proposed.csv, Capacity, Estimated) - MWSel 4 scenario, not base until Capacity's Tuesday file")
_PPD = {'low': [10, -37.8, 4.0, 35.0, 81.0, 135.0], 'base': [20, 15, 40.5, 90.5, 148, 153], 'high': [45, 80, 145, 178, 208, 208]}
_pp = {}
for _k in ["low", "base", "high"]:
    _pp[_k] = st.row(f"Proposed path delta vs model energized - {_k} (active MW; held at Q4'27 delta after)", "MW", "Capacity", f"csv: path - model_{_k} (Oklahoma, Minnesota, Newport/UK, CLT1 timing, Masmiyya/Beit Shemesh slips); Vineland IT-basis restatement excluded (no MW, no capex)",
                     kind="stock", fmt="#,##0.0", values={n: (_PPD[_k][min(n, 16) - 11] if n >= 11 else 0, "Estimated") for n in range(1, NQ + 1)}, hist_fill=False)
ppl = st.row("Proposed path delta - live scenario (leased colo; owned MW unchanged)", "MW", "Modeling", "CHOOSE(Scen, high, base, low)", kind="stock", fmt="#,##0.0",
             fn=lambda n: f"=CHOOSE(Scen,{qc(n)}{_pp['high']},{qc(n)}{_pp['base']},{qc(n)}{_pp['low']})")
rawsel = st.row("Selected block - IT MW (raw)", "MW", "Modeling", "CHOOSE(MWSel, contracted, energized, guided, energized + proposed-path delta)",
                fn=lambda n: f"=CHOOSE(MWSel,{qc(n)}{BLK['Contracted'][0]},{qc(n)}{BLK['Energized'][0]},{qc(n)}{BLK['Guided'][0]},{qc(n)}{BLK['Energized'][0]}+{qc(n)}{ppl})", kind="stock", fmt="#,##0")
rawown = st.row("Selected block - owned IT MW (raw)", "MW", "Modeling", "",
                fn=lambda n: f"=CHOOSE(MWSel,{qc(n)}{BLK['Contracted'][1]},{qc(n)}{BLK['Energized'][1]},{qc(n)}{BLK['Guided'][1]},{qc(n)}{BLK['Energized'][1]})", kind="stock", fmt="#,##0")
_aib = [i for i, s_ in enumerate(SITES) if "AIB" in s_["name"]][0]
clt = st.row("AIB CLT1 energized IT MW (live scenario; Capacity file; Reported 8-K Sep 30 2026: 50 MW, 2 x 25 MW halls ~Aug & Dec 2027)", "MW", "Capacity", "YE27 0 / 25 / 50 bear/base/bull", kind="stock", fmt="#,##0.0",
             fn=lambda n: f"=N({energized_site(_aib, n)[1:]})*$F${meta0+_aib}")
reg("Sites", "CLT1_En", clt)
st.r += 0
_mu = st.r
st.row("MW basis uplift: company-implied Q2'26 MW - Capacity named-site base (HistMWBasis 2/3)", "MW", "Filings/Capacity", "", kind="none", fmt="#,##0.0",
       fn=lambda n: (f"=CoMWQ2-INDEX(EnBaseIT,1,LastHist)" if n == 1 else None))
name(wb, "MWUp", "Sites", f"$F${_mu}")
drv0 = st.row("ACTIVE IT MW - build plan (before stop-building switch)", "MW", "Modeling", "History = selected block (default energized base); forecast shifted by EnLag (energized) or BillLag (connected guide) + MW_Delay, x MW_Haircut", bold=True,
             fn=lambda n: f"=IF({qc(n)}$5<=LastHist,{qc(n)}{rawsel}+IF(AND(HistMWBasis>=2,{qc(n)}$5>=LastHist-1),MWUp,0),INDEX($F{rawsel}:$AG{rawsel},1,MAX(LastHist,{qc(n)}$5-MW_Delay-CHOOSE(MWSel,0,EnLag,BillLag,EnLag)))*MW_Haircut+IF(HistMWBasis=2,MWUp,0))", kind="stock", fmt="#,##0")
reg("Sites", "MW_Act0", drv0)
drv = st.row("ACTIVE IT MW driving revenue (StopBuild: contract + post-term MW + non-contract MW frozen at StopIdx)", "MW", "Modeling", "v0.8", bold=True, kind="stock", fmt="#,##0",
             fn=lambda n: f"=IF(AND(StopBuild=1,{qc(n)}$5>StopIdx),{X('A_MW', n)}+{X('A_PostMW', n)}+MAX(0,INDEX(MW_Act0,1,StopIdx)-INDEX(A_MW,1,StopIdx)-INDEX(A_PostMW,1,StopIdx)),{qc(n)}{drv0})")
reg("Sites", "MW_Act", drv)
ownsh = st.row("Owned share of active IT MW", "%", "Modeling", "",
               fn=lambda n: f"=IF({qc(n)}{rawsel}>0,{qc(n)}{rawown}/{qc(n)}{rawsel},0)", kind="avg", fmt="0%")
reg("Sites", "MW_OwnSh", ownsh)
fac = st.row("Active facility MW (IT x PUE)", "MW", "Modeling", "", fn=lambda n: f"={qc(n)}{drv}*PUE", kind="stock", fmt="#,##0")
reg("Sites", "MW_Fac", fac)
st.row("Contracted IT MW (for EV/MW)", "MW", "Modeling", "", fn=lambda n: f"={qc(n)}{BLK['Contracted'][0]}", kind="stock", fmt="#,##0")
reg("Sites", "MW_Contr", st.r - 1)
riskrow = st.row("Count of Red-graded sites", "#", "Modeling", "", fn=lambda n: (f'=COUNTIF($E${meta0}:$E${rrO},"Red")' if n == 1 else None), kind="none", fmt="0")
name(wb, "RedSites", "Sites", f"$F${riskrow}")

import pickle
pickle.dump({"RN": RN}, open("/tmp/rn.pkl", "wb"))

# ------------------------------------------------------------------ PRICING
pr_ = TS(wb, "Pricing", "Chip-level list prices (on-demand / reserved / spot) and quarterly price paths by generation. Hyperscaler take-or-pay is priced separately per $/MW-yr on Contracts.")
pws = pr_.ws
pr_.section("Chip price table ($/GPU-hr) - nebius.com/prices snapshot Oct 4 2026")
hdr = ["Chip", "Gen", "On-demand prior", "On-demand eff. Oct-1-26", "Spot floor ('from')", "Reserved floor", "Within-gen weight", "Tag", "Source / method"]
for i, t in enumerate(hdr):
    pws.cell(pr_.r, 1 + i, t).font = BOLD
pr_.r += 1
c0 = pr_.r
CH = [("H100", "H", 3.85, 4.50, 0.79, 0.5, "Reported", "nebius.com/prices (Oct 4 2026 snapshot; market/data/price_snapshot_2026-10-04.csv)"),
      ("H200", "H", 4.50, 5.40, 0.79, 0.5, "Reported", "nebius.com/prices"),
      ("B200", "B", 7.15, 8.50, 0.99, 0.25, "Reported", "nebius.com/prices"),
      ("B300", "B", 7.85, 9.50, 0.99, 0.25, "Reported", "nebius.com/prices"),
      ("GB200 NVL72", "B", None, None, None, 0.25, "Not disclosed", "'Contact us' -> est = B300 x GB_Prem (CoreWeave GB200/B200 ratio)"),
      ("GB300 NVL72", "B", None, None, None, 0.25, "Not disclosed", "'Contact us' -> est = B300 x GB_Prem"),
      ("Vera Rubin", "N", None, None, None, 1.0, "Not disclosed", "No list price -> est = B300 x Rubin_Prem (placeholder)")]
b300 = c0 + 3
for i, ch in enumerate(CH):
    rr = c0 + i
    pws[f"A{rr}"] = ch[0]; pws[f"B{rr}"] = ch[1]
    if ch[2] is not None:
        inp(pws[f"C{rr}"], ch[2], "Reported", "0.00"); inp(pws[f"D{rr}"], ch[3], "Reported", "0.00"); inp(pws[f"E{rr}"], ch[4], "Reported", "0.00")
    else:
        prem = "Rubin_Prem*GB_Prem" if ch[1] == "N" else "GB_Prem"
        fml(pws[f"C{rr}"], f"=C{b300}*{prem}", "0.00"); fml(pws[f"D{rr}"], f"=D{b300}*{prem}", "0.00"); fml(pws[f"E{rr}"], f"=E{b300}", "0.00")
        for c in "CDE": pws[f"{c}{rr}"].fill = F_EST
    fml(pws[f"F{rr}"], f"=D{rr}*(1-ResDisc)", "0.00")
    inp(pws[f"G{rr}"], ch[5], "Placeholder", "0%")
    pws[f"H{rr}"] = ch[6]; pws[f"I{rr}"] = ch[7]; pws[f"I{rr}"].font = ITAL
pws[f"J{c0-1}"] = "OD Q2'26 avg"; pws[f"J{c0-1}"].font = BOLD
for i, ch in enumerate(CH):
    rr = c0 + i
    fml(pws[f"J{rr}"], "=ODq_H100" if ch[0] == "H100" else f"=C{rr}", "0.00")
c1 = c0 + len(CH) - 1
pr_.r = c1 + 2
pws[f"A{pr_.r}"] = "Generation blend"; pws[f"A{pr_.r}"].font = BOLD
for i, t in enumerate(["", "Gen", "OD prior", "OD current", "Spot", "Reserved"]):
    if t: pws.cell(pr_.r, 1 + i, t).font = BOLD
pr_.r += 1
GENS = ["H", "B", "N"]
GEN_NAME = {"H": "Hopper", "B": "Blackwell", "N": "Next-gen"}
for g in GENS:
    rr = pr_.r
    pws[f"A{rr}"] = GEN_NAME[g]; pws[f"B{rr}"] = g
    for c, nmk in zip("CDEF", ["ODp", "ODc", "Spot", "Res"]):
        fml(pws[f"{c}{rr}"], f"=SUMPRODUCT(($B${c0}:$B${c1}=$B{rr})*{c}${c0}:{c}${c1}*$G${c0}:$G${c1})/SUMIF($B${c0}:$B${c1},$B{rr},$G${c0}:$G${c1})", "0.00")
        name(wb, f"{nmk}_{g}", "Pricing", f"${c}${rr}")
    fml(pws[f"J{rr}"], f"=SUMPRODUCT(($B${c0}:$B${c1}=$B{rr})*J${c0}:J${c1}*$G${c0}:$G${c1})/SUMIF($B${c0}:$B${c1},$B{rr},$G${c0}:$G${c1})", "0.00")
    name(wb, f"ODq_{g}", "Pricing", f"$J${rr}")
    pr_.r += 1
def PXF(g, n):
    """annual list-price decline factor. PxPathMode 0 = v0.3 (from PxBaseIdx); 1 = v0.4 (year 2 onward: Blackwell after post-Rubin cut, Next-gen after launch year; Hopper from PxBaseIdx); PwrLooseOn = steeper decline from PwrLooseIdx (Blackwell/Next-gen)"""
    st = {"H": "PxBaseIdx", "B": "(PxCycIdx+4)", "N": "(NextLaunchIdx+4)"}[g]
    v03 = f"(1+PxChg_{g})^(MAX(0,{n}-PxBaseIdx)/4)"
    if g == "H":
        return v03
    v04 = (f"(1+PxChg_{g})^(MAX(0,IF(PwrLooseOn=1,MIN({n},PwrLooseIdx),{n})-{st})/4)"
           f"*IF(PwrLooseOn=1,(1+PwrLooseDecl)^(MAX(0,{n}-MAX({st},PwrLooseIdx))/4),1)")
    return f"IF(PxPathMode=1,{v04},{v03})"
pr_.section("Quarterly list-price paths by generation ($/GPU-hr); prior column before PxBaseIdx, Oct-1-26 column after; forecast x PxMult x (1+annual change)^(quarters/4)")
for g in GENS:
    od = pr_.row(f"On-demand - {GEN_NAME[g]}", "$/GPU-hr", "GPU Market", "", kind="avg", fmt="0.00",
                 fn=lambda n, g=g: f"=IF(AND(RepriceOn=1,{qc(n)}$5<=LastHist),ODq_{g},IF({qc(n)}$5<PxBaseIdx,ODp_{g},ODc_{g}))*IF({qc(n)}$5>LastHist,PxMult,1)*{PXF(g, qc(n)+'$5')}" + (f"*IF(AND(LaunchOn=1,{qc(n)}$5>=NextLaunchIdx,{qc(n)}$5<NextLaunchIdx+HeadStart/3),1+LaunchPrem,1)" if g == "N" else "")
                 + (f"*(1-PxCycleOn*PxCyc*MIN(1,MAX(0,({qc(n)}$5-PxCycIdx+1)/4)))" if g == "B" else ""))
    reg("Pricing", f"P_OD_{g}", od)
    def _newres(n, od, g):
        c = qc(n)
        w = f"MIN(1,MAX(0,({c}$5-RenewIdx+1)*RenewRate))"
        leg = f"(RealRes_B*ODq_{g}*(1-{w})+{w}*MIN(RealRes_B*ODq_{g}*(1+RenewUp),RealRes*{c}{od}))"
        L = f"INDEX(NCE_{g},1,LastHist)"; N = f"INDEX(NCA_{g},1,{n})"
        return f"IF({c}$5<=LastHist,RealRes_B*{c}{od},IF({N}>0,(MIN({L},{N})*{leg}+MAX(0,{N}-{L})*RealRes*{c}{od})/{N},RealRes*{c}{od}))"
    rs = pr_.row(f"Reserved/commit - {GEN_NAME[g]}", "$/GPU-hr", "GPU Market", "RepriceOn: legacy 6/30 book at Q2 rate, renewing RenewRate/qtr at min(+RenewUp, current list) x RealRes; GPUs added after 6/30 at current list x RealRes. Off: lagged list", kind="avg", fmt="0.00",
                 fn=lambda n, od=od, g=g: f"=IF(RepriceOn=1," + _newres(n, od, g) + f",(IF({qc(n)}$5>LastHist,MIN(1,MAX(0,({qc(n)}$5-RenewIdx+1)*RenewRate)),0)*{qc(n)}{od}+(1-IF({qc(n)}$5>LastHist,MIN(1,MAX(0,({qc(n)}$5-RenewIdx+1)*RenewRate)),0))*INDEX($F${od}:$AG${od},1,MAX(1,{qc(n)}$5-IF({qc(n)}$5>LastHist,ResLag,0))))*IF({qc(n)}$5<=LastHist,RealRes_B,RealRes))")
    reg("Pricing", f"P_RES_{g}", rs)
    sp = pr_.row(f"Spot - {GEN_NAME[g]}", "$/GPU-hr", "GPU Market", "Floor price; dynamic", kind="avg", fmt="0.00",
                 fn=lambda n, g=g: f"=Spot_{g}*IF({qc(n)}$5>LastHist,PxMult,1)*{PXF(g, qc(n)+'$5')}")
    reg("Pricing", f"P_SPOT_{g}", sp)

# ---- revenue per IT MW at each generation change + new-deal benchmark (IREN)
pr_.r += 1
_r0 = pr_.r
pws[f"A{_r0}"] = "Revenue per IT MW-yr at current (Oct-1-26) list, by generation - new-deal check vs IREN >$20M/IT MW-yr (3-yr GB300-class, ~$4.65/GPU-hr; company-reported, IT basis). Nebius '>$20M' basis-unknown; Q3 '>$40M' short-term deals excluded"; pws[f"A{_r0}"].font = BOLD
for i, t in enumerate(["", "GB200 NVL72", "GB300 NVL72", "Vera Rubin", "Rubin / GB300"]):
    if t: pws.cell(_r0 + 1, 1 + i, t).font = BOLD
_gbrow = {"GB200 NVL72": c0 + 4, "GB300 NVL72": c0 + 5, "Vera Rubin": c0 + 6}
_lines = [("GPUs per IT MW", lambda k: "=GPUMW_N" if k == "Vera Rubin" else "=GPUMW_B", "#,##0"),
          ("On-demand list $/GPU-hr (Oct-1-26; Rubin = GB300 x Rubin_Prem)", lambda k: f"=D{_gbrow[k]}*PxMult", "0.00"),
          ("Contract rate $/GPU-hr = list x RealRes (live scenario)", lambda k: f"=D{_gbrow[k]}*PxMult*RealRes", "0.00"),
          ("Gross revenue per IT MW-yr at contract rate x Util_Res ($M)", lambda k: f"=B{{gp}}*D{_gbrow[k]}*PxMult*RealRes*8760*Util_Res/1000000".replace("B{gp}", "{col}" + str(_r0 + 2)), "0.0"),
          ("Net of facility cost - LEASED (less CLT1Rent $2.2M/IT MW-yr), before power", None, "0.0"),
          ("Net of facility cost - OWNED (less owned-site O&M OM_MW), before power", None, "0.0")]
for j, (lab, f, fm) in enumerate(_lines):
    rr = _r0 + 2 + j
    pws[f"A{rr}"] = lab
    for ci, k in enumerate(["GB200 NVL72", "GB300 NVL72", "Vera Rubin"]):
        col = "BCD"[ci]
        if j == 3: frm = f"={col}{_r0+2}*{col}{_r0+4}*8760*Util_Res/1000000"
        elif j == 4: frm = f"={col}{_r0+5}-CLT1Rent"
        elif j == 5: frm = f"={col}{_r0+5}-OM_MW"
        else: frm = f(k)
        fml(pws[f"{col}{rr}"], frm, fm)
    fml(pws[f"E{rr}"], f"=IFERROR(D{rr}/C{rr},0)", "0.00x")
name(wb, "RM_GB300", "Pricing", f"$C${_r0+5}"); name(wb, "RM_Rubin", "Pricing", f"$D${_r0+5}"); name(wb, "RM_GB300_L", "Pricing", f"$C${_r0+6}"); name(wb, "RM_GB300_O", "Pricing", f"$C${_r0+7}")
name(wb, "RM_GB300_Rate", "Pricing", f"$C${_r0+4}")
rr = _r0 + 8
pws[f"A{rr}"] = "Benchmark IREN $20M/IT MW-yr gross | leased net (less $2.2M rent) | owned net (less OM_MW)"
fml(pws[f"B{rr}"], "=20", "0.0"); fml(pws[f"C{rr}"], "=20-CLT1Rent", "0.0"); fml(pws[f"D{rr}"], "=20-OM_MW", "0.0")
pws[f"A{rr+1}"] = "Rubin $/GPU-hr premium vs GB300 needed to hold revenue/IT MW (= GPUs/MW GB300 / Rubin) vs model Rubin_Prem"
fml(pws[f"B{rr+1}"], "=GPUMW_B/GPUMW_N", "0.00x"); fml(pws[f"C{rr+1}"], "=Rubin_Prem", "0.00x")
pws[f"A{rr+2}"] = "NEW-DEAL economics (Q2'26 letter), $M/IT MW-yr at NewDealIdx vintage: B multi-year | C short-term | D blended (1-STShare, STShare) | E blended net leased | F blended net owned"
fml(pws[f"B{rr+2}"], "=INDEX(ND_MYRate,1,NewDealIdx)", "0.0"); fml(pws[f"C{rr+2}"], "=INDEX(ND_STRate,1,NewDealIdx)*IF(STMode=1,STOcc,1)", "0.0")
fml(pws[f"D{rr+2}"], f"=(1-STShare)*B{rr+2}+STShare*C{rr+2}", "0.0"); fml(pws[f"E{rr+2}"], f"=D{rr+2}-CLT1Rent", "0.0"); fml(pws[f"F{rr+2}"], f"=D{rr+2}-OM_MW", "0.0")
name(wb, "RM_ND", "Pricing", f"$D${rr+2}"); name(wb, "RM_ND_L", "Pricing", f"$E${rr+2}"); name(wb, "RM_ND_O", "Pricing", f"$F${rr+2}"); name(wb, "RM_MY", "Pricing", f"$B${rr+2}")
pws[f"A{rr+3}"] = "Multi-year tranche net of facility cost vs IREN: B gross | C net leased (vs 17.8) | D net owned (vs 20 - OM_MW)"
fml(pws[f"B{rr+3}"], f"=B{rr+2}", "0.0"); fml(pws[f"C{rr+3}"], f"=B{rr+2}-CLT1Rent", "0.0"); fml(pws[f"D{rr+3}"], f"=B{rr+2}-OM_MW", "0.0")
pr_.r = rr + 5

# ------------------------------------------------------------------ CONTRACTS
ct = TS(wb, "Contracts", "Book (a): hyperscaler / investment-grade take-or-pay contracts, tranche by tranche. Fixed $/MW-yr (insensitive to list price), cancellation toggle + probability, late-delivery LDs, post-term merchant residual stream.")
cws = ct.ws
ct.section("Tranche table (Capacity owns delivery quarters; Skeptic owns cancellation / LD; Filings owns TCV)")
th = ["Tranche", "Customer", "Gen", "IT MW", "ACV $M/MW-yr", "Term (qtrs)", "Sched. qtr idx", "Billing start idx (live)", "Cancelled (0/1)", "Cancel prob.", "Implied TCV $M", "$/GPU-hr equiv.", "TCV basis $M", "Tag", "Owner", "Source / method", "Start - Bull", "Start - Base", "Start - Bear", "Billing ramp (qtrs)", "Start-qtr factor"]
for i, t in enumerate(th):
    cws.cell(ct.r, 1 + i, t).font = BOLD
ct.r += 1
t0 = ct.r
TR = []
MSCH = [(8, 8, 8), (9, 9, 9), (11, 11, 11), (11, 11, 11), (13, 14, 15), (13, 14, 15), (14, 15, 16), (14, 15, 16), (14, 15, 17)]
for k in range(9):
    TR.append((f"MSFT T{k+1}", "Microsoft", "B", "=MSFT_MW/MSFT_N", "=MSFT_TCV/(5*MSFT_MW)", 20, MSCH[k][1], None, "Reported" if k < 4 else "Estimated",
               ("Delivered (T1 Nov-25, T2 Feb-26, T3-T4 Jul-Aug-26) per Capacity tranche track" if k < 4 else "Not delivered; needs Bloom fuel cells; billing bull/base/bear per Capacity tranche track; no Q4'26 billing") +
               "; ~26.7 IT MW each (equal-size ASSUMPTION); ACV derived = TCV/(5 x MW)", MSCH[k]))
TR.append(("MSFT addendum A", "Microsoft", "B", 0, "=MSFT_TCV/(5*MSFT_MW)", 20, 12, None, "Placeholder", "Up to ~$19.4B incl. optional capacity (+$2.0B); Jan-26 addendum +2 tranches (Capacity) - MW placeholder 0"))
TR.append(("MSFT addendum B", "Microsoft", "B", 0, "=MSFT_TCV/(5*MSFT_MW)", 20, 12, None, "Placeholder", "As above"))
TR.append(("Meta #1 T1", "Meta", "B", "TCVB", "=ACV_2026", 20, 8, 1440.35, "Derived", "Nov-1-25 CISA: ~$2,880.7M total, 5-yr, 2 tranches (Dec-25, Feb-26); MW = TCV/(ACV x 5) with ACV = 2026 base $12M"))
TR.append(("Meta #1 T2", "Meta", "B", "TCVB", "=ACV_2026", 20, 9, 1440.35, "Derived", "As above"))
TR.append(("Meta #2 T1", "Meta", "N", "TCVB", "=ACV_Q2", 20, 15, 6000, "Estimated", "Mar-13-26: $12B dedicated Vera Rubin, 5-yr; company 'throughout 1H27'; base start Q3'27, bull 1H27, bear Q4'27+; split 50/50 est.; ACV = $22.5M/MW", (13, 15, 16)))
TR.append(("Meta #2 T2", "Meta", "N", "TCVB", "=ACV_Q2", 20, 16, 6000, "Estimated", "As above", (14, 16, 18)))
TR.append(("Meta backstop (contingent)", "Meta", "N", "TCVB", "=ACV_Q2", 20, 14, 15000, "Placeholder", "Up to $15B: Meta buys unsold capacity on certain clusters for 5 yrs. Contingent -> Cancelled=1 by default"))
TR.append(("Future IG contract 1", "TBD", "N", 0, "=ACV_Q2", 20, 16, None, "Placeholder", ">$40B customer commitments; RPO $37.5B at 6/30/26 - unnamed IG deals, MW placeholder 0"))
TR.append(("Future IG contract 2", "TBD", "N", 0, "=ACV_Q2", 20, 20, None, "Placeholder", ""))
for i, t in enumerate(TR):
    rr = t0 + i
    cws[f"A{rr}"] = t[0]; cws[f"B{rr}"] = t[1]; inp(cws[f"C{rr}"], t[2], "Estimated")
    if t[3] == "TCVB":
        fml(cws[f"D{rr}"], f"=M{rr}/(E{rr}*F{rr}/4)", "#,##0.0")
    elif isinstance(t[3], str):
        fml(cws[f"D{rr}"], t[3], "#,##0.0"); cws[f"D{rr}"].fill = F_EST
    else:
        inp(cws[f"D{rr}"], t[3], "Placeholder", "#,##0.0")
    fml(cws[f"E{rr}"], t[4], "0.00")
    inp(cws[f"F{rr}"], t[5], "Reported" if t[1] != "TBD" else "Placeholder", "0")
    inp(cws[f"G{rr}"], t[6], t[8] if t[8] != "Derived" else "Reported", "0")
    sc_ = t[10] if len(t) > 10 else (t[6], t[6], t[6])
    for c_, v_ in zip("QRS", sc_): inp(cws[f"{c_}{rr}"], v_, "Estimated" if len(t) > 10 else "Placeholder", "0")
    fml(cws[f"H{rr}"], f"=IF(R{rr}<=LastHist,R{rr},CHOOSE(Scen,Q{rr},R{rr},S{rr})" + (f"+IF(R{rr}>=13,2*BloomSlipOn,0)" if t[1] == "Microsoft" else "") + ")", "0", True)  # billing already started at 6/30 = fact, scenario-invariant
    fml(cws[f"T{rr}"], (f"=IF(R{rr}<=LastHist,RampMS_B,IF(AND(T59Align=1,R{rr}>=13),1,RampMS))" if t[1] == "Microsoft" else "=1"), "0", True)
    fml(cws[f"U{rr}"], (f"=IF(AND(R{rr}>=13,T59Align=1),1,DelivFactor)" if t[1] == "Microsoft" else "=DelivFactor"), "0.00", True)
    inp(cws[f"I{rr}"], 1 if "backstop" in t[0] else 0, "Placeholder", "0")
    inp(cws[f"J{rr}"], 0.0, "Placeholder", "0%")
    fml(cws[f"K{rr}"], f"=D{rr}*E{rr}*F{rr}/4", "#,##0")
    fml(cws[f"L{rr}"], f"=IFERROR(E{rr}*1000000/(INDEX(GPUMW_H:GPUMW_N,1,1)*0+CHOOSE(MATCH(C{rr},{{\"H\",\"B\",\"N\"}},0),GPUMW_H,GPUMW_B,GPUMW_N)*8760),0)", "0.00")
    if t[7] is not None: inp(cws[f"M{rr}"], t[7], "Reported" if t[8] != "Estimated" else "Estimated", "#,##0.0")
    cws[f"N{rr}"] = t[8]; cws[f"O{rr}"] = "Filings/Capacity/Skeptic"; cws[f"P{rr}"] = t[9]
    for c in "NOP": cws[f"{c}{rr}"].font = ITAL
_bs = t0 + [k for k, t in enumerate(TR) if "backstop" in t[0]][0]
fml(cws[f"D{_bs}"], "=BackstopMW", "#,##0.0"); fml(cws[f"E{_bs}"], "=BackstopPx", "0.00"); fml(cws[f"I{_bs}"], "=IF(BackstopMW>0,0,1)", "0")
for _c in "QRS": fml(cws[f"{_c}{_bs}"], "=BackstopIdx", "0")
cws[f"P{_bs}"] = "v0.8: Meta backstop toggle (BackstopMW 0/150/190 at BackstopPx; bear $14.5M); off-RPO; Rubin; counts as contracted (ABS-eligible)"
t1 = t0 + len(TR) - 1
for rr in range(t0, t1 + 1):  # simpler $/GPU-hr formula
    cws[f"L{rr}"] = f'=E{rr}*1000000/(CHOOSE(MATCH(C{rr},{{"H";"B";"N"}},0),GPUMW_H,GPUMW_B,GPUMW_N)*8760)'
ct.r = t1 + 1
cws[f"A{ct.r}"] = "Total implied TCV (excl. cancelled)"; cws[f"A{ct.r}"].font = BOLD
fml(cws[f"K{ct.r}"], f"=SUMPRODUCT(K{t0}:K{t1},1-I{t0}:I{t1})", "#,##0", True)
name(wb, "TCV_Model", "Contracts", f"$K${ct.r}")
ct.r += 1
cws[f"A{ct.r}"] = "Microsoft implied TCV vs reported (check = 0)"
fml(cws[f"K{ct.r}"], f"=SUM(K{t0}:K{t0+8})-MSFT_TCV", "#,##0", True); name(wb, "MSFT_Chk", "Contracts", f"$K${ct.r}")
ct.r += 2
TRROWS = {"rev": [], "mw": [], "ld": [], "post": [], "postmw": [], "gen": []}
ct.section("Quarterly tranche build")
for i, t in enumerate(TR):
    rr = t0 + i
    a = f"$H${rr}"; F_ = f"$F${rr}"; live = f"(1-$I${rr})*(1-$J${rr})"
    if t[0] in ("MSFT T1", "MSFT T2", "MSFT T3", "MSFT T4"):
        live = live + "*(1-VinT14Loss*VinLossOn*({qc(n)}$5>=VinLossIdx))"
    rv = ct.row(f"{t[0]} - contract revenue", "$M", "", "",
                fn=lambda n, rr=rr, a=a, F_=F_, live=live: f"={live.replace('{qc(n)}', qc(n))}*$D${rr}*$E${rr}/4*MIN(1,MAX(0,({qc(n)}$5-{a}+1)/$T${rr}))*IF({qc(n)}$5={a},$U${rr},IF(AND({qc(n)}$5>{a},{qc(n)}$5<{a}+{F_}),1,IF({qc(n)}$5={a}+{F_},1-$U${rr},0)))")
    mw = ct.row(f"{t[0]} - delivered IT MW", "MW", "", "", kind="stock", fmt="#,##0.0",
                fn=lambda n, rr=rr, a=a, F_=F_, live=live: f"={live.replace('{qc(n)}', qc(n))}*$D${rr}*AND({qc(n)}$5>={a},{qc(n)}$5<{a}+{F_})")
    ld = ct.row(f"{t[0]} - late-delivery LDs (negative)", "$M", "", "",
                fn=lambda n, rr=rr, a=a: f"=-(1-$I${rr})*LD_Pct*$D${rr}*$E${rr}/4*AND({qc(n)}$5>=$G${rr},{qc(n)}$5<{a})")
    pm = ct.row(f"{t[0]} - post-term merchant MW", "MW", "", "", kind="stock", fmt="#,##0.0",
                fn=lambda n, rr=rr, a=a, F_=F_, live=live: f"={live.replace('{qc(n)}', qc(n))}*$D${rr}*({qc(n)}$5>={a}+{F_})")
    po = ct.row(f"{t[0]} - post-term merchant revenue (residual value)", "$M", "", "GPUs x launch OD price x decay curve(age) x ResidUtil x hours",
                fn=lambda n, rr=rr, a=a, pm=pm: f'={qc(n)}{pm}*CHOOSE(MATCH($C${rr},{{"H";"B";"N"}},0),GPUMW_H*ODc_H,GPUMW_B*ODc_B,GPUMW_N*ODc_N)*ResidUtil*{HRS(n)}*INDEX(Curve,1,MAX(1,MIN(8,INT(({qc(n)}$5-{a})/4)+1)))/1000000')
    for k, v in zip(["rev", "mw", "ld", "post", "postmw"], [rv, mw, ld, po, pm]): TRROWS[k].append(v)
    TRROWS["gen"].append(rr)
ct.section("Book (a) totals")
def sumrows(rows, n): return "=" + "+".join(f"{qc(n)}{x}" for x in rows)
NMS = 11  # Microsoft rows (9 base + 2 addendum)
ct.section("Microsoft / Vineland power ceiling (gensets until Bloom fuel cells live)")
msr = ct.row("Microsoft contract revenue - uncapped", "$M", "Modeling", "", fn=lambda n: sumrows(TRROWS["rev"][:NMS], n))
msm = ct.row("Microsoft delivered IT MW - uncapped", "MW", "Modeling", "", kind="stock", fn=lambda n: sumrows(TRROWS["mw"][:NMS], n))
cei = ct.row("Power ceiling (IT MW) = genset gross MW / PUE until Bloom live", "MW", "Capacity", "", kind="stock",
             fn=lambda n: f"=IF(AND(BloomOn=1,{qc(n)}$5>=BloomIdx+2*BloomSlipOn),{qc(n)}{msm},GensetMW/VinPUE)")
_VS = [('No shutdown', 0, 0, 0, 0.55), ('Partial 50% for 2 months', 51.25, 55, 6, 0.2), ('Partial minor-source 2 months', 100.95, 55, 6, 0.05), ('Full 1 month', 102.5, 30, 0, 0.1), ('Full 2 months', 102.5, 55, 6, 0.06), ('Full 3 months', 102.5, 55, 36, 0.04)]
ct.section("Vineland shutdown scenarios (Capacity, /workspace/nebius/capacity/vineland_shutdown.xlsx; outage from Nov 7 2026; scenarios & weights = Capacity judgment)")
_vs0 = ct.r
for _i, (_nm, _mw, _d4, _d1, _p) in enumerate(_VS):
    _rr = ct.r
    ct.ws[f"A{_rr}"] = f"{_i+1}. {_nm}"
    for _c, _v, _f in [("C", _mw, "#,##0.00"), ("D", _d4, "0"), ("E", _d1, "0"), ("F", _p, "0%")]:
        inp(ct.ws[f"{_c}{_rr}"], _v, "Estimated", _f)
    ct.ws[f"G{_rr}"] = "C = IT MW offline, D = days in Q4'26, E = days in Q1'27, F = probability"
    ct.r += 1
_vs1 = ct.r - 1
_vrow = lambda col: f"${col}${_vs0}:${col}${_vs1}"
def _vday(n):
    col = {12: "D", 13: "E"}.get(n)
    if col is None: return "0"
    return (f"IF(VinShutSel=0,0,IF(VinShutSel=7,SUMPRODUCT({_vrow('C')},{_vrow(col)},{_vrow('F')}),INDEX({_vrow('C')},VinShutSel)*INDEX({_vrow(col)},VinShutSel)))")
vso = ct.row("Vineland IT MW offline (avg in quarter) - selected shutdown scenario", "MW", "Capacity", "MW x days offline / days in quarter; physical only (billing MW not reduced, so no GPU reallocation to books b/c)", kind="avg",
             fn=lambda n: f"={_vday(n)}/{HRS(n)}*24")
reg("Contracts", "VinOff", vso)
msb = ct.row("Microsoft BILLING IT MW (capped at power ceiling)", "MW", "Modeling", "", kind="stock", bold=True, fn=lambda n: f"=MIN({qc(n)}{msm},{qc(n)}{cei})")
reg("Contracts", "MS_BillMW", msb); reg("Contracts", "MS_MW", msm)
msc0 = ct.row("Microsoft revenue - capped at power ceiling (before shutdown loss)", "$M", "Modeling", "",
              fn=lambda n: f"=IF({qc(n)}{msm}>0,{qc(n)}{msr}*{qc(n)}{msb}/{qc(n)}{msm},0)")
vl = ct.row("Vineland shutdown: Microsoft revenue lost = MW-days offline x VinRevMW / 365", "$M", "Capacity", "Prob-weighted (VinShutSel 7) Q4'26 ~ $67-68M at $14.5M/IT MW-yr (Capacity, Derived)",
            fn=lambda n: f"=MIN({qc(n)}{msc0},{_vday(n)}*VinRevMW/365)")
reg("Contracts", "VinLost", vl)
_t14 = TRROWS["rev"][:4]
sla = ct.row("Vineland T1-T4 SLA credits (PLACEHOLDER VinSLA x T1-T4 revenue, Q4'26-Q1'27; bear sensitivity VinT14Loss)", "$M", "Capacity", "NOT DISCLOSED - placeholder",
             fn=lambda n: f"=-VinT14Loss*VinSLAOn*VinSLA*AND({qc(n)}$5>=12,{qc(n)}$5<=13)*(" + "+".join(f"{qc(n)}{x}" for x in _t14) + ")")
reg("Contracts", "VinSLA_R", sla)
msc = ct.row("Microsoft revenue - capped, net of shutdown loss and SLA credits", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{msc0}-{qc(n)}{vl}+{qc(n)}{sla}")
_slip = [t0 + k for k in range(4, 9)]
idl = ct.row("IDLE Microsoft IT MW (built, not billing): T5-T9 Bloom-slip quarters + lost T1-T4 before redeployment (VinIdleOn) - NOT re-sold to new-deal books", "MW", "Capacity", "v0.7c fix: v0.7b re-priced these GPUs as new deals", kind="stock",
             fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,VinIdleOn*(BloomSlipOn*(" + "+".join(f"$D${r_}*(1-$I${r_})*AND($R${r_}>=13,{qc(n)}$5>=$H${r_}-2,{qc(n)}$5<$H${r_})" for r_ in _slip) + ")"
                          f"+VinT14Loss*VinLossOn*AND({qc(n)}$5>=VinLossIdx,{qc(n)}$5<VinLossIdx+VinRedGap)*(" + "+".join(f"$D${t0+k}" for k in range(4)) + ")))")
reg("Contracts", "VinIdleMW", idl)
reg("Contracts", "MS_Rev", msc)
ct.row("Vineland shutdown: Microsoft tranches offline (equivalent, ~26.7 IT MW each)", "#", "Capacity", "", kind="avg", fmt="0.0", fn=lambda n: f"={qc(n)}{vso}/26.67")
_vw = ct.r
ct.row("Memo: probability-weighted revenue at risk at VinRevMW (independent of VinShutSel) - Q4'26 in Q, Q1'27 in R", "$M", "Capacity", "Tie: Q4'26 ~ $67M (Capacity, Derived)", kind="none",
       fn=lambda n: (f"=SUMPRODUCT({_vrow('C')},{_vrow('D')},{_vrow('F')})*VinRevMW/365" if n == 12 else (f"=SUMPRODUCT({_vrow('C')},{_vrow('E')},{_vrow('F')})*VinRevMW/365" if n == 13 else None)))
name(wb, "VinExpQ4", "Contracts", f"${qc(12)}${_vw}")
ct.row("Check: Microsoft revenue vs ~$350M/qtr ceiling (positive = above)", "$M", "Capacity/Skeptic", "", kind="none",
       fn=lambda n: f"={qc(n)}{msc}-MSFT_Ceil")
reg("Contracts", "MS_CeilChk", ct.r - 1)
orv = ct.row("Other contract revenue (Meta, IG)", "$M", "Modeling", "", fn=lambda n: sumrows(TRROWS["rev"][NMS:], n))
ra = ct.row("Book (a) contract revenue", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{msc}+{qc(n)}{orv}"); reg("Contracts", "A_Rev", ra)
bmw = ct.row("Contract BILLING IT MW (Microsoft capped + others)", "MW", "Modeling", "", kind="stock",
             fn=lambda n: f"={qc(n)}{msb}+" + "+".join(f"{qc(n)}{x}" for x in TRROWS["mw"][NMS:]))
reg("Contracts", "A_BillMW", bmw)
rl = ct.row("Late-delivery LDs", "$M", "Modeling", "", fn=lambda n: sumrows(TRROWS["ld"], n)); reg("Contracts", "A_LD", rl)
rp = ct.row("Post-term merchant revenue", "$M", "Modeling", "", fn=lambda n: sumrows(TRROWS["post"], n)); reg("Contracts", "A_Post", rp)
rm = ct.row("Delivered contract IT MW", "MW", "Modeling", "", kind="stock", fn=lambda n: sumrows(TRROWS["mw"], n)); reg("Contracts", "A_MW", rm)
_sig = [TRROWS["mw"][i] for i, t in enumerate(TR) if t[1] in ("Microsoft", "Meta")]
name(wb, "BS_MW", "Contracts", f"$F${TRROWS['mw'][[i for i, t in enumerate(TR) if 'backstop' in t[0]][0]]}:$AG${TRROWS['mw'][[i for i, t in enumerate(TR) if 'backstop' in t[0]][0]]}")
_igr = [TRROWS["mw"][i] for i, t in enumerate(TR) if t[1] not in ("Microsoft", "Meta")]
x_ = ct.row("Signed Microsoft/Meta delivered IT MW (no new prepayment; advances already in 6/30 DR)", "MW", "Modeling", "", kind="stock", fn=lambda n: sumrows(_sig, n)); reg("Contracts", "SigMW", x_)
x_ = ct.row("New hyperscaler / IG tranche delivered IT MW (HSPre prepayment)", "MW", "Modeling", "", kind="stock", fn=lambda n: sumrows(_igr, n)); reg("Contracts", "IGMW", x_)
rpm = ct.row("Post-term merchant IT MW", "MW", "Modeling", "", kind="stock", fn=lambda n: sumrows(TRROWS["postmw"], n)); reg("Contracts", "A_PostMW", rpm)
for g in GENS:
    rows = [(TRROWS["mw"][i], TRROWS["postmw"][i]) for i, t in enumerate(TR) if t[2] == g]
    msi = [i for i, t in enumerate(TR) if t[2] == g and t[1] == "Microsoft"]
    oth = [i for i, t in enumerate(TR) if t[2] == g and t[1] != "Microsoft"]
    def fg(n, rows=rows, g=g, msi=msi, oth=oth):
        if not rows: return "=0"
        c = qc(n)
        post = "+".join(f"{c}{TRROWS['postmw'][i]}" for i in msi + oth)
        o = "+".join(f"{c}{TRROWS['mw'][i]}" for i in oth) or "0"
        if msi:
            m = "(" + "+".join(f"{c}{TRROWS['mw'][i]}" for i in msi) + ")"
            if g == "B": m = f"({m}+{X('VinIdleMW', n)})"
            m = f"IF(VinCapOn=1,MIN({m},INDEX(VinEnIT,1,MAX(1,{n}-IF({n}>LastHist,EnLag,0)))),{m})"
        else: m = "0"
        return f"=({m}+{o}+{post})*GPUMW_{g}"
    x = ct.row(f"Contract + post-term GPUs - {GEN_NAME[g]}", "GPUs", "Modeling", "", kind="stock", fmt="#,##0", fn=fg)
    reg("Contracts", f"A_GPU_{g}", x)

# ------------------------------------------------------------------ REVENUE
rv_ = TS(wb, "Revenue", "Three books: (a) hyperscaler take-or-pay [Contracts], (b) reserved/commit, (c) on-demand/self-serve + spot + AI Studio. Fleet built from active IT MW x GPUs/MW by deployment generation.")
R = rv_
R.section("Capacity & GPU fleet")
mwe = R.row("Active IT MW (end)", "MW", "Sites", "", kind="stock", fmt="#,##0", e=f"=INDEX(MW_Act,1,1)", fn=lambda n: f"={X('MW_Act', n)}")
mwa = R.row("Active IT MW (average)", "MW", "Modeling", "", kind="avg", fmt="#,##0", fn=lambda n: f"=AVERAGE({qc(n-1)}{mwe},{qc(n)}{mwe})")
reg("Revenue", "MW_Avg", mwa)
nmw = R.row("New IT MW added", "MW", "Modeling", "", fmt="#,##0", fn=lambda n: f"=MAX(0,{qc(n)}{mwe}-{qc(n-1)}{mwe})")
reg("Revenue", "MW_New", nmw)
add = {}; end = {}; avgg = {}; end0 = {}
def _endf(n, g):
    return (f"=INDEX(FileGPU_{g},1,MAX(1,{n}-IF({n}>LastHist,EnLag,0)))*IF(OR(FixMgmt=0,{n}>LastHist),IFERROR({qc(n)}{mwe}/INDEX(EnBaseIT,1,MAX(1,{n}-IF({n}>LastHist,EnLag,0))),1),IF(AND(HistMWBasis>=2,{n}>=LastHist-1),IFERROR((INDEX(EnBaseIT,1,{n})+MWUp)/INDEX(EnBaseIT,1,{n}),1),1))" + (f"+IF({n}>=5,GPUAnchorD,0)" if g == "H" else ""))
for g in GENS:
    end0[g] = R.row(f"GPU fleet end - build-plan basis (before StopBuild) - {GEN_NAME[g]}", "GPUs", "Capacity/Modeling", "", kind="stock", fmt="#,##0", e=(f"=INDEX(FileGPU_H,1,1)" if g == "H" else "=0"), fn=lambda n, g=g: _endf(n, g))
    reg("Revenue", f"GPU_End0_{g}", end0[g])
for g in GENS:
    add[g] = R.row(f"GPUs added - {GEN_NAME[g]}", "GPUs", "Modeling", "= change in fleet end (fleet from Capacity GPU file)", fmt="#,##0",
                   fn=lambda n, g=g: f"=MAX(0,{qc(n)}{R.r+3}-{qc(n-1)}{R.r+3})")
    reg("Revenue", f"GPU_Add_{g}", add[g])
for g in GENS:
    me = R.r
    end[g] = R.row(f"GPU fleet end (billing) - {GEN_NAME[g]}", "GPUs", "Capacity/Modeling", "Forecast: Capacity GPUs deployed at (n - EnLag) x driver MW / energized base IT MW at (n - EnLag) - same-quarter IT pairing; = file in history", kind="stock", fmt="#,##0",
                   e=(f"=INDEX(FileGPU_H,1,1)" if g == "H" else "=0"), fn=lambda n, g=g: f"=IF(AND(StopBuild=1,{n}>StopIdx),{X('A_GPU_'+g, n)}+MAX(0,INDEX(GPU_End0_{g},1,StopIdx)-INDEX(A_GPU_{g},1,StopIdx)),{qc(n)}{end0[g]})")
    reg("Revenue", f"GPU_End_{g}", end[g])
gt = R.row("GPU fleet end - total", "GPUs", "Modeling", "", kind="stock", fmt="#,##0", bold=True, e=f"=E{end['H']}+E{end['B']}+E{end['N']}",
           fn=lambda n: f"={qc(n)}{end['H']}+{qc(n)}{end['B']}+{qc(n)}{end['N']}")
reg("Revenue", "GPU_End", gt)
gta = R.row("GPU fleet average - total", "GPUs", "Modeling", "", kind="avg", fmt="#,##0", fn=lambda n: f"=AVERAGE({qc(n-1)}{gt},{qc(n)}{gt})")
reg("Revenue", "GPU_Avg", gta)
R.section("Book (a) - hyperscaler take-or-pay (from Contracts)")
hmw = R.row("Contract delivered IT MW", "MW", "Contracts", "", kind="stock", fmt="#,##0", fn=lambda n: f"={X('A_MW', n)}")
pmw = R.row("Post-term merchant IT MW", "MW", "Contracts", "", kind="stock", fmt="#,##0", fn=lambda n: f"={X('A_PostMW', n)}")
ba = R.row("Book (a) revenue incl. LDs", "$M", "Contracts", "Fixed $/MW-yr; insensitive to list prices", bold=True, fn=lambda n: f"={X('A_Rev', n)}+{X('A_LD', n)}")
reg("Revenue", "Book_A", ba)
bpost = R.row("Post-term merchant revenue (residual value stream)", "$M", "Contracts", "Decay-curve priced", fn=lambda n: f"={X('A_Post', n)}")
reg("Revenue", "Book_Post", bpost)
R.section("Non-contract fleet available to books (b) and (c)")
nh = {}; nha = {}
for g in GENS:
    nh[g] = R.row(f"Non-contract GPUs end - {GEN_NAME[g]}", "GPUs", "Modeling", "fleet - contract GPUs - post-term GPUs", kind="stock", fmt="#,##0",
                  e=f"=E{end[g]}", fn=lambda n, g=g: f"=MAX(0,{qc(n)}{end[g]}-{X('A_GPU_'+g, n)})")
for g in GENS: reg("Revenue", f"NCE_{g}", nh[g])
for g in GENS:
    nha[g] = R.row(f"Non-contract GPUs avg - {GEN_NAME[g]}", "GPUs", "Modeling", "", kind="avg", fmt="#,##0",
                   fn=lambda n, g=g: f"=AVERAGE({qc(n-1)}{nh[g]},{qc(n)}{nh[g]})")
    reg("Revenue", f"NCA_{g}", nha[g])
ovf = R.row("Contract MW in excess of active MW (check, should be 0)", "MW", "Modeling", "", kind="stock", fmt="#,##0",
            fn=lambda n: f"=MAX(0,{qc(n)}{hmw}+{qc(n)}{pmw}-{qc(n)}{mwe})")
reg("Revenue", "MW_Over", ovf)
R.section("Books (b) and (c): capacity revenue at 100% utilization x utilization (Q2'26 utilization calibrated to actual AI cloud revenue)")
bb = {}
for g in GENS:
    bb[g] = R.row(f"Book (b) @100% util - {GEN_NAME[g]}", "$M", "Modeling", "GPUs x hrs x ResSh x realized reserved price",
                  fn=lambda n, g=g: f"={qc(n)}{nha[g]}*{HRS(n)}*{ANN('ResSh', n)}*{X('P_RES_'+g, n)}/1000000")
bbr = R.row("Book (b) @100% utilization", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{bb['H']}+{qc(n)}{bb['B']}+{qc(n)}{bb['N']}")
bc = {}
for g in GENS:
    bc[g] = R.row(f"Book (c) @100% util - {GEN_NAME[g]}", "$M", "Modeling", "GPUs x hrs x (1-ResSh) x [(1-spot) x OD x RealOD + spot x Spot] x (1+StudioUp)",
                  fn=lambda n, g=g: (f"={qc(n)}{nha[g]}*{HRS(n)}*(1-{ANN('ResSh', n)})*((1-{ANN('SpotSh', n)})*{X('P_OD_'+g, n)}*IF({qc(n)}$5<=LastHist,RealOD_B,RealOD)"
                                     f"+{ANN('SpotSh', n)}*{X('P_SPOT_'+g, n)})*(1+{ANN('StudioUp', n)})/1000000"))
bcr = R.row("Book (c) @100% utilization", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{bc['H']}+{qc(n)}{bc['B']}+{qc(n)}{bc['N']}")
rr_ = R.r
R.ws[f"A{rr_}"] = "Implied Q2'26 (LastHist) utilization = (actual AI rev - book a - post-term - asset-light) / (b + c @100%)"; R.ws[f"A{rr_}"].font = BOLD
fml(R.ws[f"E{rr_}"], (f"=(INDEX(H_AIRev,1,LastHist)-INDEX(Book_A,1,LastHist)-INDEX(Book_Post,1,LastHist)-INDEX(AssetLight,1,INDEX(Q_Yr,1,LastHist)-FYBase)/4)"
                      f"/(INDEX($F${bbr}:$AG${bbr},1,LastHist)+INDEX($F${bcr}:$AG${bcr},1,LastHist))"), "0.0%", True)
name(wb, "UtilQ2", "Revenue", f"$E${rr_}")
fml(R.ws[f"F{rr_}"], '=IF(OR(UtilQ2>1,UtilQ2<0.4),"FLAG: outside 40-100% - not forced","OK (40-100%)")', "@", True)
name(wb, "UtilFlag", "Revenue", f"$F${rr_}")
R.r += 1
ub = R.row("Utilization used - book (b)", "%", "GPU Market", "History = UtilQ2 (calibrated); forecast = Util_Res input", kind="avg", fmt="0%",
           fn=lambda n: f"=IF({qc(n)}$5<=LastHist,UtilQ2,Util_Res)")
uc = R.row("Utilization used - book (c)", "%", "GPU Market", "History = UtilQ2; forecast = Util_OD input", kind="avg", fmt="0%",
           fn=lambda n: f"=IF({qc(n)}$5<=LastHist,UtilQ2,Util_OD)")
bbt = R.row("Book (b) revenue - reserved / commit", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{bbr}*{qc(n)}{ub}")
reg("Revenue", "Book_B", bbt)
bct = R.row("Book (c) revenue - on-demand / self-serve + spot + AI Studio", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{bcr}*{qc(n)}{uc}")
reg("Revenue", "Book_C", bct)
R.section("New-deal pricing (NewDealOn): 2027+ new non-contract MW priced per IT MW-yr (Q2'26 letter); legacy Q4'26 non-contract fleet stays on GPU-hour books (b)/(c)")
lg = {}
for g in GENS:
    lg[g] = R.row(f"Legacy non-contract GPUs avg (capped at NewDealIdx-1 level) - {GEN_NAME[g]}", "GPUs", "Modeling", "", kind="avg", fmt="#,##0",
                  fn=lambda n, g=g: f"=IF({qc(n)}$5<NewDealIdx,{qc(n)}{nha[g]},MIN({qc(n)}{nha[g]},INDEX($F${nha[g]}:$AG${nha[g]},1,NewDealIdx-1)))")
lgr = R.row("Legacy books (b)+(c) revenue (per-generation GPU share of book revenue)", "$M", "Modeling", "", 
            fn=lambda n: "=" + "+".join(f"IF({qc(n)}{nha[g]}>0,{qc(n)}{lg[g]}/{qc(n)}{nha[g]},0)*({qc(n)}{bb[g]}*{qc(n)}{ub}+{qc(n)}{bc[g]}*{qc(n)}{uc})" for g in GENS))
ndm = R.row("New 2027+ non-contract IT MW (avg) = (non-contract GPUs - legacy) / GPUs per MW", "MW", "Modeling", "", kind="avg", fmt="#,##0",
            fn=lambda n: "=" + "+".join(f"({qc(n)}{nha[g]}-{qc(n)}{lg[g]})/GPUMW_{g}" for g in GENS))
reg("Revenue", "ND_MW", ndm)
_ndg = {}
for _g in ["BH", "N"]:
    _gs = ["B", "H"] if _g == "BH" else ["N"]
    _ndg[_g] = R.row(f"New 2027+ non-contract IT MW (avg) - {'GB300/Blackwell (+Hopper)' if _g == 'BH' else 'Vera Rubin'}", "MW", "Modeling", "", kind="avg", fmt="#,##0",
                     fn=lambda n, _gs=_gs: "=" + "+".join(f"({qc(n)}{nha[g]}-{qc(n)}{lg[g]})/GPUMW_{g}" for g in _gs))
_nag = {}
for _g in ["BH", "N"]:
    _nag[_g] = R.row(f"New MW added (vintage) - {'GB300' if _g == 'BH' else 'Rubin'}", "MW", "Modeling", "", fmt="#,##0", fn=lambda n, _g=_g: f"=MAX(0,{qc(n)}{_ndg[_g]}-{qc(n-1)}{_ndg[_g]})")
nda = R.row("New 2027+ MW added (vintage, both generations)", "MW", "Modeling", "", fmt="#,##0", fn=lambda n: f"={qc(n)}{_nag['BH']}+{qc(n)}{_nag['N']}")
_pf = lambda n: f"(1+MYDecl)^(MAX(0,{qc(n)}$5-(MYDeclIdx-1))/4)*IF(AND(MYStressOn=1,{qc(n)}$5>=PxCycIdx),1-MYStressCut,1)"
myb = R.row("Multi-year GB300 ACV by signing vintage ($M/IT MW-yr): v0.6 = MYRate x basis x flat-then-MYDecl path; v0.5 = x Blackwell OD list(v)/list(Q4'26)", "$M/IT MW-yr", "Filings/Modeling", "fixed for contract life", kind="avg", fmt="0.0",
            fn=lambda n: f"=MYRate*IF(NDBasis=2,ND_GrossX,1)*IF(MYPathMode=1,{_pf(n)},{X('P_OD_B', n)}/INDEX(P_OD_B,1,NewDealIdx-1))")
reg("Revenue", "ND_RateGB", myb)
mynr = R.row("Multi-year Rubin ACV by signing vintage: RubinMode 1 = RubinMW x path (direct, no IT/gross toggle); 0 = GB300 x Rubin_Prem x GPUMW_N / ND_GPUMW_GB", "$M/IT MW-yr", "GPU Market", "only generation effect; no extra $/MW decline stacked", kind="avg", fmt="0.0",
             fn=lambda n: f"=IF(RubinMode=1,RubinMW*IF(MYPathMode=1,{_pf(n)},1),{qc(n)}{myb}*IF(ND_GenOn=1,Rubin_Prem*GPUMW_N/ND_GPUMW_GB,1))")
reg("Revenue", "ND_RateRubin", mynr)
myr = R.row("Multi-year ACV by vintage - blended by generation mix of vintage adds", "$M/IT MW-yr", "Modeling", "", kind="avg", fmt="0.0",
            fn=lambda n: f"=IF({qc(n)}{nda}>0,({qc(n)}{_nag['BH']}*{qc(n)}{myb}+{qc(n)}{_nag['N']}*{qc(n)}{mynr})/{qc(n)}{nda},{qc(n)}{myb})")
reg("Revenue", "ND_MYRate", myr)
stp = R.row("Short-term tranche price index (post-Rubin cut PxCyc for 3 qtrs, then linear recovery over 2 qtrs if STRecov; x annual decline after; reprices every 6 months)", "x", "GPU Market", "", kind="avg", fmt="0.000",
            fn=lambda n: (lambda m: f"=IF({m}<PxCycIdx,1,IF({m}<PxCycIdx+3,1-PxCyc,IF(STRecov=1,IF({m}=PxCycIdx+3,1-PxCyc/2,1),1-PxCyc)))*(1+PxChg_B)^(MAX(0,{m}-(PxCycIdx+5))/4)")(f"({qc(n)}$5-MOD({qc(n)}$5-NewDealIdx,2))"))
str_ = R.row("Short-term tranche $M/IT MW-yr", "$M/IT MW-yr", "Filings", "", kind="avg", fmt="0.0", fn=lambda n: f"=STRate*{qc(n)}{stp}")
reg("Revenue", "ND_STRate", str_)
lgm = R.row("Legacy non-contract IT MW (avg)", "MW", "Modeling", "", kind="avg", fmt="#,##0", fn=lambda n: "=" + "+".join(f"{qc(n)}{lg[g]}/GPUMW_{g}" for g in GENS))
stm = R.row("Short-term tranche IT MW = STShare x TOTAL average active MW (from NewDealIdx), carved out of legacy + new non-contract MW", "MW", "GPU Market", "v0.5: share of total capacity, toggled on capacity (not revenue)", kind="avg", fmt="#,##0",
            fn=lambda n: f"=IF({qc(n)}$5<NewDealIdx,0,IF(STMode=1,MIN(STShare*{qc(n)}{mwa},{qc(n)}{lgm}+IF(STGBOnly=1,{qc(n)}{_ndg['BH']},{qc(n)}{ndm})),STShare*{qc(n)}{ndm}))")
reg("Revenue", "ST_MW", stm)
stf = R.row("Carve-out factor = short-term MW / (legacy + new non-contract MW; GB300-era only if STGBOnly)", "%", "Modeling", "", kind="avg", fmt="0.0%",
            fn=lambda n: (lambda d: f"=IF({d}>0,{qc(n)}{stm}/({d}),0)")(f"{qc(n)}{lgm}+IF(STGBOnly=1,{qc(n)}{_ndg['BH']},{qc(n)}{ndm})"))
myc = R.row("Multi-year revenue at contracted rate (all new MW at MYRate path; used for RPO sold/unsold split)", "$M", "Modeling", "",
            fn=lambda n: f"=SUMPRODUCT($F${nda}:{qc(n)}{nda},$F${myr}:{qc(n)}{myr})/4")
R.section("Contracted vs unsold new MW from RPO (Q2'26 Note 2: RPO $37.49B, 36% recognised Jul'26-Jun'28) - Capacity site tags due Tuesday")
_w = lambda rr: f"SUMPRODUCT(($F$5:$AG$5>=LastHist+1)*($F$5:$AG$5<=LastHist+8)*$F${rr}:$AG${rr})"
_rs = R.r
R.row("Window revenue Jul'26-Jun'28 already covered ex new MW: book (a) + post-term + legacy (b)+(c) (memo, col F)", "$M", "Modeling", "", kind="none",
      fn=lambda n: (f"={_w(ba)}+{_w(bpost)}+SUMPRODUCT(($F$5:$AG$5>=LastHist+1)*($F$5:$AG$5<=LastHist+8)*($F$5:$AG$5<NewDealIdx)*($F${bbt}:$AG${bbt}+$F${bct}:$AG${bct}))+SUMPRODUCT(($F$5:$AG$5>=NewDealIdx)*($F$5:$AG$5<=LastHist+8)*$F${lgr}:$AG${lgr})" if n == 1 else None))
name(wb, "RPO_Cov", "Revenue", f"$F${_rs}")
R.row("RPO backlog recognised in window remaining for new MW = MAX(0, RPO24 - covered)", "$M", "Filings", "", kind="none", fn=lambda n: ("=MAX(0,RPO_Tot*RPO24Pct-RPO_Cov)" if n == 1 else None))
name(wb, "RPO_New", "Revenue", f"$F${_rs+1}")
R.row("Sold share of new 2027+ MW through Q2'28 = MIN(1, RPO for new MW / window multi-year revenue at contracted rate)", "%", "Modeling", "", kind="none", fmt="0.0%",
      fn=lambda n: (f"=IFERROR(MIN(1,RPO_New/{_w(myc)}),1)" if n == 1 else None))
name(wb, "RPO_Sold", "Revenue", f"$F${_rs+2}")
R.row("Contracted (sold) new MW = sold share x new MW at Q2'28; vintages filled first-in", "MW", "Modeling", "", kind="none", fmt="#,##0",
      fn=lambda n: (f"=RPO_Sold*INDEX(ND_MW,1,LastHist+8)" if n == 1 else None))
name(wb, "RPO_SoldMW", "Revenue", f"$F${_rs+3}")
ndc = R.row("Cumulative new MW added (vintage)", "MW", "Modeling", "", kind="stock", fmt="#,##0", fn=lambda n: f"={qc(n-1)}{R.r}+{qc(n)}{nda}")
nds = R.row("Sold (contracted) portion of vintage", "MW", "Modeling", "", fmt="#,##0.0", fn=lambda n: f"=MIN({qc(n)}{nda},MAX(0,RPO_SoldMW-{qc(n-1)}{ndc}))")
sfac = R.row("Vintage sold/unsold weight = (sold + unsold x MYRateU/MYRate) / vintage MW", "x", "Modeling", "", kind="avg", fmt="0.000",
             fn=lambda n: f"=IF({qc(n)}{nda}>0,({qc(n)}{nds}+({qc(n)}{nda}-{qc(n)}{nds})*MYRateU/MYRate)/{qc(n)}{nda},0)")
_cf = lambda n: f"((UncCohort=0)+(UncCohort=1)*({n}<=16)+(UncCohort=2)*({n}>16))"
ssh = R.row("v0.8b: sold (RPO-contracted) share of vintage", "%", "Modeling", "", kind="avg", fmt="0.0%", fn=lambda n: f"=IF({qc(n)}{nda}>0,{qc(n)}{nds}/{qc(n)}{nda},0)")
ubh = R.row("v0.8b: UNSOLD GB300 vintage rate ($M/IT MW-yr): UncPx x path if UncPxOn, else ACV x MYRateU/MYRate", "$M/IT MW-yr", "Skeptic", "", kind="avg", fmt="0.0", fn=lambda n: f"=IF(UncPxOn=1,IF({_cf(n)}=1,UncPx,UncPxX)*{_pf(n)},IF(UncSplitOn=1,IF({_cf(n)}=1,UncGBPx,UncPxX)*{_pf(n)},{qc(n)}{myb}*MYRateU/MYRate))")
unr = R.row("v0.8b: UNSOLD Rubin vintage rate ($M/IT MW-yr): UncPx x path if UncPxOn, else Rubin ACV x MYRateU/MYRate", "$M/IT MW-yr", "Skeptic", "", kind="avg", fmt="0.0", fn=lambda n: f"=IF(UncPxOn=1,IF({_cf(n)}=1,UncPx,UncPxX)*{_pf(n)},IF(UncSplitOn=1,UncRubPx*{_pf(n)},{qc(n)}{mynr}*MYRateU/MYRate))")
reg("Revenue", "ND_UncGB", ubh); reg("Revenue", "ND_UncRubin", unr)
mrv = R.row("Multi-year tranche revenue = GB300 vintages x (1 - carve-out) + Rubin vintages x (1 - carve-out unless STGBOnly), each x ACV x sold/unsold weight /4", "$M", "Modeling", "take-or-pay, no utilization haircut; unsold MW repriced at MYRateU (headline sens $18/22/25)",
            fn=lambda n: (f"=IF(STMode=1,1-{qc(n)}{stf},1-STShare)*(SUMPRODUCT($F${myb}:{qc(n)}{myb},$F${_nag['BH']}:{qc(n)}{_nag['BH']},$F${ssh}:{qc(n)}{ssh})+SUMPRODUCT($F${ubh}:{qc(n)}{ubh},$F${_nag['BH']}:{qc(n)}{_nag['BH']},1-$F${ssh}:{qc(n)}{ssh}))/4"
                          f"+IF(STMode=1,IF(STGBOnly=1,1,1-{qc(n)}{stf}),1-STShare)*(SUMPRODUCT($F${mynr}:{qc(n)}{mynr},$F${_nag['N']}:{qc(n)}{_nag['N']},$F${ssh}:{qc(n)}{ssh})+SUMPRODUCT($F${unr}:{qc(n)}{unr},$F${_nag['N']}:{qc(n)}{_nag['N']},1-$F${ssh}:{qc(n)}{ssh}))/4"))
rnr = R.row("GB300 vintage rate at year-3 renewal (prevailing curve at vintage + 12 qtrs; beyond Q4'30 extrapolated at MYDecl)", "$M/IT MW-yr", "Contracts", "ThreeYrOn option", kind="avg", fmt="0.0",
            fn=lambda n: f"=IF({n}+12<=28,INDEX($F${myb}:$AG${myb},1,MIN(28,{n}+12)),{qc(n)}{myb}*(1+MYDecl)^3)")
rpa = R.row("3-yr deal re-pricing adjustment = ThreeYrSh x GB300 vintages renewed (age >= 12 qtrs) x (renewal rate - original) /4", "$M", "Contracts", "bear option",
            fn=lambda n: (f"=ThreeYrOn*ThreeYrSh*IF(STMode=1,1-{qc(n)}{stf},1-STShare)*SUMPRODUCT(($F$5:{qc(n)}$5<={qc(n)}$5-12)*$F${_nag['BH']}:{qc(n)}{_nag['BH']}*$F${sfac}:{qc(n)}{sfac}*($F${rnr}:{qc(n)}{rnr}-$F${myb}:{qc(n)}{myb}))/4"))
reg("Revenue", "ND_3yAdj", rpa)
rnn = R.row("v0.8b: Rubin vintage rate at year-3 renewal (prevailing Rubin curve at vintage + 12 qtrs)", "$M/IT MW-yr", "Contracts", "", kind="avg", fmt="0.0",
            fn=lambda n: f"=IF({n}+12<=28,INDEX($F${mynr}:$AG${mynr},1,MIN(28,{n}+12)),{qc(n)}{mynr}*(1+MYDecl)^3)")
_pfm = lambda m: f"(1+MYDecl)^(MAX(0,{m}-(MYDeclIdx-1))/4)*IF(AND(MYStressOn=1,{m}>=PxCycIdx),1-MYStressCut,1)"
rsg = R.row("v0.8d: LEASED GB300 renewal rate = UncGBRen x price path at renewal (vintage + 4 x term)", "$M/IT MW-yr", "Skeptic", "", kind="avg", fmt="0.0",
            fn=lambda n: f"=UncGBRen*{_pfm(f'({n}+4*UncGBTermY)')}")
dgb = R.row("v0.8d: unsold GB300 renewal step ($M/IT MW-yr): UncPxOn -> prevailing curve - UncPx; split -> UncGBRen path - UncGBPx; else 0", "$M/IT MW-yr", "Skeptic", "", kind="avg", fmt="0.0;-0.0",
            fn=lambda n: f"=IF(UncPxOn=1,{qc(n)}{rnr}-{qc(n)}{ubh},IF(UncSplitOn=1,{qc(n)}{rsg}-{qc(n)}{ubh},0))")
dnn = R.row("v0.8d: unsold Rubin renewal step: UncPxOn -> prevailing Rubin curve - UncPx at 12 qtrs; split -> 0 (5-yr, renews flat)", "$M/IT MW-yr", "Skeptic", "", kind="avg", fmt="0.0;-0.0",
            fn=lambda n: f"=IF(UncPxOn=1,{qc(n)}{rnn}-{qc(n)}{unr},0)")
reg("Revenue", "ND_NagBH", _nag['BH']); reg("Revenue", "ND_NagN", _nag['N']); reg("Revenue", "ND_SSh", ssh); reg("Revenue", "ND_dGB", dgb); reg("Revenue", "ND_dN", dnn); reg("Revenue", "ND_STF", stf)
ren = R.row("v0.8b/d: unsold re-pricing at renewal = UNSOLD vintages aged >= renewal age (GB300 UncRenAge, Rubin 12 under UncPxOn) x renewal step / 4", "$M", "Skeptic", "UncRenew",
            fn=lambda n: (f"=UncRenew*(IF(STMode=1,1-{qc(n)}{stf},1-STShare)*SUMPRODUCT(($F$5:{qc(n)}$5<={qc(n)}$5-UncRenAge)*$F${_nag['BH']}:{qc(n)}{_nag['BH']}*(1-$F${ssh}:{qc(n)}{ssh})*$F${dgb}:{qc(n)}{dgb})"
                          f"+IF(STMode=1,IF(STGBOnly=1,1,1-{qc(n)}{stf}),1-STShare)*SUMPRODUCT(($F$5:{qc(n)}$5<={qc(n)}$5-12)*$F${_nag['N']}:{qc(n)}{_nag['N']}*(1-$F${ssh}:{qc(n)}{ssh})*$F${dnn}:{qc(n)}{dnn}))/4"))
reg("Revenue", "ND_UncRen", ren)
srv = R.row("Short-term tranche revenue = short-term MW x STRate x price index x STOcc /4", "$M", "Modeling", "STOcc = occupancy between deals (Estimated 85%): $45M x 85% ~ $38M/IT MW-yr", fn=lambda n: f"={qc(n)}{stm}*{qc(n)}{str_}*IF(STMode=1,STOcc,1)/4")
nadj = R.row("New-deal adjustment = IF(NewDealOn, legacy b+c + multi-year + short-term - (b+c GPU-hour engine), 0)", "$M", "Modeling", "", bold=True,
             fn=lambda n: f"=IF(AND(NewDealOn=1,{qc(n)}$5>=NewDealIdx),{qc(n)}{lgr}*IF(STMode=1,1-{qc(n)}{stf},1)+{qc(n)}{mrv}+{qc(n)}{rpa}+{qc(n)}{ren}+{qc(n)}{srv}-{qc(n)}{bbt}-{qc(n)}{bct},0)")
reg("Revenue", "ND_Adj", nadj)
R.row("Memo: multi-year new-MW revenue per IT MW-yr (before carve-out)", "$M/IT MW-yr", "Modeling", "", kind="avg", fmt="0.0",
      fn=lambda n: f"=IF({qc(n)}{ndm}>0,{qc(n)}{myc}*4/{qc(n)}{ndm},0)")
reg("Revenue", "ND_RevMW", R.r - 1)
R.section("Revenue roll-up")
al = R.row("Asset-light partner revenue", "$M", "GPU Market", "", fn=lambda n: f"={ANN('AssetLight', n)}/4")
aim0 = R.row("AI cloud revenue - model before June-exit overlay", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{ba}+{qc(n)}{bpost}+{qc(n)}{bbt}+{qc(n)}{bct}+{qc(n)}{nadj}+{qc(n)}{al}")
jx = R.row("June exit run-rate overlay (JuneExitOn): Q3'26 >= ARR/4 before new MW, held flat", "$M", "Filings", "MAX(0, ARR/4 - model Q3'26 AI revenue); not tied to MW; held flat in $", 
           fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,JuneExitOn*MAX(0,INDEX(H_ARR,1,LastHist)/4-INDEX($E${aim0}:$AG${aim0},1,LastHist+2)))")
reg("Revenue", "JuneOv", jx)
aim = R.row("AI cloud revenue - model", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{aim0}+{qc(n)}{jx}")
reg("Revenue", "AI_Model", aim)
oth = R.row("Other businesses revenue (Avride, TripleTen, elims)", "$M", "Filings", "History = group - AI cloud (Hist)",
            fn=lambda n: (f"={X('H_Rev', n)}-{X('H_AIRev', n)}" if n <= NHIST else None))
for n in range(1, NQ + 1):
    c = qc(n)
    R.ws[f"{c}{oth}"] = f"=IF({c}$5<=LastHist,INDEX(H_Rev,1,{c}$5)-INDEX(H_AIRev,1,{c}$5),{qc(n-1)}{oth}*(1+OthRev_g))" if n > 1 else f"=INDEX(H_Rev,1,1)-INDEX(H_AIRev,1,1)"
grm = R.row("Group revenue - model", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{aim}+{qc(n)}{oth}")
reg("Revenue", "Rev_Model", grm)
gru = R.row("GROUP REVENUE (used: actual for history, model for forecast)", "$M", "Modeling", "", bold=True,
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_Rev,1,{qc(n)}$5),{qc(n)}{grm})")
reg("Revenue", "Rev", gru)
aiu = R.row("AI cloud revenue (used)", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_AIRev,1,{qc(n)}$5),{qc(n)}{aim})")
reg("Revenue", "AIRev", aiu)
R.row("Model vs actual AI cloud revenue (history calibration)", "%", "Modeling", "model / actual - 1 (history only)", kind="none", fmt="0%;-0%",
      fn=lambda n: f"=IF({qc(n)}$5<=LastHist,{qc(n)}{aim}/INDEX(H_AIRev,1,{qc(n)}$5)-1,\"\")")
reg("Revenue", "Calib", R.r - 1)
R.section("Book mix & list-price exposure")
for lab, rr in [("Book (a) share of AI cloud model revenue", ba), ("Book (b) share", bbt), ("Book (c) share", bct), ("Post-term merchant share", bpost)]:
    R.row(lab, "%", "Modeling", "", kind="none", fmt="0%", fn=lambda n, rr=rr: f"=IF({qc(n)}{aim}<>0,{qc(n)}{rr}/{qc(n)}{aim},0)",
          fy_fn=lambda y, c, rr=rr: f"=IF({c}{aim}<>0,{c}{rr}/{c}{aim},0)")
lps = R.row("List-price-sensitive revenue (b + c)", "$M", "Modeling", "Sensitivity base", fn=lambda n: f"={qc(n)}{bbt}+{qc(n)}{bct}")
reg("Revenue", "Rev_PxSens", lps)
R.section("ARR check at 6/30 - JUNE RUN-RATE basis (June list + renewals signed by 6/30; NOT H2 average or Oct-1 prices)")
def _ARR(label, f, fmt="#,##0", nm=None, note=""):
    rr = R.r; R.ws[f"A{rr}"] = label; fml(R.ws[f"E{rr}"], f, fmt, True); R.ws[f"D{rr}"] = note; R.ws[f"D{rr}"].font = ITAL
    if nm: name(wb, nm, "Revenue", f"$E${rr}")
    R.r += 1
_q2, _q3 = qc(10), qc(11)
_t1, _t2 = TRROWS["rev"][0], TRROWS["rev"][1]
_ARR("Reported ARR at 6/30/26 ($M)", "=INDEX(H_ARR,1,LastHist)", nm="ARR_Rep", note="Q2 letter; mgmt definition per Hist note = last-month AI cloud revenue x 12 (scope/segment pending Filings)")
_ARR("Model take-or-pay June run-rate x 12", f"=4*((Contracts!{_q2}{_t1}+Contracts!{_q3}{_t1}+Contracts!{_q2}{_t2}+Contracts!{_q3}{_t2})/2+Contracts!{_q2}{orv})", nm="ARR_RRa",
     note="Microsoft T1-T2 at end-June ramp point (avg of Q2, Q3 quarterly) + Meta #1 Q2; excludes T3-T4 (delivered Jul-Aug)")
_endavg = f"({qc(10)}{nh['H']}+{qc(10)}{nh['B']}+{qc(10)}{nh['N']})/({qc(10)}{nha['H']}+{qc(10)}{nha['B']}+{qc(10)}{nha['N']})"
_ARR("Model reserved + on-demand June run-rate x 12", f"=4*({qc(10)}{bbt}+{qc(10)}{bct}+{qc(10)}{bc['H']}*{qc(10)}{uc}*(ODp_H/ODq_H-1))*{_endavg}", nm="ARR_RRbc",
     note="Q2 books (b)+(c) x end/avg fleet; Hopper on-demand at June list ($3.85 H100); reserved at contracted (signed) rates")
_ARR("Model June run-rate (annualized)", "=ARR_RRa+ARR_RRbc", nm="ARR_RR")
_ARR("  per energized IT MW at 6/30 ($M/MW-yr; team ~13.3-13.5)", "=ARR_RR/INDEX(MW_En,1,LastHist)", fmt="0.0", nm="ARR_RRMW")
_ARR("UNEXPLAINED ARR RESIDUAL (reported ARR - model June run-rate) - NOT unseen capacity", "=ARR_Rep-ARR_RR", nm="ARR_Res",
     note="Pending Filings: ARR segment scope; does ARR include signed-not-billing tranches? Evidence line only for Oklahoma/Minnesota; NOT plugged")
_ARR("  MW-equivalent at model June run-rate per MW", "=IFERROR(ARR_Res/ARR_RRMW,0)", fmt="#,##0", nm="ARR_ResMW", note="team estimate ~$0.65-0.7B / ~45-50 MW")
_ARR("June 2026 AI cloud revenue (monthly) = ARR / 12 (Filings: ARR = June recognized x 12, no signed-not-billing)", "=ARR_Rep/12", nm="ARR_JunMo")
_ARR("Q2'26 average monthly AI cloud revenue", "=INDEX(H_AIRev,1,LastHist)/3", nm="ARR_Q2Mo")
_ARR("June step-up vs March (monthly, $M)", "=ARR_JunMo-MarMo", nm="ARR_Step", note="INFO: ~$88M/mo")
_ARR("  = IT MW at take-or-pay pricing (step x 12 / ACV_2026 $12M/MW) - UPPER END; fewer MW if reserved at Q2 renewal rates", "=ARR_Step*12/ACV_2026", fmt="#,##0", nm="ARR_StepMW", note="INFO only; not plugged. Toggle JuneExitOn tests Q3 from June exit")
_ARR("Cross-check: residual at team June basis = ARR - 6/30 MW x $13.5M/MW (low)", "=ARR_Rep-INDEX(MW_En,1,LastHist)*ARR_RRMW_Hi", nm="ARR_ResTB_Lo", note="team basis ~$13.3-13.5M/MW (June list + renewals signed by 6/30)")
_ARR("Cross-check: residual at team June basis = ARR - 6/30 MW x $13.3M/MW (high)", "=ARR_Rep-INDEX(MW_En,1,LastHist)*ARR_RRMW_Lo", nm="ARR_ResTB_Hi")
_ARR("  Model residual MW-equivalent at team $13.3-13.5M/MW (low end, /13.5)", "=ARR_Res/ARR_RRMW_Hi", fmt="#,##0", nm="ARR_ResMW_TLo")
_ARR("  Model residual MW-equivalent at team $13.3-13.5M/MW (high end, /13.3)", "=ARR_Res/ARR_RRMW_Lo", fmt="#,##0", nm="ARR_ResMW_THi", note="Not used in revenue, capacity or valuation (no plug)")

# ------------------------------------------------------------------ CAPEX & D&A
cx = TS(wb, "Capex_DA", "Capex by asset class (IT/GPU by generation vintage, shell, power/electrical, maintenance) and depreciation waterfalls by vintage cohort with per-vintage life and optional impairment.")
C = cx; cws2 = cx.ws
C.section("Life & impairment parameters (live)")
LQ = {}; IMPE = {}
for g in GENS:
    rr = C.r
    cws2[f"A{rr}"] = f"Depreciation life used - {GEN_NAME[g]} (quarters)"
    fml(cws2[f"F{rr}"], f"=4*IF(LifeOverride=0,Life_{g},LifeOverride)", "0", True); name(wb, f"LQ_{g}", "Capex_DA", f"$F${rr}")
    cws2[f"G{rr}"] = "Impairment qtr (999 = none)"; fml(cws2[f"K{rr}"], f"=IF(ImpQ_{g}=0,999,ImpQ_{g})", "0", True); name(wb, f"IMPE_{g}", "Capex_DA", f"$K${rr}")
    C.r += 1
for nm, lab, src in [("LQ_CapInt", "Capitalized interest life (quarters)", "=4*CapIntLife"), ("LQ_Shell", "Shell life (quarters)", "=4*Life_Shell"), ("LQ_Pow", "Power/electrical life (quarters)", "=4*Life_Pow"), ("LQ_Maint", "Maintenance capex life (quarters)", "=4*Life_Maint")]:
    rr = C.r; cws2[f"A{rr}"] = lab; fml(cws2[f"F{rr}"], src, "0", True); name(wb, nm, "Capex_DA", f"$F${rr}"); C.r += 1
C.section("Capex build")
newmw = C.row("New active IT MW (Revenue)", "MW", "Revenue", "", fmt="#,##0", fn=lambda n: f"={X('MW_New', n)}")
own = C.row("Owned share of active IT MW (Sites)", "%", "Sites", "", kind="avg", fmt="0%", fn=lambda n: f"={X('MW_OwnSh', n)}")
hcap = C.row("Historical capex (Hist, used)", "$M", "Filings", "", fn=lambda n: (f"={X('H_Capex', n)}" if n <= NHIST else None))
capg = {}
for g in GENS:
    def fcap(n, g=g):
        c = qc(n)
        lead = f"MIN(28,{c}$5+GPULead)"
        m1 = f"INDEX(GPU_Add_{g},1,{lead})*GPUCost_{g}/1000000"
        m2 = f"INDEX(GPU_Add_{g},1,{lead})/GPUMW_{g}*ITCapexMW"
        hist = f"INDEX(H_Capex,1,{c}$5)*HistGPUsh*{ANN('Mix_'+g, n)}"
        sh_ = f"IFERROR(INDEX(GPU_Add_{g},1,{lead})/(INDEX(GPU_Add_H,1,{lead})+INDEX(GPU_Add_B,1,{lead})+INDEX(GPU_Add_N,1,{lead})),{ANN('Mix_'+g, n)})"
        _ix = f"MAX(1,MIN(28,{c}$5+GPULead-EnLag))"
        _mwn = f"IF({c}$5+GPULead>29,0,IF({c}$5+GPULead>28,IF(TermBase=2,0,MW_Add31),INDEX(MW_New,1,{lead})))"
        _cu = f"IF({c}$5=LastHist+1,MAX(0,INDEX(MW_Act,1,LastHist)-INDEX(EnBaseIT,1,LastHist))*INDEX(PathRateIT,1,LastHist+1),0)"
        m3 = f"IF(FixMgmt=1,({_mwn}*INDEX(PathRateIT,1,{_ix})+{_cu}),INDEX(PathIT,1,{_ix}))*{sh_}"
        _bsx = (f"+StopBuild*IF({c}$5+GPULead>28,0,MAX(0,INDEX(BS_MW,1,{lead})-IF({lead}>1,INDEX(BS_MW,1,MAX(1,{lead}-1)),0)))*MAX(0,CPM_Own_VR-INDEX(PathRateIT,1,{_ix})-INDEX(PathRateSP,1,{_ix}))" if g == "N" else "")
        return f"=IF({c}$5<=LastHist,IF({c}$5<=10,{hist},0),IF(AND(CapexMethod<3,{c}$5+GPULead>28),0,CHOOSE(CapexMethod,{m1},{m2},{m3}){_bsx}))"
    capg[g] = C.row(f"IT/GPU capex - {GEN_NAME[g]} vintage", "$M", "Modeling", "Forecast: GPUs added (lead GPULead) x $/GPU (method 1) or $M/IT MW (method 2)",
                    e=("=OpenGPU_H" if g == "H" else "=0"), fn=fcap)
    reg("Capex_DA", f"CapGPU_{g}", capg[g])
def fdc(n, share, permw):
    c = qc(n)
    lead = f"MIN(28,{c}$5+DCLead)"
    _ix = f"MAX(1,MIN(28,{c}$5+DCLead-EnLag))"
    _mwn = f"IF({c}$5+DCLead>29,0,IF({c}$5+DCLead>28,IF(TermBase=2,0,MW_Add31),INDEX(MW_New,1,{lead})))"
    _cu = f"IF({c}$5=LastHist+1,MAX(0,INDEX(MW_Act,1,LastHist)-INDEX(EnBaseIT,1,LastHist))*INDEX(PathRateSP,1,LastHist+1),0)"
    m3 = f"IF(FixMgmt=1,({_mwn}*INDEX(PathRateSP,1,{_ix})+{_cu}),INDEX(PathSP,1,{_ix}))*{permw}/(ShellMW+PowerMW)"
    return f"=IF({c}$5<=LastHist,INDEX(H_Capex,1,{c}$5)*{share},IF(CapexMethod=3,{m3},INDEX(MW_New,1,{lead})*INDEX(MW_OwnSh,1,{lead})*{permw}))"
caps = C.row("Shell/building capex (owned sites only)", "$M", "Modeling", "New MW (lead DCLead) x owned share x ShellMW", e="=0", fn=lambda n: fdc(n, "HistShellSh", "ShellMW"))
capp = C.row("Power/electrical capex (owned sites only)", "$M", "Modeling", "New MW (lead DCLead) x owned share x PowerMW", e="=0", fn=lambda n: fdc(n, "HistPowSh", "PowerMW"))
capm = C.row("Maintenance/other capex", "$M", "Modeling", "MaintPct x revenue (forecast)", e="=0", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,MaintPct*{X('Rev', n)})")
reg("Capex_DA", "CapMaint", capm)
caplp = C.row("Lease prepayments on new leased IT MW (LeasePrepay; at shell/power timing)", "$M", "Filings", "upfront cost; amortised over Life_Pow; not replaced",
              fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,(1-StopBuild)*LeasePrepayOn*LeasePrepay*IF({qc(n)}$5+DCLead-EnLag>LastHist,INDEX(EnNewL,1,MAX(1,MIN(28,{qc(n)}$5+DCLead-EnLag))),0))")
reg("Capex_DA", "CapLPre", caplp)
def frep(n):
    c = qc(n)
    parts = [f"IF({c}$5-LQ_{g}>=0,INDEX($E${capg[g]}:$AG${capg[g]},1,{c}$5-LQ_{g}+1),0)*IF(ReplPxOn=1,(1+PxChg_{g})^(LQ_{g}/4),1)" for g in GENS]
    return f"=IF({c}$5<=LastHist,0,ReplOn*(" + "+".join(parts) + "))"
caprp = C.row("GPU end-of-life replacement capex (cohort cost at GPU life)", "$M", "Modeling", "ReplOn; replaced at original cost; depreciated at next-gen life", e="=0", fn=frep)
reg("Capex_DA", "CapRepl", caprp)
crs = C.row("GPU residual / resale proceeds at refresh, after tax (GPU_Resid x replaced cohort cost x (1 - TaxRate); NBV 0)", "$M", "GPU Market", "v0.7b: retired fleet was written to zero with no resale; revenue continues on the replacement GPUs",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,GPU_Resid*{qc(n)}{caprp}*(1-TaxRate))")
reg("Capex_DA", "CapResid", crs)
pth = C.row("Path-consistent capex (GPU + shell + power + maintenance + replacement + lease prepayments)", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{capg['H']}+{qc(n)}{capg['B']}+{qc(n)}{capg['N']}+{qc(n)}{caps}+{qc(n)}{capp}+{qc(n)}{capm}+{qc(n)}{caprp}+{qc(n)}{caplp}")
cov = C.row("Consensus capex overlay (IT cohort; CapexMode = 1, FY27 only)", "$M", "Skeptic", "= Cons FY27 capex / 4 - path capex; depreciated at Blackwell life",
            fn=lambda n: f"=IF(AND(CapexMode=1,{qc(n)}$7=2027,{qc(n)}$5>LastHist),INDEX(Cons_Cx,1,2027-FYBase)/4-{qc(n)}{pth},0)")
capt = C.row("TOTAL CAPEX", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{pth}+{qc(n)}{cov}")
reg("Capex_DA", "Capex", capt)
cnew = C.row("Memo: new-build capex (GPU + shell + power) - prepayment base", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{capg['H']}+{qc(n)}{capg['B']}+{qc(n)}{capg['N']}+{qc(n)}{caps}+{qc(n)}{capp}")
reg("Capex_DA", "CapNew", cnew)
C.row("Check: history capex split = reported (should be 0)", "$M", "Modeling", "", kind="flow",
      fn=lambda n: (f"={qc(n)}{capt}-{qc(n)}{hcap}" if n <= NHIST else None))
reg("Capex_DA", "CapexChk", C.r - 1)
C.section("GPU depreciation waterfall by vintage cohort (straight-line from PISLag quarters after spend; impairment writes down NBV of cohorts placed up to the impairment quarter)")
dep = {}; imp = {}; nbv = {}
for g in GENS:
    cr = capg[g]
    def fa(n, cr=cr, g=g, sign="<="):
        c = qc(n)
        _pl = f"(PISLag+($E$5:{c}$5>LastHist)*(GPULead-1))"
        return (f"=SUMPRODUCT($E${cr}:{c}${cr},(({c}$5-$E$5:{c}$5)>={_pl})*(({c}$5-$E$5:{c}$5)<LQ_{g}+{_pl})*($E$5:{c}$5{sign}IMPE_{g}))/LQ_{g}")
    ra_ = C.row(f"{GEN_NAME[g]}: dep. of cohorts placed <= impairment qtr (pre-write-down)", "$M", "Modeling", "", fn=fa)
    rb_ = C.row(f"{GEN_NAME[g]}: dep. of cohorts placed after impairment qtr", "$M", "Modeling", "", fn=lambda n, fa=fa: fa(n, sign=">"))
    dep[g] = C.row(f"{GEN_NAME[g]}: GPU depreciation", "$M", "Modeling", "", bold=True,
                   fn=lambda n, ra_=ra_, rb_=rb_, g=g: f"={qc(n)}{ra_}*(1-ImpPct_{g}*({qc(n)}$5>IMPE_{g}))+{qc(n)}{rb_}")
    imp[g] = C.row(f"{GEN_NAME[g]}: impairment / repricing charge", "$M", "Skeptic", "", 
                   fn=lambda n, cr=cr, ra_=ra_, g=g: f"=IF({qc(n)}$5=IMPE_{g},ImpPct_{g}*(SUMPRODUCT($E${cr}:{qc(n)}${cr},--($E$5:{qc(n)}$5<=IMPE_{g}))-SUM($F${ra_}:{qc(n)}${ra_})),0)")
    me = C.r
    nbv[g] = C.row(f"{GEN_NAME[g]}: GPU net book value", "$M", "Modeling", "", kind="stock", e=f"=E{cr}",
                   fn=lambda n, cr=cr, g=g, me=me: f"={qc(n-1)}{me}+{qc(n)}{cr}-{qc(n)}{dep[g]}-{qc(n)}{imp[g]}")
    reg("Capex_DA", f"Dep_{g}", dep[g]); reg("Capex_DA", f"Imp_{g}", imp[g]); reg("Capex_DA", f"NBV_{g}", nbv[g])
C.section("Non-GPU depreciation")
def fw(n, cr, lq, dcl=False):
    c = qc(n)
    pl = f"(PISLag+($E$5:{c}$5>LastHist)*MAX(0,DCLead-2))" if dcl else "PISLag"
    return f"=SUMPRODUCT($E${cr}:{c}${cr},(({c}$5-$E$5:{c}$5)>={pl})*(({c}$5-$E$5:{c}$5)<{lq}+{pl}))/{lq}"
ds = C.row("Shell/building depreciation", "$M", "Modeling", "", fn=lambda n: fw(n, caps, "LQ_Shell", True))
dp = C.row("Power/electrical depreciation", "$M", "Modeling", "", fn=lambda n: fw(n, capp, "LQ_Pow", True))
dm = C.row("Maintenance capex depreciation", "$M", "Modeling", "", fn=lambda n: fw(n, capm, "LQ_Maint"))
dlp = C.row("Lease-prepayment amortisation", "$M", "Modeling", "", fn=lambda n: fw(n, caplp, "LQ_Pow", True))
drp = C.row("Replacement capex depreciation", "$M", "Modeling", "", fn=lambda n: fw(n, caprp, "LQ_N"))
reg("Capex_DA", "Dep_Shell", ds); reg("Capex_DA", "Dep_Pow", dp); reg("Capex_DA", "Dep_Repl", drp); reg("Capex_DA", "CapShell", caps); reg("Capex_DA", "CapPow", capp)
do = C.row("Other D&A (intangibles)", "$M", "Filings", "", fn=lambda n: "=OtherDA")
cin = C.row("Capitalized interest (non-cash PP&E addition)", "$M", "Filings", "Hist Note 12; fcst = CapIntPct x (convert coupon + accretion + secured interest)",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_CapInt,1,{qc(n)}$5),CapIntPct*({X('CV_cpn', n)}+{X('CV_acc', n)}+{X('Sec_Int', n)}))")
reg("Capex_DA", "CapInt", cin)
dci = C.row("Capitalized-interest depreciation", "$M", "Modeling", "", fn=lambda n: fw(n, cin, "LQ_CapInt"))
reg("Capex_DA", "Dep_CI", dci)
dov = C.row("Consensus capex overlay depreciation", "$M", "Modeling", "", fn=lambda n: fw(n, cov, "LQ_B"))
C.section("D&A totals")
dgt = C.row("GPU depreciation - all vintages", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{dep['H']}+{qc(n)}{dep['B']}+{qc(n)}{dep['N']}")
reg("Capex_DA", "Dep_GPU", dgt)
dtm = C.row("Total D&A - model", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{dgt}+{qc(n)}{ds}+{qc(n)}{dp}+{qc(n)}{dm}+{qc(n)}{do}+{qc(n)}{dci}+{qc(n)}{dov}+{qc(n)}{drp}+{qc(n)}{dlp}")
reg("Capex_DA", "DA_Model", dtm)
dtu = C.row("D&A USED (actual for history, model for forecast)", "$M", "Modeling", "", bold=True,
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_DA,1,{qc(n)}$5),{qc(n)}{dtm})")
reg("Capex_DA", "DA", dtu)
C.row("Model D&A vs reported (history calibration)", "%", "Modeling", "", kind="none", fmt="0%;-0%",
      fn=lambda n: f"=IF({qc(n)}$5<=LastHist,{qc(n)}{dtm}/INDEX(H_DA,1,{qc(n)}$5)-1,\"\")")
reg("Capex_DA", "DACalib", C.r - 1)
IMT_MARK = 1
imt = C.row("Impairment charges (forecast)", "$M", "Skeptic", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{qc(n)}{imp['H']}+{qc(n)}{imp['B']}+{qc(n)}{imp['N']})")
reg("Capex_DA", "Imp", imt)
nbvt = C.row("Total PP&E net book value (model)", "$M", "Modeling", "", kind="stock", e=f"=E{nbv['H']}",
             fn=lambda n: f"={qc(n)}{nbv['H']}+{qc(n)}{nbv['B']}+{qc(n)}{nbv['N']}+SUM($F{caps}:{qc(n)}{caps})-SUM($F{ds}:{qc(n)}{ds})+SUM($F{capp}:{qc(n)}{capp})-SUM($F{dp}:{qc(n)}{dp})+SUM($F{capm}:{qc(n)}{capm})-SUM($F{dm}:{qc(n)}{dm})+SUM($F{caprp}:{qc(n)}{caprp})-SUM($F{drp}:{qc(n)}{drp})")
reg("Capex_DA", "NBV", nbvt)

# ------------------------------------------------------------------ FINANCING
fi = TS(wb, "Financing", "Convertible notes by tranche (older series + Aug-2026 2030s and 2034s), exchange, secured term loan. Convert dilution routed to Shares; non-dilutive converts counted in net debt.")
Fz = fi; fws = fi.ws
Fz.section("Convertible tranche table (owner: Filings)")
fh = ["Series", "Issue date", "Coupon", "Maturity date", "Original principal $M", "Prior conversions $M", "Exchanged $M", "Exchange qtr idx", "Exch. shares (M)",
      "Accreted value at maturity (x orig)", "Conv. rate /$1,000", "Conv. price $", "Issue idx", "Maturity idx", "ITM at SharePx (eff. px)", "Dilution factor", "Tag", "Source / method", "Carrying / principal at start (Note 12 6/30/26 amortized cost; new notes = face)", "Carrying start date"]
for i, t in enumerate(fh):
    fws.cell(Fz.r, 1 + i, t).font = BOLD
Fz.r += 1
f0 = Fz.r
TRN = [
 ("2.00% Jun-2029 (Jun-25)", dt.date(2025, 6, 5), 0.02, dt.date(2029, 6, 1), 500, 12.5, 400, "=ExchIdx", "=ExchShares/2", 1.175, 19.4363, "Reported principal/coupon/rate; prior conversions $12.5M (2025, Filings); exchange split 50/50 est.; accretion derived from Q2'26 Note 12 accreted $587.5M; maturity day est. Jun-1"),
 ("3.00% Jun-2031 (Jun-25)", dt.date(2025, 6, 5), 0.03, dt.date(2031, 6, 1), 500, 12.5, 400, "=ExchIdx", "=ExchShares/2", 1.225, 19.4363, "As above; accreted $612.5M"),
 ("1.00% Sep-2030 (Sep-25)", dt.date(2025, 9, 15), 0.01, dt.date(2030, 9, 15), 1581.25, 0, 0, 0, 0, 1.15, 7.2072, "Orig derived ($4.16B Jun+Sep total less $1.0B, split equally); accretes to 115% ($1,818.4M); rate per 20-F; maturity day est."),
 ("2.75% Sep-2032 (Sep-25)", dt.date(2025, 9, 15), 0.0275, dt.date(2032, 9, 15), 1581.25, 0, 0, 0, 0, 1.15, 7.2072, "As above"),
 ("1.25% Mar-2031 (Mar-26)", dt.date(2026, 3, 20), 0.0125, dt.date(2031, 3, 15), 2587.5, 0, 0, 0, 0, 1.20, 5.4579, "Q2'26 Note 12: $2.59B incl. greenshoe; accretes to 120%; $183.22 conv. px"),
 ("2.625% Mar-2033 (Mar-26)", dt.date(2026, 3, 20), 0.02625, dt.date(2033, 3, 15), 1750, 0, 0, 0, 0, 1.20, 5.546, "Q2'26 Note 12: $1.75B; accretes to 120%; $180.31 conv. px"),
 ("0.50% Feb-2030 (Aug-26)", dt.date(2026, 8, 24), 0.005, dt.date(2030, 2, 15), 3450, 0, 0, 0, 0, 1.10, 3.1902, "Pricing PR Aug 19 + closing PR Aug 24 2026: $3.45B incl. $450M option; accretes to 110%; $313.46 conv. px ($344.81 effective)"),
 ("4.50% Feb-2034 (Aug-26)", dt.date(2026, 8, 24), 0.045, dt.date(2034, 2, 15), 2300, 0, 0, 0, 0, 1.25, 3.0802, "$2.30B incl. $300M option; accretes to 125%; $324.65 conv. px ($405.82 effective)"),
]
NEWTR = []
for _d, _ix in [(1, "FP_CV1Idx"), (2, "FP_CV2Idx")]:
    for _s, _lab in [("S", "short"), ("L", "long")]:
        NEWTR.append((f"PLAN deal {_d} {_lab} (Estimated, Aug-26 terms)", f"=INDEX(Q_End,1,{_ix}-1)+45", f"=FP_Cpn{_s}", f"=EDATE(B{{rr}},12*FP_Ten{_s})",
                      f"=FP_On*(1-StopBuild)*FP_CVOpen*FP_CV{_d}Net*" + ("FP_ShortSh" if _s == "S" else "(1-FP_ShortSh)"), 0, 0, 0, 0, f"=FP_Acc{_s}", f"=1000/(FP_RefPx*(1+FP_Prem{_s}))",
                      "ESTIMATED 2027 funding plan; terms = Aug-26 pattern (Filings Converts): 60/40 short/long, 0.50%/4.50%, 110%/125% accretion, +40%/+45% premium, 3.5/7.5 yrs"))
NTR0 = len(TRN)
TRN = TRN + NEWTR
for i, t in enumerate(TRN):
    rr = f0 + i
    fws[f"A{rr}"] = t[0]
    if i >= NTR0:
        fml(fws[f"B{rr}"], t[1], "yyyy-mm-dd"); fml(fws[f"C{rr}"], t[2], "0.000%"); fml(fws[f"D{rr}"], t[3].format(rr=rr), "yyyy-mm-dd"); fml(fws[f"E{rr}"], t[4], "#,##0.0")
        for _c in "BCDE": fws[f"{_c}{rr}"].fill = F_EST
    else:
        inp(fws[f"B{rr}"], t[1], "Reported", "yyyy-mm-dd"); inp(fws[f"C{rr}"], t[2], "Reported", "0.000%")
        inp(fws[f"D{rr}"], t[3], "Reported" if t[3].year in (2030, 2034) and t[3].month == 2 else "Estimated", "yyyy-mm-dd")
        inp(fws[f"E{rr}"], t[4], "Reported" if i not in (2, 3) else "Estimated", "#,##0.0")
    inp(fws[f"F{rr}"], t[5], "Reported", "#,##0.0"); inp(fws[f"G{rr}"], t[6], "Reported", "#,##0.0")
    for c, v in [("H", t[7]), ("I", t[8])]:
        if isinstance(v, str): fml(fws[f"{c}{rr}"], v, "0.0")
        else: inp(fws[f"{c}{rr}"], v, "Reported", "0.0")
    if isinstance(t[9], str): fml(fws[f"J{rr}"], t[9], "0.000"); fws[f"J{rr}"].fill = F_EST
    else: inp(fws[f"J{rr}"], t[9], "Reported", "0.000")
    if isinstance(t[10], str): fml(fws[f"K{rr}"], t[10], "0.0000"); fws[f"K{rr}"].fill = F_EST
    else: inp(fws[f"K{rr}"], t[10], "Reported", "0.0000")
    fml(fws[f"L{rr}"], f"=1000/K{rr}", "0.00")
    fml(fws[f"M{rr}"], f"=IFERROR(MATCH(B{rr}-1,Q_End,1),0)+1", "0")
    fml(fws[f"N{rr}"], f"=IFERROR(MATCH(D{rr}-1,Q_End,1),0)+1", "0")
    fml(fws[f"O{rr}"], f"=IF(SharePx>=L{rr}*J{rr},1,0)", "0")
    fml(fws[f"P{rr}"], f"=CHOOSE(DilMethod,1,IF(DilPx>L{rr},1,0),0)", "0")
    _c630 = {0: 504.1 / 487.5, 1: 498.4 / 487.5, 2: 1588.2 / 1581.25, 3: 1575.0 / 1581.25, 4: 2588.6 / 2587.5, 5: 1744.7 / 1750.0}
    if i in _c630:
        inp(fws[f"S{rr}"], round(_c630[i], 6), "Reported", "0.0000"); inp(fws[f"T{rr}"], dt.date(2026, 6, 30), "Reported", "yyyy-mm-dd")
    else:
        fml(fws[f"S{rr}"], "=1", "0.0000"); fml(fws[f"T{rr}"], f"=B{rr}", "yyyy-mm-dd")
    fws[f"Q{rr}"] = "Reported/Derived"; fws[f"R{rr}"] = t[11]
    for c in "QR": fws[f"{c}{rr}"].font = ITAL
f1 = f0 + len(TRN) - 1
name(wb, "CV_Rate", "Financing", f"$K${f0}:$K${f1}"); name(wb, "CV_Px", "Financing", f"$L${f0}:$L${f1}")
Fz.r = f1 + 1
fws[f"A{Fz.r}"] = "Totals (original principal)"; fws[f"A{Fz.r}"].font = BOLD
for c in "EFGI": fml(fws[f"{c}{Fz.r}"], f"=SUM({c}{f0}:{c}{f1})", "#,##0.0", True)
Fz.r += 2
FR = {k: [] for k in ["out", "cpn", "acc", "bal", "rep", "csh", "dil", "debt"]}
Fz.section("Quarterly convert build (per tranche)")
for i, t in enumerate(TRN):
    rr = f0 + i
    o = Fz.row(f"{t[0]}: original principal outstanding", "$M", "Filings", "", kind="stock",
               fn=lambda n, rr=rr: f"=IF(AND({qc(n)}$5>=$M${rr},{qc(n)}$5<$N${rr}),$E${rr}-$F${rr}-IF({qc(n)}$5>=$H${rr},$G${rr},0),0)*($H${rr}<>0)+IF(AND({qc(n)}$5>=$M${rr},{qc(n)}$5<$N${rr}),$E${rr}-$F${rr},0)*($H${rr}=0)")
    cp = Fz.row(f"{t[0]}: cash coupon", "$M", "Filings", "", fn=lambda n, rr=rr, o=o: f"={qc(n)}{o}*$C${rr}/4")
    ac = Fz.row(f"{t[0]}: accretion (non-cash interest)", "$M", "Filings", "",
                fn=lambda n, rr=rr, o=o: f"={qc(n)}{o}*($J${rr}-1)/(($D${rr}-$B${rr})/91.3125)")
    bl = Fz.row(f"{t[0]}: accreted balance", "$M", "Filings", "", kind="stock",
                fn=lambda n, rr=rr, o=o: f"={qc(n)}{o}*(1+($J${rr}-1)*MIN(1,MAX(0,({qc(n)}$4-$B${rr})/($D${rr}-$B${rr}))))")
    cvc = Fz.row(f"{t[0]}: carrying value (Note 12 amortized cost at 6/30/26 or face at issue, accreted linearly to redemption amount)", "$M", "Filings", "", kind="stock",
                fn=lambda n, rr=rr, o=o: f"={qc(n)}{o}*($S${rr}+($J${rr}-$S${rr})*MIN(1,MAX(0,({qc(n)}$4-$T${rr})/($D${rr}-$T${rr}))))")
    reg("Financing", f"CVc_{i}", cvc)
    rp_ = Fz.row(f"{t[0]}: cash repayment at maturity (if not converted)", "$M", "Filings", "",
                 fn=lambda n, rr=rr, o=o: f"=IF({qc(n)}$5=$N${rr},{qc(n-1)}{o}*$J${rr}*(1-$O${rr}),0)")
    cs = Fz.row(f"{t[0]}: shares issued on conversion at maturity", "M sh", "Filings", "", fmt="#,##0.00",
                fn=lambda n, rr=rr, o=o: f"=IF({qc(n)}$5=$N${rr},{qc(n-1)}{o}*$K${rr}/1000*$O${rr},0)")
    dl = Fz.row(f"{t[0]}: dilutive shares (per DilMethod)", "M sh", "Modeling", "", kind="stock", fmt="#,##0.00",
                fn=lambda n, rr=rr, o=o, bl=bl: f"=IF(DilMethod=3,MAX(0,{qc(n)}{o}*$K${rr}/1000*DilPx-{qc(n)}{bl})/DilPx,{qc(n)}{o}*$K${rr}/1000*$P${rr})")
    db = Fz.row(f"{t[0]}: counted as debt in net debt", "$M", "Modeling", "", kind="stock",
                fn=lambda n, rr=rr, bl=bl: f"={qc(n)}{bl}*IF(DilMethod=3,1,1-$P${rr})")
    for k, v in zip(FR.keys(), [o, cp, ac, bl, rp_, cs, dl, db]): FR[k].append(v)
    reg("Financing", f"CVo_{i}", o); reg("Financing", f"CVb_{i}", bl); reg("Financing", f"CVd_{i}", dl); reg("Financing", f"CVx_{i}", db)
Fz.section("Convert totals")
tot = {}
for k, lab, kind, fmt_ in [("out", "Original principal outstanding", "stock", None), ("cpn", "Cash coupons", "flow", None), ("acc", "Accretion (non-cash)", "flow", None),
                           ("bal", "Accreted balance", "stock", None), ("rep", "Cash repayments at maturity", "flow", None), ("csh", "Shares issued at maturity conversion", "flow", "#,##0.00"),
                           ("dil", "Dilutive convert shares", "stock", "#,##0.00"), ("debt", "Converts counted as debt", "stock", None)]:
    tot[k] = Fz.row(lab, "M sh" if fmt_ else "$M", "Modeling", "", kind=kind, bold=True, fmt=fmt_ or "#,##0.0;(#,##0.0);-",
                    fn=lambda n, k=k: "=" + "+".join(f"{qc(n)}{x}" for x in FR[k]))
    reg("Financing", f"CV_{k}", tot[k])
Fz.section("Secured term loan ($775M, SOFR + 2.50%)")
sb = Fz.row("Secured loan balance", "$M", "Filings", "Bullet assumed (amortization n.d.)", kind="stock",
            fn=lambda n: f"=IF(AND({qc(n)}$5>=SecIdx,{qc(n)}$5<SecMatIdx),SecAmt,0)*LoanOn")
reg("Financing", "Sec_Bal", sb)
si = Fz.row("Secured loan interest", "$M", "Filings", "", fn=lambda n: f"={qc(n)}{sb}*(SOFR+SecSpread)/4")
reg("Financing", "Sec_Int", si)
sd = Fz.row("Secured loan draws (+) / repayments (-)", "$M", "Filings", "",
            fn=lambda n: f"=(IF({qc(n)}$5=SecIdx,SecAmt,0)-IF({qc(n)}$5=SecMatIdx,SecAmt,0))*LoanOn")
reg("Financing", "Sec_Flow", sd)

# ------------------------------------------------------------------ P&L
pl = TS(wb, "PnL", "Quarterly P&L. History = reported (COR/PD/SG&A GAAP incl. SBC); forecast = driver-based ex-SBC with SBC below EBITDA. Interest/tax/net income modeled for forecast only.")
P = pl
def H_or(nm_hist, model_expr, n):
    return f"=IF({qc(n)}$5<=LastHist,INDEX({nm_hist},1,{qc(n)}$5),{model_expr})"
P.section("Revenue")
rev = P.row("Group revenue", "$M", "Revenue", "", bold=True, fn=lambda n: f"={X('Rev', n)}"); reg("PnL", "PL_Rev", rev)
P.row("  AI cloud revenue", "$M", "Revenue", "", fn=lambda n: f"={X('AIRev', n)}")
P.row("  YoY growth", "%", "", "", kind="none", fmt="0%", fn=lambda n: (f"=IF({qc(n-4)}{rev}>0,{qc(n)}{rev}/{qc(n-4)}{rev}-1,\"\")" if n > 4 else None),
      fy_fn=lambda y, c: (f"=IF({CL(FY_COL0+YEARS.index(y)-1)}{rev}>0,{c}{rev}/{CL(FY_COL0+YEARS.index(y)-1)}{rev}-1,\"\")" if y > 2024 else None))
P.section("Operating costs")
_vphys = lambda n: f"IF(FixMgmt=1,MIN({X('MS_BillMW', n)},INDEX(VinEnIT,1,MAX(1,{n}-IF({n}>LastHist,EnLag,0)))),{X('MS_BillMW', n)})"
pw = P.row("Power - Vineland (Microsoft billing MW physically energized, PPA/genset opex)", "$M", "Capacity", "min(MS billing IT MW, Vineland energized IT MW) x PUE x VinPower / 4",
           fn=lambda n: f"={_vphys(n)}*IF({qc(n)}$5<=LastHist,PUE_B,PUE)*VinPower/4")
po = P.row("Power - rest of fleet", "$M", "Capacity", "(avg IT MW - Vineland physical MS MW) x LoadFactor x PUE x hrs x $/MWh",
           fn=lambda n: f"=MAX(0,{X('MW_Avg', n)}-{_vphys(n)})*LoadFactor*IF({qc(n)}$5<=LastHist,PUE_B,PUE)*{HRS(n)}*PowerPx/1000000")
rn = P.row("Rent - leased / colo / build-to-suit (shell + power infra)", "$M", "Capacity", "avg IT MW x (1 - owned share) x PUE x RentMW / 4",
           fn=lambda n: f"={X('MW_Avg', n)}*(1-{X('MW_OwnSh', n)})*IF({qc(n)}$5<=LastHist,PUE_B,PUE)*RentMW/4")
reg("PnL", "PL_Rent", rn)
rn2 = P.row("  AIB CLT1 rent differential (CLT1Rent per IT MW vs generic RentMW x PUE)", "$M", "Capacity", "CLT1 billing MW (energized lagged EnLag) x (CLT1Rent - RentMW x PUE) / 4",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,INDEX(CLT1_En,1,MAX(1,{n}-EnLag))*(CLT1Rent-RentMW*PUE)/4)")
oc = P.row("Other cost of revenue", "$M", "Modeling", "", fn=lambda n: f"=OtherCOR*{qc(n)}{rev}")
om = P.row("Owned-site O&M (staff, maintenance, insurance, property tax; NO rent)", "$M", "Capacity", "avg IT MW x owned share x (OwnOpex - power per IT MW) / 4 (OwnOpexOn) else OM_MW",
           fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{X('MW_Avg', n)}*{X('MW_OwnSh', n)}*IF(OwnOpexOn=1,MAX(0,OwnOpex-LoadFactor*PUE*8760*PowerPx/1000000),OM_MW)/4)")
cor = P.row("Cost of revenue (hist = reported GAAP)", "$M", "Modeling", "", bold=True, fn=lambda n: H_or("H_COR", f"{qc(n)}{pw}+{qc(n)}{po}+{qc(n)}{rn}+{qc(n)}{rn2}+{qc(n)}{oc}+{qc(n)}{om}", n))
pd_ = P.row("Product development (hist GAAP; fcst ex-SBC)", "$M", "Modeling", "", fn=lambda n: H_or("H_PD", f"IF(AND(OpexHold=1,{qc(n)}$7>2028),INDEX(PDpct,1,2028-FYBase),{ANN('PDpct', n)})*{qc(n)}{rev}", n))
sg = P.row("SG&A (hist GAAP; fcst ex-SBC)", "$M", "Modeling", "", fn=lambda n: H_or("H_SGA", f"IF(AND(OpexHold=1,{qc(n)}$7>2028),INDEX(SGApct,1,2028-FYBase),{ANN('SGApct', n)})*{qc(n)}{rev}", n))
P.section("Profitability")
pmo = P.row("Peer-margin overlay (sensitivity; PeerMarginOn): cost so adj. EBITDA margin <= MarginPeer from PeerMarginYr", "$M", "GPU Market", "",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,PeerMarginOn*({qc(n)}$7>=PeerMarginYr)*MAX(0,{qc(n)}{rev}-{qc(n)}{cor}-{qc(n)}{pd_}-{qc(n)}{sg}-MarginPeer*{qc(n)}{rev}))")
eb = P.row("ADJUSTED EBITDA", "$M", "Modeling", "History = reported adj. EBITDA", bold=True,
           fn=lambda n: H_or("H_EBITDA", f"{qc(n)}{rev}-{qc(n)}{cor}-{qc(n)}{pd_}-{qc(n)}{sg}-{qc(n)}{pmo}", n))
reg("PnL", "PL_EBITDA", eb)
P.row("  Adj. EBITDA margin", "%", "", "", kind="none", fmt="0%", fn=lambda n: f"=IF({qc(n)}{rev}<>0,{qc(n)}{eb}/{qc(n)}{rev},0)",
      fy_fn=lambda y, c: f"=IF({c}{rev}<>0,{c}{eb}/{c}{rev},0)")
sbc = P.row("SBC & other EBITDA adjustments", "$M", "Filings", "Hist derived = adj EBITDA - (rev - GAAP opex ex-D&A)",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,{qc(n)}{eb}-({qc(n)}{rev}-{qc(n)}{cor}-{qc(n)}{pd_}-{qc(n)}{sg}),SBCpct*{qc(n)}{rev})")
da = P.row("Depreciation & amortization", "$M", "Capex_DA", "", fn=lambda n: f"={X('DA', n)}"); reg("PnL", "PL_DA", da)
im = P.row("Impairment / repricing", "$M", "Skeptic", "", fn=lambda n: f"={X('Imp', n)}")
ebit = P.row("EBIT", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{eb}-{qc(n)}{sbc}-{qc(n)}{da}-{qc(n)}{im}")
reg("PnL", "PL_EBIT", ebit)
P.row("  EBIT margin", "%", "", "", kind="none", fmt="0%", fn=lambda n: f"=IF({qc(n)}{rev}<>0,{qc(n)}{ebit}/{qc(n)}{rev},0)",
      fy_fn=lambda y, c: f"=IF({c}{rev}<>0,{c}{ebit}/{c}{rev},0)")
P.section("Below EBIT (forecast only; history: see filings)")
def F_only(expr, n): return f"=IF({qc(n)}$5<=LastHist,0,{expr})"
i1 = P.row("Convert cash coupons", "$M", "Financing", "", fn=lambda n: F_only(X('CV_cpn', n), n))
i2 = P.row("Convert accretion (non-cash)", "$M", "Financing", "", fn=lambda n: F_only(X('CV_acc', n), n))
i3 = P.row("Secured loan interest + v0.8b ABS interest", "$M", "Financing", "", fn=lambda n: F_only(X('Sec_Int', n) + "+" + X('ABS_Int', n), n))
i4 = P.row("Funding-plug debt interest", "$M", "CashFlow", "", fn=lambda n: F_only(X('CF_PlugInt', n), n))
i5 = P.row("Interest income", "$M", "CashFlow", "", fn=lambda n: F_only(X('CF_IntInc', n), n))
i6 = P.row("Less: capitalized interest (to PP&E)", "$M", "Capex_DA", "", fn=lambda n: F_only(X('CapInt', n), n))
ebt = P.row("Pre-tax income", "$M", "Modeling", "", bold=True, fn=lambda n: F_only(f"{qc(n)}{ebit}-{qc(n)}{i1}-{qc(n)}{i2}-{qc(n)}{i3}-{qc(n)}{i4}+{qc(n)}{i5}+{qc(n)}{i6}", n))
tx = P.row("Income tax", "$M", "Modeling", "TaxRate x positive EBT (no NOL)", fn=lambda n: f"=MAX(0,{qc(n)}{ebt})*TaxRate")
reg("PnL", "PL_Tax", tx)
ni = P.row("Net income", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{ebt}-{qc(n)}{tx}"); reg("PnL", "PL_NI", ni)
P.row("Diluted shares (Shares)", "M sh", "Shares", "", kind="stock", fmt="#,##0.0", fn=lambda n: f"={X('Sh_Dil', n)}")
P.row("Diluted EPS (forecast)", "$/sh", "Modeling", "", fmt="0.00;(0.00);-", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{qc(n)}{ni}/{X('Sh_Dil', n)})",
      fy_fn=lambda y, c: f"=IFERROR({c}{ni}/INDEX(Sh_Dil,1,MATCH({c}$7,Q_Yr,1)),0)")

# ------------------------------------------------------------------ CASH FLOW
cf = TS(wb, "CashFlow", "Cash flow, deferred revenue (prepayments), funding plug (debt/equity at MinCash), debt balances, net debt. Forecast from LastHist+1.")
K = cf
K.section("Operating cash flow")
ceb = K.row("Adj. EBITDA", "$M", "PnL", "", fn=lambda n: F_only(X('PL_EBITDA', n), n))
ctx = K.row("Cash taxes", "$M", "PnL", "", fn=lambda n: F_only(f"-{X('PL_Tax', n)}", n))
K.section("Deferred revenue (customer prepayments)")
me = K.r
def _mwn(nm, n):  # MW added at billing quarter n + GPULead
    return f"IFERROR(INDEX({nm},1,MIN({NQ},{n}+GPULead))-INDEX({nm},1,MIN({NQ},{n}+GPULead)-1),0)"
_tot = lambda n: f"IFERROR(INDEX(MW_New,1,MIN({NQ},{n}+GPULead)),0)"
K.section("Prepayment base: new-build capex split by the MW it serves (billing quarter = capex quarter + GPULead)")
psg = K.row("Share of new-build capex for SIGNED Microsoft/Meta MW (0% new prepayment)", "%", "Modeling", "", kind="avg", fmt="0%",
            fn=lambda n: f"=IF({_tot(n)}>0,MIN(1,MAX(0,{_mwn('SigMW', n)})/{_tot(n)}),0)")
reg("CashFlow", "CF_SigSh", psg)
pig = K.row("Share for NEW hyperscaler / IG MW (HSPre)", "%", "Modeling", "", kind="avg", fmt="0%",
            fn=lambda n: f"=IF({_tot(n)}>0,MIN(1-{qc(n)}{psg},MAX(0,{_mwn('IGMW', n)})/{_tot(n)}),0)")
pot = K.row("Share for other new deals (Neolab / non-hyperscaler)", "%", "Modeling", "", kind="avg", fmt="0%", fn=lambda n: f"=1-{qc(n)}{psg}-{qc(n)}{pig}")
prt = K.row("Prepayment rate - other new deals (PrepayDecayOn: 38.5% pre-Rubin, PrepayPct27 in 2027, PrepayPct28 from PrepayIdx28)", "%", "Skeptic", "", kind="avg", fmt="0.0%",
            fn=lambda n: f"=IF(PrepayDecayOn=0,PrepayPct,IF({qc(n)}$5<PxCycIdx,PrepayPct,IF({qc(n)}$5<PrepayIdx28,PrepayPct27,PrepayPct28)))")
dri = K.row("Prepayments received", "$M", "Filings", "NEW prepayments only (post 6/30) = new-build capex x (IG share x HSPre + other share x rate); signed Microsoft/Meta 0%; opening DR 5,975 is recognition only",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,IF(NDPrepayOn=1,(1-StopBuild)*NDPrepayPct*{X('CapNew', n)}*(1-{qc(n)}{psg}),(1-StopBuild)*PrepayOn*IF(AND(PrepayMode>=2,{qc(n)}$5<=PrepayEndIdx),IF(AND(PrepayMode=3,{qc(n)}$5<=LastHist+2),PrepayH2/2,PrepayQ),IF(PrepayBase=1,{X('CapNew', n)}*IF(PrepayByCust=1,{qc(n)}{pig}*HSPre+{qc(n)}{pot}*{qc(n)}{prt},PrepayPct),PrepayPct*{X('Capex', n)}))))")
reg("CashFlow", "CF_Prepay", dri)
reg("CashFlow", "CF_PrepayRate", prt)
dro = K.row("Release of 6/30 deferred revenue (Reported split): DR_Cur over Q3'26-Q2'27, DR_24-DR_Cur over Q3'27-Q2'28, remainder over DR_TailQ", "$M", "Filings", "non-cash revenue; existing Microsoft/Meta advances",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,IF({qc(n)}$5<=LastHist+4,DR_Cur/4,IF({qc(n)}$5<=LastHist+8,(DR_24-DR_Cur)/4,IF({qc(n)}$5<=LastHist+8+DR_TailQ,(INDEX(H_DR,1,LastHist)-DR_24)/DR_TailQ,0))))")
reg("CashFlow", "CF_DRRelOpen", dro)
drw = K.row("Release of NEW prepayments: straight-line over DR_AmortQ from billing start (capex quarter + GPULead)", "$M", "Modeling", "slow burn-off consistent with multi-year tranches",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,SUMPRODUCT(($F$5:$AG$5+GPULead<={qc(n)}$5)*($F$5:$AG$5+GPULead+DR_AmortQ>{qc(n)}$5)*$F${dri}:$AG${dri})/DR_AmortQ)")
drr = K.row("Deferred revenue released (non-cash revenue)", "$M", "Modeling", "6/30 balance schedule + new prepayment vintages",
            fn=lambda n: f"=IF(DRSchedMode=1,{qc(n)}{dro}+{qc(n)}{drw},IF({qc(n)}$5<=LastHist,0,{qc(n-1)}{K.r+1}/DR_AmortQ))")
drb_r = K.r
drb = K.row("Deferred revenue balance", "$M", "Filings", "History = Hist H_DR", kind="stock",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_DR,1,{qc(n)}$5),{qc(n-1)}{drb_r}+{qc(n)}{dri}-{qc(n)}{drr})")
assert drb == drb_r
reg("CashFlow", "CF_DRBal", drb); reg("CashFlow", "CF_DRRel", drr)
ddr = K.row("Change in deferred revenue", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{dri}-{qc(n)}{drr}")
K.section("v0.8b Asset-backed debt on CONTRACTED MW (ABS_LTC x new-build capex x signed share; SOFR + ABS_Spread; straight-line over ABS_Term from billing start)")
absd = K.row("ABS draws", "$M", "Modeling", "with capex", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,ABS_LTC*{X('CapNew', n)}*{qc(n)}{psg})")
absr = K.row("ABS amortisation", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,SUMPRODUCT(($F$5:$AG$5+GPULead<{qc(n)}$5)*($F$5:$AG$5+GPULead+ABS_Term>={qc(n)}$5)*$F${absd}:$AG${absd})/ABS_Term)")
_ab = K.r
absb = K.row("ABS balance", "$M", "Modeling", "", kind="stock", e="=0", fn=lambda n: f"={qc(n-1)}{_ab}+{qc(n)}{absd}-{qc(n)}{absr}")
assert absb == _ab
absi = K.row("ABS interest (prior balance x (SOFR + ABS_Spread) / 4)", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{qc(n-1)}{absb}*(SOFR+ABS_Spread)/4)")
for _n, _r in (("ABS_Draw", absd), ("ABS_Rep", absr), ("ABS_Bal", absb), ("ABS_Int", absi)): reg("CashFlow", _n, _r)
K.section("Interest (cash)")
cash_r_placeholder = None
iinc_r = K.r; plgi_r = K.r + 1
# rows filled after cash/plug rows known -> use names
iinc = K.row("Interest income on prior cash", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{X('CF_Cash', n-1)}*CashYield/4)" if n > 1 else "=0")
reg("CashFlow", "CF_IntInc", iinc)
plgi = K.row("Interest on funding-plug debt (prior balance)", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{X('CF_PlugBal', n-1)}*(SOFR+PlugSpread)/4)" if n > 1 else "=0")
reg("CashFlow", "CF_PlugInt", plgi)
reg("CashFlow", "CF_PlugIntRow", plgi)
cint = K.row("Net cash interest", "$M", "Modeling", "-(coupons + secured + plug interest) + interest income",
             fn=lambda n: F_only(f"-{X('CV_cpn', n)}-{X('Sec_Int', n)}-{X('ABS_Int', n)}-{qc(n)}{plgi}+{qc(n)}{iinc}", n))
cfo = K.row("CASH FROM OPERATIONS (proxy)", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{ceb}+{qc(n)}{ctx}+{qc(n)}{cint}+{qc(n)}{ddr}")
ccx = K.row("Capex (net of after-tax GPU residual proceeds)", "$M", "Capex_DA", "", fn=lambda n: F_only(f"-{X('Capex', n)}+{X('CapResid', n)}", n))
oinv = K.row("Other investing: AIB CLT1 build prepayment (placeholder) + Inferize acquisition (press-reported, toggle) - liquidity only, not in UFCF", "$M", "Filings", "",
             fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,-CLT1Prepay/3*AND({qc(n)}$5>=12,{qc(n)}$5<=14)-InferizeOn*InferizeAmt*({qc(n)}$5=12))")
reg("CashFlow", "CF_OthInv", oinv)
fcf = K.row("FREE CASH FLOW (levered)", "$M", "Modeling", "", bold=True, fn=lambda n: f"={qc(n)}{cfo}+{qc(n)}{ccx}+{qc(n)}{oinv}")
reg("CashFlow", "CF_FCF", fcf)
ufcf = K.row("Unlevered FCF (for DCF)", "$M", "Modeling", "EBITDA - tax on EBIT + change in DR - capex",
             fn=lambda n: F_only(f"{X('PL_EBITDA', n)}-MAX(0,{X('PL_EBIT', n)}+CapIntFix*{X('Dep_CI', n)})*TaxRate+{qc(n)}{ddr}-{X('Capex', n)}+{X('CapResid', n)}", n))
reg("CashFlow", "CF_UFCF", ufcf)
K.section("Financing flows & funding plug")
cvp = K.row("Convert issuance proceeds (net): Aug-2026 2030s + 2034s (ONCE, Q3'26; 6/30 cash excludes them) + 2027 plan deals", "$M", "Financing", "Aug converts closed 8/24/26 -> counted only here (Q3'26) and in runway row '+ Aug-2026 converts gross'",
            fn=lambda n: f"=(IF({qc(n)}$5=11,AugNet,0)+(1-FP_Fee)*(" + "+".join(f"MAX(0,{X('CVo_'+str(k), n)}-{X('CVo_'+str(k), n-1)})" for k in range(NTR0, len(TRN))) + f"))*({qc(n)}$5>LastHist)")
fpc = K.row("  of which: 2027 funding-plan converts (net of fees)", "$M", "Modeling", "", fn=lambda n: f"=(1-FP_Fee)*(" + "+".join(f"MAX(0,{X('CVo_'+str(k), n)}-{X('CVo_'+str(k), n-1)})" for k in range(NTR0, len(TRN))) + f")*({qc(n)}$5>LastHist)")
reg("CashFlow", "CF_FPConv", fpc)
reg("CashFlow", "CF_CVProc", cvp)
cvr = K.row("Convert cash repayments at maturity", "$M", "Financing", "", fn=lambda n: F_only(f"-{X('CV_rep', n)}", n))
scf = K.row("Secured loan draw / (repayment)", "$M", "Financing", "", fn=lambda n: F_only(X('Sec_Flow', n), n))
atm = K.row("ATM equity proceeds (Q3'26 actual sales + funding-plan ATM at FP_ATMIdx)", "$M", "Filings", "", fn=lambda n: f"=(IF({qc(n)}$5=11,ATM_Q3*SharePx,0)+IF({qc(n)}$5=FP_ATMIdx,FP_ATMProc,0))*({qc(n)}$5>LastHist)")
reg("CashFlow", "CF_ATM", atm)
cash_r = K.r + 5
pre = K.row("Cash before funding plug", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,{qc(n-1)}{cash_r}+{qc(n)}{fcf}+{qc(n)}{cvp}+{qc(n)}{cvr}+{qc(n)}{scf}+{qc(n)}{atm}+{X('ABS_Draw', n)}-{X('ABS_Rep', n)})")
gap = K.row("Funding gap to MinCash", "$M", "Modeling", "", fn=lambda n: f"=IF({qc(n)}$5<=LastHist,0,MAX(0,MinCash-{qc(n)}{pre}))")
dpl = K.row("  Debt plug drawn", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{gap}*PlugDebtSh")
epl = K.row("  Equity plug raised", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{gap}*(1-PlugDebtSh)")
eps_ = K.row("  Equity plug shares issued", "M sh", "Modeling", "at SharePx x (1 - EqDisc)", fmt="#,##0.00", fn=lambda n: f"={qc(n)}{epl}/(SharePx*(1-EqDisc))")
reg("CashFlow", "CF_EqSh", eps_)
cash = K.row("CASH (end)", "$M", "Hist/Modeling", "History = Hist cash", kind="stock", bold=True,
             fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_Cash,1,{qc(n)}$5),{qc(n)}{pre}+{qc(n)}{gap})")
assert cash == cash_r, (cash, cash_r)
reg("CashFlow", "CF_Cash", cash)
cws_ = cf.ws; cws_[f"E{cash}"] = "=0"
K.section("Debt & net debt")
me = K.r
pb = K.row("Funding-plug debt balance", "$M", "Modeling", "", kind="stock", e="=0", fn=lambda n, me=me: f"={qc(n-1)}{me}+{qc(n)}{dpl}")
reg("CashFlow", "CF_PlugBal", pb)
gd = K.row("Gross debt (accreted converts + secured + plug)", "$M", "Modeling", "", kind="stock", fn=lambda n: f"={X('CV_bal', n)}+{X('Sec_Bal', n)}+{qc(n)}{pb}")
dc = K.row("Debt counted in net debt (excl. converts treated as equity)", "$M", "Modeling", "", kind="stock", fn=lambda n: f"={X('CV_debt', n)}+{X('Sec_Bal', n)}+{qc(n)}{pb}")
nd = K.row("NET DEBT (debt counted - cash)", "$M", "Modeling", "", kind="stock", bold=True, fn=lambda n: f"={qc(n)}{dc}-{qc(n)}{cash}")
reg("CashFlow", "CF_NetDebt", nd); reg("CashFlow", "CF_GrossDebt", gd)

# ------------------------------------------------------------------ SHARES
sh = TS(wb, "Shares", "Basic and diluted share count. Convert dilution from Financing (DilMethod); options by treasury-stock method at SharePx; NVIDIA pre-funded warrant included in economic basic.")
Sx = sh
me = Sx.r
bx = Sx.row("Issued & outstanding EXCLUDING funding-plan / convert-maturity issuance after the valuation quarter (pre-plan base: SBC drift, exchange, Q3'26 ATM actual; legacy CF plug if FP_Resize = 0)", "M sh", "Filings", "History = Hist; valuation-date count V_BASIC reads this row (no circularity)",
             kind="stock", fmt="#,##0.00", fn=lambda n, me=me: (f"=IF({qc(n)}$5<=LastHist,INDEX(H_Shares,1,{qc(n)}$5),{qc(n-1)}{me}*(1+SBCsh)"
                f"+IF({qc(n)}$5=ExchIdx,ExchShares,0)+IF({qc(n)}$5=11,ATM_Q3,0)+IF(AND({qc(n)}$5=FP_ATMIdx,{qc(n)}$5<=V_Q),FP_ATMSh,0)+IF(FP_Resize=1,0,{X('CF_EqSh', n)})+IF({qc(n)}$5<=V_Q,{X('CV_csh', n)},0))"))
reg("Shares", "Sh_BasicX", bx)
bas = Sx.row("Issued & outstanding shares (Class A + B) INCL. funding-plan issuance at the HEADLINE price basis (PTBasis) - ATM, equity top-ups, convert maturity conversions (v0.7d)", "M sh", "Modeling", "= pre-plan base + cumulative plan shares (Shares tab, plan rows below)",
             kind="stock", fmt="#,##0.00", bold=True, fn=lambda n: f"={qc(n)}{bx}+IF({qc(n)}$5<=LastHist,0,INDEX(Sh_PlanCum,1,{qc(n)}$5))")
reg("Shares", "Sh_Basic", bas)
wr = Sx.row("NVIDIA pre-funded warrant shares", "M sh", "Filings", "", kind="stock", fmt="#,##0.00", fn=lambda n: f"=IF({qc(n)}$5>=NVDA_Idx,NVDA_W,0)")
reg("Shares", "Sh_Warr", wr)
eb_ = Sx.row("Economic basic (issued + warrant)", "M sh", "Modeling", "", kind="stock", fmt="#,##0.00", bold=True, fn=lambda n: f"={qc(n)}{bas}+{qc(n)}{wr}")
reg("Shares", "Sh_EconBasic", eb_)
op = Sx.row("Options (treasury-stock method at DilPx)", "M sh", "Filings", "", kind="stock", fmt="#,##0.00", fn=lambda n: "=Options*MAX(0,1-OptK/DilPx)")
rs_ = Sx.row("Unvested RSUs (+ Eigen AI shares if EigenIn = 1)", "M sh", "Filings", "", kind="stock", fmt="#,##0.00", fn=lambda n: "=RSU+EigenIn*Eigen")
cvd = Sx.row("Convertible notes - dilutive shares (DilMethod)", "M sh", "Financing", "", kind="stock", fmt="#,##0.00", fn=lambda n: f"={X('CV_dil', n)}")
dil = Sx.row("DILUTED SHARES", "M sh", "Modeling", "", kind="stock", fmt="#,##0.00", bold=True, fn=lambda n: f"={qc(n)}{eb_}+{qc(n)}{op}+{qc(n)}{rs_}+{qc(n)}{cvd}")
reg("Shares", "Sh_Dil", dil)
Sx.row("  Convert dilution as % of economic basic", "%", "", "", kind="none", fmt="0.0%", fn=lambda n: f"={qc(n)}{cvd}/{qc(n)}{eb_}")

# ------------------------------------------------------------------ VALUATION
va = TS(wb, "Valuation", "DCF (quarterly UFCF discounted to ValDate + terminal value) with EV bridge incl. weighted off-balance-sheet commitments, SOTP, 12-month PT, cross-checks.")
V = va; vws = va.ws
V.section("DCF")
vu = V.row("Unlevered FCF", "$M", "CashFlow", "", fn=lambda n: f"={X('CF_UFCF', n)}")
vdf = V.row("Discount factor (periods after ValDate)", "x", "Modeling", "", kind="none", fmt="0.000",
            fn=lambda n: f"=IF({qc(n)}$4>ValDate,1/(1+WACC)^(({qc(n)}$4-ValDate)/365.25),0)")
vpv = V.row("PV of UFCF", "$M", "Modeling", "", fn=lambda n: f"={qc(n)}{vu}*{qc(n)}{vdf}")
V.blank()
s0 = V.r
def VL(label, formula, fmt="#,##0;(#,##0);-", nm=None, bold=False, note=""):
    rr = V.r
    vws[f"A{rr}"] = label; vws[f"A{rr}"].font = BOLD if bold else BLACK
    fml(vws[f"E{rr}"], formula, fmt, bold)
    vws[f"D{rr}"] = note; vws[f"D{rr}"].font = ITAL
    if nm: name(wb, nm, "Valuation", f"$E${rr}")
    V.r += 1
    return rr
V.section("Terminal value (terminal year = FY2030 x (1 + TG))")
ac = fyc(2030)
VL("Terminal growth used: TGMode 1 = MIN(TGCap, TermRR x TermRONC); 0 = TG", "=IF(TGMode=1,MIN(TGCap,TermRR*TermRONC),TG)", fmt="0.00%", nm="V_g")
VL("FY2030 revenue", f"=SUMIF(Q_Yr,2030,PL_Rev)", nm="V_Rev30")
VL("FY2030 adj. EBITDA", f"=SUMIF(Q_Yr,2030,PL_EBITDA)", nm="V_EB30")
VL("FY2030 D&A", f"=SUMIF(Q_Yr,2030,PL_DA)", nm="V_DA30")
VL("  MW scale to Q1'31 billing (Q4'30 active + MW energized in Q4'30) / Q4'30 active", "=IFERROR((INDEX(MW_Act,1,28)+MW_Add31)/INDEX(MW_Act,1,28),1)", fmt="0.000", nm="V_TScale")
VL("  Q4'30 adj. EBITDA x 4 (exit run-rate)", "=4*INDEX(PL_EBITDA,1,28)", nm="V_EBX")
VL("  v0.8d: annual revenue step when unsold vintages not yet renewed at Q4'30 renew (GB300 age < UncRenAge; Rubin 2028-30 under UncPxOn)", "=TVRenOn*UncRenew*(IF(STMode=1,1-INDEX(ND_STF,1,28),1-STShare)*SUMPRODUCT((Q_Idx>28-UncRenAge)*ND_NagBH*(1-ND_SSh)*ND_dGB)+IF(STMode=1,IF(STGBOnly=1,1,1-INDEX(ND_STF,1,28)),1-STShare)*SUMPRODUCT((Q_Idx>16)*ND_NagN*(1-ND_SSh)*ND_dN))", fmt="#,##0;-#,##0", nm="V_TRenRev")
VL("  v0.8d: terminal EBITDA effect of renewal = revenue step x (1 - corporate opex % rev FY30)", "=V_TRenRev*(1-(OtherCOR+INDEX(PDpct,1,2030-FYBase)+INDEX(SGApct,1,2030-FYBase)))", fmt="#,##0;-#,##0", nm="V_TRenEB")
VL("Terminal-year EBITDA (TermBase 2 = Q4'30 x 4, no MW step-up [default]; 1 = x MW scale [memo]; 0 = FY30)", "=(IF(TermBase=0,V_EB30,V_EBX*IF(TermBase=1,V_TScale,1))+V_TRenEB)*(1+V_g)", nm="V_EBT")
VL("Terminal-year D&A = maintenance capex base (ex capitalized-interest D&A when CapIntFix)", "=IF(TermBase=0,V_DA30-CapIntFix*SUMIF(Q_Yr,2030,Dep_CI),4*(INDEX(PL_DA,1,28)-CapIntFix*INDEX(Dep_CI,1,28))*IF(TermBase=1,V_TScale,1))*(1+V_g)", nm="V_DAT")
_gr = "+".join(f"SUMPRODUCT((Q_Idx>28-LQ_{g})*CapGPU_{g})/(LQ_{g}/4)*IF(ReplPxOn=1,(1+PxChg_{g})^(LQ_{g}/4),1)" for g in GENS) + "+SUMPRODUCT((Q_Idx>28-LQ_N)*CapRepl)/(LQ_N/4)*IF(ReplPxOn=1,(1+PxChg_N)^(LQ_N/4),1)"
VL("  GPU replacement need = gross GPU cost in service at Q4'30 / GPU life (per yr)", "=" + _gr, nm="V_GRepl", note="every cohort bought in the last life-years is replaced once per life")
VL("  GPU residual credit at refresh (after tax) on terminal replacement", "=V_GRepl*GPU_Resid*(1-TaxRate)", nm="V_GResid")
VL("  Non-GPU sustaining (cash lives): gross shell / Life_Shell + gross power-cooling / Life_Pow (replaced at yr 10) + FY30 maintenance capex", "=SUM(CapShell)/Life_Shell+SUM(CapPow)/Life_Pow+SUMIF(Q_Yr,2030,CapMaint)", nm="V_NonGPU")
VL("  Growth reinvestment: TGMode 1 = g / RONC x NOPAT (fundamentals); 0 = TG x net PP&E at Q4'30", "=IF(TGMode=1,IF(TermRONC>0,V_g/TermRONC,0)*V_NOPAT,TG*INDEX(NBV,1,28))", nm="V_GrowCx")
VL("  Required terminal capex = (replacement - residual credit + non-GPU) x (1+V_g) + growth", "=(V_GRepl-V_GResid+V_NonGPU)*(1+V_g)+V_GrowCx", nm="V_CXReq")
VL("Terminal-year capex (used; ReplPxOn = required capex at scenario price path, no historical-cost D&A floor)", "=IF(TermMode=1,IF(ReplPxOn=1,V_CXReq,MAX(V_DAT*MaintRatio,V_CXReq)),V_DAT*MaintRatio)", nm="V_CXT", bold=True)
VL("  Check: terminal capex >= GPU replacement need (net of residual credit)", '=IF(V_CXT>=(V_GRepl-V_GResid)*(1+V_g)-0.01,"OK","FLAG: below GPU replacement")', fmt="@", nm="V_CXChk")
VL("Terminal-year tax on EBIT (D&A for tax = max(D&A, steady-state replacement))", "=MAX(0,V_EBT-IF(TermMode=1,IF(ReplPxOn=1,(V_GRepl+V_NonGPU)*(1+V_g),MAX(V_DAT,(V_GRepl+V_NonGPU)*(1+V_g))),V_DAT))*TaxRate", nm="V_TXT")
VL("  Steady-state D&A (= maintenance capex: replacement + non-GPU)", "=IF(TermMode=1,IF(ReplPxOn=1,(V_GRepl+V_NonGPU)*(1+V_g),MAX(V_DAT,(V_GRepl+V_NonGPU)*(1+V_g))),V_DAT)", nm="V_Dss")
VL("  Terminal NOPAT = (EBITDA - steady-state D&A) - tax", "=V_EBT-V_Dss-V_TXT", nm="V_NOPAT")
VL("  Implied reinvestment rate (growth capex / NOPAT)", "=IF(V_NOPAT<>0,V_GrowCx/V_NOPAT,0)", fmt="0.0%", nm="V_RRimp")
VL("  Terminal ROIC (NOPAT / net PP&E Q4'30) - memo vs TermRONC", "=IFERROR(V_NOPAT/INDEX(NBV,1,28),0)", fmt="0.0%", nm="V_ROIC")
VL("Terminal-year UFCF", "=V_EBT-V_TXT-V_CXT", nm="V_UFT", bold=True)
VL("TV - Gordon growth", "=V_UFT/(WACC-V_g)", nm="V_TVG")
VL("TV - exit multiple (x terminal-year EBITDA)", "=V_EBT*ExitMult", nm="V_TVM")
VL("  Implied EV/EBITDA multiple of Gordon TV (cross-check vs ExitMult)", "=IF(V_EBT<>0,V_TVG/V_EBT,0)", fmt="0.0x", nm="V_ImpMult")
VL("  FY2030 adj. EBITDA margin (cross-check: consensus FY28 ~61%)", "=IF(V_Rev30<>0,V_EB30/V_Rev30,0)", fmt="0%", nm="V_M30")
VL("TV - blended (Gordon / exit multiple)", "=DCFWeight*V_TVG+(1-DCFWeight)*V_TVM", nm="V_TVB")
VL("  Run-off value: existing fleet, no reinvestment (after-tax EBITDA declining linearly to 0 over GPU life) + (v0.7c) after-tax GPU residual as each cohort retires (TVROResOn)", "=MAX(0,V_EBT-MAX(0,V_EBT-V_DAT)*TaxRate)*SUMPRODUCT((ROW($A$1:$A$10)<=LQ_N/4)*(1-(ROW($A$1:$A$10)-1)/(LQ_N/4))/(1+WACC)^ROW($A$1:$A$10))+TVROResOn*V_GResid*SUMPRODUCT((ROW($A$1:$A$10)<=LQ_N/4)/(1+WACC)^ROW($A$1:$A$10))", nm="V_TVRO")
VL("TV - used = IF(TVFloor, max(blend, run-off), blend)", "=IF(TVFloor=1,MAX(V_TVB,V_TVRO),V_TVB)", nm="V_TV", bold=True)
VL("  TV basis", '=IF(AND(TVFloor=1,V_TVRO>V_TVB),"RUN-OFF (reinvestment value-destroying)","Going concern")', fmt="@", nm="V_TVBasis")
VL("Discount factor at Q4'30", f"=INDEX($F${vdf}:$AG${vdf},1,28)", fmt="0.000", nm="V_DF28")
VL("PV of terminal value", "=V_TV*V_DF28", nm="V_PVTV")
VL("Sum of PV of quarterly UFCF", f"=SUM($F${vpv}:$AG${vpv})", nm="V_PVF")
VL("Closing deferred revenue at Q4'30 (prepaid, unreleased)", "=INDEX(CF_DRBal,1,28)", nm="V_DR30")
VL("  PV at Q4'30 of its post-horizon release (non-cash revenue inside terminal EBITDA; straight-line over DR_AmortQ/2 = avg remaining life)", "=V_DR30*SUMPRODUCT((ROW($A$1:$A$120)<=MAX(1,ROUND(DR_AmortQ/2,0)))/MAX(1,ROUND(DR_AmortQ/2,0))/(1+WACC)^(ROW($A$1:$A$120)/4))", nm="V_DRPV")
VL("Deferred-revenue run-off deduction (DRTermOn; cash already received inside the horizon)", "=-DRTermOn*V_DRPV*V_DF28", nm="V_DRAdj", note="v0.5 cash double-count fix")
VL("ENTERPRISE VALUE (DCF)", "=V_PVF+V_PVTV+V_DRAdj", nm="V_EV", bold=True)
VL("  Info: significant financing component (ASC 606) - model revenue/EBITDA include NO SFC gross-up and no SFC interest; DCF on cash flows only", "=0", note="SFC accretion is non-cash (FY25 $4.5M); if shown, carry below EBITDA with no cash outflow. Q2'26 calibration uses reported revenue, which may embed a small gross-up (immaterial; not stripped)")
VL("  Info (Estimated): implied financing cost of customer advances, annualized = (H1'26 interest expense 182.8 - convert interest 106.5) / avg H1 deferred revenue - UPPER BOUND (includes supplier deferred-payment interest)", "=(182.8-106.5)*2/AVERAGE(1577.5,4778.3,5975.2)", fmt="0.0%", nm="V_SFCRate", note="Q2'26 alone: (119.1-67.3)x4/avg(4,778,5,975) ~ 3.9%; FY25 SFC accretion $4.5M disclosed. Split vs supplier interest not disclosed")
VL("  TV share of EV", "=IF(V_EV<>0,V_PVTV/V_EV,0)", fmt="0%")
VL("MEMO (not used): DCF EV with 'include both' terminal (TermBase=1: Q4'30 spend for MW billing Q1'31 charged AND that MW in terminal, ~1.2x unsourced step-up) - STATIC, snap.py", "=0", nm="V_EVMemo", note="Memo only; conservative default excludes both")
VL("MEMO (not used): FD value/share with 'include both' terminal - STATIC, snap.py", "=0", fmt="0.00", nm="V_PSMemo")
for _nm, _lb in [("V_EVBaseS", "Scenario snapshot (STATIC, snap.py): DCF EV - BASE"), ("V_EVBullS", "Scenario snapshot (STATIC, snap.py): DCF EV - BULL"), ("V_EVBearS", "Scenario snapshot (STATIC, snap.py): DCF EV - BEAR")]:
    VL(_lb, "=0", nm=_nm)
for _nm, _lb in [("V_UtBaseS", "Scenario snapshot (STATIC): Q2'26 calibrated utilization - BASE"), ("V_UtBullS", "Scenario snapshot (STATIC): Q2'26 calibrated utilization - BULL"), ("V_UtBearS", "Scenario snapshot (STATIC): Q2'26 calibrated utilization - BEAR")]:
    VL(_lb, "=0", fmt="0.000", nm=_nm)
V.section("EV -> equity bridge (at valuation quarter = last quarter end on/before ValDate)")
VL("Valuation quarter index", "=MATCH(ValDate,Q_End,1)", fmt="0", nm="V_Q")
VL("Valuation quarter", "=INDEX(Q_Lab,1,V_Q)", fmt="@")
VL("Net debt at valuation quarter (model)", "=INDEX(CF_NetDebt,1,V_Q)", nm="V_NDm", note="Includes Aug-26 converts, $775M secured, cash after Q3 burn; ITM converts excluded (in shares)")
VL("SOTP: ClickHouse (low-high interpolation x (1 - discount))", "=(CH_Lo+SOTPPt*(CH_Hi-CH_Lo))*(1-SOTP_Disc)", nm="V_CH")
VL("SOTP: Toloka", "=(TO_Lo+SOTPPt*(TO_Hi-TO_Lo))*(1-SOTP_Disc)", nm="V_TO")
VL("SOTP: Avride", "=(AV_Lo+AVPt*(AV_Hi-AV_Lo))*(1-SOTP_Disc)", nm="V_AV")
VL("SOTP: TripleTen / other", "=Other_Val*(1-SOTP_Disc)", nm="V_OT")
VL("  Avride burn add-back: PV of AVBurnQ (flat) Q4'26-Q4'30 + terminal at V_g - ONLY when Avride is valued separately (V_AV > 0); credit value OR keep burn, not both", f"=IF(AND(AVBurnOn=1,V_AV>0),-AVBurnQ*(SUM($F${vdf}:$AG${vdf})+4*(1+V_g)/(WACC-V_g)*V_DF28),0)", nm="V_AVAdd", note="v0.7 fix of bull double count (v0.6 bull credited Avride and kept its burn in EBITDA)")
VL("SOTP total (incl. Avride burn add-back)", "=V_CH+V_TO+V_AV+V_OT+V_AVAdd", nm="V_SOTP", bold=True)
VL("Lease & power commitments: energy $5.32B + leases not commenced $12.05B + B/S operating leases $1.625B", "=OffBS_Energy+OffBS_Lease+OpLease", nm="V_OFF")
VL("Diluted shares at valuation quarter (Shares tab, DilPx)", "=INDEX(Sh_Dil,1,V_Q)", fmt="#,##0.0", nm="V_SHq")
V.section("Dilution solve at the implied valuation price (unrolled fixed-point; converts ITM at price P move from debt to shares; options by TSM at P)")
rr = V.r
for i, t in enumerate(["Tranche (at valuation qtr)", "", "", "", "Orig. principal o/s", "Accreted balance", "Conv. rate", "Conv. price", "Orig. principal at FP_PostQ", "Accreted at FP_PostQ",
                       "Maturity idx", "Matures in (V_Q, FP_PostQ]", "Principal at maturity", "Effective conv. px (x accretion)", "2027 plan deal", "Converts at maturity at SharePx (spot)", "Accretion x", "Conv. px - self-consistent basis (plan deals priced off PT: never ITM at value)", "Plan principal", "Plan conv. premium"]):
    if t: vws.cell(rr, 1 + i, t).font = BOLD
V.r += 1
tv0 = V.r
for i, t in enumerate(TRN):
    rr = V.r
    vws[f"A{rr}"] = t[0]
    fml(vws[f"E{rr}"], f"=INDEX(CVo_{i},1,V_Q)", "#,##0.0"); fml(vws[f"F{rr}"], f"=IF(CVDebtBasis=3,INDEX(CVc_{i},1,V_Q),INDEX(CVb_{i},1,V_Q))", "#,##0.0")
    fml(vws[f"G{rr}"], f"=INDEX(CV_Rate,{i+1})", "0.0000"); fml(vws[f"H{rr}"], f"=INDEX(CV_Px,{i+1})", "0.00")
    fml(vws[f"I{rr}"], f"=INDEX(CVo_{i},1,FP_PostQ)", "#,##0.0"); fml(vws[f"J{rr}"], f"=IF(CVDebtBasis=3,INDEX(CVc_{i},1,FP_PostQ),INDEX(CVb_{i},1,FP_PostQ))", "#,##0.0")
    _fr = f0 + i
    fml(vws[f"K{rr}"], f"=Financing!$N${_fr}", "0"); fml(vws[f"L{rr}"], f"=IF(AND(K{rr}>V_Q,K{rr}<=FP_PostQ),1,0)", "0")
    fml(vws[f"M{rr}"], f"=INDEX(CVo_{i},1,MAX(1,MIN({NQ},K{rr}-1)))", "#,##0.0"); fml(vws[f"N{rr}"], f"=Financing!$L${_fr}*Financing!$J${_fr}", "0.00")
    fml(vws[f"O{rr}"], "=1" if i >= NTR0 else "=0", "0"); fml(vws[f"P{rr}"], f"=Financing!$O${_fr}", "0"); fml(vws[f"Q{rr}"], f"=Financing!$J${_fr}", "0.000")
    fml(vws[f"R{rr}"], f"=IF(O{rr}=1,1E+9,H{rr})", "0.00"); fml(vws[f"S{rr}"], f"=Financing!$E${_fr}*O{rr}", "#,##0.0"); fml(vws[f"T{rr}"], f"=IF(O{rr}=1,H{rr}/FP_RefPx-1,0)", "0%")
    V.r += 1
tv1 = V.r - 1
O_ = f"$E${tv0}:$E${tv1}"; B_ = f"$F${tv0}:$F${tv1}"; RT = f"$G${tv0}:$G${tv1}"; PX = f"$H${tv0}:$H${tv1}"
OP_ = f"$I${tv0}:$I${tv1}"; BP_ = f"$J${tv0}:$J${tv1}"
for _n, _rg in [("V_TrO", O_), ("V_TrB", B_), ("V_TrRT", RT), ("V_TrPX", PX), ("V_TrOP", OP_), ("V_TrBP", BP_)]: name(wb, _n, "Valuation", _rg)
for _n, _c in [("V_TrMat", "K"), ("V_TrWin", "L"), ("V_TrPM", "M"), ("V_TrEff", "N"), ("V_TrPlan", "O"), ("V_TrSpot", "P"), ("V_TrAcc", "Q"), ("V_TrPXsc", "R"), ("V_TrPlanE", "S"), ("V_TrPrem", "T")]:
    name(wb, _n, "Valuation", f"${_c}${tv0}:${_c}${tv1}")
def _H(o, b): return f"((CVDebtBasis=1)*{o}+(CVDebtBasis>=2)*{b})"
H_ = _H(O_, B_)
VL("Cash at valuation quarter", "=INDEX(CF_Cash,1,V_Q)", nm="V_CASH")
VL("Secured + funding-plug debt at valuation quarter + PLACEHOLDER refundable deferred revenue (DRRefPct x DR balance; not disclosed)", "=INDEX(Sec_Bal,1,V_Q)+INDEX(ABS_Bal,1,V_Q)+INDEX(CF_PlugBal,1,V_Q)+DRRefPct*INDEX(CF_DRBal,1,V_Q)+StopBuild*StopCommitWt*(OffBS_Lease+OffBS_Energy)*MIN(1,MAX(0,INDEX(MW_Act0,1,16)-INDEX(MW_Act,1,16))/MAX(1,INDEX(MW_Act0,1,16)-INDEX(MW_Act0,1,LastHist)))", nm="V_ODEBT", note="v0.8: + StopBuild x StopCommitWt x (leases not commenced + energy commitments $17.4B) x share of post-6/30 YE27 MW additions that are cut - commitments with no modeled rent / power when building stops")
VL("Economic basic shares (issued + exchange + NVIDIA warrant)", "=INDEX(Sh_BasicX,1,V_Q)+INDEX(Sh_Warr,1,V_Q)", fmt="#,##0.00", nm="V_BASIC",
   note="Ties to 308.69M at Q3'26 = 271.86 + 15.77 + 21.07 (ATM 12.27M remaining excluded)")
VL("RSUs (+ Eigen if included)", "=RSU+EigenIn*Eigen", fmt="#,##0.00", nm="V_RSU")
def convS(P, o=None, b=None, rt=None, px=None):
    o = o or O_; b = b or B_; rt = rt or RT; px = px or PX; h = _H(o, b)
    return (f"CHOOSE(DilMethod,SUMPRODUCT({o},{rt})/1000,SUMPRODUCT(({P}>{px})*{o}*{rt})/1000,"
            f"SUMPRODUCT((({o}*{rt}/1000*{P}-{h})>0)*({o}*{rt}/1000*{P}-{h}))/{P})")
def convD(P, o=None, b=None, rt=None, px=None):
    o = o or O_; b = b or B_; rt = rt or RT; px = px or PX; h = _H(o, b)
    return f"CHOOSE(DilMethod,0,SUMPRODUCT(({P}<={px})*{h}),SUMPRODUCT({h}))"
def convFull(P, o=None, rt=None, px=None):  # ITM full conversion (if-converted), regardless of DilMethod
    o = o or O_; rt = rt or RT; px = px or PX
    return f"SUMPRODUCT(({P}>{px})*{o}*{rt})/1000"
def convNSS(P, o=None, b=None, rt=None):  # net-share settlement: cash H per note, shares for excess
    o = o or O_; b = b or B_; rt = rt or RT; h = _H(o, b)
    return f"SUMPRODUCT((({o}*{rt}/1000*{P}-{h})>0)*({o}*{rt}/1000*{P}-{h}))/{P}"
rr = V.r
for i, t in enumerate(["Iteration", "", "", "", "Weight 0%", "Chosen weight", "Weight 100%"]):
    if t: vws.cell(rr, 1 + i, t).font = BOLD
V.r += 1
wrow = V.r
vws[f"A{wrow}"] = "Off-B/S / lease weight"
fml(vws[f"E{wrow}"], "=0", "0%"); fml(vws[f"F{wrow}"], "=OffBSWt", "0%"); fml(vws[f"G{wrow}"], "=1", "0%")
V.r += 1
p0 = V.r
vws[f"A{p0}"] = "P0 = SharePx"
for c in "EFG": fml(vws[f"{c}{p0}"], "=SharePx", "0.00")
V.r += 1
NIT = 15
for k in range(1, NIT + 1):
    rr = V.r
    vws[f"A{rr}"] = f"P{k}"
    for c in "EFG":
        P = f"{c}{rr-1}"
        eq = f"(V_EV-{c}${wrow}*V_OFF-V_ODEBT-{convD(P)}+V_CASH+V_SOTP)"
        shs = f"(V_BASIC+V_RSU+Options*MAX(0,1-OptK/{P})+{convS(P)})"
        fml(vws[f"{c}{rr}"], f"=MAX(0.01,{eq}/{shs})", "0.00")
    V.r += 1
pl_ = V.r - 1
VL("Convergence |P15 - P14| (should be ~0)", f"=MAX(ABS(E{pl_}-E{pl_-1}),ABS(F{pl_}-F{pl_-1}),ABS(G{pl_}-G{pl_-1}))", fmt="0.0000", nm="V_CONV")
def tri(label, fexpr, fmt="#,##0;(#,##0);-", bold=False, nm=None):
    rr = V.r
    vws[f"A{rr}"] = label; vws[f"A{rr}"].font = BOLD if bold else BLACK
    for c in "EFG":
        fml(vws[f"{c}{rr}"], fexpr(c), fmt, bold)
    if nm: name(wb, nm, "Valuation", f"$F${rr}")
    V.r += 1
    return rr
tri("Solved value per share - FULLY DILUTED ($)", lambda c: f"={c}{pl_}", fmt="0.00", bold=True, nm="V_PS")
name(wb, "V_PSolve", "Valuation", f"$F${pl_}")
tri("  Convert shares at solved price (M)", lambda c: "=" + convS(f"{c}{pl_}"), fmt="#,##0.00")
cvd_r = tri("  Converts left in debt at solved price ($M)", lambda c: "=" + convD(f"{c}{pl_}"), nm="V_CVDEBT")
tri("  Options (TSM) at solved price (M)", lambda c: f"=Options*MAX(0,1-OptK/{c}{pl_})", fmt="#,##0.00")
tri("  Fully diluted shares at solved price (M)", lambda c: f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/{c}{pl_})+{convS(f'{c}{pl_}')}", fmt="#,##0.00", nm="V_SH")
tri("Equity value - fully diluted basis ($M)", lambda c: f"=V_EV-{c}${wrow}*V_OFF-V_ODEBT-{convD(f'{c}{pl_}')}+V_CASH+V_SOTP", bold=True, nm="V_EQ")
tri("Value per share on 308.69M-style economic basic, all converts as debt ($)", lambda c: f"=(V_EV-{c}${wrow}*V_OFF-V_ODEBT-SUMPRODUCT({H_})+V_CASH+V_SOTP)/V_BASIC", fmt="0.00", nm="V_PSB")
tri("  memo: convert shares at solved price - FULL CONVERSION of ITM notes (if-converted)", lambda c: "=" + convFull(f"{c}{pl_}"), fmt="#,##0.00", nm="V_CVFull")
tri("  memo: convert shares at solved price - NET-SHARE SETTLEMENT (treasury-style; cash = principal)", lambda c: "=" + convNSS(f"{c}{pl_}"), fmt="#,##0.00", nm="V_CVNSS")
tri("  memo: FD shares at solved price - full conversion / NSS basis", lambda c: f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/{c}{pl_})+{convFull(f'{c}{pl_}')}", fmt="#,##0.00", nm="V_SHFull")
tri("  memo: FD shares at solved price - NSS basis", lambda c: f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/{c}{pl_})+{convNSS(f'{c}{pl_}')}", fmt="#,##0.00", nm="V_SHNSS")
V.section("Convert waterfall at reference prices (existing notes, valuation quarter; Filings CapStructure: $13.425B original principal)")
rr = V.r
for i, t in enumerate(["Price $/sh", "", "", "", "Price", "Converting principal $M", "Full-conv. shares M", "Debt left (full conv.) $M", "NSS shares M", "FD shares (full conv.)", "FD shares (NSS)"]):
    if t: vws.cell(rr, 1 + i, t).font = BOLD
V.r += 1
for _px in ["46", "147", "242.81", "V_PSolve"]:
    rr = V.r; vws[f"A{rr}"] = f"At {_px}"
    fml(vws[f"E{rr}"], f"={_px}", "0.00"); P = f"$E{rr}"
    fml(vws[f"F{rr}"], f"=SUMPRODUCT(({P}>{PX})*{O_})", "#,##0"); fml(vws[f"G{rr}"], "=" + convFull(P), "#,##0.00")
    fml(vws[f"H{rr}"], f"=SUMPRODUCT(({P}<={PX})*{H_})", "#,##0"); fml(vws[f"I{rr}"], "=" + convNSS(P), "#,##0.00")
    fml(vws[f"J{rr}"], f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/{P})+G{rr}", "#,##0.00"); fml(vws[f"K{rr}"], f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/{P})+I{rr}", "#,##0.00")
    V.r += 1
V.section("POST-FUNDING per-share value (v0.7): equity + plan cash (ATM + net convert proceeds) / (FD shares + ATM shares + new & existing convert shares at the solved price; notes OTM at that price stay debt at principal). Converts at FP_PostQ.")
rr = V.r; vws[f"A{rr}"] = "Funding-plan cash added to equity ($M) = ATM proceeds + 2027 converts net"; fml(vws[f"E{rr}"], "=FP_On*(FP_ATMProc*(FP_ATMIdx<=FP_PostQ)+SUMPRODUCT((Q_Idx<=FP_PostQ)*CF_FPConv)+FP_Resize*(SUMPRODUCT((Q_Idx<=FP_PostQ)*FP_EqTop)-SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*CV_rep))-SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*ABS_Int*(1-TaxRate*(PL_Tax>0)))-CpnChg*SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*CV_cpn*(1-TaxRate*(PL_Tax>0))))", "#,##0"); name(wb, "V_FPCash", "Valuation", f"$E${rr}"); V.r += 1
rr = V.r; vws[f"A{rr}"] = "MEMO SPOT basis: new shares after the valuation quarter (ATM + top-ups at SharePx x (1 - disc) + maturity conversions at spot) (M)"; fml(vws[f"E{rr}"], "=FP_On*(FP_ATMSh*(FP_ATMIdx<=FP_PostQ)*(FP_ATMIdx>V_Q)+FP_Resize*(SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*FP_EqTopSh)+SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*CV_csh)))", "#,##0.00"); name(wb, "V_FPSh", "Valuation", f"$E${rr}"); V.r += 1
p0f = V.r; vws[f"A{p0f}"] = "Q0 = solved pre-funding price"
for c in "EFG": fml(vws[f"{c}{p0f}"], f"=MAX(0.01,{c}{pl_})", "0.00")
V.r += 1
for k in range(1, NIT + 1):
    rr = V.r; vws[f"A{rr}"] = f"Q{k}"
    for c in "EFG":
        P = f"{c}{rr-1}"
        eq = f"(V_EV-{c}${wrow}*V_OFF-V_ODEBT-{convD(P, OP_, BP_)}+V_CASH+V_SOTP+V_FPCash)"
        shs = f"(V_BASIC+V_FPSh+V_RSU+Options*MAX(0,1-OptK/{P})+{convS(P, OP_, BP_)})"
        fml(vws[f"{c}{rr}"], f"=MAX(0.01,{eq}/{shs})", "0.00")
    V.r += 1
plf = V.r - 1
psf_r = tri("POST-FUNDING value per share ($)", lambda c: f"=IF({c}{plf}<=0.011,0,{c}{plf})", fmt="0.00", bold=True, nm="V_PSF")
tri("  Post-funding FD shares at solved price (M)", lambda c: f"=V_BASIC+V_FPSh+V_RSU+Options*MAX(0,1-OptK/{c}{plf})+{convS(f'{c}{plf}', OP_, BP_)}", fmt="#,##0.00", nm="V_SHF")
tri("  Post-funding FD shares - full conversion / NSS memo (full)", lambda c: f"=V_BASIC+V_FPSh+V_RSU+Options*MAX(0,1-OptK/{c}{plf})+{convFull(f'{c}{plf}', OP_)}", fmt="#,##0.00", nm="V_SHFFull")
tri("  Post-funding FD shares - NSS", lambda c: f"=V_BASIC+V_FPSh+V_RSU+Options*MAX(0,1-OptK/{c}{plf})+{convNSS(f'{c}{plf}', OP_, BP_)}", fmt="#,##0.00", nm="V_SHFNSS")
tri("  Converts left in debt post-funding ($M)", lambda c: "=" + convD(f"{c}{plf}", OP_, BP_), nm="V_CVDEBTF")
_pb = lambda c: f"IF(FP_On=1,{c}{psf_r},{c}{pl_})"
tri("MEMO SPOT-ISSUANCE 12-month PT ($) - all new issuance at SharePx x (1 - disc); DEPENDS ON MARKET ACCESS AT TODAY'S PRICE (v0.7c headline)", lambda c: f"={_pb(c)}*(1+Ke)", fmt="0.00", bold=True, nm="V_PTSpot")
tri("  memo spot: value/share ($)", lambda c: f"={_pb(c)}", fmt="0.00", nm="V_PSSpot")
# ---------------- v0.7d SELF-CONSISTENT issuance price (headline) ----------------
V.section("SELF-CONSISTENT HEADLINE (v0.7d): every new share (ATM, equity top-ups, re-funding of converts that no longer convert) is issued at the solved PT itself, P = v x (1 + Ke) x (1 - SC_Disc); 2027 plan converts priced off P (conversion px = P x (1 + premium) -> never in the money at value, stay debt). Fixed point solved by unrolled iteration (linear solve each step). Cash raised is the same $ as the spot plan (the runway needs the same dollars); only the share count changes.")
def _scflag(P): return f"(({P})>=V_TrEff)*(1-V_TrPlan)"
def SC_XR(P): return f"SUMPRODUCT(V_TrWin*V_TrPM*V_TrAcc*(V_TrSpot-{_scflag(P)}))"
def SC_MS(P): return f"SUMPRODUCT(V_TrWin*V_TrPM*V_TrRT/1000*{_scflag(P)})"
VL("Equity $ issued after the valuation quarter through FP_PostQ (spot plan: ATM proceeds + top-ups) - same $ on every basis", "=FP_On*(FP_ATMProc*(FP_ATMIdx<=FP_PostQ)*(FP_ATMIdx>V_Q)+FP_Resize*SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*FP_EqTop))", nm="V_EqIss0")
VL("Check: top-up equity in or before the valuation quarter (must be 0; those shares sit in V_BASIC)", "=SUMPRODUCT((Q_Idx<=V_Q)*FP_EqTop)", nm="V_EqIssPre")
def SCROWS(ws, cols, wexpr, evexpr, start, nit, prow_cols, nm_prefix=None):
    """cols: value columns; prow_cols: matching price columns. returns last row"""
    for c, pc in zip(cols, prow_cols):
        fml(ws[f"{c}{start}"], "=SharePx", "0.00"); fml(ws[f"{pc}{start}"], "=SharePx", "0.00")
    for k in range(1, nit + 1):
        rr = start + k
        for c, pc in zip(cols, prow_cols):
            v = f"{c}{rr-1}"; P = f"{pc}{rr}"
            fml(ws[f"{pc}{rr}"], f"=MAX(0.01,{v}*(1+Ke)*(1-SC_Disc))", "0.00")
            w = wexpr(c)
            eqp = f"({evexpr}-{w}*V_OFF-V_ODEBT-{convD(v, 'V_TrOP', 'V_TrBP', 'V_TrRT', 'V_TrPXsc')}+V_CASH+V_SOTP+V_FPCash)"
            iss = f"MAX(0,V_EqIss0+FP_On*{SC_XR(P)})/((1+Ke)*(1-SC_Disc))"
            den = f"(V_BASIC+V_RSU+Options*MAX(0,1-OptK/{v})+{convS(v, 'V_TrOP', 'V_TrBP', 'V_TrRT', 'V_TrPXsc')}+FP_On*{SC_MS(P)})"
            fml(ws[f"{c}{rr}"], f"=MAX(0.01,({eqp}-{iss})/{den})", "0.00")
    return start + nit
rr = V.r
for i, t in enumerate(["Iteration (value v | issue price P)", "", "", "", "v: weight 0%", "v: chosen", "v: weight 100%", "P: 0%", "P: chosen", "P: 100%"]):
    if t: vws.cell(rr, 1 + i, t).font = BOLD
V.r += 1
_s0 = V.r
NSC = 30
vws[f"A{_s0}"] = "S0 = SharePx (start high; converges to the largest fixed point)"
_wx = lambda c: f"{c}${wrow}"
_sl = SCROWS(vws, "EFG", _wx, "V_EV", _s0, NSC, "HIJ")
for k in range(1, NSC + 1): vws[f"A{_s0+k}"] = f"S{k}"
V.r = _sl + 1
VL("Convergence |S30 - S29| (should be ~0)", f"=MAX(ABS(E{_sl}-E{_sl-1}),ABS(F{_sl}-F{_sl-1}),ABS(G{_sl}-G{_sl-1}))", fmt="0.0000", nm="V_CONVSC")
_scv = {"E": f"E{_sl}", "F": f"F{_sl}", "G": f"G{_sl}"}; _scp = {"E": f"H{_sl}", "F": f"I{_sl}", "G": f"J{_sl}"}
pssc_r = tri("SELF-CONSISTENT value per share ($) (0 = no positive fixed point: existing equity under water at any issue price)", lambda c: f"=IF({_scv[c]}<=0.011,0,{_scv[c]})", fmt="0.00", bold=True, nm="V_PSSC")
tri("SELF-CONSISTENT 12-month PT ($) = issue price of all new shares", lambda c: f"={c}{pssc_r}*(1+Ke)", fmt="0.00", bold=True, nm="V_PTSC")
tri("  Issue price used in the last iteration ($)", lambda c: f"={_scp[c]}", fmt="0.00", nm="V_PISC")
tri("  New shares issued (M): equity $ / P + maturity conversions at P", lambda c: f"=FP_On*(MAX(0,V_EqIss0+{SC_XR(_scp[c])})/{_scp[c]}+{SC_MS(_scp[c])})", fmt="#,##0.0", nm="V_NewShSC")
tri("  of which: extra equity $ to repay converts that convert at spot but not at P ($M)", lambda c: f"=FP_On*{SC_XR(_scp[c])}", nm="V_XRSC")
tri("  Post-funding FD shares at the self-consistent price (M)", lambda c: f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/{_scv[c]})+{convS(_scv[c], OP_, BP_, RT, 'V_TrPXsc')}+{c}{V.r-2}", fmt="#,##0.0", nm="V_SHSC")
tri("  2027 plan converts: shares to reserve at self-consistent terms (conv. px = P x (1 + premium)) (M)", lambda c: f"=SUMPRODUCT(V_TrPlanE/(1+V_TrPrem))/{_scp[c]}", fmt="#,##0.0", nm="V_CVResSC")
# ---------------- v0.8c equity value per CURRENT share before issuance + reconciliation to the self-consistent value ----------------
V.section("v0.8c CROSS-CHECK: equity value per CURRENT FD share before issuance = (EV - net debt incl. existing converts at carrying value) / current FD shares; new money (funding gap) enters at fair value (proceeds = claim, net 0). Reconciled to the self-consistent value per share (exact identity)")
_hX = _H("V_TrO", "V_TrB"); _hP = _H("V_TrOP", "V_TrBP")
VL("EV (DCF)", "=V_EV", nm="X_EV")
VL("Funding gap $ after the valuation quarter = equity (ATM + top-ups) + 2027 plan converts (net) + equity to refund converts that do not convert at the SC price - i.e. need NOT covered by EBITDA, prepayments, ABS", "=V_EqIss0+V_XRSC+FP_On*SUMPRODUCT((Q_Idx<=FP_PostQ)*CF_FPConv)", nm="X_Gap")
VL("Existing converts at carrying value at the valuation quarter (excl. 2027 plan deals)", f"=SUMPRODUCT({_hX}*(1-V_TrPlan))", nm="X_CVc")
VL("Equity value before issuance = EV - OffBSWt x commitments - other debt (secured, plug, ABS, refundable DR, stop-build commitments) - existing converts at carrying + cash + SOTP", "=V_EV-OffBSWt*V_OFF-V_ODEBT-X_CVc+V_CASH+V_SOTP", nm="X_E0")
VL("Current FD shares (economic basic + RSUs + options by TSM at the SC value)", "=V_BASIC+V_RSU+Options*MAX(0,1-OptK/MAX(0.01,V_PSSC))", fmt="#,##0.0", nm="X_N0")
VL("EQUITY VALUE PER CURRENT SHARE BEFORE ISSUANCE ($)", "=X_E0/X_N0", fmt="0.00", nm="X_PS0", bold=True)
VL("  x (1 + Ke) = 12-month equivalent ($), compare with the self-consistent PT", "=X_PS0*(1+Ke)", fmt="0.00", nm="X_PT0")
VL("  memo (harsh, double counts capex already in EV): (equity value - funding gap) / current shares", "=(X_E0-X_Gap)/X_N0", fmt="0.00", nm="X_PSharsh")
_v = "MAX(0.01,V_PSSC)"
VL("Recon 1: existing converts - carrying at V_Q vs cash repaid in window, refunded at SC price, and still outstanding at FP_PostQ (accretion; conversions) per current share", f"=(X_CVc-SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*CV_rep)-V_XRSC-CHOOSE(DilMethod,0,SUMPRODUCT(({_v}<=V_TrPXsc)*{_hP}*(1-V_TrPlan)),SUMPRODUCT({_hP}*(1-V_TrPlan))))/X_N0", fmt="0.00;-0.00", nm="X_R1")
VL("Recon 2: 2027 plan converts - net proceeds vs accreted claim at FP_PostQ (fees + accretion = option / financing cost)", f"=(FP_On*SUMPRODUCT((Q_Idx<=FP_PostQ)*CF_FPConv)-CHOOSE(DilMethod,0,SUMPRODUCT(({_v}<=V_TrPXsc)*{_hP}*V_TrPlan),SUMPRODUCT({_hP}*V_TrPlan)))/X_N0", fmt="0.00;-0.00", nm="X_R2")
VL("Recon 3: notes converting into shares at the SC value (share count x value)", f"=-{_v}*({convS(_v, 'V_TrOP', 'V_TrBP', 'V_TrRT', 'V_TrPXsc')}+FP_On*{SC_MS('MAX(0.01,V_PISC)')})/X_N0", fmt="0.00;-0.00", nm="X_R3")
VL("Recon 4: ABS interest charged (after tax when taxable)", "=-SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*ABS_Int*(1-TaxRate*(PL_Tax>0)))/X_N0", fmt="0.00;-0.00", nm="X_R4")
VL("Recon 4b (v0.8d): convert cash coupons after V_Q through FP_PostQ charged (after tax when taxable)", "=-CpnChg*SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*CV_cpn*(1-TaxRate*(PL_Tax>0)))/X_N0", fmt="0.00;-0.00", nm="X_R4b")
VL("Recon 5: issue premium if new equity is sold at the 12-month PT v x (1 + Ke) (v0.8d headline SC_NoPrem = 1 -> issued at v, R5 = 0)", "=MAX(0,V_EqIss0+V_XRSC)*(1-1/((1+Ke)*(1-SC_Disc)))/X_N0", fmt="0.00;-0.00", nm="X_R5")
VL("Recon 6: residual (fixed point floored at $0.01 / not converged / MAX(0) on issuance; ~0 when converged)", "=V_PSSC-X_PS0-X_R1-X_R2-X_R3-X_R4-X_R4b-X_R5", fmt="0.00;-0.00", nm="X_R6")
VL("Shares needed (issued + convert reserve) vs authority", "=V_AuthNeed", fmt="#,##0.0", nm="X_AuthNeed")
VL("CONSTRAINT FLAG", '=IF(V_AuthNeed>AuthAvail,"exceeds authority, constraint not answer","within authority")', fmt="@", nm="X_Flag")
# ---------------- path memo (chosen weight) ----------------
V.section("MEMO PATH basis (DEPENDS ON MARKET ACCESS AT TODAY'S PRICE): issue price moves linearly from SharePx at the valuation quarter to the path's own PT over FP_PathQ quarters (x (1 - disc) as spot), flat after; maturity conversions tested at the path price. Chosen weight only. Fixed point by iteration.")
def _cl(x): return f"(({x})>=1)+(({x})>0)*(({x})<1)*({x})"
_xq = "(Q_Idx-V_Q)/FP_PathQ"; _xn = "(V_TrMat-V_Q)/FP_PathQ"
rr = V.r; vws.cell(rr, 5, "PT end (P)").font = BOLD; vws.cell(rr, 6, "v (chosen)").font = BOLD; V.r += 1
_p0 = V.r; vws[f"A{_p0}"] = "T0"; fml(vws[f"F{_p0}"], "=SharePx", "0.00"); fml(vws[f"E{_p0}"], "=SharePx", "0.00")
NPT = 40
for k in range(1, NPT + 1):
    rr = _p0 + k; vws[f"A{rr}"] = f"T{k}"
    v = f"F{rr-1}"; Pe = f"E{rr}"
    fml(vws[f"E{rr}"], f"=MAX(0.01,{v}*(1+Ke))", "0.00")
    PPq = f"(SharePx+({Pe}-SharePx)*({_cl(_xq)}))"
    PPa = f"(SharePx+({Pe}-SharePx)*MAX(0,MIN(1,(FP_ATMIdx-V_Q)/FP_PathQ)))"
    PPn = f"(SharePx+({Pe}-SharePx)*({_cl(_xn)}))"
    fl = f"({PPn}>=V_TrEff)*(1-V_TrPlan)"
    nsh = (f"FP_On*(FP_ATMProc*(FP_ATMIdx<=FP_PostQ)*(FP_ATMIdx>V_Q)/({PPa}*(1-FP_ATMDisc))+FP_Resize*SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*FP_EqTop/({PPq}*(1-FP_EqDisc)))"
           f"+SUMPRODUCT(V_TrWin*V_TrPM*(V_TrRT/1000*{fl}+V_TrAcc*(V_TrSpot-{fl})/({PPn}*(1-FP_EqDisc)))))")
    eqp = f"(V_EV-OffBSWt*V_OFF-V_ODEBT-{convD(v, OP_, BP_)}+V_CASH+V_SOTP+V_FPCash)"
    den = f"(V_BASIC+V_RSU+Options*MAX(0,1-OptK/{v})+{convS(v, OP_, BP_)}+{nsh})"
    fml(vws[f"F{rr}"], f"=MAX(0.01,{eqp}/{den})", "0.00")
    if k == NPT: fml(vws[f"G{rr}"], "=" + nsh, "#,##0.0")
_pl = _p0 + NPT
V.r = _pl + 1
VL("MEMO PATH value per share ($)", f"=IF(F{_pl}<=0.011,0,F{_pl})", fmt="0.00", nm="V_PSPath")
VL("MEMO PATH 12-month PT ($) - DEPENDS ON MARKET ACCESS AT TODAY'S PRICE", "=V_PSPath*(1+Ke)", fmt="0.00", nm="V_PTPath", bold=True)
VL("  path end price used ($)", f"=E{_pl}", fmt="0.00", nm="V_PEndPath")
VL("  path new shares (M)", f"=G{_pl}", fmt="#,##0.0", nm="V_NewShPath")
VL("  path convergence |T40 - T39|", f"=ABS(F{_pl}-F{_pl-1})", fmt="0.0000", nm="V_CONVPT")
# ---------------- HEADLINE ----------------
V.section("HEADLINE (PTBasis: 1 = self-consistent [default], 2 = spot memo, 3 = path memo)")
psh_r = tri("HEADLINE value per share ($)", lambda c: (f"=CHOOSE(PTBasis,{c}{pssc_r},V_PSSpot,V_PSPath)" if c == "F" else f"=CHOOSE(PTBasis,{c}{pssc_r},{_pb(c)},{c}{pssc_r})"), fmt="0.00", bold=True, nm="V_PSHead")
tri("12-month price target ($) - HEADLINE", lambda c: f"={c}{psh_r}*(1+Ke)", fmt="0.00", bold=True, nm="V_PT")
tri("6-month value/share ($) = x (1 + Ke)^0.5", lambda c: f"={c}{psh_r}*(1+Ke)^0.5", fmt="0.00", nm="V_PT6")
tri("3-year value/share ($) = x (1 + Ke)^3 (no dividends)", lambda c: f"={c}{psh_r}*(1+Ke)^3", fmt="0.00", nm="V_PT36")
tri("Upside / (downside) vs SharePx", lambda c: f"={c}{psh_r}*(1+Ke)/SharePx-1", fmt="0%", nm="V_UP")
VL("HEADLINE new shares issued after the valuation quarter through FP_PostQ (M)", "=CHOOSE(PTBasis,V_NewShSC,V_FPSh,V_NewShPath)", fmt="#,##0.0", nm="V_NewShHead")
VL("HEADLINE 2027 plan converts: shares to reserve (M) (SC: at P x (1 + premium); spot/path: at Aug-26-style terms off SharePx)", "=CHOOSE(PTBasis,V_CVResSC,SUMPRODUCT(V_TrPlanE*V_TrRT)/1000,SUMPRODUCT(V_TrPlanE*V_TrRT)/1000)", fmt="#,##0.0", nm="V_CVResHead")
VL("AUTHORITY: shares needed (new shares + plan-convert reserve) vs ~130M available (AuthAvail; v0.8c)", "=V_NewShHead+V_CVResHead", fmt="#,##0.0", nm="V_AuthNeed", bold=True)
VL("  shortfall vs authority (M) (> 0 needs EGM approval)", "=MAX(0,V_AuthNeed-AuthAvail)", fmt="#,##0.0", nm="V_AuthGap")
VL("Net debt (secured + plug + converts left in debt - cash), chosen weight", f"=V_ODEBT+F{cvd_r}-V_CASH", nm="V_ND")
dil_r = V.r
vws[f"A{dil_r}"] = "Price used for quarterly dilution (DilPx)"; fml(vws[f"E{dil_r}"], "=IF(DilPxMode=1,SharePx,V_PSolve)", "0.00", True)
name(wb, "DilPx", "Valuation", f"$E${dil_r}")
V.r += 1
V.section("Cross-checks")
VL("NTM adj. EBITDA (4 quarters after valuation quarter)", "=SUMPRODUCT(PL_EBITDA*(Q_Idx>V_Q)*(Q_Idx<=V_Q+4))", nm="V_NTM")
VL("EV / NTM adj. EBITDA", "=IF(V_NTM>0,V_EV/V_NTM,0)", fmt="0.0x")
VL("EV / FY2027 adj. EBITDA", "=V_EV/SUMIF(Q_Yr,2027,PL_EBITDA)", fmt="0.0x")
exec(open('/workspace/nebius/model/build_part5.py').read())
