# ================= PART 5: market EV & comps, Q4 bridge, dep-yield, KPIs, Outputs, Sensitivity, Checks, Sources, README
from openpyxl.formatting.rule import FormulaRule
reg("Sites", "MW_En", BLK['Energized'][0]); reg("Sites", "MW_Gd", BLK['Guided'][0])
# ---- Valuation: market EV at SharePx & EV/MW comps
V.section("Market EV at SharePx (dilution at $242.81; ITM converts in shares and OUT of debt)")
VL("FD shares at SharePx excl. unsold ATM (M)", f"=V_BASIC+V_RSU+Options*MAX(0,1-OptK/SharePx)+{convS('SharePx')}", fmt="#,##0.00", nm="V_FDmkt",
   note="Check vs Filings 369.2M (Checks tab)")
VL("  ITM convert shares at SharePx (M)", f"={convS('SharePx')}", fmt="#,##0.00", nm="V_ITMsh")
VL("  ITM convert principal at SharePx ($M orig.)", f"=SUMPRODUCT((SharePx>{PX})*{O_})", nm="V_ITMprin")
VL("  Converts left in debt at SharePx ($M accreted)", f"={convD('SharePx')}", nm="V_CVDmkt")
V.section("MARKET EV - Filings CapStructure basis (6/30 balance sheet + Aug notes + Aug exchange); NO forecast cash")
VL("Market equity = SharePx x economic basic 308.69M ($M)", "=SharePx*308.69", nm="V_MKTEQ", note="271.86 issued + 15.77 exchange + 21.07 NVIDIA warrant")
VL("+ Older converts, carrying pro forma for Aug exchange", "=CS_ConvCarry")
VL("+ Aug-2026 converts (2030s + 2034s) gross", "=AugGross")
VL("+ Debt, current (6/30)", "=CS_CurDebt")
VL("+ Secured loan (if drawn; cash-neutral)", "=SecAmt*LoanOn")
VL("- Unrestricted cash 6/30 (excl. $1.06B restricted)", "=-CS_Cash")
VL("- Aug converts proceeds held as cash", "=-AugGross")
VL("- Secured loan proceeds held as cash", "=-SecAmt*LoanOn")
VL("MARKET EV ex-leases ($M)", "=V_MKTEQ+CS_ConvCarry+AugGross+CS_CurDebt-CS_Cash-AugGross", nm="V_MKTEV", bold=True, note="Target ~$74.7B at $242.81 (Filings)")
VL("Memo: FD-basis market EV (FD equity at SharePx; ITM older converts out of debt)", "=SharePx*V_FDmkt+AugGross+CS_CurDebt-CS_Cash-AugGross", nm="V_MKTEV_FD", note="All older converts ITM at $242.81 -> in shares; Aug notes OTM -> debt")
VL("Market EV / NTM adj. EBITDA", "=IF(V_NTM>0,V_MKTEV/V_NTM,0)", fmt="0.0x")
VL("Energized IT MW at valuation qtr (Capacity, live scenario)", "=INDEX(MW_En,1,V_Q)", fmt="#,##0", nm="V_ENMW")
VL("Contracted IT MW at valuation qtr", "=INDEX(MW_Contr,1,V_Q)", fmt="#,##0", nm="V_CTMW")
V.section("EV per ENERGIZED MW (primary capacity multiple) - three ways; EV per contracted MW = indicative only")
rr = V.r
for c, t in zip("ABEFGH", ["", "", "Nebius market", "Nebius DCF", "CoreWeave", "NBIS / CRWV"]):
    if t: vws[f"{c}{rr}"] = t; vws[f"{c}{rr}"].font = BOLD
V.r += 1
def comp(label, neb_add, cr_add, den_n, den_c, nm=None, note=""):
    rr = V.r
    vws[f"A{rr}"] = label
    fml(vws[f"E{rr}"], f"=IFERROR((V_MKTEV{neb_add})/{den_n},0)", "0.0")
    fml(vws[f"F{rr}"], f"=IFERROR((V_EV{neb_add})/{den_n},0)", "0.0")
    fml(vws[f"G{rr}"], f"=IFERROR((CRWV_EV{cr_add})/{den_c},0)", "0.0")
    fml(vws[f"H{rr}"], f"=IFERROR(E{rr}/G{rr},0)", "0.00x")
    vws[f"D{rr}"] = note; vws[f"D{rr}"].font = ITAL
    if nm: name(wb, nm, "Valuation", f"$E${rr}")
    V.r += 1
comp("$M EV / energized MW - ex-leases", "", "", "V_ENMW", "CRWV_MW", "V_EVMW1")
comp("$M EV / energized MW - incl. commenced leases both sides", "+OpLease", "+CRWV_Lease", "V_ENMW", "CRWV_MW", "V_EVMW2", "CRWV lease liab. placeholder (GPU Market)")
comp("$M EV / energized MW - incl. all commitments", "+OpLease+OffBS_Lease+OffBS_Energy", "+CRWV_Lease+CRWV_Commit", "V_ENMW", "CRWV_MW", "V_EVMW3", "CRWV commitments placeholder")
comp("$M EV / contracted MW (INDICATIVE ONLY)", "", "", "V_CTMW", "CRWV_Contr", "V_EVCMW", "pipeline, not capacity")
VL("Peer-implied EV (CRWV EV/energized MW x Nebius energized MW)", "=CRWV_EV/CRWV_MW*V_ENMW", nm="V_PEEREV")
VL("DCF EV / market EV", "=IFERROR(V_EV/V_MKTEV,0)", fmt="0.00x")
VL("Market EV / FY27 adj. EBITDA (model)", "=IFERROR(V_MKTEV/SUMIF(Q_Yr,2027,PL_EBITDA),0)", fmt="0.0x", nm="V_EVEB27")

# ---- Contracts: Q4'26 consensus bridge
ct.section("Q4'26 consensus bridge (group revenue, $M)")
cws = ct.ws
def BR(label, f, fmt="#,##0;(#,##0);-", nm=None, bold=False, note=""):
    rr = ct.r
    cws[f"A{rr}"] = label; cws[f"A{rr}"].font = BOLD if bold else BLACK
    fml(cws[f"E{rr}"], f, fmt, bold); cws[f"D{rr}"] = note; cws[f"D{rr}"].font = ITAL
    if nm: name(wb, nm, "Contracts", f"$E${rr}")
    ct.r += 1
    return rr
BR("Consensus Q4'26 revenue", "=Cons_Q4", nm="BR_Cons", note="Skeptic, Estimated")
BR("less Q2'26 group revenue", "=-Q2GroupRev", note="Reported")
BR("Gap to bridge (consensus Q4 - Q2)", "=BR_Cons-Q2GroupRev", nm="BR_Gap", bold=True)
BR("Microsoft Q2'26 run rate (TopCust toggle)", "=MS_Q2RR", nm="BR_MSQ2", note="Meta=24% -> MS 21% ~$121M; MS=24% -> ~$138M")
BR("Microsoft Q4'26 model revenue (capped)", "=INDEX(MS_Rev,1,12)", nm="BR_MSQ4", note="T1-T4 only; T5-9 need Bloom (no Q4'26 billing)")
BR("Microsoft increment", "=BR_MSQ4-BR_MSQ2", nm="BR_MSInc")
BR("Meta #1 Q2'26 run rate (TopCust toggle)", "=IF(TopCust=1,0.24,0.21)*574.9", nm="BR_MetaQ2")
BR("Meta #1 increment (flat ~$144M/qtr)", "=Meta1Q-BR_MetaQ2", nm="BR_MetaInc")
BR("Oct-1 list-price uplift - low", "=PxLine_Lo", nm="BR_PxLo", note="GPU Market, Derived, $40-125M/qtr")
BR("Oct-1 list-price uplift - high", "=PxLine_Hi", nm="BR_PxHi")
BR("RESIDUAL from Meta #2 / other deals / unseen sites - at high price line", "=BR_Gap-BR_MSInc-BR_MetaInc-BR_PxHi", nm="BR_ResLo", bold=True)
BR("RESIDUAL - at low price line", "=BR_Gap-BR_MSInc-BR_MetaInc-BR_PxLo", nm="BR_ResHi", bold=True)
BR("Goal-seek: unseen billing MW implied @ NewDealHi ($25M/MW-yr) - range", "=BR_ResLo/(NewDealHi/4)", fmt="#,##0", nm="BR_MW25")
BR("Goal-seek: unseen billing MW implied @ NewDealLo ($20M/MW-yr) - high residual", "=BR_ResHi/(NewDealLo/4)", fmt="#,##0", nm="BR_MW20")
BR("Goal-seek: unseen billing MW implied @ base rate ($12M/MW-yr) - high residual", "=BR_ResHi/(BaseRateMW/4)", fmt="#,##0", nm="BR_MW12")
BR("Check vs Skeptic/Capacity unseen estimate 83-200 MW (OK if new-deal range overlaps)", '=IF(AND(BR_MW25<=UnseenChkHi,BR_MW20>=UnseenChkLo),"OK - overlaps 83-200 MW","CHECK")', fmt="@", nm="BR_Chk")
BR("Model Q4'26 group revenue", "=INDEX(PL_Rev,1,12)", nm="BR_Model")
BR("Model vs consensus Q4'26", "=BR_Model-BR_Cons", nm="BR_vsCons")

# ---- Capex_DA: in-service PP&E roll & depreciation yield check
C.section("In-service PP&E roll-forward & depreciation-yield check (D&A / avg in-service gross PP&E, annualized)")
gx = C.row("Cash capex + capitalized interest", "$M", "Modeling", "", fn=lambda n: f"={X('Capex', n)}+{X('CapInt', n)}")
me = C.r
trn = C.row("Transfer into service", "$M", "Filings/Modeling", "Hist Derived = change in in-service gross; fcst = CIPRel x Q2'26 CIP + post-6/30 capex lagged PISLag+GPULead-1 (placed in service at billing)",
            fn=lambda n: (f"=IF({qc(n)}$5<=LastHist,IF({qc(n)}$5>=9,INDEX(H_ISG,1,{qc(n)}$5)-INDEX(H_ISG,1,{qc(n)}$5-1),0),"
                          f"IF({qc(n)}$5=LastHist+1,CIPRel1,IF({qc(n)}$5=LastHist+2,CIPRel2,0))*INDEX(H_CIP,1,LastHist)"
                          f"+IF({qc(n)}$5-(PISLag+GPULead-1)>LastHist,INDEX($E${gx}:$AG${gx},1,MAX(1,{qc(n)}$5-(PISLag+GPULead-1))),0))"))
reg("Capex_DA", "Transfer", trn)
cip = C.row("Assets not yet in use (CIP)", "$M", "Filings/Modeling", "Hist Note 7", kind="stock",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_CIP,1,{qc(n)}$5),{qc(n-1)}{C.r}+{qc(n)}{gx}-{qc(n)}{trn})")
isg = C.row("In-service gross PP&E (no retirements)", "$M", "Filings/Modeling", "Hist Derived (Note 7)", kind="stock",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_ISG,1,{qc(n)}$5),{qc(n-1)}{C.r}+{qc(n)}{trn})")
reg("Capex_DA", "ISG", isg)
ppd = C.row("PP&E depreciation - reported basis", "$M", "Filings/Modeling", "Hist Note 7; fcst = model D&A - OtherDA",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,INDEX(H_PPDep,1,{qc(n)}$5),{qc(n)}{dtm}-OtherDA)")
ppx = C.row("PP&E depreciation - ex 4->5 yr life change", "$M", "Filings/Modeling", "Hist + Note 1 benefit (43.1 / 43.0); fcst GPU dep x (LQ/16 - 1) added back",
            fn=lambda n: f"=IF({qc(n)}$5<=LastHist,{qc(n)}{ppd}+INDEX(H_DAex,1,{qc(n)}$5)-INDEX(H_DA,1,{qc(n)}$5),{qc(n)}{ppd}+{qc(n)}{dep['H']}*(LQ_H/16-1)+{qc(n)}{dep['B']}*(LQ_B/16-1)+{qc(n)}{dep['N']}*(LQ_N/16-1))")
def yld(n, num):
    if n < 2: return None
    return f'=IF(AND({qc(n-1)}{isg}>0,{qc(n)}{isg}>0),{qc(n)}{num}*4/AVERAGE({qc(n-1)}{isg},{qc(n)}{isg}),"")'
y_rep = C.row("Depreciation yield - REPORTED basis", "%", "Modeling", "~17-18% Q1-Q2'26", kind="none", fmt="0.0%", fn=lambda n: yld(n, ppd))
y_ex = C.row("Depreciation yield - EX LIFE CHANGE (primary from Q1'26)", "%", "Modeling", "~20-22% Q1-Q2'26", kind="none", fmt="0.0%", bold=True, fn=lambda n: yld(n, ppx))
reg("Capex_DA", "DepYld", y_ex)
avail = "OR(" + ",".join(f"AND({{c}}$5-AvailIdx_{k}>=0,{{c}}$5-AvailIdx_{k}<=2)" for k in ["H200", "B200", "GB300", "VR"]) + ")"
yfl = C.row("FLAG (actuals): ex-change yield drops while new chips recently reached customers (0-2 qtrs after H200/B200/GB300/VR availability)", "flag", "Modeling", "", kind="none", fmt="0",
            fn=lambda n: (None if n < 3 else f'=IF(AND({qc(n)}$5<=LastHist,ISNUMBER({qc(n)}{y_ex}),ISNUMBER({qc(n-1)}{y_ex}),{qc(n)}{y_ex}<{qc(n-1)}{y_ex},{avail.format(c=qc(n))}),1,0)'))
reg("Capex_DA", "DepYldFlag", yfl)
q3c = C.row("Q3'26 check: large transfer into service expected after Jul-Aug Vineland tranches", "text", "Filings", "Enter Q3TransAct when Q3 reported", kind="none", fmt="@",
            fn=lambda n: ('=IF(Q3TransAct=0,"PENDING (model "&TEXT(' + f"{qc(n)}{trn}" + ',"#,##0")&")",IF(Q3TransAct<Q3TransMin,"FLAG: small transfer","OK"))' if n == 11 else None))
name(wb, "Q3TransChk", "Capex_DA", f"$P${q3c}")

# ---- CashFlow: funding runway (no plug)
K.section("FUNDING RUNWAY (no funding plug): start = 6/30 unrestricted cash; each later raise its own line; only NEW prepayments counted as cash")
cfw = cf.ws
def RW(label, fn, bold=False, fmt="#,##0;(#,##0);-", note="", kind="stock"):
    return K.row(label, "$M", "Modeling", note, fn=fn, bold=bold, fmt=fmt, kind=kind)
st0 = RW("Opening liquidity: unrestricted cash 6/30/26", lambda n: ("=CS_Cash" if n == 10 else None), note="Excludes $1,056M restricted", kind="none")
r_aug = RW("+ Aug-2026 converts gross ($3.45B + $2.30B; closed 8/24/26, NOT in 6/30 cash -> added ONCE here at Q3'26)", lambda n: ("=AugGross" if n == 11 else ("=0" if n > 10 else None)), kind="flow")
r_ln = RW("+ $775M asset-backed facility (LoanOn: base 0 = drawn amount undisclosed; 1 = $775M sensitivity)", lambda n: ("=SecAmt*LoanOn" if n == 11 else ("=0" if n > 10 else None)), kind="flow")
r_atm = RW("+ ATM: legacy toggle ATMOn (only when FP_On = 0) + funding-plan ATM at FP_ATMIdx", lambda n: (f"=IF(FP_On=0,ATMOn*ATM_Rem*SharePx*({qc(n)}$5=11),0)+{X('CF_ATM', n)}" if n > 10 else None), kind="flow")
reg("CashFlow", "Liq_Aug", r_aug)
r_fpc = RW("+ 2027 funding-plan converts (net; Aug-26 terms; coupons flow through FCF below)", lambda n: (f"={X('CF_FPConv', n)}" if n > 10 else None), kind="flow")
r_fcf = RW("+ Levered FCF ex funding-plug interest (incl. NEW prepayments; DR release non-cash)", lambda n: (f"={X('CF_FCF', n)}+{X('CF_PlugIntRow', n)}" if n > 10 else None), kind="flow")
r_rep = RW("- Convert / secured repayments", lambda n: (f"=-{X('CV_rep', n)}+MIN(0,{X('Sec_Flow', n)})" if n > 11 else ("=0" if n == 11 else None)), kind="flow")
r_abs = RW("+ v0.8b ABS on contracted MW: draws - amortisation (interest in FCF)", lambda n: (f"={X('ABS_Draw', n)}-{X('ABS_Rep', n)}" if n > 10 else None), kind="flow")
me = K.r + 2
r_eq = RW("+ Re-sized plan (FP_Resize): equity top-up to FP_MinBuf at FP_EqPx (after ATM and converts)", lambda n, me=me: (f"=FP_On*FP_Resize*({qc(n)}$5<=FP_TopEnd)*MAX(0,FP_MinBuf-({qc(n-1)}{me}+{qc(n)}{r_aug}+{qc(n)}{r_ln}+{qc(n)}{r_atm}+{qc(n)}{r_fpc}+{qc(n)}{r_fcf}+{qc(n)}{r_rep}+{qc(n)}{r_abs}))" if n > 10 else None), kind="flow")
reg("CashFlow", "FP_EqTop", r_eq)
r_eqs = RW("  equity top-up shares issued (M) at SharePx x (1 - FP_EqDisc)", lambda n: (f"={qc(n)}{r_eq}/(SharePx*(1-FP_EqDisc))" if n > 10 else "=0"), kind="flow", fmt="#,##0.00")
reg("CashFlow", "FP_EqTopSh", r_eqs)
assert K.r == me
liq = RW("LIQUIDITY (end of quarter)", lambda n, me=me: ("=CS_Cash" if n == 10 else (f"={qc(n-1)}{me}+{qc(n)}{r_aug}+{qc(n)}{r_ln}+{qc(n)}{r_atm}+{qc(n)}{r_fpc}+{qc(n)}{r_fcf}+{qc(n)}{r_rep}+{qc(n)}{r_abs}+{qc(n)}{r_eq}" if n > 10 else None)), bold=True)
reg("CashFlow", "Liq", liq)
# ---- v0.7d: plan issuance into Sh_Basic at the HEADLINE price basis
def _pp(n): return f"(SharePx+(V_PEndPath-SharePx)*MAX(0,MIN(1,({qc(n)}$5-V_Q)/FP_PathQ)))"
def _plan_sh(n):
    iss = f"CHOOSE(PTBasis,V_PISC,SharePx*(1-FP_EqDisc),{_pp(n)}*(1-FP_EqDisc))"
    atm_ = f"CHOOSE(PTBasis,V_PISC,FP_ATMPx,{_pp(n)}*(1-FP_ATMDisc))"
    tst = f"CHOOSE(PTBasis,V_PISC,SharePx,{_pp(n)})"
    F_ = f"(({tst})>=V_TrEff)*(1-V_TrPlan)"
    return (f"=IF({qc(n)}$5<=MAX(LastHist,V_Q),0,IF({qc(n)}$5=FP_ATMIdx,FP_ATMProc/{atm_},0)+{X('FP_EqTop', n)}/{iss}"
            f"+SUMPRODUCT((V_TrMat={qc(n)}$5)*V_TrPM*(V_TrRT/1000*{F_}+V_TrAcc*(V_TrSpot-{F_})/{iss})))")
Sx.section("v0.7d: funding-plan & convert-maturity shares at the HEADLINE issue-price basis (PTBasis 1 self-consistent / 2 spot / 3 path) - added to Sh_Basic")
_psf = Sx.row("Plan shares issued in quarter: ATM $ / price + equity top-up $ / price + maturity conversions (tested at that price) + re-funding of notes that do not convert", "M sh", "Modeling", "after the valuation quarter only (earlier issuance is in the pre-plan base)", fmt="#,##0.00", fn=_plan_sh)
reg("Shares", "Sh_PlanFlow", _psf)
_me = Sx.r
_psc = Sx.row("Cumulative plan shares (added to Sh_Basic)", "M sh", "Modeling", "", kind="stock", fmt="#,##0.00", e="=0", fn=lambda n, me=_me: f"={qc(n-1)}{me}+{qc(n)}{_psf}")
assert _psc == _me
reg("Shares", "Sh_PlanCum", _psc)
rr_ = K.r
cfw[f"A{rr_}"] = "Starting liquidity incl. post-6/30 raises (cash + Aug gross + loan + ATM toggle)"; fml(cfw[f"E{rr_}"], "=CS_Cash+AugGross+SecAmt*LoanOn+ATMOn*ATM_Rem*SharePx*(FP_On=0)", "#,##0", True); name(wb, "Liq_Start", "CashFlow", f"$E${rr_}")
cfw[f"A{rr_+1}"] = "First quarter liquidity < 0"; fml(cfw[f"E{rr_+1}"], f'=IFERROR(INDEX(Q_Lab,1,MATCH(TRUE,INDEX($F${liq}:$AG${liq}<0,0),0)),"none to Q4\'30")', "@", True); name(wb, "Liq_NegQ", "CashFlow", f"$E${rr_+1}")
cfw[f"A{rr_+2}"] = "Memo: H2'26 new prepayments (model) vs FY26 guide remainder (>$9B - 1H26 DR build)"; fml(cfw[f"E{rr_+2}"], "=SUMPRODUCT((Q_Yr=2026)*(Q_Idx>LastHist)*CF_Prepay)", "#,##0"); fml(cfw[f"F{rr_+2}"], "=Prepay26-(DR_Open-DR_YE25)", "#,##0")
cfw[f"A{rr_+4}"] = "Comparison vs Skeptic rerun (same basis), $B"; cfw[f"A{rr_+4}"].font = BOLD
for j, t in enumerate(["", "", "", "", "Start", "Q1'27 end", "Q2'27 end", "YE27"]):
    if t: cfw.cell(rr_ + 4, 1 + j, t).font = BOLD
cfw[f"A{rr_+5}"] = "Model"; cfw[f"A{rr_+6}"] = "Skeptic rerun"; cfw[f"A{rr_+7}"] = "Model - Skeptic"
for j, (mf, sk) in enumerate([("Liq_Start", "Sk_Start"), (f"INDEX(Liq,1,13)", "Sk_Q127"), (f"INDEX(Liq,1,14)", "Sk_Q227"), (f"INDEX(Liq,1,16)", "Sk_YE27")]):
    c = CL(5 + j)
    fml(cfw[f"{c}{rr_+5}"], f"={mf}/1000", "0.0"); fml(cfw[f"{c}{rr_+6}"], f"={sk}", "0.0"); fml(cfw[f"{c}{rr_+7}"], f"={c}{rr_+5}-{c}{rr_+6}", "0.0;(0.0)")
name(wb, "Liq_Cmp", "CashFlow", f"$E${rr_+5}:$H${rr_+6}")
K.r = rr_ + 9

# ---- KPIs
kp = TS(wb, "KPIs", "Modeling-owned derived KPIs (quarterly, annualized where $/MW). All formulas from Revenue / Sites / Contracts / Capex_DA / PnL.")
KP = kp
KP.section("MW bases (IT MW)")
bill = KP.row("BILLING IT MW = active - delivered-but-power-capped contract MW", "MW", "Modeling", "", kind="stock", fmt="#,##0",
              fn=lambda n: f"=MAX(0,{X('MW_Act', n)}-({X('A_MW', n)}-{X('A_BillMW', n)}))")
k_ct = KP.row("Contracted IT MW", "MW", "Sites", "", kind="stock", fmt="#,##0", fn=lambda n: f"={X('MW_Contr', n)}")
k_en = KP.row("Energized IT MW (Capacity estimate, live scenario)", "MW", "Sites", "", kind="stock", fmt="#,##0", fn=lambda n: f"={X('MW_En', n)}")
k_ow = KP.row("Owned IT MW (active)", "MW", "Sites", "", kind="stock", fmt="#,##0", fn=lambda n: f"={X('MW_Act', n)}*{X('MW_OwnSh', n)}")
k_ls = KP.row("Leased IT MW (active; colo + build-to-suit)", "MW", "Sites", "", kind="stock", fmt="#,##0", fn=lambda n: f"={X('MW_Act', n)}*(1-{X('MW_OwnSh', n)})")
KP.section("Revenue per MW ($M per MW-yr, annualized AI cloud revenue)")
F2 = "#,##0.00;(#,##0.00);-"
def per(num, den, mult="4"):
    return lambda n: f"=IFERROR({num(n)}*{mult}/{den(n)},0)"
KP.row("Revenue per BILLING MW", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, bold=True, fn=per(lambda n: X('AIRev', n), lambda n: f"{qc(n)}{bill}"))
KP.row("Revenue per contracted MW", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=per(lambda n: X('AIRev', n), lambda n: f"{qc(n)}{k_ct}"))
KP.row("Revenue per energized MW", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=per(lambda n: X('AIRev', n), lambda n: f"{qc(n)}{k_en}"))
KP.row("Microsoft revenue per Microsoft billing MW", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=per(lambda n: X('MS_Rev', n), lambda n: X('MS_BillMW', n)))
KP.section("Fleet KPIs")
KP.row("Revenue per GPU ($K per GPU-yr)", "$K/GPU-yr", "Modeling", "", kind="none", fmt="#,##0", fn=per(lambda n: X('AIRev', n), lambda n: X('GPU_Avg', n), "4000"))
bl_ = KP.row("Implied blended $/GPU-hr", "$/GPU-hr", "Modeling", "AI revenue / (avg GPUs x hours)", kind="none", fmt="0.00",
             fn=lambda n: f"=IFERROR({X('AIRev', n)}*1000000/({X('GPU_Avg', n)}*{HRS(n)}),0)")
lw = KP.row("Fleet-weighted on-demand list price", "$/GPU-hr", "Pricing", "", kind="none", fmt="0.00",
            fn=lambda n: f"=IFERROR(({X('GPU_End_H', n)}*{X('P_OD_H', n)}+{X('GPU_End_B', n)}*{X('P_OD_B', n)}+{X('GPU_End_N', n)}*{X('P_OD_N', n)})/{X('GPU_End', n)},0)")
KP.row("Implied utilization (blended $/hr / weighted list)", "%", "Modeling", "Mixes realized-price discount with utilization", kind="none", fmt="0%",
       fn=lambda n: f"=IFERROR({qc(n)}{bl_}/{qc(n)}{lw},0)")
KP.row("Capex per new active MW ($M/MW)", "$M/MW", "Modeling", "", kind="none", fmt=F2, fn=lambda n: f"=IFERROR({X('Capex', n)}/{X('MW_New', n)},0)")
KP.row("Adj. EBITDA per billing MW ($M/MW-yr)", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, bold=True, fn=per(lambda n: X('PL_EBITDA', n), lambda n: f"{qc(n)}{bill}"))
KP.row("Adj. EBITDA per energized MW ($M/MW-yr)", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=per(lambda n: X('PL_EBITDA', n), lambda n: f"{qc(n)}{k_en}"))
KP.section("Owned vs leased cohorts ($M per MW-yr)")
pre = KP.row("Pre-rent EBITDA per active MW (both cohorts)", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2,
             fn=per(lambda n: f"({X('PL_EBITDA', n)}+{X('PL_Rent', n)})", lambda n: X('MW_Act', n)))
rpl = KP.row("Rent per leased MW", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=per(lambda n: X('PL_Rent', n), lambda n: f"{qc(n)}{k_ls}"))
KP.row("EBITDA per MW - OWNED cohort (no rent; shell/power in capex)", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=lambda n: f"={qc(n)}{pre}")
KP.row("EBITDA per MW - LEASED cohort (after rent)", "$M/MW-yr", "Modeling", "", kind="none", fmt=F2, fn=lambda n: f"={qc(n)}{pre}-{qc(n)}{rpl}")
KP.row("Capex per MW - owned (IT + ShellMW + PowerMW)", "$M/MW", "Inputs", "", kind="none", fmt=F2, fn=lambda n: "=ITCapexMW+ShellMW+PowerMW")
KP.row("Capex per MW - leased (IT only; colo case ~$30M/IT MW)", "$M/MW", "Inputs", "", kind="none", fmt=F2, fn=lambda n: "=ITCapexMW")
KP.row("Depreciation yield ex life change (Capex_DA)", "%", "Capex_DA", "", kind="none", fmt="0.0%", fn=lambda n: f"={X('DepYld', n)}")

# ---- Outputs
ou = wb.create_sheet("Outputs")
ou["A1"] = "Outputs - key metrics at valuation-date horizons"; ou["A1"].font = TITLE
ou["A2"] = "Horizon quarter = last quarter end on/before EDATE(ValDate, months). LTM = 4 quarters ending at horizon. Per-share value rolled forward at Ke."; ou["A2"].font = ITAL
setw(ou, {"A": 58, "B": 14, "C": 14, "D": 14, "E": 14, "F": 4, "G": 14, "H": 14, "I": 14, "J": 50})
hd = ["Metric", "Valuation qtr", "+6 months", "+1 year", "+3 years"]
for i, t in enumerate(hd):
    c = ou.cell(4, 1 + i, t); c.font = HDR; c.fill = HDRFILL
MO = [0, 6, 12, 36]
for i, m in enumerate(MO):
    cl = CL(2 + i)
    fml(ou[f"{cl}5"], f"=MATCH(EDATE(ValDate,{m}),Q_End,1)", "0"); fml(ou[f"{cl}6"], f"=INDEX(Q_Lab,1,{cl}5)", "@", True)
    ou[f"{cl}7"] = m
ou["A5"] = "Quarter index"; ou["A6"] = "Quarter"; ou["A7"] = "Months from ValDate"
rows = [("Revenue - quarter ($M)", "INDEX(PL_Rev,1,{i})", "#,##0"),
        ("Revenue - LTM ($M)", "SUMPRODUCT(PL_Rev*(Q_Idx>{i}-4)*(Q_Idx<={i}))", "#,##0"),
        ("Adj. EBITDA - quarter ($M)", "INDEX(PL_EBITDA,1,{i})", "#,##0;(#,##0)"),
        ("Adj. EBITDA - LTM ($M)", "SUMPRODUCT(PL_EBITDA*(Q_Idx>{i}-4)*(Q_Idx<={i}))", "#,##0;(#,##0)"),
        ("EBIT - LTM ($M)", "SUMPRODUCT(PL_EBIT*(Q_Idx>{i}-4)*(Q_Idx<={i}))", "#,##0;(#,##0)"),
        ("Free cash flow - LTM ($M)", "SUMPRODUCT(CF_FCF*(Q_Idx>{i}-4)*(Q_Idx<={i}))", "#,##0;(#,##0)"),
        ("Net debt ($M, model)", "INDEX(CF_NetDebt,1,{i})", "#,##0;(#,##0)"),
        ("Diluted shares (M, at DilPx)", "INDEX(Sh_Dil,1,{i})", "#,##0.0"),
        ("Active IT MW", "INDEX(MW_Act,1,{i})", "#,##0"),
        ("Energized IT MW (Capacity)", "INDEX(MW_En,1,{i})", "#,##0"),
        ("Implied value per share - FD, rolled at Ke ($)", "V_PS*(1+Ke)^({m}/12)", "0.00"),
        ("Implied value per share - 308.69M-style basic ($)", "V_PSB*(1+Ke)^({m}/12)", "0.00")]
for k, (lab, f, fmt) in enumerate(rows):
    rr = 8 + k
    ou[f"A{rr}"] = lab
    for i in range(4):
        cl = CL(2 + i)
        fml(ou[f"{cl}{rr}"], "=" + f.format(i=f"{cl}$5", m=f"{cl}$7"), fmt)
rr = 8 + len(rows) + 1
ou[f"A{rr}"] = "Valuation summary"; ou[f"A{rr}"].font = BOLD
for k, (lab, f, fmt) in enumerate([("DCF EV ($M)", "=V_EV", "#,##0"), ("Market EV at SharePx ($M)", "=V_MKTEV", "#,##0"),
                                   ("FD value / share at solved price ($)", "=V_PS", "0.00"), ("12-month PT ($) - SELF-CONSISTENT headline", "=V_PT", "0.00"), ("MEMO: PT if issuing at spot $242.81 (needs market access at today's price)", "=V_PTSpot", "0.00"), ("Shares needed vs ~82M authority (M)", "=V_AuthNeed", "#,##0.0"),
                                   ("Upside vs SharePx", "=V_UP", "0%"), ("Share price ($)", "=SharePx", "0.00"),
                                   ("Off-B/S weight used", "=OffBSWt", "0%"), ("Scenario", '=CHOOSE(Scen,"BULL","BASE","BEAR")', "@")]):
    ou[f"A{rr+1+k}"] = lab; fml(ou[f"B{rr+1+k}"], f, fmt, True)
rr = rr + 11
ou[f"A{rr}"] = "Model vs consensus (Skeptic, Oct 2 2026)"; ou[f"A{rr}"].font = BOLD
for i, t in enumerate(["", "FY2026", "FY2027", "FY2028"]):
    if t: ou.cell(rr, 1 + i, t).font = BOLD
for k, (lab, mrow, cnm) in enumerate([("Revenue", "PL_Rev", "Cons_Rev"), ("Adj. EBITDA", "PL_EBITDA", "Cons_EB"), ("Capex", "Capex", "Cons_Cx")]):
    base = rr + 1 + 3 * k
    ou[f"A{base}"] = f"{lab} - model"; ou[f"A{base+1}"] = f"{lab} - consensus"; ou[f"A{base+2}"] = f"{lab} - model / consensus - 1"
    for i, y in enumerate([2026, 2027, 2028]):
        cl = CL(2 + i)
        fml(ou[f"{cl}{base}"], f"=SUMIF(Q_Yr,{y},{mrow})", "#,##0")
        fml(ou[f"{cl}{base+1}"], f"=INDEX({cnm},1,{y-2023})", "#,##0")
        fml(ou[f"{cl}{base+2}"], f'=IF(N({cl}{base+1})=0,"n/a",{cl}{base}/{cl}{base+1}-1)', "0%")
rr = rr + 11
ou[f"A{rr}"] = "MW sources at YE26 / YE27 (IT MW)"; ou[f"A{rr}"].font = BOLD
for i, t in enumerate(["", "Q4'26", "Q4'27"]):
    if t: ou.cell(rr, 1 + i, t).font = BOLD
for k, (lab, nm_) in enumerate([("Energized - Capacity, IT-normalized (YE27 ~950 IT = PLACEHOLDER pending Capacity's Tue basis column; named 185/305/424 & 764/1,069/1,374 mixed basis + unseen)", "MW_En"),
                                ("Mgmt connected guide (800MW-1GW YE26; 1.8 GW YE27)", "MW_Gd"), ("Contracted", "MW_Contr"), ("Active MW driving revenue", "MW_Act")]):
    ou[f"A{rr+1+k}"] = lab
    fml(ou[f"B{rr+1+k}"], f"=INDEX({nm_},1,12)", "#,##0"); fml(ou[f"C{rr+1+k}"], f"=INDEX({nm_},1,16)", "#,##0")
rr = rr + 7
ou[f"A{rr}"] = "Comparison vs Skeptic v2 (model columns for low/high/stress are STATIC snapshots written by snap.py)"; ou[f"A{rr}"].font = BOLD
for i, t in enumerate(["", "Model", "Skeptic v2", "Model / Sk - 1"]):
    if t: ou.cell(rr + 1, 1 + i, t).font = BOLD
CMP = [("FY27 revenue - low (model: bear)", None, "Sk_Rev27Lo"), ("FY27 revenue - base (1-qtr lag)", "=SUMIF(Q_Yr,2027,PL_Rev)", "Sk_Rev27Base"),
       ("FY27 revenue - high (model: bull)", None, "Sk_Rev27Hi"), ("FY27 revenue - STRESS (2-qtr energized lag)", None, "Sk_Rev27St"),
       ("FY28 revenue - base", "=SUMIF(Q_Yr,2028,PL_Rev)", "Sk_Rev28"), ("Market EV / FY27 adj. EBITDA - Nebius", "=V_EVEB27", "Sk_EVEB27"),
       ("EV / FY27 EBITDA - CoreWeave (Skeptic)", "=CRWV_EVEB27", "CRWV_EVEB27")]
for k, (lab, mf, sk) in enumerate(CMP):
    r_ = rr + 2 + k
    ou[f"A{r_}"] = lab
    fmt_ = "0.0x" if "EBITDA" in lab else "#,##0"
    if mf: fml(ou[f"B{r_}"], mf, fmt_)
    else: ou[f"B{r_}"].number_format = fmt_
    fml(ou[f"C{r_}"], f"={sk}", fmt_); fml(ou[f"D{r_}"], f'=IFERROR(B{r_}/C{r_}-1,"")', "0%")
OUT_CMP0 = rr + 2
rr = rr + 2 + len(CMP) + 1
ou[f"A{rr}"] = "Funding runway (CashFlow tab)"; ou[f"A{rr}"].font = BOLD
ou[f"A{rr+1}"] = "Starting liquidity ($M)"; fml(ou[f"B{rr+1}"], "=Liq_Start", "#,##0")
ou[f"A{rr+2}"] = "First quarter liquidity < 0 (Skeptic rerun: Q2'27)"; fml(ou[f"B{rr+2}"], "=Liq_NegQ", "@")
ou[f"A{rr+3}"] = "Liquidity Q1'27 / Q2'27 / YE27 ($B)"; fml(ou[f"B{rr+3}"], "=INDEX(Liq,1,13)/1000", "0.0"); fml(ou[f"C{rr+3}"], "=INDEX(Liq,1,14)/1000", "0.0"); fml(ou[f"D{rr+3}"], "=INDEX(Liq,1,16)/1000", "0.0")
OUT_SNAP = rr + 5

# ---- Sensitivity
se = wb.create_sheet("Sensitivity")
se["A1"] = "Sensitivities"; se["A1"].font = TITLE
setw(se, {"A": 60, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 50})
se["A3"] = "1) List-price change from quarter after valuation: on-demand book (c) immediately; reserved book (b) only after ResLag renewal; hyperscaler book (a) unaffected. EBITDA = 100% flow-through (no revenue-linked cost)."; se["A3"].font = ITAL
for i, t in enumerate(["Revenue impact ($M)", "-20%", "-10%", "+10%", "+20%", "% of model rev (+10%)"]):
    c = se.cell(4, 1 + i, t); c.font = HDR; c.fill = HDRFILL
for k, y in enumerate([2026, 2027, 2028]):
    rr = 5 + k
    se[f"A{rr}"] = f"FY{y} revenue = EBITDA impact"
    base = f"SUMPRODUCT((Q_Yr={y})*(Q_Idx>V_Q)*(Book_C+Book_B*(Q_Idx>V_Q+ResLag)))"
    for i, d in enumerate([-0.2, -0.1, 0.1, 0.2]):
        fml(se[f"{CL(2+i)}{rr}"], f"={d}*{base}", "#,##0;(#,##0)")
    fml(se[f"F{rr}"], f"=IFERROR(D{rr}/SUMIF(Q_Yr,{y},PL_Rev),0)", "0.0%")
se["A9"] = "  Exposure mix at valuation quarter: book (a) / (b) / (c) share of AI cloud revenue"
fml(se["B9"], "=INDEX(Book_A,1,V_Q)/INDEX(AIRev,1,V_Q)", "0%"); fml(se["C9"], "=INDEX(Book_B,1,V_Q)/INDEX(AIRev,1,V_Q)", "0%"); fml(se["D9"], "=INDEX(Book_C,1,V_Q)/INDEX(AIRev,1,V_Q)", "0%")
SENS_SNAP = 11

# ---- Checks
ck = wb.create_sheet("Checks")
ck["A1"] = "Checks"; ck["A1"].font = TITLE
setw(ck, {"A": 80, "B": 16, "C": 16, "D": 12, "E": 60})
for i, t in enumerate(["Check", "Model", "Target / tolerance", "Status", "Note"]):
    c = ck.cell(3, 1 + i, t); c.font = HDR; c.fill = HDRFILL
CK = []
def CHK(label, model, target, status, note="", fmt="#,##0.00"):
    CK.append((label, model, target, status, note, fmt))
PXv = PX.replace("$H$", "Valuation!$H$", 1); B_v = "(" + _H(O_.replace("$E$", "Valuation!$E$", 1), B_.replace("$F$", "Valuation!$F$", 1)) + ")"
dbl = f"IF(DilMethod=1,SUMPRODUCT({B_v}),IF(DilMethod=2,SUMPRODUCT((V_PSolve>{PXv})*{B_v}),0))+V_CVDEBT-SUMPRODUCT({B_v})"
CHK("CRITICAL: convert counted in BOTH shares and debt (solved price, chosen weight)", f"={dbl}", 0, f'=IF(ABS(B{{r}})<0.5,"OK","DOUBLE COUNT")',
    "If-converted + debt-counted must equal total (CVDebtBasis: 1 original principal, 2 accreted); net-share (3) keeps principal as debt by design")
CHK("CRITICAL: same at SharePx (market EV)", f"=IF(DilMethod=1,SUMPRODUCT({B_v}),IF(DilMethod=2,SUMPRODUCT((SharePx>{PXv})*{B_v}),0))+V_CVDmkt-SUMPRODUCT({B_v})", 0, f'=IF(ABS(B{{r}})<0.5,"OK","DOUBLE COUNT")')
CHK("Economic basic shares at Q3'26 (M) = 271.86 + 15.77 + 21.07", "=INDEX(Sh_EconBasic,1,11)", 308.69, '=IF(ABS(B{r}-C{r})<0.05,"OK","CHECK")')
CHK("Fully diluted at $242.81 EXCL. unsold ATM vs Filings CapStructure", "=V_FDmkt", "=FD_Filings", '=IF(ABS(B{r}-C{r})<2,"OK","CHECK")', "nebius_filings.xlsx CapStructure")
CHK("Fully diluted at $242.81 INCL. 12.27M ATM vs Filings", "=V_FDmkt+ATM_Rem", "=FD_FilingsATM", '=IF(ABS(B{r}-C{r})<2,"OK","CHECK")', "ATM is a funding source, not in today's count")
CHK("ITM convert shares at $242.81 (six ITM)", "=V_ITMsh", "=ITM_Sh_Chk", '=IF(ABS(B{r}-C{r})<0.5,"OK","CHECK")', "Steps at $324.65 / $313.46 / $183.22 / $180.31 / $138.75 / $51.45")
CHK("ITM convert principal at $242.81 ($M orig.)", "=V_ITMprin", "=ITM_Prin_Chk", '=IF(ABS(B{r}-C{r})<15,"OK","CHECK")')
CHK("Market EV at $242.81 ($M), CapStructure basis vs Filings ~74.7B", "=V_MKTEV", 74700, '=IF(ABS(B{r}/C{r}-1)<0.01,"OK","CHECK")', "6/30 balance + Aug notes + exchange; unrestricted cash only")
CHK("Dilution solve convergence |P15-P14|", "=V_CONV", 0.01, '=IF(B{r}<C{r},"OK","NOT CONVERGED")', "", "0.0000")
CHK("Microsoft implied TCV vs reported ($M)", "=MSFT_Chk", 0, '=IF(ABS(B{r})<5,"OK","CHECK")')
CHK("Quarters (pre-Bloom) Microsoft revenue exceeds ~$350M/qtr genset ceiling", "=SUMPRODUCT((MS_CeilChk>0)*(Q_Idx<IF(BloomOn=1,BloomIdx,99)))", 0, '=IF(B{r}=0,"OK","CHECK")', "", "0")
CHK("Q4 bridge: unseen MW vs Skeptic/Capacity 83-200 MW", "=BR_MW25", "=UnseenChkHi", "=BR_Chk", "", "#,##0")
CHK("History capex split = reported (sum abs diff)", "=SUMPRODUCT(ABS(N(+CapexChk)))", 0, '=IF(B{r}<1,"OK","CHECK")')
CHK("Q2'26 book (a) share vs GPU Market 45% (40-46%)", "=INDEX(Book_A,1,10)/INDEX(AI_Model,1,10)", "=MixA_Q2", '=IF(AND(ROUND(B{r},2)>=0.4,ROUND(B{r},2)<=0.46),"OK","CALIBRATE")', "", "0%")
CHK("Q2'26 book (b) share vs ~50% (44-57%)", "=INDEX(Book_B,1,10)/INDEX(AI_Model,1,10)", "=MixB_Q2", '=IF(AND(B{r}>=0.44,B{r}<=0.57),"OK","CALIBRATE")', "", "0%")
CHK("Q2'26 book (c) share vs ~4% (1.3-9%)", "=INDEX(Book_C,1,10)/INDEX(AI_Model,1,10)", "=MixC_Q2", '=IF(AND(B{r}>=0.013,B{r}<=0.09),"OK","CALIBRATE")', "", "0%")
CHK("Depreciation-yield flags (drop after chip availability)", "=SUM(DepYldFlag)", 0, '=IF(B{r}=0,"OK","FLAG - review PIS dates")', "Material weakness over depreciation start dates", "0")
CHK("Q3'26 transfer into service", "=Q3TransChk", "", '=B{r}', "", "@")
CHK("Sites: named-site energized imports tie to Capacity file COMPANY row (sum abs diff)", "=SUMPRODUCT(ABS(CapTie))", 0, '=IF(B{r}<1,"OK","CHECK")', "", "#,##0.0")
CHK("Q2'26 tie: model AI cloud revenue vs actual $574.9M", "=INDEX(AI_Model,1,10)", "=INDEX(H_AIRev,1,10)", '=IF(ABS(B{r}/C{r}-1)<0.01,"OK","CHECK")', "Utilization calibrated", "#,##0.0")
CHK("Implied Q2'26 utilization (books b+c)", "=UtilQ2", "40-100%", "=UtilFlag", "Flag (not forced) if outside 40-100%", "0.0%")
CHK("(a) Q2'26 revenue per GPU-hour of capacity (annualized AI rev / GPUs / 8,760)", "=INDEX(AI_Model,1,10)*4*1000000/INDEX(GPU_End,1,10)/8760", "=Chk_RevGPUhr", '=IF(ABS(B{r}/C{r}-1)<=0.1,"OK","CHECK")', "+/-10% (GPU Market, Derived)", "0.00")
CHK("(b) Q2'26 annualized AI rev per $ avg gross servers in service", "=INDEX(AI_Model,1,10)*4/AVERAGE(INDEX(H_SRV,1,9),INDEX(H_SRV,1,10))", "=Chk_RevSrv", '=IF(ABS(B{r}/C{r}-1)<=0.1,"OK","CHECK")', "Filings, Derived", "0.00")
CHK("(b3) Capacity server ceiling: Q2'26 GPUs x $45-60k (high end) vs $5.2B servers in service", "=INDEX(GPU_End,1,10)*Chk_SrvHi/1000000", "=Chk_SrvCap", '=IF(B{r}<=C{r},"OK: "&TEXT(INDEX(GPU_End,1,10)*Chk_SrvLo/1000000,"#,##0")&"-"&TEXT(B{r},"#,##0")&" <= 5,198; implied $"&TEXT(C{r}*1000000/INDEX(GPU_End,1,10),"#,##0")&"/GPU","FLAG")', "Capacity, Estimated", "#,##0")
CHK("(b2) Modeled in-service GPU cost at 6/30 (GPUs x $/GPU) vs $5.2B servers", "=(INDEX(GPU_End_H,1,10)*GPUCost_H+INDEX(GPU_End_B,1,10)*GPUCost_B+INDEX(GPU_End_N,1,10)*GPUCost_N)/1000000", "=Chk_SrvCap", '=IF(B{r}<=C{r},"OK","FLAG: exceeds servers in service")', "", "#,##0")
CHK("(c) Q2'26 revenue per energized MW ($M/yr)", "=INDEX(AI_Model,1,10)*4/INDEX(MW_En,1,10)", "=Chk_RevEnMW", '=IF(ABS(B{r}/C{r}-1)<=0.15,"OK","CHECK")', "~$13M (2.30B / ~173 MW)", "0.0")
CHK("Q2'26 contract (book a) IT MW as share of energized MW", "=INDEX(A_MW,1,10)/INDEX(MW_En,1,10)", "info", '="INFO"', "Microsoft GPU use capped at Vineland energized IT MW (39 MW Q2'26) + Meta #1 48 MW (TCV/ACV-derived) vs 173 MW energized; v0.1 fix took Q2'26 utilization from 133% to ~92%", "0%")
CHK("Q2'26 energized MW (Capacity base, IT)", "=INDEX(MW_En,1,10)", 173, '=IF(ABS(B{r}-C{r})<3,"OK","CHECK")', "", "#,##0.0")
CHK("Q2'26 GPUs: model (Capacity 'GPUs deployed' ~80,416 + GPU anchor gap if GPUAnchorOn)", "=INDEX(GPU_End,1,10)", "=80416+GPUAnchorD", '=IF(ABS(B{r}/C{r}-1)<0.01,IF(GPUAnchorOn=1,"INFO: anchored to ~30k at 3/31/25 (+"&TEXT(GPUAnchorD,"#,##0")&" vs Capacity file); Q2 utilization "&TEXT(UtilQ2,"0.0%"),"OK"),"CHECK")', "", "#,##0")
CHK("Unexplained ARR residual at 6/30 (June run-rate basis; NOT unseen capacity)", "=ARR_Res", "=ARR_ResTeamLo&\"-\"&ARR_ResTeamHi", '="INFO: model "&TEXT(ARR_Res,"#,##0")&" = "&TEXT(ARR_ResMW_TLo,"0")&"-"&TEXT(ARR_ResMW_THi,"0")&" MW-eq at team $13.3-13.5M/MW (model June RR "&TEXT(ARR_RR,"#,##0")&" = $"&TEXT(ARR_RRMW,"0.0")&"M/MW); team-basis residual "&TEXT(ARR_ResTB_Lo,"#,##0")&"-"&TEXT(ARR_ResTB_Hi,"#,##0")', "Pending Filings on ARR scope / signed-not-billing", "#,##0")
CHK("Adj. EBITDA margin FY28 - model vs consensus (flag if > +5pp)", "=SUMIF(Q_Yr,2028,PL_EBITDA)/SUMIF(Q_Yr,2028,PL_Rev)", "=MarginCons28", '=IF(B{r}-C{r}>0.05,"FLAG: model margin > consensus +5pp","OK")', "Skeptic consensus 13.98/23.0", "0.0%")
CHK("Adj. EBITDA margin FY30 - model vs peer ~62% (flag if > +8pp)", "=SUMIF(Q_Yr,2030,PL_EBITDA)/SUMIF(Q_Yr,2030,PL_Rev)", "=MarginPeer", '=IF(B{r}-C{r}>0.08,"FLAG: above peer +8pp - review opex/MW and revenue/MW","OK")', "CoreWeave ~62% placeholder", "0.0%")
CHK("FY30 margin sensitivity: PD + SG&A held at FY28 % of revenue (vs placeholder step-down)", "=SUMIF(Q_Yr,2030,PL_EBITDA)/SUMIF(Q_Yr,2030,PL_Rev)-(1-OpexHold)*((INDEX(PDpct,1,2028-FYBase)+INDEX(SGApct,1,2028-FYBase))-(INDEX(PDpct,1,2030-FYBase)+INDEX(SGApct,1,2030-FYBase)))", "=MarginPeer", '=IF(B{r}-C{r}>0.08,"FLAG: still above peer +8pp","INFO: within +8pp of peer")', "PD 8% + SG&A 10% (FY28) instead of 6% + 7% (FY30); toggle OpexHold=1 for full model effect (Sensitivity)", "0.0%")
CHK("FY30 opex per avg IT MW ($M/MW-yr: COR + PD + SG&A ex-SBC)", "=(SUMIF(Q_Yr,2030,PL_Rev)-SUMIF(Q_Yr,2030,PL_EBITDA))/(SUMIF(Q_Yr,2030,MW_Avg)/4)", "info", '="INFO: revenue/MW "&TEXT(SUMIF(Q_Yr,2030,PL_Rev)/(SUMIF(Q_Yr,2030,MW_Avg)/4),"0.0")', "", "0.00")
CHK("FY26 capex vs company guide $20-25B (Aug 12 2026 call)", "=SUMIF(Q_Yr,2026,Capex)", "20,000-25,000", '=IF(AND(B{r}>=20000*0.9,B{r}<=25000*1.1),"OK (within guide +/-10%)","FLAG: outside guide")', "Fact base: FY2026 capex guidance $20-25B (call transcript, CFO); GPULead calibration", "#,##0")
CHK("Bull DCF EV vs 3x base (STATIC, snap.py)", "=V_EVBullS", "=3*V_EVBaseS", '=IF(V_EVBaseS=0,"run snap.py",IF(B{r}>C{r},"FLAG: bull EV > 3x base - review bull drivers","OK"))', "Bull/base EVs from snapshot", "#,##0")
CHK("Q2'26 calibration scenario-invariant: utilization bull / bear minus base (STATIC)", "=MAX(ABS(V_UtBullS-V_UtBaseS),ABS(V_UtBearS-V_UtBaseS))", 0, '=IF(V_UtBaseS=0,"run snap.py",IF(B{r}<0.0005,"OK: all scenarios tie $574.9M at "&TEXT(V_UtBaseS,"0.0%"),"FLAG: history depends on scenario"))', "", "0.0000")
CHK("Implied blended realization vs OD list, Q2'26 books (b)+(c) (billed GPU-hours at Q2 avg list) vs GPU Market 0.72-0.80", "=(INDEX(Book_B,1,LastHist)+INDEX(Book_C,1,LastHist))*1000000/((INDEX(NCA_H,1,LastHist)*ODq_H+INDEX(NCA_B,1,LastHist)*ODq_B+INDEX(NCA_N,1,LastHist)*ODq_N)*INDEX(Q_Hrs,1,LastHist)*UtilQ2)", "0.72-0.80", '=IF(AND(B{r}>=0.72,B{r}<=0.80),"OK: within GPU Market range",IF(B{r}<0.72,"INFO: below range","FLAG: above range"))', "Reserved RealRes_B 0.72 x ResSh + on-demand/spot mix", "0.000")
CHK("New-deal pricing: model new-MW gross revenue per IT MW-yr (NewDealOn: blended multi-year + short-term at Q1'27 vintage; off: GB300 at Oct-1 list x RealRes x Util_Res) vs IREN >$20M (3-yr GB300-class ~$4.65/GPU-hr)", "=IF(NewDealOn=1,RM_ND,RM_GB300)", 20, '=IF(B{r}>C{r}*1.25,"FLAG: model new-MW pricing >1.25x IREN benchmark",IF(B{r}<C{r}*0.8,"INFO: below benchmark","OK: multi-year "&TEXT(RM_MY,"0.0")&"; blended net leased "&TEXT(RM_ND_L,"0.0")&" vs 17.8; GPU-hour GB300 engine "&TEXT(RM_GB300,"0.0")&" (legacy only)"))', "Short-term tranche (STShare) at $40-50M included in blend; IREN excludes short-term", "0.0")
CHK("Vineland shutdown: probability-weighted Q4'26 Microsoft revenue at risk at VinRevMW vs Capacity ~$67M", "=VinExpQ4", 67, '=IF(ABS(B{r}-C{r})<=3,"OK: ties ~$67M (Capacity, Derived); VinShutSel="&VinShutSel,"FLAG: differs from Capacity tie-out")', "vineland_shutdown.xlsx weights; $14.5M/IT MW-yr", "0.0")
CHK("Backlog floor: model AI cloud revenue Jul'26-Jun'28 minus RPO recognised in window (~$13.5B) = implied revenue from deals signed after Q2'26", "=SUMPRODUCT((Q_Idx>LastHist)*(Q_Idx<=LastHist+8)*AI_Model)-RPO_Tot*RPO24Pct", "=RPO_Tot*RPO24Pct", '=IF(B{r}<0,"FLAG: model below contracted backlog","INFO: "&TEXT(B{r}/(B{r}+C{r}),"0%")&" of window revenue unsold at 6/30 (~"&TEXT(B{r}/(B{r}+C{r})*SUMPRODUCT((Q_Idx>LastHist)*(Q_Idx<=LastHist+8)*MW_Avg)/8,"0")&" of avg "&TEXT(SUMPRODUCT((Q_Idx>LastHist)*(Q_Idx<=LastHist+8)*MW_Avg)/8,"0")&" active MW, revenue-weighted); sold share of new 2027+ MW "&TEXT(RPO_Sold,"0%"))', "Q2'26 Note 2: RPO $37.49B, 36% Jul'26-Jun'28", "#,##0")
CHK("Backlog cash floor Jul'26-Jun'28 = RPO in window - 6/30 DR recognised within 24 months (DR_24 base $2.25B; range $2.0-2.5B)", "=RPO_Tot*RPO24Pct-DR_24", "11,000-11,500", '="INFO: range "&TEXT(RPO_Tot*RPO24Pct-2500,"#,##0")&"-"&TEXT(RPO_Tot*RPO24Pct-2000,"#,##0")&"; model cash revenue in window (AI rev - DR release) "&TEXT(SUMPRODUCT((Q_Idx>LastHist)*(Q_Idx<=LastHist+8)*(AI_Model-CF_DRRel)),"#,##0")', "DR 6/30: current $979.4M, non-current $4,995.8M (Reported)", "#,##0")
CHK("Payback (months) vs mgmt 1y10m (22), prepayments excluded: GB300 leased IT / gross; Rubin leased base / bear (no toggle)", "=PB_GB_IT", "=PB_Target", '="INFO: GB300 "&TEXT(PB_GB_IT,"0")&" / "&TEXT(PB_GB_Gr,"0")&" mo (IT / gross); Rubin base "&TEXT(PB_RU_IT,"0")&" / bear "&TEXT(PB_RU_Gr,"0")&" mo; IRR GB300 IT "&TEXT(PB_GB_IRR,"0%")&", Rubin base "&TEXT(PB_RU_IRR,"0%")&"; IT fits 22 only at ~$32M capex, or if prepayments counted (implied rev "&TEXT(PB_ImplRevPre,"0.0")&")"', "Payback_Floor tab C; basis not disclosed by management", "0.0")
CHK("EBITDA per IT MW-yr by generation (IT reading): GB300 vs Rubin at RubinMW (direct); opex fixed", "=PB_EBGB", "=PB_EBRubin", '="INFO: GB300 "&TEXT(PB_EBGB,"0.0")&" -> Rubin "&TEXT(PB_EBRubin,"0.0")&" ("&TEXT(PB_EBRubin/PB_EBGB-1,"0%")&"); Rubin leased capex "&TEXT(CPM_Lease_VR,"0.0")&" (low-medium confidence)"', "v0.7: Rubin capex flat to GB300 in base", "0.0")
CHK("Asset floor: DCF EV vs PV(RPO margin) - capex for contracted MW", "=V_EV", "=PF_EV", '=IF(B{r}<C{r},"FLAG: DCF EV below asset floor","OK: DCF EV above floor")&"; SHARED-bridge floor $/sh "&TEXT(PF_PSHead,"0.00")&" vs DCF "&TEXT(V_PSHead,"0.00")&IF(AND(BridgeFix=1,V_EV>=PF_EV,PF_PSHead>V_PSHead+0.01)," FLAG: floor/sh above DCF/sh with EV above floor","")&IF(ABS(PF_RepChk)>0.01," FLAG: DCF replica mismatch","")&"; legacy equity floor $/sh "&TEXT(PF_PS1,"0.0")&" cons / "&TEXT(PF_PS2,"0.0")&" mark / "&TEXT(PF_PS3,"0.0")&" high"', "Payback_Floor tab D", "#,##0")
CHK("Revenue double-count: prepayments in revenue? (new-deal $20-25M/MW is revenue-recognition basis excl. prepayment, Q2 deck fn 2)", "=0", 0, '="OK: prepayments are cash timing only (CashFlow: Prepayments received / DR release); revenue engine = MW x $/MW and GPU-hours, no reference to prepayments or DR"', "Formula audit v0.5", "0")
CHK("Cash double-count: DR release reduces later cash (CFO = EBITDA + prepayments - release); closing Q4'30 DR run-off deducted from EV ($M)", "=V_DRAdj", "=-V_DR30", '="INFO: within-horizon release "&TEXT(SUM(CF_DRRel),"#,##0")&" deducted from cash; post-horizon run-off PV "&TEXT(-V_DRAdj,"#,##0")&" deducted (DRTermOn="&DRTermOn&"); closing DR "&TEXT(V_DR30,"#,##0")', "v0.5 fix: before, terminal UFCF counted the unreleased prepaid balance as cash again", "#,##0")
CHK("Prepayments: total new prepayments FY26H2-FY30 / new-build capex (signed Microsoft/Meta 0%, new hyperscaler HSPre, other 38.5% -> 22.5%)", "=SUM(CF_Prepay)/SUMPRODUCT((Q_Idx>LastHist)*CapNew)", "info", '="INFO: HSPre "&TEXT(HSPre,"0%")&"; other 2028+ "&TEXT(PrepayPct28,"0.0%")&"; decay "&PrepayDecayOn', "", "0.0%")
CHK("Significant financing component (ASC 606): SFC revenue gross-up in model revenue/EBITDA used for DCF", "=0", 0, '="OK: none (no SFC gross-up, no SFC interest; DCF on cash flows). Implied advance financing cost <= "&TEXT(V_SFCRate,"0.0%")&"/yr (Estimated upper bound)"', "Reported history may embed a small SFC gross-up (not disclosed separately)", "0")
CHK("GPU history anchor: model GPUs at 3/31/2025 vs ~30,000 actual (20-F)", "=INDEX(GPU_End,1,5)", "=GPU_Q125", '=IF(ABS(B{r}-C{r})<500,"OK: anchored (gap "&TEXT(GPUAnchorD,"#,##0")&" GPUs vs Capacity file)","FLAG: Capacity file "&TEXT(B{r},"#,##0")&" vs ~30,000 actual (GPUAnchorOn=0)")', "Single GPU calibration point; Kansas City likely live Q1'25 at 5-15 MW", "#,##0")
CHK("Rubin pricing (v0.7c): implied per-GPU-hr multiple to GB300 (derived from RubinMW) vs capex multiple", "=PX_Ratio", "=PX_CapRatio", '=PX_Flag&"; $/MW price multiple "&TEXT(PX_MWRatio,"0.00")&"x vs capex multiple "&TEXT(PX_CapRatio,"0.00")&"x"', "Rubin_Prem is now an output (no 1.2x input)", "0.00")
CHK("History MW basis (1 Capacity 173 / 2 company ~190 both history & forward / 3 mismatch test)", "=HistMWBasis", 1, '="INFO: Q2 energized basis "&TEXT(INDEX(MW_Act,1,LastHist),"0")&" MW; Q2 rev/MW $"&TEXT(INDEX(H_AIRev,1,LastHist)*4/INDEX(MW_Act,1,LastHist),"0.0")&"M"', "Same basis must drive history calibration and forward path", "0")
CHK("June step-up vs March: monthly $M and IT MW at take-or-pay pricing (upper end)", "=ARR_Step", "=ARR_StepMW", '="INFO: ~"&TEXT(ARR_StepMW,"0")&" MW at $12M/MW take-or-pay (fewer if reserved at Q2 renewal rates); JuneExitOn="&JuneExitOn', "", "#,##0")
CHK("DCF EV sign (negative = plan destroys value at WACC; equity floored at 0)", "=V_EV", 0, '=IF(B{r}<0,"FLAG: EV<0; FD value/share floored at 0","OK")', "", "#,##0")
CHK("Terminal value basis", "=V_TV", "=V_TVB", "=V_TVBasis", "Run-off floor when replacement capex exceeds terminal cash generation", "#,##0")
CHK("Funding runway: first quarter liquidity < 0 (after v0.7 funding plan when FP_On; no plug)", "=Liq_NegQ", "Skeptic: Q2'27", '="INFO"', "", "@")
CHK("Opening liquidity: Aug-26 converts counted ONCE (runway gross row + CashFlow net row, Q3'26 only); 6/30 cash $8,042M excludes them", "=SUM(Liq_Aug)", "=AugGross", '=IF(AND(ABS(B{r}-C{r})<0.5,ABS(SUMPRODUCT((Q_Idx<>11)*Liq_Aug))<0.5,ABS(INDEX(H_Cash,1,LastHist)-CS_Cash)<0.5,ABS(SUM(CF_CVProc)-SUM(CF_FPConv)-AugNet)<0.5),"OK: once (runway Q3\'26 gross 5,750; CashFlow Q3\'26 net 5,680)","DOUBLE/MISSING")', "Runway row '+ Aug-2026 converts gross' (CashFlow tab) and 'Convert issuance proceeds (net)' row", "#,##0")
CHK("$775M asset-backed facility drawn in base (LoanOn)", "=SecAmt*LoanOn", 0, '=IF(LoanOn=0,"OK: $0 drawn (undisclosed); $775M sensitivity","INFO: $775M drawn sensitivity")', "SOFR+2.50%; cash, debt, interest and runway all switch together", "#,##0")
CHK("Liquidity after funding plan: min every quarter Q3'26-Q4'30 vs buffer (FP_MinBuf if re-sized, else MinCash)", "=MIN(INDEX(Liq,1,11):INDEX(Liq,1,28))", "=IF(FP_Resize=1,FP_MinBuf,MinCash)", '=IF(FP_On=0,"INFO: plan off",IF(B{r}>=C{r}-0.5,"OK: every quarter >= buffer","FLAG: plan does not close the gap"))&"; YE27 "&TEXT(INDEX(Liq,1,16),"#,##0")&"; top-up equity "&TEXT(SUM(FP_EqTop),"#,##0")&" / "&TEXT(SUM(FP_EqTopSh),"0.0")&"M sh"', "Plan is ESTIMATED (not guidance)", "#,##0")
CHK("Convert waterfall at $147 (existing notes): converting principal / full-conversion shares vs ~$3.34B / ~26.2M", "=SUMPRODUCT((147>Valuation!" + PX + ")*Valuation!" + O_ + ")", 3337.5, '=IF(ABS(B{r}-C{r})<5,"OK","CHECK")&"; shares "&TEXT(' + convFull("147").replace(O_, "Valuation!" + O_).replace(RT, "Valuation!" + RT).replace(PX, "Valuation!" + PX) + ',"0.0")&"M; at $46 debt "&TEXT(SUMPRODUCT((46<=Valuation!' + PX + ')*Valuation!' + O_ + '),"#,##0")', "Parent reference: $46 all OTM ($13.4B debt); $147 26.2M sh; $243 all but Feb-30/Feb-34", "#,##0")
CHK("Post-funding FD shares at solved price (M) vs pre-funding", "=V_SHF", "=V_SH", '="INFO: +"&TEXT(V_SHF-V_SH,"0.0")&"M (ATM "&TEXT(V_FPSh,"0.0")&"M); full-conv "&TEXT(V_SHFFull,"0.0")&" / NSS "&TEXT(V_SHFNSS,"0.0")&"; value/share pre "&TEXT(V_PS,"0.0")&" -> post "&TEXT(V_PSF,"0.0")', "", "#,##0.0")
CHK("AUTHORITY (v0.7d): shares needed at the HEADLINE basis = new shares issued through FP_PostQ + 2027 plan-convert reserve vs ~130M (v0.8c) available (AuthAvail: ~64M 20% AGM authority + treasury)", "=V_AuthNeed", "=AuthAvail", '=IF(B{r}<=C{r},"OK: within authority","FLAG: needs "&TEXT(B{r}-C{r},"#,##0")&"M more (EGM approval)")&"; issued "&TEXT(V_NewShHead,"#,##0.0")&"M + convert reserve "&TEXT(V_CVResHead,"#,##0.0")&"M; spot basis "&TEXT(V_FPSh+SUMPRODUCT(V_TrPlanE*V_TrRT)/1000,"#,##0")&"M"', "Not a model cap", "#,##0.0")
CHK("Sh_Basic tie (v0.7d): plan shares in Shares tab (after V_Q through FP_PostQ) = valuation new shares at the headline basis", "=SUMPRODUCT((Q_Idx>V_Q)*(Q_Idx<=FP_PostQ)*Sh_PlanFlow)", "=V_NewShHead", '=IF(FP_On=0,"INFO: plan off",IF(ABS(B{r}-C{r})<0.5,"OK","CHECK: "&TEXT(B{r}-C{r},"0.0")))&IF(ABS(V_EqIssPre)>0.5," FLAG: top-up at/before V_Q","")', "", "#,##0.0")
CHK("Self-consistent PT fixed point converged (|S30 - S29|) and path memo (|T40 - T39|)", "=V_CONVSC", "=V_CONVPT", '=IF(AND(V_CONVSC<0.05,V_CONVPT<0.05),"OK","CHECK convergence")&"; SC PT "&TEXT(V_PTSC,"0.00")&" | spot memo "&TEXT(V_PTSpot,"0.00")&" | path memo "&TEXT(V_PTPath,"0.00")', "", "0.0000")
CHK("NVIDIA pre-funded warrant counted exactly once: economic basic = 271.86 issued + 15.77 exchange + 21.07 warrant (= 308.69); market EV equity uses the same 308.69", "=V_BASIC", "=271.855218+ExchShares+NVDA_W", '=IF(AND(ABS(V_BASIC-(271.855218+ExchShares+NVDA_W))<0.05,ABS(V_MKTEQ/SharePx-V_BASIC)<0.05),"OK: once in basic, market EV and FD","FLAG: warrant count mismatch")', "", "#,##0.00")
CHK("Financing does not change unlevered EV (CapIntFix): UFCF tax excludes capitalized-interest D&A", "=CapIntFix", "=1", '=IF(CapIntFix=1,"OK","FLAG: capitalized interest flows into UFCF tax")', "", "0")
CHK("Avride: value credited (V_AV) vs PV burn add-back (V_AVAdd) - credit value AND add back burn, or neither", "=V_AV", "=V_AVAdd", '=IF(V_AV=0,"OK: Avride $0, burn kept in EBITDA",IF(V_AVAdd>0,"OK: burn added back (no double count)","FLAG: double count"))', "", "#,##0")
CHK("Red-graded sites (count)", "=RedSites", "", '="INFO"', "", "0")
CHK("FY27 EBIT vs Skeptic at live life", "=SUMIF(Q_Yr,2027,PL_EBIT)", "=CHOOSE(MATCH(Life_B,{4,5,6},1),SkEBIT4,SkEBIT5,SkEBIT6)", '="INFO"', "See Sensitivity 4/5/6", "#,##0")
for k, (lab, model, target, status, note, fmt) in enumerate(CK):
    rr = 4 + k
    ck[f"A{rr}"] = lab; fml(ck[f"B{rr}"], model, fmt); ck[f"C{rr}"] = target; ck[f"C{rr}"].number_format = fmt
    ck[f"D{rr}"] = status.replace("{r}", str(rr)); ck[f"D{rr}"].font = BOLD; ck[f"E{rr}"] = note; ck[f"E{rr}"].font = ITAL
ck.conditional_formatting.add(f"D4:D{4+len(CK)}", FormulaRule(formula=['ISNUMBER(SEARCH("OK",D4))'], fill=PatternFill("solid", fgColor="C6EFCE")))
ck.conditional_formatting.add(f"D4:D{4+len(CK)}", FormulaRule(formula=['AND(D4<>"INFO",NOT(ISNUMBER(SEARCH("OK",D4))))'], fill=PatternFill("solid", fgColor="FFC7CE")))

# ---- Sources
so = wb.create_sheet("Sources")
so["A1"] = "Sources & linked files"; so["A1"].font = TITLE
setw(so, {"A": 34, "B": 110, "C": 22})
SRC = [("Capacity", "/workspace/nebius/capacity/nebius_capacity.xlsx (Sites, Quarterly MW copied as values; Vineland tranche track)", "Capacity & Power"),
       ("GPU price history", "/workspace/nebius/market/gpu_price_history.xlsx (H100 decay -> residual curve)", "GPU Market"),
       ("Pricing", "nebius.com/prices snapshot Oct 4 2026 (effective Oct 1 2026)", "GPU Market"),
       ("Consensus", "/workspace/nebius/skeptic/01_priced_in.xlsx (Oct 2 2026; share price 242.81)", "Skeptic"),
       ("Filings", "/workspace/nebius/filings/nebius_filings.xlsx (Quarterly, Converts, CapStructure, SOTP, Controls_Depreciation)", "Filings & Capital"),
       ("Microsoft 6-K", "https://www.sec.gov/Archives/edgar/data/1513845/000110465925088312/tm2525580d1_6k.htm", "Filings"),
       ("Q2'26 6-K", "https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-1.htm (Notes 1, 7, 12)", "Filings"),
       ("A100 list history", "GPU Market archived A100 list data: neocloud A100 composite $1.23 -> $2.00 (Nov 2020 -> Oct 2026); hyperscaler composite -1.8%/yr (Estimated; supports 5-6 yr economic GPU life)", "GPU Market"),
       ("Nebius vs neocloud median OD list, Oct 2026", "H100 $4.50 vs $3.99; H200 $5.40 vs $5.29; B200 $8.50 vs $7.49 (GPU Market, /workspace/nebius/market/gpu_price_history.xlsx)", "GPU Market"),
       ("Comps", "CoreWeave EV $79.5B at 6/30/26, ~1,500 MW active, ~3.7 GW contracted (GPU Market, Derived)", "GPU Market")]
for k, (a_, b_, c_) in enumerate(SRC):
    so[f"A{3+k}"] = a_; so[f"B{3+k}"] = b_; so[f"C{3+k}"] = c_

# ---- Funding plan for liquidity gap (parametric, Estimated)
fu = wb.create_sheet("Funding")
fu["A1"] = "LEGACY v0.4 parametric funding illustration (v0.7 plan is wired in Inputs FP_*, CashFlow, Valuation POST-FUNDING block) - ESTIMATED (illustrative; DCF value/share already reflects capex; dilution arises only from issuing equity/converts vs model value)"; fu["A1"].font = TITLE
FR = [3]
def FL(lab, f, fm="#,##0;(#,##0)", nm=None, inp_=False, note=""):
    r = FR[0]; fu[f"A{r}"] = lab
    if inp_: inp(fu[f"B{r}"], f, "Estimated", fm)
    else: fml(fu[f"B{r}"], f, fm)
    if nm: name(wb, nm, "Funding", f"$B${r}")
    fu[f"C{r}"] = note; fu[f"C{r}"].font = ITAL
    FR[0] += 1; return r
def FH(t):
    FR[0] += 1; fu[f"A{FR[0]}"] = t; fu[f"A{FR[0]}"].font = BOLD; FR[0] += 1
FH("Gap")
FL("Minimum liquidity Q3'26-Q4'27 (no plug)", "=MIN(INDEX(Liq,1,11):INDEX(Liq,1,16))", nm="FU_Min27")
FL("Minimum liquidity Q3'26-Q4'28 (no plug)", "=MIN(INDEX(Liq,1,11):INDEX(Liq,1,20))", nm="FU_Min28")
FL("Funding need through FY27 = MinCash - min liquidity", "=MAX(0,MinCash-FU_Min27)", nm="FU_Gap27")
FL("Funding need through FY28", "=MAX(0,MinCash-FU_Min28)", nm="FU_Gap28")
FL("Plan target (1 = through FY27, 2 = through FY28)", 1, "0", nm="FU_TgtSel", inp_=True)
FL("Plan target ($M)", "=IF(FU_TgtSel=2,FU_Gap28,FU_Gap27)", nm="FU_Tgt")
FL("Model FD value/share (V_PS) and implied FD shares (M)", "=V_PS", "0.00", nm="FU_V")
FL("  FD shares at value (M) = V_EQ / V_PS", "=IFERROR(V_EQ/V_PS,0)", "#,##0.0", nm="FU_N")
FH("Source inputs (Estimated)")
FL("ATM: shares remaining (M)", "=ATM_Rem", "0.00", nm="FU_ATMsh")
FL("ATM: issue price = SharePx x (1 - discount/fees)", "=SharePx*(1-FU_ATMd)", "0.00", nm="FU_ATMpx")
FL("  ATM discount + fees", 0.02, "0.0%", nm="FU_ATMd", inp_=True)
FL("Customer prepayments: incremental beyond PrepayPct ($M)", 1500, nm="FU_PreAmt", inp_=True, note="Q2 letter: >$9B prepayments in 2026; 70% of deals, 50-60% of capex")
FL("  Prepayment price concession (% of amount, PV cost)", 0.0, "0.0%", nm="FU_PreCost", inp_=True)
FL("ABS / GPU-backed debt: LTV on GPU net book value at Q2'27", 0.5, "0%", nm="FU_LTV", inp_=True)
FL("  GPU net book value at Q2'27 ($M)", "=INDEX(NBV_H,1,14)+INDEX(NBV_B,1,14)+INDEX(NBV_N,1,14)", nm="FU_NBV")
FL("  ABS capacity ($M) = LTV x GPU NBV, capped", "=MIN(FU_LTV*FU_NBV,FU_ABSCap)", nm="FU_ABSMax")
FL("  ABS cap ($M)", 5000, nm="FU_ABSCap", inp_=True)
FL("  ABS interest rate", 0.09, "0.0%", nm="FU_ABSr", inp_=True, note="memo: cost in WACC; no share dilution")
FL("New converts: amount ($M)", 3000, nm="FU_CVAmt", inp_=True)
FL("  Conversion premium", 0.35, "0%", nm="FU_CVPrem", inp_=True)
FL("  Conversion price", "=SharePx*(1+FU_CVPrem)", "0.00", nm="FU_CVPx")
FH("Per-source effect if used ALONE at full size (value/share = (V_EQ + equity-like proceeds) / (shares + new shares))")
for k, (lab, amt, shs, val) in enumerate([
        ("ATM (all 12.27M sh)", "FU_ATMsh*FU_ATMpx", "FU_ATMsh", "(V_EQ+FU_ATMsh*FU_ATMpx)/(FU_N+FU_ATMsh)"),
        ("Customer prepayments", "FU_PreAmt", "0", "FU_V-FU_PreAmt*FU_PreCost/FU_N"),
        ("ABS / GPU-backed debt", "FU_ABSMax", "0", "FU_V"),
        ("New converts (if converted when value > conv. price)", "FU_CVAmt", "IF(FU_V>FU_CVPx,FU_CVAmt/FU_CVPx,0)", "IF(FU_V>FU_CVPx,(V_EQ+FU_CVAmt)/(FU_N+FU_CVAmt/FU_CVPx),FU_V)")]):
    r = FR[0]; fu[f"A{r}"] = lab
    fml(fu[f"B{r}"], f"={amt}", "#,##0"); fml(fu[f"C{r}"], f"={shs}", "0.00"); fml(fu[f"D{r}"], f"={val}", "0.00"); fml(fu[f"E{r}"], f"=D{r}-FU_V", "0.00;-0.00"); fml(fu[f"F{r}"], f"=IFERROR(C{r}/FU_N,0)", "0.0%")
    FR[0] += 1
fu.cell(FR[0] - 5, 2, "Proceeds $M").font = BOLD; fu.cell(FR[0] - 5, 3, "New shares M").font = BOLD; fu.cell(FR[0] - 5, 4, "Value/share $").font = BOLD; fu.cell(FR[0] - 5, 5, "Change $/sh").font = BOLD; fu.cell(FR[0] - 5, 6, "Dilution %").font = BOLD
FH("Combined plan to fill the target: order prepayments -> ATM -> ABS -> converts")
a1 = FL("Prepayments used", "=MIN(FU_PreAmt,FU_Tgt)", nm="FU_P1")
a2 = FL("ATM used ($M)", "=MIN(FU_ATMsh*FU_ATMpx,FU_Tgt-FU_P1)", nm="FU_P2")
a3 = FL("ABS used", "=MIN(FU_ABSMax,FU_Tgt-FU_P1-FU_P2)", nm="FU_P3")
a4 = FL("Converts used", "=MIN(FU_CVAmt,FU_Tgt-FU_P1-FU_P2-FU_P3)", nm="FU_P4")
FL("Unfunded remainder (needs equity plug / more sources)", "=FU_Tgt-FU_P1-FU_P2-FU_P3-FU_P4", nm="FU_Rem")
FL("New shares: ATM + converts (if converted)", "=IFERROR(FU_P2/FU_ATMpx,0)+IF(FU_V>FU_CVPx,FU_P4/FU_CVPx,0)", "0.00", nm="FU_Sh")
FL("Combined value/share", "=(V_EQ+FU_P2+IF(FU_V>FU_CVPx,FU_P4,0)-FU_P1*FU_PreCost)/(FU_N+FU_Sh)", "0.00", nm="FU_VC")
FL("Change vs model value/share ($)", "=FU_VC-FU_V", "0.00;-0.00", nm="FU_dV")
FL("Dilution (% of FD shares)", "=IFERROR(FU_Sh/FU_N,0)", "0.0%")
FL("Memo: ABS annual interest ($M)", "=FU_P3*FU_ABSr")
fu.column_dimensions["A"].width = 70
for c_ in "BCDEF": fu.column_dimensions[c_].width = 14
bri = wb.create_sheet("Bridge")
bri["A1"] = "EV bridge v0.1 -> v0.3 -> v0.4 (STATIC, filled by snap.py)"; bri["A1"].font = TITLE

# ---- README
rd["A1"] = "Nebius (NBIS) driver model - README"; rd["A1"].font = TITLE
RDL = ["Save path: /workspace/nebius/model/nebius_model.xlsx. Quarterly Q1'24-Q2'26 actuals (Hist, tagged Reported/Derived/Estimated/Not disclosed + best estimate + source), Q3'26-Q4'30 forecast, FY roll-ups AI:AO.",
       "Scenario: Inputs!D4 (1 Bull / 2 Base / 3 Bear). Blue = input; orange = placeholder; yellow = estimated/consensus; black = formula.",
       "Flow: Inputs -> Sites (contracted / energized [Capacity file] / mgmt-connected; MWSel; BillLag; PUE basis) -> Pricing -> Contracts (tranches, power ceiling, Q4 bridge) -> Revenue (3 books) -> Capex_DA (vintage cohorts, PISLag, impairment, cap. interest, dep-yield) -> Financing -> PnL -> CashFlow -> Shares -> Valuation (DCF, dilution solve, EV bridge, EV/MW) -> KPIs / Outputs / Sensitivity / Checks.",
       "Circularity: no Excel iterative calc needed. Valuation solves value/share by 15 unrolled fixed-point steps; DilPxMode 1 = SharePx for quarterly dilution.",
       "Converts: DilMethod 1 if-converted all / 2 ITM only (default; step dilution at conversion prices) / 3 net-share. Converted principal is removed from debt (Checks row 1).",
       "Static snapshot tables on Outputs / Sensitivity / Valuation memo rows / Checks (bull 3x, scenario-invariance) come from LibreOffice recalcs in snap.py; re-run snap.py after changing inputs."]
for k, t in enumerate(RDL):
    rd[f"A{3+k}"] = t
rd.column_dimensions["A"].width = 200
cv = wb.create_sheet("Cover")
cv["A1"] = "Nebius (NBIS) driver-based model"; cv["A1"].font = TITLE
cv["A2"] = "VERSION v0.8d – DRAFT FOR REVIEW (v0.8c + split uncontracted pricing leased GB300 / owned Rubin, renewal at 5-yr rate incl. terminal, convert coupons charged, new shares issued at today's value (R5 off))"
cv["A4"] = ("v0.7d changes (each a Bridge step): (1) new shares added to Sh_Basic / fully diluted at the headline issue price (Shares tab plan rows; authority check vs ~82M); "
 "(2) off-B/S commitments weight 50% -> 0%: the $19.0B energy + lease commitments are the same rent / power already deducted in EBITDA (double count removed; EV unchanged, equity +$9.5B); "
 "(3) Vineland T5-T9 billing aligned to Capacity v2 billing-start quarters (no extra 4-qtr ramp / 50% start quarter): FY27 $447M -> $1,159.5M; "
 "(4) HEADLINE PT = SELF-CONSISTENT: ATM, equity top-ups and re-funded converts issued at the solved PT itself (fixed point); spot ($242.81) and path issuance kept as MEMO rows that depend on market access at today's price. "
 "v0.7c changes kept: shared floor bridge, Rubin $16.2M derived premium, fit-out out of leased Rubin, idle Vineland MW, 3-yr option, run-off residual, re-sized funding, Vineland T1-T4 bear sensitivity.")
cv["A3"] = "Build date: 2026-10-04. v0.3 base approved (terminal excludes Q4'30 spend for post-horizon MW and that MW; GPU capex lead 3 qtrs vs billing; history scenario-invariant; bear replacement at bear price path; bull Blackwell 6.25% cut after Rubin + next-gen -1.7%/yr)."
cv["A12"] = "NOTE: Headline PT = self-consistent issuance; spot-issuance PT is a memo that assumes continued market access at today's price."; cv["A12"].font = Font(bold=True, italic=True, color="C00000")
cv["A5"] = "Key outputs (live, current scenario)"; cv["A5"].font = BOLD
for k, (lab, f, fm) in enumerate([("Scenario", '=CHOOSE(Scen,"BULL","BASE","BEAR")', "@"), ("DCF EV ($M)", "=V_EV", "#,##0"), ("FD value / share ($)", "=V_PS", "0.00"), ("12-month PT ($)", "=V_PT", "0.00"),
                                  ("First quarter liquidity < 0", "=Liq_NegQ", "@"), ("Market EV ($M)", "=V_MKTEV", "#,##0")]):
    cv[f"A{6+k}"] = lab; fml(cv[f"B{6+k}"], f, fm, True)
cv["A13"] = "Notes - open items"; cv["A13"].font = BOLD
for k, (t, own) in enumerate([
        ("OPEN: FY26 capex below the $20-25B guide (Aug 12 2026 call): v0.3 $18.1B (~$1.9B short); v0.4 split lead $16.7B (~$3.3B short); possible early 2027 shell/power or deposits not in the path. Check: Checks 'FY26 capex vs company guide'.", "Filings / Capacity"),
        ("OPEN: first negative-liquidity quarter (Q2'27 uniform 3-qtr lead vs Q3'27 split lead) is PROVISIONAL pending June MW.", "Capacity"),
        ("OPEN: Rubin $/GPU-hr premium vs GB300 (no public price; default per-MW parity). New-deal pricing check vs IREN ~$20M/IT MW-yr flags model GB300 contract rate.", "GPU Market"),
        ("OPEN: AIB CLT1 prepayment (undisclosed; default $0) and Inferize price (press $100-150M, toggle off).", "Filings"),
        ("OPEN: Vineland shutdown default OFF (VinShutSel 0); scenarios/weights are Capacity judgment (prob-weighted Q4'26 loss ~$68M at $14.5M/IT MW-yr vs Capacity ~$67M); SLA credits undisclosed; genset power saving not credited.", "Capacity"),
        ("OPEN: Capacity proposed MW path is a scenario only (MWSel 4) until Capacity's Tuesday file; new-deal $/MW basis (IT vs gross) medium-low confidence; short-term share 10% is an estimate.", "Capacity / Filings"),
        ("OPEN: ARR scope / whether $3.0B ARR includes signed-not-billing tranches; unexplained ARR residual ~$0.6B shown as a line, not plugged.", "Filings"),
        ("OPEN: YE27 ~950 MW IT normalization is a placeholder pending Capacity's basis column.", "Capacity"),
        ("OPEN: Q3'26 transfer into service (model $4,702M) pending Q3 filing; depreciation start dates (material weakness) flagged.", "Filings"),
        ("FLAGS: bull EV > 3x base (bull: multi-year $25M + short-term $50M/IT MW-yr, 6.25% cut); FY30 adj. EBITDA margin 69.9% (just inside peer +8pp; 64.9% if PD/SG&A held at FY28 %); FY26 capex $16.7B below $20-25B guide; bear EV < 0 (equity floored at 0). See Checks.", "Modeling / GPU Market")]):
    cv[f"A{14+k}"] = t; cv[f"B{14+k}"] = "Owner: " + own
cv.column_dimensions["A"].width = 150; cv.column_dimensions["B"].width = 24
# ---------------- v0.6: Payback & asset floor tab
pf = wb.create_sheet("Payback_Floor")
pf["A1"] = "Payback check, new-deal price cases, EBITDA per IT MW by generation, and asset floor (live formulas). Inputs blue."; pf["A1"].font = TITLE
def PI(r, lab, nm, v, st, note, fm="0.0"):
    pf[f"A{r}"] = lab; inp(pf[f"B{r}"], v, st, fm); pf[f"C{r}"] = note; pf[f"C{r}"].font = ITAL; name(wb, nm, "Payback_Floor", f"$B${r}")
PI(3, "Leased GB300 capex per IT MW ($M; GPU + network + storage, no shell) = CPM_Lease_GB300 (rack-based)", "PB_CapGB", "=CPM_Lease_GB300", "Derived", "v0.7b: 6.8 racks x GB_RackPx + $6.1M network (v0.7 41.3)")
PI(4, "Leased Vera Rubin capex per IT MW ($M) - live scenario = CPM_Lease_VR", "PB_CapRubin", "=CPM_Lease_VR", "Derived", "Capacity v2 ex $1M colo fit-out (v0.7c): $33.3 / 40.1 / 48.2M bull/base/bear (racks $6.7/7.8/9.1M; network $5.3M)")
PI(5, "Owned GB300 capex per IT MW ($M, incl. data hall) = CPM_Own_GB300", "PB_CapOwnGB", "=CPM_Own_GB300", "Derived", "v0.7b: rack-based + $12.5M hall (v0.7 53)")
PI(6, "Owned Vera Rubin capex per IT MW ($M) = CPM_Own_VR", "PB_CapOwnRubin", "=CPM_Own_VR", "Derived", "Capacity v2 base $52.6M (v0.7 56)")
PI(7, "Opex per IT MW-yr (rent $2.0-2.5M + power + ops), fixed across generations", "PB_Opex", 4.5, "Estimated", "Range $4-5M; does NOT fall at the Rubin switch")
PI(8, "Management payback target (months)", "PB_Target", 22, "Reported", "Q2'26 letter: 1 year 10 months, prepayments excluded (footnote basis)", "0")
PI(9, "Asset life for unlevered IRR (yrs; capex at t0, flat EBITDA, residual at end: GPUs GPU_Resid, owned data hall at HallLife)", "PB_Life", 5, "Estimated", "GPU life 5 yrs (model base)", "0")
for r_, lab_, nm_, f_, fm_ in [(3, "Owned opex per IT MW-yr (power + ops, NO rent) = OwnOpex", "PB_OpexOwn", "=IF(OwnOpexOn=1,OwnOpex,LoadFactor*PUE*8760*PowerPx/1000000+OM_MW)", "0.00"),
                             (4, "Owned data hall + power/cooling $M/IT MW (in owned capex; not replaced on GPU cycle) = CPM_SP", "PB_Hall", "=CPM_SP", "0.00"),
                             (5, "Data hall economic life (yrs; residual only) = HallLife", "PB_HallLife", "=HallLife", "0"),
                             (6, "GPU/server residual at refresh (% of servers+network) = GPU_Resid", "PB_Resid", "=GPU_Resid", "0%"),
                             (7, "WACC (NPV discount rate)", "PB_WACC", "=WACC", "0.0%"),
                             (8, "Corporate opex % revenue FY29 (other COR + PD + SG&A ex-SBC; for after-tax NPV incl. overhead)", "PB_Corp", "=OtherCOR+INDEX(PDpct,1,2029-FYBase)+INDEX(SGApct,1,2029-FYBase)", "0.0%")]:
    pf[f"E{r_}"] = lab_; fml(pf[f"F{r_}"], f_, fm_); name(wb, nm_, "Payback_Floor", f"$F${r_}")
pf["A10"] = "A) New-deal price cases - $/IT MW-yr next to implied $/GPU-hr. GB300: one linked input (ND_Px x 491 x 8,760 x ND_Occ), IT vs gross toggle. RUBIN: direct $/IT MW-yr (no IT/gross toggle)"; pf["A10"].font = BOLD
for j, t_ in enumerate(["Case", "IT reading $M/IT MW", "IT $/GPU-hr", "Gross reading $M/IT MW", "Gross $/GPU-hr", "$/GPU-hr at 400/MW (ref only)"]): pf.cell(11, 1 + j, t_).font = BOLD
pf["A12"] = "GB300 NVL72 (491 GPUs/IT MW)"; fml(pf["B12"], "=MYRate", "0.0"); fml(pf["C12"], "=ND_Px", "0.00"); fml(pf["D12"], "=MYRate*ND_GrossX", "0.0"); fml(pf["E12"], "=ND_Px*ND_GrossX", "0.00")
for i, (lab, v) in enumerate([("Vera Rubin $14.5M/IT MW [bear]", "14.5"), ("Vera Rubin $16.2M/IT MW [base]", "16.2"), ("Vera Rubin $20M/IT MW [bull]", "20"), ("Vera Rubin at live RubinMW", "RubinMW")]):
    r = 13 + i; pf[f"A{r}"] = lab
    fml(pf[f"B{r}"], f"={v}", "0.0"); fml(pf[f"C{r}"], f"={v}*1000000/(GPUMW_N*8760)", "0.00"); pf[f"D{r}"] = "same (no toggle)"; fml(pf[f"F{r}"], f"={v}*1000000/(RubinGPUMW_D*8760)", "0.00")
pf["A17"] = "Benchmarks: IREN new 3-yr GB300 ~$4.65/GPU-hr (>$20M/IT MW); Rubin: Anthropic-Nscale ~$16.2M/MW (6-yr, MW basis unknown); WhiteFiber-Prime Intellect 576 VR200 ~$108.2M/3 yrs = ~$7.15/GPU-hr (~$20M/IT MW if packages counted, medium). GB300 basis not disclosed by management."; pf["A17"].font = ITAL
pf["A19"] = "B) EBITDA per IT MW-yr by generation (opex fixed across generations; Rubin = RubinMW, no toggle)"; pf["A19"].font = BOLD
for j, t_ in enumerate(["", "GB300 IT", "Rubin (RubinMW)", "Change", "GB300 gross", "Rubin (RubinMW)", "Change"]): pf.cell(20, 1 + j, t_).font = BOLD
pf["A21"] = "Revenue per IT MW-yr"; pf["A22"] = "Opex per IT MW-yr (PB_Opex)"; pf["A23"] = "EBITDA per IT MW-yr"; pf["A24"] = "Memo: model's own opex/IT MW-yr - B: LEASED = RentMW x PUE + power (LoadFactor x PUE x 8,760 x PowerPx); C: OWNED = power + O&M, no rent (PB_OpexOwn)"
fml(pf["B21"], "=B12", "0.0"); fml(pf["C21"], "=RubinMW", "0.0"); fml(pf["E21"], "=D12", "0.0"); fml(pf["F21"], "=RubinMW", "0.0")
for c in "BCEF": fml(pf[f"{c}22"], "=PB_Opex", "0.0"); fml(pf[f"{c}23"], f"={c}21-{c}22", "0.0")
for c, a, b in [("D", "B", "C"), ("G", "E", "F")]:
    for r in (21, 23): fml(pf[f"{c}{r}"], f"={b}{r}/{a}{r}-1", "0%")
fml(pf["B24"], "=RentMW*PUE+LoadFactor*PUE*8760*PowerPx/1000000", "0.00"); name(wb, "PB_OpexModel", "Payback_Floor", "$B$24")
fml(pf["C24"], "=PB_OpexOwn", "0.00"); name(wb, "PB_OpexModelOwn", "Payback_Floor", "$C$24")
name(wb, "PB_EBGB", "Payback_Floor", "$B$23"); name(wb, "PB_EBRubin", "Payback_Floor", "$C$23")
pf["A26"] = "C) PAYBACK (months) = capex / (revenue - opex) x 12; prepayments EXCLUDED unless stated. Flag = within +/-3 months of 22. IRR = unlevered pre-tax, capex at t0, flat EBITDA over PB_Life, residual at end (col L). NPV per IT MW at WACC: N pre-tax; O after tax (site level, book D&A shield); P after tax incl. corporate opex PB_Corp"; pf["A26"].font = BOLD
for j, t_ in enumerate(["Case", "IT / Rubin payback", "flag", "Gross (MYRate x ND_GrossX) payback", "flag", "capex $M/IT MW", "opex", "revenue IT $M", "EBITDA/IT MW (IT)", "IRR (IT, with residual)", "IRR (gross, with residual)", "residual at yr PB_Life $M/MW", "IRR (IT) no residual [v0.7 basis]", "NPV/MW pre-tax @WACC", "NPV/MW after tax (site)", "NPV/MW after tax incl. corporate opex", "of which data hall $M/MW"]): pf.cell(27, 1 + j, t_).font = BOLD
_av = lambda a, b, nm="PathRateIT": "AVERAGE(" + ",".join(f"INDEX({nm},1,{k})" for k in range(a, b + 1)) + ")"
_avs = lambda a, b: "AVERAGE(" + ",".join(f"INDEX(MW_OwnSh,1,{k})" for k in range(a, b + 1)) + ")"
# (label, capex, IT revenue, gross revenue or None (= same), opex, capex net factor)
_LP = "+LeasePrepayOn*LeasePrepay"
pbr = [("Leased GB300 (capex incl. $1.6M lease prepayment)", "PB_CapGB" + _LP, "MYRate", "MYRate*ND_GrossX", "PB_Opex", "1", "0"),
       ("Leased Vera Rubin BASE: $16.2M rev / $40.1M capex ex fit-out (+ lease prepayment)", "40.1" + _LP, "16.2", None, "PB_Opex", "1", "0"),
       ("Leased Vera Rubin BULL: $20M / $33.3M (+ lease prepayment)", "33.3" + _LP, "20", None, "PB_Opex", "1", "0"),
       ("Leased Vera Rubin BEAR: $14.5M / $48.2M (+ lease prepayment)", "48.2" + _LP, "14.5", None, "PB_Opex", "1", "0"),
       ("Leased Vera Rubin: bull price $20M at base capex $40.1M (+ lease prepayment)", "40.1" + _LP, "20", None, "PB_Opex", "1", "0"),
       ("Owned GB300 (memo, rack-based servers+network + 12.5 hall), opex NO rent", "PB_CapOwnGB", "MYRate", "MYRate*ND_GrossX", "PB_OpexOwn", "1", "PB_Hall"),
       ("Owned Vera Rubin (memo, base 52.6 = 40.1 + 12.5 hall) at $16.2M, opex NO rent", "PB_CapOwnRubin", "16.2", None, "PB_OpexOwn", "1", "PB_Hall"),
       ("MODEL'S OWN: GB300-era path capex (avg PathRateIT + PathRateSP 2027) + model opex at 2027 owned share", _av(13, 16) + "+" + _av(13, 16, "PathRateSP") + f"+(1-{_avs(13, 16)})*LeasePrepayOn*LeasePrepay", "MYRate", "MYRate*ND_GrossX", f"(1-{_avs(13, 16)})*PB_OpexModel+{_avs(13, 16)}*PB_OpexModelOwn", "1", _av(13, 16, "PathRateSP")),
       ("MODEL'S OWN: Rubin-era path capex (avg PathRateIT + PathRateSP 2029) + model opex at 2029 owned share, live RubinMW", _av(21, 24) + "+" + _av(21, 24, "PathRateSP") + f"+(1-{_avs(21, 24)})*LeasePrepayOn*LeasePrepay", "RubinMW", None, f"(1-{_avs(21, 24)})*PB_OpexModel+{_avs(21, 24)}*PB_OpexModelOwn", "1", _av(21, 24, "PathRateSP")),
       ("TEST: GB300 IT at $32M capex/IT MW (IT fits 22 mo only at ~$32M = 12 x 32 / (22 - 4.5))", "32", "MYRate", "MYRate*ND_GrossX", "PB_Opex", "1", "0"),
       ("TEST: GB300 IT with lower opex $2M/IT MW-yr (ex-rent definition)", "PB_CapGB", "MYRate", "MYRate*ND_GrossX", "2", "1", "0"),
       ("TEST: GB300 IT with prepayments counted as cash (55% of capex on 70% of deals = 38.5%)", "PB_CapGB", "MYRate", "MYRate*ND_GrossX", "PB_Opex", "(1-0.55*0.7)", "0")]
_ann = "(1-(1+PB_WACC)^-PB_Life)/PB_WACC"; _d5 = "(1+PB_WACC)^-PB_Life"
for i, (lab, cap, rv, rg, ox, nf, hl) in enumerate(pbr):
    r = 28 + i; pf[f"A{r}"] = lab
    fml(pf[f"F{r}"], f"={cap}", "0.0"); fml(pf[f"G{r}"], f"={ox}", "0.00"); fml(pf[f"H{r}"], f"={rv}", "0.0"); fml(pf[f"I{r}"], f"=H{r}-G{r}", "0.0")
    fml(pf[f"Q{r}"], f"={hl}", "0.00")
    fml(pf[f"L{r}"], f"=PB_Resid*(F{r}-Q{r})+Q{r}*MAX(0,1-PB_Life/PB_HallLife)", "0.0")
    fml(pf[f"B{r}"], f"=IFERROR(F{r}*{nf}/(H{r}-G{r})*12,999)", "0.0")
    fml(pf[f"J{r}"], f"=IFERROR(RATE(PB_Life,I{r},-F{r}*{nf},L{r}),-1)", "0.0%")
    fml(pf[f"M{r}"], f"=IFERROR(RATE(PB_Life,I{r},-F{r}*{nf}),-1)", "0.0%")
    fml(pf[f"N{r}"], f"=-F{r}*{nf}+I{r}*{_ann}+L{r}*{_d5}", "0.0")
    _dep = f"((F{r}-Q{r})/PB_Life+Q{r}*(ShellMW/(ShellMW+PowerMW)/Life_Shell+PowerMW/(ShellMW+PowerMW)/Life_Pow))"
    _nbvh = f"Q{r}*MAX(0,1-PB_Life*(ShellMW/(ShellMW+PowerMW)/Life_Shell+PowerMW/(ShellMW+PowerMW)/Life_Pow))"
    _rat = f"(PB_Resid*(F{r}-Q{r})*(1-TaxRate)+Q{r}*MAX(0,1-PB_Life/PB_HallLife)-TaxRate*(Q{r}*MAX(0,1-PB_Life/PB_HallLife)-{_nbvh}))"
    fml(pf[f"O{r}"], f"=-F{r}*{nf}+(I{r}-TaxRate*MAX(0,I{r}-{_dep}))*{_ann}+{_rat}*{_d5}", "0.0")
    fml(pf[f"P{r}"], f"=-F{r}*{nf}+((I{r}-PB_Corp*H{r})-TaxRate*MAX(0,I{r}-PB_Corp*H{r}-{_dep}))*{_ann}+{_rat}*{_d5}", "0.0")
    if rg:
        fml(pf[f"D{r}"], f"=IFERROR(F{r}*{nf}/({rg}-G{r})*12,999)", "0.0"); fml(pf[f"K{r}"], f"=IFERROR(RATE(PB_Life,{rg}-G{r},-F{r}*{nf},L{r}),-1)", "0.0%")
    else:
        pf[f"D{r}"] = "n/a (no toggle)"
    for c in ("B", "D"):
        fl = chr(ord(c) + 1)
        fml(pf[f"{fl}{r}"], f'=IF(ISNUMBER({c}{r}),IF(ABS({c}{r}-PB_Target)<=3,"WITHIN +/-3","outside"),"")', "@")
rI = 28 + len(pbr)
pf[f"A{rI}"] = "Implied GB300 revenue/IT MW if management's 22 months INCLUDES prepayments (capex x 61.5% x 12 / 22 + opex)"; fml(pf[f"H{rI}"], "=PB_CapGB*(1-0.55*0.7)*12/PB_Target+PB_Opex", "0.0"); name(wb, "PB_ImplRevPre", "Payback_Floor", f"$H${rI}")
pf[f"A{rI+1}"] = ("NOTE (v0.7b): the 2027+ new-deal price is now set from CONTRACTS ($16M/IT MW base = $3.72/GPU-hr; 5-yr deals $14.5-16M: Microsoft ~$14.5M, Astra ~$15.9M), NOT calibrated to management's 22-month payback. "
                  "At $16M no row fits 22 months ex-prepayments; with prepayments counted as cash (38.5% of capex) leased GB300 does (~24 months). Management's payback therefore most likely INCLUDES prepayments and/or uses lower opex. "
                  "Col D 'gross' = MYRate x ND_GrossX (sensitivity only).")
pf[f"A{rI+1}"].font = ITAL; pf[f"A{rI+1}"].alignment = Alignment(wrap_text=True)
name(wb, "PB_GB_IT", "Payback_Floor", "$B$28"); name(wb, "PB_GB_Gr", "Payback_Floor", "$D$28"); name(wb, "PB_RU_IT", "Payback_Floor", "$B$29"); name(wb, "PB_RU_Gr", "Payback_Floor", "$B$31")
name(wb, "PB_RU_IRR", "Payback_Floor", "$J$29"); name(wb, "PB_GB_IRR", "Payback_Floor", "$J$28")
for nm_, ref_ in [("PB_NPV_LR", "$N$29"), ("PB_NPVat_LR", "$O$29"), ("PB_NPVc_LR", "$P$29"), ("PB_NPV_OR", "$N$34"), ("PB_NPVat_OR", "$O$34"), ("PB_NPVc_OR", "$P$34"),
                  ("PB_IRR_OR", "$J$34"), ("PB_IRR_OG", "$J$33"), ("PB_OR_IT", "$B$34"), ("PB_OG_IT", "$B$33"), ("PB_NPVc_MR", "$P$36"), ("PB_NPVc_MG", "$P$35")]:
    name(wb, nm_, "Payback_Floor", ref_)
pf["A44"] = "D) ASSET FLOOR: PV of EBITDA margin on signed RPO + stakes - net debt - capex still needed for contracted MW ($M)"; pf["A44"].font = BOLD
pf["A45"] = "RPO schedule assumption: 36% ($13.5B) evenly Jul'26-Jun'28 (Q2'26 Note 2); remainder 64% spread evenly Jul'28-Jun'31 (12 qtrs; ~avg remaining tenor: Microsoft 5-yr from late 2025, Meta #2 5-yr from 2027 - assumption). Margin = model adj. EBITDA margin by quarter (FY30 margin after 2030); pre-tax; discounted at WACC from ValDate."; pf["A45"].font = ITAL
_df = "1/(1+WACC)^((Q_End-ValDate)/365.25)"
rows = [("RPO total ($M)", "=RPO_Tot", "PF_RPO"),
        ("PV of margin - window Jul'26-Jun'28", f"=SUMPRODUCT((Q_Idx>LastHist)*(Q_Idx<=LastHist+8)*(RPO_Tot*RPO24Pct/8)*PL_EBITDA/PL_Rev*{_df})", "PF_PV1"),
        ("PV of margin - remainder, Jul'28-Dec'30 (10 qtrs in horizon)", f"=SUMPRODUCT((Q_Idx>LastHist+8)*(Q_Idx<=28)*(RPO_Tot*(1-RPO24Pct)/12)*PL_EBITDA/PL_Rev*{_df})", "PF_PV2"),
        ("PV of margin - remainder, Q1-Q2'31 at FY30 margin", f"=RPO_Tot*(1-RPO24Pct)/12*V_M30*(1/(1+WACC)^((INDEX(Q_End,1,28)+91-ValDate)/365.25)+1/(1+WACC)^((INDEX(Q_End,1,28)+182-ValDate)/365.25))", "PF_PV3"),
        ("PV of margin on signed RPO (pre-tax)", "=PF_PV1+PF_PV2+PF_PV3", "PF_PV"),
        ("  memo: after tax at TaxRate (no D&A shield)", "=PF_PV*(1-TaxRate)", "PF_PVat"),
        ("Contracted MW not yet delivered at 6/30 (Microsoft T5-T9 + Meta; backstop & TBD excluded)", "=" + "+".join(f"Contracts!$D${t0+i}*(Contracts!$H${t0+i}>LastHist)*(1-Contracts!$I${t0+i})" for i, t in enumerate(TR) if t[1] in ("Microsoft", "Meta") and "backstop" not in t[0]), "PF_MWund"),
        ("Capex still needed for contracted MW (Microsoft/Vineland = build-to-suit LEASE -> leased GB300 PB_CapGB [v0.7b; was owned 53]; Meta by gen: Rubin PB_CapRubin, Blackwell PB_CapGB) - gross, CIP already spent NOT netted (conservative)", "=" + "+".join(f"Contracts!$D${t0+i}*(Contracts!$H${t0+i}>LastHist)*(1-Contracts!$I${t0+i})*" + ("PB_CapGB" if t[1] == "Microsoft" else ("PB_CapRubin" if t[2] == "N" else "PB_CapGB")) for i, t in enumerate(TR) if t[1] in ("Microsoft", "Meta") and "backstop" not in t[0]), "PF_Cap"),
        ("Operating asset floor (EV-equivalent) = PV margin - capex needed", "=PF_PV-PF_Cap", "PF_EV"),
        ("Net debt at the conservative floor price (secured + converts OTM at that price at principal - cash; Q3'26)", "=V_ODEBT+" + convD("PF_PS1", "V_TrO", "V_TrB", "V_TrRT", "V_TrPX") + "-V_CASH", "PF_ND"),
        ("Stakes - conservative (carrying values, Note 5: ClickHouse 1,517.7 + Toloka 86.5 + VC 13.8 + other 2.5; Avride $0)", "=1620.5", "PF_S1"),
        ("Stakes - mark (ClickHouse $4.2B est., Toloka FV 130.6, VC/other 16.3, Avride $0)", "=CH_Hi+TO_Hi+Other_Val", "PF_S2"),
        ("Stakes - high (mark + Avride $2.75B Caplight, LOW confidence) + PV of Avride burn added back (credit value AND remove burn - v0.7 double-count fix)", "=PF_S2+AV_Hi-AVBurnOn*AVBurnQ*(SUM(Valuation!$F$" + str(vdf) + ":$AG$" + str(vdf) + ")+4*(1+V_g)/(WACC-V_g)*V_DF28)", "PF_S3"),
        ("Equity floor - conservative (at solved floor price)", "=PF_EV-V_ODEBT-" + convD("PF_PS1", "V_TrO", "V_TrB", "V_TrRT", "V_TrPX") + "+V_CASH+PF_S1", "PF_EQ1"),
        ("Equity floor - mark", "=PF_EV-V_ODEBT-" + convD("PF_PS2", "V_TrO", "V_TrB", "V_TrRT", "V_TrPX") + "+V_CASH+PF_S2", "PF_EQ2"),
        ("Equity floor - high", "=PF_EV-V_ODEBT-" + convD("PF_PS3", "V_TrO", "V_TrB", "V_TrRT", "V_TrPX") + "+V_CASH+PF_S3", "PF_EQ3"),
        ("LEGACY v0.7b floor bridge (no off-B/S commitments, no funding, carrying-value stakes) - per FD share cons / mark / high ($) [memo; superseded by shared bridge when BridgeFix=1]", "=PF_PS1x", "PF_PS1"),
        ("HEADLINE equity floor per share - SHARED bridge identical to DCF (OffBSWt x commitments, converts carrying, cash, DR placeholder, V_SOTP, funding cash & shares at the HEADLINE issue-price basis PTBasis); BridgeFix=0 -> legacy cons", "=IF(BridgeFix=1,IF(AND(PTBasis=1,FP_On=1),PF_PSSC,PF_PSH),PF_PS1x)", "PF_PSHead"),
        ("Base DCF EV (live)", "=V_EV", "PF_DCF"),
        ("FLAG: DCF EV vs operating asset floor", '=IF(V_EV<PF_EV,"FLAG: DCF EV below asset floor","OK: DCF EV above asset floor")', "PF_Flag")]
for i, (lab, f, nm) in enumerate(rows):
    r = 46 + i; pf[f"A{r}"] = lab
    fml(pf[f"B{r}"], f, "@" if nm == "PF_Flag" else "#,##0", True if nm in ("PF_EV", "PF_PV", "PF_Flag") else False); name(wb, nm, "Payback_Floor", f"$B${r}")
    if nm == "PF_PS1":
        fml(pf[f"C{r}"], "=PF_PS2x", "0.00"); fml(pf[f"D{r}"], "=PF_PS3x", "0.00"); pf[f"B{r}"].number_format = "0.00"
        name(wb, "PF_PS2", "Payback_Floor", f"$C${r}"); name(wb, "PF_PS3", "Payback_Floor", f"$D${r}")
        fml(pf[f"E{r}"], "=" + convS("PF_PS1", "V_TrO", "V_TrB", "V_TrRT", "V_TrPX"), "0.00"); pf[f"F{r}"] = "<- convert shares at conservative floor price (M)"
# floor fixed-point (no circularity): cols H/I/J = conservative/mark/high
_r0 = 46 + len(rows) + 1
pf[f"G{_r0}"] = "Floor price iteration"; pf[f"H{_r0}"] = "cons"; pf[f"I{_r0}"] = "mark"; pf[f"J{_r0}"] = "high"
for c in "HIJ": fml(pf[f"{c}{_r0+1}"], "=10", "0.00")
for k in range(2, 17):
    rr_ = _r0 + k
    for c, S_ in [("H", "PF_S1"), ("I", "PF_S2"), ("J", "PF_S3")]:
        P = f"{c}{rr_-1}"
        eq = f"(PF_EV-V_ODEBT-{convD(P, 'V_TrO', 'V_TrB', 'V_TrRT', 'V_TrPX')}+V_CASH+{S_})"
        shs = f"(V_BASIC+V_RSU+Options*MAX(0,1-OptK/{P})+{convS(P, 'V_TrO', 'V_TrB', 'V_TrRT', 'V_TrPX')})"
        fml(pf[f"{c}{rr_}"], f"=MAX(0.01,{eq}/{shs})", "0.00")
pf[f"K{_r0}"] = "SHARED pre-fund"; pf[f"L{_r0}"] = "SHARED post-fund"; pf[f"M{_r0}"] = "DCF EV same code (post)"
for c in "KLM": fml(pf[f"{c}{_r0+1}"], "=10", "0.00")
for k in range(2, 17):
    rr_ = _r0 + k
    for c, EVn in [("K", "PF_EV"), ("L", "PF_EV"), ("M", "V_EV")]:
        P = f"{c}{rr_-1}"
        if c == "K":
            eq = f"({EVn}-OffBSWt*V_OFF-V_ODEBT-{convD(P, 'V_TrO', 'V_TrB', 'V_TrRT', 'V_TrPX')}+V_CASH+V_SOTP)"
            shs = f"(V_BASIC+V_RSU+Options*MAX(0,1-OptK/{P})+{convS(P, 'V_TrO', 'V_TrB', 'V_TrRT', 'V_TrPX')})"
        else:
            eq = f"({EVn}-OffBSWt*V_OFF-V_ODEBT-{convD(P, 'V_TrOP', 'V_TrBP', 'V_TrRT', 'V_TrPX')}+V_CASH+V_SOTP+V_FPCash)"
            shs = f"(V_BASIC+V_FPSh+V_RSU+Options*MAX(0,1-OptK/{P})+{convS(P, 'V_TrOP', 'V_TrBP', 'V_TrRT', 'V_TrPX')})"
        fml(pf[f"{c}{rr_}"], f"=MAX(0.01,{eq}/{shs})", "0.00")
for c, nm_ in [("K", "PF_PSS"), ("L", "PF_PSSF"), ("M", "PF_DCFRep")]:
    fml(pf[f"{c}{_r0+17}"], f"=IF({c}{_r0+16}<=0.011,0,{c}{_r0+16})", "0.00"); name(wb, nm_, "Payback_Floor", f"${c}${_r0+17}")
pf[f"K{_r0+18}"] = "Floor value/sh on shared bridge (FP_On -> post-funding) / x(1+Ke) / DCF replica check vs V_PSF"
fml(pf[f"K{_r0+19}"], "=IF(FP_On=1,PF_PSSF,PF_PSS)", "0.00"); name(wb, "PF_PSH", "Payback_Floor", f"$K${_r0+19}")
pf[f"N{_r0}"] = "SHARED floor, SELF-CONSISTENT issue price (v)"; pf[f"O{_r0}"] = "issue price P"
_fl = SCROWS(pf, "N", lambda c: "OffBSWt", "PF_EV", _r0 + 1, 30, "O")
fml(pf[f"N{_fl+1}"], f"=IF(N{_fl}<=0.011,0,N{_fl})", "0.00"); name(wb, "PF_PSSC", "Payback_Floor", f"$N${_fl+1}")
fml(pf[f"L{_r0+19}"], "=PF_PSHead*(1+Ke)", "0.00"); name(wb, "PF_PTH", "Payback_Floor", f"$L${_r0+19}")
fml(pf[f"M{_r0+19}"], "=PF_DCFRep-IF(FP_On=1,V_PSSpot,V_PS)", "0.00"); name(wb, "PF_RepChk", "Payback_Floor", f"$M${_r0+19}")
for c, nm_ in [("H", "PF_PS1x"), ("I", "PF_PS2x"), ("J", "PF_PS3x")]:
    fml(pf[f"{c}{_r0+17}"], f"=IF({c}{_r0+16}<=0.011,0,{c}{_r0+16})", "0.00"); name(wb, nm_, "Payback_Floor", f"${c}${_r0+17}")
_x = _r0 + 20
pf[f"A{_x}"] = "E) PRICE CHECK - $/GPU-hr implied by new-deal $/IT MW (live scenario). Price is NOT calibrated off the 22-month payback."; pf[f"A{_x}"].font = BOLD
for k_, (lab_, f_) in enumerate([("GB300 multi-year new deal (MYRate, IT reading) $M/IT MW-yr | $/GPU-hr at 491", "=MYRate"), ("GB300 gross reading (MYRate x ND_GrossX) per IT MW | $/GPU-hr at 491", "=MYRate*ND_GrossX"),
                                 ("Q2'26 letter ACV midpoint $22.5M (ACV_Q2) if GB300 at 491/IT MW | $/GPU-hr (vs MSFT $3.37, Astra $3.70, IREN ~$4.65)", "=ACV_Q2"), ("Vera Rubin new deal (RubinMW) | $/GPU-hr at 320 (col D: Rubin / GB300 per GPU-hr)", "=RubinMW")]):
    rr_ = _x + 1 + k_; pf[f"A{rr_}"] = lab_; fml(pf[f"B{rr_}"], f_, "0.00"); fml(pf[f"C{rr_}"], f"=B{rr_}*1000000/({'GPUMW_N' if k_ == 3 else 491}*8760)", "0.00")
fml(pf[f"D{_x+4}"], f"=C{_x+4}/C{_x+1}", "0.00x"); name(wb, "PX_Ratio", "Payback_Floor", f"$D${_x+4}")
pf[f"E{_x}"] = "Rubin vs GB300 (leased, live scenario)"; pf[f"E{_x}"].font = BOLD
for k_, (lab_, f_, fm_, nm_) in enumerate([("per-GPU-hr price multiple", "=PX_Ratio", "0.00x", None), ("$/IT MW price multiple (RubinMW / MYRate)", "=RubinMW/MYRate", "0.00x", "PX_MWRatio"),
                                          ("capex multiple (CPM_Lease_VR / CPM_Lease_GB300)", "=CPM_Lease_VR/CPM_Lease_GB300", "0.00x", "PX_CapRatio"),
                                          ("price-to-capex: Rubin (RubinMW / CPM_Lease_VR)", "=RubinMW/CPM_Lease_VR", "0.000", "PX_PCR"), ("price-to-capex: GB300 (MYRate / CPM_Lease_GB300)", "=MYRate/CPM_Lease_GB300", "0.000", "PX_PCG"),
                                          ("FLAG if Rubin price-to-capex > GB300 by >10%", '=IF(PX_PCR>1.1*PX_PCG,"FLAG: Rubin price/capex much better than GB300 ("&TEXT(PX_PCR/PX_PCG,"0.00")&"x)","OK: Rubin price/capex "&TEXT(PX_PCR/PX_PCG,"0.00")&"x GB300")', "@", "PX_Flag")]):
    pf[f"E{_x+1+k_}"] = lab_; fml(pf[f"F{_x+1+k_}"], f_, fm_)
    if nm_: name(wb, nm_, "Payback_Floor", f"$F${_x+1+k_}")
name(wb, "PX_GB", "Payback_Floor", f"$C${_x+1}"); name(wb, "PX_GBg", "Payback_Floor", f"$C${_x+2}"); name(wb, "PX_ACV", "Payback_Floor", f"$C${_x+3}"); name(wb, "PX_VR", "Payback_Floor", f"$C${_x+4}")
_y = _x + 7
pf[f"A{_y}"] = "F) REVENUE DILUTION CROSS-CHECK ONLY (Peer KPIs payback_dilution.xlsx Nebius_Path): rev/MW = SS x (1 - k x new MW / avg MW); SS mixes legacy at L0 with new MW at P_new. NOT layered on the price path (would double count)"; pf[f"A{_y}"].font = BOLD
for k_, (lab_, nm_, v_) in enumerate([("k (dilution slope; range 0.43-0.62)", "DX_k", 0.522), ("P_new $M/IT MW-yr (Nebius-Microsoft $17.1M)", "DX_P", 17.086), ("L0 = Q2'26 AI rev x 4 / avg Q2'26 energized IT MW", "DX_L0", None)]):
    rr_ = _y + 1 + k_; pf[f"A{rr_}"] = lab_
    if v_ is None: fml(pf[f"B{rr_}"], "=574.9*4/AVERAGE(INDEX(MW_En,1,9),INDEX(MW_En,1,10))", "0.00")
    else: inp(pf[f"B{rr_}"], v_, "Estimated", "0.000")
    name(wb, nm_, "Payback_Floor", f"$B${rr_}")
_h = _y + 4
for j_, t_ in enumerate(["Quarter idx", "End energized IT MW", "Avg MW", "New MW", "SS $M/MW-yr", "Rev/MW diluted", "Implied AI rev $M", "Model AI rev $M"]): pf.cell(_h, 1 + j_, t_).font = BOLD
for k_, q_ in enumerate(range(11, 17)):
    rr_ = _h + 1 + k_
    fml(pf[f"A{rr_}"], f"={q_}", "0"); fml(pf[f"B{rr_}"], f"=INDEX(MW_En,1,{q_})", "#,##0.0"); fml(pf[f"C{rr_}"], f"=AVERAGE(INDEX(MW_En,1,{q_-1}),INDEX(MW_En,1,{q_}))", "#,##0.0")
    fml(pf[f"D{rr_}"], f"=MAX(0,INDEX(MW_En,1,{q_})-INDEX(MW_En,1,{q_-1}))", "#,##0.0")
    fml(pf[f"E{rr_}"], f"=(DX_L0*AVERAGE(INDEX(MW_En,1,9),INDEX(MW_En,1,10))+DX_P*MAX(0,C{rr_}-AVERAGE(INDEX(MW_En,1,9),INDEX(MW_En,1,10))))/C{rr_}", "0.00")
    fml(pf[f"F{rr_}"], f"=E{rr_}*(1-DX_k*D{rr_}/C{rr_})", "0.00"); fml(pf[f"G{rr_}"], f"=F{rr_}*C{rr_}/4", "#,##0"); fml(pf[f"H{rr_}"], f"=INDEX(AIRev,1,{q_})", "#,##0")
_t = _h + 7
pf[f"A{_t}"] = "FY27 AI cloud revenue: dilution cross-check | model | gap (model - cross-check)"
fml(pf[f"B{_t}"], f"=SUM(G{_h+3}:G{_h+6})", "#,##0"); fml(pf[f"C{_t}"], f"=SUM(H{_h+3}:H{_h+6})", "#,##0"); fml(pf[f"D{_t}"], f"=C{_t}-B{_t}", "#,##0")
name(wb, "DX_FY27", "Payback_Floor", f"$B${_t}"); name(wb, "DX_Model27", "Payback_Floor", f"$C${_t}"); name(wb, "DX_Gap27", "Payback_Floor", f"$D${_t}")
# ---- v0.8 G) breakeven new-deal price (NPV = 0, after tax, after corporate costs PB_Corp)
_b = _t + 3
pf[f"A{_b}"] = "G) v0.8 BREAKEVEN NEW-DEAL PRICE ($M per IT MW-yr) at which a new uncontracted MW is NPV-zero: after tax, after corporate costs (PB_Corp x revenue), capex at t0, flat EBITDA over PB_Life, residual as section C (col P basis). Solved in closed form (P is linear in price)."; pf[f"A{_b}"].font = BOLD
for j_, t_ in enumerate(["Case (section C row)", "WACC 9%", "WACC 10%", "WACC 11%", "live WACC", "current price", "capex $M/IT MW", "opex"]): pf.cell(_b + 1, 1 + j_, t_).font = BOLD
def _be(r, w):
    ann = f"((1-(1+{w})^-PB_Life)/{w})"; d5 = f"((1+{w})^-PB_Life)"
    dep = f"((F{r}-Q{r})/PB_Life+Q{r}*(ShellMW/(ShellMW+PowerMW)/Life_Shell+PowerMW/(ShellMW+PowerMW)/Life_Pow))"
    nbvh = f"Q{r}*MAX(0,1-PB_Life*(ShellMW/(ShellMW+PowerMW)/Life_Shell+PowerMW/(ShellMW+PowerMW)/Life_Pow))"
    rat = f"(PB_Resid*(F{r}-Q{r})*(1-TaxRate)+Q{r}*MAX(0,1-PB_Life/PB_HallLife)-TaxRate*(Q{r}*MAX(0,1-PB_Life/PB_HallLife)-{nbvh}))"
    ht = f"((F{r}-TaxRate*{dep}*{ann}-{rat}*{d5})/((1-TaxRate)*(1-PB_Corp)*{ann})+G{r}/(1-PB_Corp))"
    hn = f"((F{r}-{rat}*{d5})/((1-PB_Corp)*{ann})+G{r}/(1-PB_Corp))"
    return f"=IF({ht}*(1-PB_Corp)-G{r}-{dep}>=0,{ht},{hn})"
BEC = [("Leased GB300 (row 28)", 28, "BE_GB"), ("Leased Vera Rubin base capex (row 29)", 29, "BE_LR"), ("Owned GB300 (row 33)", 33, "BE_OG"), ("Owned Vera Rubin $52.6M (row 34) - Capacity owned uncontracted halls", 34, "BE_OR"),
       ("Model's own GB300-era path (row 35)", 35, "BE_MG"), ("Model's own Rubin-era path (row 36)", 36, "BE_MR")]
for k_, (lab_, r_, nm_) in enumerate(BEC):
    rr_ = _b + 2 + k_; pf[f"A{rr_}"] = lab_
    for c_, w_ in zip("BCDE", ["0.09", "0.1", "0.11", "PB_WACC"]): fml(pf[f"{c_}{rr_}"], _be(r_, w_), "0.0")
    fml(pf[f"F{rr_}"], f"=H{r_}", "0.0"); fml(pf[f"G{rr_}"], f"=F{r_}", "0.0"); fml(pf[f"H{rr_}"], f"=G{r_}", "0.00")
    for c_, w_ in zip("BCDE", ["9", "10", "11", "L"]): name(wb, f"{nm_}{w_}", "Payback_Floor", f"${c_}${rr_}")
pf[f"A{_b + 2 + len(BEC)}"] = "Breakeven excludes funding / dilution effects (self-consistent PT) and contracted-MW economics; the PT-matching price (build-plan PT = stop-building PT) is solved by full-model bisection - see Cover / v0.8 notes."
# ---- v0.8 Funding: reconciliation Capacity uncontracted capex vs model funding need
FH("v0.8 RECONCILIATION: Capacity ~$17-18B capex for uncontracted MW vs model funding need (col B = Q3'26-Q4'27; col D = Q3'26-Q4'30). Signs: + = needs funding, ( ) = funds itself")
fu.cell(FR[0], 2, "Q3'26-Q4'27").font = BOLD; fu.cell(FR[0], 4, "Q3'26-Q4'30").font = BOLD; FR[0] += 1
def _sp(expr, b): return f"=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<={b})*({expr}))"
REC = [("Capacity estimate: capex for uncontracted MW to YE27 (~$17-18B; owned Rubin ~165 MW x $52.6M + leased ~230-250 MW; Capacity, Estimated)", None, "RC_Cap"),
       ("Model: new-build capex on UNCONTRACTED MW (CapNew x (1 - signed share)) - incl. capex spent ahead of go-live (GPULead / DCLead) for sites billing in 2027-28", "CapNew*(1-CF_SigSh)", "RC_Unc"),
       ("  timing / scope difference vs Capacity (model - Capacity; col D = post-YE27 uncontracted build)", "RCDIFF", "RC_Diff"),
       ("+ new-build capex on CONTRACTED MW (Microsoft T5-T9, Meta; no new prepayment on signed deals)", "CapNew*CF_SigSh", "RC_Con"),
       ("+ maintenance capex + GPU refresh / replacement + lease prepayments - after-tax residual proceeds", "CapMaint+CapRepl+CapLPre-CapResid", "RC_Oth"),
       ("TOTAL CAPEX (net of residual)", "Capex-CapResid", "RC_Capex"),
       ("- Adj. EBITDA (rent, power, opex and corporate burn are netted here)", "-PL_EBITDA", "RC_EB"),
       ("+ cash taxes + net cash interest (coupons + secured - interest income; plug interest excluded as in the runway)", "PL_Tax+CV_cpn+Sec_Int+ABS_Int-CF_IntInc", "RC_TI"),
       ("- net new customer prepayments (new prepayments - deferred revenue released)", "-(CF_Prepay-CF_DRRel)", "RC_DR"),
       ("+ other investing (CLT1 prepayment, Inferize)", "-CF_OthInv", "RC_OI"),
       ("+ convert maturities / redemptions + secured + ABS repayments", "CV_rep-MIN(0,Sec_Flow)+ABS_Rep", "RC_Rep"),
       ("+ change in liquidity (end - 6/30 cash $8.0B; minimum buffer FP_MinBuf)", "LIQ", "RC_Liq"),
       ("GROSS EXTERNAL FUNDING NEED", "TOT", "RC_Need"),
       ("- debt-like sources: Aug-26 converts ($5.75B gross) + funding-plan converts (net) + secured draws + ABS draws on contracted MW", "-(Liq_Aug+CF_FPConv+MAX(0,Sec_Flow)+ABS_Draw)", "RC_Debt"),
       ("EQUITY NEED (ATM + equity top-ups)", "EQ", "RC_Eq"),
       ("  check: model ATM + top-ups (should equal the line above)", "CF_ATM+FP_EqTop", "RC_Chk")]
_rr = {}
for lab_, ex_, nm_ in REC:
    r = FR[0]; fu[f"A{r}"] = lab_; _rr[nm_] = r
    for c_, b_ in (("B", 16), ("D", 28)):
        if ex_ is None: f_ = "=17500" if c_ == "B" else f"=B{r}"
        elif ex_ == "RCDIFF": f_ = f"={c_}{_rr['RC_Unc']}-B{_rr['RC_Cap']}"
        elif ex_ == "LIQ": f_ = f"=INDEX(Liq,1,{b_})-CS_Cash"
        elif ex_ == "TOT": f_ = f"=SUM({c_}{_rr['RC_Capex']}:{c_}{r-1})"
        elif ex_ == "EQ": f_ = f"={c_}{_rr['RC_Need']}+{c_}{_rr['RC_Debt']}"
        else: f_ = _sp(ex_, b_)
        if ex_ is None and c_ == "B": inp(fu[f"B{r}"], 17500, "Estimated", "#,##0;(#,##0)")
        else: fml(fu[f"{c_}{r}"], f_, "#,##0;(#,##0)")
    name(wb, nm_, "Funding", f"$B${r}"); name(wb, nm_ + "30", "Funding", f"$D${r}")
    if nm_ in ("RC_Capex", "RC_Need", "RC_Eq"): fu[f"A{r}"].font = BOLD
    FR[0] += 1
fu[f"A{FR[0]}"] = "Why they differ: Capacity's figure is only the capex for uncontracted MW live by YE27. The model's funding need adds contracted-MW capex (no new prepayment on signed deals), maintenance / refresh, convert maturities, the liquidity buffer, and all 2028-30 build. It nets operating EBITDA and new prepayments. '~$43B' = model capex Q3'26-Q4'27 (col B total capex); equity need Q3'26-Q4'30 = col D."; FR[0] += 1
fu.column_dimensions["D"].width = 14
pf.column_dimensions["A"].width = 95
for c in "BCDEFG": pf.column_dimensions[c].width = 16
# ---- v0.8b FUNDING_STACK sheet (live block; static grid written by fs_grid.py)
fs = wb.create_sheet("Funding_Stack")
fs["A1"] = "v0.8b FUNDING STACK - live toggles (Inputs: NDPrepayOn/NDPrepayPct, ABS_LTC/ABS_Spread/ABS_Term, UncPxOn/UncPx/UncRenew, BackstopMW) and static grid (self-consistent PT basis)"; fs["A1"].font = BOLD
_fsr = [("New-deal prepayment toggle (NDPrepayOn) / share of new-build capex on non-signed MW", "=NDPrepayOn", "0", None), ("  NDPrepayPct", "=NDPrepayPct", "0%", None),
        ("ABS loan-to-cost on contracted-MW capex (ABS_LTC)", "=ABS_LTC", "0%", None), ("  ABS rate = SOFR + spread", "=SOFR+ABS_Spread", "0.00%", "FS_Kd"),
        ("  ABS after-tax cost = rate x (1 - TaxRate)", "=FS_Kd*(1-TaxRate)", "0.00%", "FS_Kdat"),
        ("Unsold new-MW price toggle (UncPxOn) / UncPx $M per IT MW-yr", "=IF(UncPxOn=1,UncPx,0)", "0.0", None), ("v0.8d SPLIT uncontracted pricing on (UncSplitOn): leased GB300 $M / term yrs / renewal $M; owned Rubin $M", "=IF(UncSplitOn=1,TEXT(UncGBPx,\"0.0\")&\" / \"&UncGBTermY&\"y / \"&TEXT(UncGBRen,\"0.0\")&\"; Rubin \"&TEXT(UncRubPx,\"0.0\"),\"off\")", "@", None), ("Meta backstop MW", "=BackstopMW", "0", None), ("Stop building", "=StopBuild", "0", None),
        ("New prepayments Q3'26-Q4'30 ($M)", "=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<=28)*CF_Prepay)", "#,##0", "FS_Pre"),
        ("ABS draws Q3'26-Q4'30 ($M)", "=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<=28)*ABS_Draw)", "#,##0", "FS_ABS"), ("ABS peak balance ($M)", "=MAX(ABS_Bal)", "#,##0", "FS_ABSPk"),
        ("ABS interest Q3'26-Q4'30 ($M)", "=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<=28)*ABS_Int)", "#,##0", "FS_ABSInt"),
        ("Total capex Q3'26-Q4'30 ($M)", "=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<=28)*Capex)", "#,##0", "FS_Cx"),
        ("ABS share of funding mix = ABS draws / total capex Q3'26-Q4'30", "=IFERROR(FS_ABS/FS_Cx,0)", "0.0%", "FS_wABS"),
        ("Base WACC implied for the rest of the mix", 0.11, "0.0%", "FS_WBase"),
        ("BLENDED WACC = ABS share x after-tax ABS cost + (1 - share) x base 11%", "=FS_wABS*FS_Kdat+(1-FS_wABS)*FS_WBase", "0.00%", "FS_WACCb"),
        ("Equity top-ups + ATM Q3'26-Q4'30 ($M)", "=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<=28)*(FP_EqTop+CF_ATM))", "#,##0", "FS_Eq"),
        ("2027 plan converts (net, $M)", "=SUM(CF_FPConv)", "#,##0", "FS_CV"),
        ("Self-consistent PT (live) / EV / new shares / YE27 liquidity", "=V_PT", "0.00", None)]
for k_, (lab_, f_, fm_, nm_) in enumerate(_fsr):
    r_ = 3 + k_; fs[f"A{r_}"] = lab_
    if isinstance(f_, float): inp(fs[f"B{r_}"], f_, "Input", fm_)
    else: fml(fs[f"B{r_}"], f_, fm_)
    if nm_: name(wb, nm_, "Funding_Stack", f"$B${r_}")
_r = 3 + len(_fsr) - 1
fml(fs[f"C{_r}"], "=V_EV", "#,##0"); fml(fs[f"D{_r}"], "=V_NewShHead", "#,##0.0"); fml(fs[f"E{_r}"], "=INDEX(Liq,1,16)", "#,##0")
fs["A24"] = "STATIC GRID (fs_grid.py; full recalcs) - see rows below"; fs["A24"].font = BOLD
fs.column_dimensions["A"].width = 92
for c_ in "BCDEFGHIJ": fs.column_dimensions[c_].width = 14
order = ["Cover", "README", "Bridge", "Payback_Floor", "Funding", "Funding_Stack", "Outputs", "Checks", "Inputs", "Timeline", "Hist", "Sites", "Pricing", "Contracts", "Revenue", "Capex_DA", "Financing", "PnL", "CashFlow", "Shares", "Valuation", "KPIs", "Sensitivity", "Sources"]
wb._sheets = [wb[n] for n in order]
OUT = "/workspace/nebius/model/nebius_model.xlsx"
import os
os.makedirs("/workspace/nebius/model", exist_ok=True)
wb.save(OUT)
print("saved", OUT, wb.sheetnames)
