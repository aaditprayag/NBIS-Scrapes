from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
BLUE=Font(color="0000FF"); BLK=Font(color="000000"); B=Font(bold=True)
HDR=PatternFill("solid",fgColor="1F3864"); INF=PatternFill("solid",fgColor="DDEBF7")
FB="/workspace/nebius/filings/nebius_fact_base.md"
Q2FS="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm"
F20="https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm"
MS="https://www.marketscreener.com/quote/stock/NEBIUS-GROUP-8037501/finances/"
SHL="https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf"
wb=Workbook(); A=wb.active; A.title="Assumptions"
A.append(["Input","Value","Unit","Tag","Source / note"])
C={}
def inp(key,label,val,unit,tag,src,formula=False):
    A.append([label,val,unit,tag,src]); r=A.max_row; C[key]=f"Assumptions!$B${r}"
    A.cell(r,2).font=BLK if formula else BLUE
    if not formula: A.cell(r,2).fill=INF
def sec(t): A.append([t]); A.cell(A.max_row,1).font=B
sec("Useful lives (scenario columns on Model sheet)")
inp("L1","Life scenario A (yrs) - Nebius pre-2026 policy",4,"yrs","Reported (policy)",F20+" Note 2: server & network 4.0 yrs (FY2025)")
inp("L2","Life scenario B (yrs) - Nebius stated policy from 1/1/26",5,"yrs","Reported (policy)",Q2FS+" Note 1 p.F-10; "+FB)
inp("L3","Life scenario C (yrs) - CoreWeave/Alphabet/Oracle-style",6,"yrs","Reported (peer policy)","See Peers sheet")
inp("ngl","Non-GPU (DC/infra) blended life",12,"yrs","Estimated","Nebius: buildings 20y, infrastructure 3-10y, other 2-10y (20-F Note 2); 12y blended is an assumption")
sec("Asset base at 30-Jun-2026 (Reported)")
inp("srv","Server & network equipment, gross, in service",5197.7,"$M","Reported",Q2FS+" Note 7")
inp("age","Avg age of in-service server base at 6/30/26",4,"quarters","Estimated","60% of gross added in 2025, 40% in 1H26 (Note 7 gross 3,123.6 YE25 -> 5,197.7) => ~4 qtrs")
inp("oth","Other in-service PP&E gross (infra 332.5 + other 189.5 + land/bldg 593.4)",1115.4,"$M","Reported (sum)",Q2FS+" Note 7; land not separately split - depreciated anyway (slightly overstates D&A)")
inp("niu","Assets not yet in use",7836.4,"$M","Reported",Q2FS+" Note 7")
inp("niu3","Share of not-yet-in-use placed in service in Q3'26 (rest Q4'26)",0.5,"%","Estimated","Q1 letter: 'significantly increase our capacity in place in 3Q'")
inp("gpu","GPU/server share of capex and of not-yet-in-use",0.79,"%","Estimated (mgmt via secondary)","Mgmt on Q4'25 call: ~79% GPUs, ~20% data center, <1% power (secondary transcript summary: https://www.youtube.com/watch?v=JDCUSh2okgc) - verify vs transcript")
inp("lag","Capex-to-in-service lag",2,"quarters","Estimated","Assumption; Q2 not-yet-in-use 7.8B vs Q2 capex 5.7B suggests >1 quarter")
sec("Capex (consensus / guidance / extrapolation)")
inp("cx26","FY2026 capex",23207,"$M","Estimated (consensus)",MS+" ; guide $20-25B ("+SHL+")")
inp("h1cx","1H2026 capex actual",8130.3,"$M","Reported",FB)
inp("cx27","FY2027 capex",29135,"$M","Estimated (consensus)",MS)
inp("cx28","FY2028 capex",38885,"$M","Estimated (consensus)",MS)
inp("g29","FY2029 capex growth vs FY28",0.0,"%","Estimated (extrapolation)","No consensus beyond FY28; flat ~$38.9B ~ 1.3GW/yr at ~$30M/MW, consistent with '>1GW/yr from 2027' guide")
inp("g30","FY2030 capex growth vs FY29",0.0,"%","Estimated (extrapolation)","as above")
inp("cpmw","All-in capex per MW",30.0,"$M/MW","Estimated (third party)","~$30M/MW sell-side/SA estimate; only used to translate yields into $/MW")
sec("EBITDA / revenue (FY26-FY28 consensus; FY29+ extrapolated) - quarterly FY26-27 consensus is entered on Model sheet (blue)")
inp("eb28","FY2028 adj. EBITDA consensus",13978,"$M","Estimated (consensus)",MS)
inp("rv28","FY2028 revenue consensus",22983,"$M","Estimated (consensus)",MS)
for i,w in enumerate([0.22,0.24,0.26,0.28]): inp(f"w{i}",f"FY28 quarterly weight Q{i+1}",w,"%","Estimated","Assumed ramp (sums to 100%)")
inp("yov","FY29+ EBITDA yield override (annual EBITDA / gross GPU in service); blank = consensus-implied",None,"%","Scenario input","Blank uses FY28 consensus-implied yield (Model!C52). Reference points: installed base $12M/MW ~30%; Q2 deals $20M/MW ~51% (see Reverse_DCF notes)")
sec("1H2026 actuals (for FY2026 roll-up)")
inp("h1rv","1H26 revenue",981.3,"$M","Reported",FB)
inp("h1eb","1H26 adj. EBITDA",365.7,"$M","Reported",FB)
inp("h1da","1H26 D&A (at 5-yr policy)",471.7,"$M","Reported",FB)
inp("h1ben","1H26 D&A reduction from 4->5 yr change",86.1,"$M","Reported",FB+" (Q2 -$43.0M; FY26 est. -$167.6M)")
sec("Financing / cash interest")
tr=[("2.00% 2029 (remaining)",87.5,0.02),("3.00% 2031 (remaining)",87.5,0.03),("1.00% 9/2030",1581.25,0.01),("2.75% 9/2032",1581.25,0.0275),("1.25% 3/2031",2587.5,0.0125),("2.625% 3/2033",1750,0.02625),("0.50% 2/2030",3450,0.005),("4.50% 2/2034",2300,0.045)]
rows=[]
for n,p,c in tr:
    A.append([f"Convert {n}: principal / coupon",p,"$M",c,"Reported ("+FB+" Converts)"]); r=A.max_row
    A.cell(r,2).font=BLUE; A.cell(r,2).fill=INF; A.cell(r,4).font=BLUE; A.cell(r,4).fill=INF; rows.append(r)
inp("sofr","Term SOFR assumption for $775M secured loan (SOFR+2.50%)",0.04,"%","Estimated","Assumption - verify; loan assumed fully drawn (drawn amount not disclosed)")
inp("loan","Secured loan principal",775,"$M","Reported",FB)
inp("int_q","Cash interest per quarter (coupons + loan)","=("+"+".join(f"B{r}*D{r}" for r in rows)+f"+{C['loan'].split('!')[1]}*({C['sofr'].split('!')[1]}+0.025))/4","$M","Derived","sum(principal x coupon)/4; excludes accretion (non-cash) and capitalized-interest split",formula=True)
inp("rate","Cost of incremental funding beyond pro forma cash",0.07,"%","Estimated","Between NBIS 5.1% wtd effective and CRWV 8.4% (team comps.xlsx)")
inp("cash0","Pro forma cash (start)",14497.1,"$M","Estimated","01_priced_in: 8,042.1 + 775 + ~5,680; EXCLUDES Q3 burn - the model charges Q3'26 FCF against it")
sec("Impairment toggle")
inp("imp","Impairment of 6/30/26 in-service server base (% of NBV)",0.0,"%","Scenario input","0% default; e.g. 30% tests a Hopper write-down")
inp("impq","Impairment quarter",'Q4 2027',"label","Scenario input","Must match a Model quarter label")
inp("impi","Impairment quarter index","=MATCH("+C['impq'].split('!')[1]+",Model!$C$1:$T$1,0)-1","idx","Derived","",formula=True)
sec("Reverse DCF")
inp("wacc","WACC",0.11,"%","Estimated","Assumption (sensitivity in md); not a sourced figure")
inp("tg","Terminal growth",0.03,"%","Estimated","Assumption")
inp("tax","Tax rate (terminal only; NOLs assumed shield interim)",0.258,"%","Reported (statutory)","Netherlands headline CIT 25.8%")
inp("ev","Market EV (base, pro forma ex-leases)",69620,"$M","Derived","01_priced_in.xlsx (corrected 2026-10-04)")
inp("nc","Pro forma net cash",226.4,"$M","Derived","14,497.1 - 14,270.7")
inp("sh","Pro forma basic shares",287.66,"M","Derived","271.86 + 15.8 (01_priced_in)")
inp("px","Share price 10/2/26",242.81,"$","Reported","https://stockanalysis.com/stocks/nbis/history/")
for r in A.iter_rows(min_row=1,max_row=1):
    for c in r: c.font=Font(bold=True,color="FFFFFF"); c.fill=HDR
for k,w in zip("ABCDE",[62,14,9,22,110]): A.column_dimensions[k].width=w
def a(k): return C[k]

# ---------------- Model
M=wb.create_sheet("Model")
Q=[f"Q{q} {y}" for y in range(2026,2031) for q in range(1,5)][2:]; N=len(Q)
cols=[CL(3+i) for i in range(N)]; F0,FL=cols[0],cols[-1]
def rowset(r,label,unit,vals,fmt='#,##0;(#,##0)',inputs=None,bold=False):
    M.cell(r,1,label).font=B if bold else BLK; M.cell(r,2,unit)
    for i,v in enumerate(vals):
        c=M.cell(r,3+i,v); c.number_format=fmt
        if inputs and inputs[i]: c.font=BLUE; c.fill=INF
M.cell(1,1,"Quarter").font=B
for i,q in enumerate(Q): M.cell(1,3+i,q).font=B
rowset(2,"Index","",list(range(N)),'0',inputs=None)
rowset(3,"Fiscal year","",[int(q[-4:]) for q in Q],'0')
rowset(4,"GPU generation placed in service (label)","",[f'=IF({c}2<2,"Blackwell (B200/B300/GB300)",IF({c}2<12,"Vera Rubin","Rubin Ultra / next"))' for c in cols],'@')
cons_rv=[904.1,1452,2175,2646,3304,4052]; cons_eb=[354.7,670.7,1277,1581,1883,2372]
w=[a('w0'),a('w1'),a('w2'),a('w3')]
rv=[];eb=[]
for i,c in enumerate(cols):
    if i<6: rv.append(cons_rv[i]); eb.append(cons_eb[i])
    elif i<10: rv.append(f"={a('rv28')}*{w[i-6]}"); eb.append(f"={a('eb28')}*{w[i-6]}")
    else:
        p=cols[i-1]; eb.append(f'=IF({a("yov")}="",$C$52,{a("yov")})/4*{p}16'); rv.append(f"={c}7/({a('eb28')}/{a('rv28')})")
rowset(6,"Revenue","$M",rv,inputs=[i<6 for i in range(N)])
rowset(7,"Adj. EBITDA","$M",eb,inputs=[i<6 for i in range(N)],bold=True)
rowset(8,"Adj. EBITDA margin","%",[f"={c}7/{c}6" for c in cols],'0.0%')
M.cell(9,1,"Q3'26-Q4'27 = MarketScreener quarterly consensus (blue); FY28 = annual consensus x weights; FY29+ = EBITDA yield x prior-qtr gross GPU in service")
cx=[]
for i,c in enumerate(cols):
    if i<2: cx.append(f"=({a('cx26')}-{a('h1cx')})/2")
    elif i<6: cx.append(f"={a('cx27')}/4")
    elif i<10: cx.append(f"={a('cx28')}/4")
    elif i<14: cx.append(f"={a('cx28')}*(1+{a('g29')})/4")
    else: cx.append(f"={a('cx28')}*(1+{a('g29')})*(1+{a('g30')})/4")
rowset(10,"Capex (total)","$M",cx,bold=True)
rowset(11,"  GPU/server capex","$M",[f"={c}10*{a('gpu')}" for c in cols])
rowset(12,"  Non-GPU capex (DC/infra/power)","$M",[f"={c}10*(1-{a('gpu')})" for c in cols])
def placed(c,share):
    return f"=(IF({c}$2=0,{a('niu')}*{a('niu3')},0)+IF({c}$2=1,{a('niu')}*(1-{a('niu3')}),0)+IF({c}$2>={a('lag')},INDEX(${F0}$10:${FL}$10,{c}$2-{a('lag')}+1),0))*{share}"
rowset(14,"GPU placed in service (cohort = column)","$M",[placed(c,a('gpu')) for c in cols])
rowset(15,"Non-GPU placed in service","$M",[placed(c,"(1-"+a('gpu')+")") for c in cols])
rowset(16,"Cumulative gross GPU/server in service (incl. 6/30 base, no retirements)","$M",[f"={a('srv')}+SUM(${F0}$14:{c}14)" for c in cols])
rowset(17,"Construction-in-progress memo: cum. capex not yet in service (from Q3'26)","$M",[f"=SUM(${F0}$10:{c}10)+{a('niu')}-SUM(${F0}$14:{c}14)-SUM(${F0}$15:{c}15)" for c in cols])
blocks={}
r0=19
for k,Lk in enumerate(["L1","L2","L3"]):
    L=a(Lk); r=r0+k*10; blocks[Lk]=r
    M.cell(r,1,f"DEPRECIATION @ life scenario {'ABC'[k]}").font=B; M.cell(r,2,f"={L}")
    rowset(r+1,"  Existing 6/30 server base","$M",[f"=IF({c}$2<4*{L}-{a('age')},{a('srv')}/(4*{L})*IF({c}$2>={a('impi')},1-{a('imp')},1),0)" for c in cols])
    rowset(r+2,"  New GPU cohorts (half-quarter convention)","$M",[f"=(SUMPRODUCT(${F0}$14:${FL}$14,--(${F0}$2:${FL}$2<{c}$2),--(${F0}$2:${FL}$2>{c}$2-4*{L}))+0.5*{c}$14+0.5*SUMPRODUCT(${F0}$14:${FL}$14,--(${F0}$2:${FL}$2={c}$2-4*{L})))/(4*{L})" for c in cols])
    rowset(r+3,"  Non-GPU (existing + new)","$M",[f"={a('oth')}/(4*{a('ngl')})+(SUMPRODUCT(${F0}$15:${FL}$15,--(${F0}$2:${FL}$2<{c}$2))+0.5*{c}$15)/(4*{a('ngl')})" for c in cols])
    rowset(r+4,"  Impairment charge","$M",[f"=IF({c}$2={a('impi')},{a('imp')}*{a('srv')}*MAX(0,4*{L}-{a('age')}-{c}$2)/(4*{L}),0)" for c in cols])
    rowset(r+5,"  Total D&A","$M",[f"={c}{r+1}+{c}{r+2}+{c}{r+3}" for c in cols],bold=True)
    rowset(r+6,"  EBIT (adj.; pre-SBC) = EBITDA - D&A - impairment","$M",[f"={c}$7-{c}{r+5}-{c}{r+4}" for c in cols],bold=True)
    rowset(r+7,"  EBIT margin","%",[f"={c}{r+6}/{c}$6" for c in cols],'0.0%')
    rowset(r+8,"  D&A / revenue","%",[f"={c}{r+5}/{c}$6" for c in cols],'0.0%')
rowset(50,"Cash interest (coupons + loan)","$M",[f"={a('int_q')}" for c in cols])
inc=[f"=MAX(0,-{a('cash0')})*{a('rate')}/4"]+[f"=MAX(0,-{cols[i-1]}55-{a('cash0')})*{a('rate')}/4" for i in range(1,N)]
rowset(51,"Interest on funding beyond pro forma cash","$M",inc)
M.cell(52,1,"FY28 consensus-implied EBITDA yield on avg gross GPU in service").font=B
M.cell(52,3,f"={a('eb28')}/AVERAGE(I16:L16)").number_format='0.0%'
rowset(54,"FCF = EBITDA - capex - cash interest (pre-WC/prepayments, pre-tax; independent of book life)","$M",[f"={c}7-{c}10-{c}50-{c}51" for c in cols],bold=True)
rowset(55,"Cumulative FCF from Q3'26","$M",[f"=SUM(${F0}$54:{c}54)" for c in cols])
rowset(56,"External funding need beyond pro forma cash","$M",[f"=MAX(0,-{c}55-{a('cash0')})" for c in cols])
rowset(57,"Discount factor (from ~Oct-26)","x",[f"=(1+{a('wacc')})^(-{c}$2/4)" for c in cols],'0.000')
M.column_dimensions['A'].width=70; M.column_dimensions['B'].width=8
for c in cols: M.column_dimensions[c].width=11
M.freeze_panes="C2"

# ---------------- Annual
An=wb.create_sheet("Annual"); yrs=[2026,2027,2028,2029,2030]
An.append(["$M","Basis"]+[f"FY{y}" for y in yrs])
def S(r): return f"SUMIF(Model!$C$3:$T$3,{{y}},Model!$C${r}:$T${r})"
def ann(label,basis,r,h1=None):
    row=[label,basis]
    for y in yrs:
        f="="+S(r).format(y=y)
        if y==2026 and h1: f+="+"+h1
        row.append(f)
    An.append(row)
ann("Revenue","FY26 incl. 1H actual",6,a('h1rv')); ann("Adj. EBITDA","FY26 incl. 1H actual",7,a('h1eb'))
ann("Capex","FY26 incl. 1H actual",10,a('h1cx'))
for k,Lk in enumerate(["L1","L2","L3"]):
    r=blocks[Lk]; L=a(Lk)
    h1da=f"({a('h1da')}-4*{a('h1ben')}+4*{a('h1ben')}*5/{L})"
    ann(f"D&A @ {['4','5','6'][k]}-yr","FY26 1H restated to life (D)",r+5,h1da)
    ann(f"Impairment @ {['4','5','6'][k]}-yr","",r+4)
    An.append([f"EBIT @ {['4','5','6'][k]}-yr","EBITDA less D&A less impairment"]+[f"={CL(3+i)}3-{CL(3+i)}{An.max_row-1}-{CL(3+i)}{An.max_row}" for i in range(5)])
    An.append([f"EBIT margin @ {['4','5','6'][k]}-yr",""]+[f"={CL(3+i)}{An.max_row}/{CL(3+i)}2" for i in range(5)])
ann("Cash interest","H2'26 only",50); ann("Interest on extra funding","",51)
ann("FCF (EBITDA-capex-interest)","FY26 = H2 only",54)
An.append(["External funding need, year-end (cumulative)",""]+[f"=INDEX(Model!$C$56:$T$56,MATCH(\"Q4 {y}\",Model!$C$1:$T$1,0))" for y in yrs])
An.append(["Memo: consensus FCF (MarketScreener)","",-17500,-23148,-25500,"n/a","n/a"])
for c in An[An.max_row][2:5]: c.font=BLUE
An.append(["Memo: 1H26 D&A at life L = 471.7 - 344.4 + 344.4x5/L, where 344.4 = 4 x 86.1 (since D4 - D5 = D5/4). FY26 company est. benefit: $167.6M on YE25 assets."])
for row in An.iter_rows(min_row=2):
    for c in row[2:]: c.number_format='#,##0;(#,##0)'
for r in (9,13,17):
    for c in An[r][2:]: c.number_format='0.0%'
An.column_dimensions['A'].width=44; An.column_dimensions['B'].width=30

# ---------------- Horizons
H=wb.create_sheet("Horizons")
H.append(["Mark","Quarter shown","Run-rate (qtr x4) $M:","EBITDA","D&A 4y","D&A 5y","D&A 6y","EBIT 4y","EBIT 5y","EBIT 6y","FCF","Cum. external funding need"])
for lab,q in [("6-month (Apr-2027)","Q1 2027"),("1-year (Oct-2027)","Q3 2027"),("3-year (Oct-2029)","Q3 2029")]:
    m=f"MATCH($B{H.max_row+1},Model!$C$1:$T$1,0)"
    def g(r,mult=4): return f"=INDEX(Model!$C${r}:$T${r},{m})*{mult}"
    H.append([lab,q,"",g(7),g(blocks['L1']+5),g(blocks['L2']+5),g(blocks['L3']+5),g(blocks['L1']+6),g(blocks['L2']+6),g(blocks['L3']+6),g(54),g(56,1)])
    H.cell(H.max_row,2).font=BLUE
for row in H.iter_rows(min_row=2):
    for c in row[3:]: c.number_format='#,##0;(#,##0)'
H.column_dimensions['A'].width=22

# ---------------- Reverse DCF
R=wb.create_sheet("Reverse_DCF")
R.append(["Item","4-yr","5-yr","6-yr","Note"])
lr={"L1":"B","L2":"C","L3":"D"}
def rr(label,f,note="",fmt='#,##0;(#,##0)'):
    R.append([label]+[f(Lk) for Lk in ["L1","L2","L3"]]+[note]); 
    for c in R[R.max_row][1:4]: c.number_format=fmt
    return R.max_row
r_pv=rr("PV of FCF Q3'26-Q4'30",lambda L:"=SUMPRODUCT(Model!$C$54:$T$54,Model!$C$57:$T$57)","Same for all lives (book life does not change cash)")
r_e=rr("FY2030 EBITDA",lambda L:"=Annual!G3")
r_m=rr("Steady-state replacement capex = annualized Q4'30 D&A at that life",lambda L:f"=4*(Model!$T${blocks[L]+1}+Model!$T${blocks[L]+2}+Model!$T${blocks[L]+3})","Economic life = book life assumption")
r_g=rr("Gross GPU in service within life (Q4'30)",lambda L:f"=4*{a(L)}*(Model!$T${blocks[L]+1}+Model!$T${blocks[L]+2})")
r_f=rr("Terminal FCF (yr 2031)",lambda L:f"=({lr[L]}{r_e}*(1+{a('tg')})-{lr[L]}{r_m}*(1+{a('tg')}))*(1-{a('tax')})-{a('tg')}*{lr[L]}{r_g}","(1-t)(EBITDA - replacement) - g x gross (growth reinvestment)")
r_tv=rr("Terminal value (end-2030)",lambda L:f"={lr[L]}{r_f}/({a('wacc')}-{a('tg')})")
r_pt=rr("PV of terminal value",lambda L:f"={lr[L]}{r_tv}*Model!$T$57")
r_ev=rr("Model EV",lambda L:f"={lr[L]}{r_pv}+{lr[L]}{r_pt}")
r_ps=rr("Model value per share",lambda L:f"=({lr[L]}{r_ev}+{a('nc')})/{a('sh')}","vs $242.81",'$#,##0;($#,##0)')
rr("Upside/(downside) vs price",lambda L:f"={lr[L]}{r_ps}/{a('px')}-1","",'0%')
r_er=rr("REQUIRED FY2030 EBITDA to justify market EV",lambda L:f"=(({a('ev')}-{lr[L]}{r_pv})/Model!$T$57*({a('wacc')}-{a('tg')})+{a('tg')}*{lr[L]}{r_g})/((1-{a('tax')})*(1+{a('tg')}))+{lr[L]}{r_m}","Holds interim FCF fixed (conservative: higher EBITDA would also lift interim FCF)")
r_yq=rr("Required EBITDA yield on gross GPU in service",lambda L:f"={lr[L]}{r_er}/{lr[L]}{r_g}","",'0.0%')
r_ya=rr("Modeled FY2030 EBITDA yield",lambda L:f"={lr[L]}{r_e}/{lr[L]}{r_g}","",'0.0%')
rr("Required revenue per MW (yield x capex/MW x GPU share / FY28 margin)",lambda L:f"={lr[L]}{r_yq}*{a('cpmw')}*{a('gpu')}/({a('eb28')}/{a('rv28')})","Compare: installed base ~$12M/MW; Q2 deals $20-25M/MW; Q3 short-term >$40M/MW (Q2 letter)",'$0.0"M"')
rr("Implied terminal EV / FY2030 EBITDA",lambda L:f"={lr[L]}{r_tv}/{lr[L]}{r_e}","",'0.0x')
R.append([]); R.append(["Reference yields (Derived): EBITDA yield = $/MW x EBITDA margin / (capex per MW x GPU share). Installed base $12M x 60.8% / ($30M x 0.79) = 30.8%; Q2 deals $20M => 51.3%; $25M => 64.1%."])
R.column_dimensions['A'].width=62; R.column_dimensions['E'].width=70
for k in "BCD": R.column_dimensions[k].width=14

# ---------------- Peers
P=wb.create_sheet("Peers")
P.append(["Company","Asset class","Book useful life","Recent change","Source","Tag"])
for r in [
 ("Nebius","Server & network equipment","5 yrs (from 1/1/26)","4 -> 5 yrs; -$43.0M Q2'26 D&A, ~-$167.6M FY26 est.",Q2FS+" ; "+FB,"Reported"),
 ("Nebius (prior)","Server & network equipment","4 yrs","'Our main assets, GPUs, are currently depreciated over four years' (20-F p.56)",F20,"Reported"),
 ("CoreWeave","Technology equipment (incl. GPUs)","6 yrs","5 -> 6 yrs from 1/1/2023","https://gpuinsights.net/server-useful-life-gpu-depreciation-2026/ (summarizing CRWV FY25 10-K)","Reported (via secondary)"),
 ("Microsoft","Servers & network equipment","2-6 yrs","4 -> 6 yrs from FY2023 (+$3.7B op. income FY23)","https://www.yieldtheory.app/research/michael-burry-ai-depreciation ; MSFT FY26 10-K","Reported (via secondary)"),
 ("Alphabet","Servers & some network equipment","6 yrs","to 6 yrs from 2023 (-$3.9B depreciation 2023)","https://gpuinsights.net/server-useful-life-gpu-depreciation-2026/","Reported (via secondary)"),
 ("Meta","Servers & network assets","5-5.5 yrs","most to 5.5 yrs from Jan-2025 (~-$2.9B 2025 depreciation)","https://www.yieldtheory.app/research/michael-burry-ai-depreciation","Reported (via secondary)"),
 ("Amazon","Servers; subset of servers/networking","6 yrs; subset 5 yrs","5->6 (2024) then subset 6->5 (2025) citing AI/ML pace (+$1.4B depreciation)","https://www.cnbc.com/2025/11/14/ai-gpu-depreciation-coreweave-nvidia-michael-burry.html","Reported (via secondary)"),
 ("Oracle","Servers & network equipment","6 yrs","5 -> 6 yrs from FY2025 (-$733M opex)","https://www.yieldtheory.app/research/michael-burry-ai-depreciation","Reported (via secondary)"),
 ("IREN","GPUs","not found","FY26 10-K has a useful-life schedule; GPU figure not extracted. Some GPU financings use 36-month terms","https://bizety.com/2025/09/23/gpu-depreciation-coreweave-vs-nebius/","not disclosed/not found"),
 ("Lambda","GPUs","not found (private)","","","not disclosed/not found")]:
    P.append(list(r))
P.append([]); P.append(["Note: secondary summaries used for peers; primary 10-K pages not re-read in this pass - verify before publishing."])
for k,wd in zip("ABCDEF",[16,34,20,60,80,22]): P.column_dimensions[k].width=wd

for ws in [M,An,H,R,P]:
    for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=HDR
from openpyxl.drawing.image import Image as XI
CH=wb.create_sheet("Charts"); CH["A1"]="PNG charts generated by dep/charts.py (python mirror of this workbook; values cross-checked vs LibreOffice recalculation)"
for i,f in enumerate(["03a_ebit_by_life_quarterly","03b_annual_ebitda_ebit_fcf","03c_reverse_dcf_yield","03d_da_by_cohort_5yr"]):
    im=XI(f"/workspace/nebius/skeptic/charts/{f}.png"); im.width,im.height=880,420; CH.add_image(im,f"A{3+i*23}")
L=wb.create_sheet("Legend",0)
for t in ["03_depreciation.xlsx - Nebius Skeptic (Red Team). Built 2026-10-04 ET.","Blue font/shaded = hard-coded input (sourced or assumption, see Assumptions col D/E). Black = formula.","Cohort/vintage method: each quarter's GPU placed-in-service amount is a cohort depreciated straight-line over 4/5/6 yrs (half-quarter convention). 6/30/26 in-service base depreciated over remaining life (life - avg age).","EBITDA is consensus through FY28 (MarketScreener), then extrapolated as FY28-implied EBITDA yield x gross GPU in service. Capex FY26-28 consensus, flat thereafter (input).","FCF ignores working capital/prepayments (Q2'26 OCF +$2.25B was prepayment-driven), taxes (NOLs), SBC and lease payments. Book life changes EBIT, not cash; it matters for value through the replacement capex needed at the true economic life (Reverse_DCF).","Sheets: Assumptions, Model (quarterly Q3'26-Q4'30), Annual, Horizons (6m/1y/3y), Reverse_DCF, Peers, Charts."]:
    L.append([t])
L.column_dimensions['A'].width=160
wb.save("/workspace/nebius/skeptic/03_depreciation.xlsx"); print("saved")
