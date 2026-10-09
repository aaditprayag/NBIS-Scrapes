# Builds /workspace/nebius/filings/nebius_kpi_history.xlsx  (Filings & Capital workstream)
import json, openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter as L

OUT='/workspace/nebius/filings/nebius_kpi_history.xlsx'
Q=['Q1 2024','Q2 2024','Q3 2024','Q4 2024','Q1 2025','Q2 2025','Q3 2025','Q4 2025','Q1 2026','Q2 2026']
DAYS=[91,91,92,92,90,91,92,92,90,91]
C0=3  # first quarter column (C)
BLUE=Font(color='0000FF'); BLACK=Font(color='000000'); ORANGE=Font(color='C65911',italic=True); GREY=Font(color='808080')
BOLD=Font(bold=True); TITLE=Font(bold=True,size=13); HDRF=PatternFill('solid',fgColor='DDDDDD'); SECF=PatternFill('solid',fgColor='EEF3FA')
ND='n.d.'
E='https://www.sec.gov/Archives/edgar/data/1513845/'

qr=json.load(open('/workspace/nebius/filings/quarterly_rows.json'))
S=qr['S']; ROWS=qr['rows']
S.update({
 'H1FS25':S.get('H1FS25',('6-K','2025-09-10','0001104659-25-088860','tm2525580d2_ex99-1.htm','H1 2025 interim OFR')),
 'Q1FS':S['Q1FS'],'Q2FS':S['Q2FS'],
})
def url(k):
    f,d,a,doc,desc=S[k]; return E+a.replace('-','')+'/'+doc
def cite(k,ref=''):
    f,d,a,doc,desc=S[k]; return f"{f} {d} acc {a} {doc}{(' '+ref) if ref else ''} | {url(k)}"
CAP='/workspace/nebius/capacity/nebius_capacity.xlsx'
MKT='/workspace/nebius/market/gpu_price_history.xlsx'

wb=Workbook(); wb.remove(wb.active)
readme=wb.create_sheet('README'); ws=wb.create_sheet('KPI'); pm=wb.create_sheet('Prices_Monthly'); chk=wb.create_sheet('Checks'); srcs=wb.create_sheet('Sources')
SRC=[]  # (sheet, cell, metric, period, value/formula, tag, source, note)

def style(c,tag):
    c.font={'Reported':BLUE,'Derived':BLACK,'Estimated':ORANGE,'not disclosed':GREY}.get(tag,BLACK)

# ---------------- Prices_Monthly (market team) ----------------
mw=openpyxl.load_workbook(MKT,data_only=True)
obs=mw['Price_History_Monthly']; ff=mw['Monthly_FFill']
oh=[c.value for c in obs[1]]; fh=[c.value for c in ff[1]]
months=[f'{y}-{m:02d}' for y in (2024,2025,2026) for m in range(1,13)][:30]
SKU=[('h100','Nebius | H100 SXM 80GB | on-demand | PAYG','H100 SXM on-demand'),('h200','Nebius | H200 141GB | on-demand | PAYG','H200 on-demand'),
     ('b200','Nebius | B200 | on-demand | PAYG','B200 on-demand'),('b300','Nebius | B300 | on-demand | PAYG','B300 on-demand')]
orow={r[0]:r for r in obs.iter_rows(min_row=2,values_only=True)}; frow={r[0]:r for r in ff.iter_rows(min_row=2,values_only=True)}
pm['A1']='Nebius public on-demand list prices, $/GPU-hr, monthly (GPU Market & Comps team: gpu_price_history.xlsx). Blue = observed in a Wayback/live snapshot that month (Reported); orange italic = carried forward from the last observation (Estimated, forward-fill); grey = no observation yet.'; pm['A1'].font=TITLE
pm.cell(3,1,'SKU').font=BOLD
for j,mo in enumerate(months): pm.cell(3,2+j,mo).font=BOLD
PMROW={}
for i,(k,col,lab) in enumerate(SKU):
    r=4+i; PMROW[k]=r; pm.cell(r,1,'Nebius '+lab)
    for j,mo in enumerate(months):
        o=orow[mo][oh.index(col)]; f=frow[mo][fh.index(col)]
        c=pm.cell(r,2+j)
        if o is not None: c.value=o; tag='Reported'; src=f"{MKT} Price_History_Monthly '{col}' {mo} (snapshot rows in Price_History_Long; nebius.com/prices via web.archive.org)"; note=''
        elif f is not None: c.value=f; tag='Estimated'; src=f"{MKT} Monthly_FFill '{col}' {mo}"; note='no snapshot this month; carried forward from last observation (market-team forward-fill)'
        else: c.value=ND; tag='not disclosed'; src=f"{MKT} (no observation)"; note='SKU not listed / no snapshot'
        style(c,tag); SRC.append(('Prices_Monthly',c.coordinate,'Nebius '+lab,mo,c.value,tag,src,note))
pm.column_dimensions['A'].width=28

# ---------------- KPI grid ----------------
ws['A1']='Nebius Group (NBIS) - consolidated tagged KPI history, Q1 2024 - Q2 2026. $M unless noted. Blue = Reported input; black = formula on Reported/Derived inputs; orange italic = Estimated (input, or formula using an Estimated input); grey = n.d. Per-cell sources in Sources tab.'
ws['A1'].font=TITLE
ws.cell(3,1,'Metric').font=BOLD; ws.cell(3,2,'Unit').font=BOLD
for i,q in enumerate(Q): ws.cell(3,C0+i,q).font=BOLD
ws.cell(3,C0+10,'Basis / caveat').font=BOLD; ws.cell(3,C0+11,'Owner / origin').font=BOLD
for c in range(1,C0+12): ws.cell(3,c).fill=HDRF
RIDX={}; TAGS={}; r=[4]
def sec(t):
    c=ws.cell(r[0],1,t); c.font=BOLD
    for cc in range(1,C0+12): ws.cell(r[0],cc).fill=SECF
    r[0]+=1
def newrow(key,label,unit,basis='',owner=''):
    RIDX[key]=r[0]; TAGS[key]=[None]*10
    ws.cell(r[0],1,label); ws.cell(r[0],2,unit); ws.cell(r[0],C0+10,basis); ws.cell(r[0],C0+11,owner)
    r[0]+=1; return RIDX[key]
def setcell(key,i,val,tag,src,note=''):
    row=RIDX[key]; c=ws.cell(row,C0+i)
    if tag=='not disclosed': val=ND if val in (None,'',ND) else val
    c.value=val; style(c,tag); TAGS[key][i]=tag
    SRC.append(('KPI',c.coordinate,ws.cell(row,1).value,Q[i],val,tag,src,note))
def ref(key,i): return f'{L(C0+i)}{RIDX[key]}'

# --- A. carried from our Quarterly tab
QMAP=[('grp','Revenue - group (continuing ops, latest restated basis, $M)','Group revenue (continuing ops, restated)','$M'),
      ('ai','Segment revenue - Nebius AI cloud ($M)','Nebius AI cloud segment revenue','$M'),
      ('arr','ARR - Nebius AI cloud, as stated ($M)','ARR (as stated; last-month revenue x12)','$M'),
      ('ebitda','Adjusted EBITDA - group ($M)','Adjusted EBITDA - group','$M'),
      ('ebitda_ai','Adjusted EBITDA - Nebius AI cloud segment ($M)','Adjusted EBITDA - Nebius AI cloud segment','$M'),
      ('da','Depreciation & amortization ($M)','Depreciation & amortization','$M'),
      ('capex','Capex - purchases of PP&E and intangibles ($M)','Capex - cash purchases of PP&E and intangibles','$M'),
      ('ppe','PP&E, net (end of period, $M)','PP&E, net (period end)','$M'),
      ('cash','Cash & cash equivalents (end of period, $M)','Cash & cash equivalents (period end)','$M'),
      ('debt','Total debt (current + non-current, carrying, $M)','Total debt, carrying (period end)','$M'),
      ('defrev','Deferred revenue, total (end of period, $M)','Deferred revenue, total (period end)','$M'),
      ('rpo','RPO (end of period, $M)','Remaining performance obligations (period end)','$M'),
      ('sh','Shares outstanding - total (period end)','Shares issued & outstanding (period end; excl. NVIDIA pre-funded warrant)','shares'),
      ('wavg','Weighted avg shares - basic','Weighted-average shares - basic','shares'),
      ('conc','Customer concentration (>10% of revenue, quarter)','Customer concentration (>10% of revenue)','text')]
sec('A. Financials (Filings & Capital: nebius_filings.xlsx Quarterly tab; SEC filings)')
for key,qk,label,unit in QMAP:
    basis={'arr':'Pre-Q1-26 letters: "revenue from the last month of the quarter x12"; Q1-26 on: "Nebius AI cloud revenue from the last month of the quarter multiplied by 12". 2024 values from Oct-24 investor presentation chart.',
           'ai':'Segment revenue; Q1/Q2 2024 split not disclosed (H1-24 total 17.4).','capex':'Cash flow statement, continuing ops.','sh':'Excludes 21,065,936 NVIDIA pre-funded warrant shares (unexercised).'}.get(key,'')
    newrow(key,label,unit,basis,'Filings & Capital')
    for i,(v,tag,k,rf,note) in enumerate(ROWS[qk]):
        if tag=='Not disclosed':
            setcell(key,i,ND,'not disclosed',cite(k,rf) if k else 'n/a',(v if v!=qr['ND'] else '')+((' ; '+note) if note else ''))
        else:
            setcell(key,i,v,tag if tag in('Reported','Derived','Estimated') else 'Reported',cite(k,rf),note)
# --- new rows: OCF, change in deferred revenue, PP&E detail
newrow('ocf','Net cash from operating activities - continuing ops','$M','Quarterly; latest restated basis (Toloka in discontinued ops).','Filings & Capital (new)')
OCF=[(-63.0,'Derived','H1FS25','Cash flow summary: 1H24 (162.0)','Q1-24 = 1H24 (162.0) - Q2-24 (99.0) [Q2PR25 highlights]; FY24 check: -63.0-99.0-34.8-73.1 = -269.9 (20F25)'),
     (-99.0,'Reported','Q2PR25','Highlights table (Q2-24 comparative)',''),
     (-34.8,'Reported','Q3PR25','Cash flow statement, three months (Q3-24 comparative)','Q3-25 PR highlights table shows (34.9)'),
     (-73.1,'Reported','Q4PR','Cash flow statement, three months (Q4-24 comparative)',''),
     (-184.1,'Reported','Q1PR','Cash flow statement (Q1-25 comparative, restated)','Q1-25 PR originally (197.8) incl. Toloka'),
     (-167.7,'Reported','Q2PR25','Highlights table','Q2-26 PR cash flow statement shows Q2-25 as (167.9); 1H25 (352.0) - Q1-25 (184.1) = (167.9)'),
     (-80.6,'Reported','Q3PR25','Cash flow statement, three months',''),
     (834.3,'Reported','Q4PR','Cash flow statement, three months','FY25 401.9'),
     (2258.0,'Reported','Q1PR','Cash flow statement',''),
     (2246.1,'Reported','Q2PR','Cash flow statement, three months','1H26 4,504.1')]
for i,(v,t,k,rf,n) in enumerate(OCF): setcell('ocf',i,v,t,cite(k,rf),n)
newrow('ddr','Change in deferred revenue (cash flow statement, operating)','$M','Increase = cash collected ahead of revenue (customer prepayments).','Filings & Capital (new)')
DDR=[(ND,'not disclosed','Q3PR25','','quarterly split not disclosed; 1H24 = 9M24 10.4 - Q3-24 3.7 = 6.7 (Derived)'),
     (ND,'not disclosed','Q3PR25','','see Q1 2024 note (1H24 6.7)'),
     (3.7,'Reported','Q3PR25','Cash flow statement, three months (Q3-24 comparative)',''),
     (-0.4,'Reported','Q4PR','Cash flow statement, three months (Q4-24 comparative)','FY24 10.0'),
     (2.4,'Reported','Q1PR','Cash flow statement (Q1-25 comparative)',''),
     (0.6,'Reported','Q2PR','Cash flow statement, three months (Q2-25 comparative)','1H25 3.0'),
     (-3.8,'Reported','Q3PR25','Cash flow statement, three months','9M25 (0.7)'),
     (1566.5,'Reported','Q4PR','Cash flow statement, three months','FY25 1,565.8'),
     (3198.0,'Reported','Q1PR','Cash flow statement',''),
     (1197.0,'Reported','Q2PR','Cash flow statement, three months','1H26 4,395.0')]
for i,(v,t,k,rf,n) in enumerate(DDR): setcell('ddr',i,v,t,cite(k,rf) if t!='not disclosed' else cite(k,'cash flow statement (9M24 10.4; Q3-24 3.7)'),n)
newrow('ocf_x','OCF excluding change in deferred revenue','$M','OCF minus change in deferred revenue (prepayment-free operating cash).','Derived')
PPE={'sng':('Server and network equipment, gross (in service)',[None,None,None,(794.9,'20F25','Note 7 p.F-38'),None,None,None,(3123.6,'20F25','Note 7 p.F-38'),(4692.7,'Q1FS','Note 7 p.F-19'),(5197.7,'Q2FS','Note 7 p.F-23')]),
     'aniu':('Assets not yet in use (gross; under installation incl. prepayments)',[None,None,None,(163.8,'20F25','Note 7 p.F-38'),None,None,None,(2417.4,'20F25','Note 7 p.F-38'),(2427.9,'Q1FS','Note 7 p.F-19'),(7836.4,'Q2FS','Note 7 p.F-23')]),
     'ppeg':('PP&E, gross (total before accumulated depreciation)',[None,None,None,(1083.0,'20F25','Note 7 p.F-38'),None,None,None,(6190.3,'20F25','Note 7 p.F-38'),(7976.6,'Q1FS','Note 7 p.F-19'),(14149.5,'Q2FS','Note 7 p.F-23')])}
for key,(lab,vals) in PPE.items():
    newrow(key,lab,'$M','PP&E note breakdown only in 20-F and 2026 interim statements; 2024/2025 interim OFRs give no breakdown.' if key=='sng' else ('FY24 20-F showed 163.9 for Dec-24 (later 163.8)' if key=='aniu' else ''),'Filings & Capital (new)')
    for i,x in enumerate(vals):
        if x is None: setcell(key,i,ND,'not disclosed','PP&E breakdown not in interim report for this quarter','')
        else: setcell(key,i,x[0],'Reported',cite(x[1],x[2]),'')
newrow('inserv','PP&E gross in service (= gross - assets not yet in use)','$M','','Derived')

# Transfer-into-service (filings-owned Derived from Note 7 consecutive quarter-ends)
newrow('d_aniu','Change in assets not yet in use (QoQ)','$M','Derived: period-end ANIU minus prior period-end ANIU from Note 7. Positive = net additions to installation pipeline (purchases/prepayments exceed transfers into service). Only where consecutive breakdowns exist (Q4-24→Q4-25 annual gap skipped; Q1-26 and Q2-26 QoQ).','Derived')
newrow('pis_srv','Servers placed in service (QoQ change in server & network equipment, gross)','$M','Derived: QoQ change in Server and network equipment, gross (Note 7). Proxy for servers transferred into service; does not equal -Δ ANIU because new purchases enter ANIU first. Available Q1-26 (vs YE25) and Q2-26 (vs Q1-26).','Derived')

# --- B. Capacity & Power
cw=openpyxl.load_workbook(CAP,data_only=True); cq=cw['Quarterly MW']; cg=cw['GPUs deployed']
sec('B. Capacity & Power (capacity/nebius_capacity.xlsx; Capacity & Power team - tags carried)')
MWR=[('mw','Energized MW - company total incl. plug (Capacity base case)',130,'Ties to mgmt figures where given (YE25 ~170 MW active). Capacity Quarterly MW row 130 (Total incl. plug). MIXED basis: Vineland, Modi\'in, AIB = IT MW; other sites as stated (mostly gross/unknown). Closest to Nebius "active" power; NOT connected or contracted.'),
     ('mw_lo','Energized MW - identified sites, low',124,'Excludes plug for undisclosed sites. Capacity Quarterly MW row 124.'),
     ('mw_hi','Energized MW - identified sites, high',125,'Excludes plug. Capacity Quarterly MW row 125.')]
for key,lab,row,basis in MWR:
    newrow(key,lab,'MW',basis,'Capacity & Power')
    for i in range(10):
        v=cq.cell(row,5+i).value
        setcell(key,i,v,'Estimated',f"{CAP} 'Quarterly MW'!{L(5+i)}{row} (Capacity model; sum of per-site estimates, site sources in Tags_Sources)",'Capacity tag on cell: Derived (sum) of Estimated site inputs -> carried as Estimated')
newrow('mw_mgmt','MW as stated by management','MW','YE24 = connected (Q3-25 letter chart); YE25 = ACTIVE (Q4-25 letter). Different definitions.','Capacity & Power / filings')
for i in range(10):
    if i==3: setcell('mw_mgmt',i,25,'Reported','6-K 2025-11-12 Q3-25 shareholder letter chart (connected MW YE24) | https://www.sec.gov/Archives/edgar/data/1513845/000110465925109806/tm2530882d1_ex99-2.htm','connected power')
    elif i==7: setcell('mw_mgmt',i,170,'Reported',cite('Q4LTR','highlights ("~170 MW of active power")'),'active power, approximate')
    else: setcell('mw_mgmt',i,ND,'not disclosed','n/a','no MW figure stated for this quarter end')
GEN=json.load(open('/tmp/gpucls.json'))
newrow('gpu','GPUs deployed - modeled (Capacity: energized MW x GPUs/MW)','GPUs','Modeled from MW, not a company count. Company counts ran ahead of powered MW.','Capacity & Power')
for i in range(10): setcell('gpu',i,cg.cell(34,5+i).value,'Estimated',f"{CAP} 'GPUs deployed'!{L(5+i)}34 (sum of site rows 11-29: energized base MW x GPUs/MW from gpu_mw_capex.xlsx)",'')
GL=[('g_hop','  of which Hopper (H100-class; Mantsala first 25 MW, Paris, Keflavik)','Hopper'),('g_h200','  of which H200 (Kansas City first 10 MW @ 548.5 GPUs/MW)','H200'),('g_b200','  of which B200 (Kansas City expansion, Modi\'in)','B200'),('g_b300','  of which B300 HGX (Longcross)','B300'),('g_gb300','  of which GB300 NVL72 (Vineland, Mantsala expansion)','GB300')]
for key,lab,g in GL:
    newrow(key,lab,'GPUs','Generation per site mostly ASSUMED by Capacity team (GPU vintage tab).','Capacity & Power (split by Filings)')
    for i in range(10): setcell(key,i,GEN[g][i],'Estimated',f"{CAP} 'GPUs deployed' site rows 11-29 by generation col B; mixed sites split at Mantsala 13,713 Hopper (25 MW x 548.5) and Kansas City 2,743 H200 (5 MW x 548.5)",'')
newrow('gpu_co','GPU count as stated by company (for comparison)','GPUs','Not used in formulas.','Capacity & Power / filings')
GCO={2:(10000,'Estimated',"Estimated/target - NOT a company actual. capacity 'GPUs deployed'!G31: site_list KPI pivot ~10k (from '>20k by YE24' / 'nearly doubled' commentary)"),3:(20000,'Estimated',"Estimated/target - NOT a company actual. capacity 'GPUs deployed'!H31: '>20,000 GPUs by YE2024' target (Q3-24 call, Oct-31-2024) - a forward-looking target"),4:(30000,'Reported',"20-F FY2024 (filed 2025-04-30, acc 0001558370-25-005991, nbis-20241231x20f.htm; local edgar/nbis-20241231x20f.htm.txt line ~571): 'As of March 31, 2025, we have deployed approximately 30,000 GPUs across our global data center footprint, primarily NVIDIA H200s.' | https://www.sec.gov/Archives/edgar/data/1513845/000155837025005991/nbis-20241231x20f.htm",)}
GCO_NOTE={4:"Supersedes the ~38,000 figure previously carried from the Feb-20-2025 Q4-24 call, which was a forward-looking target, not an actual."}
for i in range(10):
    if i in GCO: setcell('gpu_co',i,GCO[i][0],GCO[i][1],GCO[i][2],GCO_NOTE.get(i,''))
    else: setcell('gpu_co',i,ND,'not disclosed','Q2-26 letter / 20-F: installed GPU count not disclosed','')

# --- C. Pricing (market)
sec('C. GPU list pricing (GPU Market & Comps: market/gpu_price_history.xlsx)')
mq=mw['Price_History_Quarterly']; mqh=[c.value for c in mq[1]]
qcols=['2024Q1','2024Q2','2024Q3','2024Q4','2025Q1','2025Q2','2025Q3','2025Q4','2026Q1','2026Q2']
mqr={r[0]:r for r in mq.iter_rows(min_row=2,values_only=True)}
for k,col,lab in SKU:
    key='pq_'+k; newrow(key,f'Nebius {lab} list - quarter-end (latest snapshot in quarter)','$/GPU-hr','Public on-demand list; commitments "up to 35% less" (nebius.com/prices).','GPU Market & Comps')
    for i,qc in enumerate(qcols):
        v=mqr[col][mqh.index(qc)]; t=str(mqr[col][mqh.index(qc+' tag')])
        tag='Reported' if t.startswith('Reported') else ('Estimated' if t.startswith('Estimated') else ('Derived' if t.startswith('Derived') else 'not disclosed'))
        setcell(key,i,v if v is not None else ND,tag,f"{MKT} Price_History_Quarterly '{col}' {qc}",t)
for k,col,lab in SKU:
    key='pa_'+k; newrow(key,f'Nebius {lab} list - quarter average (3 months)','$/GPU-hr','AVERAGE of Prices_Monthly; orange if any month forward-filled.','Derived from GPU Market & Comps')
    for i in range(10):
        mcols=[L(2+3*i+m) for m in range(3)]; row=PMROW[k]
        tags=[next(s[5] for s in SRC if s[0]=='Prices_Monthly' and s[1]==f'{c}{row}') for c in mcols]
        if 'not disclosed' in tags:
            setcell(key,i,ND,'not disclosed',f"Prices_Monthly!{mcols[0]}{row}:{mcols[2]}{row}",'SKU not listed for the full quarter'); continue
        tag='Derived' if all(t=='Reported' for t in tags) else 'Estimated'
        setcell(key,i,f"=AVERAGE(Prices_Monthly!{mcols[0]}{row}:{mcols[2]}{row})",tag,f"Prices_Monthly!{mcols[0]}{row}:{mcols[2]}{row}",'formula')
h1=mw['H100_Nebius_vs_Medians']; hm={r[0]:r for r in h1.iter_rows(min_row=2,values_only=True)}
newrow('neo_med','Neocloud median H100 SXM on-demand (last month of quarter)','$/GPU-hr','Forward-filled median across neoclouds (n in source tab).','GPU Market & Comps')
for i in range(10):
    mo=months[3*i+2]; setcell('neo_med',i,hm[mo][2],'Estimated',f"{MKT} H100_Nebius_vs_Medians {mo} (n_neoclouds={hm[mo][4]})",'median of forward-filled series')
newrow('px_bx','Price used for B300 / GB300 GPUs (proxy)','$/GPU-hr','GB300 has no public Nebius price ("contact us"). Proxy = B300 quarter-avg list when listed (Q2-26), else B200 quarter-avg list (floor).','Filings (method)')
for i in range(10):
    if GEN['B300'][i]+GEN['GB300'][i]==0: setcell('px_bx',i,ND,'not disclosed','no B300/GB300 GPUs modeled','')
    elif TAGS['pa_b300'][i]!='not disclosed': setcell('px_bx',i,f"={ref('pa_b300',i)}",'Estimated','proxy: B300 list','GB300 priced at B300 HGX list')
    else: setcell('px_bx',i,f"={ref('pa_b200',i)}",'Estimated','proxy: B200 list (B300 not listed)','floor proxy')
newrow('px_blend','Fleet-weighted blended list price (quarter avg)','$/GPU-hr','Hopper priced at H100 list (H200 mix unknown -> conservative). Weights = modeled GPUs at quarter end.','Derived (Estimated inputs)')
for i in range(10):
    terms=[];
    pairs=[('g_hop','pa_h100'),('g_h200','pa_h200'),('g_b200','pa_b200')]
    ok=True
    for g,p in pairs:
        if [GEN['Hopper'],GEN['H200'],GEN['B200']][pairs.index((g,p))][i]>0:
            if TAGS[p][i]=='not disclosed': ok=False
            terms.append(f"{ref(g,i)}*{ref(p,i)}")
    if GEN['B300'][i]+GEN['GB300'][i]>0: terms.append(f"({ref('g_b300',i)}+{ref('g_gb300',i)})*{ref('px_bx',i)}")
    if ok: setcell('px_blend',i,'=('+'+'.join(terms)+f")/{ref('gpu',i)}",'Estimated','formula on KPI rows','')
    else: setcell('px_blend',i,ND,'not disclosed','','price missing for a deployed class')

# --- D. Derived KPIs
sec('D. Derived KPIs (live formulas; n.d. where an input is missing)')
newrow('hrs','Hours in quarter','hours','calendar days x 24','Derived')
for i in range(10): setcell('hrs',i,f'={DAYS[i]}*24','Derived','calendar',f'{DAYS[i]} days')
def derived(key,lab,unit,inputs,fx,basis='',owner='Derived'):
    if key not in RIDX: newrow(key,lab,unit,basis,owner)
    for i in range(10):
        need=[]
        for (k,off) in inputs:
            j=i+off
            if j<0: need.append('not disclosed'); continue
            need.append(TAGS[k][j])
        if any(t in (None,'not disclosed') for t in need):
            miss=[f"{k}@{Q[i+off] if i+off>=0 else 'pre-Q1-24'}" for (k,off),t in zip(inputs,need) if t in(None,'not disclosed')]
            setcell(key,i,ND,'not disclosed','input n.d.','missing: '+', '.join(miss)); continue
        tag='Estimated' if 'Estimated' in need else 'Derived'
        f='='+fx(lambda k,off=0: ref(k,i+off))
        setcell(key,i,f,tag,'formula: '+f,'inputs: '+', '.join(f"{k}[{Q[i+off]}]" for k,off in inputs))
derived('ocf_x','','',[('ocf',0),('ddr',0)],lambda R:f"{R('ocf')}-{R('ddr')}")
derived('inserv','','',[('ppeg',0),('aniu',0)],lambda R:f"{R('ppeg')}-{R('aniu')}")
derived('d_aniu','','',[('aniu',-1),('aniu',0)],lambda R:f"{R('aniu')}-{R('aniu',-1)}")
derived('pis_srv','','',[('sng',-1),('sng',0)],lambda R:f"{R('sng')}-{R('sng',-1)}")
derived('gpu_avg','Average GPUs in quarter (opening+closing)/2','GPUs',[('gpu',-1),('gpu',0)],lambda R:f"AVERAGE({R('gpu',-1)},{R('gpu')})",'Q1-24 n.d. (no Q4-23 model value).')
derived('px_real','Implied realized revenue per GPU-hour (AI cloud rev / avg GPUs / hours)','$/GPU-hr',[('ai',0),('gpu_avg',0),('hrs',0)],lambda R:f"{R('ai')}*1000000/({R('gpu_avg')}*{R('hrs')})",'Revenue includes non-GPU-hour items (Token Factory, storage, software, asset-light).')
derived('util','Implied utilization = AI cloud rev / (avg GPUs x hours x blended list $/GPU-hr)','%',[('ai',0),('gpu_avg',0),('hrs',0),('px_blend',0)],lambda R:f"{R('ai')}*1000000/({R('gpu_avg')}*{R('hrs')}*{R('px_blend')})",'Mixes utilization with realized-vs-list discount (commit deals priced below list); >100% possible if modeled GPUs understate fleet or revenue includes non-GPU items. Indicative only.')
derived('util_qe','Implied utilization on quarter-end GPUs (variant)','%',[('ai',0),('gpu',0),('hrs',0),('px_blend',0)],lambda R:f"{R('ai')}*1000000/({R('gpu')}*{R('hrs')}*{R('px_blend')})",'Lower bound vs average-GPU version when fleet is growing.')
derived('rev_mw_qe','AI cloud revenue per energized MW, annualized (quarter-end MW)','$M/MW-yr',[('ai',0),('mw',0)],lambda R:f"{R('ai')}*4/{R('mw')}",'Mixed IT/gross MW basis. Compare: Q2-26 letter ~$12M/MW 2026 installed base; >$20M/MW new deals.')
derived('rev_mw_avg','AI cloud revenue per energized MW, annualized (average MW)','$M/MW-yr',[('ai',0),('mw',-1),('mw',0)],lambda R:f"{R('ai')}*4/AVERAGE({R('mw',-1)},{R('mw')})",'')
derived('rev_mw_mgmt','AI cloud revenue per MW as stated by mgmt (YE MW), annualized','$M/MW-yr',[('ai',0),('mw_mgmt',0)],lambda R:f"{R('ai')}*4/{R('mw_mgmt')}",'Only where mgmt gave a MW figure (YE25 active).')
derived('rev_gpu','AI cloud revenue per GPU, annualized (average GPUs)','$K/GPU-yr',[('ai',0),('gpu_avg',0)],lambda R:f"{R('ai')}*4*1000/{R('gpu_avg')}",'')
derived('dmw','Change in energized MW vs prior quarter','MW',[('mw',-1),('mw',0)],lambda R:f"{R('mw')}-{R('mw',-1)}",'')
derived('capex_mw','Cash capex per MW added (quarter)','$M/MW',[('capex',0),('dmw',0)],lambda R:f'IF({R("dmw")}<=5,"n.m. (dMW<=5)",{R("capex")}/{R("dmw")})','Capex funds future capacity (assets not yet in use); quarterly ratio not meaningful when MW added is ~0.')
derived('capex_mw_ttm','Cash capex per MW added, trailing 4 quarters','$M/MW',[('capex',-3),('capex',-2),('capex',-1),('capex',0),('mw',-4),('mw',0)],lambda R:f'IF({R("mw")}-{R("mw",-4)}<=5,"n.m. (dMW<=5)",SUM({R("capex",-3)}:{R("capex")})/({R("mw")}-{R("mw",-4)}))','')
derived('pis_mw','Placed-in-service PP&E per MW added (gross in-service PP&E change / MW change)','$M/MW',[('inserv',-1),('inserv',0),('mw',-1),('mw',0)],lambda R:f'IF({R("mw")}-{R("mw",-1)}<=5,"n.m. (dMW<=5)",({R("inserv")}-{R("inserv",-1)})/({R("mw")}-{R("mw",-1)}))','Quarterly where consecutive breakdowns exist (Q1-26, Q2-26).')
# FY25 placed-in-service basis (Q4-24 -> Q4-25), written into Q4 2025 column of a separate row
newrow('pis_mw_fy','Placed-in-service PP&E per MW added, FY basis (Q4 prior -> Q4)','$M/MW','FY2025 only (Dec-24 -> Dec-25 breakdowns).','Derived')
for i in range(10):
    if i==7:
        f=f"=({ref('inserv',7)}-{ref('inserv',3)})/({ref('mw',7)}-{ref('mw',3)})"
        setcell('pis_mw_fy',i,f,'Estimated','formula: '+f,'Dec-24 -> Dec-25; MW Estimated')
    else: setcell('pis_mw_fy',i,ND,'not disclosed','','FY basis only computed for FY2025')
derived('arr_m','ARR / 12 = implied last-month AI cloud revenue','$M',[('arr',0)],lambda R:f"{R('arr')}/12",'Definition: last month of quarter x12 (Q2-26 letter footnote 4).')
derived('ai_m','AI cloud revenue - quarter monthly average','$M',[('ai',0)],lambda R:f"{R('ai')}/3",'')
derived('stepup','Last-month vs quarter-average step-up','%',[('arr_m',0),('ai_m',0)],lambda R:f"{R('arr_m')}/{R('ai_m')}-1",'Pre-2026 ARR scope = "revenue" (core AI infra; 2024 from presentation charts) vs AI cloud segment revenue: scope mismatch possible.')
derived('rev_srv_qe','Annualized AI cloud revenue per $1 server & network equipment in service (period end)','$/$',[('ai',0),('sng',0)],lambda R:f"{R('ai')}*4/{R('sng')}",'Gross cost basis.')
derived('rev_srv_avg','Annualized AI cloud revenue per $1 server & network equipment in service (average)','$/$',[('ai',0),('sng',-1),('sng',0)],lambda R:f"{R('ai')}*4/AVERAGE({R('sng',-1)},{R('sng')})",'')
derived('ebitda_mw','AI cloud adj. EBITDA per energized MW, annualized (average MW)','$M/MW-yr',[('ebitda_ai',0),('mw',-1),('mw',0)],lambda R:f"{R('ebitda_ai')}*4/AVERAGE({R('mw',-1)},{R('mw')})",'')
derived('ddr_capex','Change in deferred revenue as % of capex','%',[('ddr',0),('capex',0)],lambda R:f"{R('ddr')}/{R('capex')}",'Prepayment funding of capex.')

# formats
PCT={'util','util_qe','stepup','ddr_capex'}
for key,row in RIDX.items():
    for i in range(10):
        c=ws.cell(row,C0+i)
        if key in PCT: c.number_format='0.0%'
        elif key in('sh','wavg','gpu','gpu_avg','gpu_co') or key.startswith('g_'): c.number_format='#,##0'
        elif key.startswith('p') or key in('rev_srv_qe','rev_srv_avg','neo_med'): c.number_format='0.00'
        else: c.number_format='#,##0.0'
ws.column_dimensions['A'].width=62; ws.column_dimensions['B'].width=11
for i in range(10): ws.column_dimensions[L(C0+i)].width=12
ws.column_dimensions[L(C0+10)].width=70; ws.column_dimensions[L(C0+11)].width=24
ws.freeze_panes='C4'

# ---------------- Checks ----------------
chk['A1']='Checks: Q2 2026 tie-outs to filings and cross-team references (live formulas).'; chk['A1'].font=TITLE
for j,h in enumerate(['Check','Workbook value (formula)','Expected','Difference','Result','Expected source']): chk.cell(3,1+j,h).font=BOLD
K=9
checks=[('AI cloud segment revenue Q2-26 ($M)',f"=KPI!{ref('ai',K)}",574.9,cite('Q2FS','Note 15 segment table')),
 ('ARR end-Jun-26 ($M)',f"=KPI!{ref('arr',K)}",3000,cite('Q2LTR','p.9 ("ARR reached $3.0 billion")')),
 ('Deferred revenue 30-Jun-26 ($M) = current 979.4 + non-current 4,995.8',f"=KPI!{ref('defrev',K)}",'=979.4+4995.8',cite('Q2FS','balance sheet p.F-3')),
 ('RPO 30-Jun-26 ($M)',f"=KPI!{ref('rpo',K)}",37490.6,cite('Q2FS','Note 1 p.F-19')),
 ('Segment sum = group revenue Q2-26 (574.9+1.0+10.0-3.6)',f"=KPI!{ref('grp',K)}",'=574.9+1.0+10.0-3.6',cite('Q2FS','Note 15')),
 ('ARR/12 = June AI cloud revenue ($M)',f"=KPI!{ref('arr_m',K)}",250,'Q2-26 letter (ARR $3.0B / 12)'),
 ('Q2-26 AI cloud monthly average ($M)',f"=KPI!{ref('ai_m',K)}",'=574.9/3','arithmetic'),
 ('FY25 OCF: sum of quarters = 401.9 (20-F)',f"=SUM(KPI!{ref('ocf',4)}:{ref('ocf',7)})",401.9,cite('20F25','cash flow statement p.F-12')),
 ('1H26 OCF: sum = 4,504.1',f"=SUM(KPI!{ref('ocf',8)}:{ref('ocf',9)})",4504.1,cite('Q2PR','cash flow statement')),
 ('1H26 change in deferred revenue = 4,395.0',f"=SUM(KPI!{ref('ddr',8)}:{ref('ddr',9)})",4395.0,cite('Q2PR','cash flow statement')),
 ('Deferred revenue YE24 ($M)',f"=KPI!{ref('defrev',3)}",16.3,cite('Q2PR25','BS comparative Dec-24')+' / '+cite('20F25','BS')),
 ('Deferred revenue 30-Jun-25 ($M)',f"=KPI!{ref('defrev',5)}",19.3,cite('Q2PR25','BS Jun-30-2025')),
 ('Q1-26 servers placed in service = 4692.7-3123.6',f"=KPI!{ref('pis_srv',8)}",'=4692.7-3123.6',cite('Q1FS','Note 7')+' vs YE25'),
 ('Q2-26 servers placed in service = 5197.7-4692.7',f"=KPI!{ref('pis_srv',9)}",'=5197.7-4692.7',cite('Q2FS','Note 7')+' vs Q1FS'),
 ('Q1-26 Δ ANIU = 2427.9-2417.4',f"=KPI!{ref('d_aniu',8)}",'=2427.9-2417.4',cite('Q1FS','Note 7')),
 ('Q2-26 Δ ANIU = 7836.4-2427.9',f"=KPI!{ref('d_aniu',9)}",'=7836.4-2427.9',cite('Q2FS','Note 7')),
 ('FY24 OCF: sum of quarters = (269.9)',f"=SUM(KPI!{ref('ocf',0)}:{ref('ocf',3)})",-269.9,cite('20F25','cash flow statement')),
 ('PP&E gross - accumulated depreciation = net 30-Jun-26',f"=KPI!{ref('ppeg',K)}-1104.3",f"=KPI!{ref('ppe',K)}",cite('Q2FS','Note 7 p.F-23')),
 ('Energized MW Q2-26 (Capacity team reference ~193)',f"=KPI!{ref('mw',K)}",193,'Capacity & Power Quarterly MW row 130 Q2-26'),
 ('Modeled GPUs Q2-26 (Capacity team reference ~89,153)',f"=KPI!{ref('gpu',K)}",89153,'Capacity & Power GPUs deployed row 34 Q2-26 (tolerance 100)'),
 ('GPU generation split sums to total Q2-26',f"=SUM(KPI!{ref('g_hop',K)}:{ref('g_gb300',K)})",f"=KPI!{ref('gpu',K)}",'internal')]
for n,(lab,f,exp,s) in enumerate(checks):
    rr=4+n; chk.cell(rr,1,lab); chk.cell(rr,2,f); chk.cell(rr,3,exp); chk.cell(rr,4,f'=B{rr}-C{rr}')
    tol=100 if 'GPUs' in lab else (0.5 if 'MW' in lab else 0.25)
    chk.cell(rr,5,f'=IF(ABS(D{rr})<={tol},"PASS","CHECK")'); chk.cell(rr,6,s)
chk.column_dimensions['A'].width=60; chk.column_dimensions['F'].width=90
for c in 'BCD': chk.column_dimensions[c].width=14

# ---------------- README ----------------
lines=[
('Nebius Group (NBIS) consolidated KPI history','Built '+__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M')+' ET by Filings & Capital. Script: /workspace/nebius/filings/kpi/kpi_build.py. Public information only.'),
('Tabs','KPI (main grid, 10 quarters), Prices_Monthly (market inputs feeding quarter averages), Checks (tie-outs), Sources (one row per populated cell: tag, source with URL, note).'),
('Tags','Reported = number as printed in an SEC filing / company page / snapshot. Derived = live formula on Reported inputs (or arithmetic on reported totals, math shown). Estimated = a team model value or a formula using one (method in Sources). n.d. = not disclosed / input missing; never filled.'),
('Colours','Blue = Reported input; black = formula on Reported/Derived inputs; orange italic = Estimated (input, or a formula that uses an Estimated input); grey = n.d.'),
('ARR','Annualized run-rate revenue = Nebius AI cloud revenue from the last month of the quarter x 12 (Q2-26 shareholder letter, footnote 4; same wording Q1-26). Q3-25/Q4-25 letters: "revenue from the last month of the quarter multiplied by 12"; Q1-25 letter: "revenue for last month of the period multiplied by twelve". 2024 values read from the Oct-2024 investor presentation chart; Dec-24 $90M from Q4-24 PR.'),
('MW basis','Energized MW is the Capacity & Power team ESTIMATE of power physically live (closest to Nebius "active power"); not "connected" (built and equipped) or "contracted" (land + power secured). Nebius never states IT vs gross. Vineland, Modi\'in and AIB rows are IT MW; others as stated (mostly gross/unknown), so totals mix bases (1.1-1.45x PUE ambiguity). Management figures: YE24 25 MW CONNECTED; YE25 ~170 MW ACTIVE.'),
('Quarter-end vs average','Stock measures (MW, GPUs, PP&E) are quarter-end. Per-unit revenue KPIs are shown on average (opening+closing)/2 and, where noted, quarter-end. With fast build-outs the two can differ a lot (e.g., Q4-25 GPUs 31.8k -> 69.0k).'),
('GPUs','Modeled = energized base MW x GPUs per MW by generation (capacity/gpu_mw_capex.xlsx); company counts are sparse: only actual is ~30,000 deployed at 31-Mar-25 (FY24 20-F; the earlier ~38k from the Feb-25 call was a forward-looking target, superseded); Q3-24 ~10k and Q4-24 >20k are Estimated/targets and ran ahead of powered MW. Generation per site mostly assumed by the Capacity team.'),
('Implied utilization','AI cloud revenue / (average modeled GPUs x hours in quarter x fleet-weighted blended LIST price). It mixes utilization with the realized-vs-list discount (commit deals up to 35% below list; large contracts lower), and revenue includes non-GPU-hour items. Hopper priced at H100 list; GB300 (no public price) proxied by B300 list (Q2-26) or B200 list (earlier, floor). Treat as an index, not a measured utilization.'),
('Capex per MW','Cash capex / change in energized MW. Capex largely funds capacity not yet energized (assets not yet in use rose $2.4B -> $7.8B in 1H26), so quarterly ratios are "n.m." when MW added <= 5 MW; use trailing-4Q and the placed-in-service basis (change in gross in-service PP&E / MW added).'),
('June step-up','ARR/12 (implied last-month revenue) vs AI cloud quarter monthly average. Q2-26: $250M vs $191.6M (+30%); Q1-26: $160M vs $129.9M.'),
('Restatements','Revenue/OCF on latest restated basis (Toloka deconsolidated Q2-25, in discontinued ops). Small company rounding differences noted in Sources (e.g., Q2-25 OCF 167.7 vs 167.9; Q3-24 34.9 vs 34.8).'),
('Inputs','Filings: /workspace/nebius/filings/nebius_filings.xlsx (Quarterly tab via quarterly_rows.json) + edgar/*.txt. Capacity: /workspace/nebius/capacity/nebius_capacity.xlsx (Quarterly MW rows 119-125, GPUs deployed rows 11-31). Market: /workspace/nebius/market/gpu_price_history.xlsx (Price_History_Quarterly, Price_History_Monthly, Monthly_FFill, H100_Nebius_vs_Medians).'),
]
readme.column_dimensions['A'].width=26; readme.column_dimensions['B'].width=160
for n,(a,b) in enumerate(lines):
    readme.cell(1+n,1,a).font=BOLD; c=readme.cell(1+n,2,b); c.alignment=Alignment(wrap_text=True,vertical='top')

# ---------------- Sources ----------------
for j,h in enumerate(['Sheet','Cell','Metric','Period','Value / formula','Tag','Source (form date acc doc ref | URL, or file/sheet/cell)','Note / derivation']): srcs.cell(1,1+j,h).font=BOLD
for n,s in enumerate(SRC):
    for j,v in enumerate(s):
        if isinstance(v,str) and v.startswith('='): v='formula: '+v
        srcs.cell(2+n,1+j,v)
for c,w in zip('ABCDEFGH',[14,8,60,10,40,13,110,80]): srcs.column_dimensions[c].width=w
srcs.freeze_panes='A2'
wb.calculation.fullCalcOnLoad=True
wb.save(OUT); print('saved',OUT,'cells',len(SRC))
