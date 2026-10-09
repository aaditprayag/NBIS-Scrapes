import csv,json,sys,datetime as dt
sys.path.insert(0,'/workspace/nebius/skeptic/speed/code')
from curated import EXTRA_PRICES,TIMELINE,EVENTS
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter as L
B='/workspace/nebius/skeptic/speed/'
H=Font(bold=True,color='FFFFFF'); F=PatternFill('solid',fgColor='7B1F1F'); INP=PatternFill('solid',fgColor='FFF2CC'); BOLD=Font(bold=True)
def hdr(ws,r):
    for c in ws[r]:
        if c.value is not None: c.font=H; c.fill=F; c.alignment=Alignment(wrap_text=True,vertical='top')
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[L(i)].width=x
D=lambda s:dt.date.fromisoformat(s)
wb=Workbook(); ws=wb.active; ws.title='README'
for r in [['Nebius Skeptic (Red Team) | Speed / residual value / credit studies | built 2026-10-04 (ET)'],
 ['Tags: Reported = as published at source; Derived = arithmetic on reported inputs (formula in the cell, or math in the basis column); Estimated = method plus assumption, yellow input cells. "not found" means not found.'],
 ['Prices are posted list or index rates, not transacted prices. Wayback captures were downloaded by the GPU Market teammate to /workspace/nebius/wayback_raw and parsed here (code/parse_series.py). The Wayback URL is on every row.'],
 ['Sheets: S1_PriceSeries | S1_Steps | S1_Timeline | S1_Decay | S1_PrevGen | S1_HS_Entry | S1_HeadStart | S1_Reconcile | S2_Residual | S2_RentalPrecedent | S3_Credit | Sources'],
 ['Yellow cells are inputs. Change them and the outputs recalc. Python-solved values (convert-implied spread/vol, Merton) are pasted as values and labeled Estimated (code/credit_model.py).']]:
    ws.append(r)
ws['A1'].font=Font(bold=True,size=13); ws.column_dimensions['A'].width=160
# ---------- S1_PriceSeries
ps=list(csv.DictReader(open(B+'data/price_series.csv')))
for r in ps:
    if r['provider']=='Lambda' and r['date']=='2023-12-31': r['date']='2023-08-02'; r['basis']+=' [valid from 8x HGX on-demand launch 2023-08-02 through 2023-12-31]'
for p in EXTRA_PRICES:
    ps.append(dict(provider=p[0],chip=p[1],date=p[2],usd_per_gpu_hr=round(p[3],4),basis=p[4],tag=p[5],source_url=p[6],snippet=''))
ps.sort(key=lambda r:(r['provider'],r['chip'],r['date']))
s=wb.create_sheet('S1_PriceSeries'); s.append(['provider','chip','capture/effective date','$/GPU-hr','basis','tag','source_url','page text snippet']); hdr(s,1)
for r in ps: s.append([r['provider'],r['chip'],D(r['date']),float(r['usd_per_gpu_hr']),r['basis'],r['tag'],r['source_url'],r.get('snippet','')])
for c in s['C'][1:]: c.number_format='yyyy-mm-dd'
for c in s['D'][1:]: c.number_format='$0.000'
widths(s,[26,13,12,10,60,10,70,60]); s.freeze_panes='A2'; s.auto_filter.ref=s.dimensions
# ---------- S1_Steps (auto from series)
st=wb.create_sheet('S1_Steps'); st.append(['provider','chip','last capture at old price','first capture at new price','old $/GPU-hr','new $/GPU-hr','% change (Derived)','note','source (new-price capture)']); hdr(st,1)
grp={}
for r in ps: grp.setdefault((r['provider'],r['chip']),[]).append(r)
NOTES={('Nebius','H100','2026-07-12'):'Effective 2026-06-01 per Nebius customer notice ($2.95 to $3.85, +30.5%): https://knowyourinvestments.substack.com/p/nebius-watch-gpu-rent-price-hikes',
 ('Nebius','H100','2026-10-01'):'Effective 2026-10-01 (17-21% hikes): https://www.reuters.com/technology/nebius-hikes-ai-cloud-prices-again-demand-computing-power-soars-2026-09-17/',
 ('CoreWeave','H100','2025-01-01'):'BASIS CHANGE: legacy per-GPU component rate (excl. CPU/RAM) replaced by 8-GPU instance price. Not a price move.',
 ('Lambda','H100','2024-01-01'):'Lambda page: "2024 Price (effective 1/1/2024)" vs "2023 Price". A price INCREASE.'}
r0=2
for k,g in grp.items():
    for a,b in zip(g,g[1:]):
        if abs(float(a['usd_per_gpu_hr'])-float(b['usd_per_gpu_hr']))>1e-6 and 'index' not in k[0] and 'median' not in k[0] and 'read' not in k[0]:
            st.append([k[0],k[1],D(a['date']),D(b['date']),float(a['usd_per_gpu_hr']),float(b['usd_per_gpu_hr']),f'=F{st.max_row+1}/E{st.max_row+1}-1',NOTES.get((k[0],k[1],b['date']),''),b['source_url']])
for row in st.iter_rows(min_row=2):
    row[2].number_format=row[3].number_format='yyyy-mm-dd'; row[4].number_format=row[5].number_format='$0.000'; row[6].number_format='+0.0%;-0.0%'
widths(st,[22,12,13,13,10,10,10,70,70]); st.freeze_panes='A2'
# ---------- S1_Timeline
tl=wb.create_sheet('S1_Timeline')
tl.append(['chip','provider','type','first availability date','event','tag','months after first neocloud (Derived)','months after first hyperscaler (Derived)','Nebius lag vs this provider, months (+ = Nebius later)','source'])
hdr(tl,1)
chips=['H100','H200','GB200','B200']
start=2
for ch in chips:
    rows=[t for t in TIMELINE if t[0]==ch]; first=tl.max_row+1; last=first+len(rows)-1
    neb=[i for i,t in enumerate(rows) if t[1]=='Nebius'][0]+first
    for i,t in enumerate(rows):
        r=tl.max_row+1
        tl.append([t[0],t[1],t[2],D(t[3]),t[4],t[6],
          f'=(D{r}-_xlfn.MINIFS($D${first}:$D${last},$C${first}:$C${last},"Neocloud"))/30.44',
          f'=(D{r}-_xlfn.MINIFS($D${first}:$D${last},$C${first}:$C${last},"Hyperscaler"))/30.44',
          f'=($D${neb}-D{r})/30.44',t[5]])
    tl.append([None])
for row in tl.iter_rows(min_row=2):
    row[3].number_format='yyyy-mm-dd'
    for i in (6,7,8): row[i].number_format='0.0'
widths(tl,[8,12,12,13,60,14,12,12,14,80]); tl.freeze_panes='A2'
tl.cell(row=tl.max_row+2,column=1,value='Read: Nebius row in col G = months behind the first neocloud; col H = months vs first hyperscaler (positive = after). Dates are public GA or on-demand announcements, not first-customer or scale dates. Reconcile with the GPU Market time-to-deploy table when it lands in /workspace/nebius/market/.')
# ---------- S1_Decay
dc=wb.create_sheet('S1_Decay')
dc.append(['Premium decay vs. own steady state. Steady state = trough posted price before the 2026 shortage repricing (captures through 2025-12-31). Half-life = months from first capture until price <= trough + 50% of the launch gap. Bounded by capture dates.'])
dc.append(['provider','chip','launch/first price','first date','trough (pre-2026)','trough date','half-gap threshold','first capture <= threshold','half-life, months (Derived)','launch premium vs trough','latest price','latest date','latest vs trough','note / source'])
hdr(dc,2)
DEC=[('Nebius','H100',4.85,'2023-11-01',2.95,'2024-11-15','2024-09-07',4.50,'2026-10-01','Last capture above threshold 2024-08-05 ($4.85). Wayback nebius.ai/prices and nebius.com/prices.'),
 ('Lambda (8x HGX)','H100',2.59,'2023-08-02',2.99,'2024-08-15',None,3.99,'2026-10-01','Launch price was BELOW later steady state; Lambda raised to $3.49 on 2024-01-01. No launch premium to decay.'),
 ('AWS on-demand (P5)','H100',12.29,'2023-07-26',6.88,'2025-06-01','2025-06-01',6.88,'2026-10-04','One-time 44% step cut after ~22 months, not a gradual decay. AWS blog plus price-list JSON.'),
 ('AWS Capacity Blocks (p5)','H100',4.916,'2024-03-28',3.933,'2025-03-08','2025-03-08',5.191,'2026-10-01','First capture 2024-03-28 (CB launch earlier, not captured). -20% step.'),
 ('CCIR panel median','H100',3.50,'2023-09-15',None,None,None,3.64,'2026-04-15','Panel median ends the window higher than it started (CCIR). Trough not published.'),
 ('Nebius','H200',4.85,'2024-10-06',3.50,'2024-11-15','2024-11-15',5.40,'2026-10-01','Pre-GA list $4.85 (2024-10-06 capture) cut to $3.50 by GA (2024-11-15 capture) - premium gone within ~1.3 months, before availability; then flat until 2026 increases.'),
 ('AWS Capacity Blocks (p5e)','H200',5.408,'2024-10-01',4.326,'2025-03-08','2025-03-08',5.970,'2026-10-01','-20% step about 5 months after launch.'),
 ('AWS on-demand (P5en)','H200',10.551,'2024-12-03',7.912,'2025-06-01','2025-06-01',7.912,'2026-10-04','Pre-cut price Derived = post-cut / 0.75. -25% step after ~6 months.'),
 ('Nebius','B200',5.50,'2025-07-25',5.50,'2025-07-25',None,8.50,'2026-10-01','No decay before 2026; +55% since.'),
 ('Lambda (8x)','B200',4.99,'2025-09-11',4.99,'2025-09-11',None,6.69,'2026-10-01','No decay; +34% since.'),
 ('CoreWeave','B200',8.60,'2025-06-09',8.60,'2025-06-09',None,8.60,'2026-10-01','Sticky list (CRWV says the July-26 ~25% increase hit contracts, not the rate card).'),
 ('AWS Capacity Blocks','B200',8.14,'2025-05-22',8.14,'2025-05-22',None,12.355,'2026-10-01','Rising: $8.14, then $9.36 (Oct-25), $10.30 (Apr-26), $12.36 (Aug-26).'),
 ('CoreWeave','GB200 NVL72',10.50,'2025-02-09',10.50,'2025-02-09',None,10.50,'2026-10-01','Flat.'),
 ('AWS Capacity Blocks','GB200 NVL72',10.582,'2025-08-05',10.582,'2025-08-05',None,10.582,'2026-10-01','Flat.'),
 ('Nebius','GB200 NVL72',None,None,None,None,None,None,None,'not found: "Contact us" on every capture.')]
for d in DEC:
    r=dc.max_row+1
    dc.append([d[0],d[1],d[2],D(d[3]) if d[3] else None,d[4],D(d[5]) if d[5] else None,
      f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),E{r}+0.5*(C{r}-E{r}),"n/a")',D(d[6]) if d[6] else None,
      f'=IF(ISNUMBER(H{r}),(H{r}-D{r})/30.44,IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),IF(C{r}<=E{r},"no premium / no decay","not reached"),"n/a"))',
      f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),C{r}/E{r}-1,"n/a")',d[7],D(d[8]) if d[8] else None,
      f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(E{r})),K{r}/E{r}-1,"n/a")',d[9]])
for row in dc.iter_rows(min_row=3):
    for i in (3,5,7,11): row[i].number_format='yyyy-mm-dd'
    for i in (2,4,6,10): row[i].number_format='$0.000'
    row[8].number_format='0.0'; row[9].number_format=row[12].number_format='+0%;-0%'
widths(dc,[26,12,10,11,10,11,10,12,12,11,9,11,10,90]); dc.freeze_panes='C3'
# ---------- S1_PrevGen
pg=wb.create_sheet('S1_PrevGen')
pg.append(['Price of new generation relative to prior generation, same provider, same capture date (Reported prices; ratio Derived). Stable ratio = performance ladder; decaying ratio = launch premium.'])
pg.append(['provider','date','new chip','prior chip','new $/GPU-hr','prior $/GPU-hr','ratio new/prior','source']); hdr(pg,2)
PGR=[('Nebius','2024-11-15','H200','H100',3.50,2.95),('Nebius','2026-05-07','H200','H100',3.50,2.95),('Nebius','2026-07-12','H200','H100',4.50,3.85),('Nebius','2026-10-01','H200','H100',5.40,4.50),
 ('Nebius','2025-07-25','B200','H100',5.50,2.95),('Nebius','2026-05-07','B200','H100',5.50,2.95),('Nebius','2026-07-12','B200','H100',7.15,3.85),('Nebius','2026-10-01','B200','H100',8.50,4.50),
 ('Nebius','2025-07-25','B200','H200',5.50,3.50),('Nebius','2026-10-01','B200','H200',8.50,5.40),
 ('Lambda','2025-09-11','B200','H100',4.99,2.99),('Lambda','2026-03-10','B200','H100',5.74,3.44),('Lambda','2026-10-01','B200','H100',6.69,3.99),
 ('AWS Capacity Blocks','2024-10-01','H200','H100',5.408,4.916),('AWS Capacity Blocks','2025-03-08','H200','H100',4.326,3.933),('AWS Capacity Blocks','2026-08-25','H200','H100',5.970,5.191),
 ('AWS Capacity Blocks','2025-05-22','B200','H100',8.14,3.933),('AWS Capacity Blocks','2026-08-25','B200','H100',12.355,5.191),
 ('CoreWeave','2025-07-18','H200','H100',6.305,6.155),('CoreWeave','2025-07-18','B200','H100',8.60,6.155),('CoreWeave','2025-07-18','GB200 NVL72','B200',10.50,8.60),('CoreWeave','2026-10-01','B200','H100',8.60,6.155)]
for p in PGR:
    r=pg.max_row+1; pg.append([p[0],D(p[1]),p[2],p[3],p[4],p[5],f'=E{r}/F{r}','S1_PriceSeries (Wayback capture on that date)'])
for row in pg.iter_rows(min_row=3): row[1].number_format='yyyy-mm-dd'; row[4].number_format=row[5].number_format='$0.000'; row[6].number_format='0.00"x"'
widths(pg,[22,11,12,10,10,10,10,45])
# ---------- S1_HS_Entry
he=wb.create_sheet('S1_HS_Entry')
he.append(['Hyperscaler late-arrival / price-cut events vs. neocloud list price (last capture <= event-45d) and (last capture <= event+90d). % change Derived.'])
he.append(['event','event date','event','neocloud','chip priced','same chip / prior gen','before date','before $','after date','after $','% change','note','source before','source after']); hdr(he,2)
hs=list(csv.DictReader(open(B+'data/hs_entry_test.csv')))
for x in hs:
    r=he.max_row+1
    he.append([x['event'],D(x['event_date']),x['event_name'],x['neocloud'],x['chip'],x['same_gen'],x['before_date'],float(x['before']) if x['before'] else 'not observed',x['after_date'],float(x['after']) if x['after'] else 'not observed',
      f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(J{r}),L{r}=""),J{r}/H{r}-1,"n/a")',x['note'],x['src_before'],x['src_after']])
last=he.max_row
for row in he.iter_rows(min_row=3): row[1].number_format='yyyy-mm-dd'; row[10].number_format='+0.0%;-0.0%'
he.append([None]); r=he.max_row+1
he.append(['Summary (Derived, live):']); he.cell(row=r,column=1).font=BOLD
he.append(['measurable pairs, all events',f'=COUNT(K3:K{last})'])
he.append(['pairs with a step-DOWN <= -5%',f'=COUNTIF(K3:K{last},"<=-0.05")'])
he.append(['measurable pairs, 2025 events (E7-E12)',f'=SUMPRODUCT((B3:B{last}>=DATE(2025,1,1))*ISNUMBER(K3:K{last}))'])
he.append(['step-downs <= -5%, 2025 events',f'=SUMPRODUCT((B3:B{last}>=DATE(2025,1,1))*ISNUMBER(K3:K{last})*(N(+K3:K{last})<=-0.05))'])
he.append(['Read: the only step-downs (Nebius H100 -39%, Lambda H100 -14%) fall in the Aug-Nov 2024 H200 wave and hit the PRIOR generation. That wave was also a period of broad H100 supply loosening, so this is correlation, not proof that hyperscaler entry caused it. 2023 H100 entries: no pre-event neocloud captures (not observed).'])
widths(he,[7,11,40,11,12,12,11,9,11,9,9,30,50,50]); he.freeze_panes='A3'
# ---------- S1_HeadStart
hsd=wb.create_sheet('S1_HeadStart')
inputs=[('Launch premium over steady state, pi0','=S1_Decay!J3','Derived: Nebius H100 $4.85 vs $2.95 trough (link to S1_Decay)'),
 ('Premium half-life (months)','=S1_Decay!I3','Derived: Nebius H100 (link)'),
 ('Head-start N (months)',3,'Input. Note: Nebius actual head start vs first neocloud is NEGATIVE (S1_Timeline)'),
 ('GPU economic/book life (months)',48,'Nebius 20-F: GPUs depreciated over 4 years (servers/network extended to 5y from 2026-01-01)'),
 ('Steady-state price $/GPU-hr','=S1_Decay!E3','Derived: Nebius H100 trough (link)'),
 ('Utilization',0.70,'Estimated: Mercatus worked example uses 70% (https://www.mercatus-ai.com/blog/h100-resale-value)'),
 ('Hours per month',730,'8760/12'),
 ('Annual discount rate',0.12,'Estimated: illustrative equity-like rate for GPU cash flows (sensitivity only)'),
 ('Share of the generation\'s Nebius fleet online during the head-start window, f',0.15,'Estimated, illustrative: early-window cohort as a share of the generation\'s eventual fleet. Not disclosed. Change freely.'),
 ('Monthly decay factor q','=0.5^(1/B4)','Derived')]
hsd.append(['Value of an N-month head start: excess premium earned in the first N months as a share of the GPU\'s lifetime revenue']); hsd['A1'].font=BOLD
hsd.append(['input','value','source / tag']); hdr(hsd,2)
for i,(a,b,c) in enumerate(inputs): hsd.append([a,b,c]); 
for r in range(3,13):
    if r in (5,6,7,9,10,11): pass
for r in (5,6,8,10,11): hsd.cell(row=r,column=2).fill=INP
hsd['B3'].number_format='0.0%'; hsd['B4'].number_format='0.0'; hsd['B8'].number_format='0%'; hsd['B10'].number_format='0%'; hsd['B11'].number_format='0%'; hsd['B12'].number_format='0.0000'; hsd['B7'].number_format='$0.00'
hsd.append([None]); hsd.append(['OUTPUTS (Derived)']); hsd.cell(row=hsd.max_row,column=1).font=BOLD
o0=hsd.max_row+1
MT=30  # monthly table start row
outs=[('Per-GPU: excess premium in window / lifetime revenue (undiscounted)',f'=SUM(F{MT}:F{MT+119})/SUM(E{MT}:E{MT+119})','0.00%'),
 ('Per-GPU: same, discounted (PV)',f'=SUM(I{MT}:I{MT+119})/SUM(H{MT}:H{MT+119})','0.00%'),
 ('Per-GPU excess PV $',f'=SUM(I{MT}:I{MT+119})','$#,##0'),
 ('Per-GPU lifetime revenue PV $',f'=SUM(H{MT}:H{MT+119})','$#,##0'),
 ('Fleet-level (x f): head-start value as share of the generation\'s lifetime revenue (PV)',f'=B{o0+1}*B11','0.00%')]
for a,f,n in outs:
    hsd.append([a,f]); hsd.cell(row=hsd.max_row,column=2).number_format=n
hsd.append([None]); g0=hsd.max_row+1
hsd.append(['Sensitivity (undiscounted, per GPU): rows = pi0, cols = N months; uses B4 half-life and B6 life']); hsd.cell(row=g0,column=1).font=BOLD
Ns=[1,3,6,12]; Ps=[0.2,0.4,0.64,1.0]
hsd.append(['pi0 \\ N']+Ns)
for p in Ps:
    r=hsd.max_row+1; row=[p]
    for j,n in enumerate(Ns):
        c=L(2+j); row.append(f'=($A{r}*(1-$B$12^{c}${g0+1})/(1-$B$12))/($B$6+$A{r}*(1-$B$12^$B$6)/(1-$B$12))')
    hsd.append(row)
    for j in range(len(Ns)): hsd.cell(row=r,column=2+j).number_format='0.0%'
    hsd.cell(row=r,column=1).number_format='0%'
hsd.cell(row=MT-2,column=1,value='Monthly table (Derived)')
for j,v in enumerate(['month t','premium pi(t)','price $/GPU-hr','in head-start window?','revenue $ (if t<life)','excess premium $ (window only)','discount factor','PV revenue','PV excess'],1): hsd.cell(row=MT-1,column=j,value=v)
hdr(hsd,MT-1)
for t in range(120):
    r=MT+t
    for j,v in enumerate([t,f'=$B$3*$B$12^A{r}',f'=$B$7*(1+B{r})',f'=A{r}<$B$5',f'=IF(A{r}<$B$6,C{r}*$B$8*$B$9,0)',f'=IF(AND(D{r},A{r}<$B$6),$B$7*B{r}*$B$8*$B$9,0)',f'=1/(1+$B$10)^(A{r}/12)',f'=E{r}*G{r}',f'=F{r}*G{r}'],1): hsd.cell(row=r,column=j,value=v)
    hsd.cell(row=r,column=2).number_format='0.0%'; hsd.cell(row=r,column=3).number_format='$0.00'
    for c in (5,6,8,9): hsd.cell(row=r,column=c).number_format='$#,##0'
widths(hsd,[70,12,70,12,14,14,10,12,12])
# ---------- S2_Residual
rv=wb.create_sheet('S2_Residual')
rv.append(['Secondary-market residual value vs. book schedules. Age is measured from first cloud GA (H100: 2023-03-21, CoreWeave) or DGX A100 launch (2020-05-14). Prices are asks/brokered reads unless stated. eBay sold listings: not viewable from this environment (eBay returned an error page to scripted requests). Not found.'])
rv.append(['item','as-of','reference start date','age (yrs)','used low $','used high $','used central $','new/launch reference $','reference basis','residual % (central)','residual % low','residual % high','implied annual decline (geometric)','years to 15% salvage at that rate','book NBV % (4y SL)','book NBV % (5y SL)','book NBV % (6y SL)','tag','source']); hdr(rv,2)
RES=[('H100 8-GPU HGX server, used','2026-04-08','2023-03-21',150000,180000,None,(246000+315000)/2,'new 8xH100 server listings $246k-$315k midpoint (Dell XE9680 / Supermicro listings, retrieved 2026-10-04; listing date not shown)','Reported (ranges) / Derived (%)','https://hashrateindex.com/blog/used-gpu-market-pricing-deprecation-secondary-ai/ ; https://www.interpromicro.com/product/40091-dell-edge-xe9680-6u-server-with-8xh100-gpus ; https://viperatech.com/product/supermicro-superserver-sys-821ge-tnhr-sxm5-640gb-hgx-h100'),
 ('H100 PCIe 80GB, used (first publishable used read)','2026-09-01','2023-03-21',28000,32399,32290,29500,'Clustermelt new-condition H100 PCIe read, same issue','Reported / Derived','https://clustermelt.com/market-intelligence/2026-09'),
 ('H100 SXM5 module, used/refurb','2026-06-23','2023-03-21',18000,22000,None,30000,'Mercatus $30,000 OEM list reference; refurb channel $18-22k','Reported / Derived','https://www.mercatus-ai.com/blog/h100-resale-value'),
 ('H100 SXM5, 36-month residual (Mercatus survey range)','2026-06-23','2023-03-21',12000,22500,None,30000,'Mercatus 36-mo band 40-75% of $30k list','Reported','https://www.mercatus-ai.com/blog/h100-resale-value'),
 ('H100 used, mid-2024 peak scarcity','2024-07-01','2023-03-21',50000,50000,50000,32500,'retail $25-40k midpoint (hashrateindex)','Reported / Derived (mid-2024 date approx.)','https://hashrateindex.com/blog/used-gpu-market-pricing-deprecation-secondary-ai/'),
 ('A100 SXM4 80GB, used','2026-09-01','2020-05-14',9900,10000,9900,199000/8,'DGX A100 launch $199,000 / 8 GPUs (system-level, 40GB; overstates the module price)','Reported / Derived','https://clustermelt.com/market-intelligence/2026-09 ; https://nvidianews.nvidia.com/news/nvidia-ships-worlds-most-advanced-ai-system-nvidia-dgx-a100-to-fight-covid-19-third-generation-dgx-packs-record-5-petaflops-of-ai-performance'),
 ('A100 PCIe 80GB, used','2026-10-01','2020-05-14',12550,18879,12550,199000/8,'DGX A100 $199k / 8 (proxy)','Reported / Derived','https://clustermelt.com/gpus/a100-pcie-80gb/price'),
 ('A100 PCIe 40GB, used-grade (Aug) / refurb (Sep)','2026-09-01','2020-05-14',5750,8800,5750,199000/8,'DGX A100 $199k / 8 (proxy)','Reported / Derived','https://clustermelt.com/market-intelligence/2026-09'),
 ('A100 range, all SKUs (asks)','2026-04-08','2020-05-14',7800,18900,None,199000/8,'DGX A100 $199k / 8 (proxy)','Reported / Derived','https://hashrateindex.com/blog/used-gpu-market-pricing-deprecation-secondary-ai/')]
for x in RES:
    r=rv.max_row+1
    rv.append([x[0],D(x[1]),D(x[2]),f'=(B{r}-C{r})/365.25',x[3],x[4],x[5] if x[5] else f'=AVERAGE(E{r}:F{r})',x[6],x[7],f'=G{r}/H{r}',f'=E{r}/H{r}',f'=F{r}/H{r}',
      f'=J{r}^(1/D{r})-1',f'=IF(J{r}<1,LN(0.15)/LN(1+M{r}),"n/a (residual >= 100%)")',f'=MAX(0,1-D{r}/4)',f'=MAX(0,1-D{r}/5)',f'=MAX(0,1-D{r}/6)',x[8],x[9]])
for row in rv.iter_rows(min_row=3):
    row[1].number_format=row[2].number_format='yyyy-mm-dd'; row[3].number_format='0.0'
    for i in (4,5,6,7): row[i].number_format='$#,##0'
    for i in (9,10,11,14,15,16): row[i].number_format='0%'
    row[12].number_format='+0.0%;-0.0%'; row[13].number_format='0.0'
widths(rv,[44,11,11,7,10,10,10,11,48,9,9,9,10,10,9,9,9,18,80]); rv.freeze_panes='B3'
n=rv.max_row; rv.append([None])
rv.append(['Read: at ~3.5 years H100 hardware is marked at ~53-64% of new (servers) and ~60-73% (refurb modules), vs 4y straight-line book NBV of ~13% and 6y of ~42%. A100 at ~6.3 years: ~40-50% of a system-level launch proxy. On today\'s marks, book lives look CONSERVATIVE, not aggressive. But the marks are regime-dependent: H100 went from ~$50k (mid-2024) to "steep discounts" and back. The 2026 marks reflect a shortage.'])
# ---------- S2_RentalPrecedent
rp=wb.create_sheet('S2_RentalPrecedent')
rp.append(['A100 / V100 rental precedent for Hopper years 3-5. Posted on-demand $/GPU-hr. Age from first cloud availability (A100: AWS p4d 2020-11-02; H100: 2023-03-21). A100 points reuse the teammate\'s Wayback captures (parsed here) plus CCIR\'s nine-year panel.'])
rp.append(['series','date','start date','age (yrs)','$/GPU-hr','vs first point in series','tag','source']); hdr(rp,2)
RP=[('CCIR A100 panel median','2021-09-15','2020-11-02',2.20),('CCIR A100 panel median','2022-06-15','2020-11-02',3.05),('CCIR A100 panel median','2025-07-15','2020-11-02',1.75),('CCIR A100 panel median','2026-04-15','2020-11-02',1.99),
 ('AWS p4d on-demand (A100 40GB)','2020-11-02','2020-11-02',32.77/8),('AWS p4d on-demand (A100 40GB)','2025-06-01','2020-11-02',21.957642/8),
 ('Nebius A100 SXM4','2023-11-01','2020-11-02',2.66),('Nebius A100 SXM4','2024-09-07','2020-11-02',2.66),
 ('CoreWeave A100 80GB (component rate)','2024-01-08','2020-11-02',2.21),('CoreWeave A100 (8x instance /8)','2025-05-11','2020-11-02',2.70),
 ('Lambda 8x A100 80GB','2023-08-02','2020-11-02',1.50),('Lambda 8x A100 80GB','2024-01-01','2020-11-02',1.79),('Lambda 8x A100 80GB','2026-03-10','2020-11-02',2.06),('Lambda 8x A100 80GB','2026-04-01','2020-11-02',2.79),
 ('CCIR V100 panel median','2019-06-15','2017-10-25',2.69),('CCIR V100 panel median','2026-04-15','2017-10-25',0.81),
 ('CCIR H100 panel median','2023-09-15','2023-03-21',3.50),('CCIR H100 panel median','2026-04-15','2023-03-21',3.64),('Nebius H100','2026-10-01','2023-03-21',4.50)]
first={}
for x in RP:
    r=rp.max_row+1
    if x[0] not in first: first[x[0]]=r
    src={'CCIR':'https://ccir.io/research/rental-rates-by-generation','AWS ':'https://techcrunch.com/2020/11/02/aws-launches-its-next-gen-gpu-instances/ ; /workspace/nebius/raw/aws_od_use1.json'}.get(x[0][:4],'S1_PriceSeries (Wayback)')
    note='Reported' if x[0].startswith('CCIR') or x[0].startswith('Nebius') else 'Derived (instance/8)' if '/8' in x[0] or 'AWS' in x[0] or 'Lambda' in x[0] else 'Reported'
    rp.append([x[0],D(x[1]),D(x[2]),f'=(B{r}-C{r})/365.25',round(x[3],4),f'=E{r}/E${first[x[0]]}-1',note,src])
for row in rp.iter_rows(min_row=3): row[1].number_format=row[2].number_format='yyyy-mm-dd'; row[3].number_format='0.0'; row[4].number_format='$0.00'; row[5].number_format='+0%;-0%'
for rr in (17,18):
    rp.cell(row=rr,column=3).value=None; rp.cell(row=rr,column=4).value='n/a'; rp.cell(row=rr,column=7).value='Reported (V100 launch date not verified here; age not computed)'
    rp.cell(row=rr,column=6).value='=E{0}/E$17-1'.format(rr)
rp.append([None]); p0=rp.max_row+1
rp.append(['H100 years 3-5 projection (Estimated: apply precedent rates to H100)']); rp.cell(row=p0,column=1).font=BOLD
rp.append(['input/output','value','method / source']); hdr(rp,p0+1)
rows=[('H100 base posted price (CCIR median, Apr-2026, age ~3.1y)','=E20','CCIR'),
 ('Nebius H100 Oct-2026 on-demand','=E21','Wayback/nebius.com'),
 ('A100 trough / peak ratio (2025 trough vs 2022 peak)','=E5/E4','Derived from CCIR'),
 ('V100 log-linear decline per year','=-0.211','Reported fit, CCIR'),
 ('A100 endpoint CAGR per year','=-0.022','Reported, CCIR'),
 ('Scenario A, A100-like: Nebius Oct-26 $4.50 treated as peak x trough/peak','=B{0}*B{1}','Derived'),
 ('Scenario B, V100-like: CCIR base x (1-21.1%)^(5-3.1) (age 5 = 2028)','=B{2}*(1+B{3})^(5-3.1)','Derived'),
 ('Scenario C, A100 endpoint-like: base x (1-2.2%)^(5-3.1)','=B{2}*(1+B{4})^(5-3.1)','Derived'),
 ('Scenario A vs Nebius Oct-26 price','=B{5}/B{0}-1','Derived'),('Scenario B vs Nebius Oct-26 price','=B{6}/B{0}-1','Derived'),('Scenario C vs Nebius Oct-26 price','=B{7}/B{0}-1','Derived')]
base=rp.max_row+1
idx=[base+i for i in range(len(rows))]
for i,(a,f,src) in enumerate(rows):
    f=f.replace('{0}',str(idx[1])).replace('{1}',str(idx[2])).replace('{2}',str(idx[0])).replace('{3}',str(idx[3])).replace('{4}',str(idx[4])).replace('{5}',str(idx[5])).replace('{6}',str(idx[6])).replace('{7}',str(idx[7]))
    rp.append([a,f,src])
for i in idx[:2]+idx[5:8]: rp.cell(row=i,column=2).number_format='$0.00'
for i in idx[2:5]+idx[8:]: rp.cell(row=i,column=2).number_format='+0.0%;-0.0%' if i not in (idx[2],) else '0.00'
rp.append([None]); rp.append(['Cash-cost floor (Estimated): power + colo + ops ~$0.40/GPU-hr (Mercatus worked example, H100-class). A100 at $1.75-2.79 in years 5-6, and V100 at $0.81 at ~8.5 years, are both still above that floor. On a cash basis, economic life looks longer than 6 years at POSTED prices. Posted is not transacted, and older-generation utilization is not disclosed.'])
widths(rp,[62,12,11,8,9,10,22,70])
# ---------- S3_Credit
cm=json.load(open(B+'data/credit_model_out.json'))
cr=wb.create_sheet('S3_Credit')
cr.append(['Credit cross-check. NBIS convert secondary trading levels: NOT FOUND. Notes are Rule 144A; FINRA TRACE public API returned 401 (auth required); no press quote found. Levels below are issue terms (Reported), model values (Estimated), or exchange-implied (Derived).']); cr['A1'].font=BOLD
cr.append([None]); cr.append(['A. CoreWeave straight bonds (secondary levels)']); cr.cell(row=3,column=1).font=BOLD
cr.append(['bond','ISIN','price','YTW','coupon','maturity','price date','G-spread bp','mod dur','implied UST (Derived) = YTW - G-spread','source']); hdr(cr,4)
for b in [('CRWV 9.25% 2030','US21873SAB43',93.481,0.1145,0.0925,'2030-06-01',651,2.90),('CRWV 9.00% 2031','US21873SAC26',89.342,0.1221,0.09,'2031-02-01',723,3.37),('CRWV 9.75% 2031','US21873SAG30',87.024,0.1339,0.0975,'2031-10-01',838,3.74),('CRWV 9.625% 2032','US21873SAK42',84.996,0.1341,0.09625,'2032-07-15',836,4.04)]:
    r=cr.max_row+1; cr.append([b[0],b[1],b[2],b[3],b[4],D(b[5]),D('2026-10-01'),b[6],b[7],f'=D{r}-H{r}/10000','https://bondterminal.com/issuers/coreweave-inc (Reported, updated 2 Oct 26)'])
    cr.cell(row=r,column=4).number_format=cr.cell(row=r,column=5).number_format=cr.cell(row=r,column=10).number_format='0.00%'; cr.cell(row=r,column=6).number_format=cr.cell(row=r,column=7).number_format='yyyy-mm-dd'
cr.append([None]); cr.append(['B. Convertible terms and bond-only yield (if never converted, bought at par)']); cr.cell(row=cr.max_row,column=1).font=BOLD
cr.append(['note','pricing date','coupon','accretes to (% of orig.)','years to maturity','share price at pricing','conversion price','conversion premium (Derived)','effective conv. price at maturity (Derived)','shares per $1,000','UST at tenor (interp., Treasury par curve)','bond-only YTM at par (Derived, RATE)','bond-only spread bp (Derived)','parity at issue (per 100)','share price 2026-10-02','parity 2026-10-02 (per 100)','size $M','source']); hdr(cr,cr.max_row)
NT=[('NBIS 0.50% 2030','2026-08-19',0.005,1.10,cm['notes']['NBIS 0.50% 2030']['T'],223.90,313.46,3.1902,cm['notes']['NBIS 0.50% 2030']['r'],242.81,3450,'https://www.sec.gov/Archives/edgar/data/1513845/000110465926098924/tm2623617d1_ex99-1.htm ; https://ccir.io/credit/issuer/nebius'),
    ('NBIS 4.50% 2034','2026-08-19',0.045,1.25,cm['notes']['NBIS 4.50% 2034']['T'],223.90,324.65,3.0802,cm['notes']['NBIS 4.50% 2034']['r'],242.81,2300,'same'),
    ('CRWV 2.875% 2033','2026-09-17',0.02875,1.00,cm['notes']['CRWV 2.875% 2033']['T'],79.88,97.85,10.2194,cm['notes']['CRWV 2.875% 2033']['r'],89.62,4200,'/workspace/nebius/comps_raw/crwv_0922__ex991pricing.htm (CRWV 8-K ex99.1, 2026-09-18); size incl. $500M option per parts/comps.csv')]
nt0=cr.max_row+1
for x in NT:
    r=cr.max_row+1
    cr.append([x[0],D(x[1]),x[2],x[3],round(x[4],3),x[5],x[6],f'=G{r}/F{r}-1',f'=G{r}*D{r}',x[7],round(x[8],5),f'=RATE(ROUND(E{r}*2,0),C{r}*100/2,-100,D{r}*100)*2',f'=(L{r}-K{r})*10000',f'=J{r}/10*F{r}',x[9],f'=J{r}/10*O{r}',x[10],x[11]])
    for c,fm in [(2,'yyyy-mm-dd'),(3,'0.000%'),(4,'0%'),(6,'$0.00'),(7,'$0.00'),(8,'0.0%'),(9,'$0.00'),(11,'0.00%'),(12,'0.00%'),(13,'0'),(14,'0.0'),(15,'$0.00'),(16,'0.0')]: cr.cell(row=r,column=c).number_format=fm
cr.append(['Note: CRWV 2033 bond-only yield is BELOW Treasuries. Its buyers are paying for the option. NBIS 2034 needs a 4.5% coupon plus 125% accretion (bond-only ~UST+259bp) even before any option value. That is bond-like economics on an equity-linked note.'])
cr.append([None]); cr.append(['C. Model-implied credit spread (bp) at a given long-dated vol: solve spread so model value = 100 at pricing (Estimated, code/credit_model.py)']); cr.cell(row=cr.max_row,column=1).font=BOLD
cr.append(['note','vol 30%','vol 40%','vol 50%','vol 60%','vol 70%']); hdr(cr,cr.max_row)
for k in ['NBIS 0.50% 2030','NBIS 4.50% 2034','CRWV 2.875% 2033']:
    d=cm['notes'][k]['implied_spread_bp_by_vol']; cr.append([k]+[d[v] for v in ('30','40','50','60','70')])
cr.append(['Model-implied vol (%) at a given credit spread (Estimated)','spr 200bp','spr 400bp','spr 600bp','spr 836bp (= CRWV 9.625% 2032 G-spread)']); hdr(cr,cr.max_row)
for k in ['NBIS 0.50% 2030','NBIS 4.50% 2034','CRWV 2.875% 2033']:
    d=cm['notes'][k]['implied_vol_by_spread']; cr.append([k]+[('no solution' if d[s]>=199 else d[s]) for s in ('200','400','600','836')])
cr.append(['Calibration (Estimated): at CRWV\'s own straight-bond spread (836bp) its new convert implies 55.0% vol, vs CRWV 30-day IV ~70-72% (ratio ~0.77). Applying the same ratio to NBIS 30-day IV 74.6% gives ~57-58% long-dated vol. At that vol, NBIS converts imply ~690-740bp (2030s) and ~950-1,000bp (2034s). That is CRWV-like or wider, despite NBIS having far lower leverage.'])
cr.append([None]); cr.append(['D. Mark-to-model, 2026-10-02 (Estimated; per 100 original; NOT a trading price)']); cr.cell(row=cr.max_row,column=1).font=BOLD
keys=list(cm['notes']['NBIS 0.50% 2030']['mtm_2026-10-02'].keys())
cr.append(['note']+keys+['parity']); hdr(cr,cr.max_row)
for k in ['NBIS 0.50% 2030','NBIS 4.50% 2034','CRWV 2.875% 2033']:
    cr.append([k]+[cm['notes'][k]['mtm_2026-10-02'][q] for q in keys]+[cm['notes'][k]['parity_2026-10-02']])
cr.append([None]); cr.append(['E. Exchange-implied level of NBIS legacy June-2025 notes (Derived)']); cr.cell(row=cr.max_row,column=1).font=BOLD
e=cr.max_row+1
for a,v,n in [('original principal exchanged $M',800,'6-K ex99.1 2026-08-24'),('shares delivered (M)',15.8,'same (approx.)'),('NBIS close 2026-08-19 $',223.90,'pricing 6-K'),('conversion price $',51.45,'June-2025 notes'),
              ('value delivered $M','=B{1}*B{2}',''),('value per 100 original','=B{4}/B{0}*100','Derived'),('parity per 100 = 100/conv x price','=100/B{3}*B{2}','Derived'),('premium over parity','=B{5}/B{6}-1','Derived')]:
    r=cr.max_row+1; ix=[e+i for i in range(8)]
    cr.append([a,v if not isinstance(v,str) else v.format(*ix),n])
cr.cell(row=e+7,column=2).number_format='0.0%'
cr.append([None]); cr.append(['F. Equity-implied (Merton) 5-yr credit spread vs. credit-market spread (Estimated)']); cr.cell(row=cr.max_row,column=1).font=BOLD
cr.append(['issuer','equity $M','equity vol','debt face $M','asset value $M (solved)','asset vol (solved)','risk-neutral 5y PD','Merton spread bp','spread if equity vol 60%','spread if 90%','spread if 110%','credit-market reference','source']); hdr(cr,cr.max_row)
for k,ref in [('CRWV','HY bonds G-spread 651-838bp (section A)'),('NBIS','convert-implied 347-856bp at 40-50% vol (section C); secured facility SOFR+250 (IG-offtaker collateral)')]:
    m=cm['merton'][k]; cr.append([k,m['E'],m['sE'],m['D'],round(m['V']),round(m['sV'],3),round(m['PD_rn'],3),round(m['spread_bp']),round(m['spread_bp_sE60']),round(m['spread_bp_sE90']),round(m['spread_bp_sE110']),ref,m['src']])
cr.append(['Merton caveats: one zero-coupon debt claim at 5y, 30-day IV used as asset-vol input (overstates long-dated vol), no recovery modelling. Use the gap vs credit spreads directionally only.'])
cr.append([None]); cr.append(['G. Other credit reference points']); cr.cell(row=cr.max_row,column=1).font=BOLD
for a in ['Nebius $775M senior secured GPU facility, 1M SOFR + 2.50%, due 2030-10-31, collateral = GPUs plus IG-customer contracted cash flows (Reported): https://ccir.io/credit/issuer/nebius',
          'Nebius booked effective rates on 2025 converts: 2029 7.06%, 2030 4.15%, 2031 6.87%, 2032 4.88% (FY2025 20-F, via https://r40.io/news/nbis-converts-accretion-effective-rate/); Mar-26 notes 4.98% / 5.20% (ccir.io)',
          'UST par curve 2026-10-02: 2y 4.83%, 3y 4.96%, 5y 5.06%, 7y 5.17%, 10y 5.28% (data/ust_2026.csv from home.treasury.gov). On 2026-08-19: 3y 4.25%, 5y 4.35%, 7y 4.48%. Rates are ~70bp higher since NBIS priced, which lowers the bond floor.',
          'NBIS 30-day IV 74.6% (https://www.alphaquery.com/stock/NBIS/volatility-option-statistics/30-day/iv-mean, 2026-10-02/03); CRWV IV30 ~70-72% (https://opti-view.com/underlying/CRWV/implied-volatility, 2026-10-02)']:
    cr.append([a])
widths(cr,[44,13,10,12,12,12,12,12,12,12,12,12,12,12,12,12,10,60])
# ---------- S1_Reconcile (GPU Market teammate data, reused per parent steering)
import json as _json
rc=wb.create_sheet('S1_Reconcile')
rc.append(['Reconciliation with GPU Market teammate outputs (reused, not re-scraped). Sources: /workspace/nebius/parts/time_to_deploy.xlsx (Lag_Ranking, Launch_Premium_Decay) and /workspace/nebius/market/scrapers/time_to_deploy/out/lag_summary.json ; /workspace/nebius/market/price_history_monthly.csv ; /workspace/nebius/market/gpu_price_history.xlsx (built 2026-10-04 18:16 ET)'])
rc['A1'].font=Font(bold=True)
rc.append(['chip','teammate: Nebius lag vs first provider (days)','teammate: first provider','this study: Nebius months behind first neocloud (S1_Timeline)','this study in days (Derived = months*30.44)','difference days','note'])
hdr(rc,2)
lag=_json.load(open('/workspace/nebius/market/scrapers/time_to_deploy/out/lag_summary.json'))
mine={'H100':7.3916,'H200':2.5953,'B200 (HGX)':2.7924,'GB200 NVL72':4.1721}
for row in lag:
    r=rc.max_row+1
    m=mine.get(row[0])
    rc.append([row[0],row[8],row[6],m,(f'=D{r}*30.44' if m is not None else 'n/a'),(f'=B{r}-E{r}' if m is not None else 'n/a'),
      'B200: teammate dates Nebius by first public price (2025-07-25) vs GCP; this study uses the 2025-06-11 announcement' if row[0].startswith('B200') else ''])
rc.append([None]); rc.append(['A100 on-demand $/GPU-hr from teammate monthly pivot (last observed headline price in month; blank = no capture). Used for A100 years 3-6 rental precedent.']); rc.cell(row=rc.max_row,column=1).font=Font(bold=True)
import csv as _csv
mrows=list(_csv.reader(open('/workspace/nebius/market/price_history_monthly.csv')))
cols=[i for i,c in enumerate(mrows[0]) if 'A100' in c and 'on-demand' in c]
rc.append(['month']+[mrows[0][i] for i in cols]); hdr(rc,rc.max_row)
for mr in mrows[1:]:
    vals=[float(mr[i]) if mr[i] else None for i in cols]
    if any(v is not None for v in vals): rc.append([mr[0]]+vals)
rc.append([None]); rc.append(['Read: teammate lag table matches this study within a few days for H100/H200/GB200. Teammate A100 series (Lambda $1.79 flat Jan-24..Jan-26 then $2.79; Crusoe 80GB $1.65-2.00; RunPod $1.89->$1.59; CoreWeave $2.70 flat) confirms A100 at age 4-6y still rents at $1.5-2.8, above the ~$0.40 cash floor. Note teammate captured Nebius H200 at $4.85 on 2024-10-06 (pre-GA), which this study added to S1_PriceSeries. The teammate Launch_Premium_Decay sheet lists Nebius H200 launch as $3.50 on 2024-11-15; this study treats the pre-GA $4.85 list (2024-10-06) as the launch price, so the H200 premium was cut ~28% before GA. Per parent steering, part 1 lag ranking defers to the teammate study; this workbook adds only half-life and the head-start NPV.'])
widths(rc,[16,22,40,22,18,14,60])
# ---------- Sources
so=wb.create_sheet('Sources'); so.append(['source','what','retrieved (ET)']); hdr(so,1)
for a,b in [('https://web.archive.org (captures listed per row in S1_PriceSeries)','Provider pricing pages: Nebius, CoreWeave, Lambda, AWS Capacity Blocks'),
 ('/workspace/nebius/raw/aws_od_use1.json','AWS on-demand price list (GPU Market teammate pull)'),('https://ccir.io/research/rental-rates-by-generation','9-year posted-rate panel, V100/A100/H100'),
 ('https://www.silicondata.com/products/silicon-index?gpu=h100','SDH100RT index'),('https://www.silicondata.com/blog/h100-rental-prices-by-region-q3-2026','hyperscaler premium compression'),
 ('https://newsletter.semianalysis.com/p/the-great-gpu-shortage-rental-capacity','H100 1-yr contract $1.70 (Oct-25) to $2.35 (Mar-26)'),
 ('https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf','Nebius Q2-26 letter: >30% higher older-gen pricing vs Q1'),
 ('https://www.reuters.com/technology/nebius-hikes-ai-cloud-prices-again-demand-computing-power-soars-2026-09-17/','Oct-1 price hike'),
 ('https://knowyourinvestments.substack.com/p/nebius-watch-gpu-rent-price-hikes','Jun-1 hike $2.95 to $3.85'),
 ('https://hashrateindex.com/blog/used-gpu-market-pricing-deprecation-secondary-ai/','used GPU market, 2026-04-08'),('https://clustermelt.com/market-intelligence/2026-09','used GPU reads 2026-09-01'),
 ('https://www.mercatus-ai.com/blog/h100-resale-value','H100 residual bands 2026-06-23'),('https://bondterminal.com/issuers/coreweave-inc','CRWV bond levels 2026-10-01/02'),('https://ccir.io/credit/issuer/nebius','NBIS capital stack'),
 ('/workspace/nebius/src/20f.txt','Nebius 20-F useful lives (GPUs 4y; servers/network 4 to 5y from 2026-01-01; -$167.6M FY26 D&A)')]:
    so.append([a,b,'2026-10-04'])
widths(so,[90,70,14])
wb.save(B+'speed_study.xlsx'); print('saved')
