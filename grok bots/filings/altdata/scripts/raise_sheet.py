import os,sys
sys.path.insert(0,os.path.dirname(__file__))
from openpyxl.styles import Font,Alignment
QS=['Q3-24','Q4-24','Q1-25','Q2-25','Q3-25','Q4-25','Q1-26','Q2-26']
# (id, date, quarter, instrument, gross $M, src_key, detail)
RAISES=[
('F0','2024-07-15','Q3-24','Starting cash after Yandex divestment (not a raise; funding pool)',2288.2,'Q3_24_PR','Cash & equivalents at Sep-30-2024 $2,288.2M; Q2-24 PR: "the Group has USD 2.5 billion cash and cash equivalents and no debt" after Jul-15-2024 closing.'),
('F1','2024-12-02','Q4-24','Equity private placement (Accel, NVIDIA, Orbis)',700.0,'DEC24_PIPE','"USD 700 million private placement financing"'),
('F2','2025-06-02','Q2-25','Convertible notes 2029 (2.00%) + 2031 (3.00%)',1000.0,'JUN25_CB','$500M + $500M'),
('F3','2025-09-15','Q3-25','Public equity offering 12.43M Class A @ $92.50 (incl. greenshoe Sep-22)',1150.0,'SEP25_SHOE','"aggregate gross proceeds from the Class A offering to approximately $1.15 billion"'),
('F4','2025-09-15','Q3-25','Convertible notes 2030 (1.00%) + 2032 (2.75%)',3160.0,'SEP25_CLOSE','"total aggregate original principal amount of the Notes is approximately $3.16 billion"'),
('F5','2025-Q4','Q4-25','Customer prepayments (deferred revenue increase, CFO line)',1566.5,'Q4_25_PR','Microsoft/Meta strategic customer advances; FY25 20-F: cash advances $982.5M'),
('F6','2026-03-11','Q1-26','NVIDIA pre-funded warrants (21.07M shares)',2000.0,'NVDA','"NVIDIA will invest $2 billion in Nebius"'),
('F7','2026-03-20','Q1-26','Convertible notes 2031 (1.25%) + 2033 (2.625%)',4337.5,'MAR26_CLOSE','"total aggregate original principal amount of the Notes to date is $4.3375 billion"'),
('F8','2026-Q1','Q1-26','Customer prepayments (deferred revenue increase, CFO line)',3198.0,'Q1_26_OFR','"advances received under customer agreements of $3,198.0 million"'),
('F9','2026-05/06','Q2-26','ATM equity (sale of treasury shares), May-Jun 2026',2846.7,'Q2_26_FS','CF "Proceeds from sale of treasury shares" 2,846.7; 12.3M of 25M ATM shares remained at Jun-30'),
('F10','2026-Q2','Q2-26','Customer prepayments (deferred revenue increase, CFO line)',1197.0,'Q2_26_PR',''),
('F11','2026-07-17','Q3-26','Senior secured GPU-backed facility (SOFR+2.50%, matures Oct-31-2030)',775.0,'JUL26_SECURED','"first senior secured debt facility for approximately $775 million"'),
('F12','2026-08-24','Q3-26','Convertible notes 2030 (0.50%) $3.45B + 2034 (4.50%) $2.30B',5750.0,'AUG26_CLOSE','Plus exchange of $800M existing 2029/2031 notes into ~15.8M shares (non-cash)'),
]
CAPEX={'Q3-24':167.0,'Q4-24':417.5,'Q1-25':543.9,'Q2-25':510.6,'Q3-25':955.5,'Q4-25':2056.0,'Q1-26':2472.9,'Q2-26':5657.4}
CAPEX_SRC={'Q3-24':'Q3_24_PR','Q4-24':'Q4_25_PR','Q1-25':'Q1_26_PR','Q2-25':'Q2_26_PR','Q3-25':'Q3_25_PR','Q4-25':'Q4_25_PR','Q1-26':'Q1_26_PR','Q2-26':'Q2_26_PR'}
PPE={'Q2-24':328.3,'Q3-24':475.3,'Q4-24':846.7,'Q1-25':1334.1,'Q2-25':1789.4,'Q3-25':3314.4,'Q4-25':5553.3,'Q1-26':7131.7,'Q2-26':13045.2}
ARR={'Q3-24':121,'Q4-24':90,'Q1-25':249,'Q2-25':430,'Q3-25':551,'Q4-25':1250,'Q1-26':1900,'Q2-26':3000}
ARR_SRC={'Q3-24':'OCT24_PRES','Q4-24':'Q4_24_PR','Q1-25':'Q1_25_SHL','Q2-25':'Q2_25_SHL','Q3-25':'Q3_25_SHL','Q4-25':'Q4_25_SHL','Q1-26':'Q1_26_SHL','Q2-26':'Q2_26_SHL'}
CAP=[('2024-11','Paris H200 cluster in operation','Q4_24_PR'),('2025-03','Iceland colocation opened (~10 MW)','Q1_25_SHL'),('2025-04','Kansas City colocation online (5 MW phase 1)','Q1_25_SHL'),
('2025-03-31','~45 MW connected capacity; ~30,000 GPUs deployed','FY24_20F'),('2025-11','Vineland NJ: Microsoft tranche 1 delivered; Meta tranche 1 delivered','FY25_20F'),
('2025-12-31','~170 MW active power (vs 100 MW target)','Q4_25_SHL'),('2026-02','Microsoft tranche 2 + Meta tranche 2 delivered','FY25_20F'),
('2026-07','"latest planned capacity tranche" delivered to Microsoft','JUL26_SECURED'),('2026-12-31 (target)','800 MW-1 GW connected; active "throughout the first half of 2027"','CALL_Q2_26')]
def fifo():
    pool=[[r[0],r[2],r[4]] for r in RAISES if r[2] in QS]; out={r[0]:None for r in RAISES}
    i=0
    for q in QS:
        spend=CAPEX[q]
        avail=[p for p in pool if QS.index(p[1])<=QS.index(q) and p[2]>0]
        for p in avail:
            take=min(p[2],spend); p[2]-=take; spend-=take
            if p[2]<=1e-6 and out[p[0]] is None: out[p[0]]=q
            if spend<=0: break
    rem={p[0]:round(p[2],1) for p in pool}
    return out,rem
def lagfit():
    import statistics as st
    d=[ARR[QS[i]]-ARR[QS[i-1]] for i in range(1,len(QS))]
    res=[]
    for k in range(0,4):
        xs=[];ys=[]
        for i in range(1,len(QS)):
            if i-k>=0: xs.append(CAPEX[QS[i-k]]); ys.append(ARR[QS[i]]-ARR[QS[i-1]])
        if len(xs)>=3:
            mx,my=st.mean(xs),st.mean(ys); cov=sum((a-mx)*(b-my) for a,b in zip(xs,ys)); vx=sum((a-mx)**2 for a in xs); vy=sum((b-my)**2 for b in ys)
            res.append((k,len(xs),round(cov/(vx*vy)**.5,3),round(sum(ys)/sum(xs),3)))
    return res
def write(wb,sheet,header,widths,url):
    ws=sheet('Raise_to_Deploy'); ws['A1']='Speed of capital raise -> capex -> capacity live ($M). Method below.'; ws['A1'].font=Font(bold=True)
    out,rem=fifo(); r=3
    header(ws,r,['ID','Date','Quarter','Source of funds','Gross $M','FIFO: quarter fully deployed into capex','Raise->capex lag (quarters)','Unspent at Q2-26 ($M, FIFO)','Source URL','Detail / quote'])
    for fid,d,q,inst,amt,k,det in RAISES:
        r+=1; fq=out.get(fid); lag=(QS.index(fq)-QS.index(q)) if fq and q in QS else None
        vals=[fid,d,q,inst,amt,fq or ('after Q2-26' if q in QS else 'n/a (post Q2-26)'),lag,rem.get(fid),url(k),det]
        for j,v in enumerate(vals,1): ws.cell(row=r,column=j,value=v).alignment=Alignment(wrap_text=True,vertical='top')
    r+=2; header(ws,r,['Quarter']+QS); r+=1
    rows=[('Cash capex (purchases of P&E & intangibles)',CAPEX),('PP&E, net (quarter-end)',PPE),('Change in PP&E, net',{q:round(PPE[q]-PPE[QS[i-1]] if i else PPE[q]-PPE['Q2-24'],1) for i,q in enumerate(QS)}),('ARR, Nebius AI cloud ($M)',ARR),('Change in ARR',{q:(ARR[q]-ARR[QS[i-1]] if i else None) for i,q in enumerate(QS)}),('Cumulative funds raised (incl. starting cash & prepayments)',{q:round(sum(x[4] for x in RAISES if x[2] in QS and QS.index(x[2])<=QS.index(qq)) ,1) for qq in QS for q in [qq]}),('Cumulative cash capex',{q:round(sum(CAPEX[x] for x in QS[:i+1]),1) for i,q in enumerate(QS)})]
    for lab,dct in rows:
        ws.cell(row=r,column=1,value=lab)
        for j,q in enumerate(QS,2): ws.cell(row=r,column=j,value=dct.get(q))
        r+=1
    ws.cell(row=r,column=1,value='Capex source URLs: '+' | '.join(f'{q}: {url(k)}' for q,k in CAPEX_SRC.items())); r+=1
    ws.cell(row=r,column=1,value='ARR source URLs: '+' | '.join(f'{q}: {url(k)}' for q,k in ARR_SRC.items())+'  (Sep-24 = $121M per Oct-18-2024 presentation, preliminary)'); r+=2
    header(ws,r,['Capex(t-k) vs dARR(t): lag k (quarters)','n','Pearson r','Sum dARR / sum capex (ARR $ per capex $)']); r+=1
    for k,n,c,ratio in lagfit():
        for j,v in enumerate([k,n,c,ratio],1): ws.cell(row=r,column=j,value=v)
        r+=1
    r+=1; header(ws,r,['Date','Announced capacity go-live / milestone','Source URL']); r+=1
    for d,t,k in CAP:
        ws.cell(row=r,column=1,value=d); ws.cell(row=r,column=2,value=t); ws.cell(row=r,column=3,value=url(k)); r+=1
    r+=1
    M=['METHOD',
    '1) Funding sources are logged at gross amount in the quarter received: equity, converts, NVIDIA warrants, ATM, secured debt and customer prepayments (deferred-revenue increase in operating cash flow). Starting cash after the Yandex divestment is included as F0 because it funded the 2H24 capex.',
    '2) FIFO allocation: each quarter\'s cash capex is charged to the oldest unspent funds. "Raise->capex lag" = quarters until a source is fully absorbed. This is an upper bound on deployment speed: it ignores opex, M&A ($251.6M H1-26), taxes and the cash buffer. Money is fungible, so read it as aggregate pace, not literal use of proceeds.',
    '3) Capex->capacity lag. Management describes three stages (Nov-11-2025 call): land+power ~1% of capex; data-center build ("connected") ~18-20%; GPUs ~80%. Connected->revenue is "anywhere from 6-12 weeks" (Andrey Korolenko, Nov-11-2025 call). The lag regression correlates quarterly capex(t-k) with ARR change(t), k=0-3. n is only 4-7, so treat it as indicative only.',
    '4) Raise->capacity = raise->capex (FIFO) + capex->capacity (best-fit k), cross-checked against the announced go-live dates above.',
    'CAVEATS: cash capex excludes vendor-deferred capex (see Prepay_Commitments: "Acquired P&E not yet paid for" $973M in H1-26). Q3-26 raises (F11, F12) are after the last reported capex quarter.']
    for m in M:
        ws.cell(row=r,column=1,value=m).alignment=Alignment(wrap_text=True); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10); ws.row_dimensions[r].height=45; r+=1
    widths(ws,[44,12,10,52,11,16,12,14,40,60])
if __name__=='__main__':
    print(fifo()); print(lagfit())
