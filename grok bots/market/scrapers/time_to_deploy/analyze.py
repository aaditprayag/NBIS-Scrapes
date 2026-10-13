"""Analysis layer: lags, launch premium/decay, pricing-page change log. Imported by build_workbook.py."""
import csv, datetime as dt, math, statistics as st, collections, os
from data_curated import MILESTONES, AVAIL, REF, ANN, GENS
BASE=os.path.dirname(os.path.abspath(__file__))
TODAY=dt.date(2026,10,4)
D=lambda s: dt.date.fromisoformat(s) if s else None
CHIP2GEN={'H100':'H100','H200':'H200','B200':'B200 (HGX)','GB200 NVL72':'GB200 NVL72','B300':'B300 / GB300 NVL72','GB300 NVL72':'B300 / GB300 NVL72'}
ONDEMAND_PEERS=['Nebius','CoreWeave','Lambda','Crusoe','Together AI']
PRED={'H200':['H100'],'B200':['H200','H100'],'B300':['B200'],'GB200 NVL72':['B200','H200'],'GB300 NVL72':['GB200 NVL72','B300']}
def load_obs():
    rows=list(csv.DictReader(open(f'{BASE}/out/price_obs.csv')))
    for r in rows: r['d']=D(r['date']); r['p']=float(r['price_gpu_hr']) if r['price_gpu_hr'] else None
    return rows
def series(rows):
    """(provider,chip) -> sorted list of (date, price or None, archive_url, basis); current-price rows only."""
    s=collections.defaultdict(dict)
    for r in rows:
        if 'before' in r['basis'] or r['basis'].startswith('Announced'): continue
        k=(r['provider'],r['chip']); prev=s[k].get(r['d'])
        if prev is None or (prev[0] is None and r['p'] is not None): s[k][r['d']]=(r['p'],r['archive_url'],r['basis'])
    return {k:sorted((d,)+v for d,v in s[k].items()) for k in s}
def asof(ser,d):
    """last observed non-null price on/before d; None if d after last observation+45d or before first."""
    best=None
    for o in ser:
        if o[0]<=d and o[1] is not None: best=o
    if not ser or d>ser[-1][0]+dt.timedelta(days=45): return None
    return best
def addm(d,m):
    y=d.year+(d.month-1+m)//12; mo=(d.month-1+m)%12+1
    import calendar; return dt.date(y,mo,min(d.day,calendar.monthrange(y,mo)[1]))
def lag_table():
    out=[]
    for g in GENS:
        ref=D(REF[g]); rows=[]
        for a in AVAIL:
            if a[1]!=g: continue
            ga=D(a[7])
            lag=(ga-ref).days if ga else None
            cens=None if ga else (TODAY-ref).days
            rows.append(dict(gen=g,provider=a[0],ref=ref,ann=D(ANN[g]),ga=ga,lag=lag,lag_from_ann=(ga-D(ANN[g])).days if ga else None,censored_min=cens,ga_tag=a[8],ga_src=a[10],ga_def=a[9]))
        ranked=sorted([r for r in rows if r['lag'] is not None],key=lambda r:r['lag'])
        for i,r in enumerate(ranked): r['rank']=i+1
        n=len(ranked)
        for r in rows:
            r['n_ranked']=n
            if r['lag'] is None: r['rank']=None
        peers=[r['lag'] for r in ranked if r['provider']!='Nebius']
        neb=[r for r in rows if r['provider']=='Nebius']
        med=st.median(peers) if peers else None
        for r in rows:
            r['peer_median_ex_nebius']=med; r['fastest']=ranked[0]['provider'] if ranked else None
            r['nebius_minus_peer_median']=(neb[0]['lag']-med) if (neb and neb[0]['lag'] is not None and med is not None) else None
        out+=sorted(rows,key=lambda r:(r['lag'] is None, r['lag'] if r['lag'] is not None else 0, r['provider']))
    return out
def fit_monthly(ser,start,end):
    """sample step function at start, start+1m,... <=end; OLS ln(p) on month index. returns (pct_per_month, n_points, points)"""
    pts=[];m=0
    while True:
        d=addm(start,m)
        if d>end: break
        o=asof(ser,d)
        if o: pts.append((m,o[1]))
        m+=1
    if len(pts)<3: return None,len(pts),pts
    xs=[x for x,_ in pts]; ys=[math.log(y) for _,y in pts]; mx=st.mean(xs); my=st.mean(ys)
    b=sum((x-mx)*(y-my) for x,y in zip(xs,ys))/sum((x-mx)**2 for x in xs)
    return (math.exp(b)-1)*100,len(pts),pts
def premium_decay(S):
    out=[]
    for (prov,chip),ser in sorted(S.items()):
        if prov=='AWS' or chip=='H100 PCIe' and prov!='Lambda': pass
        priced=[o for o in ser if o[1] is not None]
        if not priced: continue
        L=priced[0]; ld=L[0]; last=ser[-1][0]
        ga=[a[7] for a in AVAIL if a[0]==prov and CHIP2GEN.get(chip.replace(' PCIe',''))==a[1]]
        ga=D(ga[0]) if ga and ga[0] else None
        row=dict(provider=prov,chip=chip,provider_ga=ga,first_price_gap=(ld-ga).days if ga else None,is_launch=bool(ga and (ld-ga).days<=60),launch_obs=ld,launch_price=L[1],launch_url=L[2],launch_basis=L[3],last_obs=last)
        for m in (3,6,12):
            o=asof(ser,addm(ld,m)); row[f'p{m}']=o[1] if o and addm(ld,m)<=last else None
            row[f'u{m}']=o[2] if o and addm(ld,m)<=last else None
            row[f'chg{m}']=(row[f'p{m}']/L[1]-1)*100 if row[f'p{m}'] else None
        # predecessor premium same provider same date
        row['pred']=None;row['pred_price']=None;row['prem_pred']=None
        for pc in PRED.get(chip,[]):
            if (prov,pc) in S:
                o=asof(S[(prov,pc)],ld)
                if o: row['pred']=pc;row['pred_price']=o[1];row['prem_pred']=(L[1]/o[1]-1)*100;break
        # peer median same chip same date
        peers=[]
        if prov in ONDEMAND_PEERS:
            for p2 in ONDEMAND_PEERS:
                if p2==prov or (p2,chip) not in S: continue
                o=asof(S[(p2,chip)],ld)
                if o: peers.append((p2,o[1]))
        row['peers']=peers; row['peer_med']=st.median([p for _,p in peers]) if peers else None
        row['prem_peer']=(L[1]/row['peer_med']-1)*100 if row['peer_med'] else None
        r12,n12,_=fit_monthly(ser,ld,min(addm(ld,12),last)); rall,nall,_=fit_monthly(ser,ld,last)
        row.update(fit12=r12,n12=n12,fitall=rall,nall=nall)
        out.append(row)
    return out
def change_log(rows):
    """Within each page_dir, compare consecutive captures; events merged across dirs per provider."""
    by=collections.defaultdict(lambda: collections.defaultdict(dict))
    for r in rows:
        if r['provider'] not in ('Nebius','CoreWeave','Lambda'): continue
        if 'before' in r['basis']: continue
        key=r['chip']+(' [announced future price]' if r['basis'].startswith('Announced') else '')
        by[(r['provider'],r['page_dir'])][r['ts']][key]=(r['p'],r['archive_url'])
    # include captures with zero parsed chips so removals are visible
    ev=[]
    for (prov,pdir),snaps in by.items():
        ts=sorted(snaps)
        for a,b in zip(ts,ts[1:]):
            A,B=snaps[a],snaps[b]
            for c in sorted(set(A)|set(B)):
                if c in B and c not in A: t='future price announced' if 'announced' in c else 'chip added'; old=None; new=B[c][0]
                elif c in A and c not in B: t='chip removed (or not parsed)'; old=A[c][0]; new=None
                else:
                    old,new=A[c][0],B[c][0]
                    if old==new: continue
                    t='price published' if old is None else ('price withdrawn (Contact sales)' if new is None else 'price change')
                ev.append(dict(provider=prov,page=pdir,chip=c,type=t,old=old,new=new,prev_ts=a,ts=b,
                               prev_url=(A.get(c) or list(A.values())[0])[1] if A else '',url=(B.get(c) or list(B.values())[0])[1] if B else ''))
    # merge duplicates across page dirs: same provider/chip/type/new within 60 days -> keep earliest
    ev.sort(key=lambda e:(e['provider'],e['chip'],e['type'],str(e['new']),e['ts']))
    keep=[]
    for e in ev:
        if keep and all(keep[-1][k]==e[k] for k in ('provider','chip','type')) and keep[-1]['new']==e['new'] and \
           (D(e['ts'][:4]+'-'+e['ts'][4:6]+'-'+e['ts'][6:8])-D(keep[-1]['ts'][:4]+'-'+keep[-1]['ts'][4:6]+'-'+keep[-1]['ts'][6:8])).days<60:
            keep[-1].setdefault('also',[]).append(e['page']); continue
        keep.append(e)
    # drop: announced-row disappearance (became effective); within-dir add/price events (replaced by provider timeline below)
    keep=[e for e in keep if not ('announced' in e['chip'] and e['type'].startswith('chip removed')) and
          (e['type'].startswith('chip removed') or e['type']=='future price announced')]
    # provider-level timeline across page dirs (handles domain/page moves)
    S=series(rows)
    for (prov,chip),ser in S.items():
        if prov not in ('Nebius','CoreWeave','Lambda'): continue
        f=ser[0]; ts=f[2].split('/web/')[1].split('id_')[0] if '/web/' in f[2] else ''
        keep.append(dict(provider=prov,page='(provider timeline)',chip=chip,type='chip added',old=None,new=f[1],prev_ts=prev_capture(rows,prov,f[0]),ts=ts,prev_url='',url=f[2]))
        for a,b in zip(ser,ser[1:]):
            if a[1]==b[1]: continue
            t='price published' if a[1] is None else ('price withdrawn (Contact sales)' if b[1] is None else 'price change')
            ts=b[2].split('/web/')[1].split('id_')[0]; pts=a[2].split('/web/')[1].split('id_')[0]
            keep.append(dict(provider=prov,page='(provider timeline)',chip=chip,type=t,old=a[1],new=b[1],prev_ts=pts,ts=ts,prev_url=a[2],url=b[2]))
    for e in keep:
        e['note']=''
        if e['provider']=='CoreWeave' and e['chip']=='H100' and e['old']==4.76: e['note']='DEFINITIONAL: CoreWeave moved from component pricing (GPU-only $4.76) to bundled 8-GPU instance price ($49.24/8) on new /pricing page; not a like-for-like price change'
        if e['provider']=='Nebius' and e['ts'][:8] in ('20260712','20260625'): e['note']='Change bracketed by captures 2026-05-22 (old price, ?utm getdeploying variant) and 2026-06-25 (new price)'
        if 'announced' in e['chip']: e['note']='Page showed new prices "from Oct 1, 2026" alongside current prices'
    keep.sort(key=lambda e:(e['provider'],e['ts'],e['chip']))
    return keep
def prev_capture(rows,prov,d):
    ds=sorted({r['ts'] for r in rows if r['provider']==prov and r['d']<d})
    return ds[-1] if ds else ''

