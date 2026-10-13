"""Hyperscaler-entry test: neocloud list price 45 days before each event vs. 90 days after (last Wayback capture on/before each date)."""
import csv,datetime as dt,sys
sys.path.insert(0,'/workspace/nebius/skeptic/speed/code')
from curated import EVENTS,PRIOR
D=lambda s:dt.date.fromisoformat(s)
rows=list(csv.DictReader(open('/workspace/nebius/skeptic/speed/data/price_series.csv')))
# Lambda 2023 price column applies from 8x HGX on-demand launch 2023-08-02 (Lambda blog) through 2023-12-31
for r in rows:
    if r['provider']=='Lambda' and r['date']=='2023-12-31': r['date']='2023-08-02'
ser={}
for r in rows: ser.setdefault((r['provider'],r['chip']),[]).append((D(r['date']),float(r['usd_per_gpu_hr']),r['source_url']))
for k in ser: ser[k].sort()
def at(p,c,d,maxgap=200):
    s=[x for x in ser.get((p,c),[]) if x[0]<=d]
    if not s or (d-s[-1][0]).days>maxgap: return None
    return s[-1]
out=[]
for eid,ed,name,chip,typ in EVENTS:
    e=D(ed)
    for c in PRIOR[chip]:
        for p in ['Nebius','Lambda','CoreWeave']:
            b=at(p,c,e-dt.timedelta(days=45)); a=at(p,c,e+dt.timedelta(days=90))
            if b is None and a is None: continue
            out.append(dict(event=eid,event_date=ed,event_name=name,neocloud=p,chip=c,same_gen='same chip' if c in (chip,'GB200 NVL72' if chip=='GB200' else chip) else 'prior gen',
              before_date=b[0].isoformat() if b else 'not observed',before=b[1] if b else None,after_date=a[0].isoformat() if a else 'not observed',after=a[1] if a else None,
              src_before=b[2] if b else '',src_after=a[2] if a else '',
              note='BASIS CHANGE (CoreWeave component rate card -> instance price), not a price move; excluded' if (p=='CoreWeave' and b and a and b[0].year<2025<=a[0].year and c=='H100') else ''))
csv.DictWriter(open('/workspace/nebius/skeptic/speed/data/hs_entry_test.csv','w',newline=''),list(out[0])).writeheader() if out else None
with open('/workspace/nebius/skeptic/speed/data/hs_entry_test.csv','w',newline='') as f:
    w=csv.DictWriter(f,list(out[0])); w.writeheader(); w.writerows(out)
for o in out:
    ch=(o['after']/o['before']-1) if (o['before'] and o['after']) else None
    print(o['event'],o['event_date'],o['neocloud'],o['chip'],o['same_gen'],o['before'],o['after'],f'{ch:+.1%}' if ch is not None else 'n/a')
