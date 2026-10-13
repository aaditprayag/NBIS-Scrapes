import csv,collections,sys
rows=list(csv.DictReader(open('/workspace/nebius/market/scrapers/time_to_deploy/out/price_obs.csv')))
by=collections.defaultdict(list)
for r in rows:
    if 'before' in r['basis'] or r['basis'].startswith('Announced'): continue
    by[(r['provider'],r['chip'])].append((r['date'],r['price_gpu_hr']))
for k in sorted(by):
    v=sorted(by[k]); out=[];last=None
    for d,p in v:
        if p!=last: out.append(f'{d}:{p or "n/a"}'); last=p
    print(k,len(v),'first',v[0][0],'last',v[-1][0],' | '.join(out))
