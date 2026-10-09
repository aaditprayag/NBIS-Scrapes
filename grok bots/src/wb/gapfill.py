"""Fill months missing from the exact-URL CDX pass using query-string/alternate variants (prefix match)."""
import sys,os,json,time
sys.path.insert(0,os.path.dirname(__file__))
from cdx import cdx
from fetch import fetch, RAW
label=sys.argv[1]; prefixes=sys.argv[2:]
log=f'{RAW}/fetchlog_{label}.jsonl'
have={json.loads(l)['ts'][:6] for l in open(log) if json.loads(l).get('ts') and json.loads(l).get('status') in ('ok','cached')} if os.path.exists(log) else set()
want=[f'{y}{m:02d}' for y in (2024,2025,2026) for m in range(1,13) if f'{y}{m:02d}'<='202610']
cands={}
for p in prefixes:
    rows=cdx(p,frm='2024',collapse='timestamp:6',match='prefix') or []
    for r in rows:
        mo=r[1][:6]
        if mo in want and mo not in have and r[3].startswith('text/html'):
            cands.setdefault(mo,[]).append(r)
    time.sleep(1)
with open(log,'a') as L:
    for mo in sorted(cands):
        r=sorted(cands[mo],key=lambda r:len(r[2]))[0]
        fn,st=fetch(r[1],r[2])
        L.write(json.dumps({'url':prefixes[0],'orig':r[2],'ts':r[1],'mime':r[3],'len':r[6] if len(r)>6 else None,'file':fn,'status':st,'gapfill':True})+'\n'); L.flush()
        print(mo,r[2][:100],st)
print('missing after gapfill:',[m for m in want if m not in have and m not in cands])
