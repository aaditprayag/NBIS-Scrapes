import requests, time, json, sys, os, re, hashlib
sys.path.insert(0,os.path.dirname(__file__))
from cdx import cdx, S
RAW='/workspace/nebius/wayback_raw'
def safe(u): return re.sub(r'[^A-Za-z0-9]+','_',u).strip('_')
def fetch(ts,url):
    d=f'{RAW}/{safe(url)}'; os.makedirs(d,exist_ok=True)
    fn=f'{d}/{ts}.html'
    if os.path.exists(fn) and os.path.getsize(fn)>0: return fn,'cached'
    w=f'http://web.archive.org/web/{ts}id_/{url}'
    for i in range(5):
        try:
            r=S.get(w,timeout=90)
            if r.status_code==200:
                open(fn,'wb').write(r.content); time.sleep(1); return fn,'ok'
            st=r.status_code
            if st in (404,403): return None,f'http {st}'
        except Exception as e: st=str(e)[:80]
        time.sleep(4*(i+1))
    return None,f'fail {st}'
if __name__=='__main__':
    # args: label url [url...]
    label=sys.argv[1]
    log=open(f'{RAW}/fetchlog_{label}.jsonl','a')
    for u in sys.argv[2:]:
        rows=cdx(u, frm=os.environ.get('WB_FROM') or ('2024' if 'nebius.ai' not in u else '2023'))
        if rows is None:
            log.write(json.dumps({'url':u,'status':'cdx_fail'})+'\n'); log.flush(); continue
        if not rows:
            log.write(json.dumps({'url':u,'status':'no_snapshots'})+'\n'); log.flush(); continue
        for r in rows:
            ts,orig,mime=r[1],r[2],r[3]
            fn,st=fetch(ts,orig)
            log.write(json.dumps({'url':u,'orig':orig,'ts':ts,'mime':mime,'len':r[6] if len(r)>6 else None,'file':fn,'status':st})+'\n'); log.flush()
        time.sleep(1)
