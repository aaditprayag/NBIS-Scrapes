"""Full (digest-collapsed) Wayback CDX listing for pricing pages -> wayback_raw/cdx_full_<label>.json
Usage: python cdx_full.py label url [url ...]
Each row: [urlkey,timestamp,original,mimetype,statuscode,digest,length]. collapse=digest keeps only rows whose content hash changed
vs the previous capture, so it is a cheap change-candidate list (digest also changes on trivial HTML noise)."""
import sys, json, time, os, requests
RAW='/workspace/nebius/wayback_raw'
S=requests.Session(); S.headers['User-Agent']='Mozilla/5.0 (research; aadit nyu equity project)'
def cdx(url, frm='2023', to='2026'):
    p={'url':url,'output':'json','filter':'statuscode:200','from':frm,'to':to,'collapse':'digest'}
    for i in range(6):
        try:
            r=S.get('http://web.archive.org/cdx/search/cdx',params=p,timeout=120)
            if r.status_code==200:
                j=r.json() if r.text.strip() else []
                return j[1:] if j else []
        except Exception as e: print('err',e,file=sys.stderr)
        time.sleep(3*(i+1))
    return None
if __name__=='__main__':
    label=sys.argv[1]; out={}
    for u in sys.argv[2:]:
        out[u]=cdx(u); print(u, None if out[u] is None else len(out[u]), flush=True); time.sleep(1.2)
    json.dump(out,open(f'{RAW}/cdx_full_{label}.json','w'))
