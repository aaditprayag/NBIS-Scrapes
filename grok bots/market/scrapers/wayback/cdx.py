import requests, time, json, sys
S=requests.Session(); S.headers['User-Agent']='Mozilla/5.0 (research; aadit nyu equity project)'
def cdx(url, frm='2023', to='2026', collapse='timestamp:6', match=None, extra=None):
    p={'url':url,'output':'json','filter':'statuscode:200','from':frm,'to':to}
    if collapse: p['collapse']=collapse
    if match: p['matchType']=match
    if extra: p.update(extra)
    for i in range(6):
        try:
            r=S.get('http://web.archive.org/cdx/search/cdx',params=p,timeout=90)
            if r.status_code==200:
                j=r.json() if r.text.strip() else []
                return j[1:] if j else []
            print('status',r.status_code,file=sys.stderr)
        except Exception as e: print('err',e,file=sys.stderr)
        time.sleep(3*(i+1))
    return None
if __name__=='__main__':
    for u in sys.argv[1:]:
        rows=cdx(u)
        print(u, None if rows is None else len(rows), [r[1][:6] for r in rows] if rows else '')
        time.sleep(1)
