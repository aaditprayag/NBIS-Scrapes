#!/usr/bin/env python3
"""Fetch archived ATS job-board JSON snapshots (Wayback id_ raw captures listed in raw/wayback/cdx_*.json),
patiently (one request at a time; on 429/TLS reset wait 3 min, up to ~3 h), then parse each into the same
normalized row format as jobs_snapshot.py -> jobs_history.csv (snapshot_ts, company, id, title, department, location, offices).
Run: python3 wayback_jobs.py [fetch|parse|all]"""
import json, os, subprocess, time, csv, sys, glob
D=os.path.dirname(os.path.abspath(__file__)); W=f'{D}/raw/wayback'
SETS={'Nebius':'cdx_gh_nebius_api.json','CoreWeave':'cdx_gh_coreweave_api.json','Lambda':'cdx_ashby_lambda_api.json'}
def fetch():
    deadline=time.time()+3*3600
    for co,f in SETS.items():
        rows=json.load(open(f'{W}/{f}'))[1:]; os.makedirs(f'{W}/jobs_{co}',exist_ok=True)
        for ts,orig,*_ in rows:
            out=f'{W}/jobs_{co}/{ts}.json'
            while not (os.path.exists(out) and os.path.getsize(out)>1000):
                if time.time()>deadline: print('deadline'); return
                r=subprocess.run(['curl','-s','-m','180','-o',out,'-w','%{http_code}',f'https://web.archive.org/web/{ts}id_/{orig}'],capture_output=True,text=True)
                print(co,ts,r.stdout,flush=True)
                if r.stdout=='200': time.sleep(8); break
                if os.path.exists(out): os.remove(out)
                time.sleep(180)
def parse():
    out=[]
    for co in SETS:
        for f in sorted(glob.glob(f'{W}/jobs_{co}/*.json')):
            ts=os.path.basename(f)[:-5]
            try:
                import gzip; b=open(f,'rb').read(); b=gzip.decompress(b) if b[:2]==b'\x1f\x8b' else b; j=json.loads(b)
            except Exception as e: print('bad',f,e); continue
            for x in j.get('jobs',[]):
                if co=='Lambda':
                    out.append([ts,co,x.get('id'),x.get('title'),x.get('department') or '',x.get('location',''),'; '.join([x.get('location','')]+[s.get('location','') for s in x.get('secondaryLocations',[])]),x.get('publishedAt',''),x.get('publishedAt','')])
                else:
                    out.append([ts,co,x['id'],x['title'],'; '.join(d['name'] for d in x.get('departments',[])),x.get('location',{}).get('name',''),'; '.join(o['name'] for o in x.get('offices',[])),x.get('updated_at',''),x.get('first_published','')])
    with open(f'{D}/jobs_history.csv','w',newline='') as fh:
        w=csv.writer(fh); w.writerow(['snapshot_ts','company','id','title','department','location','offices','updated','first_published']); w.writerows(out)
    print(len(out))
if __name__=='__main__':
    m=sys.argv[1] if len(sys.argv)>1 else 'all'
    if m in('fetch','all'): fetch()
    if m in('parse','all'): parse()
