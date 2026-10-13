#!/usr/bin/env python3
"""Snapshot public ATS job boards (Nebius & CoreWeave Greenhouse, Lambda Ashby).
Saves raw JSON to raw/ and a normalized CSV jobs_<date>.csv: company,id,title,department,location,country,updated."""
import json, subprocess, datetime, csv, os
D=os.path.dirname(os.path.abspath(__file__)); R=os.path.join(D,'raw'); os.makedirs(R,exist_ok=True)
UA="Mozilla/5.0 (research; Aadit Prayag NYU Stern)"
today=datetime.date.today().isoformat()
SRC={'Nebius':'https://boards-api.greenhouse.io/v1/boards/nebius/jobs?content=true',
     'CoreWeave':'https://boards-api.greenhouse.io/v1/boards/coreweave/jobs?content=true',
     'Lambda':'https://api.ashbyhq.com/posting-api/job-board/lambda?includeCompensation=false'}
def get(u): return json.loads(subprocess.check_output(['curl','-s','-A',UA,u]))
rows=[]
for co,u in SRC.items():
    j=get(u); json.dump(j,open(f'{R}/{co.lower()}_jobs_{today}.json','w'))
    if 'jobs' in j and co!='Lambda':
        for x in j['jobs']:
            rows.append([co,x['id'],x['title'],'; '.join(d['name'] for d in x.get('departments',[])),
                         x.get('location',{}).get('name',''),'; '.join(o['name'] for o in x.get('offices',[])),x.get('updated_at',''),x.get('first_published','')])
    else:
        for x in j['jobs']:
            rows.append([co,x['id'],x['title'],x.get('department','') or '', x.get('location',''),
                         '; '.join([x.get('location','')]+[s.get('location','') for s in x.get('secondaryLocations',[])]),x.get('publishedAt',''),x.get('publishedAt','')])
with open(f'{D}/jobs_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['company','id','title','department','location','offices','updated','first_published']); w.writerows(rows)
for co in SRC: print(co, sum(r[0]==co for r in rows))
