#!/usr/bin/env python3
"""Public GitHub org stats via the search API (unauthenticated; 10 req/min).
Writes raw/gh_<org>_<date>.json and github_<date>.csv with repo, stars, forks, created, pushed, archived."""
import json, subprocess, datetime, csv, os, sys, time
D=os.path.dirname(os.path.abspath(__file__)); today=datetime.date.today().isoformat()
ORGS=sys.argv[1:] or ['nebius','coreweave','lambdal']
rows=[]
for org in ORGS:
    items=[]
    for page in (1,2,3):
        j=json.loads(subprocess.check_output(['curl','-s','-H','Accept: application/vnd.github+json',
            f'https://api.github.com/search/repositories?q=org:{org}&sort=stars&order=desc&per_page=100&page={page}']))
        if 'items' not in j: print(org,'ERR',j.get('message')); break
        items+=j['items']
        if len(items)>=j['total_count']: break
        time.sleep(7)
    json.dump(items,open(f'{D}/raw/gh_{org}_{today}.json','w'))
    for x in items: rows.append([org,x['name'],x['stargazers_count'],x['forks_count'],x['created_at'][:10],x['pushed_at'][:10],x['archived'],x['fork']])
    print(org,len(items),sum(x['stargazers_count'] for x in items)); time.sleep(7)
with open(f'{D}/github_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['org','repo','stars','forks','created','pushed','archived','is_fork']); w.writerows(rows)
