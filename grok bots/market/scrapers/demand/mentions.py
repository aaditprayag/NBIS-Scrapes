#!/usr/bin/env python3
"""Quarterly mention counts: Hacker News (Algolia API, stories+comments) and Stack Overflow (StackExchange API, questions).
Output mentions_<date>.csv."""
import json, subprocess, datetime, csv, os, time, urllib.parse
D=os.path.dirname(os.path.abspath(__file__)); today=datetime.date.today().isoformat()
TERMS={'Nebius':'nebius','CoreWeave':'coreweave','Lambda Labs':'"lambda labs"'}
def qs():
    for y in (2024,2025,2026):
        for q in range(4):
            a=datetime.datetime(y,3*q+1,1); b=datetime.datetime(y+(q==3),(3*q+3)%12+1,1)
            if a>datetime.datetime.now(): return
            yield f'{y}Q{q+1}',int(a.timestamp()),int(b.timestamp())
def get(u):
    return json.loads(subprocess.check_output(['curl','-s','--compressed',u]))
rows=[]
for name,t in TERMS.items():
    for lab,a,b in qs():
        for tag in ('story','comment'):
            u=f'https://hn.algolia.com/api/v1/search?query={urllib.parse.quote(t)}&tags={tag}&numericFilters={urllib.parse.quote(f'created_at_i>={a},created_at_i<{b}')}&hitsPerPage=0'
            rows.append(['HN',name,tag,lab,get(u).get('nbHits')])
        u=f'https://api.stackexchange.com/2.3/search/excerpts?site=stackoverflow&q={urllib.parse.quote(t)}&fromdate={a}&todate={b-1}&filter=total'
        j=get(u); rows.append(['StackOverflow',name,'posts (q+a, full text)',lab,j.get('total')]); time.sleep(0.3)
with open(f'{D}/mentions_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['site','entity','type','quarter','count']); w.writerows(rows)
print(len(rows))
