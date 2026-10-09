#!/usr/bin/env python3
"""Tranco rank (30-day Dowdall-combined, PLD-level) at each quarter-end, Q1'24..latest, plus current daily ranks.
Downloads each daily top-1M list (streamed, not stored) and looks up target domains. Output tranco_<date>.csv."""
import json, subprocess, datetime, csv, os, time
D=os.path.dirname(os.path.abspath(__file__)); today=datetime.date.today()
DOM=['nebius.com','nebius.ai','coreweave.com','lambda.ai','lambdalabs.com','crusoe.ai','together.ai','nscale.com']
dates=[]
for y in (2024,2025,2026):
    for m,dd in ((3,31),(6,30),(9,30),(12,31)):
        d=datetime.date(y,m,dd)
        if d<today: dates.append(d)
dates.append(today-datetime.timedelta(days=1))
rows=[]
for d in dates:
    meta=json.loads(subprocess.check_output(['curl','-s',f'https://tranco-list.eu/api/lists/date/{d}']))
    if not meta.get('available'): print(d,'n/a'); continue
    txt=subprocess.run(['curl','-s','-L',meta['download']],capture_output=True,text=True).stdout
    rank={}
    for line in txt.splitlines():
        r,_,dom=line.partition(',')
        if dom.strip() in DOM: rank[dom.strip()]=int(r)
    for dom in DOM: rows.append([str(d),meta['list_id'],dom,rank.get(dom,'>1,000,000')])
    print(d,meta['list_id'],{k:rank.get(k) for k in DOM}); time.sleep(1.5)
with open(f'{D}/tranco_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['date','tranco_list_id','domain','rank']); w.writerows(rows)
