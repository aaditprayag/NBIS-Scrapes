#!/usr/bin/env python3
"""Nebius AI Cloud docs changelog cadence (docs.nebius.com/changelog.md, weekly <Update> blocks).
Counts entries per week by type (Added/Updated/Fixed/Removed) and rolls up by quarter of week-end date.
Also flags self-serve-relevant entries (spot/preemptible/on-demand/pricing/billing/Token Factory/serverless).
Outputs docs_changelog_<date>.csv (weekly) and docs_changelog_selfserve_<date>.csv."""
import re, subprocess, datetime, csv, os, collections
D=os.path.dirname(os.path.abspath(__file__)); today=datetime.date.today().isoformat()
md=subprocess.check_output(['curl','-s','https://docs.nebius.com/changelog.md']).decode('utf-8','ignore')
open(f'{D}/raw/docs_changelog_{today}.md','w').write(md)
M={m:i+1 for i,m in enumerate(['January','February','March','April','May','June','July','August','September','October','November','December'])}
rows=[]; ss=[]
for lab,body in re.findall(r'<Update label="([^"]+)">(.*?)</Update>',md,flags=re.S):
    ms=re.findall("|".join(M),lab); ys=re.findall(r"20\d\d",lab)
    if not ys: continue
    y=int(ys[-1]); d=int(re.findall(r"(\d{1,2})\D*20\d\d",lab)[-1])
    end=datetime.date(y,M[ms[-1]],d)
    items=re.findall(r'^\s*\*\s+(.*)$',body,flags=re.M)
    c=collections.Counter((re.match(r'\*(\w+):\*',i) or [None,'Other'])[1] for i in items)
    rows.append([lab,str(end),f'{end.year}Q{(end.month-1)//3+1}',len(items),c.get('Added',0),c.get('Updated',0),c.get('Fixed',0),c.get('Removed',0)])
    for i in items:
        if re.search(r'spot|preemptible|on-demand|pricing|billing|Token Factory|serverless|self-service|free tier|credits|capacity block|reservation',i,re.I):
            ss.append([str(end),re.sub(r'\s+',' ',i)[:400]])
with open(f'{D}/docs_changelog_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['week','week_end','quarter','entries','added','updated','fixed','removed']); w.writerows(rows)
with open(f'{D}/docs_changelog_selfserve_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['week_end','entry']); w.writerows(ss)
q=collections.defaultdict(lambda:[0,0,0])
for r in rows: q[r[2]][0]+=1; q[r[2]][1]+=r[3]; q[r[2]][2]+=r[4]
for k in sorted(q): print(k,'weeks',q[k][0],'entries',q[k][1],'added',q[k][2])
print(len(ss),'self-serve-relevant entries'); [print(x) for x in ss[:40]]
