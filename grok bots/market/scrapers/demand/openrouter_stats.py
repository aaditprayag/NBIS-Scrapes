#!/usr/bin/env python3
"""OpenRouter public data for Nebius Token Factory:
 - provider leaderboard (openrouter.ai/providers RSC payload): modelCount, dailyTokens, monthlyTokens for every provider -> share & rank
 - Nebius provider page chartData: daily tokens by model (last ~90 days)
 - per-model endpoint pricing for Nebius (/api/v1/models/{id}/endpoints) -> implied OpenRouter-routed revenue (Derived)
Outputs openrouter_providers_<date>.csv, openrouter_nebius_daily_<date>.csv, openrouter_nebius_models_<date>.csv"""
import re, json, subprocess, datetime, csv, os, time, collections
D=os.path.dirname(os.path.abspath(__file__)); R=f'{D}/raw'; today=datetime.date.today().isoformat()
UA="Mozilla/5.0 (X11; Linux x86_64) Chrome/124"
def get(u): return subprocess.check_output(['curl','-s','-A',UA,u]).decode('utf-8','ignore')
def rsc(s): return "".join(re.findall(r'self\.__next_f\.push\(\[1,"(.*?)"\]\)',s,flags=re.S)).encode().decode('unicode_escape',errors='ignore')
pp=get('https://openrouter.ai/providers'); open(f'{R}/or_providers_page_{today}.html','w').write(pp); big=rsc(pp)
prov={}
for m in re.finditer(r'\{"name":"([^"]+)","slug":"([^"]+)","displayName":"([^"]*)".*?"modelCount":(\d+).*?"dailyTokens":(\d+|null),"monthlyTokens":(\d+|null)',big):
    prov[m.group(2)]=[m.group(3),int(m.group(4)),int(m.group(5)) if m.group(5)!='null' else 0,int(m.group(6)) if m.group(6)!='null' else 0]
tot=sum(v[3] for v in prov.values()); r=sorted(prov.items(),key=lambda kv:-kv[1][3])
with open(f'{D}/openrouter_providers_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['rank_by_monthly_tokens','slug','display_name','model_count','daily_tokens','monthly_tokens','share_of_listed_monthly_tokens'])
    for i,(k,v) in enumerate(r): w.writerow([i+1,k]+v+[v[3]/tot])
pg=get('https://openrouter.ai/provider/nebius'); open(f'{R}/or_provider_nebius_{today}.html','w').write(pg); big=rsc(pg)
k=big.index('"chartData":')+len('"chartData":'); depth=0
for n,ch in enumerate(big[k:]):
    depth+= (ch=='[')-(ch==']')
    if depth==0: end=k+n+1; break
cd=json.loads(big[k:end])
models=sorted({m for d in cd for m in d['ys'] if m!='Others'})
with open(f'{D}/openrouter_nebius_daily_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['date','total_tokens']+models+['Others'])
    for d in cd: w.writerow([d['x'][:10],sum(d['ys'].values())]+[d['ys'].get(m,0) for m in models]+[d['ys'].get('Others',0)])
# pricing of Nebius endpoints for every model Nebius serves (from provider page) + charted models
served=sorted(set(re.findall(r'"provider_slug":"nebius[^"]*".{0,600}?"model_variant_permaslug":"([^"]+)"',big)) | set(models))
last30=collections.Counter()
for d in cd[-31:-1]:
    for m,v in d['ys'].items(): last30[m]+=v
rows=[]
for m in sorted(set(models)|set(x.split(':')[0] for x in served)):
    base=re.sub(r'-\d\d-\d\d$','',m)  # permaslug date suffix
    j={}
    for cand in (m,base):
        try:
            j=json.loads(get(f'https://openrouter.ai/api/v1/models/{cand}/endpoints')); 
            if 'data' in j: break
        except Exception: pass
        time.sleep(0.3)
    eps=[e for e in j.get('data',{}).get('endpoints',[]) if 'ebius' in e.get('provider_name','')]
    nprov=len(j.get('data',{}).get('endpoints',[]))
    pin=float(eps[0]['pricing']['prompt'])*1e6 if eps else None; pout=float(eps[0]['pricing']['completion'])*1e6 if eps else None
    rows.append([m,len(eps),nprov,pin,pout,last30.get(m,0)])
with open(f'{D}/openrouter_nebius_models_{today}.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['model','nebius_endpoints','total_provider_endpoints','nebius_price_in_$per_M','nebius_price_out_$per_M','nebius_tokens_last30d']); w.writerows(rows)
print('providers',len(prov),'nebius rank',[i+1 for i,(k,v) in enumerate(r) if k=='nebius'], prov.get('nebius'), 'share',prov['nebius'][3]/tot)
for x in rows: print(x)
print('Others last30',last30.get('Others'))
