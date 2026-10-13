#!/usr/bin/env python3
"""Fetch every nebius.com/customer-stories/* page listed in the server sitemap; extract title, description,
published date (JSON-LD / meta if present), and sentences mentioning GPUs, $ or clusters. Output customer_stories_<date>.csv."""
import re, subprocess, datetime, csv, os, json, html, time
D=os.path.dirname(os.path.abspath(__file__)); R=f'{D}/raw/stories'; os.makedirs(R,exist_ok=True)
UA="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"
today=datetime.date.today().isoformat()
sm=subprocess.check_output(['curl','-s','-L','-A',UA,'https://nebius.com/server-sitemap.xml']).decode()
urls=sorted(set(u for u in re.findall(r'<loc>(.*?)</loc>',sm) if '/customer-stories/' in u))
rows=[]
for u in urls:
    slug=u.rstrip('/').split('/')[-1]; f=f'{R}/{slug}.html'
    if not os.path.exists(f): subprocess.run(['curl','-s','-L','-A',UA,u,'-o',f]); time.sleep(0.5)
    s=open(f,errors='ignore').read()
    title=html.unescape((re.search(r'<title>(.*?)</title>',s,re.S) or [None,''])[1]).strip()
    desc=html.unescape((re.search(r'<meta name="description" content="(.*?)"',s) or [None,''])[1])
    dates=re.findall(r'"(?:datePublished|publishedAt|createdAt|publishDate|date)":"(20\d\d-\d\d-\d\d)',s.replace('\\"','"'))
    t=re.sub(r'<script.*?</script>|<style.*?</style>','',s,flags=re.S); t=html.unescape(re.sub(r'<[^>]+>',' ',t)); t=re.sub(r'\s+',' ',t)
    sents=[x.strip() for x in re.split(r'(?<=[.!?])\s',t) if re.search(r'\b(GPU|GPUs|H100|H200|B200|GB200|GB300|Blackwell|\$\d|million|cluster|nodes)\b',x)]
    rows.append([slug,u,title,desc,min(dates) if dates else '',' | '.join(dict.fromkeys(sents))[:1500]])
with open(f'{D}/customer_stories_{today}.csv','w',newline='') as fh:
    w=csv.writer(fh); w.writerow(['slug','url','title','description','date_in_page','gpu_or_spend_mentions']); w.writerows(rows)
print(len(rows))
