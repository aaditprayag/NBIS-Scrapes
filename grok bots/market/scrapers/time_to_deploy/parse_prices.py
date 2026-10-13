"""Parse cached Wayback pricing-page text into a long table of on-demand $/GPU-hr per chip per snapshot.
Input : cache_text/<page_dir>/<ts>.txt (built by totext.py from /workspace/nebius/wayback_raw)
Output: out/price_obs.csv  (provider,page_dir,ts,date,chip,price_gpu_hr,basis,raw,archive_url)
Every value is 'Reported' (as printed) or 'Derived' (instance price / GPU count; math in 'basis')."""
import re, os, glob, csv, json
BASE=os.path.dirname(os.path.abspath(__file__)); TXT=f'{BASE}/cache_text'; OUT=f'{BASE}/out'
ORIG={  # page_dir -> original URL (for archive links)
 'https_nebius_ai_prices':'https://nebius.ai/prices','https_nebius_com_prices':'https://nebius.com/prices',
 'https_www_coreweave_com_gpu_cloud_pricing':'https://www.coreweave.com/gpu-cloud-pricing','https_coreweave_com_gpu_cloud_pricing':'https://coreweave.com/gpu-cloud-pricing',
 'http_www_coreweave_com_gpu_cloud_pricing':'http://www.coreweave.com/gpu-cloud-pricing','https_www_coreweave_com_pricing':'https://www.coreweave.com/pricing','https_coreweave_com_pricing':'https://coreweave.com/pricing',
 'https_lambdalabs_com_service_gpu_cloud':'https://lambdalabs.com/service/gpu-cloud','https_lambda_ai_service_gpu_cloud':'https://lambda.ai/service/gpu-cloud','https_lambda_ai_pricing':'https://lambda.ai/pricing',
 'https_crusoe_ai_cloud_pricing':'https://crusoe.ai/cloud/pricing','https_www_crusoe_ai_cloud_pricing':'https://www.crusoe.ai/cloud/pricing',
 'https_www_together_ai_pricing':'https://www.together.ai/pricing','https_aws_amazon_com_ec2_capacityblocks_pricing':'https://aws.amazon.com/ec2/capacityblocks/pricing/'}
# Nebius pricing page also lives under URL variants (nebius.ai ?hss_channel=..., nebius.com ?utm_...); include every cached Nebius
# variant so the old nebius.ai domain (Oct 2024, first H200 listing) and gap-fill captures are not missed (fix 2026-10-04).
import json as _json
for _f in glob.glob('/workspace/nebius/wayback_raw/fetchlog_*.jsonl'):
    for _l in open(_f):
        _d=_json.loads(_l)
        if _d.get('file') and _d.get('orig') and 'nebius' in _d['orig'] and 'ai-studio' not in _d['orig'] and 'ai_studio' not in _d['file']:
            ORIG.setdefault(os.path.basename(os.path.dirname(_d['file'])),_d['orig'])
PROV={'nebius':'Nebius','coreweave':'CoreWeave','lambda':'Lambda','crusoe':'Crusoe','together':'Together AI','aws':'AWS'}
def prov(d):
    for k,v in PROV.items():
        if k in d: return v
def num(s): return float(s.replace(',','').replace('$','').strip())
def chipnorm(c):
    c=c.upper()
    for k in ['GB300','GB200','B300','B200','H200','H100']:
        if k in c: return k+(' NVL72' if k.startswith('GB') else '')
def nebius(t,ts=''):
    r=[]
    for m in re.finditer(r'GPU type: (H100|H200|A100)[^.]*\. PAYG price – \$([\d.]+)/h',t):
        if m.group(1)!='A100': r.append((m.group(1),num(m.group(2)),'Reported: PAYG $/GPU-hr',m.group(0)))
    for m in re.finditer(r'1 × (H200|H100) SXM \d+ \d+ \d+ \$([\d.]+)',t):
        r.append((m.group(1),num(m.group(2)),'Reported: 1-GPU on-demand',m.group(0)))
    for m in re.finditer(r'NVIDIA (H200|H100) GPU \d+ \d+ \d+ \$([\d.]+)',t):
        r.append((m.group(1),num(m.group(2)),'Reported: on-demand $/hr per GPU',m.group(0)))
    two = 'Preemptible, GPU-hour On-demand, GPU-hour' in t
    eff = re.search(r'GPU-hour \(Effective ([A-Z][a-z]+ \d+, \d{4})\)',t)
    for m in re.finditer(r'NVIDIA (?:HGX )?(B300|B200|H200|H100|GB300 NVL72\*?|GB200 NVL72\*?) \d+ \d+ (\$[\d.]+|Contact us|––)(?: (\$[\d.]+|Contact us|––))?',t):
        a,b=m.group(2),m.group(3); chip=chipnorm(m.group(1))
        if two: v=b; basis='Reported: on-demand column (2nd; 1st is preemptible)'
        elif eff and b:
            import datetime
            ed=datetime.datetime.strptime(eff.group(1),'%B %d, %Y').strftime('%Y%m%d')
            if ts[:8] < ed:   # new price announced but not yet effective -> current = first column
                r.append((chip,num(b) if b.startswith('$') else None,'Announced: future on-demand price effective '+eff.group(1),m.group(0)))
                v=a; basis='Reported: current on-demand column (new price announced, effective '+eff.group(1)+')'
            else:
                r.append((chip,num(a) if a.startswith('$') else None,'Reported: on-demand column before '+eff.group(1),m.group(0)))
                v=b; basis='Reported: on-demand column "Effective '+eff.group(1)+'"'
        else: v=a; basis='Reported: on-demand $/GPU-hr'
        r.append((chip,num(v) if v and v.startswith('$') else None,basis+('' if v and v.startswith('$') else ' (no public price: '+str(v)+')'),m.group(0)))
    return r
def coreweave(t):
    r=[]
    for m in re.finditer(r'NVIDIA HGX (H100|H200) 80 48 256 \$([\d.]+)',t):  # old per-GPU component table
        r.append((m.group(1),num(m.group(2)),'Reported: HGX per-GPU hourly (component pricing)',m.group(0)))
    seen=set()
    for m in re.finditer(r'NVIDIA (?:HGX )?(GB300 NVL72|GB200 NVL72|B300|B200|H200|H100) (\d)(?:\^1)? (\d+) (\d+) ([\d,]+) ([\d.]+) (\$[\d,.]+|Contact sales)',t):
        chip=chipnorm(m.group(1)); n=int(m.group(2))
        if chip in seen: continue
        seen.add(chip); p=m.group(7)
        if p.startswith('$'): r.append((chip,round(num(p)/n,4),f'Derived: instance {p}/hr ÷ {n} GPUs',m.group(0)))
        else: r.append((chip,None,'No public price (Contact sales)',m.group(0)))
    return r
def lambda_(t):
    r=[]; seen=set()
    for m in re.finditer(r'(?:8x )?NVIDIA (H100 SXM|B200 SXM6|H200 SXM|B300 SXM6?)(?: New)?(?: Price)? (\d+) GB 208 (?:(?!CONTACT|Reserved)[^$]){0,40}?\$([\d.,]+)( / hr| / GPU / hr)?',t):
        chip=chipnorm(m.group(1))
        if chip in seen: continue
        seen.add(chip); p=num(m.group(3))
        if m.group(4)==' / hr' and p>15: r.append((chip,round(p/8,4),f'Derived: 8x instance ${p}/hr ÷ 8',m.group(0)))
        else: r.append((chip,p,'Reported: 8x on-demand $/GPU-hr',m.group(0)))
    m=re.search(r'1x NVIDIA H100 PCIe(?: New)? 80 GB 26 [^$]{0,40}?\$([\d.]+)',t)
    if m: r.append(('H100 PCIe',num(m.group(1)),'Reported: 1x H100 PCIe on-demand $/GPU-hr',m.group(0)))
    return r
def crusoe(t):
    r=[]; seen=set()
    for m in re.finditer(r'NVIDIA (GB300|GB200|B300|B200|H200|H100) (?:Tensor Core GPU )?\d+GB (?:SXM|HGX|NVL72)(?: On-Demand)? (\$ ?[\d.]+|Contact [Ss]ales)',t):
        chip=chipnorm(m.group(1))
        if chip in seen: continue
        seen.add(chip); p=m.group(2)
        r.append((chip,num(p) if p.startswith('$') else None,'Reported: on-demand $/GPU-hr' if p.startswith('$') else 'No public price (Contact sales)',m.group(0)))
    return r
def together(t):
    r=[]; seen=set()
    i=t.find('Pay as you go'); s=t[i:] if i>=0 else t
    for m in re.finditer(r'NVIDIA (?:HGX )?(GB300 NVL72|GB200 NVL72|B300|B200|H200|H100)(?: SXM)?(?: \d+GB)? (\$[\d.]+|Contact us|Contact sales|—)',s):
        chip=chipnorm(m.group(1))
        if chip in seen: continue
        seen.add(chip); p=m.group(2)
        r.append((chip,num(p) if p.startswith('$') else None,'Reported: first on-demand $/GPU-hr listed' if p.startswith('$') else 'No public price ('+p+')',m.group(0)))
    return r
def aws(t):
    r=[]
    m=re.search(r'Effective ([A-Z][a-z]+ \d+, \d{4}), hourly rates per accelerator[^:]*: (.{0,400})',t)
    if m:
        for k,c in [('P6-B300','B300'),('P6-B200','B200'),('P5en','H200'),('P5e','H200e'),('P5','H100')]:
            mm=re.search(re.escape(k)+r' at \$(\d+\.\d+)',m.group(2))
            if mm and c!='H200e': r.append((c,num(mm.group(1)),'Reported: EC2 Capacity Block $/accelerator-hr, effective '+m.group(1)+' ('+k+')',mm.group(0)))
    return r
P={'Nebius':nebius,'CoreWeave':coreweave,'Lambda':lambda_,'Crusoe':crusoe,'Together AI':together,'AWS':aws}
def main():
    os.makedirs(OUT,exist_ok=True); rows=[]
    for d in sorted(os.listdir(TXT)):
        if d not in ORIG: continue
        pv=prov(d)
        for f in sorted(glob.glob(f'{TXT}/{d}/*.txt')):
            ts=os.path.basename(f)[:-4]; t=open(f).read()
            for chip,price,basis,raw in (P[pv](t,ts) if pv=='Nebius' else P[pv](t)):
                rows.append(dict(provider=pv,page_dir=d,ts=ts,date=f'{ts[:4]}-{ts[4:6]}-{ts[6:8]}',chip=chip,price_gpu_hr=price,basis=basis,raw=raw[:200],
                                 archive_url=f'https://web.archive.org/web/{ts}/{ORIG[d]}'))
    with open(f'{OUT}/price_obs.csv','w',newline='') as fh:
        w=csv.DictWriter(fh,fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
    print(len(rows),'rows')
if __name__=='__main__': main()
