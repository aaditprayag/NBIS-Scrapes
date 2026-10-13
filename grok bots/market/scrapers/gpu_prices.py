#!/usr/bin/env python3
"""Re-runnable public GPU cloud price scraper (Nebius research).
Usage: python3 gpu_prices.py [--providers nebius,coreweave,...] [--out DIR]
Writes DIR/price_snapshot_<YYYY-MM-DD>.csv with columns:
provider,chip,config,pricing_type,term,quoted_price,quoted_unit,usd_per_gpu_hr,region,notes,source_url,retrieved_date
Only records numbers parsed from source pages/APIs. Failures are logged as rows with pricing_type='ERROR'.
"""
import re, html, json, gzip, csv, sys, argparse, datetime, statistics, urllib.request, urllib.parse
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/128 Safari/537.36'
TODAY=datetime.date.today().isoformat()
COLS=['provider','chip','config','pricing_type','term','quoted_price','quoted_unit','usd_per_gpu_hr','region','notes','source_url','retrieved_date']
CHIPS=['GB300 NVL72','GB200 NVL72','B300','B200','H200','H100','MI355X','MI325X','MI300X']

def get(url, raw=False, timeout=45):
    req=urllib.request.Request(url, headers={'User-Agent':UA,'Accept-Encoding':'identity'})
    b=urllib.request.urlopen(req, timeout=timeout).read()
    if b[:2]==b'\x1f\x8b': b=gzip.decompress(b)
    return b if raw else b.decode('utf-8','ignore')

def text(url):
    s=get(url)
    t=re.sub(r'<script.*?</script>|<style.*?</style>|<svg.*?</svg>','',s,flags=re.S)
    t=re.sub(r'<[^>]+>','|',t); t=html.unescape(t); return re.sub(r'\s*\|[\s|]*','|',t)

def money(s):
    m=re.search(r'\$\s*([0-9][0-9,]*\.?[0-9]*)',s); return float(m.group(1).replace(',','')) if m else None

def row(out,provider,chip,config,pt,term,qp,qu,usd,region,notes,url):
    out.append(dict(provider=provider,chip=chip,config=config,pricing_type=pt,term=term,quoted_price=qp,quoted_unit=qu,
        usd_per_gpu_hr=None if usd is None else round(usd,4),region=region,notes=notes,source_url=url,retrieved_date=TODAY))

def norm_chip(name):
    n=name.upper().replace('HGX ','').replace('NVIDIA ','')
    for c in CHIPS:
        if c.replace(' NVL72','') in n: return c if 'NVL72' in c or c.startswith(('MI','B','H')) else c
    return None

# ---------------- providers ----------------
def nebius(out):
    U='https://nebius.com/prices'; t=text(U)
    sec=t[t.find('NVIDIA GPU Instances'):t.find('Preemptible NVIDIA GPU Instances')]
    hdr='Effective' in sec
    for m in re.finditer(r'NVIDIA (GB300 NVL72|GB200 NVL72|HGX B300|HGX B200|HGX H200|HGX H100)\|[^|]*\|[^|]*\|([^|]*)\|([^|]*)\|',sec):
        chip=m.group(1).replace('HGX ',''); a,b=m.group(2),m.group(3)
        cur=b if hdr else a
        v=money(cur)
        row(out,'Nebius',chip,'NVIDIA '+m.group(1),'on-demand','hourly',v if v else cur.strip(),'$/GPU-hour' if v else 'n/a',v,'Not stated',
            'Current = "Effective Oct 1, 2026" column' if hdr else '',U)
        if hdr and money(a):
            row(out,'Nebius',chip,'NVIDIA '+m.group(1),'on-demand (superseded, prior column)','hourly',money(a),'$/GPU-hour',money(a),'Not stated','Prior column; excluded from stats.',U)
    sp=t[t.find('Preemptible NVIDIA GPU Instances'):t.find('CPU-only instances')]
    for m in re.finditer(r'NVIDIA HGX (B300|B200|H200|H100)\|[^$]*?from \$([0-9.]+)',sp):
        row(out,'Nebius',m.group(1),'NVIDIA HGX '+m.group(1)+' preemptible','spot/preemptible','dynamic','from $'+m.group(2),'$/GPU-hour (floor)',float(m.group(2)),'Not stated','Spot floor; dynamic.',U)

def coreweave(out):
    U='https://www.coreweave.com/pricing'; t=text(U)
    na=t.find('REGION: NORTH AMERICA'); eu=t.find('REGION: EUROPE'); end=t.find('CPU Model',eu)
    for reg,seg in [('North America',t[na:eu]),('Europe',t[eu:end])]:
        for m in re.finditer(r'NVIDIA (GB300 NVL72|GB200 NVL72|HGX B300|HGX B200|HGX H200|HGX H100)\|On-Demand Price:\|((?:\$[0-9.,]+\|)?)/ Hour\|Spot Price:\|((?:\$[0-9.,]+|N/A)\|)?/ Hour',seg):
            chip=m.group(1).replace('HGX ',''); n=4 if 'NVL72' in chip else 8
            od,sp=money(m.group(2) or ''),money(m.group(3) or '')
            row(out,'CoreWeave',chip,f'{m.group(1)} instance ({n} GPUs)','on-demand','hourly',od if od else 'Contact sales',f'$/instance-hour ({n} GPUs)' if od else 'n/a',od/n if od else None,reg,'',U)
            if sp: row(out,'CoreWeave',chip,f'{m.group(1)} instance ({n} GPUs)','spot/preemptible','hourly',sp,f'$/instance-hour ({n} GPUs)',sp/n,reg,'',U)

def coreweave_text_fallback(t): pass

def lambda_(out):
    U='https://lambda.ai/pricing'; t=text(U)
    for m in re.finditer(r'NVIDIA (B200 SXM6|H100 SXM|H100 PCIe)\|[^|]*\|(\d+)\|[^|]*\|[^|]*\|\$([0-9.]+)',t):
        chip='B200' if 'B200' in m.group(1) else ('H100' if 'SXM' in m.group(1) else 'H100 PCIe')
        row(out,'Lambda',chip,f'{m.group(1)} instance ({m.group(2)} vCPU)','on-demand','hourly',float(m.group(3)),'$/GPU-hour',float(m.group(3)),'Not stated','',U)
    for m in re.finditer(r'NVIDIA (HGX B200|H100)\|(2 weeks[^|]*)\|([0-9+]+)\|\$([0-9.]+)',t):
        chip='B200' if 'B200' in m.group(1) else 'H100'
        row(out,'Lambda',chip,f'1-Click Cluster {m.group(3)} GPUs','reserved/committed',m.group(2),float(m.group(4)),'$/GPU-hour',float(m.group(4)),'Not stated','',U)

def crusoe(out):
    U='https://www.crusoe.ai/cloud/pricing'; t=text(U)
    seg=t[t.find('GPU instances pricing'):t.find('CPU instances pricing')]
    for m in re.finditer(r'(NVIDIA|AMD) (GB200|B200|H200|H100|MI355X|MI300X)\|([^|]*)\|(?:(HGX|NVL72|SXM|PCIe)\|)?([^|]*)\|([^|]*)\|',seg):
        chip=m.group(2)+(' NVL72' if m.group(2)=='GB200' else ''); od=money(m.group(5))
        if m.group(4)=='PCIe': continue
        row(out,'Crusoe',chip,f'{m.group(2)} {m.group(3)} {m.group(4) or ""}'.strip(),'on-demand','hourly',od if od else m.group(5),'$/GPU-hour' if od else 'n/a',od,'Not stated',f'Spot: {m.group(6)}',U)
    i=t.find('Self-Serve Deployments pricing'); seg=t[i:t.find('Tailored Deployments',i)]
    for m in re.finditer(r'(NVIDIA|AMD) (H100|H200|B200|MI355x)\|[^|]*\|(?:HGX\|)?\$([0-9.]+)',seg):
        row(out,'Crusoe',m.group(2).upper(),'Managed Inference self-serve deployment','inference/dedicated endpoint','hourly',float(m.group(3)),'$/hour',float(m.group(3)),'Not stated','',U)

def together(out):
    U='https://www.together.ai/pricing'; t=text(U)
    i=t.find('Preemptible Compute|ON-Demand|Reserved'); seg=t[i:i+1500]
    for m in re.finditer(r'NVIDIA\|(HGX H100|HGX H200|HGX B200|HGX B300|GB200 NVL72|GB300 NVL72)\|((?:[^|]*\|){2,6}?)(?=NVIDIA|Sandbox)',seg):
        chip=m.group(1).replace('HGX ',''); cells=[c for c in m.group(2).split('|') if c]
        labels=[('spot/preemptible','pay as you go'),('on-demand','hourly'),('reserved/committed','7-30 days'),('reserved/committed','31-90 days'),('reserved/committed','91-180 days'),('reserved/committed','181+ days')]
        for (pt,term),c in zip(labels,cells):
            v=money(c)
            row(out,'Together AI',chip,'GPU Clusters '+m.group(1),pt,term,v if v else c,'$/GPU-hour' if v else 'n/a',v,'Not stated','',U)

def runpod(out):
    U='https://www.runpod.io/pricing'; t=text(U)
    seg=t[t.find('Pods|'):t.find('Serverless|Cost effective')]
    for m in re.finditer(r'\|(B300|H200|B200|H100 SXM|H100 PCIe|H100 NVL)\|(\d+ GB[^$]*?)\$\|([0-9.]+)\|/hr',seg):
        chip={'H100 SXM':'H100'}.get(m.group(1),m.group(1))
        row(out,'RunPod',chip,'Pod 1x '+m.group(1),'on-demand','per-second',float(m.group(3)),'$/hr (1 GPU)',float(m.group(3)),'31 regions','Community/Secure tab not disambiguated in static HTML.',U)

def aws_capacity_blocks(out):
    U='https://aws.amazon.com/ec2/capacityblocks/pricing/'; t=text(U)
    seen=set()
    for m in re.finditer(r'\|((?:p5|p5e|p5en|p6-b200|p6-b300)\.48xlarge|u-p6e-gb200x\d+)\|([^|]+)\|\$([0-9,.]+)(?: USD)? \(\$([0-9.]+) USD\)\|(\d+) x (\w+)',t):
        key=(m.group(1),m.group(3))
        if key in seen: continue
        seen.add(key); n=int(m.group(5)); chip={'u-p6e':'GB200 NVL72'}.get(m.group(1)[:5], m.group(6))
        row(out,'AWS',chip,f'{m.group(1)} ({n}x {m.group(6)})','reserved/committed','Capacity Block',float(m.group(3).replace(',','')),f'$/instance-hour ({n} GPUs)',float(m.group(4)),m.group(2),'first region listed at this rate',U)

def aws_ondemand(out, region='US East (N. Virginia)'):
    U='https://b0.p.awsstatic.com/pricing/2.0/meteredUnitMaps/ec2/USD/current/ec2-ondemand-without-sec-sel/'+urllib.parse.quote(region)+'/Linux/index.json'
    d=json.loads(get(U,raw=True)); reg=list(d['regions'].values())[0]
    gp={'p5.48xlarge':('H100',8),'p5.4xlarge':('H100',1),'p5e.48xlarge':('H200',8),'p5en.48xlarge':('H200',8),'p6-b200.48xlarge':('B200',8),'p6-b300.48xlarge':('B300',8)}
    for v in reg.values():
        it=v.get('Instance Type')
        if it in gp:
            chip,n=gp[it]; p=float(v['price'])
            row(out,'AWS',chip,it,'on-demand','pay as you go',p,f'$/instance-hour ({n} GPUs)',p/n,region,'Linux',U)

def azure(out, region_by_sku=None):
    base="https://prices.azure.com/api/retail/prices?$filter="
    skus=region_by_sku or {'Standard_ND96isr_H100_v5':('H100',8,'eastus'),'Standard_ND96isr_H200_v5':('H200',8,'eastus2'),'Standard_ND128isr_NDR_GB200_v6':('GB200 NVL72',4,'eastus'),'Standard_ND96isr_MI300X_v5':('MI300X',8,'eastus2')}
    for sku,(chip,n,reg) in skus.items():
        f=f"serviceName eq 'Virtual Machines' and armRegionName eq '{reg}' and armSkuName eq '{sku}'"
        U=base+urllib.parse.quote(f); d=json.loads(get(U))
        od=None
        items=[i for i in d['Items'] if 'Windows' not in i['productName']]
        for i in items:
            if i['type']=='Consumption' and not any(x in i['skuName'] for x in ('Spot','Low Priority')): od=i['retailPrice']
        for i in items:
            p=i['retailPrice']
            if i['type']=='Consumption':
                if 'Low Priority' in i['skuName']: continue
                pt='spot/preemptible' if 'Spot' in i['skuName'] else 'on-demand'
                if pt.startswith('spot') and p==od: pt='spot (anomalous meter = on-demand; excluded)'
                row(out,'Azure',chip,sku,pt,'hourly',p,f'$/VM-hour ({n} GPUs)',p/n,reg,'',U)
            elif i['type']=='Reservation':
                yrs=int(i['reservationTerm'].split()[0])
                row(out,'Azure',chip,sku,'reserved/committed',i['reservationTerm'],p,'$ total per VM for term',p/(8760*yrs)/n,reg,'total/(8760*yrs)/GPUs',U)

def oracle(out):
    U='https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/?currencyCode=USD'; d=json.loads(get(U))
    pat={'GPU - H100':'H100','GPU - H200':'H200','GPU - B200':'B200','GPU - B300':'B300','GPU - GB200':'GB200 NVL72','GPU - GB300':'GB300 NVL72','GPU - MI300X':'MI300X','GPU - MI355X':'MI355X'}
    for i in d['items']:
        n=i['displayName']
        if 'AI Enterprise' in n or 'Cloud@Customer' in n: continue
        for k,chip in pat.items():
            if n.rstrip().endswith(k) and i.get('metricName')=='GPU Per Hour':
                v=[p['value'] for l in i['currencyCodeLocalizations'] for p in l['prices'] if p['model']=='PAY_AS_YOU_GO'][0]
                row(out,'Oracle OCI',chip,f"{i['partNumber']} {n}",'on-demand','pay as you go',v,'$/GPU-hour',v,'global list','',U)

def gcp(out):
    U='https://cloud.google.com/products/compute/pricing/accelerator-optimized'; t=text(U)
    labels=[('on-demand','pay as you go'),('DWS flex-start','Flex-start'),('reserved/committed','DWS Calendar mode'),('spot/preemptible','spot'),('reserved/committed','1-yr CUD'),('reserved/committed','3-yr CUD')]
    for m in re.finditer(r'\|(a4-highgpu-8g|a3-ultragpu-8g|a3-megagpu-8g|a3-highgpu-8g)\|Nvidia (\w+)\|',t):
        toks=t[m.end():m.end()+800].split('|')
        cells=[c.strip() for c in toks if c.strip().startswith(('$','N/A'))][:6]
        for (pt,term),c in zip(labels,cells):
            v=money(c.replace(' ',''))
            row(out,'Google Cloud',m.group(2),m.group(1)+' (8 GPUs)',pt,term,v if v else c,'$/VM-hour (8 GPUs)' if v else 'n/a',v/8 if v else None,'Not stated','',U)

def vast(out):
    for g,chip in [('H100 SXM','H100'),('H200','H200'),('B200','B200'),('B300','B300')]:
        q={"gpu_name":{"eq":g},"rentable":{"eq":True},"rented":{"eq":False},"verified":{"eq":True},"type":"on-demand","limit":500}
        U='https://console.vast.ai/api/v0/bundles/?q='+urllib.parse.quote(json.dumps(q))
        o=json.loads(get(U))['offers']; p=sorted(x['dph_total']/x['num_gpus'] for x in o if x.get('num_gpus'))
        if p:
            med=statistics.median(p)
            row(out,'Vast.ai',chip,g+' verified on-demand','on-demand','per-second',round(med,3),'$/GPU-hour (median of live offers)',med,'global',f'n={len(p)} min {p[0]:.3f} max {p[-1]:.3f}','https://console.vast.ai/api/v0/bundles/')

PROVIDERS={'nebius':nebius,'coreweave':coreweave,'lambda':lambda_,'crusoe':crusoe,'together':together,'runpod':runpod,
 'aws_cb':aws_capacity_blocks,'aws_od':aws_ondemand,'azure':azure,'oracle':oracle,'gcp':gcp,'vast':vast}

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--providers',default=','.join(PROVIDERS)); ap.add_argument('--out',default='/workspace/nebius/market/data')
    a=ap.parse_args(); out=[]
    for p in a.providers.split(','):
        n0=len(out)
        try: PROVIDERS[p](out)
        except Exception as e: row(out,p,'','','ERROR','',str(e)[:200],'',None,'','scrape failed','',)
        print(f'{p:10s} {len(out)-n0:3d} rows', file=sys.stderr)
    import os; os.makedirs(a.out,exist_ok=True)
    fn=f'{a.out}/price_snapshot_{TODAY}.csv'
    with open(fn,'w',newline='') as f:
        w=csv.DictWriter(f,COLS); w.writeheader(); w.writerows(out)
    print(fn)
if __name__=='__main__': main()
