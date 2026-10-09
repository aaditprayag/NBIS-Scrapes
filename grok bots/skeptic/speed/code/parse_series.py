"""Parse dated on-demand $/GPU-hr from Wayback snapshots already downloaded by the GPU Market teammate
(/workspace/nebius/wayback_raw/<dir>/<timestamp>.html, converted to text by wb2txt.py).
Output: data/price_series.csv  (one row per provider x chip x snapshot). Every row carries the Wayback URL."""
import re,glob,os,csv
TXT='/workspace/nebius/skeptic/speed/txt'
ORIG={'https_nebius_ai_prices':'https://nebius.ai/prices','https_nebius_com_prices':'https://nebius.com/prices',
 'https_coreweave_com_gpu_cloud_pricing':'https://coreweave.com/gpu-cloud-pricing','https_www_coreweave_com_gpu_cloud_pricing':'https://www.coreweave.com/gpu-cloud-pricing',
 'http_www_coreweave_com_gpu_cloud_pricing':'http://www.coreweave.com/gpu-cloud-pricing','https_coreweave_com_pricing':'https://coreweave.com/pricing',
 'https_www_coreweave_com_pricing':'https://www.coreweave.com/pricing','https_lambdalabs_com_service_gpu_cloud':'https://lambdalabs.com/service/gpu-cloud',
 'https_lambda_ai_service_gpu_cloud':'https://lambda.ai/service/gpu-cloud','https_lambda_ai_pricing':'https://lambda.ai/pricing',
 'https_aws_amazon_com_ec2_capacityblocks_pricing':'https://aws.amazon.com/ec2/capacityblocks/pricing/'}
rows=[]
def add(prov,chip,ts,price,basis,snip,d):
    rows.append(dict(provider=prov,chip=chip,date=f'{ts[:4]}-{ts[4:6]}-{ts[6:8]}',usd_per_gpu_hr=round(price,4),basis=basis,
        snippet=snip[:160].replace('\n',' '),source_url=f'https://web.archive.org/web/{ts}/{ORIG[d]}',tag='Reported' if 'derived' not in basis else 'Derived'))
for f in sorted(glob.glob(TXT+'/*.txt')):
    d,ts=os.path.basename(f)[:-4].split('__'); t=open(f,errors='ignore').read()
    if d not in ORIG: continue
    if d=='https_nebius_ai_prices':
        for m in re.finditer(r'GPU type: (H100 SXM5|A100 SXM4|L40S)\. PAYG price – \$([\d.]+)/h',t):
            add('Nebius',m.group(1).split()[0],ts,float(m.group(2)),'PAYG on-demand per GPU-hr',m.group(0),d)
    elif d=='https_nebius_com_prices':
        two = 'Preemptible, GPU-hour|On-demand' in t; eff='Effective October 1, 2026' in t
        pats=[r'1 × (H100|H200) SXM\|\d+\|\d+\|\d+\|\$([\d.]+)', r'NVIDIA (H200|H100) GPU\|\d+\|\d+\|\d+\|\$([\d.]+)',
              r'NVIDIA (?:HGX )?(B300|B200|H200|H100)\|\d+\|\d+\|\$([\d.]+)(?:\|\$([\d.]+))?']
        seen=set()
        for p in pats:
            for m in re.finditer(p,t):
                chip=m.group(1)
                if chip in seen: continue
                seen.add(chip)
                if m.lastindex>=3 and m.group(3):
                    if two: add('Nebius',chip,ts,float(m.group(3)),'on-demand per GPU-hr (2nd col; 1st col=preemptible)',m.group(0),d)
                    elif eff:
                        add('Nebius',chip,ts,float(m.group(2)),'on-demand per GPU-hr, column in force through 2026-09-30',m.group(0),d)
                        rows.append(dict(rows[-1],date='2026-10-01',usd_per_gpu_hr=float(m.group(3)),basis='on-demand per GPU-hr, column "Effective October 1, 2026"'))
                    else: add('Nebius',chip,ts,float(m.group(2)),'on-demand per GPU-hr',m.group(0),d)
                else: add('Nebius',chip,ts,float(m.group(2)),'on-demand per GPU-hr',m.group(0),d)
    elif 'coreweave' in d:
        seen=set()
        for m in re.finditer(r'NVIDIA HGX (H100)\|80\|48\|256\|\$([\d.]+)|(A100 80GB NVLINK)\|80\|48\|256\|\$([\d.]+)',t):
            chip='H100' if m.group(1) else 'A100 80GB'
            if chip in seen: continue
            seen.add(chip); add('CoreWeave',chip,ts,float(m.group(2) or m.group(4)),'GPU component price per GPU-hr (excl. CPU/RAM, legacy rate card)',m.group(0),d)
        for m in re.finditer(r'NVIDIA (?:HGX )?(GB200 NVL72|B200|H200|H100|A100)\|(\d+)(?:\^1)?\|\d+\|\d+\|[\d,]+\|[\d.]+\|\$([\d,.]+)',t):
            chip=m.group(1); n=int(m.group(2))
            if chip in seen: continue
            seen.add(chip); p=float(m.group(3).replace(',',''))
            add('CoreWeave',chip,ts,p/n,f'instance on-demand ${p}/hr / {n} GPUs (derived)',m.group(0),d)
    elif 'lambda' in d:
        seen=set()
        for m in re.finditer(r'(?<![G\w])(?:8x )?NVIDIA (H100 SXM|A100 SXM|B200 SXM6)\|(?:New(?: Price)?\|)?(80|40|180) GB\|(208|240|124)\|[^$]{0,60}?\$([\d.]+)( / hr| / GPU / hr)?(?:\|\$([\d.]+) / hr)?',t):
            chip={'H100 SXM':'H100','A100 SXM':'A100 80GB' if m.group(2)=='80' else 'A100 40GB','B200 SXM6':'B200'}[m.group(1)]
            if chip in seen: continue
            seen.add(chip)
            if m.group(6):  # Jan-2024 page: '2023 Price' and '2024 Price (effective 1/1/2024)'
                add('Lambda',chip,'20231231',float(m.group(4))/8,'8x instance on-demand, "2023 Price" column /8 (derived)',m.group(0),d)
                add('Lambda',chip,ts,float(m.group(6))/8,'8x instance on-demand, "2024 Price (effective 1/1/2024)" /8 (derived)',m.group(0),d)
                continue
            p=float(m.group(4))
            if m.group(5)==' / hr' or p>15: add('Lambda',chip,ts,p/8,'8x instance on-demand /8 (derived)',m.group(0),d)
            else: add('Lambda',chip,ts,p,'8x instance on-demand per GPU-hr',m.group(0),d)
    elif 'capacityblocks' in d:
        specs=[('H100','p5\\.48xlarge\\|US East \\(N\\. Virginia\\)'),('H200','p5e\\.48xlarge\\|US East \\(Ohio\\)'),('H200 (p5en)','p5en\\.48xlarge\\|US East \\(N\\. Virginia\\)'),
               ('B200','p6-b200\\.48xlarge\\|US East \\(N\\. Virginia\\)'),('B200','p6-b200\\.48xlarge\\|US West \\(Oregon\\)'),('GB200 NVL72','p6e-gb200x72(?: \\(UltraServer\\))?\\|[^$]{0,60}'),('A100 40GB','p4d\\.24xlarge\\|US East \\(Ohio\\)')]
        seen=set()
        for chip,p in specs:
            if chip in seen: continue
            m=re.search(p+r'\|\$([\d.,]+) USD ?\(? ?\$?([\d.]+)? ?(?:USD)?\)?\|\$?([\d.]+)?',t) or re.search(p+r'\|\$([\d.,]+) \(\$([\d.]+) USD\)',t)
            if not m: continue
            per=m.group(2) or m.group(3)
            if per is None: continue
            seen.add(chip); add('AWS Capacity Blocks',chip,ts,float(per),'EC2 Capacity Block effective rate per accelerator-hr (reported)',m.group(0),d)
with open('/workspace/nebius/skeptic/speed/data/price_series.csv','w',newline='') as fo:
    w=csv.DictWriter(fo,list(rows[0])); w.writeheader(); w.writerows(sorted(rows,key=lambda r:(r['provider'],r['chip'],r['date'])))
print(len(rows))
