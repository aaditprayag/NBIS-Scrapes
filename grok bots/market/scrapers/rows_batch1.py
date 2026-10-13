D='2026-10-04'
NEB='https://nebius.com/prices'; CW='https://www.coreweave.com/pricing'; LAM='https://lambda.ai/pricing'
CRU='https://www.crusoe.ai/cloud/pricing'; TOG='https://www.together.ai/pricing'; RP='https://www.runpod.io/pricing'
AWSCB='https://aws.amazon.com/ec2/capacityblocks/pricing/'; VP='https://www.voltagepark.com/pricing'; IR='https://iren.com/ai-cloud'
R=[]
def r(p,chip,cfg,pt,term,qp,qu,usd,reg,notes,url): R.append(dict(provider=p,chip=chip,config=cfg,pricing_type=pt,term=term,quoted_price=qp,quoted_unit=qu,usd_per_gpu_hr=usd,region=reg,notes=notes,source_url=url,retrieved_date=D))
# Nebius
nb_note='Page shows two columns: "On-demand, GPU-hour" and "GPU-hour (Effective October 1, 2026)". As of 2026-10-04 the Oct-1 column is treated as the CURRENT on-demand rate.'
for chip,cfg,old,new in [('B300','NVIDIA HGX B300 (24 vCPU/346GB RAM per GPU)',7.85,9.50),('B200','NVIDIA HGX B200 (20 vCPU/224GB per GPU)',7.15,8.50),('H200','NVIDIA HGX H200 (16 vCPU/200GB per GPU)',4.50,5.40),('H100','NVIDIA HGX H100 (16 vCPU/200GB per GPU)',3.85,4.50)]:
    r('Nebius',chip,cfg,'on-demand','hourly, no commitment',new,'$/GPU-hour',new,'Not stated on page',nb_note,NEB)
    r('Nebius',chip,cfg,'on-demand (superseded, pre-Oct-1-2026 column)','hourly',old,'$/GPU-hour',old,'Not stated on page','Prior on-demand column still displayed on page; excluded from summary stats.',NEB)
for chip,cfg in [('GB300 NVL72','NVIDIA GB300 NVL72 (112 vCPU/800GB)'),('GB200 NVL72','NVIDIA GB200 NVL72 (112 vCPU/800GB)')]:
    r('Nebius',chip,cfg,'on-demand','hourly','Contact us','n/a',None,'Not stated on page','Both price columns read "Contact us".',NEB)
for chip,v in [('B300',0.99),('B200',0.99),('H200',0.79),('H100',0.79)]:
    r('Nebius',chip,'NVIDIA HGX '+chip+' preemptible','spot/preemptible','dynamic (can change every 15 min)','from $%.2f'%v,'$/GPU-hour (floor)',v,'Not stated on page','"from" floor price; spot is dynamic, max = $0.01 below on-demand. Floor, not a typical clearing price.',NEB)
r('Nebius','All (H100-GB300)','Commitment','reserved/committed','multi-month','up to 35% less than on-demand','% discount',None,'n/a','No $ rate published; "contact the sales team".',NEB)
# CoreWeave
for reg in ['North America','Europe']:
    for chip,cfg,od,sp,inst in [('GB200 NVL72','4 GPU instance (2x GB200 superchips), 186GB VRAM/GPU',42.00,None,4),('B200','HGX B200 8-GPU node',68.80,{'North America':34.11,'Europe':34.87}[reg],8),('H100','HGX H100 8-GPU node',49.24,{'North America':19.71,'Europe':19.51}[reg],8),('H200','HGX H200 8-GPU node',50.44,{'North America':20.93,'Europe':20.64}[reg],8)]:
        r('CoreWeave',chip,cfg,'on-demand','hourly',od,'$/instance-hour (%d GPUs)'%inst,round(od/inst,4),reg,'Kubernetes-native; InfiniBand not stated on pricing page.',CW)
        if sp: r('CoreWeave',chip,cfg,'spot/preemptible','hourly',sp,'$/instance-hour (%d GPUs)'%inst,round(sp/inst,4),reg,'',CW)
    r('CoreWeave','B300','HGX B300 8-GPU node','on-demand','hourly','Contact sales','n/a',None,reg,'',CW)
    r('CoreWeave','B300','HGX B300 8-GPU node','spot/preemptible','hourly',{'North America':35.84,'Europe':36.70}[reg],'$/instance-hour (8 GPUs)',round({'North America':35.84,'Europe':36.70}[reg]/8,4),reg,'',CW)
    r('CoreWeave','GB300 NVL72','4 GPU instance (2x GB300 superchips)','on-demand','hourly','Contact sales','n/a',None,reg,'',CW)
for chip,v in [('GB200 NVL72',10.50),('B200',8.60),('H100',6.16),('H200',6.31)]:
    r('CoreWeave',chip,'Inference single-GPU price','inference/dedicated endpoint','hourly',v,'$/GPU-hour',v,'North America','"Exclusively applicable to CoreWeave inference platform customers."',CW)
r('CoreWeave','All','Reserved Compute Capacity','reserved/committed','committed','up to 60% discount vs on-demand','% discount',None,'n/a','No $ rate; contact sales.',CW)
# Lambda
for chip,cfg,v in [('B200','8x B200 SXM6 instance',6.69),('B200','4x B200 SXM6',6.79),('B200','2x B200 SXM6',6.89),('B200','1x B200 SXM6',6.99),('H100','8x H100 SXM instance',3.99),('H100','4x H100 SXM',4.09),('H100','2x H100 SXM',4.19),('H100','1x H100 SXM',4.29),('H100 PCIe','1x H100 PCIe',3.29)]:
    r('Lambda',chip,cfg,'on-demand','hourly, self-serve',v,'$/GPU-hour (plus tax)',v,'Not stated on page','Instances pricing.',LAM)
for chip,n,v in [('B200',16,9.86),('B200',64,9.36),('B200','256+',8.87),('H100',16,6.16),('H100',64,5.85),('H100',256,5.54)]:
    r('Lambda',chip,'1-Click Cluster, %s GPUs'%n,'reserved/committed','2 weeks - 1 year',v,'$/GPU-hour (plus tax)',v,'Not stated on page','1-Click Clusters (16-2,000+ GPUs); "Talk to our team". Note: priced ABOVE Lambda self-serve instances.',LAM)
for chip in ['B200','H100']:
    r('Lambda',chip,'1-Click Cluster 16+ GPUs','reserved/committed','1 year+','—','n/a',None,'Not stated on page','Contact sales ("Talk to our team"); lowest prices via reserved.',LAM)
# Crusoe
for chip,cfg,v in [('H200','141GB HGX',4.29),('H100','80GB HGX',3.90),('MI300X','AMD MI300X 192GB',3.45)]:
    r('Crusoe',chip,cfg,'on-demand','hourly',v,'$/GPU-hour',v,'Not stated on page','Spot: "Contact sales".',CRU)
for chip,cfg in [('GB200 NVL72','186GB NVL72'),('B200','180GB HGX'),('MI355X','AMD MI355X 288GB')]:
    r('Crusoe',chip,cfg,'on-demand','hourly','Contact sales','n/a',None,'Not stated on page','On-demand and spot both "Contact sales".',CRU)
for chip,v in [('H100',5.50),('H200',6.00),('B200',9.65),('MI355X',7.40)]:
    r('Crusoe',chip,'Managed Inference self-serve dedicated deployment','inference/dedicated endpoint','hourly',v,'$/hour per GPU (as listed by GPU type)',v,'Not stated on page','Self-Serve Deployments; contact sales for monthly/volume.',CRU)
# Together
for chip,pre,od,r1,r2,r3 in [('H100',1.99,3.99,3.69,3.45,3.19),('H200',2.99,5.99,4.99,4.15,3.99),('B200',4.09,8.19,7.99,7.79,6.79)]:
    r('Together AI',chip,'GPU Clusters HGX '+chip,'on-demand','hourly',od,'$/GPU-hour',od,'Not stated on page','GPU Clusters.',TOG)
    r('Together AI',chip,'GPU Clusters HGX '+chip,'spot/preemptible','pay as you go',pre,'$/GPU-hour',pre,'Not stated on page','"Preemptible Compute".',TOG)
    for t,v in [('7-30 days',r1),('31-90 days',r2),('91-180 days',r3)]:
        r('Together AI',chip,'GPU Clusters HGX '+chip,'reserved/committed',t,v,'$/GPU-hour',v,'Not stated on page','',TOG)
    r('Together AI',chip,'GPU Clusters HGX '+chip,'reserved/committed','181+ days','Contact us','n/a',None,'Not stated on page','',TOG)
r('Together AI','B300','GPU Clusters HGX B300','on-demand','hourly',9.99,'$/GPU-hour',9.99,'Not stated on page','',TOG)
r('Together AI','B300','GPU Clusters HGX B300','spot/preemptible','pay as you go',4.99,'$/GPU-hour',4.99,'Not stated on page','',TOG)
r('Together AI','B300','GPU Clusters HGX B300','reserved/committed','all terms','Contact us','n/a',None,'Not stated on page','',TOG)
for chip in ['GB200 NVL72','GB300 NVL72']:
    r('Together AI',chip,'GPU Clusters','on-demand','hourly','—','n/a',None,'Not stated on page','Shown as "—" (not offered on-demand); reserved = Contact us.',TOG)
for chip,v in [('H100',5.49),('B200',8.99)]:
    r('Together AI',chip,'Dedicated Inference HGX '+chip,'inference/dedicated endpoint','hourly',v,'$/GPU-hour',v,'Not stated on page','Dedicated Inference; H200/B300/GB200/GB300 "Contact us".',TOG)
# RunPod
rpn='Pods pricing as rendered in static HTML (page has Community Cloud / Secure Cloud toggle; static HTML did not disambiguate which tier). Page "Updated September 27, 2026". Marketplace-style reference.'
for chip,cfg,v in [('B300','Pod 1x B300 288GB',7.89),('H200','Pod 1x H200 141GB',4.59),('B200','Pod 1x B200 180GB',6.79),('H100','Pod 1x H100 SXM 80GB',3.49),('H100 PCIe','Pod 1x H100 PCIe',2.89),('H100 NVL','Pod 1x H100 NVL 94GB',3.19)]:
    r('RunPod',chip,cfg,'on-demand','per-second billing',v,'$/hr (1 GPU)',v,'31 global regions (not per-SKU)',rpn,RP)
r('RunPod','H200','Instant Cluster H200 SXM','on-demand','hourly, no commitment',4.31,'$/hr per GPU',4.31,'Not stated','Clusters up to 64 GPUs; H100 SXM & B200 clusters "Contact sales".',RP)
for chip,v in [('B300',9.98),('B200',8.64),('H200',5.93),('H100',4.79)]:
    r('RunPod',chip,'Serverless worker','inference/dedicated endpoint','per-second',v,'$/hr',v,'Not stated','Serverless (flex) worker price.',RP)
r('RunPod','H100/H200/B200','Reserved Clusters','reserved/committed','1mo/3mo/6mo/12mo/12mo+','Contact sales','n/a',None,'n/a','',RP)
# AWS capacity blocks
cb='EC2 Capacity Blocks for ML effective hourly rate (short-term reservation; rates are dynamic). Page shows per-instance and per-accelerator.'
for chip,cfg,inst,n,acc,reg in [('H100','p5.48xlarge (8x H100)',41.528,8,5.191,'US East (N. Virginia)/Ohio/Oregon/N. California'),('H100','p5.48xlarge (8x H100)',37.76,8,4.720,'Tokyo/Jakarta/Mumbai/Sydney/London/Stockholm/Sao Paulo'),('H200','p5e.48xlarge (8x H200)',47.76,8,5.97,'US East (Ohio)/US West/EU/APAC'),('H200','p5en.48xlarge (8x H200)',54.920,8,6.865,'US East (Ohio/N. Virginia)/US West'),('H200','p5en.48xlarge (8x H200)',49.928,8,6.241,'London/Spain/Jakarta/Mumbai/Tokyo etc.'),('B200','p6-b200.48xlarge (8x B200)',98.84,8,12.355,'US East/US West (Oregon)/Mumbai/Hyderabad'),('B300','p6-b300.48xlarge (8x B300)',112.32,8,14.04,'US West (Oregon)/US East (N. Virginia)/Hyderabad/Jakarta'),('GB200 NVL72','u-p6e-gb200x72 UltraServer (72x B200)',761.904,72,10.582,'US East (Dallas) Local Zone'),('GB200 NVL72','u-p6e-gb200x36 UltraServer (36x B200)',380.952,36,10.582,'US East (Dallas) Local Zone')]:
    r('AWS',chip,cfg,'reserved/committed','Capacity Block (short-term, fixed dates)',inst,'$/instance-hour (%d GPUs)'%n,acc,reg,cb+' EFA networking.',AWSCB)
# Voltage Park
for cfg,t in [('On-Demand, Ethernet (1-1016 HGX H100)','hourly'),('On-demand, 3200 Gbps InfiniBand (8-1016 HGX H100)','hourly'),('Long-Term Reserve 32-8000+ H100, 3200 Gbps InfiniBand','6+ month contracts')]:
    r('Voltage Park',  'H100',cfg,'on-demand' if cfg.lower().startswith('on') else 'reserved/committed',t,'Contact for pricing','n/a',None,'Not stated','Voltage Park has merged with Lightning AI.',VP)
r('Voltage Park','H100','1x H100, no contract','on-demand (floor quote, FAQ)','hourly','starting at $1.99','$/GPU-hour (floor)',1.99,'Not stated','FAQ text: "rent one H100 GPU for 1 hour starting at $1.99 without a contract." Floor price from FAQ, not the pricing table.',VP)
# IREN
for chip in ['H100','H200','B200','B300','GB300 NVL72']:
    r('IREN',chip,'IREN Cloud '+chip,'on-demand','n/a','Contact sales','n/a',None,'Not stated','All SKUs "Contact sales".',IR)
