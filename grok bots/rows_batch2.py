D='2026-10-04'
R=[]
def r(p,chip,cfg,pt,term,qp,qu,usd,reg,notes,url): R.append(dict(provider=p,chip=chip,config=cfg,pricing_type=pt,term=term,quoted_price=qp,quoted_unit=qu,usd_per_gpu_hr=None if usd is None else round(usd,4),region=reg,notes=notes,source_url=url,retrieved_date=D))
# Oracle OCI (public price-list API that backs oracle.com/cloud/price-list)
OAPI='https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/?currencyCode=USD (backs https://www.oracle.com/cloud/compute/gpu/pricing/)'
for part,chip,cfg,v,ic in [('B98415','H100','BM.GPU.H100.8 (8x H100 80GB)',10,'8x2x200 Gb/s RDMA'),('B110519','H200','BM.GPU.H200.8 (8x H200)',10,'8x1x400 Gb/s RDMA'),('B110978','B200','BM.GPU.B200.8 (8x B200)',14,'8x400 Gb/s'),('B112237','B300','BM.GPU.B300.8 (8x B300 263GB)',15,'8x800 Gb/s'),('B110979','GB200 NVL72','BM.GPU.GB200.4 (4x B200 189GB, NVL72)',16,'4x400 Gb/s InfiniBand'),('B112140','GB300 NVL72','BM.GPU.GB300.4 (4x B300 279GB, NVL72)',18,'4x800 Gb/s'),('B109485','MI300X','BM.GPU.MI300X.8 (8x MI300X)',6,'8x1x400 Gb/s RDMA'),('B111758','MI355X','BM.GPU.MI355X.8 (8x MI355X 288GB)',8.6,'8x1x400 Gb/s RDMA')]:
    r('Oracle OCI',chip,cfg,'on-demand','pay as you go',v,'$/GPU-hour',v,'Not stated (OCI list price is global)',f'Part {part}; cluster network {ic} per gpu pricing page. NVIDIA AI Enterprise add-on billed separately.',OAPI)
# Azure (Retail Prices API)
AZ='https://prices.azure.com/api/retail/prices (Azure Retail Prices API; Linux, Consumption/Reservation)'
def az(chip,sku,n,reg,od,spot,lowpri,res,ic):
    r('Azure',chip,sku+f' ({n} GPUs)','on-demand','pay as you go',od,f'$/VM-hour ({n} GPUs)',od/n,reg,ic,AZ)
    if spot is not None: r('Azure',chip,sku,'spot/preemptible' if spot!=od else 'spot (anomalous meter = on-demand; excluded)','spot',spot,f'$/VM-hour ({n} GPUs)',spot/n,reg,'Spot meter as listed.'+(' NOTE: spot meter equals on-demand rate in API for this SKU.' if spot==od else ''),AZ)
    for term,tot,yrs in res:
        r('Azure',chip,sku,'reserved/committed',term,tot,f'$ total per VM for {term} reservation',tot/(8760*yrs)/n,reg,f'Reservation listed as total term price; hourly = total/(8760 x {yrs})/{n}.',AZ)
az('H100','Standard_ND96isr_H100_v5',8,'eastus',98.32,18.169536,19.664,[('1 Year',551221.0,1),('3 Years',1134310.0,3),('5 Years',1722566.0,5)],'InfiniBand (ND isr). Low Priority $19.664/VM-hr. Same rates eastus2/westus3; southcentralus $117.968.')
az('H200','Standard_ND96isr_H200_v5',8,'eastus2',84.80,84.80,16.96,[('1 Year',407526.0,1),('3 Years',1109592.0,3)],'InfiniBand. Low Priority $16.96/VM-hr. Reservation meters effective 2026-08-01.')
az('GB200 NVL72','Standard_ND128isr_NDR_GB200_v6',4,'eastus',108.16,108.16,None,[('1 Year',606388.0,1),('3 Years',1250676.0,3)],'4 Blackwell GPUs/VM per Microsoft Learn; 4x400 Gb/s Quantum-2 InfiniBand; NVL72 = 18 VMs.')
az('MI300X','Standard_ND96isr_MI300X_v5',8,'eastus2',48.0,8.8704,9.6,[('1 Year',269107.0,1),('3 Years',553772.0,3)],'Low Priority $9.60/VM-hr.')
r('Azure','B200/GB300','ND B200 / ND GB300 SKUs','on-demand','n/a','Not found','n/a',None,'all','No B200 or GB300 VM meters returned by Retail Prices API on 2026-10-04.',AZ)
# GCP
G='https://cloud.google.com/products/compute/pricing/accelerator-optimized'
gn='Region not stated on page (Google default list region). Components incl. vCPU/RAM/local SSD.'
for chip,cfg,od,flex,cal,spot,c1,c3 in [('B200','a4-highgpu-8g (8x B200)',None,64.44,90.22,39.6336,88.9272,56.7072),('H200','a3-ultragpu-8g (8x H200)',84.806908493,42.40,59.36,50.874449096,58.471933151,37.208420822),('H100','a3-megagpu-8g (8x H100 Mega)',93.400712807,40.32,44.00,55.894488548,64.206558727,40.650103355),('H100','a3-highgpu-8g (8x H100)',88.490000119,38.32,41.60,52.962296548,61.383674231,38.864383195)]:
    if od: r('Google Cloud',chip,cfg,'on-demand','pay as you go',od,'$/VM-hour (8 GPUs)',od/8,'Not stated',gn,G)
    else: r('Google Cloud',chip,cfg,'on-demand','pay as you go','N/A','n/a',None,'Not stated','On-demand listed as N/A; use DWS/CUD/Spot.',G)
    r('Google Cloud',chip,cfg,'DWS flex-start','Flex-start (variable)',flex,'$/VM-hour (8 GPUs)',flex/8,'Not stated','Dynamic Workload Scheduler Flex-start; variable price.',G)
    r('Google Cloud',chip,cfg,'reserved/committed','DWS Calendar mode (short-term reservation)',cal,'$/VM-hour (8 GPUs)',cal/8,'Not stated','DWS Calendar Mode.',G)
    r('Google Cloud',chip,cfg,'spot/preemptible','spot (variable daily)',spot,'$/VM-hour (8 GPUs)',spot/8,'Not stated','Current Spot pricing.',G)
    r('Google Cloud',chip,cfg,'reserved/committed','1-yr CUD',c1,'$/VM-hour (8 GPUs)',c1/8,'Not stated','Compute Resource CUD 1 Year.',G)
    r('Google Cloud',chip,cfg,'reserved/committed','3-yr CUD',c3,'$/VM-hour (8 GPUs)',c3/8,'Not stated','Compute Resource CUD 3 Year.',G)
r('Google Cloud','GB200 NVL72','A4X','on-demand','n/a','Not listed','n/a',None,'n/a','A4X (GB200) not listed on accelerator-optimized pricing page.',G)
# AWS on-demand
A='https://b0.p.awsstatic.com/pricing/2.0/meteredUnitMaps/ec2/USD/current/ec2-ondemand-without-sec-sel/US%20East%20(N.%20Virginia)/Linux/index.json (data behind https://aws.amazon.com/ec2/pricing/on-demand/)'
for chip,cfg,p,n,ic in [('H100','p5.48xlarge (8x H100)',55.04,8,'EFA 3200 Gbps'),('H100','p5.4xlarge (1x H100)',6.88,1,''),('H200','p5en.48xlarge (8x H200)',63.296,8,''),('B200','p6-b200.48xlarge (8x B200)',113.9328,8,'Network 3200 Gigabit'),('B300','p6-b300.48xlarge (8x B300)',142.416,8,'Network 6400 Gigabit')]:
    r('AWS',chip,cfg,'on-demand','pay as you go',p,f'$/instance-hour ({n} GPUs)',p/n,'US East (N. Virginia); same in Ohio',('Linux. '+ic).strip(),A)
r('AWS','H200','p5e.48xlarge','on-demand','n/a','Not listed','n/a',None,'US East (N. Virginia)','p5e has no on-demand meter in us-east-1 file (Capacity Blocks only).',A)
# Vast.ai marketplace (live API)
V='https://console.vast.ai/api/v0/bundles/ (public offer search API; https://vast.ai/pricing is JS-rendered)'
for chip,cfg,n,mn,med,mx in [('H100','H100 SXM, verified, on-demand, rentable',13,1.950,3.580,5.162),('H200','H200 SXM, verified, on-demand',6,4.303,5.062,5.121),('B200','B200, verified, on-demand',30,6.251,7.813,9.385),('B300','B300, verified, on-demand',2,9.940,12.470,15.001)]:
    r('Vast.ai',chip,cfg,'on-demand','per-second, marketplace',med,'$/GPU-hour (median of live offers, dph_total/num_gpus)',med,'Global marketplace',f'Snapshot of {n} live verified offers at ~17:55 ET: min ${mn:.3f}, median ${med:.3f}, max ${mx:.3f}. Computed from listed offers (not a single list price). Interruptible & reserved (1/3/6 mo) tiers exist but not sampled.',V)
r('Vast.ai','GB200 NVL72/MI300X','-','on-demand','n/a','No offers','n/a',None,'Global','0 live verified offers on 2026-10-04.',V)
# No public pricing
r('Nscale','GB300 NVL72 / H100 / H200 / B200','Reservations','reserved/committed','multi-month','Not published','n/a',None,'UK/EU/US','nscale.com/pricing returns 404; docs say GB300 NVL72 reservation pricing shown only in console. Third-party aggregators list numbers but not used (not first-party).','https://docs.nscale.com/docs/compute/reservations')
r('Hyperbolic','H100/H200/B200','On-demand / Reserved','on-demand','hourly','Not visible','n/a',None,'n/a','Homepage describes on-demand and reserved (<1yr) but shows no $ rates; /pricing 404; curl 403.','https://hyperbolic.ai/')
r('FluidStack','H100/H200/B200/GB200','-','reserved/committed','n/a','Not visible','n/a',None,'n/a','fluidstack.io/pricing 404 via fetch / 403 via curl; no public rate card found.','https://www.fluidstack.io/pricing')
