import csv,statistics,collections
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
SRC='/workspace/nebius/price_snapshot_2026-10-04.csv'; OUT='/workspace/nebius/market/charts/gpu_price_ondemand_2026-10-04.png'
HS={'AWS','Azure','Google Cloud','Oracle OCI'}
rows=[r for r in csv.DictReader(open(SRC)) if r['pricing_type']=='on-demand' and r['usd_per_gpu_hr']]
chips=['H100','H200','B200','B300']
pv=collections.defaultdict(lambda: collections.defaultdict(list))
for r in rows:
    if r['chip'] in chips: pv[r['chip']][r['provider']].append(float(r['usd_per_gpu_hr']))
data={}
for c in chips:
    p={k:statistics.median(v) for k,v in pv[c].items()}
    neo=[v for k,v in p.items() if k!='Nebius' and k not in HS]; hs=[v for k,v in p.items() if k in HS]
    data[c]=(p.get('Nebius'),statistics.median(neo),statistics.median(hs),p)
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':11})
fig,ax=plt.subplots(figsize=(12,6.9),dpi=200)
w=0.26; cols=['#1F3864','#8FAADC','#BFBFBF']; lab=['Nebius (list, eff. Oct 1 2026)','Neocloud peer median','Hyperscaler median (AWS/Azure/GCP/OCI)']
for i,c in enumerate(chips):
    vals=data[c][:3]
    for j,v in enumerate(vals):
        x=i+(j-1)*w; b=ax.bar(x,v,w*0.95,color=cols[j],label=lab[j] if i==0 else None,zorder=2)
        ax.text(x,v+0.2,f'${v:.2f}',ha='center',va='bottom',fontsize=9.5,fontweight='bold' if j==0 else 'normal')
    nb,nm,hm,_=data[c]
ax.set_xticks(range(len(chips))); ax.set_xticklabels([f'NVIDIA {c}\nNebius {(data[c][0]/data[c][1]-1)*100:+.0f}% vs neocloud\n{(data[c][0]/data[c][2]-1)*100:+.0f}% vs hyperscaler' for c in chips],fontsize=10.5)
ax.tick_params(axis='x',pad=6); ax.yaxis.set_major_locator(matplotlib.ticker.MultipleLocator(2))
ax.set_ylabel('On-demand list price, $ per GPU-hour'); ax.set_ylim(0,max(d[2] for d in data.values())*1.15)
ax.yaxis.set_major_formatter(matplotlib.ticker.StrMethodFormatter('${x:,.0f}'))
for s in ['top','right']: ax.spines[s].set_visible(False)
ax.grid(axis='y',color='#E5E5E5',zorder=0); ax.legend(frameon=False,loc='upper left',fontsize=10)
ax.set_title('Nebius: 7-14% above neocloud median (H100-B200), 40-58% below hyperscalers\nPublic on-demand GPU cloud list prices, snapshot Oct 4, 2026',loc='left',fontsize=13.5,fontweight='bold')
peers=lambda c:', '.join(f'{k} {v:.2f}' for k,v in sorted(data[c][3].items(),key=lambda z:z[1]) if k!='Nebius')
foot=('Notes: provider value = median of that provider\'s on-demand SKUs (e.g., 1x-8x sizes), converted to $/GPU-hr from per-node quotes (e.g., CoreWeave $49.24 per 8x H100 node-hr). '
 'Neocloud peers: CoreWeave, Lambda, Crusoe, Together AI, RunPod, Vast.ai (median of live verified offers). Excludes contact-sales SKUs, spot, and commitments; Nebius commitments up to 35% below list. '
 'Nebius list price changed Oct 1, 2026 (H100 $3.85 -> $4.50).\n'
 'Sources (retrieved 2026-10-04): nebius.com/prices; coreweave.com/pricing; lambda.ai/pricing; crusoe.ai/cloud/pricing; together.ai/pricing; runpod.io/pricing; console.vast.ai API; '
 'AWS EC2 on-demand price list (us-east-1); Azure Retail Prices API (eastus/eastus2); cloud.google.com accelerator-optimized pricing; Oracle OCI price list API.')
import textwrap
fig.text(0.01,0.01,'\n'.join(textwrap.fill(p,215) for p in foot.split('\n')).replace('$','\\$'),fontsize=6.8,color='#555555',va='bottom')
plt.subplots_adjust(left=0.07,right=0.98,top=0.86,bottom=0.27)
fig.savefig(OUT); print(OUT); 
for c in chips: print(c,data[c][:3],peers(c))
