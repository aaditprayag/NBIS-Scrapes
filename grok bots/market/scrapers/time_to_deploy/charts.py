"""Deck charts: lag by generation and post-launch price path."""
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, numpy as np
import analyze as A
from data_curated import GENS
OUT='/workspace/nebius/market/charts/'
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.spines.top':False,'axes.spines.right':False})
COL={'Nebius':'#E4572E','CoreWeave':'#1F3864','Lambda':'#7F7F7F','Crusoe':'#A5A5A5','AWS':'#FF9900','Azure':'#0078D4','GCP':'#34A853','Oracle OCI':'#C74634','Nscale':'#8E7CC3','Together AI':'#B4C7E7'}
PROV=['Nebius','CoreWeave','Lambda','Crusoe','Oracle OCI','AWS','Azure','GCP','Nscale','Together AI']
L=A.lag_table()
fig,ax=plt.subplots(figsize=(13,6.2)); w=0.085
gl=[g for g in GENS]
for i,p in enumerate(PROV):
    xs=[];ys=[];cs=[]
    for k,g in enumerate(gl):
        r=next((r for r in L if r['gen']==g and r['provider']==p),None)
        if not r: continue
        x=k+(i-len(PROV)/2+0.5)*w
        if r['lag'] is not None:
            ax.bar(x,r['lag'],w,color=COL[p],edgecolor='black' if p=='Nebius' else 'none',lw=1.2,label=p if k==next(j for j,gg in enumerate(gl) if any(rr['provider']==p for rr in L if rr['gen']==gg)) else None)
            if p=='Nebius': ax.text(x,r['lag']+8,f"{r['lag']}",ha='center',fontsize=8,fontweight='bold',color=COL[p])
        elif p in ('Nebius','CoreWeave') or g=='Vera Rubin NVL72':
            ax.bar(x,r['censored_min'],w,color='white',edgecolor=COL[p],hatch='///',lw=0.8)
handles,labels=ax.get_legend_handles_labels()
from matplotlib.patches import Patch
handles.append(Patch(facecolor='white',edgecolor='grey',hatch='///')); labels.append('Not yet GA (min. lag to 10/4/26)')
ax.legend(handles,labels,ncol=4,frameon=False,fontsize=8.5,loc='upper right')
ax.set_xticks(range(len(gl))); ax.set_xticklabels([g.replace(' / ','/\n') for g in gl])
ax.set_ylabel('Days from NVIDIA reference GA to provider GA'); ax.set_ylim(0,620)
ax.set_title('Time-to-deploy: Nebius trails CoreWeave on every NVIDIA generation since H200',loc='left',fontweight='bold',fontsize=13)
ax.grid(axis='y',alpha=.25)
clip=[r for r in L if r['lag'] and r['lag']>ax.get_ylim()[1]]
for r in clip: ax.text(gl.index(r['gen'])+(PROV.index(r['provider'])-4.5)*w+0.02,ax.get_ylim()[1]*0.97,f"{r['provider']}: {r['lag']}d (clipped)",fontsize=7,ha='left',va='top',rotation=90)
fig.text(0.01,0.01,'Provider GA = earliest public GA / first public price (press releases, blogs, Wayback captures of pricing pages; upper bounds at ~monthly capture cadence). NVIDIA reference GA: H100 full production 9/20/22; H200 Q2-24 (est.); Blackwell\n'
 'start of Q4 FY25 ramp 10/28/24 (est.); B300/GB300 start of Q2 FY26 4/28/25 (est.); Vera Rubin 8/1/26 (est., Q2 FY27 call). Sources: NVIDIA newsroom & earnings calls; provider newsrooms; web.archive.org. Full source list: time_to_deploy.xlsx.',fontsize=7,color='#555')
plt.tight_layout(rect=(0,0.05,1,1)); fig.savefig(OUT+'time_to_deploy_lag_by_generation.png',dpi=200); plt.close()
# ---- price path after launch
S=A.series(A.load_obs())
SER=[('Nebius','H100'),('Nebius','H200'),('Nebius','B200'),('Nebius','B300'),('CoreWeave','H100'),('CoreWeave','B200'),('CoreWeave','GB200 NVL72'),
     ('Lambda','H100 PCIe'),('Lambda','H100'),('Lambda','B200'),('Crusoe','H200')]
fig,axs=plt.subplots(1,2,figsize=(14,6),gridspec_kw={'width_ratios':[1.25,1]})
mk={'H100':'o','H100 PCIe':'v','H200':'s','B200':'D','B300':'^','GB200 NVL72':'P'}
for prov,chip in SER:
    ser=[o for o in S[(prov,chip)] if o[1] is not None]
    if prov=='CoreWeave' and chip=='H100': ser=[o for o in ser if o[0].year<2025]  # drop definitional change
    t0=ser[0][0]
    xs=[];ys=[]
    for o in ser: xs.append(o[0]); ys.append(o[1])
    last=S[(prov,chip)][-1][0]; xs.append(last); ys.append(ys[-1])
    kw=dict(color=COL[prov],lw=2.6 if prov=='Nebius' else 1.4,marker=mk[chip],ms=4,alpha=1 if prov=='Nebius' else .8,drawstyle='steps-post')
    axs[0].plot(xs,ys,label=f'{prov} {chip}',**kw)
    m=[(x-t0).days/30.44 for x in xs]; axs[1].plot(m,[100*y/ys[0] for y in ys],**kw)
axs[0].set_title('On-demand list $/GPU-hr (Wayback captures)',loc='left',fontweight='bold')
axs[0].set_ylabel('$ / GPU-hour'); axs[0].legend(fontsize=7.5,frameon=False,ncol=2); axs[0].grid(alpha=.25)
axs[0].axvspan(np.datetime64('2026-01-01'),np.datetime64('2026-10-04'),color='#E4572E',alpha=.06); axs[0].text(np.datetime64('2026-01-10'),1.75,'2026: list prices rising',fontsize=8,color='#E4572E')
axs[1].axhline(100,color='black',lw=.6); axs[1].set_xlim(0,36)
axs[1].set_title('Indexed to first listed price (=100) vs months since listing',loc='left',fontweight='bold'); axs[1].set_xlabel('Months since chip first priced on provider page'); axs[1].grid(alpha=.25)
fig.suptitle('Launch pricing: no consistent Nebius launch premium; 2023-25 decay reversed by 2026 price hikes',x=0.01,ha='left',fontweight='bold',fontsize=13)
fig.text(0.01,0.01,'Sources: Wayback Machine captures (web.archive.org, id_ raw) of nebius.ai/prices, nebius.com/prices, coreweave.com/pricing & /gpu-cloud-pricing, lambdalabs.com/service/gpu-cloud, lambda.ai/pricing, crusoe.ai/cloud/pricing. Step lines = price in effect between captures\n'
 '(changes are dated to first capture showing them). CoreWeave H100 shown only through 2024 (2025 page moved to bundled instance pricing). Nebius 10/1/26 prices announced on page from ~9/23/26. GB200/GB300 at Nebius & CoreWeave GB300 = "contact sales" (not shown).',fontsize=7,color='#555')
plt.tight_layout(rect=(0,0.06,1,0.95)); fig.savefig(OUT+'launch_price_decay.png',dpi=200); plt.close()
print('ok')
