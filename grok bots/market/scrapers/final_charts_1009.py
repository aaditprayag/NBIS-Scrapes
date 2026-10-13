import pandas as pd, matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
from matplotlib.patches import Patch
# ---------- 1. same-chip list price history (final) ----------
def load(f):
    W=pd.read_excel(f'parts/{f}_monthly.xlsx',sheet_name='Wide_FFill'); M=pd.read_excel(f'parts/{f}_monthly.xlsx',sheet_name='Medians')
    W=W.merge(M[['year_month','neocloud_od_median']],on='year_month'); W=W[W.year_month>='2025-01'].copy(); W['d']=pd.to_datetime(W.year_month); return W
fig,axs=plt.subplots(2,2,figsize=(13.33,7.2),dpi=150); axs=axs.ravel()
C={'neb':'#0B5FFF','med':'#7A7A7A','aws':'#FF9900','cw':'#2CA58D','oci':'#C74634','az':'#6B4FBB'}
for ax,(f,t) in zip(axs,[('h100_sxm','H100 SXM'),('h200','H200'),('b200','B200'),('gb200_nvl72','GB200 NVL72')]):
    W=load(f)
    def pl(col,lab,c,ls='-',lw=2.2):
        if col in W and W[col].notna().any():
            ax.plot(W.d,W[col],ls,color=c,lw=lw,label=lab); v=W[col].dropna().iloc[-1]
            ax.annotate(f"${v:.2f}",(W.d.iloc[-1],v),xytext=(5,-11 if lab.startswith('Neocloud') else 0),textcoords='offset points',color=c,fontsize=9.5,fontweight='bold',va='center')
    if f!='gb200_nvl72':
        pl('AWS | reserved','AWS Capacity Block (term)',C['aws']); pl('Nebius | on-demand','Nebius on-demand',C['neb'],lw=2.8); pl('neocloud_od_median','Neocloud on-demand median',C['med'],'--',1.8)
    else:
        pl('Oracle OCI | on-demand','Oracle OCI on-demand',C['oci']); pl('Azure | reserved','Azure reserved',C['az']); pl('CoreWeave | on-demand','CoreWeave on-demand',C['cw'])
        ax.text(0.02,0.90,'Nebius GB200: "contact sales", no list price',transform=ax.transAxes,fontsize=9,color=C['neb'],style='italic')
    ax.set_title(t,fontsize=13,fontweight='bold',loc='left'); ax.set_ylabel('$ / GPU-hr',fontsize=9)
    ax.grid(alpha=.25); ax.spines[['top','right']].set_visible(False); ax.set_ylim(bottom=0)
    ax.legend(fontsize=8,frameon=False,loc='lower left'); ax.set_xlim(pd.Timestamp('2025-01-01'),pd.Timestamp('2026-12-10'))
fig.suptitle('GPU list prices turned up into October: a demand signal, not contract pricing',fontsize=16,fontweight='bold',x=0.01,ha='left')
fig.text(0.01,0.905,'Same chip, same pricing type. Nebius raised H100 / H200 / B200 on-demand by 17% / 20% / 19% (effective Oct 1); unchanged at the Oct 9 re-check. \nAWS raised Capacity Blocks 15% (Oct 7). On-demand + Token Factory is ~4% of Nebius revenue on our estimate (range 1–9%) [E], so list prices do not move the contract band.',fontsize=10,ha='left',color='#333')
fig.text(0.01,0.008,'Source: provider pricing pages (live scrape Oct 8, re-checked Oct 9, 2026) and Wayback Machine history. Monthly last-observed list price, Jan 2025 to Oct 2026. Neocloud median covers Nebius, CoreWeave, Lambda, Crusoe,\nTogether AI and RunPod where listed; RunPod\'s Oct 9 visible increase is unconfirmed (page structured data still shows Oct 8 prices) and is excluded. List prices Reported; medians Derived; forward-filled between observations.',fontsize=7.3,color='#666',ha='left')
plt.tight_layout(rect=[0,0.05,1,0.875]); plt.savefig('market/charts/gpu_list_price_history_final.png',facecolor='white'); plt.close()

# ---------- 2. contract band slide ----------
pr=[('Microsoft–IREN (GB300)\n200 IT MW reported',9.70,'R/D','#2C5F8A',False),
    ('Meta–Nebius #1\nMW not disclosed (~50 est.)',11.52,'E','#C0392B',True),
    ('Microsoft–Nscale Narvik (GB300)\nIT MW from GPU count',11.71,'D','#2C5F8A',False),
    ('Microsoft–Nscale all SOWs\nIT MW from GPU counts',13.35,'D, term P','#2C5F8A',False),
    ('Microsoft–Nebius (Vineland)\n240 IT MW team estimate',14.49,'E, house','#C0392B',False),
    ('OpenAI–Oracle Stargate\n4.5 GW unlabeled, ÷1.2 if gross',16.00,'P (IT read E)','#8E9AAF',True),
    ('Anthropic–Nscale (Vera Rubin)†\n460 IT MW reported',16.16,'R/D, term P','#6B4FBB',False),
    ('Anthropic–Lambda\n~350 MW, basis unstated',16.67,'P','#2C5F8A',False)]
fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=150)
y=list(range(len(pr)))[::-1]
ax.axvspan(9.7,16.7,color='#DCE8F5',zorder=0)
for yi,(lab,v,tag,c,h) in zip(y,pr):
    ax.barh(yi,v,color='white' if h else c,edgecolor=c,hatch='///' if h else None,lw=1.5,height=0.62,zorder=2)
    ax.text(v+0.25,yi,f'\\${v:.2f}M  [{tag}]',va='center',fontsize=10,fontweight='bold',zorder=3)
ax.set_yticks(y); ax.set_yticklabels([p[0] for p in pr],fontsize=9.5)
for x,lab,c,yy in [(22.68,'Clear spot,\nhigh path\n\\$22.68M [D]','#E67E22',7.3),(24.51,'Clear spot,\nbase path\n\\$24.51M [D]','#D35400',5.3),(27.6,'Street\nnon-MSFT\nhurdle\n~\\$27.6M [D]','#B03A2E',3.3)]:
    ax.axvline(x,color=c,ls='--',lw=1.8,zorder=1); ax.text(x+0.15,yy,lab,color=c,fontsize=9,va='top',fontweight='bold',bbox=dict(fc='white',ec='none',pad=1))
# above-band, no disclosed MW (markers only, not prints)
ax.plot([20,25],[-1.05,-1.05],color='#555',lw=6,solid_capstyle='butt',alpha=.55)
ax.text(22.5,-0.62,'Nebius Q2’26 four deals \\$20–25M/MW [R]',fontsize=8.5,color='#444',va='center',ha='center',bbox=dict(fc='white',ec='none',pad=1),zorder=4); ax.text(25.3,-1.05,'no MW, basis or\nterm disclosed',fontsize=8,color='#444',va='center')
ax.plot([22.4],[-2.05],marker='<',ms=10,color='#555')
ax.text(21.9,-2.05,'Bitdeer A201 ≤\\$22.4M ceiling [D, basis P], ~15 MW (not a print)',fontsize=8.5,color='#444',va='center',ha='right')
ax.set_ylim(-2.5,len(pr)-0.4); ax.set_xlim(0,30)
ax.text(13.2,-0.75,'Contract band: \\$9.7–16.7M as reported',ha='center',va='top',fontsize=10.5,color='#1F4E79',fontweight='bold')
ax.set_xlabel('$M of contract revenue per MW per year (TCV ÷ term ÷ MW)',fontsize=10)
ax.grid(axis='x',alpha=.25); ax.spines[['top','right']].set_visible(False)
fig.suptitle('Every disclosed contract print sits below the price the Street needs from Nebius',fontsize=16,fontweight='bold',x=0.01,ha='left')
fig.text(0.01,0.885,'5–6 year full-stack GPU contracts with stated or normalized MW.\nManagement\'s four Q2 deals yield \\$20–25M per MW; the Street needs ~\\$27.6M per IT MW from the whole non-MSFT book on our MW.',fontsize=10.3,color='#333',ha='left')
fig.text(0.01,0.012,'Tags: R reported, D derived from reported inputs, E our estimate, P press/secondary source. Hatched = MW estimated. Hurdles are per IT MW; Lambda and Bitdeer MW bases unstated, so the band is "as reported". † Monarch campus civil-only (Grade Yellow).\nSources: IREN 8-K 11/3/25; Aker-Nscale and Nscale releases; Nscale S-1 9/18/26; Nebius 6-Ks 9/8/25, 11/12/25, Q2\'26 letter 8/12/26 p.3–4; WSJ/DCD (Stargate); Needham 10/2 note (Lambda, Bitdeer); Bitdeer releases. \nDetail: contract_economics.xlsx Print_Table.',fontsize=7.3,color='#666',ha='left')
plt.tight_layout(rect=[0,0.055,1,0.87]); plt.savefig('market/charts/contract_band_slide_final.png',facecolor='white')
