import pandas as pd, numpy as np, matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
a=pd.read_csv('market/data/price_snapshot_2026-10-04.csv'); b=pd.read_csv('market/data/price_snapshot_2026-10-08.csv')
key=['provider','chip','config','pricing_type','term','region']
for d in (a,b):
    for k in key: d[k]=d[k].fillna('').astype(str)
m=a.merge(b,on=key,how='outer',suffixes=('_1004','_1008'),indicator=True)
m=m[m.chip.isin(['H100','H200','B200','GB200 NVL72'])&~m.pricing_type.str.contains('superseded')]
m['bucket']=np.where(m.pricing_type.str.startswith('reserved'),'term',np.where(m.pricing_type.str.startswith('spot'),'spot','on-demand'))
m['chg_usd']=m.usd_per_gpu_hr_1008-m.usd_per_gpu_hr_1004; m['chg_pct']=m.chg_usd/m.usd_per_gpu_hr_1004
def st(r):
    if r._merge=='left_only': return 'removed (AWS now one global CB rate)' if r.provider=='AWS' else 'removed'
    if r._merge=='right_only': return 'new'
    if pd.isna(r.usd_per_gpu_hr_1004) and pd.isna(r.usd_per_gpu_hr_1008): return 'no public price both dates'
    return 'CHANGED' if abs(r.chg_usd)>1e-6 else 'unchanged'
m['status']=m.apply(st,axis=1); m['tag']='Reported (list); chg Derived'
out=m[['chip','bucket','provider','config','pricing_type','term','region','usd_per_gpu_hr_1004','usd_per_gpu_hr_1008','chg_usd','chg_pct','status','source_url_1008','tag']].sort_values(['chip','bucket','status','provider'])
summ=out.groupby(['chip','bucket','status']).size().unstack(fill_value=0)
with pd.ExcelWriter('market/data/price_diff_2026-10-08_vs_2026-10-04.xlsx') as w:
    pd.DataFrame({'item':['Scope','Method','Headline','Nebius','Noise','Source'],'text':[
     'H100/H200/B200/GB200 NVL72, on-demand + term (reserved/Capacity Block) + spot, live list prices.',
     'gpu_prices.py live scrape 2026-10-08 08:32 ET vs 2026-10-04 snapshot; joined on provider/chip/config/pricing_type/term/region.',
     'Only material move: AWS EC2 Capacity Blocks +15.0% effective Oct 7 [R]: H100 p5 $5.191->$5.970, H200 p5e $5.970->$6.866, p5en $6.865->$7.895, B200 $12.355->$14.208 (GovCloud $14.801). AWS now one rate in every region, so Tokyo H100 $4.72 and Stockholm H200 $6.24 rows are gone (+26.5% there) [D]. GB200 UltraServer CB $10.582 unchanged. AWS says On-Demand and Savings Plans unchanged.',
     'Nebius list unchanged vs Oct 4 [R]: OD H100 $4.50, H200 $5.40, B200 $8.50 (Oct 1 effective; prior $3.85/$4.50/$7.15), spot $0.79/$0.79/$0.99; GB200/GB300 contact sales. No Rubin list.',
     'Vast.ai marketplace and GCP spot moves are dynamic noise. Neocloud term (Lambda, Crusoe, Together) unchanged. No GB200 change anywhere.',
     'https://aws.amazon.com/ec2/capacityblocks/pricing/ (announcement: effective Oct 7, 2026)']}).to_excel(w,sheet_name='README',index=False)
    out.to_excel(w,sheet_name='Diff',index=False); summ.reset_index().to_excel(w,sheet_name='Summary',index=False)
out.to_csv('market/data/price_diff_2026-10-08_vs_2026-10-04.csv',index=False)
print(out[out.status!='unchanged'][['chip','bucket','provider','region','usd_per_gpu_hr_1004','usd_per_gpu_hr_1008','chg_pct','status']].to_string())

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
        pl('AWS | reserved','AWS Capacity Block (term)',C['aws'])
        pl('Nebius | on-demand','Nebius on-demand',C['neb'],lw=2.8)
        pl('neocloud_od_median','Neocloud on-demand median',C['med'],'--',1.8)
    else:
        pl('Oracle OCI | on-demand','Oracle OCI on-demand',C['oci']); pl('Azure | reserved','Azure reserved',C['az']); pl('CoreWeave | on-demand','CoreWeave on-demand',C['cw'])
        ax.text(0.02,0.90,'Nebius GB200: "contact sales", no list price',transform=ax.transAxes,fontsize=9,color=C['neb'],style='italic')
    ax.set_title(t,fontsize=13,fontweight='bold',loc='left'); ax.set_ylabel('$ / GPU-hr',fontsize=9)
    ax.grid(alpha=.25); ax.spines[['top','right']].set_visible(False); ax.set_ylim(bottom=0)
    ax.legend(fontsize=8,frameon=False,loc='lower left'); ax.set_xlim(pd.Timestamp('2025-01-01'),pd.Timestamp('2026-12-10'))
fig.suptitle('GPU list prices turned up into October: a demand signal, not contract pricing',fontsize=16,fontweight='bold',x=0.01,ha='left')
fig.text(0.01,0.905,'Nebius raised H100 / H200 / B200 on-demand by 17% / 20% / 19% (effective Oct 1). AWS raised Capacity Blocks 15% (Oct 7). Monthly last-observed list price,\nJan 2025 to Oct 8, 2026.',fontsize=10.5,ha='left',color='#333')
fig.text(0.01,0.008,'Source: provider pricing pages (live scrape Oct 8, 2026) and Wayback Machine history. Neocloud median covers Nebius, CoreWeave, Lambda,\nCrusoe, Together AI and RunPod where listed. List prices Reported; medians Derived; forward-filled between observations.',fontsize=7.5,color='#666',ha='left')
plt.tight_layout(rect=[0,0.045,1,0.885]); plt.savefig('market/charts/gpu_list_price_monthly_2026-10-08.png',facecolor='white')
