import sys,os
sys.path.insert(0,os.path.dirname(__file__))
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from guidance_data import G
from prepay_data import Q,R
OUT='/workspace/nebius/filings/altdata/'
plt.rcParams.update({'font.size':10,'font.family':'DejaVu Sans','text.parse_math':False})
# --- Guidance hit rate chart ---
cats=['ARR','Revenue','Capex','MW','GPU']
st=['BEAT','HIT','SLIPPED','MISS','PENDING','NOT VERIFIABLE']
col={'BEAT':'#2e7d32','HIT':'#81c784','SLIPPED':'#ffb300','MISS':'#c62828','PENDING':'#90caf9','NOT VERIFIABLE':'#bdbdbd'}
fig,ax=plt.subplots(1,2,figsize=(12,4.8),gridspec_kw={'width_ratios':[1.6,1]})
left=[0]*len(cats)
for s in st:
    v=[sum(1 for g in G if g['cat']==c and g['status']==s) for c in cats]
    ax[0].barh(cats,v,left=left,color=col[s],label=s.title())
    left=[a+b for a,b in zip(left,v)]
ax[0].invert_yaxis(); ax[0].set_xlabel('# forward-looking statements'); ax[0].legend(ncol=3,fontsize=8,loc='lower right',frameon=False)
ax[0].set_title('Nebius guidance outcomes by metric (Oct-2024 to Aug-2026)',loc='left',fontweight='bold')
for sp in ['top','right']: ax[0].spines[sp].set_visible(False)
res=[g for g in G if g['status'] in('HIT','BEAT','MISS','SLIPPED')]
rate=[]
for c in cats:
    rr=[g for g in res if g['cat']==c]; rate.append(100*sum(g['status'] in('HIT','BEAT') for g in rr)/len(rr) if rr else 0)
allr=100*sum(g['status'] in('HIT','BEAT') for g in res)/len(res)
b=ax[1].bar(cats+['All'],rate+[allr],color=['#1f3864']*5+['#c00000'])
for x,v,n in zip(b,rate+[allr],[sum(1 for g in res if g['cat']==c) for c in cats]+[len(res)]):
    ax[1].text(x.get_x()+x.get_width()/2,v+2,f'{v:.0f}%\n(n={n})',ha='center',fontsize=8)
ax[1].set_ylim(0,120); ax[1].set_ylabel('Hit/beat rate of resolved statements (%)'); ax[1].set_title('Hit rate (resolved only)',loc='left',fontweight='bold')
for sp in ['top','right']: ax[1].spines[sp].set_visible(False)
fig.text(0.01,0.01,'Misses: Dec-24 ARR ($90M vs $170-190M), FY24 revenue & capex, FY25 capex ($4.07B vs ~$5B guided Nov-25). Source: Nebius 6-K/20-F (SEC EDGAR), call transcripts (StockAnalysis/Quartr). See altdata.xlsx Guidance_HitRate.',fontsize=7,color='#555')
plt.tight_layout(rect=(0,0.04,1,1)); plt.savefig(OUT+'guidance_hitrate.png',dpi=200); plt.close()
# --- Prepay chart ---
d={m:v for m,v,_,_ in R}
dr=d['Deferred revenue (contract liability), total']; rpo=d['Remaining performance obligations (RPO)']; pp=d['Property & equipment, net']
qs=Q[1:]
fig,ax=plt.subplots(figsize=(11,5))
x=range(len(qs))
ax.bar([i-0.2 for i in x],[pp.get(q,0)/1000 for q in qs],0.4,color='#9e9e9e',label='PP&E, net ($B)')
ax.bar([i+0.2 for i in x],[dr.get(q,0)/1000 for q in qs],0.4,color='#1f3864',label='Deferred revenue / customer prepayments ($B)')
for i,q in enumerate(qs):
    if dr.get(q,0)>1000: ax.text(i+0.2,dr[q]/1000+0.2,f'{dr[q]/1000:.1f}',ha='center',fontsize=8,color='#1f3864')
ax.set_xticks(list(x)); ax.set_xticklabels(qs); ax.set_ylabel('$B (bars)')
ax2=ax.twinx(); rq=[q for q in qs if q in rpo]
ax2.plot([qs.index(q) for q in rq],[rpo[q]/1000 for q in rq],'o-',color='#c00000',lw=2,label='RPO / backlog ($B, rhs)')
for q in rq: ax2.text(qs.index(q),rpo[q]/1000+1.2,f'${rpo[q]/1000:.1f}B',ha='center',color='#c00000',fontsize=8)
ax2.set_ylim(0,45); ax2.set_ylabel('RPO $B')
h1,l1=ax.get_legend_handles_labels(); h2,l2=ax2.get_legend_handles_labels(); ax.legend(h1+h2,l1+l2,loc='upper left',frameon=False)
ax.set_title('Customer prepayments now fund ~46% of Nebius PP&E; backlog $37.5B at Q2-26',loc='left',fontweight='bold')
fig.text(0.01,0.01,'Deferred revenue ~$15-19M through Q3-25 (pre-hyperscaler). Q2-26: deferred rev $5.98B vs PP&E $13.05B; energy commitments $5.3B (vs $29M at YE25); unstarted leases $12.1B. Source: Nebius 6-K/20-F, SEC EDGAR.',fontsize=7,color='#555')
plt.tight_layout(rect=(0,0.04,1,1)); plt.savefig(OUT+'prepay_commitments.png',dpi=200); plt.close()
print('charts ok')
