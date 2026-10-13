"""Charts for the speed / residual / credit studies. Run after build_workbook.py."""
import csv, json, sys, datetime as dt
sys.path.insert(0, '/workspace/nebius/skeptic/speed/code')
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt; import numpy as np
from curated import EXTRA_PRICES, EVENTS
B='/workspace/nebius/skeptic/speed/'; C=B+'charts/'
D=lambda s: dt.datetime.strptime(s,'%Y-%m-%d')
ps=list(csv.DictReader(open(B+'data/price_series.csv')))
for p in EXTRA_PRICES: ps.append(dict(provider=p[0],chip=p[1],date=p[2],usd_per_gpu_hr=p[3]))
def ser(prov,chip):
    pts=sorted((D(r['date']),float(r['usd_per_gpu_hr'])) for r in ps if r['provider']==prov and r['chip']==chip)
    return pts
SRC='Sources: Wayback captures of provider pricing pages (/workspace/nebius/wayback_raw), AWS price list, CCIR; see speed_study.xlsx'
# 1 H100 price paths
fig,ax=plt.subplots(figsize=(11,6))
for prov,lab in [('Nebius','Nebius on-demand'),('Lambda','Lambda 8x on-demand'),('CoreWeave','CoreWeave (2024 component / 2025+ instance/8)'),('AWS Capacity Blocks','AWS Capacity Blocks'),('AWS on-demand','AWS P5 on-demand'),('CCIR posted panel median','CCIR panel median')]:
    s=ser(prov,'H100')
    if s: ax.step([a for a,_ in s],[b for _,b in s],where='post',marker='o',ms=3,label=lab)
for e in EVENTS:
    if e[4]=='HS entry': ax.axvline(D(e[1]),color='grey',ls=':',lw=0.8); ax.text(D(e[1]),12.6,e[0],fontsize=7,rotation=90,va='top')
ax.set_ylabel('$/GPU-hr (posted)'); ax.set_title('H100 posted prices vs hyperscaler entry dates (dotted)\nNebius launch premium (+64% vs trough) gone in ~12 months; 2026 rise is market-wide')
ax.legend(fontsize=8); ax.grid(alpha=.3); fig.text(0.01,0.01,SRC,fontsize=7); fig.tight_layout(rect=(0,0.03,1,1)); fig.savefig(C+'s1_h100_price_paths.png',dpi=140); plt.close()
# 2 Nebius all chips
fig,ax=plt.subplots(figsize=(10,5.5))
for chip in ['H100','H200','B200','B300','A100','L40S']:
    s=ser('Nebius',chip)
    if s: ax.step([a for a,_ in s],[b for _,b in s],where='post',marker='o',ms=3,label=chip)
ax.set_title('Nebius on-demand list prices by chip (Wayback + 2026 notices)'); ax.set_ylabel('$/GPU-hr'); ax.legend(); ax.grid(alpha=.3)
fig.text(0.01,0.01,SRC,fontsize=7); fig.tight_layout(rect=(0,0.03,1,1)); fig.savefig(C+'s1_nebius_all_chips.png',dpi=140); plt.close()
# 3 head-start heatmap (same math as S1_HeadStart)
def share(pi0,N,hl=10.2168,L=48,Pss=2.95):
    q=0.5**(1/hl); t=np.arange(L); pi=pi0*q**t; rev=Pss*(1+pi); exc=np.where(t<N,Pss*pi,0); return exc.sum()/rev.sum()
pis=[0.2,0.4,0.64,1.0]; Ns=[1,3,6,12]
M=np.array([[share(p,n) for n in Ns] for p in pis])
fig,ax=plt.subplots(figsize=(7,5)); im=ax.imshow(M*100,cmap='Reds',origin='lower')
for i in range(len(pis)):
    for j in range(len(Ns)): ax.text(j,i,f'{M[i,j]*100:.1f}%',ha='center',va='center')
ax.set_xticks(range(len(Ns)),[str(n) for n in Ns]); ax.set_yticks(range(len(pis)),[f'{p:.0%}' for p in pis])
ax.set_xlabel('head start N (months)'); ax.set_ylabel('launch premium pi0'); ax.set_title('Excess launch-premium revenue as % of per-GPU lifetime revenue\n(half-life 10.2 mo, 48-mo life; fleet-level = x share of fleet online early, ~15%)',fontsize=10)
fig.colorbar(im,label='%'); fig.text(0.01,0.01,'Derived: S1_HeadStart in speed_study.xlsx. Nebius actual head start vs first neocloud is negative (2-7 months behind).',fontsize=7)
fig.tight_layout(rect=(0,0.03,1,1)); fig.savefig(C+'s1_headstart_heatmap.png',dpi=140); plt.close()
# 4 residual vs book
fig,ax=plt.subplots(figsize=(10,5.5)); yrs=np.linspace(0,7,100)
for life,ls in [(4,'--'),(5,'-.'),(6,':')]: ax.plot(yrs,np.clip(1-yrs/life,0,1)*100,ls,color='k',label=f'{life}-yr straight-line book NBV')
mk=[('H100 8-GPU server used (hashrateindex, Apr-26)',3.05,53.5,64.2),('H100 PCIe used vs new (Clustermelt, Sep-26)',3.45,94.9,109.8),('H100 SXM refurb vs $30k (Mercatus, Jun-26)',3.26,60,73.3),('H100 36-mo survey band (Mercatus)',3.26,40,75),('H100 mid-2024 peak scarcity',1.28,153.8,153.8),('A100 SXM4 80GB used vs DGX/8 (Sep-26)',6.3,39.8,40.2),('A100 PCIe 80GB used (Oct-26)',6.38,50.5,75.9),('A100 40GB used/refurb (Aug-Sep 26)',6.3,23.1,35.4)]
for lab,x,lo,hi in mk: ax.errorbar(x,(lo+hi)/2,yerr=[[ (lo+hi)/2-lo],[hi-(lo+hi)/2]],fmt='o',capsize=4,label=lab)
ax.set_xlabel('age since GA (years)'); ax.set_ylabel('% of new price'); ax.set_ylim(0,170); ax.grid(alpha=.3); ax.legend(fontsize=7,loc='upper right')
ax.set_title('Used GPU marks vs book depreciation schedules\nMarks sit ABOVE 4-6 yr book NBV today, but 2026 is a shortage market (regime-dependent)')
fig.text(0.01,0.01,'Sources: hashrateindex 2026-04-08, Clustermelt 2026-09-01, Mercatus 2026-06-23, listings Sep/Oct 2026 (asks, not sold prices). See S2_Residual.',fontsize=7)
fig.tight_layout(rect=(0,0.03,1,1)); fig.savefig(C+'s2_residual_vs_book.png',dpi=140); plt.close()
# 5 rental precedent: A100 (teammate monthly pivot) + CCIR A100/V100
rows=list(csv.reader(open('/workspace/nebius/market/price_history_monthly.csv'))); h=rows[0]
fig,ax=plt.subplots(figsize=(11,6))
for want in ['Lambda | A100 80GB | on-demand | PAYG','Crusoe | A100 80GB | on-demand | PAYG','RunPod | A100 80GB | on-demand | PAYG','CoreWeave | A100 | on-demand | PAYG','Nebius | A100 | on-demand | PAYG']:
    i=h.index(want); pts=[(dt.datetime.strptime(r[0]+'-15','%Y-%m-%d'),float(r[i])) for r in rows[1:] if r[i]]
    ax.plot([a for a,_ in pts],[b for _,b in pts],marker='.',label=want.split(' | ')[0]+' A100 (GPU Market teammate)')
for chip,c in [('A100','tab:red'),('V100','tab:purple'),('H100','tab:green')]:
    s=ser('CCIR posted panel median',chip); ax.plot([a for a,_ in s],[b for _,b in s],'s--',color=c,label=f'CCIR {chip} panel median')
s=ser('AWS on-demand','A100 40GB'); ax.plot([a for a,_ in s],[b for _,b in s],'^-',color='k',label='AWS p4d on-demand /8')
ax.axhline(0.40,color='grey',ls=':'); ax.text(dt.datetime(2019,7,1),0.45,'~\\$0.40 cash-cost floor (Est., Mercatus)',fontsize=8)
ax.set_ylabel('$/GPU-hr'); ax.set_title('Older-generation rental precedent: A100 at age 4-6 yrs still \\$1.5-2.8; V100 \\$0.81 at ~8.5 yrs\n(both above cash floor -> economic life > 6 yrs at posted prices)')
ax.legend(fontsize=7); ax.grid(alpha=.3)
fig.text(0.01,0.01,'Sources: /workspace/nebius/market/price_history_monthly.csv (GPU Market teammate, Wayback); ccir.io/research/rental-rates-by-generation; AWS.',fontsize=7)
fig.tight_layout(rect=(0,0.03,1,1)); fig.savefig(C+'s2_rental_precedent.png',dpi=140); plt.close()
# 6 credit: implied spread vs vol, with CRWV bond G-spreads
cm=json.load(open(B+'data/credit_model_out.json'))['notes']
fig,ax=plt.subplots(figsize=(10,5.5))
for k,v in cm.items():
    xs=sorted(int(x) for x in v['implied_spread_bp_by_vol']); ax.plot(xs,[v['implied_spread_bp_by_vol'][str(x)] for x in xs],marker='o',label=k+' convert (model-implied)')
for lab,g in [('CRWV 9.25% 2030',651),('CRWV 9% 2031',723),('CRWV 9.75% 2031 / 9.625% 2032',837)]: ax.axhline(g,color='grey',ls=':',lw=.8); ax.text(30.5,g+10,f'{lab} bond G+{g if g!=837 else "838/836"}',fontsize=7)
ax.axvspan(55,60,color='orange',alpha=.15); ax.text(55.3,120,'calibrated NBIS\nlong-dated vol\n~57-58%',fontsize=8)
ax.set_xlabel('assumed long-dated equity vol (%)'); ax.set_ylabel('credit spread over UST implied at issue (bp)')
ax.set_title('Convert-implied credit spreads: NBIS ~690-1,000bp at calibrated vol\n= CoreWeave high-yield territory despite far lower leverage (Estimated)')
ax.legend(fontsize=8); ax.grid(alpha=.3)
fig.text(0.01,0.01,'Sources: SEC 6-K pricing (2026-08-19), CRWV 8-K (2026-09-17), bondterminal.com (2026-10-02), Treasury par curve; model in code/credit_model.py.',fontsize=7)
fig.tight_layout(rect=(0,0.03,1,1)); fig.savefig(C+'s3_convert_implied_spread.png',dpi=140); plt.close()
print('charts done')
