import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt, json
src=open('model.py').read().split('if __name__')[0]
def run(**ov):
    g={}; s=src
    exec(s,g)
    if ov:
        g['A'].update(ov); g2={}; 
        code=s.replace("A=dict(","A0=dict(").replace("rv=[904.1","A=A0; A.update(OV)\nrv=[904.1")
        g2['OV']=ov; exec(code,g2); return g2
    return g
base=run()
out='/workspace/nebius/skeptic/charts'; import os; os.makedirs(out,exist_ok=True)
Q=base['Q']; x=range(len(Q))
# 1 quarterly EBITDA vs EBIT
plt.figure(figsize=(11,5.2))
plt.plot(x,base['eb'],'k-',lw=2.5,label='Adj. EBITDA (consensus to FY28, extrapolated after)')
for L,c in zip([4,5,6],['#C00000','#1F3864','#70AD47']):
    plt.plot(x,base['D'][L]['ebit'],color=c,lw=2,label=f'EBIT (adj.) @ {L}-yr GPU life'+(' - Nebius policy' if L==5 else ' - old policy' if L==4 else ' - CRWV-style'))
plt.axhline(0,color='grey',lw=.8); plt.xticks(x,Q,rotation=60,fontsize=8); plt.ylabel('$M per quarter')
plt.title('NBIS: EBITDA vs EBIT under 4/5/6-yr GPU useful lives (Q3 2026 to Q4 2030)\nSame cash, different reported earnings')
plt.legend(fontsize=8); plt.grid(alpha=.3); plt.tight_layout(); plt.savefig(f'{out}/03a_ebit_by_life_quarterly.png',dpi=150); plt.close()
# 2 annual bars
yrs=[2026,2027,2028,2029,2030]; fy=base['fy']
ebA=[fy(base['eb'],y)+(365.7 if y==2026 else 0) for y in yrs]
import numpy as np; w=0.16; X=np.arange(5)
plt.figure(figsize=(11,5.2))
plt.bar(X-2*w,ebA,w,color='k',label='Adj. EBITDA')
for i,(L,c) in enumerate(zip([4,5,6],['#C00000','#1F3864','#70AD47'])):
    e=[fy(base['D'][L]['ebit'],y)+((365.7-base['H1']['da'][L]) if y==2026 else 0) for y in yrs]
    plt.bar(X+(i-1)*w,e,w,color=c,label=f'EBIT @ {L}-yr')
plt.bar(X+2*w,[fy(base['fcf'],y) for y in yrs],w,color='#ED7D31',label='FCF (EBITDA-capex-interest; FY26=H2)')
plt.xticks(X,[f'FY{y}' for y in yrs]); plt.axhline(0,color='grey',lw=.8); plt.ylabel('$M'); plt.legend(fontsize=8); plt.grid(alpha=.3,axis='y')
plt.title('NBIS annual: EBITDA vs EBIT (4/5/6-yr) vs FCF'); plt.tight_layout(); plt.savefig(f'{out}/03b_annual_ebitda_ebit_fcf.png',dpi=150); plt.close()
# 3 reverse DCF sensitivity: value/share vs FY29+ EBITDA yield
ys=[0.25,0.28,0.30,0.35,0.40,0.45,0.50,0.55,0.60]; sens={}
plt.figure(figsize=(10,5.2))
for L,c in zip([4,5,6],['#C00000','#1F3864','#70AD47']):
    v=[run(yov=y)['rdcf'][L]['ps'] for y in ys]; sens[L]=v
    plt.plot([y*100 for y in ys],v,'o-',color=c,label=f'{L}-yr economic life')
plt.axhline(242.81,color='k',ls='--',label='Price $242.81 (10/2/26)'); plt.axhline(0,color='grey',lw=.8)
for y,t,yy in [(base['yld_c'],'FY28 consensus-\nimplied ~27.6%',-470),(0.308,'Installed base\n$12M/MW ~30.8%',-300),(0.513,'Q2 deals\n$20M/MW ~51%',-300)]:
    plt.axvline(y*100,color='grey',ls=':'); plt.text(y*100+0.4,yy,t,fontsize=7)
plt.xlabel('Steady-state EBITDA yield on gross GPU capital, FY2029+ (%)'); plt.ylabel('DCF value per share ($)')
plt.title('Reverse DCF: what the price needs (WACC 11%, g 3%, capex/MW $30M)\nPrice requires ~44-55% fleet-wide yield = ~$17-21M revenue/MW for the asset life')
plt.legend(fontsize=8,loc='upper left'); plt.grid(alpha=.3); plt.tight_layout(); plt.savefig(f'{out}/03c_reverse_dcf_yield.png',dpi=150); plt.close()
# WACC sensitivity at Q2-deal yield 0.513
ws=[0.09,0.10,0.11,0.12]; wsens={L:[round(run(yov=0.513,wacc=wc)['rdcf'][L]['ps']) for wc in ws] for L in [4,5,6]}
req={wc:{L:round(run(wacc=wc)['rdcf'][L]['yreq'],3) for L in [4,5,6]} for wc in ws}
# 4 D&A by cohort at 5y (stack)
L=5; R=base['D'][L]
plt.figure(figsize=(11,5))
plt.stackplot(x,R['ex'],R['new'],R['ng'],labels=['6/30/26 in-service base (Hopper/early Blackwell)','New GPU cohorts Q3 26+ (Blackwell -> Vera Rubin)','Non-GPU DC/infra'],colors=['#A5A5A5','#1F3864','#FFC000'])
plt.plot(x,base['eb'],'k--',label='Adj. EBITDA'); plt.xticks(x,Q,rotation=60,fontsize=8); plt.ylabel('$M per quarter'); plt.legend(fontsize=8,loc='upper left')
plt.title('D&A by cohort at Nebius 5-yr policy vs EBITDA'); plt.grid(alpha=.3); plt.tight_layout(); plt.savefig(f'{out}/03d_da_by_cohort_5yr.png',dpi=150); plt.close()
json.dump(dict(ys=ys,sens={k:[round(v) for v in vs] for k,vs in sens.items()},ws=ws,wsens=wsens,req={str(k):v for k,v in req.items()},yld_c=base['yld_c']),open('sens.json','w'),indent=1)
print(json.dumps(dict(sens={k:[round(v) for v in vs] for k,vs in sens.items()},wsens=wsens,req=req)))
