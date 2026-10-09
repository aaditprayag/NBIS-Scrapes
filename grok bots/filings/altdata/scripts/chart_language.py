import json,os
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
plt.rcParams.update({'font.size':10,'text.parse_math':False})
A='/workspace/nebius/filings/altdata/'
R=json.load(open(A+'scripts/language_results.json'))
Q=['Q1-24','Q2-24','Q3-24','Q4-24','Q1-25','Q2-25','Q3-25','Q4-25','Q1-26','Q2-26']
fam=[('supply_constrained_per10k','Supply-constrained / sold out','#c00000'),('demand_pipeline_per10k','Demand / pipeline','#1f3864'),('pricing_competition_per10k','Pricing / competition','#ed7d31')]
fig,ax=plt.subplots(1,2,figsize=(13,4.8),sharey=True)
for a,kind,title in [(ax[0],'call','Earnings calls'),(ax[1],'letter','Shareholder letters (hollow = press release; no letter issued)')]:
    docs={d['q']:d for d in R if d['kind'].strip('*')==kind or (kind=='letter' and d['kind']=='release*')}
    for k,lab,c in fam:
        xs=[i for i,q in enumerate(Q) if q in docs]; ys=[docs[Q[i]][k] for i in xs]
        a.plot(xs,ys,'o-',color=c,lw=2,label=lab)
        if kind=='letter':
            pr=[i for i in xs if docs[Q[i]]['kind']=='release*']
            a.plot(pr,[docs[Q[i]][k] for i in pr],'o',mfc='white',color=c)
    a.set_xticks(range(len(Q))); a.set_xticklabels(Q,rotation=45); a.set_title(title,loc='left',fontweight='bold')
    for sp in ['top','right']: a.spines[sp].set_visible(False)
    a.grid(axis='y',alpha=.3)
ax[0].set_ylabel('Mentions per 10,000 words'); fig.legend(*ax[0].get_legend_handles_labels(),frameon=False,loc='lower center',ncol=3,bbox_to_anchor=(0.5,0.035))
ax[0].annotate('Q3-25: "sold out" x9,\n"bottleneck" x6',xy=(6,31.4),xytext=(2.6,44),arrowprops=dict(arrowstyle='->',color='#c00000'),color='#c00000',fontsize=8)
ax[0].annotate('Q2-26: pricing language\nat record; mostly price-UP',xy=(9,51.3),xytext=(7.0,62),arrowprops=dict(arrowstyle='->',color='#ed7d31'),color='#ed7d31',fontsize=8)
fig.suptitle('Nebius management language: from "sold out" to "pricing power"',x=0.01,ha='left',fontweight='bold')
fig.text(0.01,0.01,'No calls Q1-24/Q2-24 (Yandex divestment, trading halt). Phrase dictionary in altdata.xlsx Language sheet. Sources: SEC EDGAR 6-K ex 99.1/99.2; transcripts via StockAnalysis.com (Quartr).',fontsize=7,color='#555')
plt.tight_layout(rect=(0,0.1,1,0.95)); plt.savefig(A+'language_trends.png',dpi=200); print('ok')
