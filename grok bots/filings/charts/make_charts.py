"""Deck-ready charts for the NBIS pitch. Data = values in nebius_filings.xlsx (Quarterly / CapStructure tabs),
which cite SEC filings cell by cell. Run: python3 make_charts.py [price]"""
import sys, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
OUT='/workspace/nebius/filings/charts/'
NAVY='#1F3A5F'; TEAL='#2A9D8F'; GREY='#8D99AE'; ORANGE='#E76F51'; GOLD='#E9C46A'; LIGHT='#D9DEE6'
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':11,'axes.spines.top':False,'axes.spines.right':False,
  'axes.edgecolor':'#555','axes.titleweight':'bold','axes.titlesize':14,'axes.titlelocation':'left','figure.dpi':200})
SRC='Source: Nebius Group SEC filings (6-K results PRs, shareholder letters, interim FS, 20-F); see nebius_filings.xlsx for cell-level citations.'
def E(t): return t.replace('$',r'\$')
def foot(fig,extra=''):
    fig.text(0.01,0.01,(SRC+(' '+extra if extra else '')).replace('$',r'\$'),fontsize=7,color='#666',ha='left',wrap=True)
Q=['Q1\'24','Q2\'24','Q3\'24','Q4\'24','Q1\'25','Q2\'25','Q3\'25','Q4\'25','Q1\'26','Q2\'26']
rev=[9.7,14.5,32.1,35.2,50.9,105.1,146.1,227.7,399.0,582.3]
# segments Q3'24 onward (H1'24 quarterly split not disclosed)
seg_q=Q[2:]
ai=[27.1,23.8,41.4,93.7,131.1,214.1,389.7,574.9]
tt=[6.3,12.3,10.5,12.3,16.2,15.1,11.6,10.0]
av=[0.1,0.1,0.2,0.2,0.2,0.7,0.9,1.0]
el=[-1.4,-1.0,-1.2,-1.1,-1.4,-2.2,-3.2,-3.6]
arr=[32,45,121,90,249,430,551,1250,1920,3000]
ebitda=[-58.4,-58.1,-45.9,-63.9,-53.7,-21.0,-5.2,15.0,129.5,236.2]
capex=[58.9,159.0,172.1,417.5,543.9,510.6,955.5,2056.0,2472.9,5657.4]
da=[8.9,11.4,23.5,33.3,49.1,75.2,99.0,180.7,212.0,259.7]
rpo_q=['Q4\'25','Q1\'26','Q2\'26']; rpo=[21333.0,33585.3,37490.6]; defrev=[1577.5,4778.1,5975.2]
def lab(ax,bars,fmt='{:,.0f}',dy=0.01,neg=False):
    top=ax.get_ylim()[1]
    for b in bars:
        h=b.get_height(); y=b.get_y()+h if h>=0 else b.get_y()
        ax.text(b.get_x()+b.get_width()/2, (y+top*dy) if h>=0 else (y-top*0.05), fmt.format(h if h>=0 else h), ha='center',fontsize=8,color='#333')
# 1 Revenue by segment
fig,ax=plt.subplots(figsize=(11,6))
import numpy as np
x=np.arange(len(seg_q))
b1=ax.bar(x,ai,color=NAVY,label='Nebius AI cloud'); b2=ax.bar(x,tt,bottom=ai,color=TEAL,label='TripleTen')
b3=ax.bar(x,av,bottom=[a+t for a,t in zip(ai,tt)],color=GOLD,label='Avride')
ax.bar(x,el,color=LIGHT,label='Eliminations')
tot=rev[2:]
for i,t in enumerate(tot): ax.text(i,ai[i]+tt[i]+av[i]+8,f'${t:,.0f}M',ha='center',fontsize=9,fontweight='bold')
ax.set_xticks(x,seg_q); ax.set_ylabel('$M'); ax.legend(frameon=False,loc='upper left')
ax.set_title(E('Revenue by segment ($M, continuing ops) - AI cloud is ~99% of growth'))
ax.yaxis.set_major_formatter(FuncFormatter(lambda v,p:f'{v:,.0f}'))
foot(fig,'Q4\'24 and Q4\'25 segment values derived (FY less 9M). H1\'24 quarterly segment split not disclosed. Toloka in discontinued ops.')
fig.tight_layout(rect=(0,0.04,1,1)); fig.savefig(OUT+'kpi_revenue_by_segment.png'); plt.close(fig)
# 2 ARR
fig,ax=plt.subplots(figsize=(11,6))
bars=ax.bar(Q,arr,color=[NAVY]*10); ax.set_ylim(0,max(arr)*1.15); lab(ax,bars,'${:,.0f}M')
ax.axhspan(7000,9000,alpha=0) 
ax.set_title(E('Nebius AI cloud ARR ($M, as stated by company)')); ax.set_ylabel('$M')
ax.text(9.4,max(arr)*1.08,'FY26 ARR guide: $7-9B',ha='right',fontsize=10,color=ORANGE,fontweight='bold')
foot(fig,'2024 quarters per Oct-24 investor presentation chart; Dec-24 $90M (missed $170-190M guide). ARR = month revenue x12 (company definition).')
fig.tight_layout(rect=(0,0.04,1,1)); fig.savefig(OUT+'kpi_arr.png'); plt.close(fig)
# 3 Adj EBITDA
fig,ax=plt.subplots(figsize=(11,6))
bars=ax.bar(Q,ebitda,color=[TEAL if v>=0 else ORANGE for v in ebitda]); ax.axhline(0,color='#333',lw=0.8)
ax.set_ylim(-110,max(ebitda)*1.22)
m=[e/r*100 for e,r in zip(ebitda,rev)]
for i,(b,v) in enumerate(zip(bars,ebitda)):
    t=f'{v:,.1f}'+(f'\n({m[i]:.0f}%)' if i>=4 else '')
    ax.text(b.get_x()+b.get_width()/2, v+6 if v>=0 else v-8, t, ha='center', va='bottom' if v>=0 else 'top', fontsize=8.5)
ax.set_title(E('Group adjusted EBITDA ($M; margin % in brackets) - FY26 guide ~40%')); ax.set_ylabel('$M')
foot(fig,'Non-GAAP as defined by company. Margins shown from Q1\'25 (2024 margins below -100%).')
fig.tight_layout(rect=(0,0.04,1,1)); fig.savefig(OUT+'kpi_adj_ebitda.png'); plt.close(fig)
# 4 Capex vs D&A
fig,ax=plt.subplots(figsize=(11,6)); x=np.arange(10); w=0.4
b1=ax.bar(x-w/2,capex,w,color=NAVY,label='Capex (PP&E + intangibles purchases)'); b2=ax.bar(x+w/2,da,w,color=GREY,label='D&A')
ax.set_ylim(0,max(capex)*1.15)
for b,v in zip(b1,capex): ax.text(b.get_x()+b.get_width()/2,v+60,f'{v:,.0f}',ha='center',fontsize=8)
for b,v in zip(b2,da): ax.text(b.get_x()+b.get_width()/2,v+60,f'{v:,.0f}',ha='center',fontsize=7,color='#555')
ax.set_xticks(x,Q); ax.legend(frameon=False,loc='upper left'); ax.set_ylabel('$M'); ax.yaxis.set_major_formatter(FuncFormatter(lambda v,p:f'{v:,.0f}'))
ax.set_title(E('Capex vs D&A ($M) - capex 22x D&A in Q2\'26; FY26 capex guide $20-25B (call)'))
foot(fig,'Server life extended 4->5 yrs from 1-Jan-2026 (cut Q2\'26 depreciation by $43.0M).')
fig.tight_layout(rect=(0,0.04,1,1)); fig.savefig(OUT+'kpi_capex_vs_da.png'); plt.close(fig)
# 5 RPO / deferred revenue
fig,ax=plt.subplots(figsize=(9,6)); x=np.arange(3)
b1=ax.bar(x-0.2,rpo,0.4,color=NAVY,label='RPO'); b2=ax.bar(x+0.2,defrev,0.4,color=TEAL,label='Deferred revenue (incl. in RPO)')
ax.set_ylim(0,max(rpo)*1.15)
for b,v in zip(list(b1)+list(b2),rpo+defrev): ax.text(b.get_x()+b.get_width()/2,v+400,f'${v/1000:,.1f}B',ha='center',fontsize=9)
ax.set_xticks(x,rpo_q); ax.yaxis.set_major_formatter(FuncFormatter(lambda v,p:f'{v:,.0f}')); ax.legend(frameon=False,loc='upper left'); ax.set_ylabel('$M')
ax.set_title(E('Remaining performance obligations & deferred revenue ($M)'))
foot(fig,'RPO first disclosed at YE25. Deferred revenue = current + non-current (customer prepayments).')
fig.tight_layout(rect=(0,0.04,1,1)); fig.savefig(OUT+'kpi_rpo_deferred_revenue.png'); plt.close(fig)
# 6 Dilution waterfall (replicates CapStructure formulas)
P=float(sys.argv[1]) if len(sys.argv)>1 else 242.81
iss=271855218; exch=int((4403.18*4e5+4413.18*4e5+4e8*0.02*79/360+4e8*0.03*79/360)/223.9); nv=21065936
rsu=6234091; opt=6740600*(1-88.65/P) if P>88.65 else 0
cv=[('Jun-29/31\n$51.45',2*87.5*19.4363*1000,51.45),('Sep-30/32\n$138.75',2*1581.25*7.2072*1000,138.75),('Mar-31\n$183.22',2587.5*5.4579*1000,183.22),('Mar-33\n$180.31',1750*5.546*1000,180.31),('Feb-30\n$313.46',3450*3.1902*1000,313.46),('Feb-34\n$324.65',2300*3.0802*1000,324.65)]
atm=12270507
steps=[('Issued\n6/30/26',iss,'base'),('Aug-26\nexchange',exch,'add'),('NVIDIA\nwarrant',nv,'add'),('Basic',None,'sub'),('RSUs',rsu,'add'),('Options\n(TSM)',opt,'add')]
for n,s,k in cv: steps.append((n,s if P>k else 0,'add' if P>k else 'otm'))
steps+=[('ATM\nremaining',atm,'add'),('Fully\ndiluted',None,'sub')]
fig,ax=plt.subplots(figsize=(13,6.5)); cum=0; xs=[]
for i,(n,v,k) in enumerate(steps):
    if k=='sub':
        ax.bar(i,cum/1e6,color=NAVY); ax.text(i,cum/1e6+3,f'{cum/1e6:,.1f}M',ha='center',fontweight='bold',fontsize=10)
    elif k=='base':
        ax.bar(i,v/1e6,color=GREY); cum=v; ax.text(i,v/1e6+3,f'{v/1e6:,.1f}',ha='center',fontsize=9)
    else:
        col=TEAL if k=='add' else LIGHT
        ax.bar(i,v/1e6,bottom=cum/1e6,color=col,edgecolor='#999' if k=='otm' else None,hatch='//' if k=='otm' else None)
        ax.text(i,(cum+v)/1e6+3,'OTM\n(0)' if k=='otm' else f'+{v/1e6:,.1f}',ha='center',fontsize=8,color='#333')
        cum+=v
    xs.append(n)
ax.set_xticks(range(len(xs)),xs,fontsize=8.5); ax.set_ylim(250,cum/1e6*1.08); ax.set_ylabel('Shares (M)')
basic=iss+exch+nv
ax.set_title(E(f'NBIS dilution waterfall at ${P:,.2f}: {cum/1e6:,.1f}M fully diluted (+{(cum/basic-1)*100:.1f}% vs basic {basic/1e6:,.1f}M)'))
foot(fig,f'Converts if-converted at initial rates on remaining original principal; OTM notes excluded (all-in incl. OTM: +18.1M). No capped calls. NVIDIA pre-funded warrant counted once (in basic EPS, not in issued count). Price input ${P:,.2f} (Oct 2, 2026 close).')
fig.tight_layout(rect=(0,0.05,1,1)); fig.savefig(OUT+'dilution_waterfall.png'); plt.close(fig)
print('FD',cum,'basic',basic)
