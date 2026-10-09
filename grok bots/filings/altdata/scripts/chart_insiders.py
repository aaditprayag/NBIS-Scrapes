import os,sys; sys.path.insert(0,os.path.dirname(__file__))
import matplotlib; matplotlib.use('Agg'); matplotlib.rcParams['text.parse_math']=False
import matplotlib.pyplot as plt, numpy as np
import insiders_sheet as I
A,P,N,NP=I.aggregates(); QS=I.QS[4:]  # Q1-25..Q3-26
x=np.arange(len(QS)); w=0.38
f4p=[P[q][0]/1e6 if q in P else 0 for q in QS]; f4r=[(P[q][1]+P[q][2])/1e6 if q in P else 0 for q in QS]
n_p=[NP[q][0]/1e6 if q in NP else 0 for q in QS]; n_o=[NP[q][1]/1e6 if q in NP else 0 for q in QS]
CAP=140
fig,ax=plt.subplots(figsize=(12,6.2))
b1=ax.bar(x-w/2,np.minimum(n_p,CAP),w,color='#9DC3E6',label='Form 144 notice: 10b5-1 plan')
ax.bar(x-w/2,n_o,w,bottom=np.minimum(n_p,CAP),color='#DEEBF7',edgecolor='#9DC3E6',label='Form 144 notice: sell-to-cover/vesting')
ax.bar(x+w/2,f4p,w,color='#C00000',label='Form 4 executed: 10b5-1 plan')
ax.bar(x+w/2,f4r,w,bottom=f4p,color='#F4B183',label='Form 4 executed: RSU tax sell-to-cover')
for i,q in enumerate(QS):
    if n_p[i]>CAP: ax.text(x[i]-w/2,CAP+2,f"${n_p[i]:.0f}M\n(Volozh 10b5-1 plan\nnotice, up to 4.13M sh,\nSep-25 to Feb-26)",ha='center',va='bottom',fontsize=8)
    t=f4p[i]+f4r[i]
    if t>0: ax.text(x[i]+w/2,t+2,f"${t:.0f}M",ha='center',fontsize=9,color='#C00000')
ax.annotate('Korolenko 500k sh @ $203\non Q1-26 results day (May-13)\n(10b5-1 adopted Dec-2-25)',xy=(x[5]+w/2,105),xytext=(x[5]-1.6,118),fontsize=8,arrowprops=dict(arrowstyle='->'))
ax.axvline(x[4]+0.5-0.5+0.15,color='grey',ls=':',lw=1)
ax.text(x[4]+0.18,175,'Mar-18-26: Section 16 starts for\nFPI insiders (HFIA Act) -\nno Form 4 data before this',fontsize=8,color='grey',va='top')
ev={'Q3-25':'Sep-25: $1.15B equity + $3.16B CB;\nMSFT deal','Q1-26':'NVDA $2B; $4.3B CB','Q2-26':'ATM $2.85B','Q3-26':'$5.75B CB'}
for i,q in enumerate(QS):
    if q in ev: ax.text(x[i],-24,ev[q],ha='center',fontsize=7.5,color='#1F3864')
ax.set_xticks(x); ax.set_xticklabels(QS+[''] * 0); ax.set_ylim(-30,200)
ax.set_ylabel('$M'); ax.set_title('Nebius insider selling: Form 144 notices vs Form 4 executed sales, by quarter (Q3-26 = to Sep-30)',fontsize=12)
ax.legend(loc='upper left',fontsize=8.5); ax.axhline(0,color='black',lw=0.6)
fig.text(0.01,0.01,'Source: SEC EDGAR, NBIS (CIK 1513845) Forms 4, 4/A, 144 raw XML (detail + URLs in altdata.xlsx "Insiders"). Form 144 = notice at market value on notice date, not proof of sale. Q3-26 includes filings through Sep-29-2026.',fontsize=7.5)
plt.tight_layout(rect=(0,0.03,1,1)); plt.savefig('/workspace/nebius/filings/altdata/insiders.png',dpi=160)
