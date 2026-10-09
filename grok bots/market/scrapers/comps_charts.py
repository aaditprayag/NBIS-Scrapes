#!/usr/bin/env python3
"""Deck charts from /workspace/nebius/parts/comps_values.json (built by parts/build_comps.py).
Outputs PNGs to /workspace/nebius/market/charts/. Only companies with sourced inputs are plotted."""
import json, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
V={v["company"]:v for v in json.load(open("/workspace/nebius/parts/comps_values.json"))}
OUT="/workspace/nebius/market/charts/"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":10,"axes.spines.top":False,"axes.spines.right":False})
NAVY,TEAL,GREY,ORANGE="#1F3864","#2E86AB","#8C8C8C","#E07A1F"
def num(x): return isinstance(x,(int,float))

# 1) EV / contracted MW vs revenue growth
fig,ax=plt.subplots(figsize=(9,5.5),dpi=200)
pts=[]; excl=[]
for name,v in V.items():
    if "OCI" in name: continue
    if num(v.get("ev_cmw")) and num(v.get("yoy")): pts.append((name,v["yoy"],v["ev_cmw"]))
    else: excl.append(name)
for name,x,y in pts:
    priv = name in ("Nscale",)
    c = NAVY if name.startswith("Nebius") else (GREY if priv else TEAL)
    ax.scatter(x,y,s=160,color=c,edgecolor="white",zorder=3,marker="D" if priv else "o")
    lab={"Nebius Group":"Nebius","Nscale":"Nscale*"}.get(name,name)
    left = name=="CoreWeave"
    ax.annotate(f"{lab}\n${y:.1f}M/MW, {x:+.0f}%",(x,y),xytext=(-10,8) if left else (8,6),textcoords="offset points",fontsize=9,color=c,weight="bold",ha="right" if left else "left")
ax.set_xscale("symlog",linthresh=100); ax.set_xlabel("Revenue growth, latest quarter YoY (%, symlog scale)")
ax.set_ylabel("Pro forma EV / contracted MW ($M per MW)")
ax.set_title("EV per contracted MW vs. revenue growth",loc="left",weight="bold",fontsize=13,color=NAVY)
ax.grid(alpha=.25); ax.set_ylim(0,max(p[2] for p in pts)*1.25); ax.set_xlim(-60,3000)
foot=("Source: company filings (NBIS 6-K Q2'26, CRWV 10-Q Q2'26, IREN 10-K FY26, Nscale S-1 9/18/26); prices Yahoo Finance close 10/2/26. "
 "EV pro forma for post-balance-sheet converts/exchanges/warrants.\nContracted MW definitions differ (NBIS/CRWV contracted power; IREN ~5 GW grid agreements; Nscale active+contracted). "
 "*Nscale: private, EV = Mar-26 Series C valuation + net debt, growth = H1/H1.\nExcluded: "+", ".join(excl)+
 " (Oracle: MW not disclosed; Crusoe: no reported revenue growth; Lambda: no MW or growth disclosed).")
fig.text(0.01,0.005,foot,fontsize=6.5,color="#555",va="bottom",wrap=True)
plt.tight_layout(rect=(0,0.11,1,1)); fig.savefig(OUT+"ev_per_contracted_mw_vs_growth.png"); plt.close(fig)

# 2) EV/EBITDA (LTM + NTM) and cost of debt
names=[n for n,v in V.items() if "OCI" not in n and (num(v.get("ev_ebitda_ltm")) or num(v.get("cod")))]
excl2=[n for n in V if n not in names and "OCI" not in n]
lab=lambda n: n.replace(" Group","").replace(" (consolidated)","")
fig,(a1,a2)=plt.subplots(1,2,figsize=(11,5),dpi=200,gridspec_kw={"width_ratios":[1.5,1]})
import numpy as np
x=np.arange(len(names)); w=.38
ltm=[V[n]["ev_ebitda_ltm"] if num(V[n].get("ev_ebitda_ltm")) else np.nan for n in names]
ntm=[V[n]["ev_ebitda_ntm"] if num(V[n].get("ev_ebitda_ntm")) else np.nan for n in names]
b1=a1.bar(x-w/2,ltm,w,color=NAVY,label="EV / LTM adj. EBITDA")
b2=a1.bar(x+w/2,ntm,w,color=TEAL,label="EV / NTM adj. EBITDA (consensus)")
a1.set_yscale("log"); a1.set_xticks(x,[lab(n) for n in names]); a1.set_ylabel("x (log scale)")
for bars in (b1,b2):
    for r in bars:
        h=r.get_height()
        if np.isfinite(h): a1.text(r.get_x()+r.get_width()/2,h*1.07,f"{h:.1f}x",ha="center",fontsize=8)
for i,n in enumerate(names):
    if not np.isfinite(ntm[i]): a1.text(x[i]+w/2,1.2,"n/a",ha="center",fontsize=7,color=GREY)
a1.legend(frameon=False,fontsize=8,loc="upper right"); a1.set_title("EV / adj. EBITDA",loc="left",weight="bold",color=NAVY)
a1.set_ylim(1,600)
cod=[V[n]["cod"] if num(V[n].get("cod")) else np.nan for n in names]
b3=a2.bar(x,cod,.55,color=[NAVY if n.startswith("Nebius") else ORANGE for n in names])
for r in b3:
    h=r.get_height()
    if np.isfinite(h): a2.text(r.get_x()+r.get_width()/2,h+.2,f"{h:.1f}%",ha="center",fontsize=8)
a2.set_xticks(x,[lab(n) for n in names]); a2.set_ylabel("%"); a2.set_title("Wtd avg effective cost of debt",loc="left",weight="bold",color=NAVY)
fig.suptitle("Valuation vs. funding cost",x=0.01,ha="left",weight="bold",fontsize=13,color=NAVY)
foot=("Source: company filings; EV pro forma (prices Yahoo close 10/2/26). LTM adj. EBITDA = company-reported (Oracle derived: op. income + D&A + SBC). NTM consensus: NBIS Q4'26-Q3'27 quarterly cons, CRWV calendarized FY26/FY27 cons "
 "(MarketScreener/Yahoo, via team Skeptic sheet 10/4/26); n/a = not sourced.\nCost of debt: principal-weighted effective rates from debt notes (NBIS, CRWV, IREN; converts carry low cash coupons, effective rate incl. discount accretion); "
 "Oracle = annualized interest expense / avg debt (net of capitalized interest). IREN EBITDA is mostly Bitcoin mining.\nExcluded: "+", ".join(excl2)+" (private; EBITDA/LTM and weighted debt cost not disclosed or negative).")
fig.text(0.01,0.005,foot,fontsize=6.3,color="#555",va="bottom",wrap=True)
plt.tight_layout(rect=(0,0.13,1,0.95)); fig.savefig(OUT+"ev_ebitda_and_cost_of_debt.png"); plt.close(fig)
print("ok", excl, excl2)
