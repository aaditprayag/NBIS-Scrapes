import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.dates as md
from matplotlib.patches import Patch
from datetime import date, timedelta
ev=[
(date(2025,11,15),"Nov '25","T1 live on 16 gensets","done","R"),
(date(2026,2,10),"Feb '26","T2 live (32 gensets)","done","R"),
(date(2026,5,20),"May 20 '26","Nebius–Bloom deal: 328 MW 'operational in 2026'","done","R"),
(date(2026,7,25),"Jul–Aug '26","T3–T4 delivered (62 gensets, ~107 IT MW ceiling)","done","R/E"),
(date(2026,8,10),"Aug 10 '26","City stop-construction order on Bloom units","risk","R"),
(date(2026,8,17),"Aug 17 '26","Planning Board approves Bloom/LNG site plan","done","R"),
(date(2026,9,22),"Sep 22 '26","NJDEP genset order: $1.07M, 45-day permit clock","risk","R"),
(date(2026,10,13),"Oct 12–15 '26","DataOne hearing-request window (watch)","pending","E"),
(date(2026,10,14),"Oct 14 '26","Planning Board (agenda posted: no DataOne item)","pending","R"),
(date(2026,10,22),"Oct 22 '26","Independence MO TRO hearing: Phase 1 timing only, under 1.5% of FY27 billed","pending","R/D"),
(date(2026,10,27),"Oct 27 '26","Bloom Q3 call (first chance for delivery news)","pending","E"),
(date(2026,11,8),"~Nov 6–9 '26","NJDEP genset filing deadline","pending","E"),
(date(2026,11,10),"Oct 22 / Nov 10 '26","Nebius Q3 letter: tranche count","pending","E"),
(date(2027,5,15),"Q2 '27","Base case: T5–T6 bill","model","E"),
(date(2027,8,15),"Q3 '27","Base case: T7–T9 bill (240 IT MW)","model","E"),
]
col={"done":"#2e7d32","risk":"#c62828","pending":"#f9a825","model":"#1565c0"}
fig,(a1,ax)=plt.subplots(2,1,figsize=(13,9.5),sharex=True,gridspec_kw={"height_ratios":[1,2.3]})
x0,x1=date(2025,9,1),date(2028,1,1)
# top: IT MW path
t=26.67
steps=[(x0,0),(date(2025,11,15),t),(date(2026,2,10),2*t),(date(2026,8,1),4*t),(date(2027,5,15),6*t),(date(2027,8,15),9*t),(x1,9*t)]
xs=[s[0] for s in steps]; ys=[s[1] for s in steps]
a1.step(xs,ys,where="post",color="#1565c0",lw=2.2,label="Vineland IT MW delivered/billed, base (9 × 26.67 IT MW, equal split) [E]")
a1.axhline(107,color="#c62828",ls=":",lw=1.3); a1.text(date(2025,9,10),110,"~107 IT MW ceiling on 62 gensets + ~5 MW grid (÷1.2 PUE) [E]",fontsize=8,color="#c62828")
a1.axhline(240,color="#555",ls="--",lw=1); a1.text(date(2025,9,10),246,"Microsoft contract: 240 IT MW, 9 tranches [E]; \\$14.49M per IT MW [E, house]",fontsize=8,color="#333")
for d,n in [(date(2025,11,15),16),(date(2026,2,10),32),(date(2026,6,15),62)]:
    a1.annotate(f"{n} gensets [R]",(d,0),xytext=(0,6),textcoords="offset points",fontsize=7.5,color="#2e7d32",ha="center")
a1.axvspan(date(2026,11,6),date(2027,5,15),color="#fff3e0",zorder=0); a1.text(date(2026,11,20),180,"T5–T9 need Bloom:\nno public delivery date",fontsize=8.5,color="#e65100")
a1.set_ylim(0,275); a1.set_ylabel("IT MW"); a1.legend(loc="lower right",fontsize=8,frameon=False)
for s_ in ["right","top"]: a1.spines[s_].set_visible(False)
a1.axvline(date(2026,10,9),color="#888",ls="--",lw=1)
# bottom: events
n=len(ev)
for i,(d,lab,l,s,tg) in enumerate(ev):
    y=n-i
    ax.hlines(y,x0,d,color="#eeeeee",lw=0.8)
    ax.scatter([d],[y],color=col[s],s=60,zorder=3)
    ax.text(d+timedelta(days=8),y,f"{lab}  {l} [{tg}]",va="center",fontsize=8.3)
ax.axvline(date(2026,10,9),color="#888",ls="--",lw=1); ax.text(date(2026,10,9),n+0.8,"Today (Oct 9)",ha="center",fontsize=8,color="#555")
ax.text(date(2025,9,5),-0.4,"UNDATED gates for T5–T9 (no public date): (1) city lifts Aug 10 stop order   (2) NJDEP fuel-cell air permit   (3) Bloom delivery / start-up disclosed",fontsize=9,color="#c62828",fontweight="bold")
ax.set_ylim(-1,n+1.3); ax.set_yticks([]); ax.set_xlim(x0,x1)
ax.xaxis.set_major_locator(md.MonthLocator(interval=2)); ax.xaxis.set_major_formatter(md.DateFormatter("%b '%y"))
for s_ in ["left","right","top"]: ax.spines[s_].set_visible(False)
ax.legend(handles=[Patch(color=col[k],label=v) for k,v in [("done","Happened"),("risk","Risk event"),("pending","Upcoming / watch"),("model","Our billing base case")]],loc="upper right",fontsize=8,frameon=False)
fig.suptitle("Vineland (Microsoft) power & permitting: T1–T4 (~107 IT MW) run on unpermitted gensets; T5–T9 gated on Bloom fuel cells (as of Oct 9, 2026)",fontsize=11)
fig.text(0.01,0.012,"All MW are IT MW. Locks: FY27 avg billed 531.8 base / 659.2 high; Vineland 182.3 MW-yr. Shutdown risk: \\$122–224M Q4'26 revenue, \\$79M probability-weighted [D]. Independence is Phase 1 only, 0 MW at YE26.\nSources: Nebius 20-F & Q2'26 letter; Nebius–Bloom PR 5/20/26; NJDEP AO 9/22/26; Floodlight (DataOne letter); Vineland PB agendas 8/17, 10/14; Jackson County TRO coverage. R=Reported, D=Derived, E=Estimated.",fontsize=7,color="#555")
plt.tight_layout(rect=(0,0.04,1,0.97)); plt.savefig("/workspace/nebius/capacity/vineland_power_permitting_exhibit.png",dpi=170)
plt.savefig("/workspace/nebius/capacity/vineland_power_permitting_exhibit.pdf")
