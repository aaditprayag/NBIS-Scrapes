import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.dates as md
from datetime import date
ev=[ # date, label, status, tag
(date(2025,11,15),"T1 live on 16 gensets","done","R"),
(date(2026,2,10),"T2 live (32 gensets)","done","R"),
(date(2026,5,20),"Nebius–Bloom deal: 328 MW 'operational in 2026'","done","R"),
(date(2026,7,25),"T3–T4 delivered (62 gensets, ~107 IT MW ceiling)","done","R/E"),
(date(2026,8,10),"City stop-construction order on Bloom units","risk","R"),
(date(2026,8,17),"Planning Board approves Bloom/LNG site plan","done","R"),
(date(2026,9,22),"NJDEP genset order: $1.07M, 45-day permit clock","risk","R"),
(date(2026,10,13),"DataOne hearing-request window Oct 12–15 (watch)","pending","E"),
(date(2026,10,14),"Planning Board Oct 14 (agenda posted: no DataOne item)","pending","R"),
(date(2026,10,27),"Bloom Q3 call (first chance for delivery news)","pending","E"),
(date(2026,11,8),"NJDEP genset filing deadline ~Nov 6–9","pending","E"),
(date(2026,11,10),"Nebius Q3 letter: tranche count (Oct 22 or Nov 10)","pending","E"),
(date(2027,5,15),"Base case: T5–T6 bill Q2'27","model","E"),
(date(2027,8,15),"Base case: T7–T9 bill Q3'27","model","E"),
]
APPROX={'T1 live': "Nov '25", 'T2 live': "Feb '26", 'T3–T4 delivered': "Jul–Aug '26", 'hearing-request': "Oct 12–15 '26", 'genset filing deadline': "~Nov 6–9 '26", 'Nebius Q3 letter': "Oct 22 / Nov 10 '26", 'T5–T6': "Q2 '27", 'T7–T9': "Q3 '27"}
def lab(d,l):
    for k,v in APPROX.items():
        if k in l: return v
    return f"{d:%b %d '%y}"
col={"done":"#2e7d32","risk":"#c62828","pending":"#f9a825","model":"#1565c0"}
fig,ax=plt.subplots(figsize=(13,7))
n=len(ev)
for i,(d,l,s,t) in enumerate(ev):
    y=n-i
    ax.hlines(y,date(2025,9,1),d,color="#e0e0e0",lw=0.8)
    ax.scatter([d],[y],color=col[s],s=70,zorder=3)
    ax.text(d+__import__("datetime").timedelta(days=8),y,f"{lab(d,l)}  {l} [{t}]",va="center",fontsize=8.5)
ax.axvline(date(2026,10,9),color="#888",ls="--",lw=1); ax.text(date(2026,10,9),n+0.9,"Today (Oct 9)",ha="center",fontsize=8,color="#555")
ax.text(date(2025,9,5),-0.6,"UNDATED gates for T5–T9 (no public date): (1) city lifts Aug 10 stop order   (2) NJDEP fuel-cell air permit   (3) Bloom delivery / start-up disclosed",fontsize=9,color="#c62828",fontweight="bold")
ax.set_ylim(-1.2,n+1.4); ax.set_yticks([]); ax.set_xlim(date(2025,9,1),date(2028,4,1))
ax.xaxis.set_major_locator(md.MonthLocator(interval=2)); ax.xaxis.set_major_formatter(md.DateFormatter("%b '%y"))
for s_ in ["left","right","top"]: ax.spines[s_].set_visible(False)
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=col[k],label=v) for k,v in [("done","Happened"),("risk","Risk event"),("pending","Upcoming / watch"),("model","Our billing base case")]],loc="upper right",fontsize=8,frameon=False)
ax.set_title("Vineland (Microsoft): T1–T4 run on gensets; T5–T9 gated on Bloom fuel cells — no public Bloom delivery date (as of Oct 9, 2026)",fontsize=10.5)
fig.text(0.01,0.01,"Sources: Nebius 20-F & Q2'26 letter; Nebius–Bloom PR 5/20/26; NJDEP AO 9/22/26; DataOne NJDEP letter (Floodlight); Vineland PB agendas 8/17, 10/14; WHYY/Hunterbrook; MarketBeat (Bloom date). R=Reported, E=Estimated.",fontsize=7,color="#555")
plt.tight_layout(rect=(0,0.03,1,1)); plt.savefig("/workspace/nebius/capacity/vineland_bloom_timeline.png",dpi=170)
