# 2026-10-08 ~10:00 ET: re-run on the model v0.9e comparison.xlsx (no lock values in this figure; KPI values unchanged).
# Memo-sized figure strip (v0.9d refresh Mon Oct 5 2026 ET: panel A adds K02b billed-IT); every value read from peers/comparison.xlsx (no new numbers).
import openpyxl, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
wb = openpyxl.load_workbook("/workspace/nebius/peers/comparison.xlsx", data_only=True)
cmp_ = {r[0].value: [c.value for c in r] for r in wb["Comparison"].iter_rows(min_row=4) if r[0].value}
hdr = [c.value for c in wb["Comparison"][3]]
peers = ["Nebius","CoreWeave","IREN","Oracle OCI","Nscale","Lambda","Crusoe"]
col = {p: hdr.index(p) for p in peers}
def v(k,p):
    x = cmp_[k][col[p]]
    return x if isinstance(x,(int,float)) else None
con = {r[0].value: [c.value for c in r] for r in wb["Contracts_$MWyr"].iter_rows(min_row=3) if r[0].value}
plt.rcParams.update({"font.family":"Liberation Sans","font.size":7})
NEB="#1f4e8c"; GRY="#9aa3ad"
fig, ax = plt.subplots(1,3, figsize=(7.4,2.25), gridspec_kw={"width_ratios":[1.05,1,1.15]})
# A: revenue per active MW (K01 period-end, K02 average energized, K02b average billed IT; same defs as Comparison)
lab=["Nebius","CRWV","IREN","Nscale","Lambda","Crusoe"]; ps=["Nebius","CoreWeave","IREN","Nscale","Lambda","Crusoe"]
ser=[("K01",""),("K02","////"),("K02b","xxxx")]
x=range(len(ps)); w=0.27
a=ax[0]
for j,(k,h) in enumerate(ser):
    for i,p in enumerate(ps):
        val=v(k,p)
        if val is None: continue
        c = NEB if p=="Nebius" else GRY
        xx=i+(j-1)*w
        a.bar(xx,val,w,color=c,hatch=h,edgecolor="white",linewidth=0)
        a.text(xx,val+0.25,f"{val:.1f}",ha="center",va="bottom",fontsize=4.8,fontweight="bold" if p=="Nebius" else None,rotation=90)
a.set_xticks(list(x)); a.set_xticklabels(lab,fontsize=5.6,rotation=30,ha="right")
a.set_title("A. Revenue per MW ($M/MW-yr)",fontsize=7.2,fontweight="bold",loc="left")
a.set_ylim(0,17.5)
from matplotlib.patches import Patch
a.legend(handles=[Patch(color=GRY,label="K01 period-end MW"),Patch(facecolor=GRY,hatch="////",edgecolor="white",label="K02 avg energized MW"),Patch(facecolor=GRY,hatch="xxxx",edgecolor="white",label="K02b avg billed IT MW")],fontsize=5.0,frameon=False,loc="upper right")
# B: capex / PP&E per MW
ps2=["Nebius","CoreWeave","Oracle OCI","Nscale"]; lab2=["Nebius","CRWV","Oracle","Nscale"]
ks=[("K13","T4Q capex / MW added"),("K14","Net PP&E / active MW"),("K15","Gross in-service PP&E / MW")]
b=ax[1]; w=0.26; hat=["","////","...."]
for j,(k,l) in enumerate(ks):
    for i,p in enumerate(ps2):
        val=v(k,p)
        if val is None: continue
        c = NEB if p=="Nebius" else GRY
        b.bar(i+(j-1)*w,val,w,color=c,hatch=hat[j],edgecolor="white",linewidth=0)
        b.text(i+(j-1)*w,val+1.2,f"{val:.0f}",ha="center",fontsize=5.6)
med=cmp_["K13"][hdr.index("Peer median (formula; comparable peers)")]
b.axhline(med,color="#c0392b",lw=0.7,ls="--")
b.set_xticks(range(4)); b.set_xticklabels(lab2,fontsize=5.8); b.set_ylim(0,125)
b.set_title("B. Capital per MW ($M/MW)",fontsize=7.2,fontweight="bold",loc="left")
from matplotlib.lines import Line2D
b.legend(handles=[Patch(facecolor=GRY,hatch=h,edgecolor="white",label=l) for h,(k,l) in zip(hat,ks)]+[Line2D([0],[0],color="#c0392b",lw=0.7,ls="--",label=f"peer median capex/MW ({med:.0f})")],fontsize=5.4,frameon=False,loc="upper right",ncol=1)
# C: contract $/MW-yr
# 2026-10-09: Nebius-Microsoft shown only as the house $14.49M per IT MW [E] = model Contracts!E12 (comparison.xlsx N1 col W)
assert abs(con["N1"][22] - 14.494) < 0.001
rows=[("Nebius-MSFT (house, E)",con["N1"][22],True),
      ("IREN-MSFT",con["P1"][21],False),("IREN-NVIDIA",con["NEW1"][21],False),
      ("Nscale-MSFT (IT)",con["P2"][22],False),("Nscale-Anthropic",con["O1"][21],False),
      ("Lambda-Anthropic (P, 6y, Med.)",con["NEW2"][21],False),("Oracle-OpenAI (P, IT E)",con["O2"][22],False),
      ("Facility leases (shell+power)",con["F4"][21],None)]
c=ax[2]
for i,(l,val,neb) in enumerate(rows[::-1]):
    col_ = NEB if neb else ("#d9dde2" if neb is None else GRY)
    c.barh(i,val,0.7,color=col_); c.text(val+0.2,i,(f"{val:.2f}" if neb else f"{val:.1f}"),va="center",fontsize=5.6)
c.set_yticks(range(len(rows))); c.set_yticklabels([r[0] for r in rows[::-1]],fontsize=5.8)
c.set_xlim(0,19.5)
c.set_title("C. Contract value ($M per MW-yr)",fontsize=7.2,fontweight="bold",loc="left")
for a_ in ax:
    a_.spines[["top","right"]].set_visible(False); a_.tick_params(axis="y",labelsize=5.8)
fig.text(0.005,0.014,"Source: peers/comparison.xlsx (Comparison, Contracts_$MWyr). C: Nebius-MSFT = house \\$14.49M per IT MW (E, model Contracts!E12). Nebius MW = Capacity estimate (E); K02b = avg billed IT MW (Nebius 176.1; CRWV K02b = K02).\n"
         "Latest period per peer: Lambda Q2'25 press; Crusoe gross MW, FY26E revenue; IREN K02 Sep-25. C: facility lease = Applied Digital-CoreWeave 1.8 (pack range 1.2-1.9). Comp band 9.7-16.7 (top Lambda-Anthropic, 6-yr, press, Medium).",fontsize=4.9,color="#555")
fig.tight_layout(rect=(0,0.07,1,1),w_pad=1.0)
fig.savefig("/workspace/nebius/peers/memo/fig_unit_economics.png",dpi=300)
print("ok", med)
