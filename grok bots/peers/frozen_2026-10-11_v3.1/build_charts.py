# PNG charts for the peer comparison. Reads cached (recalculated) values from peers/comparison.xlsx Chart_Data. Author: Grok Bot, 2026-10-04 ET.
# v0.9d refresh 2026-10-05 ET: rev chart adds K02b (billed-IT); new nebius_billing_lock_v09d.png (billing IT MW lock + implied $/billed MW).
# 2026-10-08 ~08:46 ET (b): billing chart -> nebius_billing_lock.png; FY27 tranche locks 536.5/663.8 + ex-Microsoft peer bar; share convention 358.56M debt / 376.65M if-converted.
# 2026-10-08 ET: billing-lock chart footnote -> share basis 376.6M PF primary (358.3M v0.9d memo) + Q3'26 results '~Nov 10 (est.)'. Chart values unchanged.
# 2026-10-08 ~10:00 ET (v0.9e): billing-lock chart + footnotes re-pointed to model v0.9e locks (FY27 531.8/659.2 engine-linked; ex-MSFT = (Street - $2,643.8M) / (base - 182.3)); values read from comparison.xlsx.
import openpyxl, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np, textwrap
OUT = "/workspace/nebius/peers/charts"
wb = openpyxl.load_workbook("/workspace/nebius/peers/comparison.xlsx", data_only=True); D = wb["Chart_Data"]; C = wb["Comparison"]
COLS = ["Nebius", "CoreWeave", "IREN", "Oracle OCI", "Nscale", "Lambda", "Crusoe"]
CLR = {"Nebius": "#1f4e9c", "CoreWeave": "#7f7f7f", "IREN": "#2ca02c", "Oracle OCI": "#d62728", "Nscale": "#9467bd", "Lambda": "#8c564b", "Crusoe": "#ff7f0e"}
# tag/as-of lookup from Comparison
crow = {C.cell(r, 1).value: r for r in range(4, 60) if C.cell(r, 1).value}
def tag(kid, name):
    r = crow[kid]; c = 4 if name == "Nebius" else 7 + 3 * COLS[1:].index(name)
    return "\n".join(str(x) for x in (C.cell(r, c + 1).value, C.cell(r, c + 2).value) if x)
KNAME = {C.cell(r, 1).value: C.cell(r, 2).value for r in range(4, 60) if C.cell(r, 1).value}
def leglab(lab):
    kid = lab.split(" ")[0]
    return textwrap.fill(f"{kid} {KNAME.get(kid, lab)}", 34)
blocks = {}
r = 1
while r <= D.max_row:
    a = D.cell(r, 1).value
    if a and D.cell(r, 2).value in ("Nebius", "$M/MW-yr", "Active low"):
        rows = []; rr = r + 1
        while rr <= D.max_row and D.cell(rr, 1).value: rows.append(rr); rr += 1
        blocks[a] = (r, rows); r = rr
    r += 1
def val(x): return x if isinstance(x, (int, float)) else np.nan
def grouped(title, fname, ylabel, note, log=False, kid_for_tags=None):
    r0, rows = blocks[title]
    labels = [D.cell(rr, 1).value for rr in rows]
    vals = [[val(D.cell(rr, j).value) for j in range(2, 9)] for rr in rows]
    fig, ax = plt.subplots(figsize=(12, 5.5)); x = np.arange(len(COLS)); w = 0.8 / len(rows)
    for i, (lab, vs) in enumerate(zip(labels, vals)):
        b = ax.bar(x + (i - (len(rows) - 1) / 2) * w, vs, w, label=leglab(lab), alpha=0.9 if i == 0 else 0.55, color=[CLR[c] for c in COLS], edgecolor="black", hatch=None if i == 0 else ("//" if i == 1 else "xx"))
        for xi, v in zip(x + (i - (len(rows) - 1) / 2) * w, vs):
            if not np.isnan(v): ax.text(xi, v, f"{v:,.1f}", ha="center", va="bottom", fontsize=8)
    kid = labels[0].split(" ")[0]
    ax.set_xticks(x); ax.set_xticklabels([f"{c}\n{tag(kid, c)}" if kid in crow else c for c in COLS], fontsize=7.5); ax.set_xlim(-0.6, len(COLS) - 0.4)
    for i, c in enumerate(COLS):
        if all(np.isnan(v[i]) for v in vals): ax.text(i, 0, "n/d", ha="center", va="bottom", color="grey", fontsize=8)
    if log: ax.set_yscale("log")
    ax.set_ylabel(ylabel); ax.set_title(title)
    if not log: ax.set_ylim(top=np.nanmax(np.array(vals, dtype=float)) * 1.15)
    ax.legend(fontsize=8, loc="upper left", bbox_to_anchor=(1.01, 1), borderaxespad=0)
    fig.tight_layout()
    fig.text(0.01, -0.02, textwrap.fill(note, 170), fontsize=7, color="dimgray", va="top")
    p = f"{OUT}/{fname}"; fig.savefig(p, dpi=150, bbox_inches="tight"); plt.close(fig); print(p)
src = "Source: peers/comparison.xlsx Comparison (tags R/P/D/E under each name; latest period per peer). "
_kr = {D.cell(r, 1).value.split(" ")[0]: r for r in range(1, D.max_row + 1) if isinstance(D.cell(r, 1).value, str) and D.cell(r, 1).value.startswith("K0")}
_k02, _k02b = D.cell(_kr["K02"], 2).value, D.cell(_kr["K02b"], 2).value
grouped("Revenue per active / billing MW ($M/MW-yr)", "rev_per_active_mw.png", "$M per MW-yr", src + f"K01/K02 on Capacity ENERGIZED MW (Nebius K02 {_k02:.2f} on avg 187.25 MW = (181.5+193)/2, Capacity base case, E). K02b = BILLED-IT basis: Nebius {_k02b:.2f} on avg Capacity billing IT MW 176.1 (Q1/Q2'26 165.5/186.8); CoreWeave K02b = K02 (active ~ billed); other peers n/d billed MW. MW bases differ (Crusoe GROSS, FY26E revenue; Lambda Q2'25 press). Oracle has no installed-MW disclosure. billed-MW locks (model v0.9e): YE26 258/284, FY27 avg 531.8/659.2 engine-linked (Nebius_Inputs).")
grouped("Realized $/GPU-hr vs list", "rev_per_gpu_hr.png", "$ per GPU-hr", src + "K07 = revenue / (avg GPUs x hours); Nebius GPUs modeled by Capacity, revenue includes non-GPU-hour items. K08 = fleet-weighted list (IREN = own $2.00 assumption).")
grouped("Capex / PP&E per MW ($M/MW)", "capex_per_mw.png", "$M per MW", src + "K13 trailing-4Q capex / MW added (Nebius capex funds 2027 capacity). K14 net PP&E incl. CIP. Crusoe 12.5 shell-only excluded.")
grouped("Cost of debt (%)", "cost_of_debt.png", "%", src + "K28 principal-weighted effective (Nscale 18.7% legacy $137M excluded); K36 interest x4 / avg debt. Marginal secured debt (PX09): Nscale ~6.25% E, Lambda 6.78% R, Crusoe ~9.1% E.")
grouped("Customer concentration (%)", "top_customer_pct.png", "% of revenue", src + "K24 largest customer; K26 combined >=10% customers (Lambda: Amazon + Microsoft 50%, press). Oracle: no >=10% customer.")
grouped("RPO / annualized revenue (years)", "rpo_to_annualized_revenue.png", "years (log)", src + "Nscale ~201 yrs reflects pre-ramp revenue vs $56.4B RPO. Oracle RPO company-wide.", log=True)
grouped("Depreciation yield (%) and useful life (yrs)", "depreciation_yield_useful_life.png", "% / years", src + "K17 = D&A x4 / avg in-service gross PP&E; K16 = GPU/server useful life policy (shorter = more conservative).")
# MW
r0, rows = blocks["Active vs contracted MW (latest; mixed IT/gross)"]
act = [val(D.cell(rows[0], j).value) for j in range(2, 9)]; con = [val(D.cell(rows[1], j).value) for j in range(2, 9)]
fig, ax = plt.subplots(figsize=(11, 5.5)); x = np.arange(7)
ax.bar(x - 0.2, act, 0.4, label="Active / energized MW", color=[CLR[c] for c in COLS], edgecolor="black")
ax.bar(x + 0.2, con, 0.4, label="Contracted MW", color=[CLR[c] for c in COLS], alpha=0.4, edgecolor="black", hatch="//")
_oi = COLS.index("Oracle OCI")  # v3: Oracle contracted is a floor ('>10 GW secured', FQ1 FY27 call) - label it as such
for k_, (xi, v) in enumerate(list(zip(x - 0.2, act)) + list(zip(x + 0.2, con))):
    if not np.isnan(v): ax.text(xi, v, ">10 GW" if k_ == len(COLS) + _oi else f"{v:,.0f}", ha="center", va="bottom", fontsize=8)
ax.set_yscale("log"); ax.set_xticks(x); ax.set_xticklabels(COLS); ax.set_ylabel("MW (log)"); ax.set_title("Active vs contracted MW"); ax.set_ylim(top=np.nanmax(con) * 2.5)
ax.legend(fontsize=8, loc="upper left", bbox_to_anchor=(1.01, 1), borderaxespad=0)
fig.tight_layout(); fig.text(0.01, -0.02, textwrap.fill(D.cell(rows[1] + 1, 1).value, 170), fontsize=7, color="dimgray", va="top")
fig.savefig(f"{OUT}/active_vs_contracted_mw.png", dpi=150, bbox_inches="tight"); plt.close(fig); print(f"{OUT}/active_vs_contracted_mw.png")
# contracts
r0, rows = blocks["Contracts: $M per MW-yr (IT MW where available, else stated MW)"]
KS = wb["Contracts_$MWyr"]; KL = {KS.cell(r, 1).value: f"{KS.cell(r, 3).value} - {KS.cell(r, 4).value}" for r in range(4, KS.max_row + 1) if KS.cell(r, 2).value}
labs = [D.cell(rr, 1).value for rr in rows]; vs = [val(D.cell(rr, 2).value) for rr in rows]
labs = [f"{l.split(' ')[0]} " + textwrap.shorten(KL.get(l.split(' ')[0], l), 52, placeholder="...") for l in labs]
keep = [(l, v) for l, v in zip(labs, vs) if not np.isnan(v)]
fig, ax = plt.subplots(figsize=(12, 8))
cols = ["#bbbbbb" if l.split(" ")[0][:1] in "FR" else ("#1f4e9c" if l.startswith("N") and not l.startswith("NEW") else "#ff7f0e" if l.startswith("NEW") else "#7f7f7f") for l, _ in keep]
ax.barh([l for l, _ in keep][::-1], [v for _, v in keep][::-1], color=cols[::-1], edgecolor="black")
for i, (l, v) in enumerate(keep[::-1]): ax.text(v, i, f" {v:,.2f}", va="center", fontsize=8)
ax.set_xlabel("$M per MW-yr (TCV / term / MW; IT MW where available)"); ax.set_title("Implied contract $M per MW-yr\n(blue = Nebius, dark grey = reused GPU Market rows, orange = new rows, light grey = facility-only leases)", fontsize=11)
ax.set_xlim(right=max(v for _, v in keep) * 1.15)
NOTE = ("Source: comparison.xlsx Contracts_$MWyr (reused rows cite parts/contract_economics.xlsx). Nebius-Microsoft \\$14.49M per IT MW [E, house] (model Contracts!E12; 11.60 and 17.08 superseded); Meta #1 on ~50 MW low-confidence estimate; Cerebras naive (ramp). Facility leases are shell+power only. Compute comp band 9.7-16.7: top Lambda-Anthropic ~16.7 ($35B / 6-yr / ~350 MW; press, Medium confidence), low IREN-Microsoft 9.7.")
fig.tight_layout(); fig.text(0.01, -0.01, textwrap.fill(NOTE, 170), fontsize=7, color="dimgray", va="top")
fig.savefig(f"{OUT}/contracts_usd_per_mw_yr.png", dpi=150, bbox_inches="tight"); plt.close(fig); print(f"{OUT}/contracts_usd_per_mw_yr.png")
# sites
r0, rows = blocks["Sites: evidenced-active (base) vs announced/contracted MW"]
labs = [D.cell(rr, 1).value for rr in rows]
lo, ba, hi, an, ye = ([val(D.cell(rr, j).value) for rr in rows] for j in range(2, 7))
fig, ax = plt.subplots(figsize=(11, 6)); y = np.arange(len(labs))
ax.barh(y + 0.2, an, 0.4, color="#dddddd", edgecolor="black", label="Announced / contracted MW")
ax.barh(y - 0.2, ba, 0.4, color=["#1f4e9c" if l.startswith("Nebius") else "#ff7f0e" for l in labs], edgecolor="black", label="Evidenced-active MW (base)")
err = np.array([[max(b - l, 0) if not np.isnan(b) else 0 for b, l in zip(ba, lo)], [max(h - b, 0) if not np.isnan(b) else 0 for b, h in zip(ba, hi)]])
ax.errorbar(ba, y - 0.2, xerr=err, fmt="none", ecolor="black", capsize=3, label="Active low-high")
ax.scatter(ye, y - 0.2, marker="D", color="red", zorder=5, label="YE26 base (site ENERGIZED, not billing)")
for i in range(len(labs)):
    ax.text(max(np.nan_to_num(an[i]), np.nan_to_num(hi[i])) + 10, y[i], f"active {ba[i]:,.0f} / {an[i]:,.0f}", va="center", fontsize=7)
ax.set_yticks(y); ax.set_yticklabels(labs, fontsize=8); ax.invert_yaxis(); ax.set_xlabel("MW (IT where stated; Nebius contracted mostly gross/unknown)")
ax.set_title("Key sites: evidenced-active vs announced MW (as of 2026-10-08; Abilene per Oracle Sep 10)"); ax.set_xlim(right=np.nanmax(an) * 1.18)
ax.legend(fontsize=7, loc="upper left", bbox_to_anchor=(1.01, 1), borderaxespad=0)
NOTE = ("Source: comparison.xlsx Sites (peers/sites/sites.xlsx + site_verification.md; capacity/nebius_capacity.xlsx Quarterly MW Q3'26/Q4'26 + site_verification.md). Same Green/Yellow/Red method; peer rows match the four site appendices (peers/sites/appendix_*.pdf). Crusoe Abilene active 618 MW [D on R]: Oracle FQ1 FY27 call (Sep 10) - 6 of 8 buildings / 75% delivered; low 421 = Epoch (P); IT vs gross basis unstated (was 200/421/421 on Oct 5). Diamonds = per-site YE26 ENERGIZED base; Nebius company BILLED IT MW locks (model v0.9e): YE26 258 base / 284 high, FY27 avg 531.8 / 659.2 (engine-linked) - see nebius_billing_lock.png.")
fig.tight_layout(); fig.text(0.01, -0.01, textwrap.fill(NOTE, 170), fontsize=7, color="dimgray", va="top")
fig.savefig(f"{OUT}/sites_active_vs_announced_mw.png", dpi=150, bbox_inches="tight"); plt.close(fig); print(f"{OUT}/sites_active_vs_announced_mw.png")
# Nebius billed IT MW locks (model v0.9e) + implied $ per billed IT MW
_br = next(r for r in range(1, D.max_row + 1) if D.cell(r, 1).value == "Nebius billed IT MW path - locks (model v0.9e)")
rows = []; rr = _br + 1
while rr <= D.max_row and D.cell(rr, 1).value: rows.append(rr); rr += 1
labs = [D.cell(r, 1).value for r in rows]; mw = [val(D.cell(r, 2).value) for r in rows]; per = [val(D.cell(r, 4).value) for r in rows]
fig, ax = plt.subplots(figsize=(12.5, 5.8)); x = np.arange(len(rows))
cl = ["#7f7f7f", "#1f4e9c", "#9dc3e6", "#1f4e9c", "#ed7d31", "#9dc3e6"][:len(rows)]
ax.bar(x, mw, 0.6, color=cl, edgecolor="black")
for xi, m_, p_ in zip(x, mw, per):
    ax.text(xi, m_, f"{m_:,.0f} IT MW\n${p_:,.1f}M / billed MW", ha="center", va="bottom", fontsize=8.5)
ax.set_xticks(x); ax.set_xticklabels([textwrap.fill(l, 18) for l in labs], fontsize=8); ax.set_ylabel("Billing IT MW"); ax.set_ylim(top=np.nanmax(mw) * 1.25)
ax.set_title("Nebius billed IT MW locks - model v0.9e (base dark / high light / orange = ex-Microsoft peer comparator)\nand implied $M per billed IT MW-yr", fontsize=10)
_ni = {wb["Nebius_Inputs"].cell(r, 1).value: wb["Nebius_Inputs"].cell(r, 3).value for r in range(1, wb["Nebius_Inputs"].max_row + 1)}
NOTE = ("Source: comparison.xlsx Chart_Data <- Nebius_Inputs N_BILL_* (model/nebius_model.xlsx v0.9e: YE26 Judge_Slide B34/B35; FY27 Drivers B34/C34 engine-linked = non-Vineland path + Vineland billed MW-yr; Q2: capacity/nebius_capacity.xlsx Quarterly MW row 147). "
        "Q2'26 = K02b (AI cloud rev x4 / avg billing IT MW). YE26 = run-rate guide \\$7B on base / \\$9B on high (same pairing as peers/price_or_mw.xlsx). FY27 = Street \\$12.3B / avg billed MW (blended, includes the \\$14.5M/MW Microsoft anchor); "
        f"orange = (Street - Microsoft \\${_ni['N_MSFT_FY27_REV']:,.1f}M) / (base - Vineland {_ni['N_MSFT_FY27_MW']:.1f} billed MW-yr) = \\${_ni['D_STREET_FY27_NONMSFT']:.2f}M = model Drivers B9, the figure to compare with peers. "
        f"At \\$16M comps the Street needs 768.75 MW = {_ni['D_COV16_BASE']:.2f}x base / {_ni['D_COV16_HIGH']:.2f}x high. Billed-IT basis throughout. Shares not charted (no per-share/EV-per-MW figure): FD 358.56M with new converts as debt (below \\$313.46 / \\$324.65 conversion prices) / 376.65M if-converted; v0.9d 358.34M historical. Nebius Q3'26 results ~Nov 10 (est.; not announced).")
fig.tight_layout(); fig.text(0.01, -0.01, textwrap.fill(NOTE, 170), fontsize=7, color="dimgray", va="top")
fig.savefig(f"{OUT}/nebius_billing_lock.png", dpi=150, bbox_inches="tight"); plt.close(fig); print(f"{OUT}/nebius_billing_lock.png")
