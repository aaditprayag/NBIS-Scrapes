#!/usr/bin/env python3
# peers/charts/price_or_mw.png from recalculated peers/price_or_mw.xlsx Implied_Price (2026-10-08 ET; replaces the ad-hoc Oct-5 chart).
# 2026-10-09 N1 ruling: note cites Microsoft-Nebius only as $14.49M per IT MW [E, house] (Contracts!E12).
# ~10:00 ET: labels/footnote re-pointed to model v0.9e locks (values come from the workbook).
import openpyxl, textwrap, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
ip = openpyxl.load_workbook("/workspace/nebius/peers/price_or_mw.xlsx", data_only=True)["Implied_Price"]
nc = openpyxl.load_workbook("/workspace/nebius/peers/price_or_mw.xlsx", data_only=True)["Nebius_Check"]
f = lambda r: ip.cell(r, 6).value; d = lambda r: ip.cell(r, 4).value
BARS = [  # row, label, colour
 (18, f"Nebius Street\nEX-MSFT\n${ip.cell(18, 2).value/1000:.2f}B / {d(18):,.1f} MW", "#5c0000"),
 (5,  f"Nebius Street\nblended\n$12.3B / {d(5):,.1f} MW", "#8b0000"),
 (24, f"Nebius Street\n/ FY27 HIGH\n$12.3B / {d(24):,.1f} MW", "#c06060"),
 (6,  f"Nebius YE26\nARR low\n$7B / {d(6):,.1f} MW", "#cd5c5c"),
 (7,  f"Nebius YE26\nARR high\n$9B / {d(7):,.1f} MW", "#e9967a"),
 (8,  "CoreWeave Street\n/ BASE MW", "#1f4e79"), (9, "CoreWeave Street\n/ LOW MW", "#5b9bd5"),
 (12, "IREN YE26 ARR\n/ 300 IT MW", "#548235"), (15, "IREN FY27 ARR-path\nproxy", "#a9d18e"),
 (16, "Oracle OCI run-rate\n/ cum. delivered MW", "#ed7d31")]
for r, _, _ in BARS: assert isinstance(f(r), (int, float)), (r, f(r))
bm = nc["B37"].value
fig, ax = plt.subplots(figsize=(15, 6.5))
xs = range(len(BARS))
ax.axhspan(13, 18, color="#fff2cc", alpha=0.8, label="$13–18 slide band", zorder=0)
ax.bar(xs, [f(r) for r, _, _ in BARS], 0.7, color=[c for _, _, c in BARS], zorder=2)
for x, (r, _, _) in zip(xs, BARS): ax.text(x, f(r) + 0.6, f"{f(r):.1f}", ha="center", fontweight="bold", fontsize=10)
ax.axhline(bm, color="#5c0000", ls="--", lw=1.4, label=f"Nebius Street ex-MSFT ~${bm:.1f}M (peer benchmark)", zorder=3)
ax.set_xticks(list(xs)); ax.set_xticklabels([l for _, l, _ in BARS], fontsize=8)
ax.set_ylabel("$M per billed IT MW-yr"); ax.set_ylim(0, 40); ax.legend(loc="upper right", fontsize=9)
ax.set_title("PRICE-OR-MW: FY27 rev (or run-rate) ÷ avg billed IT MW — FY27 on model v0.9e billed MW", fontweight="bold")
note = ("Sources: peers/price_or_mw.xlsx (K02b billed-IT basis). FY27 avg billed IT MW = model v0.9e locks BASE {b:.1f} / HIGH {h:.1f} "
        "(model/nebius_model.xlsx Drivers B34/C34; engine-linked, Vineland at billed Microsoft tranches). "
        "Ex-MSFT = (\u002412.3B \u2212 \u0024{ms:,.1f}M Microsoft FY27, Drivers B32) / ({b:.1f} \u2212 {vin:.1f} Vineland billed MW-yr) = model Drivers B9 \u2014 the figure to compare with peers (peer \u0024/MW has no Microsoft anchor; Microsoft\u2013Nebius = \u002414.49M per IT MW [E, house], model Contracts!E12). "
        "Shaded $13\u201318 = this test's slide hypothesis band (revenue per billed MW); separate contract comp band (TCV/term/MW) = \u00249.7\u201316.7M, top Lambda\u2013Anthropic ~\u002416.7M (6-yr, press, Medium). "
        "YE26 high = \u00249B / 284.2. CoreWeave YE27 MW Estimated; IREN AI GAAP consensus n-d (ARR used); Oracle MW = cumulative delivered (likely gross), revenue includes non-AI OCI.").format(
        b=d(5), h=d(24), ms=nc["B34"].value, vin=nc["B36"].value)
fig.tight_layout(); fig.text(0.01, -0.01, textwrap.fill(note.replace("$", r"\$"), 200), fontsize=7.5, color="dimgray", va="top")
out = "/workspace/nebius/peers/charts/price_or_mw.png"; fig.savefig(out, dpi=150, bbox_inches="tight"); print(out)
