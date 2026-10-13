#!/usr/bin/env python3
# Builds rubin_price_points.xlsx. Every figure is sourced; derived math shown in notes.
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
H=8760; GB_A=5.10; GB_B=4.65; RUB_MW=320; GB_MW=491
def mw(p): return p*RUB_MW*H/1e6
rows=[]
def add(**k): rows.append(k)

# 1 Anchor: Anthropic-Nscale
p1=44.6e9/6/460e0/1e6   # $M per IT MW-yr
g1=p1*1e6/(RUB_MW*H)
add(source="Nscale F-1 Ex-10.25 (Anthropic order, pricing redacted) + F-1 'up to $44.6B'; Bloomberg/CNBC 26-Aug-2026 ($45B, 6 yr, ~460 MW)",
    date="2026-08-26 / 2026-09-18", provider="Nscale (customer: Anthropic)", chip="Vera Rubin NVL72 (VR200)",
    p=g1, term="6 yr; first capacity late 2027", loc="Contract (take-or-pay after acceptance; tranches)",
    mwv=p1, tag="Derived",
    url="http://archive.fast-edgar.com/20260918/A82LC22FZZ2R89ZK22ZC2ZYQJ8SJZZ22Z28Q/ck0002110365-ex10_25.htm ; https://www.bloomberg.com/news/articles/2026-08-26/anthropic-to-pay-nscale-45-billion-for-ai-computing-power ; https://superpowerdaily.com/posts/nscale-s-filing-leaves-103-4-billion-tcv-without-a-disclosed-payment-floor",
    notes=f"EXISTING ANCHOR (re-verified). $44.6B / 6 yr / 460 MW = ${p1:.2f}M per IT MW-yr; / ({RUB_MW} GPU/MW x 8,760) = ${g1:.2f}/GPU-hr. Using press $45B: $16.30M/MW-yr, $5.82/GPU-hr. Per-GPU price is NOT disclosed (redacted in Ex-10.25; formula = hours x GPUs x price/GPU-hr). 460 MW is press-sourced (people familiar). GPU-hr figure is highly sensitive to GPUs/MW: at 400 GPU/MW (NVIDIA DSX figure) = ${p1*1e6/(400*H):.2f} (0.91x of $5.10). Ratable, 100%-billed assumption -> floor.")

# 2 WhiteFiber - Prime Intellect (NEW)
tcv=108.2e6; n=576; yrs=3
g2=tcv/(n*H*yrs)
add(source="WhiteFiber Q2-26 results press release (corrected) + 8-K Ex-99.1 investor deck",
    date="2026-08-12", provider="WhiteFiber (WYFI) (customer: Prime Intellect)", chip="NVIDIA VR200 (Vera Rubin), 576 GPUs (= 8 NVL72 racks)",
    p=g2, term="3 yr; service targeted Q2 2027; Canada", loc="Contract (cloud services agreement)",
    mwv=mw(g2), tag="Derived",
    url="https://www.prnewswire.com/news-releases/whitefiber-inc-reports-second-quarter-2026-results-302848971.html ; https://www.sec.gov/Archives/edgar/data/2042022/000121390026099424/ea030525701ex99-1.htm",
    notes=f"PR: '3-year ... deploy 576 NVIDIA VR200 GPUs in Canada ... approximately $108 million'; TradingView/Zacks cites $108.2M. Math: $108.2M / (576 x 8,760 x 3) = ${g2:.2f}/GPU-hr (ratable, 100% of hours billed, all TCV treated as compute -> may include storage/network; any ramp would RAISE the rate). $/IT MW-yr uses model's {RUB_MW} GPU/MW assumption (not disclosed). Small deal (~2.5-3 MW), shorter tenor, AI-native counterparty -> expect above hyperscaler-anchor pricing.")

# 3 WhiteFiber same-provider B300 reference
g3=87.5e6/(576*H*5)
add(source="WhiteFiber Q2-26 results press release (same release) - REFERENCE ROW (B300, not Rubin)",
    date="2026-08-12", provider="WhiteFiber (customer undisclosed; Iceland)", chip="NVIDIA B300 (HGX), 576 GPUs",
    p=g3, term="5 yr initial term; Iceland; + revenue-share upside", loc="Contract",
    mwv=None, tag="Derived",
    url="https://www.prnewswire.com/news-releases/whitefiber-inc-reports-second-quarter-2026-results-302848971.html",
    notes=f"$87.5M / (576 x 8,760 x 5) = ${g3:.2f}/GPU-hr. Same provider, same GPU count: VR200 3-yr ${g2:.2f} vs B300 5-yr ${g3:.2f} = {g2/g3:.2f}x. NOT apples-to-apples (HGX B300 vs rack-scale VR NVL72; 5 vs 3 yr; Iceland vs Canada; revenue share excluded). Shown only as an upper-bound sanity check on within-provider premium.")

# 4 Civo list
g4=11.00
add(source="Civo Vera Rubin page + CTO blog 'Why we reserved 2,016 Vera Rubins'",
    date="2026-05-26 (blog); page live 2026-10-04", provider="Civo (UK)", chip="Vera Rubin NVL72 (2,016 GPUs reserved)",
    p=g4, term="Reserved only; Q1 2027 delivery; term/deposit unspecified ('from')", loc="List (indicative early-access 'from' price)",
    mwv=mw(g4), tag="Reported",
    url="https://www.civo.com/ai/vera-rubin-gpu ; https://www.civo.com/blog/why-we-reserved-2016-vera-rubins-and-what-it-took ; https://www.civo.com/pricing",
    notes="'Early access pricing from $11.00/hr ... rates varying by deposit level and contract length.' Unit is not explicitly labelled per-GPU; assumed per GPU-hr (Civo prices GPUs per GPU; a per-rack $11 is implausible). Same-provider reference (Civo pricing page, 2026-10-04): 8x B300 node $59.92/hr on-demand ($7.49/GPU), $43.92/hr 36-mo ($5.49/GPU) -> VR 'from' $11.00 = 1.47x B300 on-demand, 2.00x B300 36-mo. Marketing list price, UK sovereign niche; LOW weight for a 2027+ contracted-MW model.")

# 5 SemiAnalysis TCO (cost side)
add(source="SemiAnalysis, 'Vera Rubin NVL72 vs GB200 NVL72? Inference TCO' (InferenceX) + InferenceX compare page (July 2026 AI Cloud TCO model)",
    date="2026-07-23", provider="SemiAnalysis (cost estimate, not a price)", chip="VR200 NVL72 vs GB300 NVL72",
    p=3.57, term="n/a (ownership TCO per GPU-hr)", loc="Cost (TCO) estimate",
    mwv=None, tag="Estimated",
    url="https://inferencex.semianalysis.com/blog/vera-rubin-nvl72-vs-gb200-nvl72-inference ; https://inferencex.semianalysis.com/compare-per-dollar/deepseek-v4-gb300-vs-vr200",
    notes="SemiAnalysis TCO model: Rubin $3.57 vs GB300 $2.36/GPU-hr (operator ownership; 'not rental prices') = 1.51x; hyperscaler-volume tier $3.61 vs $2.31 = 1.56x. Premium columns for this row show the COST ratio vs GB300 TCO, not vs $5.10/$4.65 rents. Method is SemiAnalysis's (capex+opex+DC), not ours.")

# 6 Morgan Stanley BOM
add(source="Morgan Stanley (Howard Kao) Rubin rack BOM, via Locsic / Wccftech",
    date="2026-05-21", provider="Morgan Stanley (hardware ASP estimate)", chip="VR200 NVL72 rack vs GB300 NVL72 rack",
    p=None, term="n/a", loc="Cost (system ASP) estimate",
    mwv=None, tag="Estimated",
    url="https://locsic.com/thinking/morgan-stanley-rubin-bom-teardown/ ; https://wccftech.com/bernstein-warns-nvidias-vera-rubin-racks-will-hit-9-1-million-as-hbm4-prices-triple-to-53-per-gigabyte/",
    notes="ODM rack price ~$7.8M VR200 vs ~$4.0M GB300 (72 GPUs each) = 1.95x per GPU (MS: +95%). If hyperscaler buys SOCAMM direct: ~$6.7M = 1.68x. Bernstein (8-Jun-2026): up to $9.1M on $53/GB HBM4 -> 2.28x vs MS's $4.0M GB300 (mixed sources). Hardware only, ex-DC.")

# 7 Tom's Hardware quotes
add(source="Tom's Hardware (Anton Shilov), unnamed source + DigiTimes",
    date="2026-03-24", provider="Industry quotes (hardware)", chip="VR200 NVL72 vs GB300 NVL72",
    p=None, term="n/a", loc="Cost (system quote)",
    mwv=None, tag="Reported",
    url="https://www.tomshardware.com/tech-industry/artificial-intelligence/price-of-nvidias-vera-rubin-nvl72-racks-skyrockets-to-as-much-as-usd8-8-million-apiece-but-server-makers-margins-will-be-tight-nvidia-is-moving-closer-to-shipping-entire-full-scale-systems",
    notes="VR200 NVL72 quoted $5-7M (incl. ~$1M NAND) vs GB300 NVL72 $6-6.5M (inference-optimized config) -> 0.77-1.17x. CONFLICTS with MS ($4M GB300); GB300 config likely loaded. Low weight; shown for completeness.")

# 8 Not derivable
add(source="Nebius 6-K + Ex-99.1 (Meta Infrastructure Service Agreement)",
    date="2026-03-16", provider="Nebius (customer: Meta)", chip="Vera Rubin (dedicated clusters)",
    p=None, term="5 yr per order; tranches from early 2027", loc="Contract",
    mwv=None, tag="Reported (no unit price derivable)",
    url="https://www.sec.gov/Archives/edgar/data/1513845/000110465926027886/tm268879d1_6k.htm",
    notes="$12B dedicated Rubin capacity (incl. storage/connectivity) + up to $15B unsold-capacity backstop. No MW, GPU count or $/GPU-hr disclosed -> cannot derive. Break-even (existing contract_economics): ~149 IT MW to match Anthropic-Nscale $16.16M/MW-yr.")
add(source="Nscale/Microsoft 1.35 GW Vera Rubin LOI press release; Nscale Europe Rubin release",
    date="2026-03-16", provider="Nscale (customer: Microsoft)", chip="Vera Rubin NVL72",
    p=None, term="multi-year compute term + DC lease; tranches from late 2027", loc="LOI",
    mwv=None, tag="Reported (no $)",
    url="https://www.nscale.com/press-releases/nscale-west-virginia-ai-factory ; https://www.nscale.com/press-releases/nscale-nvidia-vera-rubin",
    notes="No dollar value or unit price. (Nscale F-1 'up to $43.8B' Microsoft SOWs are mostly GB300-identified; not attributable to Rubin.)")
add(source="Public rate trackers: SAVRN AI Index, ComputeStacker, CoreWeave Rubin page/pricing, TensorBay, Accruex, Spheron",
    date="2026-09-16 to 2026-10-04", provider="CoreWeave, Lambda, Crusoe, Together, Nebius, Oracle, Azure, AWS, GCP, Vast/SF Compute etc.", chip="Vera Rubin NVL72",
    p=None, term="n/a", loc="List (none published)",
    mwv=None, tag="Reported (absence)",
    url="https://savrn.com/ai-index/pricing/gpus/vera-rubin ; https://computestacker.com/insights/vera-rubin-nvl72-coreweave/ ; https://coreweave.com/products/nvidia-vera-rubin ; https://tensorbay.io/gpu/vr200-nvl72/ ; https://accruex.com/vera-rubin/",
    notes="SAVRN: '0 on-demand, 0 reserved and 0 spot prices from 0 providers'; 'NVIDIA has published no Vera Rubin list price'. CoreWeave GA'd VR NVL72 30-Sep-2026 (Cognition first) - capacity-planning/contact only. TensorBay/Accruex: quote on request. No Ornn/Silicon Data Rubin index found. Spheron's '30-50% premium' is an unsourced expectation - excluded.")
add(source="Nebius Q2-26 earnings call (Motley Fool transcript) - CONTEXT, not Rubin",
    date="2026-08-19", provider="Nebius", chip="Blackwell generation (auction); mid-term deals gen unspecified",
    p=None, term="1-3 yr mid-term; <=6 mo short-term", loc="Contract (disclosed yields)",
    mwv=None, tag="Reported (excluded - not Rubin)",
    url="https://www.fool.com/earnings/call-transcripts/2026/08/19/nebius-nbis-q2-2026-earnings-call-transcript/",
    notes="Auction 'cleared at the highest price we have seen for the Blackwell generation ... 15% above'. Mid-term deals $20-25M/MW; short-term $40-50M/MW. Analyst asked whether $40M/MW reflected Rubin pricing; management did not confirm (answered on Rubin deployment timing). Do NOT use as Rubin premium.")

wb=Workbook(); ws=wb.active; ws.title="Rubin_price_points"
hdr=["source","date","provider","chip","$/GPU-hr","term","list_or_contract","implied premium vs GB300 $5.10","implied premium vs GB300 $4.65",
     "$/IT MW/yr ($M) if derivable","implied rev/IT MW vs GB300 @ $5.10 x 491 GPU/MW ($21.94M)","tag","URL","notes"]
ws.append(hdr)
gbmw=GB_A*GB_MW*H/1e6
for r in rows:
    p=r["p"]
    if r["loc"].startswith("Cost (TCO)"):
        pa=round(3.57/2.36,2); pb="1.56x (hyperscaler tier 3.61/2.31)"
    elif p is None: pa=pb=None
    else: pa=round(p/GB_A,2); pb=round(p/GB_B,2)
    m=r["mwv"]
    ws.append([r["source"],r["date"],r["provider"],r["chip"],round(p,2) if p else None,r["term"],r["loc"],pa,pb,
               round(m,2) if m else None, round(m/gbmw,2) if m else None, r["tag"],r["url"],r["notes"]])
ws2=wb.create_sheet("GPUs_per_MW_check")
ws2.append(["source","date","figure","implied Rubin GPUs per MW","tag","URL","notes"])
ws2.append(["NVIDIA Vera Rubin NVL72 product page - '100 MW AI Factory Specifications'","live 2026-10-04","40K Rubin GPUs in 100 MW (DSX, MaxLPS)",400,"Reported","https://www.nvidia.com/en-us/data-center/vera-rubin-nvl72/","Footnote: 'at-scale AI factory using NVIDIA DSX with MaxLPS'. Not stated whether 100 MW is IT or utility; likely an optimistic, power-smoothed design point."])
ws2.append(["Hashrate Index (Luxor) spec breakdown","2026-03-20","Rubin GPU ~1.8 kW (Max Q) / ~2.3 kW (Max P) TGP; NVL72 rack ~190 kW / ~230 kW","379 (190 kW) / 313 (230 kW) per MW of rack power","Derived","https://hashrateindex.com/blog/nvidia-vera-rubin-nvl72-specs-breakdown/","72 / 0.190 = 379; 72 / 0.230 = 313. Rack power only; excludes network/storage/CPU racks (~3-18% more) -> ~300-365 per IT MW. Not an NVIDIA datasheet figure."])
ws2.append(["Tom's Hardware","2026-03-24","GPU power 1.4 kW (Blackwell Ultra) -> 1.8 kW and potentially higher","n/a","Reported","https://www.tomshardware.com/tech-industry/artificial-intelligence/price-of-nvidias-vera-rubin-nvl72-racks-skyrockets-to-as-much-as-usd8-8-million-apiece-but-server-makers-margins-will-be-tight-nvidia-is-moving-closer-to-shipping-entire-full-scale-systems","Consistent with Max-Q 1.8 kW."])
ws2.append(["Third-party read of Nscale (Shanu Mathew thread)","2026-08","~430,000 GPUs / 1,350 MW","319","Derived (by third party)","https://www.unrollnow.com/status/2092716897766723983","Attributed to an Nscale PR, but the Nscale 16-Mar-2026 release does not state 430k GPUs - treat as unverified."])
ws2.append(["Model assumption (build_model.py GPUMW_N)","-","320 GPU/IT MW (~3.1 kW/GPU all-in IT)",320,"Placeholder","/workspace/nebius/model/build_model.py","Verdict: plausible but at the conservative (low-density) end. Evidence range ~300-400. Break-even premium = 491/GPU-per-MW: 1.53x @320, 1.40x @350, 1.29x @380, 1.23x @400."])
for w in (ws,ws2):
    for c in w[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78"); c.alignment=Alignment(wrap_text=True,vertical="top")
    for row in w.iter_rows(min_row=2):
        for c in row: c.alignment=Alignment(wrap_text=True,vertical="top")
    w.freeze_panes="A2"
for i,wd in enumerate([38,16,26,26,10,26,22,14,14,14,16,16,50,90],1): ws.column_dimensions[get_column_letter(i)].width=wd
for i,wd in enumerate([40,14,40,22,16,50,80],1): ws2.column_dimensions[get_column_letter(i)].width=wd
out="/workspace/nebius/parts/_rebuild/rubin_price_points.xlsx"; wb.save(out); print("saved",out)
for r in rows: print(r["provider"][:40], r["p"] and round(r["p"],3), r["mwv"] and round(r["mwv"],2))
print("GB300 rev/MW",gbmw)
