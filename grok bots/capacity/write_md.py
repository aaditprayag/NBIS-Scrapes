import pickle
cols,R,K,T,S=pickle.load(open('/workspace/nebius/capacity/data.pkl','rb'))
Y,tot=pickle.load(open('/workspace/nebius/capacity/ye26.pkl','rb'))
esc=lambda x:str(x).replace('|','\\|').replace('\n',' ')
L=[]
L.append("# Nebius Group (NBIS) – Announced data center sites & capacity targets\n")
L.append("_Compiled Oct 4, 2026 (ET). Step 1 of independent capacity verification. Everything is **as announced** by Nebius or credible press/partners. Anything marked **INFERENCE** or **Estimated** is mine. 'n/d' = not disclosed. Coordinates are approximate (city or intersection level) and are only meant for mapping._\n")
L.append("## 1. Site table (23 rows)\n")
short=["#","Site","City/Region","Country","Lat","Lon","Type","Partner/Landlord","Announced MW – initial phase","Announced MW – full build-out","Timeline (announced; revisions)","GPU generation(s)","Anchor customer","Status (as of Oct 4 2026)","Sources (date)"]
idx=[cols.index(c) for c in short]
L.append("| "+" | ".join(short)+" |"); L.append("|"+"---|"*len(short))
for r in R: L.append("| "+" | ".join(esc(r[i]) for i in idx)+" |")
L.append("\nThe IT-vs-gross MW split is **not disclosed for any site** except AIB (50 MW critical IT on a 65 MW utility load). The CSV/XLSX also include a coordinate-basis column.\n")
L.append("## 2. Company-level capacity targets and how they changed\n")
L.append("**Definitions** (Q1'26 letter, May 13 2026, "+S['q1_26']+"):\n- **Contracted** = 'power secured by contracted land and power commitments'\n- **Connected** = 'power connected into fully built and equipped data centers'\n- **Active** = 'power being consumed by installed, operational IT equipment and available for revenue generation'\n\nThe Q2'25 letter (Aug 7 2025) said connected power 'includes both active power and power that is fully provisioned in our data centers and can be activated immediately upon GPU installation'.\n")
L.append("| Date | Target / disclosure | Source |\n|---|---|---|")
for t in T: L.append("| "+" | ".join(esc(x) for x in t)+" |")
L.append("\n**Path of the YE2026 contracted-power target:** 1 GW (Aug-25), then >2.5 GW (Nov-25), >3 GW (Feb-26), >4 GW (May-26) and 5 GW (Aug-26). The **YE2026 connected-power** guide has stayed at **800 MW–1 GW** since Nov 11 2025. The YE2025 active-power target was 100 MW and the actual came in at ~170 MW.\n")
L.append("## 3. YE2026 connected-power guide vs. a per-site energization build (priority test)\n")
L.append("**Exact wording of the guide:**\n- Nov 11 2025 (Q3'25 letter): 'we expect to have 800MW to 1GW of connected power (i.e, power connected to built data centers) by the end of 2026'. "+S['q3_25']+"\n- Feb 12 2026 (Q4'25 letter): 'we also enter 2026 with strengthened conviction in achieving our 800 MW to 1 GW of connected power by the end of this year'. "+S['q4_25']+"\n- May 13 2026 (Q1'26 letter): 'We continue to expect 800MW to 1GW of connected power by year-end.' "+S['q1_26']+"\n- Aug 12 2026: the Q2'26 letter does **not** restate the guide. On the call, CIPO Andrey Korolenko said: 'We still expect to meet this guidance from 800 MW to 1 GW of connected power this year… connected power represents the data centers, and there are a few steps coming from the data centers to the revenues. You have to commission the data center, build the network, build the clusters, deploy the platform, then onboarding the customers… That takes a few months… In terms of our guidance from 800 MW to 1 GW, yes, that is true about the connected, but I would think that about being active throughout the first half of 2027.' "+S['q2_26call']+"\n")
L.append("**Definitional nuance:** 'Connected' means built and equipped data-center power that is ready for GPUs. It is **not** energized IT load that is billing (that is 'active'), and it is **not** secured grid power (that is 'contracted'). Management has explicitly said active power, and therefore revenue, ramps **through 1H 2027**. A model that bills 0.8–1 GW for all of 2027 runs ahead of what management has said. Nebius has never said whether the MW figures are IT or gross/utility load. Mäntsälä's PUE is ~1.1, so the gap is small for owned sites but could matter for colocation sites.\n")
L.append("**Per-site build (INFERENCE; includes only sites with a disclosed MW and a disclosed YE26 date):**\n")
L.append("| Site | Low MW | Base MW | High MW | Basis | Source |\n|---|---|---|---|---|---|")
for y in Y: L.append("| "+" | ".join(esc(x) for x in y)+" |")
L.append(f"| **TOTAL identified** | **{tot[0]:.0f}** | **{tot[1]:.0f}** | **{tot[2]:.0f}** | vs. guide 800–1,000 MW connected | |")
L.append(f"\nThe identified sites reach ~800 MW only in the high case. That case needs every stated YE26 milestone to be hit: Vineland at its full 300 MW, Béthune at 120 MW, Birmingham phase 1 at ~100 MW, and Beit Shemesh fully delivered. To reach the guide in the base case, about {800-tot[1]:.0f}–{1000-tot[1]:.0f} MW would have to come from sites where MW is not disclosed: the Independence initial phase (IPL power 2H26), Minnesota, Oklahoma, France's second colocation site, the third UK site, and the KC/Vineland upside.\n")
L.append("## 4. Historical KPIs (Q1'24–Q2'26)\nThe full site × quarter table is in the XLSX tabs 'Site x Quarter KPIs' and 'KPI pivot'. Every cell is tagged Reported, Derived or Estimated. The company-level hard datapoints are:\n- ~10k GPUs at Q3'24 (Estimated: '>20k by YE24' target ÷ 'nearly doubled QoQ' in Q4)\n- >20k GPUs at Q4'24 (Estimated from the target)\n- 25 MW connected at YE24 (Reported, Q3'25 chart)\n- ~38k GPUs at end-Q1'25, mostly H200 (Reported, Feb 20 2025 call)\n- ~170 MW active at YE25 (Reported)\n\nNo active MW or GPU count is disclosed for Q1'24, Q2'24, Q2'25, Q3'25, Q1'26 or Q2'26.\n")
L.append("## 5. Gaps / conflicting information\n")
G=[
"**Site count.** The Q4'25 letter says 16 sites, the 20-F says 'more than 16', and the Sept 2026 local press (Pottsville Republican Herald) says 19. Reuters (Mar 31 2026) calls Lappeenranta Nebius's '10th site', which conflicts with the other counts. My table has 23 rows, including unnamed sites and the AIB deal.",
"**Lappeenranta ownership.** Nebius calls it a 'new owned location' (Q1'26 letter). Developer Polarnode headlines that 'Nebius to operate Polarnode's 310 MW campus'. Lease vs. owned matters for capex.",
"**Béthune model.** The 20-F classes the Feb-2026 French sites as colocation, and DCmag/Reuters say Azur Datacenter finances the build. Nebius PRs call it 'building an AI factory'. The full build-out date is '2027' (Reuters/DCD) vs. 'end-2027' (DCmag/GNT).",
"**Owned sites total.** The Q1'26 letter says 'owned facilities will deliver 3 GW across five sites'. My INFERENCE: Independence 1.2 + Pennsylvania 1.2 + Lappeenranta 0.31 + Birmingham ~0.3 (press figure only) + Mäntsälä 0.075 = ~3.085 GW. Birmingham's 300 MW comes from press, not from Nebius.",
"**Independence MO capacity and buildings.** Capacity was first reported as 800 MW and is now up to 1.2 GW. Nebius's page says four buildings over 5–7 years; DCD says up to 10 buildings.",
"**Birmingham timing.** Nebius's community page targets initial-phase energization in Q4 2026, while the Q1'26 letter says the site is operational in 2027.",
"**Pennsylvania phase 1.** The Q1'26 call said 250–300 MW by end-2027 and 1.2 GW by ~2030; DCD/site plans say 260 MW by Oct 2027 and 660 MW by 2029. Earlier, Volozh reportedly said 250–350 MW live by end-2027.",
"**Contracted power is not site-attributed.** The 5 GW YE26 target compares with ~4.0 GW of announced site MW (INFERENCE sum, which includes UK 65 MW + Longcross 16 MW). That leaves ~1 GW unattributed. Management also mentions 'hundreds of MW' of behind-the-meter generation and the Bloom partnership.",
"**Anchor customers by site.** Microsoft is tied to Vineland only. The Meta #1 deal (~$2.9B, Nov 2025, tranches Dec-25/Feb-26) has no site disclosed. Meta #2 ($12B dedicated + up to $15B contingent, Vera Rubin, from early 2027) covers 'multiple locations' that are not named.",
"**Vineland MW actually live is not disclosed.** Tranche counts are known (9 + 2 tranches, all delivered to date per Q2'26) but not their MW. Power has moved to Bloom fuel cells, and the site-plan amendment hearing was adjourned (Q2'26 call).",
"**UK.** £1.7bn covers 'four sites' / 'three new deployments' totalling 65 MW by 2027. Only Harlow (22 MW) and the Longcross expansion are named. Whether 65 MW includes the original 16 MW at Longcross is ambiguous.",
"**Not disclosed for any site:** MW for Paris, Oklahoma, Minnesota or France's second site, and the landlords for Oklahoma and Minnesota. The Kansas City phase-2 MW is also not disclosed.",
"**New site not yet in a Nebius release:** AIB Data Centers (Sept 30 2026), 50 MW critical IT in the southeastern US.",
"**Third-party aggregator claims are excluded** (worldmetrics, Northwise, DC Atlas) except where noted as Estimated/low-confidence (Paris 5 MW; Harlow 4.4 MW split).",
]
for g in G: L.append("- "+g)
L.append("\nFiles: site_list.md / site_list.csv / site_list.xlsx (tabs: Sites, Site x Quarter KPIs, KPI pivot, YE2026 energized vs guide, Company targets history) in /workspace/nebius/capacity/. Source PDFs are cached in /workspace/nebius/src/.")
open('/workspace/nebius/capacity/site_list.md','w').write("\n".join(L))
print(len(L))
