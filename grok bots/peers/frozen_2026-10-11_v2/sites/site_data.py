# Content for the peer site appendices (Oct 8 2026 ET). Every line carries a tag (R reported, P press/third party, D derived,
# E estimated, n/d not disclosed) and a source number. Read by build_site_appendices.py.
ASOF = "Thu Oct 8, 2026, ~3:15 PM ET"
DEFS = ("Definitions (team, as in capacity/vineland_site_appendix): Announced = company-stated build-out. Connected = power delivered into fully built, "
        "equipped halls (can sit idle until GPUs arrive); grid/utility MW is shown on its own row and is NOT IT MW. Active = power used by installed, operating IT. "
        "Billed = revenue-generating IT MW. Tags: R reported (company/regulator), P press/third party, D derived, E estimated, n/d not disclosed. [n] = sources.")
T5 = "https://www.tceq.texas.gov/assets/public/permitting/air/reports/applications/titlev-pending-permits.html"
NSR = "https://www.tceq.texas.gov/assets/public/permitting/air/reports/applications/nsr-pending-permits.html"
GOV = "https://gov.texas.gov/news/post/governor-abbott-directs-tceq-to-halt-data-center-permits"
CLAW = "https://computelaw.blog/brief/texas-governor-pauses-environmental-permits-for-data-centers-until-state-audits-end/"
ORCL = "https://www.fool.com/earnings/call-transcripts/2026/09/11/oracle-orcl-q1-2027-earnings-call-transcript/"

SITES = {}

SITES["coreweave_denton"] = dict(
 title="Appendix: CoreWeave Denton, TX (Core Scientific landlord) site status, as of " + ASOF,
 sub="Nebius Peer KPIs. Public sources only. Grade: Green (unchanged since Oct 5). Same layout and definitions as the Vineland appendix.",
 facts=[["Owner / landlord", "Core Scientific (Core Scientific Denton LLC); City of Denton ground lease to 12/31/2035, extends with the PPA", "R", "1"],
        ["Tenant / operator", "CoreWeave (colocation licence; CoreWeave pays utility costs); end customer n/d", "R", "1"],
        ["Location / build", "8171 Jim Christal Rd, Denton TX; ~64 ac (78.85 ac after 2024 amendment); 8 bldgs A–H, 370,284 sf; ex-bitcoin campus, HPC conversion from Q3'24; GC T5", "R", "1, 7"]],
 mw=[["Announced", "262 IT / 394 grid", "full build", "R", "1, 2"],
     ["Connected (grid / utility)", "297 energized + 74 to energize", "grid, not IT", "R", "3"],
     ["Connected (halls)", "~260", "IT; 'substantially complete'", "R", "5"],
     ["Active", "~260 (250–262)", "IT; landlord says billing", "D", "5"],
     ["Billed", "~260", "IT; CORZ billing CoreWeave, mid-Jul", "R", "5"],
     ["YE26 billed L / B / H", "260 / 262 / 262", "IT", "E", "6"]],
 mw_note=("Billed = landlord (CORZ) billing CoreWeave; CoreWeave's own billed MW at Denton n/d. Apr 21 8-K: Bldgs D–G ~132 MW billing, rest H2'26; Jul 28 deck: ~260 MW "
          "billing of >430 MW CORZ total [1, 5]. City (Sep 22) cites 391 MW electrical capacity [8]."),
 chart=dict(ann=262, grid=297, grid_lbl="297 grid\n(+74 pending)", conn=260, act=(250, 260, 262), bill=260,
            tags=("R", "R", "R", "D", "R")),
 power=[["Utility / ISO", "Denton Municipal Electric PPA (~394 MW potential); ERCOT", "R", "2, 7"],
        ["ERCOT status", "297 MW 'Base Load, Pathway (a)' (energized before 3/25/2022), outside Batch Zero; +74 MW validated in ERCOT 2025 RTP, outside Batch Zero", "R", "3"],
        ["Substations / grid", "2 new on-site substations (CORZ-funded); DME/ERCOT upgrades by 2029; city: transmission near limits", "R", "7, 8"],
        ["PPA headroom", "DME (Sep 22): CORZ not yet at its PPA MW cap; no extra power requested; asked to lease more DME land 3–4 months earlier", "R", "8"],
        ["On-site generation", "Emergency / standby gensets only (PBR); campus is grid-powered", "R", "9"]],
 permits=[["TCEQ PBR 179143", "§106.511 standby engines/turbines, 'Core Scientific Data Center Site 1 And 2 Denton', ack 2025-02-11 (RN111355798)", "R", "9"],
          ["Major NSR / PSD / Title V", "None found: no Denton DC entry on TCEQ Title V list (data 10/8) or NSR list (run 10/8)", "R absence", "10"],
          ["City", "PPA + land-lease amendments approved 2024-11-19; no city tax incentives requested", "R", "7"]],
 evidence=[["Sentinel-2 Oct 5 v Sep 15", "Fully roofed campus, no new pads; white-roof 17.6 v 18.7 ha (same metric; illumination noise). Not load evidence", "D; read E", "11"],
           ["2023 → 2026 imagery", "White-roof ~9.8 → ~18.7 ha; south pad from Oct-25; thermal confounded by mining-era load", "D; E", "6"],
           ["Contract economics", "CORZ–CoreWeave ~590 MW, '$10B+' over the term → ~$1.4M per MW-yr colocation (D, ~12 yrs). Facility rent, not compute $/MW. Denton $ n/d", "R; D", "5"]],
 what=("(1) CORZ Q3 results ~Oct 23 [E] and CoreWeave Q3 ~early Nov [E]: confirm 262 MW complete/billing and a date for the +74 MW (bull: expansion past 262) [5, 12]. "
       "(2) Denton moratorium: 2nd hearing Oct 27, vote as early as Dec 1. City attorney: it 'does not necessarily stop an existing project', but CORZ's extra-land request could be caught "
       "(bear for expansion only) [8]. (3) ERCOT large-load audit reports Dec 10, PUCT Dec 17: reclassifying the +74 MW is the bear [3, 13]. "
       "Biggest risk: low for the 262 MW in service; only expansion optionality is exposed."),
 changed="Since Oct 5: Denton moratorium process (Oct 27 hearing; vote ≥ Dec 1) and CORZ land request added; Oct 5 Sentinel-2 scene shows no change; MW unchanged.",
 sources=[
  "[1] CORZ 8-K Ex.99.2 (2026-04-21) https://www.sec.gov/Archives/edgar/data/1839341/000119312526165121/d149019dex992.htm (acc. Oct 8)",
  "[2] CORZ/CoreWeave PR (2025-02-26) https://www.businesswire.com/news/home/20250226866161/en/Core-Scientific-and-CoreWeave-Announce-%241.2-Billion-Expansion-at-Denton-TX-Site (acc. Oct 5)",
  "[3] CORZ 8-K (2026-09-10, ERCOT status) https://investors.corescientific.com/sec-filings/all-sec-filings/content/0001839341-26-000023/0001839341-26-000023.pdf (acc. Oct 8)",
  "[5] CORZ Q2'26 deck (2026-07-28) https://www.sec.gov/Archives/edgar/data/1839341/000183934126000013/q2fy26earningsdeck728am.htm (acc. Oct 5; sites/docs/corz_q226_deck.txt)",
  "[6] Team: peers/sites/site_verification.md §1; sites.xlsx; strips/denton_tx_strip.png",
  "[7] City of Denton (2024-11-20) https://www.cityofdenton.com/CivicAlerts.aspx?AID=922 (acc. Oct 5)",
  "[8] Community Impact (2026-09-24) https://communityimpact.com/denton/government/denton-officials-weigh-data-center-moratorium-during-first-public-hearing/ (acc. Oct 8)",
  "[9] TCEQ PBR 179143 https://records.tceq.texas.gov/cs/idcplg?IdcService=TCEQ_EXTERNAL_SEARCH_GET_FILE&Rendition=Web&dID=8656717 ; sites/docs/denton_tceq_pbr_179143.pdf (acc. Oct 5)",
  f"[10] TCEQ Title V pending (data 10/8/2026) {T5} ; NSR pending (run 10/8/2026) {NSR} (acc. Oct 8)",
  "[11] Sentinel-2 L2A S2B_14SPB_20261005 (Earth Search); sites/imagery/denton_tx/detail/denton_tx_2026-10-05_crop_oct8.png",
  "[12] CORZ Q3 date (third party) https://www.marketbeat.com/earnings/reports/2026-10-23-core-scientific-inc-stock/ (acc. Oct 8)",
  f"[13] Gov. Abbott (2026-09-21) {GOV} ; ERCOT audit timeline {CLAW} (acc. Oct 8)"])

SITES["iren_childress"] = dict(
 title="Appendix: IREN Childress, TX (Microsoft Horizons) site status, as of " + ASOF,
 sub="Nebius Peer KPIs. Public sources only. Grade: Green (H1–2) / Yellow (H3–4 by Q4'26), unchanged since Oct 5. Site grid MW includes legacy bitcoin load.",
 facts=[["Owner / operator", "IREN (SPV Ie US Development Holdings 3 Inc.); IREN-owned 576 ac", "R", "1, 2"],
        ["Tenants", "Microsoft: Horizons 1–4, 4 × 50 MW IT liquid-cooled GB300 NVL72, $9.7B / 5 yrs. NVIDIA: $3.4B in ~60 MW of existing halls from early 2027", "R", "3, 4"],
        ["Location / legacy", "620 FM 1033, Childress County TX; Phases 1–5 = 650 MW bitcoin 'operating data centers' (Aug-25), winding down", "R", "2, 5"]],
 mw=[["Announced", "750 grid; 200 IT (H1–4) + ~60 (NVIDIA)", "H5–6 (2027) MW n/d", "R", "3, 4, 6"],
     ["Connected (grid / utility)", "750", "grid incl. mining load", "R", "6"],
     ["Connected (halls)", "~100 (H1 + H2 commissioning)", "IT", "D", "3"],
     ["Active", "75 (50–100)", "IT, AI only", "D", "3, 7"],
     ["Billed", "50 (H1 accepted Aug 13)", "IT; H2 n/d until acceptance", "D on R", "7"],
     ["YE26 billed L / B / H", "100 / 150 / 200", "IT", "E", "8"]],
 mw_note=("Site grid load ≠ AI IT: the 750 MW interconnection still carries mining load. H2 'commissioning underway' and H3–4 'late-stage construction, targeting delivery "
          "in Q4 2026' (Aug 27) [3]; no H2 acceptance found as of Oct 8."),
 chart=dict(ann=200, grid=750, grid_lbl="750 grid\n(incl. mining)", conn=100, act=(50, 75, 100), bill=50, tags=("R", "R", "D", "D", "D")),
 power=[["Utility / ISO", "AEP Texas / ERCOT; IREN-owned substations, dual diverse connection", "R", "2"],
        ["ERCOT status", "Existing energized load (raised 600 → 750 MW on 2024-07-24 after ERCOT study); not a Batch Zero new-interconnect case; 2nd bulk transformer energized for H1–2", "R", "5, 6"],
        ["Regional grid", "ERCOT-endorsed AEP Childress Area Transmission Improvement (Tier 2, ~$77.3M, ISD Nov 2031, CCN needed): area reliability, not site supply", "R", "9"],
        ["On-site generation", "None found; Title V filing implies permitted emission units (unit list n/d)", "R abs.; E", "10"]],
 permits=[["TCEQ Title V #4951", "'Childress Data Center', Ie US Development Holdings 3; received 2026-08-17; Administrative Complete (list data 10/8, unchanged since Oct 5)", "R", "10"],
          ["Major NSR / PSD", "None found (NSR pending list run 10/8: no IREN / Childress DC entry)", "R absence", "11"],
          ["Sep 21 state pause", "TCEQ issues no data-center permits until the ERCOT audit; Childress is on third-party lists of pending DC applications, so #4951 waits; operations continue meanwhile [E]", "R; P; E", "12"],
          ["County", "Ch. 312 abatements / Iris Reinvestment Zone (IE US entities)", "R", "13"]],
 evidence=[["Sentinel-2 Oct 5 v Sep 28", "Mining sheds + new west/north halls + grey SE facility; white-roof 9.0 v 12.8 ha (same metric; sun/view noise, no demolition seen). H2–H4 not separable at 10 m", "D; read E", "14"],
           ["2023 → 2026 imagery", "White-roof 0.5 → 19–25 ha (Mar–Jul 26); thermal dominated by mining", "D; E", "8"],
           ["Contract economics", "Microsoft $9.7B / 5 yrs / 200 IT MW = $9.7M per IT MW-yr (bottom of the $9.7–16.7M comp band). NVIDIA $3.4B / 5 / ~60 = ~$11.3M (MW unlabeled). IREN: 'active discussions at ~$25M per MW (IT)', not signed", "R; D", "3, 15"]],
 what=("(1) Horizon 2 acceptance PR and H3–4 delivery in Q4'26: each accepted 50 MW adds billed MW; a slip past Dec 31 leaves YE26 at the 100 MW low [3, 7]. "
       "(2) IREN Q1 FY27 results ~early–mid Nov [E]: H2–H4 status and NVIDIA 60 MW ramp [3]. "
       "(3) TCEQ compliance report Oct 19; ERCOT audit Dec 10 / PUCT Dec 17: #4951 stays unissued under the pause; bear only if the state moves against operating sites [12]. "
       "Biggest trigger: H2–H4 acceptance timing (YE26 billed 100 / 150 / 200)."),
 changed="Since Oct 5: no IREN 8-K; #4951 still Administrative Complete (TCEQ data 10/8) and now explicitly inside the Sep 21 permit pause; H2 still not accepted; Oct 5 Sentinel-2 adds nothing.",
 sources=[
  "[1] Texas Comptroller data-center list (LD926028) https://comptroller.texas.gov/taxes/data-centers/data-center-lists.php (acc. Oct 5)",
  "[2] IREN Childress page https://iren.com/data-centers/childress (acc. Oct 5)",
  "[3] IREN FY26 results (2026-08-27) https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm (acc. Oct 8; sites/docs/iren_fy26.txt)",
  "[4] IREN Q3 FY26 (2026-05-07, NVIDIA $3.4B) via peers/iren/iren_kpis.md",
  "[5] IREN Aug-2025 update https://iren.gcs-web.com/static-files/ec8ec2a6-a1d1-4bb6-9cb9-bc2d436b9db9 (acc. Oct 5)",
  "[6] IREN PR (2024-07-24) https://www.nasdaq.com/press-release/iren-increases-childress-power-capacity-600mw-750mw-2024-07-24 (acc. Oct 5)",
  "[7] IREN 8-K Ex.99.1 (2026-08-13, H1 accepted) https://www.sec.gov/Archives/edgar/data/1878848/000114036126032638/ef20080141_ex99-1.htm (acc. Oct 8)",
  "[8] Team: peers/sites/site_verification.md §2; sites.xlsx; strips/childress_tx_strip.png",
  "[9] PUC Docket 55999 Item 173 https://interchange.puc.texas.gov/Documents/55999_173_1556587.PDF (acc. Oct 5)",
  f"[10] TCEQ Title V pending (data 10/8/2026) {T5} (acc. Oct 8)",
  f"[11] TCEQ NSR pending (run 10/8/2026) {NSR} (acc. Oct 8)",
  f"[12] Gov. Abbott (2026-09-21) {GOV} ; {CLAW} (acc. Oct 8)",
  "[13] Childress County tax abatements https://www.childresstx.us/pages/tax-abatements ; sites/docs/childress_county_abatement_notice_2023-08.pdf (acc. Oct 5)",
  "[14] Sentinel-2 L2A S2A_14SLD_20261005 (Earth Search); sites/imagery/childress_tx/detail/childress_tx_2026-10-05_crop_oct8.png",
  "[15] peers/comparison.xlsx Contracts_$MWyr rows P1 / NEW1 (comp band $9.7–16.7M)"])

SITES["crusoe_abilene"] = dict(
 title="Appendix: Crusoe Abilene, TX (Stargate; Oracle / OpenAI tenant) site status, as of " + ASOF,
 sub="Nebius Peer KPIs. Public sources only. Grade: Green (B1–6 delivered) / Yellow (B7–8 by YE26; expansion permit frozen). Oracle's Abilene MW is counted here only, not under Oracle.",
 facts=[["Land / developer", "Lancium (Clean Campus land, substations); Crusoe developer/builder via financing JV (Blue Owl per team files)", "R; P", "1, 14"],
        ["Tenant / end user", "Oracle (OCI) leases all 8 buildings; OpenAI is the end user (Stargate); NVIDIA GB200 NVL72", "R", "1, 2"],
        ["Location", "Longhorn Data Center, 615 FM 2404, Abilene, Taylor County (32.508, −99.777). Adjacent Crusoe–Microsoft 900 MW campus (2 × 336 IT, 2027+) excluded", "R", "7, 14"]],
 mw=[["Announced", "1.2 GW gross / 8 bldgs; ~824 IT implied", "618 ÷ 0.75; Epoch ~843 IT", "R; D; P", "1, 2, 3"],
     ["Connected (grid / utility)", "200 (138 kV) + 1,000 (345 kV) subs; ~360 on-site gas", "grid / on-site, not IT", "R", "4, 5"],
     ["Connected (halls)", "618 delivered (6 of 8 bldgs, 75%)", "basis unstated; read as IT", "R; basis E", "2"],
     ["Active", "618 (421–618)", "IT; delivered + accepted (R) read as active; low = Epoch to Jul-27", "D on R; P", "2, 3"],
     ["Billed", "~618 (Oracle → OpenAI)", "acceptance 'only 24 hours'", "D", "2"],
     ["YE26 billed L / B / H", "618 / 721 / 824", "IT; B7–8 ≈ 103 each", "E", "2, 3"]],
 mw_note=("CHANGED vs Oct 5 (active 200 / 421 / 421): Oracle's Sep 10 call (FQ1 FY27) says 6 of 8 buildings, 618 MW and 75% of capacity are delivered to the customer, with "
          "131k GPUs delivered there in the quarter [2]. The JV's lease billing to Oracle is n/d."),
 chart=dict(ann=824, grid=1200, grid_lbl="1,200 grid subs\n(+~360 on-site)", conn=618, act=(421, 618, 618), bill=618, tags=("D", "R", "R", "D", "D")),
 power=[["Utility / ISO", "ERCOT; AEP interconnection via Lancium 138 kV and 345 kV substations", "R", "4"],
        ["345 kV substation", "1st transformer 2025-12-19; all 5 MPTs by 2026-03-10 (3 temporary); permanent swaps targeted Oct 2026", "R", "4"],
        ["On-site generation", "Std Permit 177263: 5 Solar Titan 350 + 5 GE LM2500 ≈ 360.5 MW, 8,760 h/yr, on-site use only; ~62 diesel gensets (PBR, reported 177262). Crusoe: turbines are backup", "R; P", "5, 6"]],
 permits=[["177263 (issued)", "Abilene DC 1 LLC dba Crusoe; modification reviewed Jan 2025", "R", "5"],
          ["182126 / PSDTX1688 / GHGPSDTX263", "Crusoe Energy Systems, Longhorn DC (RN112029061): filed 2025-11-07; NORI 11-19; NAPD + draft 2026-08-05; TCEQ list: 'Technically Complete' (data 8/28, run 10/8). NOT issued", "R", "7, 8"],
          ["Scope of 182126", "Bldgs 9–10 / added on-site generation: 39 simple-cycle turbines, up to ~1,330 MW (EIP) vs 41 turbines + 18 gensets (Tribune)", "R; P", "7, 9, 6"],
          ["Sep 21 state pause", "No TCEQ data-center permits until ERCOT audit; TCEQ compliance report due Oct 19; ERCOT audit reports Dec 10, PUCT Dec 17", "R; P", "10, 11"],
          ["Incentives", "City 85% / Taylor County 80% 10-yr abatements, one agreement per building", "P", "12"]],
 evidence=[["Sentinel-2 Oct 5 v Sep 15", "All 8 halls roofed with chiller yards; large new works on the southern (Microsoft) campus. Beige roofs read low on the white-roof metric (14.9 v 10.9 ha); energized halls not visible at 10 m", "D; read E", "13"],
           ["Thermal", "Landsat Jun–Aug anomaly +0.40 → −0.41 °C: not informative (red soil)", "E", "14"],
           ["Contract economics", "Oracle–OpenAI ~$300B / 5 yrs (WSJ; not in Oracle filings) = $13.3M per stated MW-yr, $16.0M per IT MW-yr if 4.5 GW is gross: inside the $9.7–16.7M band. Site lease n/d. JV shell + power ≈ $12.5M per gross MW", "P; D", "15, 16"]],
 what=("(1) TCEQ report to the Governor Oct 19: as written the pause freezes 182126 (Bldgs 9–10, ~1.3 GW gas) only; extension to modifications or turbine operation would hit B1–8 backup (bear) [10, 11]. "
       "(2) Permanent 345 kV transformer swaps 'by Oct 2026' and Oracle FQ2 FY27 (~mid-Dec, E): 8 of 8 delivered (~824) is the bull; Epoch projected 843 IT ~Nov 1 [3, 4]. "
       "(3) ERCOT audit Dec 10 / PUCT Dec 17; Legislature from Jan 2027 on incentives [11]. Biggest risk: growth past 8 buildings is frozen until the audit; the 618 MW in service is not."),
 changed="Since Oct 5: active base 421 → 618 MW on Oracle's Sep 10 statement (in the Oracle pack, not previously in the site view); 182126 confirmed not issued (Technically Complete); EIP turbine count; ERCOT Dec 10/17 dates.",
 sources=[
  "[1] Crusoe PR (2025-09-30) https://www.crusoe.ai/resources/newsroom/crusoe-announces-flagship-abilene-data-center-is-live (acc. Oct 8); 1.2 GW / 8 bldgs per Crusoe JV PR 2025-05-21 (sites.xlsx)",
  f"[2] Oracle FQ1 FY27 call (2026-09-10; transcript 09-11) {ORCL} (acc. Oct 8)",
  "[3] Epoch AI https://epoch.ai/data/ai-data-centers/directory/openai-stargate-abilene (acc. Oct 5)",
  "[4] Mortenson https://www.mortenson.com/projects/abilene-data-center-development (acc. Oct 5)",
  "[5] TCEQ Std Permit 177263 tech review: sites/docs/abilene_tceq_177263_tech_review.pdf",
  "[6] Texas Tribune / Floodlight (2026-07-09) https://www.texastribune.org/2026/07/09/texas-data-centers-ai-power-plants-pollution-state-permits/ (acc. Oct 5)",
  "[7] TCEQ 182126 plain-language summary https://www.tceq.texas.gov/downloads/permitting/air/bilingual/pending-permit-notices/182126-pls-english.pdf ; sites/docs/abilene_tceq_182126_napd.pdf, _nori.pdf (acc. Oct 8)",
  f"[8] TCEQ NSR pending (run 10/8/2026) {NSR} (acc. Oct 8)",
  "[9] EIP Oil & Gas Watch https://oilandgaswatch.org/alert/fc44e871-e7e9-4d4e-af22-bce4133e16d3 (acc. Oct 8)",
  f"[10] Gov. Abbott (2026-09-21) {GOV} (acc. Oct 8)",
  f"[11] Texas Lawbook https://texaslawbook.net/abbott-orders-statewide-pause-on-texas-data-center-permits/ ; {CLAW} (acc. Oct 8)",
  "[12] DCOA (2025-05-01) https://developabilene.com/2025/05/01/project-radiance-powering-the-future-of-ai-in-abilene-texas/ (acc. Oct 5)",
  "[13] Sentinel-2 L2A S2A_14SMA_20261005 (Earth Search); sites/imagery/abilene_tx/detail/abilene_tx_2026-10-05_crop_oct8.png",
  "[14] Team: peers/sites/site_verification.md §4; sites.xlsx; strips/abilene_tx_strip.png; thermal/abilene_tx_landsat_st.csv; peers/oracle/oracle_kpis.xlsx Inputs Q40 (Blue Owl)",
  "[15] peers/comparison.xlsx Contracts_$MWyr O2 (WSJ) / NEW8",
  "[16] peers/private/private_peers.md (Crusoe K14, D)"])

SITES["oracle_port_washington"] = dict(
 title="Appendix: Oracle (OCI) Port Washington, WI (Vantage 'Lighthouse'; Stargate) site status, as of " + ASOF,
 sub=("Why this site: Oracle's largest live AI site is Abilene, the SAME campus as Crusoe Abilene (Oracle is the tenant), so it is counted once, under Crusoe. Port Washington is Oracle's "
      "separable key site. Shackelford County TX (Oracle: 'our next gigawatt scale campus', not delivered in FQ1) not built: thinner public record (gap). Grade: Yellow/Red vs end-2027 handover (unchanged)."),
 facts=[["Developer / owner", "Vantage Data Centers (Wisco Acquisition Co.); $15B; 4 bldgs, ~2.5M sf", "R; P", "1, 2"],
        ["Tenant / end user", "Oracle (OCI) for OpenAI (Stargate); turnover in phases from end-2027", "P", "2"],
        ["Location", "North Port Washington, Ozaukee County WI; 672 ac (later phases up to 3.5 GW, press)", "P", "1, 3"]],
 mw=[["Announced", "~902 IT / 1.3 GW electrical", "Vantage IT v utility total", "R; P", "1, 4"],
     ["Connected (grid / utility)", "0", "~310 requested for 2027; 438 peak 2027", "P", "4"],
     ["Connected (halls)", "0", "20–25% complete; 2 of 4 halls enclosed", "R; P", "2, 5"],
     ["Active", "0", "IT", "R", "6"],
     ["Billed", "0", "IT", "R", "6"],
     ["YE26 / YE27 L / B / H", "0 / 0 / 0; YE27 0 / ~310 / 438", "YE27 = grid MW", "E", "4, 5"]],
 mw_note=("Oracle (Sep 10): 'neither Shackleford nor New Mexico or Wisconsin or Michigan were delivered in Q1' [6]. Utility filings (via press) show Vantage asked ATC for ~310 MW in 2027 "
          "and 1.3 GW from 2028 [4]."),
 chart=dict(ann=902, grid=0, grid_lbl="0 grid\n(1.3 GW requested)", conn=0, act=(0, 0, 0), bill=0, tags=("R", "P", "R", "R", "R")),
 power=[["Utility", "We Energies (very-large-customer tariff; PSC 6630-TE-113 collateral rules; Oracle dropped its court challenge Aug 17)", "R; P", "3, 7"],
        ["Transmission", "ATC Ozaukee County project $1.39–1.67B, new Decker substation; construction Dec 2026 – Dec 2027", "R", "8"],
        ["PSC status", "Completeness revoked Aug 2026; case closed Sep 10; ATC refiled (Sep 14 per Wisconsin Watch, Sep 18 per TechRadar); completeness decision expected ~Oct 19", "P", "4, 5, 9"],
        ["Aterio timing", "Partial power Oct 2027 earliest / Dec 2027 base / Jun 2028 extended; full 1.3 GW Aug 2028 / Oct 2028 / Apr 2029", "P", "5"],
        ["On-site generation", "Up to 45 emergency diesel engines, 87 MW: backup only", "R", "10"]],
 permits=[["WDNR air", "Construction 25-MIN-094 + operation 24618528A-F01, final 2026-06-12", "R", "10"],
          ["PSC (ATC line)", "Refiled; completeness pending (~Oct 19)", "R; P", "8, 9"],
          ["City", "TIF $458.6M", "P", "1"]],
 evidence=[["Sentinel-2 Oct 6 v Sep 11", "More roofs closed: white-roof 28.9 v 22.9 ha (+6 ha, same metric); still a construction site; no energized power plant visible at 10 m", "D; read E", "11"],
           ["Press photos", "Ozaukee Press aerial 2026-09-21: first 555k sf building enclosed; Sep 18 tour; 2,800–3,000 workers/day", "P", "2"],
           ["Contract economics", "Oracle–OpenAI ~$300B (WSJ) program-level, $13.3–16.0M per MW-yr; site lease n/d", "P; D", "12"]],
 what=("(1) PSC completeness decision ~Oct 19 (with ATC's MISO voltage-device filing): a clean finding keeps Aterio's base (partial power ~Dec 2027); another rejection moves to the "
       "extended case (Jun 2028) [5]. (2) Oracle FQ2 FY27 (~mid-Dec, E): any change to 'H2 2027' Wisconsin delivery [6]. (3) Watch for a temporary / bypass supply for the ~310 MW 2027 "
       "request; Vantage says portions go live while the rest is built [2, 4]. Biggest risk: grid, not construction: halls can be ready before power (connected 0 into late 2027)."),
 changed="Since Oct 5: PSC completeness ~Oct 19 trigger, ATC refile-date conflict and the ~310 MW 2027 request added; Oracle confirms no Wisconsin delivery in FQ1; Oct 6 Sentinel-2 +6 ha roof.",
 sources=[
  "[1] Ozaukee Press (2025-10-22) https://www.ozaukeepress.com/content/openai-oracle-run-stargate-port-data-center (acc. Oct 5)",
  "[2] Ozaukee Press (2026-09-23) https://www.ozaukeepress.com/content/data-center-size-%E2%80%98-insane%E2%80%99 (acc. Oct 8)",
  "[3] Data Center Census https://datacentercensus.com/dc/port-washington-lighthouse-wi (acc. Oct 8)",
  "[4] TechRadar (2026-10-04) https://www.techradar.com/pro/a-meaningful-risk-oracles-massive-1-3-gw-wisconsin-ai-campus-faces-severe-power-delays-as-grid-review-restarts-pushing-customer-delivery-past-2027 (acc. Oct 8)",
  "[5] Aterio (BusinessWire 2026-09-30) https://www.streetinsider.com/Business+Wire/Aterio+Flags+Delay+Risk+for+Oracle%E2%80%99s+1.3+GW+Wisconsin+Data+Center/27125698.html ; https://constructionreviewonline.com/project-lighthouse-delays-threaten-oracles-2027-wisconsin-data-center-launch/ (acc. Oct 8)",
  f"[6] Oracle FQ1 FY27 call (2026-09-10) {ORCL} (acc. Oct 8)",
  "[7] PSC doc 593736 https://apps.psc.wi.gov/ERF/ERFview/viewdoc.aspx?docid=593736 (acc. Oct 5)",
  "[8] PSC doc 597276 https://apps.psc.wi.gov/ERF/ERFview/viewdoc.aspx?docid=597276 (acc. Oct 5)",
  "[9] Wisconsin Watch https://wisconsinwatch.org/2026/09/wisconsin-port-washington-data-center-atc-restarts-push-grid-upgrades-transmission-utility/ (acc. Oct 5)",
  "[10] WDNR 25-MIN-094 final https://www.wpr.org/wp-content/uploads/2026/07/dnr.state_.wi_.us_Apps_AM_WARP_Documents_PROD_Final_63297.pdf ; sites/docs/wi_dnr_vantage_25-MIN-094_final.pdf (acc. Oct 5)",
  "[11] Sentinel-2 L2A S2B_16TDP_20261006 (Earth Search); sites/imagery/portwash_wi/detail/portwash_wi_2026-10-06_crop_oct8.png",
  "[12] peers/comparison.xlsx Contracts_$MWyr O2; team peers/sites/site_verification.md §6; sites.xlsx"])
