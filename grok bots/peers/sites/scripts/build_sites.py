#!/usr/bin/env python3
"""Builds /workspace/nebius/peers/sites/sites.xlsx. Same structure/method as
/workspace/nebius/capacity/altdata/scripts/build_timeline.py (Timeline rows: site, date, evidence type,
observation, source, confidence) plus a one-row-per-site summary (Sites) for peer comparison.
MW basis: IT (critical) MW unless stated. 'Active' = evidence-supported energized & in service as of 2026-10-04 ET."""
import csv, glob, json, os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
OUT='/workspace/nebius/peers/sites'
S2='Sentinel-2 L2A (Element84 Earth Search STAC);'
def img(site,date): return f'{S2} imagery/{site}/detail/{site}_{date}_crop.png'

# ------------------------------------------------------------------ SITES (one row per peer)
SITES=[
dict(peer='CoreWeave', site='Denton TX (Core Scientific campus, 8171 Jim Christal Rd)', key='denton_tx', basis='IT (critical) MW',
 announced='~262 MW IT (394 MW grid) at full build', contracted='~260 MW IT leased to CoreWeave (CORZ Q2-26 deck)',
 act_L=250, act_B=260, act_H=262, ye_L=260, ye_B=262, ye_H=262,
 active_basis='Landlord-reported billing: Bldgs D-G ~132 MW billing Mar-2026 (8-K Ex99.2 2026-04-21); ~260 MW leased & ~260 MW billing, "substantially complete" (Q2-26 deck 2026-07-28). Low = rounding haircut on "~260".',
 energization='Pre-existing 297 MW grid load energized before 3/25/2022 (ERCOT "Base Load Pathway (a)", CORZ 8-K 2026-09-10); initial HPC billing H2-2025; +74 MW in ERCOT 2025 RTP still to energize',
 power='Denton Municipal Electric PPA (~394 MW potential); ERCOT: existing 297 MW conditionally approved Base Load Pathway (a) (energized pre-3/25/2022), outside Batch Zero; +74 MW validated in 2025 RTP, also outside Batch Zero (CORZ 8-K 2026-09-10). CORZ funds 2 new on-site substations; DME/ERCOT transmission upgrades by 2029 (City of Denton 2024-11-20)',
 permits='City of Denton lease & PPA amendments approved 2024-11-19 (78.85 ac); no city tax incentives (state incentives applied for). TCEQ PBR Registration 179143 acknowledged 2025-02-11 under 30 TAC §106.511 (emergency/standby engines) for Core Scientific Data Center Site 1 And 2 Denton (RN111355798 / CN605945302; docs/denton_tceq_pbr_179143.pdf). No major NSR air permit found in this pass (grid-powered campus). T5 Construction GC',
 imagery='S2 2023-02-03..2026-09-15 (25 scenes). White-roof area in crop 6.8-9.8 ha (2023) -> 12.4 (Jul-25) -> 18.6-18.7 ha (Jun/Sep-26); new south pad from Oct-25',
 thermal='Landsat: see Thermal tab (pre-existing mining load => cannot isolate AI conversion)',
 grade='Green', confidence='High (landlord SEC filings; imagery consistent)',
 key_images='strips/denton_tx_strip.png; imagery/denton_tx/detail/denton_tx_2023-06-28_crop.png; imagery/denton_tx/detail/denton_tx_2026-09-15_crop.png'),
dict(peer='IREN', site='Childress TX (620 FM 1033)', key='childress_tx', basis='IT MW for AI; grid MW for site',
 announced='750 MW grid site; Horizons 1-4 = 4 x 50 MW IT liquid-cooled (+H5-6 2027); 60 MW IT air-cooled for NVIDIA contract (from early 2027)',
 contracted='200 MW IT Microsoft ($9.7B/5 yr) = Horizons 1-4; NVIDIA $3.4B within 60 MW of existing halls (ramp from early 2027)',
 act_L=50, act_B=75, act_H=100, ye_L=100, ye_B=150, ye_H=200,
 active_basis='Horizon 1 (50 MW IT) delivered/accepted by Microsoft 2026-08-13 (SEC Ex99.1); Horizon 2 commissioning underway (FY26 results 2026-08-27) => 50-100 MW IT AI. Legacy bitcoin load (650 MW "operating data centers" Aug-2025) separate and being wound down.',
 energization='Site energized for mining since 2022-24 (Phases 1-5 = 650 MW, IREN Aug-2025 update); second bulk substation transformer energized for Horizon 1-2 (Aug-2025 update); grid raised 600->750 MW 2024-07-24 (AEP/ERCOT)',
 power='AEP/ERCOT 750 MW interconnection (amended connection agreement after ERCOT network studies; IREN PR 2024-07-24); IREN-owned on-site substations with dual feed; second bulk transformer energized for Horizon 1-2 (Aug-2025 update). Site already energized (mining Phases 1-5); not a Batch Zero new-interconnection story (contrast Sweetwater). Separate AEP Childress Area Transmission Improvement Project (ERCOT Tier 2, ~$77.3M, ISD Nov-2031) addresses county voltage/thermal reliability from load growth — regional upgrade, not the site\'s 750 MW interconnection itself (PUC 55999 Item 173 / docs/ercot_aep_childress_tier2_endorsement.pdf)',
 permits='TCEQ Title V Operating Permit application #4951 (Ie US Development Holdings 3 Inc. — IREN SPV; Texas Comptroller DC list LD926028), Area Name Childress Data Center, received 2026-08-17, Administrative Complete (tceq.texas.gov Title V pending list, data updated 2026-10-05). Title V = operating consolidation, not a pending major NSR construction permit. No public NSR major / EGU standard-permit registration for on-site turbines found this pass (grid-connected campus). Childress County Chapter 312 tax abatements + Iris Reinvestment Zone for IE US Development Holdings 3 / IE US Hardware 1-4 (county notice Aug-2023; ~$1.4B improvements contemplated; amended/restated filings on childresstx.us/pages/tax-abatements). TDLR TABS2024024946 / TABS2026003980 confirm same SPV at 620 FM 1033. Search scope for air: TCEQ Title V pending list; public web/TCEQ records search for Childress Data Center / IE US / IREN PBR & standard permits — only Title V surfaced.',
 imagery='S2 2023-03-03..2026-09-28 (25 scenes). White-roof 0.5 ha (Mar-23) -> 10.8 (Feb-25) -> 17-18 (Jul/Aug-25) -> 19-25 ha (Mar-Jul-26); new west halls Oct-25..Mar-26; new grey facility SE + north halls by Sep-26',
 thermal='Landsat: see Thermal tab (mining load dominates)',
 grade='Green (H1-2) / Yellow (H3-4 by Q4-26)', confidence='High for 50 MW; Medium for 75-100 MW',
 key_images='strips/childress_tx_strip.png; imagery/childress_tx/detail/childress_tx_2024-07-30_crop.png; imagery/childress_tx/detail/childress_tx_2026-09-28_crop.png'),
dict(peer='Nscale', site='Kvandal (Narvik/Bjerkvik) NO - Stargate Norway', key='kvandal_no', basis='MW (grid allocation / IT not split)',
 announced='230 MW grid allocated (Statnett) + up to 290 MW more (paused); HENT KNB4-6 = 100 MW buildings',
 contracted='S-1: Norway owned sites 100 + 100 MW contracted (not active); Vattenfall PPA phase 1 2027-31',
 act_L=0, act_B=0, act_H=0, ye_L=0, ye_B=0, ye_H=0,
 active_basis='Under construction: groundbreaking 2025-11-19 (NRK); HENT 2026-09-02 completion "during 2027"; Vattenfall PPA operations from 2027; imagery shows first shells only from Jun-2026.',
 energization='None evidenced. Nordkraft/Statnett grid allocation 230 MW; Statnett paused new allocations in N. Norway Apr-2026 (phase 2 uncertain)',
 power='Statnett 230 MW allocated; Vattenfall PPA 2027-2031',
 permits='Kvandal North regulated to industry Apr-2024 (345.no); NVE notice 2026-09-29: construction started without required waste-heat analysis, deadline 2026-11-15 or coercive fines; Eksfin NOK 2.2B loan; Kvandal South DC Facility $725M + $65M VAT (S-1)',
 imagery='S2 2023-06-26..2026-10-03 (13 snow-free scenes). Graded pads, no buildings 2023-Oct-25; building shells on SW pad and NE pad from 2026-06-10; long halls by 2026-09-11',
 thermal='Not run (pre-operational; high latitude/snow)',
 grade='Yellow (2027 start; NVE/grid risks)', confidence='High that active = 0',
 key_images='strips/kvandal_no_strip.png; imagery/kvandal_no/detail/kvandal_no_2025-10-04_crop.png; imagery/kvandal_no/detail/kvandal_no_2026-09-11_crop.png'),
dict(peer='Crusoe', site='Abilene TX (Lancium Clean Campus; Stargate - Oracle/OpenAI)', key='abilene_tx', basis='IT MW',
 announced='1.2 GW gross / 8 buildings (Crusoe JV PR 2025-05-21); ~824 MW IT implied (618 / 0.75, D; Epoch ~843 IT, P). Microsoft 900 MW (2 x 336 MW IT) adjacent campus 2027+ excluded',
 contracted='All 8 buildings leased to Oracle (for OpenAI); 600 MW expansion reportedly not proceeding (Mar-2026, secondary)',
 act_L=421, act_B=618, act_H=618, ye_L=618, ye_B=721, ye_H=824,  # Oct 8: was 200/421/421, YE26 421/632/843 (Oct 5 history)
 active_basis='[D on R] Oracle FQ1 FY27 call (2026-09-10): 6 of 8 buildings, 618 MW and 75% of capacity delivered to the customer; 131k GPUs delivered there in the quarter. Delivered + accepted (R) read as active (D). IT vs gross basis unstated (read as IT, E). Low 421 = Epoch AI satellite model (Bldgs 1-4, ~2026-05-23; P). Billed ~618 (Oracle -> OpenAI; D); JV lease billing to Oracle n/d. History: Oct 5 view 200/421/421 (Oracle Apr-2026 "200MW is already operational"; Epoch 421).',
 energization='138 kV 200 MW substation (built May-2022, 3rd bay H2-2024; AEP interconnection / slack span to Abilene NW). 1 GW 345 kV greenfield substation (12.5 ac): first MPT energized 2025-12-19; all 5 MPTs by 2026-03-10 (3 temporary); permanent transformer swaps targeted by Oct-2026 (Mortenson). On-site generation: 10 gas turbines (~360.5 MW) under Std Permit 177263 + ~62 diesel gensets under minor PBR (Texas Tribune 2026-07-09)',
 power='ERCOT grid via Lancium 138 kV (200 MW) + 345 kV (1 GW) substations (Mortenson EPC; AEP interconnection). Behind-the-meter: Standard Permit Reg. 177263 — 5× Solar Titan 350 (38 MW) + 5× GE LM2500 (34.1 MW) = ~360.5 MW, 8,760 hrs/yr authorized, onsite use only, no grid export (TCEQ tech review Jan-2025; RN112029079 / CN606295137; docs/abilene_tceq_177263_tech_review.pdf). Crusoe has stated turbines are backup (Tribune)',
 permits='Minor: Std Permit Reg. 177263 (originally 2024; mod approved ~2025-01-22) for 10 turbines; diesel backup gensets under PBR (reported Reg. 177262 / §106.511; Tribune: 62 units). Major pending: State AQ Permit 182126 + PSDTX1688 + GHGPSDTX263 (Crusoe Technologies / Crusoe Energy Systems; Longhorn Data Center RN112029061 @ 615 FM 2404); filed 2025-11-07; NORI 2025-11-19; NAPD + draft permit 2026-08-05 (docs/abilene_tceq_182126_*.pdf). Unit counts for expansion (41 turbines + 18 gensets) per Texas Tribune/Floodlight 2026-07-09 — NAPD does not enumerate units in the public notice. City of Abilene 85% / Taylor County 80% 10-yr abatements per building. Note: Gov. Abbott 2026 directive to halt new data-center TCEQ permits pending ERCOT/TWDB audits (update due 2026-10-19) is a process risk for the pending major, not for already-issued 177263.',
 imagery='S2 2023-01-22..2026-09-15 (25 scenes). Pre-construction Jul-23; N pair roofed + middle row pads Apr-25; 6 roofs Jul-25; chillers visible Oct-25; all 8 roofed Mar-26; new southern earthworks (Microsoft campus) Jun-Sep-26',
 thermal='Landsat: see Thermal tab (daytime; bare red soil confounds)',
 grade='Green (B1-6 delivered) / Yellow (B7-8 by YE26; expansion permit 182126 frozen)', confidence='Medium-High (618 MW company-stated delivered, R; active read D on R; IT vs gross basis unstated)',
 key_images='strips/abilene_tx_strip.png; imagery/abilene_tx/detail/abilene_tx_2025-04-10_crop.png; imagery/abilene_tx/detail/abilene_tx_2026-09-15_crop.png'),
dict(peer='Lambda', site='Kansas City MO (Project Blitz, 11155 N Airworld Dr)', key='kcairworld_mo', basis='IT MW',
 announced='24 MW early 2026; >100 MW potential (Lambda blog 2025-10-28)', contracted='Single customer (unnamed); >10,000 Blackwell Ultra GPUs',
 act_L=0, act_B=12, act_H=24, ye_L=12, ye_B=24, ye_H=24,
 active_basis='No public confirmation of operation. MO DNR letter 2026-04-15: commissioning from 2026-06-01 on 5 x 3 MW rental engines (prime, Jun-Oct) + up to 6 x 3 MW; Evergy full supply late summer 2026. Evergy (KSHB, Jun-2026): enough power "available today" to fully operate. Base = half of 24 MW.',
 energization='Partial Evergy service expected Jun-2026, full late summer 2026 (MO DNR 2026-04-15); not confirmed',
 power='Evergy; temporary rental gensets (classed nonroad, no permit)',
 permits='MO DNR de minimis construction permit 102025-010 issued 2025-10-29 (site ID 165-0051); MO DNR 2026-03-061 applicability letter 2026-04-15; Port KC bond application withdrawn Jun-2026; city building permits active (KC Star)',
 imagery='S2 2023-02-03..2026-10-02 (26 scenes). Existing ex-BofA building; white-roof 1.6-1.9 ha (2023-25) -> 2.1-2.5 ha (Apr-Oct-26): small rooftop/yard changes only (interior retrofit; 10 m too coarse)',
 thermal='Not run (crop ~5 x 4 Landsat TIRS pixels; meaningless)',
 grade='Yellow (no go-live confirmation)', confidence='Low',
 key_images='strips/kcairworld_mo_strip.png; imagery/kcairworld_mo/detail/kcairworld_mo_2025-10-20_crop.png; imagery/kcairworld_mo/detail/kcairworld_mo_2026-10-02_crop.png'),
dict(peer='Oracle (OCI)', site='Port Washington WI (Vantage "Lighthouse"; Stargate)', key='portwash_wi', basis='IT MW (~902 MW IT of 1.3 GW electrical)',
 announced='~1.3 GW electrical / ~902 MW IT, 4 buildings, 672 ac, $15B (Ozaukee Press 2025-10-22)', contracted='Leased to Oracle for OpenAI (Stargate); handover from end-2027',
 act_L=0, act_B=0, act_H=0, ye_L=0, ye_B=0, ye_H=0,
 active_basis='No grid power: ATC Ozaukee County interconnection (new Decker substation) construction Dec-2026..Dec-2027; PSC completeness revoked Aug-2026, refiled 2026-09-14/15; Aterio: partial power earliest Oct-2027, base Dec-2027.',
 energization='None; earliest partial Oct-2027 (Aterio, third party)',
 power='We Energies VLC tariff / Oracle service agreement (PSC doc 593736); ATC $1.39-1.67B project (PSC doc 597276)',
 permits='WDNR air construction permit 25-MIN-094 + operation permit 24618528A-F01 final 2026-06-12: up to 45 emergency diesel engines (15x1.5 + 18x1.75 + 12x2.75 MW = 87 MW backup); city TIF $458.6M',
 imagery='S2 2023-02-12..2026-09-11 (23 scenes). Farmland to Oct-25; mass earthworks Mar-26; large white roofs ~21.6 ha Jun-26 -> ~22.9 ha Sep-26 + dark roofs/foundations; 20-25% complete (Ozaukee Press 2026-09-23)',
 thermal='Landsat run as construction control (no load) - see Thermal tab',
 grade='Yellow/Red (end-2027 handover vs grid timeline)', confidence='High that active = 0',
 key_images='strips/portwash_wi_strip.png; imagery/portwash_wi/detail/portwash_wi_2025-10-04_crop.png; imagery/portwash_wi/detail/portwash_wi_2026-09-11_crop.png'),
dict(peer='Nscale', site='Monarch Compute Campus, Mason County WV (Point Pleasant)', key='monarch_wv', basis='IT MW',
 announced="Phase 1: 2 GW gross / ~1.35-1.37 GW IT by H1'28 (S-1 / nscale.com); ~8 GW gross by 2031; $20B+ planned capex",
 contracted='460 MW IT in S-1 facilities table (US Owned, Tier 3, Contracted). Anthropic GPU Services Agreements up to $44.6B (signed 2026-08-25; Vera Rubin NVL72; four tranches). Financing for Anthropic NOT committed as of S-1 date. Implied ACV ~$16.2M/IT MW-yr if 6-yr term (press).',
 act_L=0, act_B=0, act_H=0, ye_L=0, ye_B=0, ye_H=0, h128_L=0, h128_B=460, h128_H=1370,
 active_basis='Site prep / civil only as of 2026-09-25 Sentinel-2: ~85-100 ha bare ground, graded pads, no completed hall roofs (white-roof still ~2 ha = pre-existing hangars). Site manager (2026-09-20): still civil, about to go vertical. Active IT = 0.',
 energization="None. Off-grid gas microgrid (no AEP/FirstEnergy/PJM interconnection for primary load). Air permit app plans startup ~2027-01-01; S-1 targets H1'28 for 2 GW / 1.37 GW IT. Air construction permit R14-0042 still DRAFT as of 2026-09-29 (public comment into late Oct 2026).",
 power='Behind-the-meter: 864 x 2.5 MW natural-gas RICE = 2.16 GW nameplate (WVDEP R14-0042); peak site demand ~1.8 GW (permit app). Caterpillar partnership. 14-mile gas pipeline (Putnam/Mason; stormwater permit). Closed-loop cooling (no municipal drinking water). Not on public grid.',
 permits='WVDEP air R14-0042 applied 2026-04-20, draft 2026-09-29 (not final). Stormwater enforcement order 2026-07-14. High Impact Data Center certification announced 2026-09-03 (Gov. Morrisey). $62.5M forgivable state loan to affiliate Mountaineer GigaSystem (press; $25M funded / $37.5M waived). Land: option on site at permit filing (app).',
 imagery='S2 2023-03-30..2026-09-25 (21 scenes). Pre-construction through Apr-2026 (airport runway intact). Bare-ha 5 -> 48 (May-28) -> 95-99 (Jun-Aug) -> 85 (Sep-25). White-roof flat ~1-2 ha. Graded pads / earthworks; no hall shells yet.',
 thermal='Landsat ST completed (construction control); Jun-Aug anomaly near 0 — not a power meter. thermal/monarch_wv_landsat_st.csv',
 grade="Yellow (460 MW H1'28 path) / Yellow-Red (full 1.37 GW H1'28)", confidence="High that active=0; Medium that 460 MW contract is real; Low-Medium that 460 MW is on-track for H1'28 (financing + draft permit + early construction)",
 key_images='strips/monarch_wv_strip.png; imagery/monarch_wv/detail/monarch_wv_2026-04-08_crop.png; imagery/monarch_wv/detail/monarch_wv_2026-06-02_crop.png; imagery/monarch_wv/detail/monarch_wv_2026-09-25_crop.png'),
]

# ------------------------------------------------------------------ TIMELINE evidence rows
T=[
# DENTON
('Denton TX','2023-06-28','Imagery','Pre-conversion Core Scientific bitcoin campus: rows of mining sheds; white-roof in crop ~9.8 ha',img('denton_tx','2023-06-28'),'high'),
('Denton TX','2024-11-19','Municipal','Denton City Council approves PPA + land-lease amendments (78.85 ac); CORZ pays interconnection incl. 2 new on-site substations; no city tax incentives; DME/ERCOT transmission fixes by 2029','https://www.cityofdenton.com/CivicAlerts.aspx?AID=922','high'),
('Denton TX','2025-02-26','Company','CoreWeave +70 MW expansion -> critical IT just over 260 MW; CORZ funds ~$104M powered shell','https://www.businesswire.com/news/home/20250226866161/en/Core-Scientific-and-CoreWeave-Announce-%241.2-Billion-Expansion-at-Denton-TX-Site','high'),
('Denton TX','2025-07-22','Imagery','New dark equipment yard E of campus and new building segments; white-roof ~12.4 ha',img('denton_tx','2025-07-22'),'medium'),
('Denton TX','2025-10-10','Imagery','Large new L-shaped graded pad S of campus (expansion/laydown)',img('denton_tx','2025-10-10'),'medium'),
('Denton TX','2026-03 (as of)','Company (SEC)','Bldgs D-G (~132 MW) operational & billing (50% complete); 262 MW IT / 394 MW grid; 8 buildings A-H; 370,284 sf; PUE ~1.5; remainder H2-2026','https://investors.corescientific.com/sec-filings/all-sec-filings/content/0001193125-26-165121/d149019dex992.htm','high'),
('Denton TX','2026-06-07','Imagery','New white roofs S/E; white-roof ~18.6 ha',img('denton_tx','2026-06-07'),'medium'),
('Denton TX','2026-07-28','Company (SEC)','Denton ~260 MW leased, ~260 MW billing, "substantially complete"','https://www.sec.gov/Archives/edgar/data/1839341/000183934126000013/q2fy26earningsdeck728am.htm','high'),
('Denton TX','2026-09-10','ERCOT','ERCOT conditionally approves existing 297 MW as Base Load Pathway (a) (energized pre-3/25/2022), outside Batch Zero; +74 MW to be energized in 2025 RTP','https://investors.corescientific.com/sec-filings/all-sec-filings/content/0001839341-26-000023/0001839341-26-000023.pdf','high'),
('Denton TX','2026-09-15','Imagery','Campus fully roofed; white-roof ~18.7 ha (vs 9.8 ha 2023)',img('denton_tx','2026-09-15'),'medium'),
# CHILDRESS
('Childress TX','2023-06-16','Imagery','Early IREN site: substation yard + first mining building',img('childress_tx','2023-06-16'),'high'),
('Childress TX','2024-07-24','Utility','Grid capacity raised 600 -> 750 MW (AEP/ERCOT amendments)','https://iren.com/data-centers/childress','high'),
('Childress TX','2023-08 / ongoing','Tax abatement','Childress County Chapter 312 tax abatements + Iris Reinvestment Zone for IE US Development Holdings 3 Inc. / IE US Hardware 1-4; ~$1.4B improvements contemplated (county notice Aug-2023); amended/restated agreements posted on county tax-abatements page','https://www.childresstx.us/pages/tax-abatements ; docs/childress_county_abatement_notice_2023-08.pdf','high'),
('Childress TX','2025-11-05/10','Transmission (regional)','ERCOT endorses AEP Childress Area Transmission Improvement Project (Tier 2, ~$77.3M Option 1A, ISD Nov-2031) to fix Childress County voltage/thermal issues from load growth; CCN required — regional reliability, distinct from IREN site 750 MW interconnection','https://interchange.puc.texas.gov/Documents/55999_173_1556587.PDF ; docs/ercot_aep_childress_tier2_endorsement.pdf','high'),
('Childress TX','2026-08-17','Air permit','TCEQ Title V Operating Permit app #4951 by Ie US Development Holdings 3 Inc. for Childress Data Center; Administrative Complete (TCEQ pending list updated 2026-10-05). No pending major NSR construction permit found this pass','https://www.tceq.texas.gov/assets/public/permitting/air/reports/applications/titlev-pending-permits.html','high'),
('Childress TX','2024-07-30','Imagery','First large building block (mining halls); white-roof ~4.6 ha',img('childress_tx','2024-07-30'),'high'),
('Childress TX','2025-05-03','Imagery','Second large block complete; white-roof ~13.2 ha',img('childress_tx','2025-05-03'),'high'),
('Childress TX','2025-08','Company','650 MW operating data centers (Phases 1-5 = 100+100+150+150+150); second bulk substation transformer energized for Horizon 1 & 2 (100 MW IT)','https://iren.gcs-web.com/static-files/ec8ec2a6-a1d1-4bb6-9cb9-bc2d436b9db9','high'),
('Childress TX','2025-10-18','Imagery','New halls under construction on W side of campus (likely Horizon liquid-cooled halls; attribution unconfirmed)',img('childress_tx','2025-10-18'),'low'),
('Childress TX','2026-03-12','Imagery','W halls roofed; new pads N of campus',img('childress_tx','2026-03-12'),'medium'),
('Childress TX','2026-05-07','Company','Q3 FY26: $3.4B NVIDIA contract, air-cooled Blackwell within 60 MW of existing Childress halls, ramp from early 2027; mining being decommissioned','IREN Q3 FY26 results (2026-05-07)','high'),
('Childress TX','2026-08-13','Company (SEC)','Horizon 1 (50 MW IT, D2C liquid, GB300 NVL72) delivered and accepted by Microsoft','https://www.sec.gov/Archives/edgar/data/1878848/000114036126032638/ef20080141_ex99-1.htm','high'),
('Childress TX','2026-08-27','Company (SEC)','FY26: Horizon 2 commissioning; Horizons 3-4 late-stage construction targeting Q4 2026; MSFT 200 MW IT $9.7B','https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm','high'),
('Childress TX','2026-09-28','Imagery','New grey-roofed facility SE of campus and new halls N; white-roof 14.8 ha (scene haze; Jul ~19.7)',img('childress_tx','2026-09-28'),'medium'),
# KVANDAL
('Kvandal NO','2024-04','Municipal','Kvandal North regulated to industrial use; Nordkraft transformer station nearby','https://www.345.no/2024/04/ett-steg-naermere-datasenter-i-narvik-kommune/','high'),
('Kvandal NO','2025-10-04','Imagery','Graded industrial pads, no buildings',img('kvandal_no','2025-10-04'),'high'),
('Kvandal NO','2025-11-19','Company/Press','Groundbreaking (foundation stone) for data centre at Kvandal','https://www.nrk.no/nordland/legger-ned-grunnstein-for-datasenter-pa-kvandal-1.17658846','high'),
('Kvandal NO','2026-04/05','Grid','230 MW allocated by Statnett; Statnett paused new allocations in N. Norway -> phase 2 (+290 MW) uncertain; phase 1 32,000 m2, ~NOK 8B, full operation in about a year','https://www.nrk.no/nordland/nscale-sar-tvil-om-videre-satsing-pa-gigant-datasenter-i-narvik-1.17893773','high'),
('Kvandal NO','2026','PPA','Vattenfall PPA for phase 1, 2027-2031; operations start 2027','https://group.vattenfall.com/press-and-media/pressreleases/2026/vattenfall-and-nscale-partner-to-support-ai-infrastructure-growth-in-norway/','high'),
('Kvandal NO','2026-06-10','Imagery','First building shells on SW pad + building on NE pad',img('kvandal_no','2026-06-10'),'medium'),
('Kvandal NO','2026-07','Financing (SEC S-1)','Kvandal South DC Facility: $725M term loan + $65M VAT facility; matures 6 yrs after construction completion','/workspace/nebius/comps_raw/nscale_s1.htm.txt (Nscale S-1)','high'),
('Kvandal NO','2026-09-02','Contractor','HENT signs 2 more buildings (75 MW); HENT total KNB4-6 = 100 MW; completion during 2027','https://www.hent.no/2026/09/02/hent-har-signert-kontrakt-for-utbygging-av-datasenteret-til-nscale-i-narvik-regionen/','high'),
('Kvandal NO','2026-09-11','Imagery','Long hall shells on SW pad; no cooling yards resolvable',img('kvandal_no','2026-09-11'),'medium'),
('Kvandal NO','2026-09-29','Regulator','NVE: construction started without required waste-heat analysis; deadline 2026-11-15 or coercive fines','https://www.nrk.no/nordland/nscale-varsles-om-tvangsmulkt-1.18039577','high'),
# ABILENE
('Abilene TX','2022-05..2024 H2','Substation','138 kV 200 MW substation built from May 2022; 3rd bay added H2-2024','https://www.mortenson.com/projects/abilene-data-center-development','high'),
('Abilene TX','2023-07-18','Imagery','Pre-construction (substation pad only)',img('abilene_tx','2023-07-18'),'high'),
('Abilene TX','2024-08 / 2025-01-22','Air permit (minor)','TCEQ Standard Permit Reg. 177263 (Abilene DC 1 LLC dba Crusoe; RN112029079): 5× Titan 350 (38 MW) + 5× GE LM2500 (34.1 MW) = ~360.5 MW, 8760 hrs/yr, onsite only. Orig. registered ~2024-08; mod tech-review approved Jan-2025. Diesel backup gensets under separate PBR (Tribune: 62 units)','docs/abilene_tceq_177263_tech_review.pdf ; https://www.texastribune.org/2026/07/09/texas-data-centers-ai-power-plants-pollution-state-permits/','high'),
('Abilene TX','2025-11-07','Air permit (major pending)','Crusoe files State AQ Permit 182126 + PSDTX1688 + GHGPSDTX263 for Longhorn Data Center additional generation (615 FM 2404; RN112029061). NORI 2025-11-19; NAPD + draft permit 2026-08-05. Expansion unit counts (41 turbines + 18 gensets) per Tribune/Floodlight — not enumerated in NAPD text','docs/abilene_tceq_182126_napd.pdf ; docs/abilene_tceq_182126_nori.pdf ; docs/abilene_tceq_182126_pls.pdf','high'),

('Abilene TX','2024-06','Company','Construction began June 2024; first two buildings energized within a year','https://www.crusoe.ai/resources/newsroom/crusoe-announces-flagship-abilene-data-center-is-live','high'),
('Abilene TX','2025 (Apr)','Tax abatement','City of Abilene 85% / Taylor County 80% 10-yr abatements; separate agreement per building, starting as each completes','https://developabilene.com/2025/05/01/project-radiance-powering-the-future-of-ai-in-abilene-texas/ ; https://finance.yahoo.com/news/taylor-county-estimated-see-18-222915156.html','medium'),
('Abilene TX','2025-04-10','Imagery','North pair (Bldgs 1-2) roofed; middle row pads and partial roofs',img('abilene_tx','2025-04-10'),'high'),
('Abilene TX','2025-06','Company','Oracle began delivering first GB200 racks June 2025','https://www.crusoe.ai/resources/newsroom/crusoe-announces-flagship-abilene-data-center-is-live','high'),
('Abilene TX','2025-07-30','Imagery','Six buildings roofed (N pair + middle row of 4); W pair under construction',img('abilene_tx','2025-07-30'),'high'),
('Abilene TX','2025-09-30','Company','Crusoe: first phase "live" on OCI','https://www.crusoe.ai/resources/newsroom/crusoe-announces-flagship-abilene-data-center-is-live','high'),
('Abilene TX','2025-10-03','Imagery','Chiller/equipment rows (blue) visible on W buildings',img('abilene_tx','2025-10-03'),'medium'),
('Abilene TX','2025-12-19','Substation','1 GW 345 kV substation: first transformer energized; all 5 MPTs energized by 2026-03-10 (3 temporary; permanent swaps by Oct-2026)','https://www.mortenson.com/projects/abilene-data-center-development','high'),
('Abilene TX','2025-12-24','3rd-party satellite','Epoch AI: Building 2 fully operational -> 211 MW IT (TDLR filing date)','https://epoch.ai/data/ai-data-centers/directory/openai-stargate-abilene','medium'),
('Abilene TX','2026-03-26','Imagery','All 8 buildings roofed',img('abilene_tx','2026-03-26'),'high'),
('Abilene TX','2026-03 (late)','Thermal (3rd party)','Thermal image shows an on-site gas turbine operating','https://www.texastribune.org/2026/07/09/texas-data-centers-ai-power-plants-pollution-state-permits/','high'),
('Abilene TX','2026-04-22','Company','Oracle post: "In Abilene, 200MW is already operational"','https://epoch.ai/data/ai-data-centers/directory/openai-stargate-abilene (quotes Oracle)','high'),
('Abilene TX','2026-05-23','3rd-party satellite','Epoch AI estimate: Bldgs 3-4 operational -> 421 MW IT; 320 Airedale chillers','https://epoch.ai/data/ai-data-centers/directory/openai-stargate-abilene','medium'),
('Abilene TX','2026-07-28','3rd-party satellite','Epoch AI: Bldgs 5-8 structurally complete, not confirmed operational; projected 843 MW IT ~2026-11-01','https://epoch.ai/data/ai-data-centers/directory/openai-stargate-abilene','medium'),
('Abilene TX','2026-09-10','Company','Oracle FQ1 FY27 call: 6 of 8 buildings, 618 MW, 75% of capacity delivered to the customer; 131k GPUs delivered in the quarter (IT vs gross basis unstated). Replaces 421 MW (Epoch) as the active base','https://www.fool.com/earnings/call-transcripts/2026/09/11/oracle-orcl-q1-2027-earnings-call-transcript/','high'),
('Abilene TX','2026-09-15','Imagery','Campus complete externally; major new earthworks/structures to S (Microsoft 900 MW campus)',img('abilene_tx','2026-09-15'),'high'),
# LAMBDA KC
('Kansas City MO (Lambda)','2025-10-28','Company','Lambda: 24 MW early 2026, >100 MW potential, >10k Blackwell Ultra, single customer','https://lambda.ai/blog/lambda-to-build-a-100mw-ai-factory-in-kansas-city-mo','high'),
('Kansas City MO (Lambda)','2025-10-29','Air permit','MO DNR de minimis construction permit 102025-010 (Airworld Data Center)','https://dnr.mo.gov/air/business-industry/air-permits/airworld-data-center-kansas-city-102025-010','high'),
('Kansas City MO (Lambda)','2026-04-15','Regulator','MO DNR: Russell close to turnover; commissioning from 2026-06-01; Evergy partial power Jun 1, full late summer 2026; 5 x 3 MW rental prime engines Jun-Oct + up to 6 x 3 MW (nonroad, no permit)','https://dnr.mo.gov/sites/dnr/files/vfc/2026/04/main/airworld-blitz-data-center-kansas-city-2026-03-061.pdf','high'),
('Kansas City MO (Lambda)','2026-06','Incentive','Port KC bond application withdrawn; Evergy: site has enough power available today to fully operate; Lambda no comment','https://www.kshb.com/news/local-news/data-centers/northland-ai-data-center-withdraws-bond-application-with-port-kc','high'),
('Kansas City MO (Lambda)','2023-2026','Imagery','Existing building; only minor rooftop/yard brightening Apr-Oct 2026 (white-roof 1.6-1.9 -> 2.1-2.5 ha); retrofit not resolvable at 10 m',img('kcairworld_mo','2026-10-02'),'medium'),
('Kansas City MO (Lambda)','2026-10-04','Gap','No public confirmation found that Blitz is operational / billing','searches 2026-10-04 (Lambda blog, DCD, KC Star, KSHB)','medium'),
# PORT WASHINGTON
('Port Washington WI','2025-10-04','Imagery','Farmland (pre-construction)',img('portwash_wi','2025-10-04'),'high'),
('Port Washington WI','2025-10-22','Press','1.3 GW electrical / ~902 MW IT, 4 buildings, 672 ac, $15B, complete 2028; TIF $458.6M','https://www.ozaukeepress.com/content/openai-oracle-run-stargate-port-data-center','high'),
('Port Washington WI','2026-03-27','Imagery','Mass earthworks across campus',img('portwash_wi','2026-03-27'),'high'),
('Port Washington WI','2026-06-12','Air permit','WDNR final 25-MIN-094 / 24618528A-F01: up to 45 emergency diesel engines (87 MW total)','https://www.wpr.org/wp-content/uploads/2026/07/dnr.state_.wi_.us_Apps_AM_WARP_Documents_PROD_Final_63297.pdf ; docs/wi_dnr_vantage_25-MIN-094_final.pdf','high'),
('Port Washington WI','2026-06-23','Imagery','Large white building roofs; white-roof ~21.6 ha',img('portwash_wi','2026-06-23'),'high'),
('Port Washington WI','2026-08/09','Utility','PSC revoked completeness of ATC Ozaukee County interconnection (Aug); ATC refiled 2026-09-14/15; project $1.39-1.67B, construction Dec-2026..Dec-2027, new Decker substation','https://wisconsinwatch.org/2026/09/wisconsin-port-washington-data-center-atc-restarts-push-grid-upgrades-transmission-utility/ ; https://apps.psc.wi.gov/ERF/ERFview/viewdoc.aspx?docid=597276','high'),
('Port Washington WI','2026-09-11','Imagery','More roofs closed, dark roofing; white-roof ~22.9 ha',img('portwash_wi','2026-09-11'),'high'),
('Port Washington WI','2026-09-23','Press','20-25% complete; 2,800-3,000 workers; first building topped May 12; handover to Oracle/OpenAI from end-2027','https://www.ozaukeepress.com/content/data-center-size-%E2%80%98-insane%E2%80%99','high'),
('Port Washington WI','2026-10-01','3rd party','Aterio: 2 of 4 halls enclosed; partial power earliest Oct-2027, base Dec-2027, extended Jun-2028','https://constructionreviewonline.com/project-lighthouse-delays-threaten-oracles-2027-wisconsin-data-center-launch/','medium'),
# MONARCH WV
('Monarch WV','2026-03-31','Acquisition (SEC)','Nscale acquires AIPCorp / Monarch Compute Campus, Mason County WV: ~2,250 ac; initial 2 GW gross (1.37 GW IT) expected online H1 2028; ~8 GW gross planned 2031. Consideration ~$793M (cash + Series C shares)','Nscale S-1 (/workspace/nebius/comps_raw/nscale_s1.htm.txt); private_peers N25','high'),
('Monarch WV','2026-04-08','Imagery','Pre/early construction: Mason County Airport runway and hangars intact; forest/fields; white-roof ~1.3 ha',img('monarch_wv','2026-04-08'),'high'),
('Monarch WV','2026-04-20','Air permit','Monarch Cloud Campus LLC files WVDEP major-source construction permit R14-0042: 2.16 GW off-grid gas RICE (864 x 2.5 MW) + auxiliaries; 3 H-shaped halls / 12 power nodes; peak demand ~1.8 GW; site option (not deeded); planned startup ~2027-01-01','docs/monarch_r14_application.pdf ; https://ukozvhzgrcnkcatjqfoo.supabase.co/storage/v1/object/public/03-00-air-construction/2026.04.20_Monarch_Compute_Campus_Power_Plant_R14-0042_Application_c514f1ef','high'),
('Monarch WV','2026-05-28','Imagery','Mass grading begins: bare ground in crop jumps to ~48 ha (from ~5 ha Apr)',img('monarch_wv','2026-05-28'),'high'),
('Monarch WV','2026-06-02','Imagery','Campus-scale earthworks ~95 ha bare; graded pads west of runway; no hall roofs',img('monarch_wv','2026-06-02'),'high'),
('Monarch WV','2026-06','Utility (gas)','14-mile natural-gas pipeline proposed through Putnam and Mason counties to fuel the microgrid (WVDEP stormwater/general permit process)','https://www.wsaz.com/2026/06/16/pipeline-project-proposed-power-mason-county-data-center/','high'),
('Monarch WV','2026-07-14','Regulatory','WVDEP unilateral enforcement order: stormwater/mud left site into neighboring properties; stop land-disturbing work in one drainage area lacking controls','https://www.wvnews.com/news/wvnews/public-hearings-set-for-monarch-cloud-campus-air-permit-in-mason-county/article_e2d06911-0746-4a33-a9ba-aaa8e1bf01c0.html','high'),
('Monarch WV','2026-08-25','Contract (SEC)','Anthropic GPU Services Agreements (four tranches) for dedicated Vera Rubin NVL72 infrastructure at Monarch; aggregate payments up to ~$44.6B. Best-efforts financing covenant; as of S-1 date, no binding financing commitments','Nscale S-1; https://www.datacenterfrontier.com/machine-learning/news/55407627/anthropic-openai-keep-expanding-the-ai-data-center-map-and-the-financing-gets-harder','high'),
('Monarch WV','2026-08-31','Contracted MW (SEC)','S-1 facilities table lists United States Owned 460 MW Tier 3 Contracted (Monarch Anthropic slice). Full Phase 1 1.37 GW IT is announced, not all contracted — do not add 1.37 GW into the 1.37 GW company-wide active+contracted total','Nscale S-1 facilities table; private_peers N22/N25 note','high'),
('Monarch WV','2026-09-03','Incentive','Gov. Morrisey announces WV High Impact Data Center certification for Monarch; claims total client investment $69.2B (Anthropic + undisclosed $24.2B — MW for the add-on not disclosed)','https://wvpublic.org/story/energy-environment/second-certified-data-center-announced-for-mason-county/','medium'),
('Monarch WV','2026-09-20','Construction (press)','Site manager: civil work continuing, about to go vertical; ~900 workers on site + pipeline','https://wvmetronews.com/2026/09/20/monarch-compute-campus-to-go-vertical-soon/','high'),
('Monarch WV','2026-09-25','Imagery','Still grading/pads; bare ~85 ha; white-roof ~2.0 ha (hangars); no completed DC halls',img('monarch_wv','2026-09-25'),'high'),
('Monarch WV','2026-09-29','Air permit','WVDEP issues DRAFT R14-0042 (not final). Address: 30 Monarch Way, Ohio River Rd (WV-62) & Airport Rd; Lat/Lon 38.914258/-82.100935. Public hearings/comments into late Oct 2026','docs/monarch_r14_draft_permit.pdf ; https://www.publicnow.com/view/0138C564AC6E5357F14FFC0F7637030F6CAC29FB','high'),
('Monarch WV','n/a','Interconnection','No utility/ISO queue: campus is designed as a state-certified AI microgrid, fully behind-the-meter / off public grid (Nscale site; Caterpillar generation). Primary power risk is gas supply + air permit + engine delivery, not PJM','https://www.nscale.com/ai-infrastructure/monarch','high'),
]
NAME={'denton_tx':'Denton TX','childress_tx':'Childress TX','kvandal_no':'Kvandal NO','abilene_tx':'Abilene TX','kcairworld_mo':'Kansas City MO (Lambda)','portwash_wi':'Port Washington WI','monarch_wv':'Monarch WV'}
# thermal summaries appended as Timeline rows
for k,n in NAME.items():
    f=f'{OUT}/thermal/{k}_landsat_st.csv'
    if not os.path.exists(f): continue
    rows=list(csv.DictReader(open(f))); summ=[]
    for y in ('2023','2024','2025','2026'):
        v=[float(r['anomaly_C']) for r in rows if r['date'][:4]==y and r['date'][5:7] in ('06','07','08')]
        if v: summ.append(f"{y} {sum(v)/len(v):+.2f}C (n={len(v)})")
    T.append((n,'2023-2026','Thermal (Landsat)','Daytime ST campus-minus-3km-background, Jun-Aug mean: '+'; '.join(summ)+'. 100 m TIRS, ~10:30 overpass; albedo/bare soil confound; not a power meter',f'thermal/{k}_landsat_st.csv (Planetary Computer landsat-c2-l2)','low'))

wb=Workbook(); H=Font(bold=True,color='FFFFFF'); F=PatternFill('solid',fgColor='305496')
def sheet(ws,hdr,rows,widths):
    ws.append(hdr)
    for c in ws[1]: c.font=H; c.fill=F; c.alignment=Alignment(wrap_text=True,vertical='top')
    for r in rows: ws.append(list(r))
    for i,w in enumerate(widths): ws.column_dimensions[chr(65+i) if i<26 else 'A'+chr(65+i-26)].width=w
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
    ws.freeze_panes='C2'
ws=wb.active; ws.title='Sites'
hdr=['Peer','Site','MW basis','Announced MW','Contracted MW','Active low (MW)','Active base (MW)','Active high (MW)','Active basis (evidence)','YE26 low','YE26 base','YE26 high',"H1'28 low","H1'28 base","H1'28 high",'Energization-date evidence','Power / interconnection','Permits / filings / incentives','Imagery dates & observations','Thermal','Grade','Confidence','Key images']
sheet(ws,hdr,[[s['peer'],s['site'],s['basis'],s['announced'],s['contracted'],s['act_L'],s['act_B'],s['act_H'],s['active_basis'],s['ye_L'],s['ye_B'],s['ye_H'],s.get('h128_L',s['ye_L']),s.get('h128_B',s['ye_B']),s.get('h128_H',s['ye_H']),s['energization'],s['power'],s['permits'],s['imagery'],s['thermal'],s['grade'],s['confidence'],s['key_images']] for s in SITES],
      [12,30,14,34,34,9,9,9,60,8,8,8,8,8,8,50,40,55,55,30,20,26,50])
ws.append([]); ws.append(['Notes: Active = evidence-supported energized and in service as of 2026-10-04 ET (Abilene updated 2026-10-08 to Oracle\'s Sep 10 statement: 618 MW delivered, D on R; IT vs gross basis unstated; Oct 5 view 200/421/421 is history). Low/Base/High follow Nebius site_verification.md convention. YE26 = estimate (see site_verification.md). Grades: Green = strong evidence; Yellow = partial/slippage; Red = contradictory/no evidence.'])
ws2=wb.create_sheet('Timeline'); sheet(ws2,['Site','Date','Evidence type','Observation','Source / image path','Confidence'],T,[22,16,18,90,80,11])
# Imagery metrics
ws3=wb.create_sheet('Imagery_metrics'); rows=[]
for k in NAME:
    f=f'{OUT}/imagery/{k}/detail_metrics.csv'
    if os.path.exists(f):
        for r in csv.DictReader(open(f)): rows.append(list(r.values()))
        hdrm=list(csv.DictReader(open(f)).fieldnames)
sheet(ws3,hdrm,rows,[14]*len(hdrm))
ws4=wb.create_sheet('Imagery_AOI_metrics'); rows=[]; keys=['site','date','scene','aoi_cloud','valid_frac','ndvi_mean','ndbi_mean','bsi_mean','builtbare_ha','brightroof_ha','chip']
for k in NAME:
    m=json.load(open(f'{OUT}/imagery/{k}/metrics.json'))
    for r in m['rows']:
        if r.get('date'): rows.append([r.get(x) for x in keys])
sheet(ws4,keys,rows,[14,11,30,9,9,9,9,9,11,11,70])
ws5=wb.create_sheet('Thermal'); rows=[]
for k in NAME:
    f=f'{OUT}/thermal/{k}_landsat_st.csv'
    if os.path.exists(f): rows+= [list(r.values()) for r in csv.DictReader(open(f))]
sheet(ws5,['site','date','id','campus_C','background_C','anomaly_C'],rows,[14,11,40,10,12,10])
ws6=wb.create_sheet('Sites_AOI'); import sys; sys.path.insert(0,os.path.dirname(__file__)); from sites import SITES as AOI
crops=json.load(open(f'{OUT}/scripts/crops_extra.json'))
sheet(ws6,['site','lat','lon','half_m','cadence','detail crop [W,S,E,N]','basis'],[[k,v['lat'],v['lon'],v['half_m'],v['cadence']+' (2023-01..2026-10) + monthly 2025-07..2026-10',str(crops.get(k)),v['basis']] for k,v in AOI.items()],[14,9,10,8,30,40,100])
wb.save(f'{OUT}/sites.xlsx'); print('saved',len(T),'timeline rows')
json.dump(SITES,open(f'{OUT}/scripts/sites_summary.json','w'),indent=1)
