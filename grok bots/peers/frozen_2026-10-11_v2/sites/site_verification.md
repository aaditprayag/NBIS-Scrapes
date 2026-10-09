# Peer key sites: ACTIVE vs ANNOUNCED MW (power, permit and satellite evidence)

_As of Mon Oct 5, 2026 (ET). Appended Monarch WV. Power/permit verification pass for Denton, Childress, Abilene completed later the same day. Public information only. This uses the same method as the Nebius alt-data pack (`/workspace/nebius/capacity/altdata/`, `site_verification.md`) so results are directly comparable: Sentinel-2 L2A quarterly 2023-01→2026-10 plus a monthly pass from 2025-07 (Element84 Earth Search, SCL-screened <5% cloud), tight campus crops with the same roof/bare-ground metrics, and Landsat 8/9 ST_B10 campus-minus-3 km thermal (Planetary Computer). The grades also match: **Green** = strong evidence the stated go-live is on track; **Yellow** = partial evidence or slippage risk; **Red** = contradictory or no evidence. "Active" = evidence-supported energized and in service as of 2026-10-04. MW is IT (critical) unless stated. Workbook: `sites.xlsx` (Sites, Timeline, Imagery_metrics, Imagery_AOI_metrics, Thermal, Sites_AOI)._

## Summary grade table

| Peer | Key site | Announced | Contracted | Active now L / B / H (MW) | YE26 L / B / H | Grade | Confidence |
|---|---|---|---|---|---|---|---|
| CoreWeave | Denton TX (Core Scientific) | ~262 MW IT / 394 MW grid | ~260 MW IT | **250 / 260 / 262** | 260 / 262 / 262 | **Green** | High |
| IREN | Childress TX | 750 MW grid; 4 x 50 MW IT Horizons (+H5-6 2027) | 200 MW IT (MSFT) + 60 MW (NVIDIA, 2027) | **50 / 75 / 100** AI IT (plus legacy mining load, winding down) | 100 / 150 / 200 | **Green** (H1-2) / **Yellow** (H3-4 Q4-26) | High (50) / Medium |
| Nscale | Kvandal (Narvik) NO | 230 MW grid allocated (+290 paused); 100 MW HENT buildings | 100 + 100 MW (S-1, not active) | **0 / 0 / 0** | 0 / 0 / 0 | **Yellow** (2027) | High |
| Nscale | **Monarch Compute Campus, Mason Co. WV** | Phase 1: 2 GW gross / 1.37 GW IT by H1'28; ~8 GW by 2031 | **460 MW IT** (Anthropic / Vera Rubin; S-1 table). Financing not committed | **0 / 0 / 0** | 0 / 0 / 0 (H1'28: **0 / 460 / 1,370**) | **Yellow** (460) / **Yellow-Red** (1.37 GW) | High (active=0); Medium (contract); Low-Med (H1'28) |
| Crusoe | Abilene TX (Stargate) | 1.2 GW gross / 8 bldgs (~824 MW IT implied = 618 ÷ 0.75; Epoch ~843) | All 8 bldgs to Oracle/OpenAI | **421 / 618 / 618** [D on R]¹ | 618 / 721 / 824 | **Green** (B1-6 delivered) / **Yellow** (B7-8 by YE; expansion permit frozen) | Medium-High |
| Lambda | Kansas City MO (Blitz) | 24 MW (>100 MW potential) | 1 unnamed customer | **0 / 12 / 24** | 12 / 24 / 24 | **Yellow** | Low |
| Oracle (OCI) | Port Washington WI (Vantage Lighthouse) | ~902 MW IT / 1.3 GW | Oracle for OpenAI | **0 / 0 / 0** | 0 / 0 / 0 | **Yellow/Red** (end-2027 handover vs grid) | High |

¹ **Updated Thu Oct 8, 2026 (ET):** Oracle's FQ1 FY27 call (Sep 10) says 6 of 8 Abilene buildings, 618 MW and 75% of capacity are delivered to the customer; delivered + accepted (R) is read as active (D). IT vs gross basis unstated. The Oct 5 row (200 / 421 / 421 active; YE26 421 / 632 / 843; Green B1-4 / Yellow B5-8) is kept below as history only. See `appendix_crusoe_abilene.pdf`.

How this compares with Nebius (same method, `/workspace/nebius/capacity/site_verification.md`): Nebius's largest site, Vineland NJ, is graded Yellow/Red at 57 / 82 / 107 MW IT for YE26, running on unpermitted temporary gensets. Among peers, only Denton and Abilene have hundreds of MW of evidenced-active IT on the grid. Kvandal, Port Washington and Monarch WV are at 0 active MW; Monarch's 460 MW Anthropic contract is real in the S-1 but still Yellow on the H1'28 path (financing uncommitted, draft air permit, civil-only imagery).

Oracle overlap: Abilene is Oracle's largest live site, but it is counted once, under Crusoe. Port Washington is used as Oracle's own key site, and it shows **no active MW**. That reflects the state of Oracle's next-wave capacity, not OCI's existing fleet.

## Site detail and sources

### 1. CoreWeave: Denton, TX (Core Scientific landlord). Green
- **Announced and contracted:** 262 MW critical IT / 394 MW grid; 8 buildings A–H; 370,284 sf; ~64 ac on City of Denton land; PUE ~1.5; T5 Construction is GC. Work began Q3 2024. Source: [CORZ 8-K Ex.99.2, 2026-04-21](https://investors.corescientific.com/sec-filings/all-sec-filings/content/0001193125-26-165121/d149019dex992.htm). The +70 MW expansion took it to just over 260 MW ([PR 2025-02-26](https://www.businesswire.com/news/home/20250226866161/en/Core-Scientific-and-CoreWeave-Announce-%241.2-Billion-Expansion-at-Denton-TX-Site)).
- **Active (high confidence):**
  - Buildings D–G (~132 MW) were billing in March 2026 (same 8-K).
  - ~260 MW leased and ~260 MW billing, "substantially complete" ([Q2-26 deck, 2026-07-28](https://www.sec.gov/Archives/edgar/data/1839341/000183934126000013/q2fy26earningsdeck728am.htm)).
- **Power and ERCOT (verified 2026-10-05):** the existing 297 MW load was energized before 3/25/2022 and conditionally approved by ERCOT as "Base Load Pathway (a)", outside Batch Zero. Another 74 MW was previously incorporated in and validated by ERCOT's 2025 RTP study cases and is also outside Batch Zero ([8-K 2026-09-10](https://investors.corescientific.com/sec-filings/all-sec-filings/content/0001839341-26-000023/0001839341-26-000023.pdf)).
- **Municipal:** Denton City Council approved the PPA and land-lease amendments on 2024-11-19 (78.85 ac). CORZ pays for interconnection, including two new on-site substations. No city tax incentives were requested. Transmission upgrades are due by 2029 ([City of Denton 2024-11-20](https://www.cityofdenton.com/CivicAlerts.aspx?AID=922)).
- **Air permits (verified 2026-10-05):** TCEQ PBR Registration **179143**, acknowledged **2025-02-11**, under 30 TAC §106.511 (portable/emergency/standby engines and turbines) for "Core Scientific Data Center Site 1 And 2 Denton" (RN111355798 / CN605945302; [TCEQ ack](https://records.tceq.texas.gov/cs/idcplg?IdcService=TCEQ_EXTERNAL_SEARCH_GET_FILE&Rendition=Web&dID=8656717); `docs/denton_tceq_pbr_179143.pdf`). No major NSR air permit found this pass — campus is grid-powered via Denton Municipal Electric; PBR covers emergency/standby gensets only. Confidence: **High** on PBR; **Medium** that no major NSR exists (absence based on public TCEQ/web search).
- **Imagery:**
  - 2023: a pre-conversion bitcoin-shed campus; crop white-roof ~9.8 ha (`imagery/denton_tx/detail/denton_tx_2023-06-28_crop.png`).
  - Jul 2025: a new equipment yard to the east.
  - Oct 2025: a large southern pad.
  - Jun–Sep 2026: white-roof ~18.6–18.7 ha (`.../denton_tx_2026-09-15_crop.png`; `strips/denton_tx_strip.png`).
- **Thermal:** the campus had load before the AI conversion, so Landsat can't separate HPC from mining. The Jun–Aug anomaly rose from −1.56 °C (2023) to −0.76 °C (2026). Low confidence.

### 2. IREN: Childress, TX. Green for H1–2; Yellow for H3–4 by Q4-26
- **Site and power (verified 2026-10-05):**
  - 620 FM 1033; 576 ac; IREN-owned on-site substations with dual physically diverse ERCOT connection ([iren.com](https://iren.com/data-centers/childress)).
  - Grid raised 600→750 MW on 2024-07-24 after ERCOT network-study approval and an amended AEP connection agreement ([IREN PR](https://www.nasdaq.com/press-release/iren-increases-childress-power-capacity-600mw-750mw-2024-07-24)).
  - Aug 2025: 650 MW of operating data centers (Phases 1–5, bitcoin), and a second bulk transformer energized for Horizon 1–2 ([IREN Aug-2025 update](https://iren.gcs-web.com/static-files/ec8ec2a6-a1d1-4bb6-9cb9-bc2d436b9db9)).
  - Site is already energized load — not a Batch Zero *new* interconnection story (contrast IREN Sweetwater). Confidence: **High**.
  - **Regional transmission (distinct from site interconnect):** ERCOT endorsed AEP's Childress Area Transmission Improvement Project (Tier 2, ~$77.3M Option 1A, ISD Nov 2031) on 2025-11-05/10 to address Childress County voltage/thermal violations from load growth; CCN required ([PUC 55999 Item 173](https://interchange.puc.texas.gov/Documents/55999_173_1556587.PDF); `docs/ercot_aep_childress_tier2_endorsement.pdf`). This is area reliability for future load, not the mechanism that delivered IREN's existing 750 MW.
- **Active AI (high confidence for 50 MW):** Horizon 1 (50 MW IT, GB300 NVL72, direct-to-chip liquid) was delivered to and accepted by Microsoft on 2026-08-13 ([SEC Ex.99.1](https://www.sec.gov/Archives/edgar/data/1878848/000114036126032638/ef20080141_ex99-1.htm)).
- **Medium confidence for up to 100 MW:** Horizon 2 is "commissioning underway". Horizons 3–4 are in "late-stage construction, targeting delivery in Q4 2026" ([FY26 results 2026-08-27](https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm)).
- **Contracts:**
  - Microsoft: 200 MW IT, $9.7B.
  - NVIDIA: $3.4B within 60 MW of existing halls, from early 2027 (Q3 FY26, 2026-05-07).
  - Bitcoin mining is being decommissioned. The mining load still sits inside the energized 750 MW, so **site grid load ≠ AI IT**.
- **Imagery** (`strips/childress_tx_strip.png`):
  - White-roof 0.5 ha (Mar 2023) → 4.6 (Jul 2024) → 13.2 (May 2025) → ~17–18 (Jul–Aug 2025) → 19–25 ha (Mar–Jul 2026).
  - New halls on the west side appeared Oct 2025 – Mar 2026. Attribution to the Horizons is likely but not confirmed (low).
  - A new grey-roofed facility to the SE and new northern halls by 2026-09-28 (`.../childress_tx_2026-09-28_crop.png`).
- **Permits / incentives (gap closed 2026-10-05):**
  - **TCEQ Title V:** Operating Permit application **#4951**, applicant **Ie US Development Holdings 3 Inc.** (IREN SPV; confirmed via [Texas Comptroller data-center list](https://comptroller.texas.gov/taxes/data-centers/data-center-lists.php) LD926028 and TDLR TABS2024024946 / TABS2026003980 at 620 FM 1033). Area name "Childress Data Center"; received **2026-08-17**; status **Administrative Complete** ([TCEQ Title V pending list](https://www.tceq.texas.gov/assets/public/permitting/air/reports/applications/titlev-pending-permits.html), data updated 2026-10-05). Title V consolidates operating requirements — it is **not** a pending major NSR construction permit.
  - **NSR / on-site generation:** no public major NSR or EGU standard-permit registration for Childress turbines found this pass. Campus is AEP/ERCOT grid-connected (unlike Abilene's behind-the-meter gas plant). Search scope: TCEQ Title V pending list; public web + TCEQ records queries for "Childress Data Center" / "IE US Development" / IREN PBR & standard permits — only Title V surfaced. **Confirmed absence of a found major NSR** within that scope (not a claim that none exists in unindexed TCEQ files).
  - **County Chapter 312:** Childress County tax abatements + Iris Reinvestment Zone for IE US Development Holdings 3 / IE US Hardware 1–4; Aug-2023 notice contemplated ~$1.4B improvements; amended/restated agreements posted on the [county tax-abatements page](https://www.childresstx.us/pages/tax-abatements) (`docs/childress_county_abatement_notice_2023-08.pdf`). Confidence: **High** on existence; **Medium** on exact MW/phase terms in the latest amendments (full agreement PDFs not OCR'd this pass).
- **Thermal:** dominated by mining; the anomaly is negative (cool white roofs). Not useful.

### 3. Nscale: Kvandal (Narvik/Bjerkvik), Norway. Yellow (operations 2027)
- **Announced and contracted:**
  - Statnett allocated 230 MW. Another ~290 MW is uncertain because Statnett paused new allocations in northern Norway in April 2026.
  - Phase 1 is 32,000 m², ~NOK 8B, "full operation in about a year" ([NRK](https://www.nrk.no/nordland/nscale-sar-tvil-om-videre-satsing-pa-gigant-datasenter-i-narvik-1.17893773)).
  - HENT has 3 buildings (KNB4–6), 100 MW, completing during 2027 ([HENT 2026-09-02](https://www.hent.no/2026/09/02/hent-har-signert-kontrakt-for-utbygging-av-datasenteret-til-nscale-i-narvik-regionen/)).
  - The Vattenfall PPA runs 2027–31, with operations starting 2027 ([Vattenfall](https://group.vattenfall.com/press-and-media/pressreleases/2026/vattenfall-and-nscale-partner-to-support-ai-infrastructure-growth-in-norway/)).
  - Financing: Kvandal South DC Facility $725M + $65M VAT (Nscale S-1, Jul 2026).
- **Active: 0 (high confidence).** Groundbreaking was 2025-11-19 ([NRK](https://www.nrk.no/nordland/legger-ned-grunnstein-for-datasenter-pa-kvandal-1.17658846)).
- **Imagery:** graded pads only through 2025-10-04 (`.../kvandal_no_2025-10-04_crop.png`). The first shells appear 2026-06-10, and long halls by 2026-09-11 (`.../kvandal_no_2026-09-11_crop.png`; `strips/kvandal_no_strip.png`).
- **Regulatory risk:** on 2026-09-29, NVE warned of coercive fines because construction started without the required waste-heat analysis. The deadline is 2026-11-15 ([NRK](https://www.nrk.no/nordland/nscale-varsles-om-tvangsmulkt-1.18039577)).
- **Site location:** the north-Bjerkvik industrial pads are identified from the zoning description ([345.no 2024-04](https://www.345.no/2024/04/ett-steg-naermere-datasenter-i-narvik-kommune/)) plus visible new construction. Medium confidence that this crop covers all of Kvandal North/South.

### 4. Crusoe: Abilene, TX (Lancium Clean Campus; Oracle/OpenAI Stargate). Green for B1–4; Yellow for B5–8 by YE26
- **Power (verified 2026-10-05):**
  - 138 kV 200 MW substation, built from May 2022, with a third bay added H2 2024; AEP interconnection / slack span to Abilene NW ([Mortenson](https://www.mortenson.com/projects/abilene-data-center-development)).
  - 1 GW 345 kV greenfield substation (12.5 ac): first transformer energized 2025-12-19; all 5 MPTs by 2026-03-10 (3 temporary); permanent swaps targeted by Oct 2026 (same Mortenson page). Confidence: **High**.
  - Behind-the-meter generation: TCEQ **Standard Permit Registration 177263** (Abilene DC 1, LLC dba Crusoe; RN112029079 / CN606295137) — **5× Solar Titan 350 (38 MW) + 5× GE LM2500 (34.1 MW) = ~360.5 MW**, 8,760 hrs/yr authorized, onsite use only / no grid export. Originally registered ~2024-08; modification tech-reviewed/approved Jan 2025 (`docs/abilene_tceq_177263_tech_review.pdf`). Diesel backup gensets under separate PBR (Tribune: **62** units). Late-March 2026 thermal image shows a turbine running ([Texas Tribune / Floodlight 2026-07-09](https://www.texastribune.org/2026/07/09/texas-data-centers-ai-power-plants-pollution-state-permits/)). Crusoe told Floodlight turbines are for backup.
- **Air permits — minor vs major (verified 2026-10-05):**
  - **Issued (minor):** Std Permit **177263** for the 10 turbines (above); diesel PBR (reported Reg. 177262 / §106.511).
  - **Pending (major):** State Air Quality Permit **182126** + **PSDTX1688** + **GHGPSDTX263** (Crusoe Technologies LLC / Crusoe Energy Systems LLC; Longhorn Data Center RN112029061 @ 615 FM 2404). Filed **2025-11-07**; NORI **2025-11-19**; NAPD + draft permit **2026-08-05** (`docs/abilene_tceq_182126_napd.pdf`, `_nori.pdf`, `_pls.pdf`). NORI describes authorization for Longhorn Buildings 9–10 / additional onsite generation. Expansion unit counts (**41 turbines + 18 gensets**) are from Tribune/Floodlight — the NAPD public notice does not enumerate units. Process risk: Gov. Abbott directed TCEQ to halt new data-center permits pending ERCOT/TWDB audits (agency update due **2026-10-19**).
- **Incentives:** City of Abilene 85% and Taylor County 80% 10-year abatements, with a separate agreement per building ([DCOA 2025-05-01](https://developabilene.com/2025/05/01/project-radiance-powering-the-future-of-ai-in-abilene-texas/)). Medium confidence.
- **Active:**
  - Company-stated: Crusoe said the first phase was "live" on 2025-09-30, with construction from June 2024 and GB200 racks from June 2025 ([Crusoe](https://www.crusoe.ai/resources/newsroom/crusoe-announces-flagship-abilene-data-center-is-live)). Oracle said on 2026-04-22 that "200MW is already operational" (quoted by Epoch). This gives the **low = 200 MW**.
  - Third-party satellite model ([Epoch AI](https://epoch.ai/data/ai-data-centers/directory/openai-stargate-abilene), imagery to 2026-07-27): 211 MW by 2025-12-24 and **421 MW by ~2026-05-23** (base), 320 chillers. Buildings 5–8 are roofed but not confirmed operational; Epoch projects 843 MW around 2026-11-01. The YE26 base of 632 MW is our haircut, assuming half of B5–8 by year-end. _(Oct 5 history; superseded Oct 8 by Oracle's Sep 10 statement of 618 MW delivered, see note ¹ above.)_
- **Imagery** (`strips/abilene_tx_strip.png`):
  - Pre-construction on 2023-07-18.
  - North pair roofed by 2025-04-10 (`.../abilene_tx_2025-04-10_crop.png`).
  - 6 roofs by 2025-07-30.
  - Chiller rows by 2025-10-03.
  - All 8 roofed by 2026-03-26.
  - Large new southern works (the adjacent Microsoft campus) Jun–Sep 2026 (`.../abilene_tx_2026-09-15_crop.png`).
  - The white-roof metric reads low here because the roofs are beige; visual reads are used instead.
- **Thermal:** the Jun–Aug anomaly went from +0.40 °C (2025) to −0.41 °C (2026). Not informative, because red bare soil confounds the measurement.

### 5. Lambda: Kansas City, MO (Project Blitz, 11155 N Airworld Dr). Yellow
- **Announced:** 24 MW in early 2026, >100 MW potential, >10k Blackwell Ultra GPUs, single customer ([Lambda 2025-10-28](https://lambda.ai/blog/lambda-to-build-a-100mw-ai-factory-in-kansas-city-mo)).
- **Permits:**
  - MO DNR de minimis permit 102025-010, issued 2025-10-29 ([MO DNR](https://dnr.mo.gov/air/business-industry/air-permits/airworld-data-center-kansas-city-102025-010)).
  - MO DNR letter of 2026-04-15: commissioning from 2026-06-01 on 5 × 3 MW rental prime-power engines (Jun–Oct), plus up to 6 × 3 MW more. Evergy power partial at first, full by late summer 2026 ([PDF](https://dnr.mo.gov/sites/dnr/files/vfc/2026/04/main/airworld-blitz-data-center-kansas-city-2026-03-061.pdf)).
  - The Port KC bond application was withdrawn in June 2026. Evergy said power is "available today" ([KSHB](https://www.kshb.com/news/local-news/data-centers/northland-ai-data-center-withdraws-bond-application-with-port-kc)).
- **Active: unconfirmed.** No Lambda, customer or press confirmation of go-live was found as of 2026-10-04. The base of 12 MW is half of the 24 MW, weighted by the commissioning plan. **Low confidence.**
- **Imagery:** this is an existing building, and an interior retrofit can't be resolved at 10 m. There is only minor brightening of the roof/yard in Apr–Oct 2026 (`strips/kcairworld_mo_strip.png`). Thermal was not run because the crop is only ~5×4 TIRS pixels.

### 6. Oracle (OCI): Port Washington, WI (Vantage "Lighthouse", Stargate). Yellow/Red vs end-2027 handover
- **Announced:** ~1.3 GW electrical / ~902 MW IT, 4 buildings, 672 ac, $15B, TIF $458.6M ([Ozaukee Press 2025-10-22](https://www.ozaukeepress.com/content/openai-oracle-run-stargate-port-data-center)).
- **Construction status:** 20–25% complete with ~3,000 workers. The first building was topped out on 2026-05-12. Handover is expected from end-2027 ([Ozaukee Press 2026-09-23](https://www.ozaukeepress.com/content/data-center-size-%E2%80%98-insane%E2%80%99)).
- **Air permit:** WDNR final permit 25-MIN-094 issued 2026-06-12. It covers up to 45 emergency diesel engines totalling 87 MW of backup (`docs/wi_dnr_vantage_25-MIN-094_final.pdf`; [WPR copy](https://www.wpr.org/wp-content/uploads/2026/07/dnr.state_.wi_.us_Apps_AM_WARP_Documents_PROD_Final_63297.pdf)).
- **Grid:**
  - The ATC Ozaukee County interconnection ($1.39–1.67B, new Decker substation) is scheduled for construction Dec 2026 – Dec 2027 ([PSC doc 597276](https://apps.psc.wi.gov/ERF/ERFview/viewdoc.aspx?docid=597276)).
  - The PSC revoked completeness in Aug 2026; ATC refiled on 2026-09-14/15 ([Wisconsin Watch](https://wisconsinwatch.org/2026/09/wisconsin-port-washington-data-center-atc-restarts-push-grid-upgrades-transmission-utility/)).
  - Aterio (third party): partial power earliest Oct 2027, base Dec 2027 ([Construction Review 2026-10-01](https://constructionreviewonline.com/project-lighthouse-delays-threaten-oracles-2027-wisconsin-data-center-launch/)).
- **Active: 0 (high confidence).**
- **Imagery:**
  - Farmland through 2025-10-04 (`.../portwash_wi_2025-10-04_crop.png`).
  - Mass earthworks by 2026-03-27.
  - ~21.6 ha of white roof by 2026-06-23 and ~22.9 ha by 2026-09-11 (`.../portwash_wi_2026-09-11_crop.png`; `strips/portwash_wi_strip.png`).
  - This site serves as the method's "construction-only" control. Its thermal anomaly is ≈0 (+0.4 °C on average in 2026, n=3).



### 7. Nscale: Monarch Compute Campus, Mason County WV (Point Pleasant). Yellow (460 MW path) / Yellow-Red (full 1.37 GW by H1'28)

_Location:_ 30 Monarch Way, Ohio River Road (WV-62) & Airport Rd, north of Point Pleasant (former Mason County Airport). **38.914258°N, −82.100935°W** ([WVDEP draft R14-0042, 2026-09-29](docs/monarch_r14_draft_permit.pdf)). ~2,250 acres. Mason County.

#### Announced vs contracted vs active

| | MW (IT unless noted) | Source / confidence |
|---|---|---|
| **Announced Phase 1** | 2 GW gross / **1.35–1.37 GW IT** by H1'28; ~8 GW gross by 2031 | S-1 (Mar 31 2026 acquisition narrative); [nscale.com](https://www.nscale.com/ai-infrastructure/monarch) (1.35 GW). High |
| **Contracted (S-1 table)** | **460 MW** US Owned Tier 3 Contracted | S-1 facilities table (site #17). Only this slice sits in the company-wide 1.37 GW active+contracted total — do **not** add the full Phase 1 on top of N22. High |
| **Anthropic deal** | up to **$44.6B**, Vera Rubin NVL72, four tranches (signed 2026-08-25) | S-1. Press (Bloomberg/Reuters via WV MetroNews): ~460 MW / ~$45B / 6 yrs → **~$16.2M ACV per IT MW-yr**. High for S-1 dollars; Medium for the 6-yr/460 MW press mapping |
| **Active now (Oct 2026)** | **0 / 0 / 0** | Imagery + company: civil/grading only. High |
| **YE26** | **0 / 0 / 0** | No halls roofed as of 2026-09-25; air permit still draft. High |
| **H1'28 L / B / H** | **0 / 460 / 1,370** | Low = financing fails or permit/build slips past H1'28; Base = Anthropic 460 MW delivered; High = full Phase 1 IT. Medium / Low-Medium |

**Is the 460 MW "real" on the S-1 timeline?** Commercially **yes** — it is explicitly in the S-1 contracted facilities table and backed by signed Anthropic GSAs. Execution is **not yet Green**: (1) S-1 says financing for Anthropic is **not** committed; (2) air permit R14-0042 is still a **draft** (public comment into late Oct 2026); (3) as of 2026-09-25 the site is still graded pads with **no completed hall roofs**; (4) land was under option at the April permit filing; (5) the permit-app startup date (~Jan 1 2027) looks optimistic vs the S-1 H1'28 target and vs imagery. Grade **Yellow** for 460 MW by H1'28; **Yellow-Red** for the full 1.37 GW by H1'28.

#### Power / interconnection (not a utility queue story)

- **Off-grid AI microgrid** — no AEP / FirstEnergy / PJM interconnection for primary load ([Nscale](https://www.nscale.com/ai-infrastructure/monarch); Caterpillar generation partner).
- WVDEP air application: **864 × 2.5 MW** natural-gas RICE = **2.16 GW** nameplate; peak site demand ~**1.8 GW**; 3 H-shaped halls × 4 power nodes; closed-loop cooling ([application PDF](docs/monarch_r14_application.pdf)).
- **14-mile gas pipeline** Putnam/Mason counties ([WSAZ 2026-06-16](https://www.wsaz.com/2026/06/16/pipeline-project-proposed-power-mason-county-data-center/)).
- Primary risks are gas supply, engine delivery, air-permit finalization and financing — **not** PJM queue position.

#### Permits / incentives

- Air: R14-0042 applied **2026-04-20**, draft **2026-09-29** (not final). Facility ID 053-00136.
- Stormwater: WVDEP enforcement order **2026-07-14** after sediment left the site ([WV News](https://www.wvnews.com/news/wvnews/public-hearings-set-for-monarch-cloud-campus-air-permit-in-mason-county/article_e2d06911-0746-4a33-a9ba-aaa8e1bf01c0.html)).
- High Impact Data Center certification announced **2026-09-03** ([WVPB](https://wvpublic.org/story/energy-environment/second-certified-data-center-announced-for-mason-county/)); Morrisey also claimed $69.2B total client investment (Anthropic + undisclosed $24.2B — add-on MW not disclosed).
- State forgivable loan $62.5M to affiliate Mountaineer GigaSystem (press; $25M funded / $37.5M waived) — secondary.
- No local PILOT/abatement found in this pass (gap); HB 2014 limits county ordinance authority over High Impact sites.

#### Imagery (Sentinel-2, same method)

- 21 scenes, 2023-03-30 → 2026-09-25. Crop `[-82.122, 38.902, -82.088, 38.925]`.
- **2026-04-08:** runway/hangars intact (`imagery/monarch_wv/detail/monarch_wv_2026-04-08_crop.png`).
- **2026-05-28 → 06-02:** bare ground 48 → ~95 ha — mass grading (`.../monarch_wv_2026-06-02_crop.png`).
- **2026-09-25:** ~85 ha bare, graded pads, white-roof still ~2 ha (hangars); **no completed DC halls** (`.../monarch_wv_2026-09-25_crop.png`).
- Strip: `strips/monarch_wv_strip.png` (full AOI: `strips/monarch_wv_full_strip.png`).
- Site manager 2026-09-20: civil ongoing, about to go vertical; ~900 workers ([WV MetroNews](https://wvmetronews.com/2026/09/20/monarch-compute-campus-to-go-vertical-soon/)).

#### Gaps specific to Monarch

1. Anthropic tranche MW schedule and financing close date not disclosed beyond "four tranches" / best-efforts financing.
2. Air permit not final; Title V to follow within 12 months of issuance.
3. Land deed/transfer completion not confirmed (option at April filing; S-1 lists land-transfer risk).
4. Gas pipeline permit status not tracked to a final approval in this pass.
5. Landsat thermal completed (construction control): Jun-Aug anomaly ~0 to slightly positive/negative; not informative for MW. See thermal/monarch_wv_landsat_st.csv.
6. Undisclosed $24.2B "client investment" (Morrisey) has no MW attribution — ignored for contracted MW.

## Key images
- `strips/abilene_tx_strip.png`: 8 buildings appear in sequence, 2023→2026.
- `strips/childress_tx_strip.png`: mining campus → new AI halls.
- `strips/portwash_wi_strip.png`: farmland → 22 ha of roofs in 11 months, with no power.
- `strips/kvandal_no_strip.png`: bare pads → first shells in 2026.
- `strips/monarch_wv_strip.png`: airport/fields Apr-2026 → ~95 ha graded pads by Jun/Sep-2026; no halls yet.
- `strips/denton_tx_strip.png`: bitcoin campus → fully roofed HPC campus.
- Full-AOI context: `strips/*_full_strip.png`.
- Thermal plots: `thermal/*_landsat_st.png`.

## Not disclosed / unresolved (gaps)
1. **IT vs grid basis:** Childress active AI IT can't be separated from the legacy mining load on the grid. Kvandal's 230 MW is a grid allocation, not IT.
2. **Lambda KC:** no go-live confirmation, no customer name, and no Evergy energization record.
3. **Abilene:** the 421 MW figure is a third-party satellite model (Epoch). The only company figure found is 200 MW (Apr 2026). The ERCOT large-load filing for Lancium wasn't retrieved. _(Oct 5 history; superseded Oct 8 by Oracle's Sep 10 statement of 618 MW delivered, see note ¹ above.)_
4. **ERCOT large-load queue (Childress/Abilene):** not accessible in this pass. Only the CORZ 8-K covers ERCOT status, and only for Denton.
5. **Permits — residual gaps after 2026-10-05 pass:** Childress Title V #4951 and Chapter 312 abatements now documented; **no Childress major NSR / EGU standard-permit registration number found** in public search (absence within stated scope). Abilene minor (177263) and major pending (182126/PSDTX1688/GHGPSDTX263) now documented with primary TCEQ PDFs; diesel PBR registration number (reported 177262) not independently downloaded this pass. Denton PBR 179143 documented; no Denton major NSR found. Still open: Taylor County/TDLR filings for Abilene (Epoch cites TDLR); Narvik municipal building permits; full Childress abatement MW/phase schedules (agreement PDFs not OCR'd).
6. **Thermal:** Landsat daytime surface temperature (100 m, ~10:30 local) didn't separate operating from non-operating sites. Anomalies are mostly negative because of cool white roofs. As with the Nebius sites, it is **low confidence**.
7. **Imagery resolution:** Sentinel-2 (10 m) can't resolve chillers, generators or substation equipment reliably. Statements about equipment rely on Epoch, Mortenson and press.
8. **Tooling:** the OSM Overpass API and the US Census geocoder failed. Sites were located by Nominatim plus visual search. AOIs and crops are listed in `sites.xlsx` › Sites_AOI.
