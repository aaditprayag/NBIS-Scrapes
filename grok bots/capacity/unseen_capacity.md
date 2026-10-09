# Nebius (NBIS): unexplained connected power, unseen sites and calibration
_As of Sun Oct 4, 2026 (ET); GPU anchor updated 10:45 PM ET. Public sources only: no logins, no paid data, nobody contacted. Tags: **R** = Reported, **D** = Derived, **E** = Estimated. Companion files: `mw_path_proposed.csv` (proposed path) and `vineland_shutdown.xlsx` (shutdown scenarios). `nebius_capacity.xlsx` was not modified by this note; it was later revised to v2 (backup `nebius_capacity_v1.xlsx`)._

## 1. The ~496 MW YE26 plug: what explains it
- **The plug.** Management guide is 800 MW to 1 GW **connected** at YE26, reaffirmed on the Q2'26 call on Aug 12, 2026 (R; `/workspace/nebius/src/q2t.txt`). Identified-site energized base at Q4'26 is 304.5, so the plug is 495.5 MW (D; nebius_capacity.xlsx rows 121–124).
- **Main finding: the plug is mostly a basis mismatch, not missing sites.**
  - "Connected" = "power connected into fully built and equipped data centers" (R; Q1'26 letter). It is not limited to active IT load.
  - The identified-site rows in the model are energized/active (IT) figures.
  - Rebuilding YE26 on a connected/shell basis gives about **941 MW base (561 low / 1,195 high)** (E). That sits inside the 800–1,000 MW guide.
- **What fills the gap at YE26 (base, E):**

| Item | YE26 connected (base) | Active in model | Gap filled | Basis | Source |
|---|---|---|---|---|---|
| Vineland at full shell (gross) | 300 | 81.8 | ~218 | Mgmt: "Vineland is part of the 2026 capacity and connected power" (R) | q2t.txt, Aug 12 2026 |
| Independence MO (IPL 200 MW late Q3'26) | 100 (0–200) | 0 | ~100 | Utility feed, built portion | site_verification.md |
| Birmingham AL (115 kV feed by Dec 31) | 50 (0–100) | 0 | ~50 | Initial energization, MW not disclosed | site_verification.md |
| Béthune FR (mgmt 120 MW Q4'26) | 120 | 30 | ~90 | Mgmt figure | nebius_capacity.xlsx row 27 |
| Oklahoma colo | 100 (50–150) | 0 (n/d) | ~100 | Map category ">100MW" | Q2'26 letter p.2 |
| Minnesota colo | 15 | 0 (n/d) | 15 | IronGate MSP01 15 MW | see §2 |
| Newport + UK 3rd | 15 (0–30) | 0 (n/d) | 15 | 43 MW by 2027 | see §2 |
| KC to 40 (from 20), Israel, Madrid, Tallinn, UK at contract level | — | — | ~40–60 | Contracted MW | model rows |

- **Confidence: medium-low.** The direction is solid (active vs connected basis). Site-by-site numbers are Estimated.

## 2. Lead-by-lead

### Oklahoma (colo, town not disclosed): sizing >100 MW category (R); active E 10 / 25 / 50 MW for Q3'26 / Q4'26 / Q1'27; low confidence
- **What Nebius says**
  - The Q2'26 letter map marks Oklahoma as **"Sites >100MW"** (R). https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf (Aug 12 2026)
  - The 20-F says Nebius "secured the contracts needed to commence operations" for OK (and MN) colo in Feb 2026 (R).
- **Jobs.** Greenhouse public API (https://boards-api.greenhouse.io/v1/boards/nebius/jobs?content=true, pulled Oct 4 2026) shows **8 on-site roles**. Location is listed only as "Oklahoma, United States".

| Job ID | Role | Posted (2026) |
|---|---|---|
| 4780568101 | IT DC Technician | Mar 2 |
| 4866023101 | DC IT Technician ("onsite in our Oklahoma data center") | May 28 |
| 4902968101 | Security | Jun 23 |
| 4942839101 | Admin Services Mgr ("continued growth of the facility… liaison with building management", i.e. colo) | Aug 17 |
| 4965870101 | Warehouse | Sep 15 |
| 4994675101 | DC Site Manager | Sep 30 |
| 4995854101 | Critical Infrastructure Engineer (CDUs, in-rack busbar, so GB300/Vera Rubin-class liquid cooling) | Oct 2 |

- **Location clues (geocodes only):** a Ladders listing shows "Oklahoma City 73160"; Aterio shows Tulsa 74117. Both are likely generic geocodes.
- **Landlord, interconnection, permits and abatements: none found.** Checked:
  - OKC and Tulsa project lists; both cities have moratoria
  - Core Scientific Muskogee: building 2 unleased; "newest customer" delivers 2H27
  - CloudBurst OKC: first 30 MW in Q4'26, delayed
  - Cerebras/Scale OKC: 10 MW
  - No OCC or SPP filing naming Nebius was found.
- **Inference (E):** the site is live at a small scale (techs on site since about May). It is scaling into liquid-cooled racks in Q4'26 and Q1'27.

### Minnesota (Woodbury, us-north1): E 5 / 10 / 15 MW; low-medium confidence
- us-north1 (Woodbury MN, B300) is a public region (R). https://docs.nebius.com/overview/regions
- Jobs (R):
  - Data Center Technician 4765032101 (Feb 20): "onsite in our Minnesota data center"
  - IT Infra Engineer, RMA & Hardware Diagnostics 4765077101 (Feb 24)
- Woodbury colo supply, per Woodbury News Net, Sep 9 2026 (R): https://woodburynewsnet.org/7933/business/what-woodbury-residents-should-know-about-irongates-two-data-centers/
  - IronGate MSP01 (401 Bielenberg Dr): 15 MW live, **fully leased**, 66 MW full build
  - MSP02: 4 MW live plus 10 MW under construction, 6 MW leased
  - Woodbury has a moratorium.
- Nebius-at-IronGate is an inference (E). The map shows MN as a plain dot (<100 MW) (R).
- No PUC or MISO filing names Nebius.

### UK beyond Longcross and Harlow: Newport + 3rd site = 43 MW by 2027 (D); confidence medium (total), low (split)
- **The 43 MW total.**
  - Jun 8 2026 PR: three new UK deployments "combined… will reach 65 MW when fully ramped up in 2027" (R). https://nebius.com/newsroom/nebius-expands-in-uk-with-more-nvidia-powered-infrastructure-more-customers-and-more-cloud-capabilities-for-agentic-and-enterprise-ai
  - 65 minus Harlow 22 = **43 MW** for Newport plus the 3rd site (D).
- **Newport, Wales = Vantage CWL1.** Announced Aug 13 2026; MW not disclosed (R).
  - Vantage PR: https://vantage-dc.com/news/vantage-data-centers-and-nebius-expand-uk-ai-infrastructure-with-first-deployment-in-south-wales-ai-growth-zone/
  - DCD: https://www.datacenterdynamics.com/en/news/nebius-to-lease-data-center-capacity-from-vantage-in-newport-wales/
  - Jobs: Newport DC IT Tech (Jun 19) and Warehouse (Aug 27).
- **UK 3rd site, low-confidence lead:** Green Mountain LON-East (Romford) leased 14 MW to an unnamed neocloud (Sep 1 2026). https://www.datacenterdynamics.com/en/news/green-mountain-secures-neocloud-customer-at-data-center-in-london-uk/
  - Supporting clue: Nebius "IT Infrastructure Engineer (East London)" job 4769160101 (Feb 2).
  - If it is Nebius, Newport ≈ 29 MW (D, low).

### Israel beyond Modi'in and Masmiyya: no additional site found; Beit Shemesh (already modelled, 58 MW) is the only other one
- Deep TASE on Mega Or's Q2 report (Aug 10 2026): https://www.deeptase.co.il/en/analysis/4873
  - The text says customer delivery is Q3'26 for Yoav/Masmiyya and staged Q3'26–Q1'27 for Beit Shemesh.
  - The presentation shows construction completion **Q4'26 (Yoav)** and **Q2'27 (Beit Shemesh)** (R).
  - I have pushed both later in the proposed path (E; medium-low confidence).

### France 2nd site (beyond PA10 and Béthune): MW not disclosed; location near Paris (map); landlord not found
- eu-west2 (France, B300) is a public region (R; docs.nebius.com/overview/regions).
- Paris job 4753448101 (Feb 4) (R).
- Checked Eclairion, Data4, Telehouse and Digital Realty: no Nebius disclosure. The map shows <100 MW.
- Kept at **n/d** in the path, so no number was invented.

### AIB CLT1 (new; per parent): 50 MW critical IT, 2 × 25 MW halls; path 0 / 25 / 50 MW IT by YE27
- **Location (R):** CLT-01 campus, **389 Tyger River Industrial Dr, Moore, SC 29369, Spartanburg County**.
  - 10-K: "40 MW facility in Spartanburg County… ~17.6 acres of leased land". https://www.sec.gov/Archives/edgar/data/2070542/000121390026036810/ea0282624-10k_blockchain.htm
  - Address: colomap / Compute Atlas.
- **Utility (R):** **Lockhart Power**. The ground lease requires Lockhart Power, the only utility for the industrial park (10-K above).
- **Power (R):** 15-yr ESA executed May 2026, expanding from 40 to 65 MW over the existing 34.5 kV line with "no significant additional electrical infrastructure upgrades". https://www.sec.gov/Archives/edgar/data/2070542/000121390026063565/ea029232801ex99-1.htm (May 27 2026)
- **Deal (R):** binding 12-yr MCSA plus 2 × 5-yr renewals; delivered in two halls. https://www.globenewswire.com/news-release/2026/09/30/3371855/0/en/aib-data-centers-signs-contract-with-nebius-for-ai-data-center-capacity.html (Sep 30 2026); 8-K dated Oct 1 2026.
- **Investor deck (R):** "ESA service scheduled to begin October 1, 2026"; "Full construction design submitted for permits"; contract value ≥$1.3B. http://archive.fast-edgar.com/20260930/AB2ZM22DL222P26L222Q2MY2R6GDZ222ZQ82/ea030723301ex99-1.htm
- **Hall timing (R, secondary):** halls at about months 10 and 14, so **~Aug 2027 and ~Dec 2027** (D). Source: savrn tracker summarising the deck, https://savrn.com/data-center-capital-atlas/south-carolina
- **Permits and construction:** permit application submitted; **no permit issuance or construction start found** (R).
- **Local context (R):** Spartanburg County Council said it would not vote on new data-center items, including FILOT abatements, until the state legislature sets rules. This followed the Feb 27 2026 TigerDC "Project Spero" withdrawal from the **same Tyger River Industrial Park**.
  - AIB CEO Jerry Tang was also TigerDC's CEO (WSPA).
  - CLT-01 is an existing permitted facility being repositioned. Risk: county or state action delays permits (E).
  - Sources: https://www.datacenterdynamics.com/en/news/tigerdc-pulls-planned-data-center-project-from-south-carolina-county-after-officials-deny-tax-deal/ ; https://www.wspa.com/news/local-news/live-spartanburg-co-council-members-legislators-provide-update-on-proposed-ai-data-center/
- **Implied colo price (D):** ≥$1.3B / 12 yrs / 50 MW ≈ ≥$2.2M per MW IT per year. That is colocation rent, not Nebius revenue.
- **Path (E):** low 0 / base 25 / high 50 MW IT at YE27. The base differs from the model only in Q3'27: 25 (hall 1 ~Aug 2027) vs 12.5.

## 3. June 2026 revenue step-up: what started billing?
- **Derived from Reported figures (q1_26.txt, q2_26.txt):**
  - ARR = last-month revenue × 12 (R).
  - Revenue by month (D):

| Month(s) | Calculation | Monthly revenue |
|---|---|---|
| March 2026 | $1.92B ARR / 12 | ≈ $160M |
| April–May 2026 average | ($574.9M Q2 AI cloud − $250M June) / 2 | ≈ $162M |
| June 2026 | $3.0B ARR / 12 | ≈ $250M |

  - The step-up is **≈ $88M per month, and it lands entirely in June**: April–May were flat with March.
- **MW equivalent (D):** ≈ 72 MW at $14.49M per IT MW per year (restated Oct 9; the earlier ≈90 MW used the retired $11.6M gross-MW rate); ≈ 88 MW at the ~$12M "2026 base" ACV per MW; ≈ 42–52 MW at $20–25M/MW.
- **Note 7 constraint (parent):** only ~$0.76B of equipment moved into service in Q2'26. At ~$25M per MW for GPUs plus network (gpu_mw_capex.md), that is **~30 MW of new capacity** (D). New Q2 capacity alone cannot explain ~90 MW.
- **Management's own attribution (R, Q2 call):** Q2 revenue was "driven by capacity added in Q1, higher utilization, asset-light, Token Factory, acquisitions". Capacity "deployed late in Q2" contributes from Q3.
- **Most plausible explanation (E, low-medium):** Q1-added capacity started billing late, i.e. contract start or acceptance in June. Candidates:
  - **Vineland T2** (delivered Feb 2026; 62 gensets all running only by June 2026, per Floodlight Oct 1)
  - Mäntsälä fill (62.5 → 75 MW)
  - Q1-built OK/MN capacity
  - Plus price uplift (+30% on older GPUs).
- **No specific site or tranche found that began billing in June 2026.** The Q2 letter places the two newest Vineland tranches in the month before Aug 12, i.e. Jul–Aug, not June (R).

## 4. Calibration issues (parent request)
**(a) $12M per installed MW implies ~190 MW vs the model's 173.**
- Q2 AI cloud revenue $574.9M × 4 / $12M ≈ **192 MW average billing-equivalent** (D).
- At the June exit rate: $3.0B / $12M ≈ **250 MW** (D).
- The model has 173 MW identified at Q2'26 and already needed a 22 MW plug at YE25 (170 Reported active vs 147.7 identified).
- **Most likely missing MW (E):**
  1. **Oklahoma and Minnesota colos:** live from ~Q2'26 per on-site jobs; ~10–25 MW together.
  2. **Kansas City Patmos:** the model holds 20 MW; landlord states 40 MW at full build; worth up to +15.
  3. **Vineland:** the model's 39 MW at Q2'26 understates what 62 gensets could carry, but billing is by tranche.
  4. **Mäntsälä:** 75 MW is the facility figure. If IT ≈ facility it is fully counted; no upside unless #2 started early.
- **Pricing caveat:** ARR per MW also embeds price (Q2 deals at $20–25M/MW). A blended $12M understates MW only if mix is old-gen.

**(b) GPU history vs company disclosures.**
- Nebius **does not publish a quarterly fleet GPU count** in its letters. Q1'25, Q1'26 and Q2'26 letters describe generations, not totals (R).
- **Company statements I found:**

| Date | Statement | Tag | Source |
|---|---|---|---|
| Q4'24 | "Total deployed base of NVIDIA Hopper GPUs… nearly doubled quarter-over-quarter" | R | https://nebius.com/newsroom/nebius-group-n-v-announces-fourth-quarter-and-full-year-2024-financial-results |
| Dec 2024 | Plan for ">22,000 Blackwell GPUs" in US + Finland in 2025 | R | same release |
| Nov 2024 | Kansas City "thousands of Hopper GPUs" | R | same release |
| — | Mäntsälä expansion "up to 60,000 GPUs" | R | press |
| Nov 2025 | Longcross 4,000 B300 + 3,000 phase 2 | R | https://www.datacenterdynamics.com/en/news/nebius-launches-nvidia-gpu-cluster-in-london-uk/ |
| — | Modi'in 4,000 B200 | R | site_list.md |

- **GPU anchor, verified Oct 4 (v2):** the FY24 20-F (filed Apr 30 2025) says "As of March 31, 2025, we have deployed approximately 30,000 GPUs across our global data center footprint, primarily NVIDIA H200s" (R). https://www.sec.gov/Archives/edgar/data/1513845/000155837025005991/nbis-20241231x20f.htm
  - The "~38k" call figure is superseded.
  - "~10k in 2024" still has no primary source.
  - The v2 model (KC from Q1'25 at 5–15 MW) gives ~27.4k modelled GPUs, 91% of 30k (D).
- **Physical check (D, gpu_mw_capex.md):** 38k GPUs on ~35 MW ≈ 1,086 per MW, which is impossible for Hopper. Company-style counts include GPUs delivered but not powered.
- **Sites most likely behind the Q1'25 gap (E):**
  - **Kansas City:** live end-Q1'25; the model has 0 MW at Q1'25 and 5 at Q2'25.
  - **Keflavik:** first capacity March 2025.
  - **Mäntsälä:** expansion hardware delivered ahead of power; the model holds 25 MW through Q2'25.
  - **Delivered-not-energized GPUs** for Q2'25 Blackwell rollouts.
- **2024 overcount (13.7k modelled vs ~10k):** most likely the model energizes Paris (Nov 2024) and Mäntsälä Q4'24 adds too early. "Nearly doubled QoQ" in Q4'24 implies a low Q3'24 base.

## 5. Vineland items (detail in vineland_shutdown.xlsx, Sources tab)
- **NJDEP order (R):** AONOCAPA dated **Sep 22 2026**; penalty $1,072,000; 62 CAT G3516H × 1,982 kW = 122.9 MW.
  - https://dep.nj.gov/wp-content/uploads/newsrel/2026/dataone75894-aonocapa.pdf (bot-protected; text via snippets)
  - Governor release: https://www.nj.gov/governor/news/2026/20260922a.shtml
- **Deadlines (D):** 45-day deadline **Nov 6 2026** if received Sep 22 (receipt date not disclosed; ~Nov 9 if mailed). Both are before the Nov 10 report. Hearing request due ~Oct 12.
- **What triggers a cease (R):** failing to **submit** applications within 45 days. DataOne says it "will apply" (WHYY Sep 23; ENR Sep 23).
- **Air-permit notice / comment period:** none found for gensets or Bloom (R absence). An earlier Bergen-engine application was withdrawn May 27 2026. The NJDEP BOSS notices page is bot-protected.
- **Bloom energized:** no evidence.
  - Stop-construction order Aug 10 2026 (Hunterbrook, https://hntrbrk.com/breaking-news/vineland); no public record of it being lifted.
  - Sentinel-2 (Sep 16 cloud-affected; Oct 1 hazy) and Landsat thermal (Sep 17 anomaly +1.5 °C vs +1.2 °C Aug 24) do not resolve fuel cells. **Not verifiable.**
- **Planning Board:** Oct 14 2026 6:30 pm; agenda not yet posted. **City Council:** Oct 13 agenda has no DataOne item.
- **Tranches:** T1 Nov 2025, T2 Feb 2026, T3–T4 ~mid-Jul to mid-Aug 2026. **No June 2026 tranche found.**
