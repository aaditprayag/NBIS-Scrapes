# Vineland (DataOne / Nebius / Microsoft): permit and filing watch list and 2027 Bloom dependence
_As of Mon Oct 5, 2026, ~12:40 PM ET (Bloom check §1b; Q3 date tag fixed to E). Originally drafted 8:30 AM ET. Public sources only: nothing was logged into and nobody was contacted._
_Tags: **R** = Reported, **D** = Derived, **E** = Estimated. Companion files: `vineland_shutdown.xlsx` (outage scenarios) and `nebius_capacity.xlsx` (Vineland tranche track tab and Bloom-slip rows)._

## 0. Key dates

### Today's watch (Mon Oct 5)
- **Hearing-request window ~Oct 12–15** remains open (20 days after receipt of the Sep 22 AONOCAPA; exact receipt date not disclosed).
- **Oct 14 Planning Board** agenda not yet posted — next check Oct 7 via routine.
- **~Nov 6 NJDEP** genset permit/operating-certificate filing deadline (Nov 9 if mailed). DataMiner/DocMiner still captcha-blocked. **Oct 5 PM Bloom check:** no public PCP/Certificate filing found; no Bloom delivery/install evidence; Aug 10 stop-order still no public lift — see §1b.

| Date (ET) | Event | Tag | Source |
|---|---|---|---|
| Sep 22 2026 | NJDEP Administrative Order and Penalty Assessment (EA ID PEA260001-75894, Facility/Program Interest ID **75894**) issued. $1,072,000 penalty; 62 CAT G3516H engines × 1,982 kW | R | https://dep.nj.gov/wp-content/uploads/newsrel/2026/dataone75894-aonocapa.pdf ; https://www.nj.gov/governor/news/2026/20260922a.shtml |
| **~Oct 12–15 2026** | Hearing-request deadline: 20 days after receipt. Without a request the penalty is final on day 21. A request does **not** stay the corrective orders | D (receipt date not disclosed) | https://morristownminute.org/1-07-million-penalty-vineland-data-center/ |
| Oct 13 2026, 5:30 pm | Vineland City Council: no DataOne item on the agenda | R | https://vinelandnj.api.civicclerk.com/v1/Events |
| **Oct 14 2026, 6:30 pm** | Vineland Planning Board regular meeting; agenda not yet posted (posts about 1 week ahead). Alternate date Oct 29 | R | https://www.vinelandcity.org/event/planning-board-meeting-60/ |
| **Nov 6 2026** (Nov 9 if mailed) | 45-day deadline to **submit** permit and operating-certificate applications. If not submitted, DataOne must cease operating the gensets | D | AONOCAPA text (via snippets); ENR Sep 23 |
| ~Nov 10 2026 (est.) | Nebius Q3'26 results | E | Third-party estimate — **not company-announced**; do not tag Reported |

## 0b. Vineland revenue at risk: one table (added Oct 4 2026, 11:20 PM ET)

All on the Microsoft rate of **$14.5M per IT MW per year** (D). The cases are not exclusive, and the probabilities are my judgment (E), not market-implied.

| Case | What happens | MW IT hit | Revenue hit | Kind of loss | Probability (E) | Probability-weighted |
|---|---|---|---|---|---|---|
| A. Generator shutdown or cap after the NJDEP deadline (~Nov 6–9) | Gensets curtailed or stopped for 1–3 months until permits or Bloom arrive | ~51–103 on T1–T4 | $122–374M one-off in Q4'26–Q1'27 (shutdown scenarios); **$79M probability-weighted** ($63M on the $11.6M gross basis, of which $54M falls in FY26) | Lost revenue on delivered capacity, plus possible SLA credits (not disclosed) | 45% some curtailment, 55% none | ~$79M |
| B. Bloom fuel cells slip 2 quarters | T5–T9 start billing 2 quarters later | 133 billing; −66.5 energized at YE27 | **−$966M of 2027 revenue** ($1,159.5M → $193.3M), mostly moved into 2028 | Timing shift, small value hit | 35% (stop-construction order Aug 10, no public lift; no Bloom air-permit notice found) | ~$338M of 2027 revenue |
| C. Permanent loss of T1–T4 | Gensets stop before Bloom is live, and Microsoft uses its termination rights on the delivered tranches | ~107 | **−$1,547M per year** for the rest of the contract | Permanent | ~5% (needs a long outage plus Microsoft choosing to walk; it needs the capacity) | ~$77M a year |

**One number for the pitch:** the base case is B, with about $1B of 2027 revenue at risk from Bloom timing. The earlier $54–79M is only case A, the short shutdown. C is the tail risk at ~$1.5B a year. Next evidence: the NJDEP permit filing (by ~Nov 6–9), any lift of the Bloom stop-construction order, and the tranche count in the Q3 letter (~Nov 10 est., not company-announced).


## 0c. Locked billing path (team, Oct 5 AM)
Modeling and Skeptic locked these figures for the judge slide (Derived from `nebius_capacity.xlsx` v2 billing rows with 1-quarter lag):
| Path | YE26 billing MW | FY27 average billing MW |
|---|---|---|
| Base | ~258 | ~475 |
| High (like-for-like) | **~284** | **~645** |

Do **not** use the earlier ~350 / ~710 shortcuts (simple YE energized averages × 80–85%). The high-path gap to the ~769 MW needed at $16M comps is about **19%**.

## 1. Filings that would show DataOne applied, and current status

| # | What to check | URL | What a filing would look like | Status (Oct 4 2026) |
|---|---|---|---|---|
| 1 | **NJDEP DataMiner**: air permit reports by Program Interest **75894** (DataOne, 3963 S Lincoln Ave) | https://njems.nj.gov/DataMiner. Use "Search by Category > Air" or a PI # search. Help guide: https://dep.nj.gov/wp-content/uploads/docminer/dmhelpguide.pdf | A new **Preconstruction Permit (PCP) and Certificate to Operate application**. It would show an activity/permit number (format like "PCP2600xx" or a BOP number), the submittal date and status "Application received / under technical review" against PI 75894 | **Not verifiable from here.** The site is bot-protected (Incapsula) for automated access. Needs a manual browser check. No press report of a filing yet; WHYY Sep 23: DataOne "will apply" (R) |
| 2 | **NJDEP DocMiner**: documents under PI 75894 | https://dep.nj.gov/docminer/ | Application package PDFs (forms, emissions calcs, possibly air dispersion modelling), DEP deficiency letters (14-day response clock), any consent order | Bot-protected; not verifiable here. Morristown Minute (Sep 22) used the DEP document index to show the earlier **Bergen-engine application was revised several times and withdrawn May 27 2026** (R) |
| 3 | **NJDEP Bureau of Stationary Sources public notices** (draft permits for comment) | https://dep.nj.gov/boss/public-notices/ | "Notice of intent to issue / modify air permit… Program Interest No. 75894", with a **30–60 day comment period** and possibly a hearing | Bot-protected. **No DataOne notice found** by web search (R absence). A major-source permit (Floodlight: potential NOx far above the 25 tpy threshold) needs public notice. It cannot be issued by Nov 6 (E) |
| 4 | **NJ Environmental Justice law** (N.J.S.A. 13:1D-157; AO 2021-25) | https://dep.nj.gov/ej/ (EJ mapping tool and decisions) | EJ impact statement plus a public hearing notice published by the applicant. Bloom/Equinix Secaucus followed this route: PCP filed Mar 13 2023, 60-day comment Nov 4 2023–Jan 3 2024, hearing Dec 5 2023 | Not found for DataOne. Whether the site lies in an overburdened community is **not verified** (check the EJ map) |
| 5 | **Bloom fuel-cell air permits** (PI 75894, or a Bloom-held PI) | Same DataMiner / BOSS pages. Precedent: https://dep.nj.gov/wp-content/uploads/ej/bloom-equinix-notice-12052023.pdf | N.J.A.C. 7:27-8.2(c) lists fuel-cell systems explicitly (https://www.law.cornell.edu/regulations/new-jersey/N-J-A-C-7-27-8-2). Bloom's Secaucus fuel cell needed a PCP (PCP 230001) (R). A ~328 MW natural-gas fuel-cell plant would most likely need a PCP plus operating certificate (E) | DataOne says it "has the necessary approvals in place" for the fuel cells (R, WHYY Sep 23). **No public Bloom PCP, notice or comment period found** (R absence). Confidence that they are fully permitted: low |
| 6 | **Vineland construction office**: lifting of the Aug 6 (LNG) and Aug 10 (Bloom) stop-construction orders; building/electrical permits and certificate of occupancy | City construction office (OPRA). Orders were published by Sustain SJ (Facebook) and reported by Hunterbrook: https://hntrbrk.com/breaking-news/vineland | A "release of stop construction order" or new UCC permits for the Bloom Power areas / LNG tank. The order states no CO is issued until penalties are paid | **No public record of a lift** (R absence). The site plan was approved Aug 17 (R). The building office declined comment (Hunterbrook) |
| 7 | **Vineland Planning Board**: memorialization of the Aug 17 amended site-plan resolution, any further DataOne amendment, Oct 14 agenda | https://www.vinelandcity.org/Archive/?dir=Planning%20Board/Agendas/2026 ; minutes: https://www.vinelandcity.org/Archive/?dir=Planning%20Board/Minutes/2026 | An agenda item "DataOne Vineland LLC, Block 7503 Lot 35.01 – resolution / amended site plan / minor change" | Sep 9 agenda: none. Latest minutes posted: Aug 12 (none). **Oct 14 agenda: not posted yet** as of Oct 5 AM. Routine checks Oct 7 / 9 / 12 at 6:41 PM ET; agenda typically posts ~1 week ahead (~Oct 7). |
| 8 | **Vineland City Council / Vineland Municipal Electric Utility (VMEU)**: service agreements and resolutions (VMEU is governed by the Council, not BPU-rate-regulated) | CivicClerk API https://vinelandnj.api.civicclerk.com/v1/Events (agenda packets via GetMeetingFileStream) | A resolution for added VMEU service, an interconnection or standby agreement, or a PILOT amendment | Oct 13 agenda: no item. Oct 27: no files yet. Context: VMEU serves only ~5 MW of mechanical load. ACE withdrew the 49 MW Cumberland load need from PJM planning by late Jul 2026 (secondary: https://measuredai.substack.com/p/microsoft-new-jersey-ai-data-center). Not regulated by BPU (SNJ Today) |
| 9 | **NJ BPU** dockets (only relevant if ACE or another investor-owned utility service is sought) | https://publicaccess.bpu.state.nj.us/ | A docket naming DataOne / Northeast Precast / Vineland large load | Bot-protected (403). No press reference to a BPU docket (R absence) |
| 10 | **PJM**: load-forecast or interconnection entries for ACE / Vineland | PJM planning and TEAC materials (pjm.com) | Re-entry of a Cumberland County large-load adjustment | 49 MW entry withdrawn (secondary source, above) |
| 11 | **Pinelands Commission** | https://www.nj.gov/pinelands/appli/ | Application number for 3963 S Lincoln Ave | **Not applicable.** Pinelands Alliance: "Unlike Vineland, the [Williamstown] property… is within the Pinelands boundaries" (R): https://pinelandsalliance.org/dataone-vineland-town-hall/ |
| 12 | **Administrative hearing request** at NJ Office of Administrative Law | OAL docket search (njoal) | A transmitted contested case "NJDEP v. DataOne" | Watch after ~Oct 12–15. DataOne "disagrees" with the generator determination (R), so a request is plausible (E) |
| 13 | **Nebius Q3'26 letter (~Nov 10 est., not company-announced)** and any 6-K before it | https://group.nebius.com/sec-filings | Tranche count delivered; SLA credits; Bloom energization; Vineland revenue | Next disclosure |

**Indirect signals to watch:**
- Resident noise reports. Noise quieted from ~Sep 22 and returned Sep 30 (Inquirer Oct 1, R).
- Floodlight/Guardian drone thermal follow-ups.
- Bloom Q3 results (early Nov) commentary on the "first deployment".


## 1b. Oct 5 PM Bloom check (T5–T9 delivery / install evidence)
_Search completed Mon Oct 5, 2026, ~12:50 PM ET. Public sources only; NJDEP DataMiner/DocMiner/BOSS still Incapsula-blocked from this box (captcha; `njdep_dataminer_captcha_wall.png`). No permit numbers invented._

### Bloom fuel-cell delivery / install timing — **none found that advances T5–T9**
| Check | Result | Tag | Source |
|---|---|---|---|
| Web: "Bloom Energy Vineland", "DataOne fuel cell", "Bloom Vineland delivery", "Bloom Power Module Vineland" | Only previously known: May 14/20 2026 Nebius–Bloom agreement (~328 MW installed / ~250 MW guaranteed, three 10-yr phases, "operational during 2026"); Aug 17 Planning Board amended site plan with Bloom Power areas; Aug 10 stop-construction order on Bloom units | R (prior) | https://nebius.com/newsroom/nebius-and-bloom-energy-partner-to-power-ai-infrastructure-build-out ; WHYY / ENR / Hunterbrook |
| Bloom Q2'26 earnings release (Jul 28 2026) | Record revenue / raised FY26 guide; **no Nebius, DataOne, or Vineland named**; no site-level COD | R absence | https://investor.bloomenergy.com/press-releases/press-release-details/2026/Bloom-Energy-Reports-Record-Second-Quarter-2026-Financial-Results-and-Raises-Full-Year-2026-Guidance/default.aspx |
| IDF / Oaktree $1.7B Bloom financing for Nebius (Jul 16 2026) | Project finance for Bloom fuel cells serving Nebius AI build-out; **no Vineland COD, module delivery date, or install progress** | R | https://www.prnewswire.com/news-releases/industrial-development-funding-and-oaktree-announce-1-7-billion-project-investment-in-bloom-energy-fuel-cells-for-nebius-ai-infrastructure-build-out-302827566.html |
| Bloom Q2'26 call commentary (secondary) | Offtake / IDF takes delivery even if projects delay (secondary transcript summaries); **no public Vineland install milestone** | E (secondary) | Seeking Alpha / Roic AI transcript pages (fetch timed out) |
| Stop-construction order Aug 10 lift | **No public record of lift** as of Oct 5 PM (same as AM). Site plan approved Aug 17; UCC / CO still required to resume Bloom work | R absence | Hunterbrook https://hntrbrk.com/breaking-news/vineland ; WHYY stop-work articles |
| NJDEP Bloom / fuel-cell PCP or BOSS public notice for PI 75894 | **None found** via web search; DataMiner/DocMiner/BOSS captcha-blocked. DataOne claim "necessary approvals in place" for fuel cells (R, WHYY Sep 23) still unverified in public DEP notices | R absence | BOSS https://dep.nj.gov/boss/public-notices/ (blocked); WHYY Sep 23 |
| Vineland Planning Board / construction | Oct 14 meeting still scheduled 6:30 pm; **Oct 14 agenda not posted** (archive latest = Agenda 9-9-26). Posted hearing notice for Oct 14 is **Take 5 Auto Service**, not DataOne | R | https://www.vinelandcity.org/Archive/?dir=Planning%20Board/Agendas/2026 ; https://www.vinelandcity.org/wp-content/uploads/2026/09/Notice-Of-Public-Hearing-Take-5-Post-10.3.pdf |
| Local press (Hunterbrook, WHYY, Inquirer, Sustain SJ, Floodlight) since ~Sep 22 | Coverage of genset fine / 45-day clock; **no article reporting Bloom module delivery, energization, or stop-order lift** | R absence | Floodlight Sep 23; WHYY; ENR Sep 23 |

**Verdict (E, high confidence on absence):** No public evidence that Bloom modules have been delivered or installed for Vineland T5–T9 since the Aug 10 stop-work and Aug 17 site-plan approval. Case B Bloom-slip (−~$966M 2027 rev) probability **unchanged** at ~35% (no positive COD signal; stop-order still unlifted in public record).

### Genset Nov 6 risk — filing status update
| Item | Status Oct 5 PM | Tag |
|---|---|---|
| AONOCAPA EA ID PEA260001-75894 (Sep 22) | Unchanged; $1,072,000; 62 × CAT G3516H; 45 days to **submit** PCP + Certificate to Operate applications or cease | R |
| Deadline | **~Nov 6 2026** (Nov 9 if mailed); receipt date still not disclosed | D |
| Public filing of PCP / Certificate applications since Sep 22 | **None found** in press or searchable DEP notices. DataOne said it "**will** apply" (WHYY/ENR/Floodlight Sep 23) — future tense; no follow-up reporting a submittal | R absence |
| DataMiner / DocMiner PI 75894 | Still captcha-blocked from this box — cannot confirm or invent activity/permit numbers | R (block) |
| Hearing-request window | Still ~Oct 12–15 (20 days after receipt) | D |

**Nov risk:** Unchanged. Without a visible application by ~Nov 6–9, cease-operation order on gensets remains live risk for T1–T4 (~107 IT MW ceiling). Partial-curtailment / consent-order path still the base expectation if they file on time (E).

### Implications for Case B (−~$966M 2027 rev)
- **No change** to Case B probability from this PM check.
- What would cut Case B odds: public lift of Aug 10 Bloom stop-order; Bloom/NJDEP PCP notice; Nebius or Bloom disclosure of module delivery / first-fire / COD; satellite/thermal evidence of fuel-cell arrays (not found in prior Oct 1 imagery note).
- What would raise Case B odds: explicit DEP or city statement that Bloom work remains barred; Bloom Q3 commentary of delay; missed genset filing forcing dual power-path stress.

## 2. 2027 dependence on Bloom tranches T5–T9
**Tranche MW size search (Oct 5, 2026 AM) — result: still Estimated.**
Searched Nebius FY24/FY25 20-F text (`/workspace/nebius/src/20f.txt`), Q1'26 and Q2'26 shareholder letters, and the Q2'26 call transcript (`q2t.txt`). Disclosures state the Microsoft GPU services are deployed in **nine tranches during 2025 and 2026** (R) and name delivery dates for early tranches (T1 Nov 2025, T2 Feb 2026; two further tranches in the month before the Q2 letter) (R), but **no filing or transcript discloses megawatts per tranche or that tranches are equal**. Therefore **240 IT MW ÷ 9 ≈ 26.7 MW IT per tranche remains Estimated (E)**; revenue-at-risk math below inherits that assumption. If early tranches are smaller (suggested by the Q2 top-customer 24% revenue cross-check in `site_verification.md`), T5–T9 could be larger than 26.7 each — treat 26.7 as a working midpoint, not a Reported figure.

**Assumptions:**
- Equal split of 240 MW IT across 9 tranches gives **26.7 MW IT per tranche** (E; tranche sizes not disclosed).
- Microsoft rate $17,392.9M ÷ 5 yrs ÷ 240 MW IT = **$14.5M per IT MW per year** (D), so **$386.5M per tranche per year** (D).
- Billing starts on the first day of the base billing quarter (nebius_capacity.xlsx Vineland tranche track).
- Share of IT MW uses YE27 IT-normalized base energized of **1,082 MW** (v2 model, D).
- Revenue share uses the YE26 ARR guide of **$7–9B** (R, Q1'26 letter, reiterated Q2'26) as the run-rate proxy. There is no 2027 revenue guide.

| Tranche | MW IT | Energized (base) | Billing start base (bear / bull) | 2027 revenue, base ($M) | 2027 revenue if Bloom slips 2Q ($M) | Share of YE27 Nebius IT MW | Annual revenue as share of $7–9B ARR |
|---|---|---|---|---|---|---|---|
| T5 | 26.7 | Q1'27 | Q2'27 (Q3'27 / Q1'27) | 289.9 | 96.6 (Q4'27) | 2.5% | 4.3–5.5% |
| T6 | 26.7 | Q1'27 | Q2'27 (Q3'27 / Q1'27) | 289.9 | 96.6 (Q4'27) | 2.5% | 4.3–5.5% |
| T7 | 26.7 | Q2'27 | Q3'27 (Q4'27 / Q2'27) | 193.3 | 0 (Q1'28) | 2.5% | 4.3–5.5% |
| T8 | 26.7 | Q2'27 | Q3'27 (Q4'27 / Q2'27) | 193.3 | 0 (Q1'28) | 2.5% | 4.3–5.5% |
| T9 | 26.7 | Q2'27 | Q3'27 (Q1'28 / Q2'27) | 193.3 | 0 (Q1'28) | 2.5% | 4.3–5.5% |
| **T5–T9** | **133.3** | | | **1,159.5** | **193.3** | **12.3%** | **21–28%** ($1,932.5M per year) |

**Formulas:**
- 2027 revenue per tranche = 26.7 × $14.5M × (months billed in 2027 ÷ 12).
- 2027 revenue at risk from a 2-quarter Bloom slip = 1,159.5 − 193.3 = **≈ $966M** (D).
- Energized-MW effect of the slip: Vineland is −86.7 MW IT at Q2'27 and −66.5 MW IT at YE27 vs base (nebius_capacity.xlsx Quarterly MW, "SCENARIO: Bloom slips 2 quarters").
- Billing-MW effect of the slip: −80 MW at YE27 (D).

**Dependencies:**
- Bloom supply: ~328 MW installed / ~250 MW guaranteed in three phases. Nebius buys the power under 10-year phases, fees up to $2.6B (R, Nebius 6-K via measuredai). Mapping Bloom phases to tranches is **not disclosed**.
- 250 MW guaranteed plus 5 MW grid ≈ 255 MW firm supply. At PUE 1.2 that supports ~212 MW IT, below the 240 MW IT cap (D). The last ~1 tranche may rely on non-guaranteed Bloom headroom (E, low).
- Microsoft holds service-credit and per-tranche termination rights for delivery delays (R, 20-F).

## 3. What would change the view
- **Filing appears in DataMiner / DocMiner by ~Nov 6:** the full-shutdown risk drops sharply. Likely next step is a consent order letting gensets run until Bloom, with an NOx cap. Partial-curtailment risk remains.
- **No filing, or DEP enforces "obtain":** the full-shutdown scenarios in vineland_shutdown.xlsx apply (−102.5 MW IT; $98–179M Q4'26 revenue at risk at $11.6M per MW).
- **Bloom PCP notice published:** fuel cells need a 30–60 day comment period, and possibly an EJ hearing, before an operating certificate. That puts T5–T9 at risk of the Bloom-slip case.
