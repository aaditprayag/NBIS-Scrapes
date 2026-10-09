# Nebius site tracker — Wednesday handoff
_As of Mon Oct 5, 2026, ~10:30 AM ET (colo gap pass + owned power path). Public sources only._

**Files:** `site_tracker_wed.xlsx` · `owned_campus_power_path.md` · source model `nebius_capacity.xlsx` v2 · lock note `vineland_watch.md` §0c · tranche note `vineland_watch.md` §2.

## Company totals (IT-normalized energized; Derived from capacity v2)

| | YE26 | YE27 | YE28 |
|---|---:|---:|---:|
| **Energized IT base** | **301.7** | **1,081.9** | **1,652.8** |
| Energized IT low | 189.5 | 794.6 | 1,407.6 |
| Energized IT high | 414.0 | 1,369.2 | 1,897.9 |
| Connected base (gross/shell, E) | 941 | 1,882 | n/d |
| Billing workbook YE exit (IT, D) | 257.7 | 852.6 | 1,344.0 |

## Locked billing (Modeling / Skeptic — do not change)

| Path | YE26 billing MW | FY27 **average** billing MW |
|---|---:|---:|
| **Base** | **~258** | **~475** |
| **High** | **~284** | **~645** |

- Method: 1-quarter lag from capacity v2 energized / Vineland billed-tranche timing (`vineland_watch.md` §0c).
- YE26 base cross-check: workbook billing row = 257.7 ≈ 258.
- **Drop** ~350 / ~710 shortcuts (old YE-energized × 80–85%).
- **Microsoft tranche MW:** still **Estimated** at 26.7 IT (240÷9); no per-tranche MW in 20-F / Q1–Q2 letters / Q2 call (`vineland_watch.md` §2).
- Do **not** use YE27 exit billing (853) or the raw Q1–Q4'27 mean (~594) as the FY27 average — use the locked **475 / 645**.

## Sites with gaps (updated Oct 5 ~10:30 AM ET — re-search complete)

All five colo gaps were re-searched (landlord PRs, Nebius letters/maps, docs.nebius.com, utility/planning, secondary trackers). **Nebius absolute announced MW remains n/d for every one.** Facility/category leads are documented; none change the locked billing path.

| Site | Announced MW finding | Landlord / lead | Tag | Still missing? |
|---|---|---|---|---|
| Newport UK (Vantage CWL1) | Nebius lease MW **n/d** (R). Campus ≤~148 MW critical IT is Vantage's rating, not Nebius share. UK residual **43 MW** = 65−22 Harlow for Newport+UK3 (D) | Vantage CWL1 (R, Aug 13 2026 PR) | n/d / D | YES (Nebius MW) |
| UK 3rd (GM LON-East lead) | **14 MW** leased to *unnamed* neocloud (R, GM Sep 1). Nebius **not named** | Green Mountain LON-East, Romford | E (lead) | YES (identity + whether Nebius) |
| Minnesota (Woodbury / us-north1) | Nebius tenant MW **n/d**. IronGate MSP01: **15 MW live / 66 MW full**, fully leased (R). Nebius-at-IronGate = inference | IronGate lead (E) | E / R (facility) | partial |
| Oklahoma colo | Map category **">100 MW"** only (R). Absolute MW, town, landlord, OCC/SPP docket: **none found** | n/d | R (category) | YES |
| France 2nd (eu-west2) | Region live B300 (R). Physical site / landlord / MW: **n/d**. Distinct from Equinix PA10 | n/d | n/d | YES (+ no energized path) |

**Missing ownership disclosure (1):** UK 3rd — Sites tab "not disclosed" (Basis treats as leased if lead confirmed).

**Billing lock:** unchanged — Base ~258 YE26 / ~475 FY27 avg; High ~284 / ~645. No colo MW discovery justified a revision.

## Power path summary — owned campuses (see `owned_campus_power_path.md`)

| Campus | Utility | Key feed / COD | YE26 connected meaning | First credible IT | Grade |
|---|---|---|---|---|---|
| Independence MO | IPL / IPP Blue Valley (+ NextEra/Evergy interim) | 200 MW utility late Q3/Oct 2026 (R); Blue Valley Ph1 ~200–250 MW Oct 2027 (R) | Power feed + shell, **not** IT | ~Q3 2027 | Green |
| Butler Twp PA | PPL (1.2 GW ESA) | 230 kV switchyard + 2 laterals: build late 2027–2028/29 (R). Mgmt "260 MW Oct 2027" **not supported** | Shell/earthworks only | Base ~2028 | **Red** |
| Birmingham AL | Alabama Power | 115 kV NLT Dec 31 2026; 230 kV Oct 2027 (R). 32×400 kW gas = commissioning only | Feed + partial shell | into 2027 (litigation Mar 2027) | Yellow |

None of the three owned campuses support locked YE26 billing.

## So-what for Modeling (one paragraph)

Use the locked billing path (~258 / ~475 base; ~284 / ~645 high), not YE energized exits and not the old ~350/~710 shortcuts. YE26 energized IT base is ~302 MW but only ~258 MW bills by year-end after the 1Q lag and Vineland tranche timing; YE27 energized IT base (~1,082) is an exit stock — FY27 revenue should run off ~475 MW average billing (base), which is ~56% of the YE27 exit and implies the ramp is still back-half loaded. High path still sits ~19% below the ~769 MW that would be needed to clear $16M/MW comps at the high ARR. Biggest site-level swings into YE27: Vineland Bloom T5–T9 (~133 IT MW), Oklahoma/Bethune/Beit Shemesh lease ramps, and owned campuses (Independence / Birmingham / Lappeenranta) that barely contribute until 2H27–2028. Five colo sites still lack disclosed MW (Newport, UK3, MN, OK, France 2nd), so the IT total is a named-site floor plus those estimates.

## Tab map (`site_tracker_wed.xlsx`)

| Tab | Contents |
|---|---|
| README | Build stamp, standards, locked numbers, YE IT snapshot |
| By_Site | 24 sites × YE26/27/28 low/base/high, IT-norm est., ownership, GPU, customer, contracted vs unsold, grade, Q3'26–Q4'28 base path, missing flags |
| Company_Totals | IT / gross / connected / stated / billing series through Q4'28 + YE snapshot |
| Billing_Lock | Hardcoded ~258/284/475/645 + dropped shortcuts + lag explanation |
| Sources | File and disclosure map |

Tags: **R** Reported · **D** Derived · **E** Estimated. Blue cells = locked/hardcoded inputs.
