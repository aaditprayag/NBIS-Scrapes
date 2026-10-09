# YE26 active vs connected MW by site
_As of Mon Oct 5, 2026, ~12:50 PM ET. Public sources only; nobody contacted._
_Tags: **R** Reported · **D** Derived · **E** Estimated. Companion: `ye26_active_vs_connected.xlsx`._

## Company anchors (do not contradict locks)
| Metric | Value | Tag | Source |
|---|---|---|---|
| YE26 active stated/mixed (identified) | 319.5 | D | `nebius_capacity.xlsx` Quarterly MW R126 |
| YE26 active IT-normalized base / high | 301.7 / 414.0 | D | Company_Totals / Billing_Lock |
| YE26 connected L / B / H | **561 / 941 / 1,195** | E | Quarterly MW R141–143; `unseen_capacity.md` |
| Mgmt connected guide YE26 | 800–1,000 | R | Q1/Q2'26 letters; Q2'26 call |
| LOCKED YE26 billing base / high | ~258 / ~284 | D | `vineland_watch.md` §0c |
| LOCKED FY27 avg billing base / high | ~475 / ~645 | D | same |

**One-line for judges — connected:** Nebius's disclosed metric = **power delivered into fully built and equipped data-center halls** (shell + fit-out ready; can sit idle until GPUs/onboarding). It is **not** (a) land+power *contracted* (~5 GW), (b) substation/grid access alone without a built hall, or (c) billing/active IT. (R: Q1'26 letter; Q2'25 letter: includes active plus power "fully provisioned… and can be activated immediately upon GPU installation"; Q3'25: "power connected to built data centers".)

**Companion:** *active* = power consumed by installed operational IT (R). Company YE26 connected guide **800 MW–1 GW** (R); our **941** is an Estimated reconstruction of that company build target, not a Nebius site-level disclosure — site splits stay [E].

## Headline gaps (connected base − active base)
- **Vineland NJ (DataOne)**: gap **218 MW** (active 81.8 → connected 300); grade Yellow/Red
- **Independence MO (Eastgate)**: gap **100 MW** (active 0 → connected 100); grade Green
- **Bethune FR (Azur)**: gap **90 MW** (active 30 → connected 120); grade Yellow
- **Oklahoma colo (unnamed)**: gap **75 MW** (active 25 → connected 100); grade Yellow/Red
- **Birmingham AL (BHM01, Oxmoor)**: gap **50 MW** (active 0 → connected 50); grade Yellow

Largest gaps are **shell / power-path ahead of billing IT** — Vineland full shell vs genset-capped IT; Independence/Birmingham utility-feed plugs with 0 YE26 energized; Béthune mgmt 120 vs model 30; Oklahoma map >100 vs model 25 active.

## Site table (summary)
Full columns (incl. L/H and sources) in the xlsx. Active L/B/H = stated/mixed from capacity model (**D**). Connected = Estimated attribution reconstructing `unseen_capacity.md` §1 so base sums to **941**.

| Site | Active base | Active high | Connected base | Gap | Grade |
|---|---:|---:|---:|---:|---|
| Vineland NJ (DataOne) | 81.8 | 106.6 | 300 | 218.2 | Yellow/Red |
| Independence MO (Eastgate) | 0 | 0 | 100 | 100.0 | Green |
| Butler Twp PA (Highridge) | 0 | 0 | 0 | 0.0 | Red |
| Birmingham AL (BHM01, Oxmoor) | 0 | 0 | 50 | 50.0 | Yellow |
| Bethune FR (Azur) | 30 | 60 | 120 | 90.0 | Yellow |
| Lappeenranta FI (Pajarila) | 0 | 0 | 0 | 0.0 | Green |
| Mantsala FI #1 | 75 | 75 | 75 | 0.0 | Green |
| Mantsala FI #2 (Kapuli) | 0 | 0 | 0 | 0.0 | Yellow |
| Kansas City MO (Patmos) | 20 | 35 | 40 | 20.0 | Green |
| Paris (Equinix PA10) | 5 | 5 | 5 | 0.0 | Green |
| Keflavik IS (Verne) | 10 | 10 | 10 | 0.0 | Green |
| London Longcross (Ark LP01) | 17.5 | 19 | 19 | 1.5 | Green |
| Harlow UK (Kao Data) | 2.2 | 4.4 | 22 | 19.8 | Green |
| Newport UK (Vantage CWL1) | 5 | 10 | 10 | 5.0 | Yellow |
| UK 3rd (lead: Green Mountain LON-East) | 0 | 0 | 5 | 5.0 | Red |
| Modi'in IL (Mega Or) | 8 | 8 | 8 | 0.0 | Green |
| Masmiyya IL (Mega DC) | 5.5 | 11 | 22 | 16.5 | Yellow |
| Beit Shemesh IL (Mega DC) | 0 | 0 | 0 | 0.0 | Yellow |
| Madrid-Getafe ES (Merlin) | 13.5 | 18 | 18 | 4.5 | Green |
| Tallinn EE (Greenergy Huuru) | 11 | 22 | 22 | 11.0 | Yellow |
| Minnesota colo (Woodbury) | 10 | 15 | 15 | 5.0 | Yellow |
| Oklahoma colo (unnamed) | 25 | 40 | 100 | 75.0 | Yellow/Red |
| France 2nd colo (eu-west2 lead) | n/d | n/d | n/d | n/d | Yellow |
| AIB CLT1 (Moore SC) | 0 | 0 | 0 | 0.0 | Red |
| **TOTAL (excl. FR2 n/d)** | **319.5** | **439.0** | **941** | **621.5** | |

Site connected high sum ≈1156 vs company high 1,195 — residual ~39 left in France 2nd n/d / range (not invented).

## Method / caveats
1. Active figures are **Derived** from the locked capacity workbook — do not re-estimate.
2. Connected site splits are **Estimated** (Nebius does not publish per-site connected). Confidence medium-low on sizes; direction (active ≪ connected) is solid.
3. Owned campuses: Independence ~100 and Birmingham ~50 of YE26 connected are **power/shell**, not billing (`owned_campus_power_path.md`).
4. Vineland connected 300 = **gross shell**; active 81.8 = **IT** on gensets. Basis mismatch intentional (same as capacity model).
5. Billing path unchanged by this table (1Q lag + Vineland tranche timing).

## Files
- `/workspace/nebius/capacity/ye26_active_vs_connected.xlsx`
- `/workspace/nebius/capacity/ye26_active_vs_connected.md`
