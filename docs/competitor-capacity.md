# Competitor capacity near each Nebius campus

`neighbors.csv` answers which buildings sit within 60 km. This file answers how many megawatts sit near each campus, and when they come online. It goes out to 200 km, by phase, with the quarter each phase comes online and the distance and round trip to each campus.

The build is `scripts/build_competitors.py`. The research is `data/competitor_projects.jsonl`. The outputs are `data/out/competitor_phases.csv`, `data/out/competitors.csv`, `data/out/wide/competitors_ring1.csv`, and `data/out/wide/competitors_ring2.csv`.

## What was collected

Every other AI, cloud, and colocation project within about 200 km of the 22 Nebius campuses: hyperscalers, neoclouds, colocation and wholesale operators, and national AI projects. Russia and Belarus are left out. Within 50 km, every building with a public megawatt figure is in, including halls that were running before 2024. From 50 to 200 km, every campus of 10 MW or more is in, plus every hyperscaler, neocloud, or AI project of any size.

Each record is one phase: a building, a hall, or a power block. A phase has its megawatts, the basis of that number (IT, utility, or unstated), the quarter it comes online, and the status (operating, under construction, permitted, announced). Each megawatt figure and each date has a source URL and a short quote from the page. Sources are company filings and releases, local permits and council papers, utility filings, DCD and local press, and operator facility pages. Baxtel and datacentermap are used only when nothing primary was found.

The research came in 20 slices, by campus or region. Joined, there are 698 records. 45 are the same project id in two slices, and two more are the same project under a second name, within 1 km and with the same phases. That leaves 651 projects and 883 phases. 283 projects have at least one phase near a campus with both megawatts and a date.

No megawatt figure is a guess. If a range or a floor is all that is public, the floor is used and the basis says so. If the megawatts or the date are not public, the phase stays in the CSVs with `not public` and is left out of the totals.

## Rings

Ring 1 is within 50 km of the campus. Ring 2 is 50 to 200 km. Distance is a straight line. `rtt_ms` is the distance divided by 100: light in fiber, there and back, on a straight path. Real fiber routes are longer, so it is a floor.

Campus pins are the ones in `data/ledger.json`, the same pins as the map. Most are city level, so a distance is to the city pin, not to the hall. Oklahoma City has no pin in the ledger and Hyderabad is not in it. Those two use the city centre. The research itself was gathered around address-level pins. The two sets of pins are mostly within 10 km of each other. Three are further apart: Paris (10 km, the ledger pin is in the city, PA10 is in Saint-Denis), Hüüru (12 km, the ledger pin is Tallinn), and Spartanburg County (29 km, the ledger pin is county level). Near the 200 km edge of those three, a project could be missing.

## Columns

`competitor_phases.csv` is one row per phase, with the operator, owner, tenant, coordinates and their basis, the megawatts and online quarter with their basis, and the source URL and quote for each.

`competitors.csv` is one row per campus and phase within 200 km. A phase near two campuses is on two rows. `counted` is `yes`, or the reason it is not in the totals: `MW not public`, `no online date public`, `online after Q4 2030`.

The wide files have one row per campus, quarter-end dates across, and the fiscal quarter on the row above, from Q1 2024 through Q4 2030. A cell is the megawatts online by that quarter end. A phase that was running before 2024 counts from Q1 2024. A year with no quarter is Q4 of that year. A 0 means no counted phase in that ring by that date. The rings around one campus never overlap, but neighboring campuses share projects, so do not add campuses together.

## What it shows

Megawatts online, end of year:

| Campus | Ring 1, 2026 | Ring 1, 2028 | Ring 1, 2030 | Ring 2, 2026 | Ring 2, 2030 |
|---|---:|---:|---:|---:|---:|
| Paris (PA10) | 695 | 1,477 | 2,485 | 0 | 600 |
| Longcross | 809 | 1,190 | 2,323 | 290 | 930 |
| Harlow | 388 | 661 | 2,254 | 639 | 927 |
| Kansas City | 364 | 1,397 | 1,397 | 0 | 0 |
| Getafe | 272 | 570 | 1,168 | 0 | 426 |
| Hyderabad | 407 | 687 | 827 | 0 | 0 |
| Independence | 364 | 797 | 797 | 0 | 600 |
| Oklahoma City | 228 | 398 | 398 | 0 | 0 |
| Woodbury | 368 | 368 | 368 | 458 | 608 |
| Masmiyya, Beit Shemesh | 102 | 200 | 330 | 67 | 168 |
| Butler Township | 0 | 350 | 350 | 726 | 1,096 |
| Mäntsälä | 64 | 284 | 284 | 165 | 635 |
| Modi'in | 114 | 154 | 284 | 55 | 214 |
| Keflavík | 256 | 256 | 256 | 0 | 0 |
| Vineland | 0 | 120 | 120 | 763 | 1,363 |
| Newport | 77 | 117 | 117 | 677 | 1,898 |
| Birmingham | 20 | 95 | 95 | 314 | 869 |
| Hüüru | 10 | 10 | 10 | 213 | 903 |
| Béthune | 2 | 2 | 2 | 570 | 2,226 |
| Lappeenranta | 0 | 0 | 0 | 170 | 750 |
| Spartanburg County | 0 | 0 | 0 | 72 | 1,062 |

Longcross, Harlow, and Paris have the most dated capacity within 50 km, and it grows three to six times by 2030. Lappeenranta, Béthune, Hüüru, and Spartanburg County have almost nothing within 50 km. Their competition is 50 to 200 km away. Vineland and Butler Township are the same: nothing dated within 50 km before 2027, and 0.7 GW or more in ring 2 by the end of 2026.

Late years are thin. 359 of the 883 phases have no public online date and 236 have no public megawatts, and none of those are in the totals. The 2029 and 2030 columns are a floor.
