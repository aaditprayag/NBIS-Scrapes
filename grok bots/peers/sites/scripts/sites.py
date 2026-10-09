# Peer-site AOIs (same method as /workspace/nebius/capacity/altdata/scripts/sites.py). half_m = AOI half-width (m).
# Coordinates: Nominatim/OSM geocodes of addresses from filings, refined visually from Sentinel-2 chips (see basis).
SITES = {
 'denton_tx':      dict(lat=33.2150, lon=-97.2110, half_m=1300, cadence='Q', basis='Nominatim/OSM: Jim Christal Rd west segment (33.2207,-97.2088) + Denton Energy Center (33.2149,-97.2098); campus at 8171 Jim Christal Rd per Core Scientific 8-K 2026-04-21; mining-shed rows located visually W/SW of DEC'),
 'childress_tx':   dict(lat=34.392,  lon=-100.045, half_m=2500, cadence='Q', basis='located visually in a 10 km search AOI (34.40,-100.10): building rows E of Childress near FM 1033 (address 620 FM 1033 per ColoMap/Texas AI Docket)'),
 'kvandal_no':     dict(lat=68.565,  lon=17.565,  half_m=3000, cadence='Q', basis='approx: north Bjerkvik, NW side of E6 (Kvandal North industrial plan, 345.no 2024-04); refined visually'),
 'abilene_tx':     dict(lat=32.5029, lon=-99.7834, half_m=1800, cadence='Q', basis='Nominatim/OSM: Lancium Clean Campus, Abilene TX'),
 'kcairworld_mo':  dict(lat=39.2960, lon=-94.6762, half_m=600,  cadence='Q', basis='Nominatim: 11155 N Airworld Dr, Kansas City MO (Lambda Project Blitz; ex-Bank of America DC)'),
 'portwash_wi':    dict(lat=43.435,  lon=-87.875,  half_m=2500, cadence='Q', basis='approx: north Port Washington between I-43 / Dixie Rd / Interurban Trail; tour address 1374 Lake Dr geocodes 43.4420,-87.8546; refined visually'),
 'monarch_wv':     dict(lat=38.9143, lon=-82.1009, half_m=2500, cadence='Q', basis='WVDEP draft air permit R14-0042 (2026-09-29): 30 Monarch Way, Ohio River Rd (WV-62) & Airport Rd, Point Pleasant, Mason County WV; Lat/Lon 38.914258/-82.100935; former Mason County Airport runway'),
}
