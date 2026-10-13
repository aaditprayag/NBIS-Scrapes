# Site AOIs for Nebius alt-data work. Coordinates: geocoded address points (Nominatim/OSM)
# or site_list.csv approximations; 'basis' records which. half_m = AOI half-width in metres.
SITES = {
 'vineland_nj':   dict(lat=39.4282, lon=-75.0184, half_m=1200, cadence='M', basis='Nominatim: 3963 S Lincoln Ave, Vineland NJ (Block 7503 Lot 35.01, PILOT ordinance)'),
 'mantsala_fi':   dict(lat=60.6289, lon=25.2620, half_m=1500, cadence='Q', basis='Nominatim: Moreenikatu 6, Mantsala (Nebius DC Oy registered visiting address); AOI also covers Kapuli area ~1 km'),
 'kansascity_mo': dict(lat=39.0940, lon=-94.5794, half_m=400,  cadence='Q', basis='Nominatim: 1601 McGee St (Patmos / former KC Star press)'),
 'birmingham_al': dict(lat=33.4291, lon=-86.8851, half_m=800,  cadence='Q', basis='Nominatim: 201 Milan Pkwy, Birmingham AL'),
 'independence_mo': dict(lat=39.103, lon=-94.308, half_m=1600, cadence='Q', basis='visually located: active earthworks NE of Little Blue Pkwy / MO-78 interchange in S2 chips (Oct 2026); Nebius/City FAQ location description'),
 'butler_pa':     dict(lat=40.735, lon=-76.300, half_m=2500, cadence='Q', basis='approx: Highridge Park Rd (OSM 40.7346,-76.3022) / Gordon Mountain Rd (OSM 40.7409,-76.3251) near I-81; wide AOI'),
 'bethune_fr':    dict(lat=50.5397, lon=2.6590,  half_m=900,  cadence='Q', basis='Nominatim/OSM: Bridgestone plant, Bethune'),
 'keflavik_is':   dict(lat=63.9716, lon=-22.5734, half_m=1000, cadence='Q', basis='Nominatim: Asbru campus (Verne); approx'),
 'lappeenranta_fi': dict(lat=61.06, lon=28.19, half_m=2000, cadence='Q', basis='approx: site_list city-level (Pajarila)'),
 'longcross_uk':  dict(lat=51.3806, lon=-0.5946, half_m=600,  cadence='Q', basis='Nominatim/OSM: Ark Longcross Park'),
}
