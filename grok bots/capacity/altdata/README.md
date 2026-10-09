# Nebius alt-data: physical build & energization dating (as of 2026-10-04)

Public/free data only. No logins, no paid services, no contact with anyone.

## Outputs
- `altdata_timelines.xlsx`: **Timeline** (site, date, evidence type, observation, source, confidence; 80 rows), **Vineland_MW** (power-bounded MW estimate; uses formulas), **Imagery_metrics** (tight-crop roof/bare metrics), **Imagery_AOI_metrics**, **Thermal**, **Sites_AOI** (coordinates and where each came from).
- `imagery/<site>/`: true-colour PNG chip for each selected scene (`<site>_<date>.png`, 10 m pixels upsampled 4x) plus `.npz` index arrays and `metrics.json`. `imagery/<site>/detail/`: tight campus crops (8x upsampled).
- `strips/<site>_strip.png`: deck-ready time strips. The ones that show clear change are **vineland_nj, birmingham_al, bethune_fr, independence_mo, mantsala_fi, butler_pa**. The keflavik, longcross and kansascity strips show no interpretable Nebius-specific change. `_UNLOCATED_lappeenranta_fi_strip.png` is a city-centre AOI: the campus was not located, so don't use it.
- `thermal/vineland_nj_landsat_st.{csv,png}`: Landsat 8/9 surface-temperature anomaly.
- `docs/`: downloaded primary documents (Q2'26 shareholder letter).

## Methods
1. **Scene selection** (`scripts/s2_timeline.py`). Queries the Element84 Earth Search STAC (`sentinel-2-l2a`) for each AOI in `scripts/sites.py`, by quarter (monthly for Vineland), 2023-01 to 2026-10-04. Candidates are ranked by scene cloud cover. For each, the SCL band is read over the AOI only, and the first scene with <5% cloud/shadow, <20% snow and <5% nodata is kept. COG windows are read directly over HTTPS (no download, no login). Per-AOI metrics: NDVI, NDBI, BSI, built/bare ha and bright-roof ha. AOI-wide metrics are dominated by farmland seasonality, so don't use them for change dating.
2. **Tight-crop metrics** (`scripts/detail.py`, crops in-script and in `scripts/crops_extra.json`). Re-reads true reflectance for each selected scene over a campus crop and computes:
   - `roof_white_ha`: NDVI<0.15, mean RGB>0.18, blue/red>0.80. The neutral-colour test excludes yellow sand.
   - `roof_grey_ha`
   - `bare_ha`: earthworks.
   Each crop also gets a stretched 8x PNG and a strip. Compare same-season dates: sun angle and winter shadow move the numbers by ±1-2 ha.
3. **Full-AOI strips** (`scripts/strip_full.py`) for sites without a tight crop.
4. **Thermal** (`scripts/landsat_thermal.py`). Earth Search's `landsat-c2-l2` assets sit in the requester-pays `usgs-landsat` bucket, which needs AWS credentials. So Microsoft Planetary Computer STAC is used instead (anonymous SAS token, free). ST_B10 is converted to °C. Value = campus crop mean minus a 3 km background box, clear-sky only (QA_PIXEL bits 0-4).
   - **Limits:** native TIRS resolution is 100 m. The ~10:30 am overpass means solar heating of roofs and sand dominates. Not a power meter. Treat as low confidence.
5. **Timeline build** (`scripts/build_timeline.py`): rebuilds the workbook. Evidence rows are hard-coded with sources.
6. **Coordinates.** OSM Nominatim geocodes of addresses from permits and filings: 3963 S Lincoln Ave Vineland (Block 7503 Lot 35.01); Moreenikatu 6 Mäntsälä; 1601 McGee St KC; 201 Milan Pkwy Birmingham; the Bridgestone/SIG plant in Béthune. Independence and Butler Twp were located visually from the earthworks, so their parcels aren't confirmed.

Rerun: `../../.venv_geo/bin/python scripts/s2_timeline.py vineland_nj && ../../.venv_geo/bin/python scripts/detail.py vineland_nj && ../../.venv_geo/bin/python scripts/build_timeline.py`. The venv is at `/workspace/nebius/.venv_geo`: rasterio, pystac-client, Pillow, openpyxl, matplotlib.

## Not measurable / not done
- Interior fit-out, GPU install and energization are invisible at 10 m. Imagery only shows earthworks, roofs and large equipment arrays.
- Paris (Equinix PA10) and KC (Patmos interior retrofit) are colocation inside existing buildings, so satellite can't resolve them. Israel sites were not attempted.
- Bloom units and gensets can't be counted at 10 m. Unit counts come from NJDEP and from Floodlight's thermal drone.
- PJM queue: Vineland has no generator interconnection queue position we could find. The project is behind the meter. The ACE 49 MW load withdrawal comes from a secondary source (measuredai substack) and has **not** been verified in PJM TEAC/load-forecast documents.
- Fingrid (Mäntsälä) and Evergy (KC) filings were not found or checked.
- ImportYeti limits anonymous use to about 25 page views per IP and then requires a login. Only the free Nebius company page was read; nothing was bypassed.
- The NJDEP order PDF is bot-protected for curl/WebFetch. Its contents are quoted from search snippets and the ENR/Floodlight coverage.
