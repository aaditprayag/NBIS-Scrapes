Peer-site alt-data scripts. Copied from /workspace/nebius/capacity/altdata/scripts with OUT=/workspace/nebius/peers/sites; methods are unchanged.
Python: /workspace/nebius/.venv_geo/bin/python
1. sites.py: AOIs. crops_extra.json: tight campus crops [W,S,E,N]
2. s2_timeline.py <site>: Sentinel-2 quarterly pass. Peer addition: S2_START=YYYY-MM-DD S2_CADENCE=M runs a denser pass and merges it into metrics.json
3. detail.py <site>: crop metrics + strips/<site>_strip.png
4. strip_full.py <site>: full-AOI strip, written to strips/<site>_full_strip.png
5. landsat_thermal.py <site>: Landsat ST anomaly, written to thermal/
6. preview_grid.py <site> <date>: lat/lon grid helper for locating crops
7. build_sites.py: builds sites.xlsx and sites_summary.json
