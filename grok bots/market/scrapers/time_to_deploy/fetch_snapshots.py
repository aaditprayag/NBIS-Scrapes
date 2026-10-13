"""Fetch specific Wayback captures (raw, id_ modifier) into /workspace/nebius/wayback_raw/<page>/<ts>.html, ~1s between requests.
Reuses /workspace/nebius/src/wb/fetch.py (shared cache with the price-history workstream). Usage: python fetch_snapshots.py URL TS [TS ...]"""
import sys, time
sys.path.insert(0,'/workspace/nebius/market/scrapers/wayback')
from fetch import fetch
url=sys.argv[1]
for ts in sys.argv[2:]:
    fn,st=fetch(ts,url); print(ts,st,fn,flush=True); time.sleep(1)
