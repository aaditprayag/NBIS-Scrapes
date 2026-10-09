#!/usr/bin/env python3
"""Landsat 8/9 Collection-2 L2 surface temperature (ST_B10) anomaly: campus crop minus a 3 km background
box (campus excluded), using Microsoft Planetary Computer STAC (free, anonymous SAS token; no login).
ST_B10 native 100 m TIRS (delivered at 30 m) => only campus-scale (>~10 ha) signals; daytime ~10:30-11:00
local overpass, so solar heating of roofs/bare sand dominates; NOT a direct power meter. Clear-sky only
(QA_PIXEL cloud/shadow bits masked). Usage: landsat_thermal.py site"""
import sys, os, json, csv, math, requests, warnings
warnings.filterwarnings('ignore')
import numpy as np, rasterio
from rasterio.warp import transform_bounds
from rasterio.windows import from_bounds
OUT='/workspace/nebius/peers/sites'
sys.path.insert(0,os.path.dirname(__file__)); from detail import CROPS
API='https://planetarycomputer.microsoft.com/api/stac/v1'
import time
def token():
    for i in range(8):
        try: return requests.get('https://planetarycomputer.microsoft.com/api/sas/v1/token/landsat-c2-l2',timeout=60).json()['token']
        except Exception: time.sleep(5)
    raise SystemExit('Planetary Computer SAS token endpoint unavailable')
tok=token()
def rd(href,bb):
    with rasterio.open(href+'?'+tok) as ds:
        w=from_bounds(*transform_bounds('EPSG:4326',ds.crs,*bb),ds.transform); return ds.read(1,window=w,boundless=True,fill_value=0)
site=sys.argv[1]; c=CROPS[site]; cx=(c[0]+c[2])/2; cy=(c[1]+c[3])/2
d=1500/111320; bg=[cx-d/math.cos(math.radians(cy)),cy-d,cx+d/math.cos(math.radians(cy)),cy+d]
feats=[]; body={'collections':['landsat-c2-l2'],'bbox':list(c),'datetime':'2023-01-01/2026-10-04','limit':200,
 'query':{'eo:cloud_cover':{'lt':30},'platform':{'in':['landsat-8','landsat-9']}}}
r=requests.post(API+'/search',json=body,timeout=120).json(); feats=r['features']
rows=[]
for f in sorted(feats,key=lambda f:f['properties']['datetime']):
    try:
        st=rd(f['assets']['lwir11']['href'],bg).astype('float32'); qa=rd(f['assets']['qa_pixel']['href'],bg)
    except Exception as e: continue
    ok=(st>0)&((qa&0b11111)==0)   # fill, dilated cloud, cirrus, cloud, shadow bits clear
    if ok.mean()<0.9: continue
    T=st*0.00341802+149.0-273.15
    H,W=T.shape; x0=int((c[0]-bg[0])/(bg[2]-bg[0])*W); x1=int((c[2]-bg[0])/(bg[2]-bg[0])*W); y0=int((bg[3]-c[3])/(bg[3]-bg[1])*H); y1=int((bg[3]-c[1])/(bg[3]-bg[1])*H)
    m=np.zeros_like(ok); m[y0:y1,x0:x1]=True
    if (ok&m).sum()<20: continue
    rows.append(dict(site=site,date=f['properties']['datetime'][:10],id=f['id'],campus_C=round(float(T[ok&m].mean()),2),
                     background_C=round(float(T[ok&~m].mean()),2),anomaly_C=round(float(T[ok&m].mean()-T[ok&~m].mean()),2)))
    print(rows[-1]['date'],rows[-1]['campus_C'],rows[-1]['anomaly_C'])
os.makedirs(f'{OUT}/thermal',exist_ok=True)
with open(f'{OUT}/thermal/{site}_landsat_st.csv','w',newline='') as fh:
    w=csv.DictWriter(fh,fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt, datetime as dt
x=[dt.date.fromisoformat(r['date']) for r in rows]; y=[r['anomaly_C'] for r in rows]
plt.figure(figsize=(9,3.2)); plt.plot(x,y,'o-',ms=3); plt.axhline(0,c='k',lw=.5)
plt.title(f'{site}: Landsat 8/9 ST_B10 campus-minus-background (°C), daytime, clear-sky'); plt.ylabel('°C'); plt.tight_layout()
plt.savefig(f'{OUT}/thermal/{site}_landsat_st.png',dpi=130)
