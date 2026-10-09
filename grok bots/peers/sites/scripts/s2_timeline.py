#!/usr/bin/env python3
"""Sentinel-2 L2A quarterly/monthly timeline per site via Element84 Earth Search STAC (no login).
For each period: search scenes intersecting AOI, rank by scene cloud cover, check AOI-local cloud/snow
with the SCL band, keep the clearest (<5% cloud/shadow, <20% snow). Saves true-color PNG chip + metrics.
Metrics (AOI-wide, valid pixels only): mean NDVI, NDBI=(SWIR16-NIR)/(SWIR16+NIR), BSI,
built/bare fraction = NDVI<0.2 & NDBI>-0.05 (impervious or bare soil), bright-roof fraction =
NDVI<0.15 & mean(B,G,R)>0.25 (white/metal roofs; 10 m pixels). Areas in hectares.
Usage: s2_timeline.py [site ...]"""
import sys, os, json, math, datetime as dt, warnings
warnings.filterwarnings('ignore')
import numpy as np, rasterio
from rasterio.warp import transform_bounds
from rasterio.windows import from_bounds
from pystac_client import Client
from PIL import Image
sys.path.insert(0, os.path.dirname(__file__))
from sites import SITES
os.environ.update(GDAL_DISABLE_READDIR_ON_OPEN='EMPTY_DIR', AWS_NO_SIGN_REQUEST='YES', GDAL_HTTP_MULTIRANGE='YES', CPL_VSIL_CURL_ALLOWED_EXTENSIONS='.tif')
OUT='/workspace/nebius/peers/sites'
API = 'https://earth-search.aws.element84.com/v1'
START, END = dt.date(2023,1,1), dt.date(2026,10,4)
# Peer-sites addition: optional env overrides for a denser recent pass (rows are merged into metrics.json by date)
if os.environ.get('S2_START'): START=dt.date.fromisoformat(os.environ['S2_START'])

def bbox(s):
    dlat = s['half_m']/111320; dlon = s['half_m']/(111320*math.cos(math.radians(s['lat'])))
    return [s['lon']-dlon, s['lat']-dlat, s['lon']+dlon, s['lat']+dlat]

def periods(cad):
    out=[]; y,m=START.year,START.month; step=1 if cad=='M' else 3
    while dt.date(y,m,1)<=END:
        m2=m+step; y2=y+(m2-1)//12; m2=(m2-1)%12+1
        out.append((dt.date(y,m,1), min(dt.date(y2,m2,1)-dt.timedelta(days=1), END))); y,m=y2,m2
    return out

def read(href, bb, shape=None):
    with rasterio.open(href) as ds:
        b = transform_bounds('EPSG:4326', ds.crs, *bb)
        w = from_bounds(*b, ds.transform)
        kw = dict(window=w, boundless=True, fill_value=0)
        if shape: kw.update(out_shape=shape, resampling=rasterio.enums.Resampling.nearest)
        return ds.read(1, **kw).astype('float32')

def refl(item, key, bb, shape=None):
    a = read(item.assets[key].href, bb, shape)
    off = 0 if item.properties.get('earthsearch:boa_offset_applied', True) else (-1000 if float(item.properties.get('s2:processing_baseline','0'))>=4 else 0)
    return (a+off)/10000.0

def process(name):
    s=SITES[name]; bb=bbox(s); cl=Client.open(API); rows=[]
    os.makedirs(f'{OUT}/imagery/{name}', exist_ok=True)
    for p0,p1 in periods(os.environ.get('S2_CADENCE', s['cadence'])):
        its = list(cl.search(collections=['sentinel-2-l2a'], bbox=bb, datetime=f'{p0}/{p1}', query={'eo:cloud_cover':{'lt':60}}, max_items=60).items())
        its.sort(key=lambda i: i.properties.get('eo:cloud_cover',100))
        best=None
        for it in its[:8]:
            try:
                scl = read(it.assets['scl'].href, bb)
            except Exception as e:
                continue
            n=scl.size; nod=(scl==0).sum()/n
            cloud=np.isin(scl,[3,8,9,10]).sum()/n; snow=(scl==11).sum()/n
            if nod>0.05: continue
            if cloud<0.05 and snow<0.2: best=(it,scl,cloud,snow); break
        if not best:
            rows.append(dict(site=name, period=f'{p0}', date=None, note='no clear scene (cloud/snow/darkness)')); print(name,p0,'none'); continue
        it,scl,cloud,snow=best
        r=refl(it,'red',bb); shp=r.shape
        g=refl(it,'green',bb); b=refl(it,'blue',bb); nir=refl(it,'nir',bb)
        sw=refl(it,'swir16',bb,shp); scl=read(it.assets['scl'].href,bb,shp)
        valid=~np.isin(scl,[0,1,3,8,9,10,11])
        ndvi=(nir-r)/(nir+r+1e-6); ndbi=(sw-nir)/(sw+nir+1e-6); bsi=((sw+r)-(nir+b))/((sw+r)+(nir+b)+1e-6)
        vis=(r+g+b)/3
        built=valid&(ndvi<0.2)&(ndbi>-0.05); roof=valid&(ndvi<0.15)&(vis>0.25)
        px_ha=0.01  # 10 m pixel = 100 m2
        d=it.datetime.date().isoformat()
        rgb=np.clip(np.dstack([r,g,b])/0.25,0,1)**(1/1.6)
        im=Image.fromarray((rgb*255).astype('uint8')).resize((shp[1]*4,shp[0]*4), Image.NEAREST)
        fn=f'{OUT}/imagery/{name}/{name}_{d}.png'; im.save(fn)
        row=dict(site=name, period=f'{p0}', date=d, scene=it.id, scene_cloud=it.properties.get('eo:cloud_cover'),
                 aoi_cloud=round(float(cloud),4), aoi_snow=round(float(snow),4), valid_frac=round(float(valid.mean()),3),
                 ndvi_mean=round(float(ndvi[valid].mean()),4), ndbi_mean=round(float(ndbi[valid].mean()),4), bsi_mean=round(float(bsi[valid].mean()),4),
                 builtbare_ha=round(float(built.sum()*px_ha),2), brightroof_ha=round(float(roof.sum()*px_ha),2), chip=fn)
        np.savez_compressed(f'{OUT}/imagery/{name}/{name}_{d}.npz', ndvi=ndvi, ndbi=ndbi, vis=vis, valid=valid, br=(b/(r+1e-6)).astype('float32'))
        rows.append(row); print(name,d,row['builtbare_ha'],row['brightroof_ha'])
    mf=f'{OUT}/imagery/{name}/metrics.json'
    if os.environ.get('S2_START') and os.path.exists(mf):  # merge with earlier pass
        old=[r for r in json.load(open(mf))['rows'] if r.get('date')]
        seen={r['date'] for r in rows if r.get('date')}
        rows=sorted([r for r in old if r['date'] not in seen]+[r for r in rows if r.get('date')], key=lambda r:r['date'])
    json.dump(dict(site=name, bbox=bb, **s, rows=rows), open(mf,'w'), indent=1)

if __name__=='__main__':
    for n in (sys.argv[1:] or SITES): process(n)
