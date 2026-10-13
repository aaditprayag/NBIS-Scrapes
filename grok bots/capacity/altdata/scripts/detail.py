#!/usr/bin/env python3
"""Re-read true reflectance for a tight campus crop for every scene chosen by s2_timeline.py, compute
 roof_white_ha: NDVI<0.15, mean(RGB)>0.18, blue/red>0.80 (neutral/white roofs; excludes yellow sand),
 roof_grey_ha:  NDVI<0.15, 0.07<mean(RGB)<=0.18, blue/red>0.80, NDBI<0.15 (grey roofs/equipment/asphalt/concrete),
 bare_ha:       NDVI<0.2, blue/red<=0.80, mean>0.12 (exposed soil/sand = earthworks),
and write detail_metrics.csv + deck strip strips/<site>_strip.png (stretch 0.35, 2 rows x 4).
Usage: detail.py site [YYYY-MM ...target months for strip]"""
import sys, os, json, csv, math, datetime as dt, warnings
warnings.filterwarnings('ignore')
import numpy as np, rasterio
from rasterio.warp import transform_bounds
from rasterio.windows import from_bounds
from PIL import Image, ImageDraw
from pystac_client import Client
os.environ.update(GDAL_DISABLE_READDIR_ON_OPEN='EMPTY_DIR', AWS_NO_SIGN_REQUEST='YES')
OUT='/workspace/nebius/capacity/altdata'
CROPS={'vineland_nj':(-75.0200,39.4215,-75.0075,39.4295),'birmingham_al':(-86.8885,33.4250,-86.8805,33.4320),
       'kansascity_mo':(-94.5815,39.0925,-94.5775,39.0955)}
CROPS.update(json.load(open(f'{OUT}/scripts/crops_extra.json')) if os.path.exists(f'{OUT}/scripts/crops_extra.json') else {})
def rd(href,bb,shape=None):
    with rasterio.open(href) as ds:
        w=from_bounds(*transform_bounds('EPSG:4326',ds.crs,*bb),ds.transform)
        kw=dict(window=w,boundless=True,fill_value=0)
        if shape: kw.update(out_shape=shape)
        return ds.read(1,**kw).astype('float32')/10000
def main(site,targets):
    meta=json.load(open(f'{OUT}/imagery/{site}/metrics.json')); bb=CROPS[site]
    cl=Client.open('https://earth-search.aws.element84.com/v1'); rows=[]; os.makedirs(f'{OUT}/imagery/{site}/detail',exist_ok=True)
    for r in meta['rows']:
        if not r.get('scene'): continue
        it=next(cl.search(collections=['sentinel-2-l2a'],ids=[r['scene']]).items())
        R=rd(it.assets['red'].href,bb); G=rd(it.assets['green'].href,bb); B=rd(it.assets['blue'].href,bb); N=rd(it.assets['nir'].href,bb)
        S=rd(it.assets['swir16'].href,bb,R.shape)
        if (R==0).mean()>0.05: continue
        ndvi=(N-R)/(N+R+1e-6); ndbi=(S-N)/(S+N+1e-6); vis=(R+G+B)/3; br=B/(R+1e-6)
        white=(ndvi<0.15)&(vis>0.18)&(br>0.80); grey=(ndvi<0.15)&(vis>0.07)&(vis<=0.18)&(br>0.80)&(ndbi<0.15); bare=(ndvi<0.2)&(br<=0.80)&(vis>0.12)
        img=np.clip(np.dstack([R,G,B])/0.35,0,1)**(1/1.4); fn=f"{OUT}/imagery/{site}/detail/{site}_{r['date']}_crop.png"
        Image.fromarray((img*255).astype('uint8')).resize((R.shape[1]*8,R.shape[0]*8),Image.LANCZOS).save(fn)
        rows.append(dict(site=site,date=r['date'],scene=r['scene'],crop_ha=round(R.size*0.01,1),roof_white_ha=round(white.sum()*0.01,2),
                         roof_grey_ha=round(grey.sum()*0.01,2),bare_ha=round(bare.sum()*0.01,2),ndvi_mean=round(float(ndvi.mean()),3),png=fn))
        print(site,r['date'],rows[-1]['roof_white_ha'],rows[-1]['roof_grey_ha'],rows[-1]['bare_ha'])
    with open(f'{OUT}/imagery/{site}/detail_metrics.csv','w',newline='') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
    pick=[]
    for t in targets:
        td=dt.date.fromisoformat(t+'-15'); c=min(rows,key=lambda q:abs((dt.date.fromisoformat(q['date'])-td).days))
        if abs((dt.date.fromisoformat(c['date'])-td).days)<=50 and c not in pick: pick.append(c)
    tiles=[]
    for q in pick:
        im=Image.open(q['png']); sc=360/im.width; im=im.resize((360,int(im.height*sc)),Image.LANCZOS)
        d=ImageDraw.Draw(im); d.rectangle([0,0,359,20],fill=(0,0,0)); d.text((5,4),f"{q['date']}  white roof ~{q['roof_white_ha']} ha",fill=(255,255,255)); tiles.append(im)
    H=max(t.height for t in tiles); nc=4; nr=math.ceil(len(tiles)/nc)
    st=Image.new('RGB',(nc*370,nr*(H+10)+30),(255,255,255))
    for i,t in enumerate(tiles): st.paste(t,((i%nc)*370,(i//nc)*(H+10)))
    ImageDraw.Draw(st).text((5,nr*(H+10)+8),f"{site} | Sentinel-2 L2A true colour, 10 m (Copernicus/ESA via Element84 Earth Search). Crop {bb}. Roof area = est. from bright neutral low-NDVI pixels.",fill=(0,0,0))
    os.makedirs(f'{OUT}/strips',exist_ok=True); st.save(f'{OUT}/strips/{site}_strip.png')
if __name__=='__main__':
    main(sys.argv[1], sys.argv[2:] or ['2023-07','2024-07','2025-04','2025-07','2025-10','2026-03','2026-06','2026-09'])
