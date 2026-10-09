#!/usr/bin/env python3
"""Tight-crop metrics + deck time-strips from s2_timeline.py outputs.
For each site, CROPS gives a lon/lat box around the actual campus (located visually on chips /
from address). Within the crop: bright-roof ha (NDVI<0.15 & mean visible refl>0.20), built/bare ha
(NDVI<0.2 & NDBI>-0.05), and impervious-or-bare share. Outputs crop_metrics.csv per site and a
strip PNG (strips/<site>_strip.png) using the clearest scene closest to selected target months.
NOTE: 10 m pixels; roofs < ~0.1 ha and dark (grey/black) roofs are under-counted; bare soil after
harvest inflates built/bare in farmland AOIs, so compare same-season scenes."""
import json, os, sys, csv, math, datetime as dt
import numpy as np
from PIL import Image, ImageDraw, ImageFont
OUT='/workspace/nebius/capacity/altdata'
CROPS={  # lon_min, lat_min, lon_max, lat_max
 'vineland_nj':(-75.0200,39.4215,-75.0075,39.4295),
 'birmingham_al':(-86.8885,33.4255,-86.8810,33.4315),
 'independence_mo':None, 'butler_pa':None, 'mantsala_fi':None, 'bethune_fr':None,
 'kansascity_mo':(-94.5815,39.0925,-94.5775,39.0955),
 'keflavik_is':None,'lappeenranta_fi':None,'longcross_uk':None,
}
def crop_idx(meta, box, shape):
    w,s,e,n=meta['bbox']; H,W=shape
    if box is None: return 0,H,0,W
    x0=int((box[0]-w)/(e-w)*W); x1=int(math.ceil((box[2]-w)/(e-w)*W))
    y0=int((n-box[3])/(n-s)*H); y1=int(math.ceil((n-box[1])/(n-s)*H))
    return max(0,y0),min(H,y1),max(0,x0),min(W,x1)
def run(site, targets=None):
    meta=json.load(open(f'{OUT}/imagery/{site}/metrics.json')); box=CROPS.get(site)
    rows=[]
    for r in meta['rows']:
        if not r.get('date'): continue
        z=np.load(r['chip'].replace('.png','.npz')); ndvi,ndbi,vis,valid=z['ndvi'],z['ndbi'],z['vis'],z['valid']; br=z['br'] if 'br' in z else None
        y0,y1,x0,x1=crop_idx(meta,box,ndvi.shape)
        nd,nb,vi,va=[a[y0:y1,x0:x1] for a in (ndvi,ndbi,vis,valid)]
        brc=br[y0:y1,x0:x1] if br is not None else np.ones_like(nd)
        if va.mean()<0.9 or not np.isfinite(nd[va]).all() or abs(nd[va].mean())>1: continue
        roof=va&(nd<0.15)&(vi>0.20)&(brc>0.85); bb=va&(nd<0.2)&(nb>-0.05)  # roof: neutral/white (blue/red>0.85) excludes yellow sand
        rows.append(dict(site=site,date=r['date'],crop_ha=round(va.size*0.01,1),brightroof_ha=round(roof.sum()*0.01,2),
                         builtbare_ha=round(bb.sum()*0.01,2),ndvi_crop=round(float(nd[va].mean()),3),chip=r['chip'],crop=(y0,y1,x0,x1)))
    with open(f'{OUT}/imagery/{site}/crop_metrics.csv','w',newline='') as f:
        wr=csv.DictWriter(f,fieldnames=[k for k in rows[0] if k not in('chip','crop')],extrasaction='ignore'); wr.writeheader(); wr.writerows(rows)
    # strip
    if targets is None: targets=['2023-07','2024-07','2025-03','2025-07','2025-10','2026-03','2026-06','2026-09']
    pick=[]
    for t in targets:
        td=dt.date.fromisoformat(t+'-15')
        c=min(rows,key=lambda r:abs((dt.date.fromisoformat(r['date'])-td).days))
        if abs((dt.date.fromisoformat(c['date'])-td).days)<=50 and c not in pick: pick.append(c)
    tiles=[]
    for r in pick:
        im=Image.open(r['chip']); y0,y1,x0,x1=[v*4 for v in r['crop']]
        pad=40; im=im.crop((max(0,x0-pad),max(0,y0-pad),min(im.width,x1+pad),min(im.height,y1+pad)))
        sc=320/im.width; im=im.resize((320,int(im.height*sc)),Image.LANCZOS)
        d=ImageDraw.Draw(im); d.rectangle([0,0,319,22],fill=(0,0,0)); d.text((5,4),f"{r['date']}  roof~{r['brightroof_ha']}ha",fill=(255,255,255))
        tiles.append(im)
    H=max(t.height for t in tiles); nc=4; nr=math.ceil(len(tiles)/nc)
    strip=Image.new('RGB',(nc*330,nr*(H+10)+40),(255,255,255))
    for i,t in enumerate(tiles): strip.paste(t,((i%nc)*330,(i//nc)*(H+10)))
    d=ImageDraw.Draw(strip); d.text((5,nr*(H+10)+8),f"{site}: Sentinel-2 L2A true colour (10 m), Element84 Earth Search. 'roof' = bright low-NDVI pixels in crop (est.). Source: Copernicus/ESA via AWS Open Data.",fill=(0,0,0))
    os.makedirs(f'{OUT}/strips',exist_ok=True); strip.save(f'{OUT}/strips/{site}_strip.png')
    for r in rows: print(site,r['date'],r['brightroof_ha'],r['builtbare_ha'],r['ndvi_crop'])
if __name__=='__main__':
    for s in sys.argv[1:]: run(s)
