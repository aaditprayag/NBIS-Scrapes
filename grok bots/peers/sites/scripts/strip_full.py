#!/usr/bin/env python3
"""Deck strip from full-AOI chips (imagery/<site>/*.png). Usage: strip_full.py site [YYYY-MM ...]"""
import sys, json, math, datetime as dt
from PIL import Image, ImageDraw
OUT='/workspace/nebius/peers/sites'
site=sys.argv[1]; targets=sys.argv[2:] or ['2023-07','2024-07','2025-04','2025-09','2026-01','2026-04','2026-07','2026-09']
m=json.load(open(f'{OUT}/imagery/{site}/metrics.json')); rows=[r for r in m['rows'] if r.get('date') and abs(r.get('ndvi_mean',0))<=1]
pick=[]
for t in targets:
    td=dt.date.fromisoformat(t+'-15'); c=min(rows,key=lambda q:abs((dt.date.fromisoformat(q['date'])-td).days))
    if abs((dt.date.fromisoformat(c['date'])-td).days)<=60 and c not in pick: pick.append(c)
tiles=[]
for q in pick:
    im=Image.open(q['chip']); im=im.resize((360,int(im.height*360/im.width)),Image.LANCZOS)
    d=ImageDraw.Draw(im); d.rectangle([0,0,359,20],fill=(0,0,0)); d.text((5,4),q['date'],fill=(255,255,255)); tiles.append(im)
H=max(t.height for t in tiles); nc=4; nr=math.ceil(len(tiles)/nc)
st=Image.new('RGB',(nc*370,nr*(H+10)+30),(255,255,255))
for i,t in enumerate(tiles): st.paste(t,((i%nc)*370,(i//nc)*(H+10)))
ImageDraw.Draw(st).text((5,nr*(H+10)+8),f"{site} | Sentinel-2 L2A true colour 10 m (Copernicus/ESA via Element84 Earth Search) | AOI {[round(v,4) for v in m['bbox']]}",fill=(0,0,0))
st.save(f'{OUT}/strips/{site}_full_strip.png'); print(site,[q['date'] for q in pick])
