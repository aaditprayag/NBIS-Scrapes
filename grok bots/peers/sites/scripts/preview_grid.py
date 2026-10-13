#!/usr/bin/env python3
"""Locator helper: downsized AOI chip with lat/lon grid (used to pick tight campus crops). Usage: preview_grid.py site date [out]"""
import sys, json
from PIL import Image, ImageDraw
OUT='/workspace/nebius/peers/sites'
site,date=sys.argv[1],sys.argv[2]; out=sys.argv[3] if len(sys.argv)>3 else f'/tmp/prev/{site}_{date}.png'
m=json.load(open(f'{OUT}/imagery/{site}/metrics.json')); x0,y0,x1,y1=m['bbox']
im=Image.open(f'{OUT}/imagery/{site}/{site}_{date}.png'); S=1000; im=im.resize((S,int(im.height*S/im.width))); H=im.height
d=ImageDraw.Draw(im); n=10
for i in range(n+1):
    x=i*S/n; lon=x0+(x1-x0)*i/n; d.line([(x,0),(x,H)],fill=(255,255,0),width=1); d.text((x+2,2),f'{lon:.4f}',fill=(255,255,0))
    y=i*H/n; lat=y1-(y1-y0)*i/n; d.line([(0,y),(S,y)],fill=(0,255,255),width=1); d.text((2,y+2),f'{lat:.4f}',fill=(0,255,255))
im.save(out); print(out)
