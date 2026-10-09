import sys,math,warnings,requests,numpy as np,rasterio,time
warnings.filterwarnings('ignore')
sys.path.insert(0,'/workspace/nebius/capacity/altdata/scripts')
from detail import CROPS
from pystac_client import Client
from rasterio.warp import transform_bounds
from rasterio.windows import from_bounds
from PIL import Image
c=CROPS['vineland_nj']
def rd(href,bb,tok=''):
    with rasterio.open(href+(('?'+tok) if tok else '')) as ds:
        w=from_bounds(*transform_bounds('EPSG:4326',ds.crs,*bb),ds.transform); return ds.read(1,window=w,boundless=True,fill_value=0)
cl=Client.open('https://earth-search.aws.element84.com/v1')
for d in ['2026-09-08','2026-09-16','2026-10-01']:
    its=list(cl.search(collections=['sentinel-2-l2a'],bbox=list(c),datetime=d,max_items=5).items())
    for it in its:
        scl=rd(it.assets['scl'].href,c)
        cloud=np.isin(scl,[3,8,9,10]).mean(); nod=(scl==0).mean()
        print(d,it.id,'crop cloud/shadow frac %.2f nodata %.2f'%(cloud,nod))
        if cloud<0.2 and nod<0.05:
            rgb=np.dstack([rd(it.assets[b].href,c) for b in ['red','green','blue']]).astype('float32')/10000
            img=np.clip(rgb/0.3,0,1)**(1/1.6)
            im=Image.fromarray((img*255).astype('uint8')); im=im.resize((im.width*8,im.height*8),Image.NEAREST); im.save(f'vineland_{d}.png')
            nir=rd(it.assets['nir'].href,c).astype('float32')/10000
            ndvi=(nir-rgb[...,0])/(nir+rgb[...,0]+1e-6)
            white=((ndvi<0.15)&(rgb.mean(2)>0.18)&(rgb[...,2]/(rgb[...,0]+1e-6)>0.8)).sum()*0.01
            dark=((ndvi<0.15)&(rgb.mean(2)<0.06)).sum()*0.01
            print('   roof_white_ha %.2f dark_lowNDVI_ha %.2f'%(white,dark))
        break
tok=requests.get('https://planetarycomputer.microsoft.com/api/sas/v1/token/landsat-c2-l2',timeout=60).json()['token']
pc=Client.open('https://planetarycomputer.microsoft.com/api/stac/v1')
cx=(c[0]+c[2])/2; cy=(c[1]+c[3])/2; dd=1500/111320; bg=[cx-dd/math.cos(math.radians(cy)),cy-dd,cx+dd/math.cos(math.radians(cy)),cy+dd]
for it in sorted(pc.search(collections=['landsat-c2-l2'],bbox=list(c),datetime='2026-08-20/2026-10-04',max_items=50).items(),key=lambda i:i.datetime):
    st=rd(it.assets['lwir11'].href,bg,tok).astype('float32'); qa=rd(it.assets['qa_pixel'].href,bg,tok)
    ok=(st>0)&((qa&0b11111)==0)
    T=st*0.00341802+149.0-273.15; H,W=T.shape
    x0=int((c[0]-bg[0])/(bg[2]-bg[0])*W); x1=int((c[2]-bg[0])/(bg[2]-bg[0])*W); y0=int((bg[3]-c[3])/(bg[3]-bg[1])*H); y1=int((bg[3]-c[1])/(bg[3]-bg[1])*H)
    m=np.zeros_like(ok); m[y0:y1,x0:x1]=True
    if ok.mean()<0.9 or (ok&m).sum()<20: print('LS',it.id,'clear frac %.2f -> skipped'%ok.mean()); continue
    print('LS',it.datetime.date(),it.id,'campus %.2f bg %.2f anomaly %.2f C'%(T[ok&m].mean(),T[ok&~m].mean(),T[ok&m].mean()-T[ok&~m].mean()))
