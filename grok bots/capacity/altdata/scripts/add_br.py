# backfill blue/red ratio into existing npz files from saved PNG chips is lossy; instead recompute from PNG (gamma-stretched) ratio as proxy
import glob, numpy as np
from PIL import Image
for f in glob.glob('/workspace/nebius/capacity/altdata/imagery/*/*.npz'):
    z=dict(np.load(f))
    if 'br' in z: continue
    im=np.asarray(Image.open(f.replace('.npz','.png'))).astype('float32')[::4,::4]
    r=(im[...,0]/255)**1.6; b=(im[...,2]/255)**1.6
    z['br']=(b/(r+1e-6)).astype('float32')[:z['ndvi'].shape[0],:z['ndvi'].shape[1]]
    np.savez_compressed(f,**z)
