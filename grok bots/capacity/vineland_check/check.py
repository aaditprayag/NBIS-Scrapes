import sys,warnings
warnings.filterwarnings('ignore')
sys.path.insert(0,'/workspace/nebius/capacity/altdata/scripts')
from detail import CROPS
from pystac_client import Client
c=CROPS['vineland_nj']
cl=Client.open('https://earth-search.aws.element84.com/v1')
for it in sorted(cl.search(collections=['sentinel-2-l2a'],bbox=list(c),datetime='2026-09-01/2026-10-04',max_items=50).items(),key=lambda i:i.datetime):
    print('S2',it.datetime.isoformat()[:16],it.id,'cloud',it.properties.get('eo:cloud_cover'))
pc=Client.open('https://planetarycomputer.microsoft.com/api/stac/v1')
for it in sorted(pc.search(collections=['landsat-c2-l2'],bbox=list(c),datetime='2026-08-20/2026-10-04',max_items=50).items(),key=lambda i:i.datetime):
    print('LS',it.datetime.isoformat()[:16],it.id,'cloud',it.properties.get('eo:cloud_cover'))
