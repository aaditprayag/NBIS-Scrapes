"""Convert every cached Wayback HTML under wayback_raw/<page>/<ts>.html into plain text at
/workspace/nebius/market/scrapers/time_to_deploy/cache_text/<page>/<ts>.txt (idempotent)."""
import sys,glob,os
sys.path.insert(0,'/workspace/nebius/market/scrapers/wayback')
from txt import text
RAW='/workspace/nebius/wayback_raw'; OUT='/workspace/nebius/market/scrapers/time_to_deploy/cache_text'
for f in sorted(glob.glob(RAW+'/*/*.html')):
    d=os.path.basename(os.path.dirname(f)); o=f'{OUT}/{d}/{os.path.basename(f)[:-5]}.txt'
    if os.path.exists(o) and os.path.getmtime(o)>=os.path.getmtime(f): continue
    os.makedirs(os.path.dirname(o),exist_ok=True)
    try: open(o,'w').write(text(f))
    except Exception as e: print('ERR',f,e)
