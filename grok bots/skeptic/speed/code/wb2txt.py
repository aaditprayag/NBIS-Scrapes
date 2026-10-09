# Convert Wayback raw HTML snapshots (downloaded by GPU Market teammate to /workspace/nebius/wayback_raw) to pipe-delimited text
import re,html,os,glob
SRC='/workspace/nebius/wayback_raw'; OUT='/workspace/nebius/skeptic/speed/txt'
for d in glob.glob(SRC+'/*/'):
    for f in glob.glob(d+'*.html'):
        s=open(f,errors='ignore').read()
        t=re.sub(r'<script.*?</script>|<style.*?</style>|<svg.*?</svg>','',s,flags=re.S)
        t=re.sub(r'<[^>]+>','|',t); t=html.unescape(t).replace('\xa0',' '); t=re.sub(r'\s*\|[\s|]*','|',t)
        o=f'{OUT}/{os.path.basename(d[:-1])}__{os.path.basename(f)[:-5]}.txt'
        open(o,'w').write(t)
