"""GCP pricing pages embed price tables in iframes (<page>_<hash>.frame). Collect frame URLs from cached page snapshots
and fetch each frame from Wayback at the parent snapshot's timestamp (Wayback serves nearest capture).
Usage: python gcp_frames.py   -> appends to wayback_raw/fetchlog_gcpframe.jsonl"""
import glob, re, os, sys, json, gzip
sys.path.insert(0, os.path.dirname(__file__))
from fetch import fetch, RAW
from cdx import S
log = open(f'{RAW}/fetchlog_gcpframe.jsonl', 'a')
done = {json.loads(l).get('frame') for l in open(f'{RAW}/fetchlog_gcpframe.jsonl')} if os.path.getsize(f'{RAW}/fetchlog_gcpframe.jsonl') else set()
for page in ['gpus_pricing', 'vm_instance_pricing']:
    for f in sorted(glob.glob(f'{RAW}/https_cloud_google_com_compute_{page}/*.html')):
        ts = os.path.basename(f)[:14]
        b = open(f, 'rb').read()
        try: b = gzip.decompress(b)
        except Exception: pass
        s = b.decode('utf8', 'ignore')
        frames = sorted(set(re.findall(r'src="((?:https://cloud\.google\.com)?/(?:frame/)?compute/[a-z\-]+pricing[^"]*?\.frame)"', s)))
        for fr in frames:
            u = fr if fr.startswith('http') else 'https://cloud.google.com' + fr
            if u in done: continue
            fn, st = fetch(ts, u)
            log.write(json.dumps({'url': u, 'orig': u, 'ts': ts, 'parent': page, 'frame': u, 'file': fn, 'status': st}) + '\n'); log.flush()
            done.add(u); print(ts, page, st, u[-60:]); sys.stdout.flush()
