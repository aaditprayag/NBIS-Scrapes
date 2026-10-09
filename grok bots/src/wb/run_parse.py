import json,glob,os,sys,re
sys.path.insert(0,os.path.dirname(__file__))
from txt import text
from parse import PARSERS, money, oracle_api, vantage
from txt import raw
RAW='/workspace/nebius/wayback_raw'
PROV={'oracle_api':'Oracle OCI','vantage_aws':'AWS','vantage_az':'Azure','awsp5':'AWS','gcpvm':'Google Cloud','azure':'Azure','runpod':'RunPod','vast':'Vast.ai','nebius':'Nebius','coreweave':'CoreWeave','coreweave2024':'CoreWeave','cwcks':'CoreWeave','lambda':'Lambda','together':'Together AI','crusoe':'Crusoe','aws':'AWS','oracle':'Oracle OCI','gcp':'Google Cloud','nscale':'Nscale'}
def load_logs(only=None):
    rows=[]
    for f in glob.glob(f'{RAW}/fetchlog_*.jsonl'):
        lab=f.split('fetchlog_')[1][:-6]
        if only and lab not in only: continue
        for l in open(f):
            d=json.loads(l); d['label']=lab; d['provider']=PROV.get(lab,lab); rows.append(d)
    return rows
def parse_all(only=None):
    recs=[]; cov=[]
    for d in load_logs(only):
        if not d.get('ts'):
            cov.append(dict(provider=d['provider'],url=d['url'],month=None,snapshot_timestamp=None,status=d['status'],n_prices=0,note='CDX returned no 200 snapshots' if d['status']=='no_snapshots' else d['status'])); continue
        p=PARSERS.get(d['provider'])
        if d['label']=='oracle_api': p=lambda t,f=d.get('file'): oracle_api(raw(f)) if f else []
        elif d['label']=='vantage_aws': p=lambda t: vantage(t,'p5.48xlarge')
        elif d['label']=='vantage_az': p=lambda t: vantage(t,'nd96isr-h100-v5')
        elif d['label'] in ('awsp5','gcpvm','gcp','azure','oracle','vast'): p=lambda t: []
        st=d['status']; n=0; note=''
        if d.get('file') and p:
            t=text(d['file'])
            rs=p(t)
            for r in rs:
                if r['pricing_type']=='PARSE_WARNING': note+=f"warn:{r['term']}; "; continue
                q=money(r['quoted_price'])
                usd=None
                if q is not None:
                    unit=r['quoted_unit']
                    if 'minute' in unit: usd=round(q*60,4)
                    elif r['gpus_per_unit'] and r['gpus_per_unit']>1 and 'instance' in unit or 'VM-hour' in unit: usd=round(q/r['gpus_per_unit'],4)
                    else: usd=q
                r.update(provider=d['provider'],usd_per_gpu_hr=usd,snapshot_timestamp=d['ts'],
                         snapshot_url=f"https://web.archive.org/web/{d['ts']}/{d['orig']}",original_url=d['orig'])
                recs.append(r); n+=1
            if not rs: note+='page fetched but no GPU prices in static HTML (JS-rendered or prices not listed); '
        elif not p: note+='no parser'
        cov.append(dict(provider=d['provider'],url=d['orig'],month=d['ts'][:4]+'-'+d['ts'][4:6],snapshot_timestamp=d['ts'],status=st,n_prices=n,note=note.strip()))
    return recs,cov
if __name__=='__main__':
    only=sys.argv[1].split(',') if len(sys.argv)>1 else None
    recs,cov=parse_all(only)
    import collections
    for r in sorted(recs,key=lambda r:(r['provider'],r['snapshot_timestamp'])):
        if '-v' in sys.argv or (r['headline'] and r['chip'] in ('H100','H200','B200','GB200','B300')):
            print(r['provider'][:8],r['snapshot_timestamp'][:8],r['chip'],r['pricing_type'],r['term'][:30],r['quoted_price'],r['usd_per_gpu_hr'],r['config'][:35],'H' if r['headline'] else '-')
    for c in cov:
        if c['n_prices']==0 or c['note']: print('COV',c)
