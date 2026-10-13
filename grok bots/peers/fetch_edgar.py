# usage: python3 fetch_edgar.py <cik> <accession> <tag> <regex> <outdir>
import sys,json,subprocess,time,os,re
UA="Aadit Prayag research aadit@example.com"
cik,acc,tag,pat,outd=sys.argv[1:6]
a=acc.replace('-','')
base=f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{a}"
idx=json.loads(subprocess.check_output(['curl','-s','-A',UA,base+'/index.json']))
for it in idx['directory']['item']:
    n=it['name']
    if n.endswith('.htm') and 'index' not in n and re.search(pat,n,re.I):
        out=os.path.join(outd,f"{tag}__{n}")
        subprocess.run(['curl','-s','-A',UA,base+'/'+n,'-o',out]); time.sleep(0.25)
        subprocess.run(['python3','/workspace/nebius/peers/coreweave/raw/h2t.py',out])
        print(out, os.path.getsize(out+'.txt'), base+'/'+n)
