import re,json
out={}
files={'k_fy24__orcl-20240531.htm.txt':('May-24','May-23'),'k_fy25__orcl-20250531.htm.txt':('May-25','May-24'),'k_fy26__orcl-20260531.htm.txt':('May-26','May-25')}
num=lambda s: None if s.strip() in ('—','-','') else float(s.replace(',','').replace('$','').strip())
for f,(cur,pri) in files.items():
    t=open('/workspace/nebius/peers/oracle/raw/'+f).read()
    i=t.find('Notes payable and other borrowings consisted of the following')
    j=t.find('Total senior notes and other borrowings',i)
    seg=t[i:j]
    seg=re.sub(r'\d+ \s*Table of Contents( Index to Financial Statements)?','',seg)
    seg=re.sub(r'\| May 31, \| \d{4} \| \d{4} \| \(Amounts in millions\) \| Date of\s*Issuance \| Amount \| Effective Interest\s+Rate \| Amount \| Effective Interest\s+Rate \|','|',seg)
    cat='Senior notes (fixed)'
    rows=[]
    # tokens split by |
    toks=[x.strip() for x in seg.split('|')]
    k=0
    while k<len(toks):
        x=toks[k]
        if re.match(r'^(Fixed-rate senior notes|Floating-rate senior notes|Term loan credit agreement|Commercial paper|Other borrowings)',x):
            if x.startswith('Fixed'): cat='Senior notes (fixed)'
            elif x.startswith('Floating'): cat='Floating-rate notes'
            elif x.startswith('Term'): cat='Term loans'
        if re.match(r'^(\$|€|Commercial paper notes|Other borrowings due)',x) and k+4<len(toks):
            name=x
            c = 'Commercial paper' if x.startswith('Commercial') else ('Other' if x.startswith('Other borrowings') else cat)
            # find the next 4 numeric-ish tokens (skip issuance date token)
            vals=[]; m=k+1
            while len(vals)<4 and m<len(toks):
                y=toks[m]
                if re.match(r'^[\d,\.]+%?$|^—$|^N\.A\.?$',y): vals.append(y)
                elif re.match(r'^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$',y): pass
                elif y=='' or y=='$': pass
                else: break
                m+=1
            if len(vals)==4:
                a1,r1,a0,r0=vals
                rf=lambda r: None if r.startswith('N.A') or r=='—' else float(r.rstrip('%'))
                rows.append(dict(name=name,cat=c,**{cur:(num(a1),rf(r1)),pri:(num(a0),rf(r0))}))
                k=m; continue
        k+=1
    for snap in (cur,pri):
        p=sum((r[snap][0] or 0) for r in rows)
        w=sum((r[snap][0] or 0)*(r[snap][1] or 0) for r in rows if r[snap][1] is not None)
        print(f,snap,'rows',len(rows),'principal',round(p,1),'wavg',round(w/p,3))
    out[f]=dict(cur=cur,pri=pri,rows=rows)
json.dump(out,open('/workspace/nebius/peers/oracle/work/orcl_debt_tables.json','w'),indent=1)
