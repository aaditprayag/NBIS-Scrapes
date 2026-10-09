import re,json
F={'Jun-2029':'tm2516825d1_ex99-4','Jun-2031':'tm2516825d1_ex99-5','Sep-2030':'tm2525580d4_ex4-1','Sep-2032':'tm2525580d4_ex4-3','Mar-2031':'tm268409d4_ex4-1','Mar-2033':'tm268409d4_ex4-3','Feb-2030':'tm2623863d1_ex4-1','Feb-2034':'tm2623863d1_ex4-3'}
out={}
for k,f in F.items():
    L=open(f+'.htm.txt').read().split('\n')
    cand=[n for n,l in enumerate(L) if re.match(r'^\s*(Make-Whole Fundamental Change Effective )?Date\s*\|.*\$',l) or re.match(r'^\s*Make-Whole Fundamental Change Effective Date\s*\|.*\$',l)]
    i=cand[-1]
    hdr=[x for x in (c.strip() for c in L[i].split('|')[1:]) if x and x!='$']
    rows=[]
    for l in L[i+1:i+15]:
        p=[c.strip() for c in l.split('|')]
        if re.search(r'^[A-Z][a-z]+ \d+, 20\d\d$',p[0]): rows.append([p[0]]+[x for x in p[1:] if x and x!='$'])
        elif rows: break
    t=re.sub(r'\s+',' ',' '.join(L))
    m=re.search(r'(?:exceed|exceeds) ([\d.]+) (?:shares|Ordinary Shares)',t)
    out[k]={'file':f,'hdr':hdr,'rows':rows,'max':m.group(1) if m else None}
    print(k,out[k]['max'],len(rows),hdr[0],hdr[-1],rows[0][:2])
json.dump(out,open('makewhole.json','w'),indent=1)
