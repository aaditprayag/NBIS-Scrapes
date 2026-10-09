import sys,re
f=sys.argv[1]; L=open(f).read().split('\n')
def page(i):
    for j in range(i,min(i+400,len(L))):
        s=L[j].strip()
        if re.fullmatch(r'(F-)?\d{1,3}',s) and j+1<len(L) and 'Table of Contents' in L[j+1]: return s
        if re.fullmatch(r'(F-)?\d{1,3}',s): return s
    return '?'
for pat in sys.argv[2:]:
    for i,l in enumerate(L):
        if re.search(pat,l,re.I): print(f'[p.{page(i)} L{i+1}]',l[:1200]); 
