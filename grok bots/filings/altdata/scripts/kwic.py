import sys,re
f=sys.argv[1]; pat=re.compile(sys.argv[2],re.I); w=int(sys.argv[3]) if len(sys.argv)>3 else 300
t=re.sub(r'\s+',' ',open(f,errors='ignore').read())
last=-10**9
for m in pat.finditer(t):
    if m.start()-last<w: continue
    last=m.start()
    print('…'+t[max(0,m.start()-w):m.end()+w]+'…\n')
