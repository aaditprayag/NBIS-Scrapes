import sys,time
from cdx import cdx
frm=sys.argv[1]
for u in sys.argv[2:]:
    pre=u.endswith('*'); 
    rows=cdx(u.rstrip('*'),frm=frm,match='prefix' if pre else None,collapse='timestamp:6' if not pre else 'urlkey')
    if rows is None: print(u,'FAIL'); continue
    if pre: print(u,len(rows),sorted(set(r[2] for r in rows))[:40])
    else: print(u,len(rows),' '.join(r[1][:6] for r in rows))
    sys.stdout.flush(); time.sleep(1)
