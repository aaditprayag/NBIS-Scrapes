import sys,re,gzip
from bs4 import BeautifulSoup
def raw(fn):
    d=open(fn,'rb').read()
    if d[:2]==b'\x1f\x8b':
        try: d=gzip.decompress(d)
        except Exception: pass
    elif not d.lstrip()[:1] in (b'<',b'{',b'['):
        try:
            import brotli; d=brotli.decompress(d)
        except Exception: pass
    return d.decode('utf-8','ignore')
def text(fn):
    s=BeautifulSoup(raw(fn),'lxml')
    for t in s(['script','style','noscript']): t.decompose()
    return re.sub(r'\s+',' ',s.get_text(' ',strip=True))
if __name__=='__main__':
    t=text(sys.argv[1]); k=sys.argv[2] if len(sys.argv)>2 else 'H100'
    n=int(sys.argv[3]) if len(sys.argv)>3 else 2500
    i=t.find(k); print(len(t), t[max(0,i-300):i+n])
