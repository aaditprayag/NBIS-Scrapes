import re,sys,os
sys.path.insert(0,os.path.dirname(__file__))
from guidance_data import G,W
from sources import S
E='/workspace/nebius/filings/edgar/'; A='/workspace/nebius/filings/altdata/'
LOCAL={'Q2_24_PR':E+'tmb-20240630xex99d1.htm.txt','OCT24_PRES':E+'tm2426436d1_ex99-1.htm.txt','Q3_24_PR':E+'tm2427209d1_ex99-1.htm.txt',
'DEC24_PIPE':E+'tm2429858d1_ex99-1.htm.txt','Q4_24_PR':E+'tm257185d1_ex99-1.htm.txt','FY24_20F':A+'raw/nbis-20241231x20f.txt',
'Q1_25_PR':E+'tm2515580d1_ex99-1.htm.txt','Q1_25_SHL':E+'tm2515580d1_ex99-2.htm.txt','Q2_25_PR':E+'tm2522866d1_ex99-1.htm.txt',
'Q2_25_SHL':E+'tm2522866d1_ex99-2.htm.txt','Q3_25_PR':E+'tm2530882d1_ex99-1.htm.txt','Q3_25_SHL':E+'tm2530882d1_ex99-2.htm.txt',
'Q4_25_PR':E+'tm266173d1_ex99-1.htm.txt','Q4_25_SHL':E+'tm266173d1_ex99-2.htm.txt','FY25_20F':E+'nbis-20251231x20f.htm.txt',
'Q1_26_PR':E+'tm2614392d1_ex99-1.htm.txt','Q1_26_SHL':E+'tm2614392d1_ex99-2.htm.txt','Q1_26_OFR':E+'nbis-20260331xex99d1.htm.txt','Q1_26_FS':E+'nbis-20260331xex99d2.htm.txt',
'JUL26_SECURED':E+'tm2620683d1_ex99-1.htm.txt','Q2_26_PR':E+'tm2622968d1_ex99-1.htm.txt','Q2_26_SHL':E+'tm2622968d1_ex99-2.htm.txt','Q2_26_OFR':E+'nbis-20260812xex99d1.htm.txt','Q2_26_FS':E+'nbis-20260812xex99d2.htm.txt'}
for q in ['Q3_24','Q4_24','Q1_25','Q2_25','Q3_25','Q4_25','Q1_26','Q2_26']:
    LOCAL['CALL_'+q]=A+'transcripts/transcript_'+q.lower().replace('_','-').replace('q','q',1)[:2]+'-20'+q[-2:]+'.txt'
LOCAL['FOOL_Q2_26']=A+'transcripts/transcript_q2-2026_fool.txt'
def norm(s): return re.sub(r'\s+',' ',s.replace('\u200b',' ').replace('\xa0',' ')).strip().lower()
cache={}
def found(key,quote):
    if key not in LOCAL or not quote: return None
    if key not in cache: cache[key]=norm(open(LOCAL[key],errors='ignore').read())
    parts=[p for p in re.split(r'\.\.\.|\[.*?\]',quote) if p.strip()]
    return all(norm(p) in cache[key] for p in parts)
bad=0
for g in G:
    r=found(g['src'],g['quote']); r2=found(g['asrc'],g['aquote'])
    if r is False or r2 is False: bad+=1; print('FAIL',g['id'],r,r2)
for d,k,_,q in W:
    if found(k,q) is False: bad+=1; print('FAIL W',d,k)
print('checked',len(G),'+',len(W),'bad',bad)
