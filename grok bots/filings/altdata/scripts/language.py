import re,os,sys,json
sys.path.insert(0,os.path.dirname(__file__))
from sources import S
E='/workspace/nebius/filings/edgar/'; A='/workspace/nebius/filings/altdata/'
Q=['Q1-24','Q2-24','Q3-24','Q4-24','Q1-25','Q2-25','Q3-25','Q4-25','Q1-26','Q2-26']
# Documents: (quarter, type, src_key, local file). 'letter' = shareholder letter (ex 99.2); for Q2-24..Q4-24 no letter existed -> press release used, flagged.
DOCS=[('Q2-24','release*', 'Q2_24_PR',E+'tmb-20240630xex99d1.htm.txt'),
('Q3-24','release*','Q3_24_PR',E+'tm2427209d1_ex99-1.htm.txt'),('Q4-24','release*','Q4_24_PR',E+'tm257185d1_ex99-1.htm.txt'),
('Q1-25','letter','Q1_25_SHL',E+'tm2515580d1_ex99-2.htm.txt'),('Q2-25','letter','Q2_25_SHL',E+'tm2522866d1_ex99-2.htm.txt'),
('Q3-25','letter','Q3_25_SHL',E+'tm2530882d1_ex99-2.htm.txt'),('Q4-25','letter','Q4_25_SHL',E+'tm266173d1_ex99-2.htm.txt'),
('Q1-26','letter','Q1_26_SHL',E+'tm2614392d1_ex99-2.htm.txt'),('Q2-26','letter','Q2_26_SHL',E+'tm2622968d1_ex99-2.htm.txt')]
for q in ['Q3-24','Q4-24','Q1-25','Q2-25','Q3-25','Q4-25','Q1-26','Q2-26']:
    DOCS.append((q,'call','CALL_'+q.replace('-','_'),A+'transcripts/transcript_'+q[:2].lower()+'-20'+q[-2:]+'.txt'))
# ---------------- Phrase dictionary (regex, case-insensitive) ----------------
FAM={
'supply_constrained':[
 r'sold[- ]out', r'sell(ing)? out', r'sells? out', r'capacity[- ]constrain\w*', r'supply[- ]constrain\w*', r'\bconstrain(ed|t|ts)\b',
 r'bottlenecks?', r'demand (continued to |again |significantly |continues to |remained |remains |far )?(exceed\w*|outpac\w*|outstrip\w*) (available )?(supply|capacity)',
 r'(exceed\w*|outpac\w*|outstrip\w*) (available )?(supply|capacity)', r'(supply|capacity) shortages?', r'shortages? of (gpus?|capacity|compute|power)',
 r'(peak|full|near[- ]peak) utili[sz]ation', r'fully utili[sz]ed', r'only (real )?limitation', r'limited by (the amount of )?capacity',
 r'lack of capacity', r'waitlists?|wait lists?', r'could sell (our|the) entire', r'can sell it', r'everything we build', r'supply[- ]demand (imbalance|gap)'],
'demand_pipeline':[
 r'\bdemand\b', r'pipelines?', r'customer interest', r'(strong|significant|growing|robust|tremendous) interest', r'inbound', r'backlog',
 r'bookings?', r'new logos?', r'customer wins?', r'\bdeal flow\b', r'\bRPO\b|remaining performance obligations?', r'contracted revenue', r'(signed|closed) (a |several |four |new )?(deals?|contracts?)'],
'pricing_competition':[
 r'price pressure', r'pricing pressure', r'\bpric(e|es|ing|ed)\b', r'competition|competitive|competitors?', r'discount\w*', r'commodit\w*', r'oversuppl\w*|glut',
 r'per[- ]hour|\$/hr|per gpu[- ]hour', r'\bper (mw|megawatt)\b'],
}
POS_PRICE=[r'(increas\w*|higher|rising|rise|raised?|strong|stronger) (prices|pricing)', r'pric(es|ing) (increas\w*|rose|rising|went up|higher|strengthen\w*|continued to rise)', r'higher[- ]priced', r'repric\w* higher', r'(at a|significant) premium']
NEG_PRICE=[r'price (pressure|declines?|decreases?|erosion|war)', r'pricing pressure', r'(lower|falling|declining) (prices|pricing)', r'commoditi\w*', r'oversuppl\w*|glut', r'discount\w*']
def strip_boiler(t,kind):
    # drop forward-looking statements / safe-harbor boilerplate
    i=re.search(r'Forward[- ]looking statements? This (document|press release|presentation)',t,re.I)
    if i: t=t[:i.start()]
    j=re.search(r'FORWARD-LOOKING STATEMENTS',t)
    if j: t=t[:j.start()]
    return t
EXCL=[r'(share|offering|repurchase|exercise|purchase|conversion|stock|closing|strike|issue) prices?', r'prices? per (class a )?share', r'priced (the|its|an?) (offering|notes)', r'pricing of (the|its|an?|upsized)? ?(private |public )?(offering|notes)', r'announces pricing', r'market price', r'priced at sofr']
def count(text):
    low=text.lower(); mask=list(low); out={}
    for p in EXCL:
        for m in re.finditer(p,low,re.I):
            for k in range(m.start(),m.end()): mask[k]='#'
    # supply family first, then mask so "demand exceeds supply" isn't also counted as demand
    for fam in ['supply_constrained','demand_pipeline','pricing_competition']:
        n=0; hits={}
        for p in FAM[fam]:
            s=''.join(mask)
            for m in re.finditer(p,s,re.I):
                n+=1; hits[m.group(0)]=hits.get(m.group(0),0)+1
                for k in range(m.start(),m.end()): mask[k]='#'
        out[fam]=(n,hits)
    neg=sum(len(re.findall(p,low,re.I)) for p in NEG_PRICE)
    pos=sum(len(re.findall(p,low,re.I)) for p in POS_PRICE)
    return out,(neg,pos)
def run():
    os.makedirs(A+'mined_text',exist_ok=True); rows=[]
    for q,kind,key,f in DOCS:
        raw=re.sub(r'\s+',' ',open(f,errors='ignore').read())
        if kind=='call':
            raw=raw.split('\n',1)[-1] if raw.startswith('SOURCE') else raw
        t=strip_boiler(raw,kind)
        words=len(re.findall(r"[A-Za-z][A-Za-z'\-]*",t))
        c,neg=count(t)
        open(A+f'mined_text/{q}_{kind.strip("*")}.txt','w').write(f'SOURCE: {S[key][0]}\nDATE: {S[key][1]}\nTYPE: {kind}\nWORDS (alpha tokens, boilerplate stripped): {words}\n\n{t}')
        rows.append(dict(q=q,kind=kind,src=key,url=S[key][0],date=S[key][1],words=words,
            **{fam:c[fam][0] for fam in c},**{fam+'_per10k':round(c[fam][0]/words*1e4,1) for fam in c},
            neg_price=neg[0],neg_price_per10k=round(neg[0]/words*1e4,1),pos_price=neg[1],pos_price_per10k=round(neg[1]/words*1e4,1),**{fam+'_hits':json.dumps(c[fam][1]) for fam in c}))
    json.dump(rows,open(A+'scripts/language_results.json','w'),indent=1)
    return rows
if __name__=='__main__':
    for r in run(): print(r['q'],r['kind'],r['words'],r['supply_constrained_per10k'],r['demand_pipeline_per10k'],r['pricing_competition_per10k'],'neg',r['neg_price'],'pos',r['pos_price'])
