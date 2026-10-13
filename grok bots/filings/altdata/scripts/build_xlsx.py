import sys,os,importlib
sys.path.insert(0,os.path.dirname(__file__))
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from sources import S
OUT='/workspace/nebius/filings/altdata/'
wb=Workbook(); first=True
H=Font(bold=True,color='FFFFFF'); HF=PatternFill('solid',fgColor='1F3864')
COL={'HIT':'C6EFCE','BEAT':'A9D08E','MISS':'FFC7CE','SLIPPED':'FFEB9C','PENDING':'DDEBF7','NOT VERIFIABLE':'E7E6E6'}
def sheet(name):
    global first
    if first: ws=wb.active; ws.title=name; first=False
    else: ws=wb.create_sheet(name)
    return ws
def header(ws,row,cols):
    for j,c in enumerate(cols,1):
        x=ws.cell(row=row,column=j,value=c); x.font=H; x.fill=HF; x.alignment=Alignment(wrap_text=True,vertical='top')
def widths(ws,w):
    for j,x in enumerate(w,1): ws.column_dimensions[get_column_letter(j)].width=x
def url(k): return S[k][0] if k in S else ''
# ---------- Guidance ----------
from guidance_data import G,W
ws=sheet('Guidance_HitRate')
ws['A1']='Nebius (NBIS) forward-looking statements log, Q3-2024 to Q2-2026 (+ pending targets). All quotes verbatim (machine-checked against source text by scripts/verify_quotes.py).'; ws['A1'].font=Font(bold=True)
ws['A2']='Status: BEAT/HIT = met by target date; MISS = not met; SLIPPED = met late (slip in quarters); PENDING = target date in future; NOT VERIFIABLE = actual never disclosed.'
res=[g for g in G if g['status'] in ('HIT','BEAT','MISS','SLIPPED')]
cats=['ARR','Revenue','Capex','MW','GPU']
r=4; header(ws,r,['Summary','Resolved','Hit/Beat','Miss','Slipped','Hit rate','Pending','Not verifiable'])
for c in cats+['ALL']:
    gg=[g for g in G if c=='ALL' or g['cat']==c]; rr=[g for g in gg if g in res]
    h=sum(g['status'] in('HIT','BEAT') for g in rr); m=sum(g['status']=='MISS' for g in rr); s=sum(g['status']=='SLIPPED' for g in rr)
    r+=1; ws.append([c,len(rr),h,m,s,(h/len(rr)) if rr else None,sum(g['status']=='PENDING' for g in gg),sum(g['status']=='NOT VERIFIABLE' for g in gg)])
    ws.cell(row=r,column=6).number_format='0%'
r+=1; ws.cell(row=r,column=1,value='Note: MW "contracted" targets (MW-8..MW-10) were each beaten early, but contracted power is the cheapest metric (~1% of capex per CEO, Nov-11-2025 call). Excluding contracted-MW rows, hit rate is shown below.')
rr=[g for g in res if g['id'] not in ('MW-8','MW-9','MW-10','MW-4')]; h=sum(g['status'] in('HIT','BEAT') for g in rr)
r+=1; ws.cell(row=r,column=1,value='ALL ex-contracted-MW'); ws.cell(row=r,column=2,value=len(rr)); ws.cell(row=r,column=3,value=h); ws.cell(row=r,column=6,value=h/len(rr)).number_format='0%'
r+=2; cols=['ID','Category','Metric','Date said','Source doc','Source URL','Exact quote','Target value','Target date','Actual reported','Actual source URL','Actual quote','Status','Slip (quarters)','Notes']
header(ws,r,cols); top=r
for g in G:
    r+=1
    vals=[g['id'],g['cat'],g['metric'],g['said'],g['src'],url(g['src']),g['quote'],g['target'],g['tdate'],g['actual'],url(g['asrc']) if g['asrc'] else ('https://assets.nebius.com/assets/6590d0a0-0fc4-4396-bc3f-f52a9ccb770e/20260331%20Nebius%20to%20construct%20310%20MW%20AI%20factory%20in%20Finland.pdf' if g['id']=='MW-1' else ''),g['aquote'],g['status'],g['slip'],g['note']]
    for j,v in enumerate(vals,1):
        x=ws.cell(row=r,column=j,value=v); x.alignment=Alignment(wrap_text=True,vertical='top')
    ws.cell(row=r,column=13).fill=PatternFill('solid',fgColor=COL.get(g['status'],'FFFFFF'))
widths(ws,[9,9,26,11,12,30,60,16,11,28,30,45,12,10,45]); ws.freeze_panes=ws.cell(row=top+1,column=4)
ws=sheet('MW_Wording')
ws['A1']='Wording-change history for Nebius MW definitions (contracted / connected / active). Verbatim quotes.'; ws['A1'].font=Font(bold=True)
header(ws,3,['Date','Source doc','Source URL','What changed','Exact quote'])
for d,k,what,q in W:
    ws.append([d,k,url(k),what,q])
for row in ws.iter_rows(min_row=4):
    for x in row: x.alignment=Alignment(wrap_text=True,vertical='top')
widths(ws,[11,13,40,45,90])
# ---------- Prepay ----------
from prepay_data import Q,R
ws=sheet('Prepay_Commitments')
ws['A1']='Contract liabilities, customer prepayments, RPO and purchase/lease commitments by quarter-end ($M unless %). Blank = not disclosed / not available.'; ws['A1'].font=Font(bold=True)
header(ws,3,['Metric']+Q+['Notes','Sources (by quarter)'])
rows={}
for m,vals,srcs,note in R:
    if m.startswith('Deferred revenue / PP&E'): continue
    line=[m]+[vals.get(q) for q in Q]+[note,' | '.join(f'{q}: {url(k)}' for q,k in srcs.items())]
    ws.append(line); rows[m]=vals
dr=rows['Deferred revenue (contract liability), total']; pp=rows['Property & equipment, net']
ws.append(['Deferred revenue / PP&E net (%) [derived]']+[round(dr[q]/pp[q]*100,1) if q in dr and q in pp else None for q in Q]+['Derived: customer cash funding as share of installed asset base.',''])
for row in ws.iter_rows(min_row=4):
    for x in row: x.alignment=Alignment(wrap_text=True,vertical='top')
    for x in row[1:11]: x.number_format='#,##0.0'
widths(ws,[46]+[10]*10+[70,80]); ws.freeze_panes='B4'
# ---------- optional sheets from other modules ----------
for mod in ['language_sheet','raise_sheet','insiders_sheet']:
    try: m=importlib.import_module(mod)
    except ModuleNotFoundError: continue
    m.write(wb,sheet,header,widths,url)
wb.save(OUT+'altdata.xlsx'); print('saved',wb.sheetnames)
