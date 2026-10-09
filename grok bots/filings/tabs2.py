# Executed inside build.py (shares globals: wb, S, url, cite, BLUE, BOLD, HDR, Font, PatternFill, get_column_letter)
import json as _json
from openpyxl.styles import Alignment, Border, Side
GREEN=Font(color='008000'); BLK=Font(color='000000')
TITLE=Font(bold=True,size=13); SEC=PatternFill('solid',fgColor='DCE6F1'); TOT=Font(bold=True)
def hdr(ws,r,vals):
    for j,v in enumerate(vals):
        c=ws.cell(r,j+1,v); c.font=BOLD; c.fill=HDR; c.alignment=Alignment(wrap_text=True,vertical='top')
def sec(ws,r,text,ncol=8):
    ws.cell(r,1,text).font=Font(bold=True,size=11)
    for j in range(1,ncol+1): ws.cell(r,j).fill=SEC
def src(k,ref=''): return cite(k,ref)

# ================= Indenture_Terms =================
it=wb.create_sheet('Indenture_Terms')
it.append(['Convertible note indenture terms (all values Reported from indenture exhibits unless tagged). Settlement / redemption / make-whole / fundamental change.'])
it['A1'].font=TITLE
hdr(it,2,['Series','Indenture (key)','Default settlement method','Settlement options at company election','Net-share settlement (Combination Settlement)','Dutch withholding override (Sec. 3.10)','Holder conversion before earliest redemption date','Provisional (soft) call: earliest redemption date','Soft-call trigger','Redemption price','Partial redemption limit','Fundamental change put','Holder put on fixed dates','Make-whole: table range (share price)','Make-whole: max conversion rate (sh per $1k)','Initial conv. rate','Capped call','Tag','Source'])
MWJ=_json.load(open('/workspace/nebius/filings/edgar/makewhole.json'))
INDS=[('Jun-2029','2.00% due Jun-2029','IND29','Dec 10, 2026','130% x conv. price x Accretion Ratio, >=20 of 30 consecutive trading days',19.4363,'not included (June 2025 indentures pre-date Sec. 3.10 language; verify)'),
('Jun-2031','3.00% due Jun-2031','IND31','Jun 10, 2027','130% x conv. price x Accretion Ratio, >=20 of 30 trading days',19.4363,'not included (June 2025 indentures; verify)'),
('Sep-2030','1.00% due Sep-2030','IND30S','Sep 20, 2028','130% x conv. price x Accretion Ratio, >=20 of 30 trading days',7.2072,'yes'),
('Sep-2032','2.75% due Sep-2032','IND32','Sep 20, 2028','130% x conv. price x Accretion Ratio, >=20 of 30 trading days',7.2072,'yes'),
('Mar-2031','1.25% due Mar-2031','IND31M','Mar 20, 2029','130% x conv. price x Accretion Ratio, >=20 of 30 trading days',5.4579,'yes'),
('Mar-2033','2.625% due Mar-2033','IND33','Mar 20, 2030','130% x conv. price x Accretion Ratio, >=20 of 30 trading days',5.5460,'yes'),
('Feb-2030','0.50% due Feb-2030','IND30F','Feb 21, 2028','130% x conv. price x Accretion Ratio, >=20 of 30 trading days',3.1902,'yes'),
('Feb-2034','4.50% due Feb-2034','IND34','Aug 21, 2028','150% if redeemed Aug 21, 2028 - Aug 20, 2029; 130% thereafter (x conv. price x Accretion Ratio, >=20 of 30 trading days)',3.0802,'yes')]
for key,name,k,date,trig,rate,w in INDS:
    m=MWJ[key]
    it.append([name,f"{k}: {S[k][3]}",'"Default Settlement Method means Physical Settlement" (shares)','Physical, Cash or Combination Settlement; company may change the default and may irrevocably fix or eliminate a method (Sec. 5.03(A))',
      'Available: Combination Settlement = cash up to a Specified Dollar Amount per $1,000 original principal + shares for the excess. Not the default.',
      w,'Only on specified conditions (130% stock-price test, 98% trading-price test, certain corporate events); freely in final 2 months (per Q2FS Note 12)',
      date,trig,'Cash = Accreted Principal Amount + accrued & unpaid interest','Partial redemption only if >= $100M accreted principal remains outstanding',
      'Cash = Accreted Principal Amount + accrued & unpaid interest (Sec. 4.02); make-whole additional shares on conversion in connection with a Make-Whole Fundamental Change',
      'none found (text search for optional repurchase/put dates: 0 hits)',
      f"{m['hdr'][0]} to {m['hdr'][-1]} ({len(m['rows'])} effective dates)",float(m['max']),rate,'none disclosed','Reported',src(k)])
it.append([])
it.append(['Notes: (1) The FY2025 20-F states for the Sep-25 notes "Upon conversion, the Notes are settled entirely in shares" and for the Jun-25 notes "settled in shares" (20-F p.F-47/48) - consistent with the Physical default. (2) Q2FS Note 1: "In accordance with ASU 2020-06, share settlement is presumed" for diluted EPS. (3) Accretion: conversion rate is per $1,000 ORIGINAL principal; accretion raises the cash claim (redemption/put/maturity) but not the share count. (4) Aug-26 exchange (Ex.4.5 form): $4,403.18 per $1,000 of 2029s and $4,413.18 per $1,000 of 2031s + accrued interest from 6/5/26, paid in shares at the 8/19/26 close ($223.90); parity = 19.4363 x 223.90 = $4,351.79 -> ~1.2% premium (Derived).', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', src('EXAG')])
for col,w in zip('ABCDEFGHIJKLMNOPQRS',[20,30,26,40,40,34,40,16,40,30,30,40,22,26,14,12,12,10,60]): it.column_dimensions[col].width=w
for row in it.iter_rows(min_row=2):
    for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')

# ================= MakeWhole_Tables =================
mwt=wb.create_sheet('MakeWhole_Tables')
mwt.append(['Make-whole tables: additional shares per $1,000 original principal on conversion in connection with a Make-Whole Fundamental Change (Reported, verbatim from indentures; straight-line interpolation per indenture; zero above top price; capped at the max conversion rate)'])
mwt['A1'].font=TITLE
KEYS={'Jun-2029':'IND29','Jun-2031':'IND31','Sep-2030':'IND30S','Sep-2032':'IND32','Mar-2031':'IND31M','Mar-2033':'IND33','Feb-2030':'IND30F','Feb-2034':'IND34'}
for key,m in MWJ.items():
    mwt.append([]); r=mwt.max_row+1
    mwt.cell(r,1,f"{key} notes - max conversion rate {m['max']} - source: {src(KEYS[key])}").font=BOLD
    mwt.append(['Effective date \\ Share price ($)']+[float(h.replace('$','').replace(',','')) for h in m['hdr']])
    for c in mwt[mwt.max_row]: c.font=BOLD; c.fill=HDR
    for row in m['rows']:
        mwt.append([row[0]]+[float(x) for x in row[1:]])
mwt.column_dimensions['A'].width=34

# ================= CapStructure =================
cs=wb.create_sheet('CapStructure')
cs['A1']='Nebius Group (NBIS) - capital structure & dilution waterfall. $M unless stated; shares in units. Blue = hard-coded input (sourced); black = formula. As of 30-Jun-2026 balance sheet + disclosed post-period events through 1-Oct-2026.'
cs['A1'].font=TITLE
cs['A3']='Share price input ($/share)'; cs['B3']=242.81; cs['B3'].font=Font(color='0000FF',bold=True); cs['B3'].number_format='$#,##0.00'
cs['C3']='Input - Oct 2, 2026 NBIS close as provided by user (not from an SEC filing; change freely)'
cs['A4']='Price date'; cs['B4']='2026-10-02'; cs['B4'].font=BLUE
P='$B$3'
# ---- Convert dilution table
sec(cs,6,'A. Convertible notes - share count at input price (feeds waterfall)',13)
hdr(cs,7,['Series','Remaining orig. principal ($M)','Conv. rate (sh/$1k orig)','Conv. price ($)','ITM at input price?','If-converted shares','If-converted shares (ITM only)','NSS: cash per $1k (input)','NSS shares (Combination, cash = H)','Conversion value ($M)','Accreted principal at maturity per $1k','Tag','Source / derivation'])
CV=[('2.00% Jun-2029',87.5,19.4363,51.45,1200,'Derived: 500 - 12.5 converted (2025) - 400 exchanged (Aug-26)','IND29','JUN25; 20F25 p.F-48; AUG26C'),
('3.00% Jun-2031',87.5,19.4363,51.45,1250,'Derived: 500 - 12.5 - 400','IND31','JUN25; 20F25 p.F-48; AUG26C'),
('1.00% Sep-2030',1581.25,7.2072,138.75,1150,'Derived: Note 12 accreted 1,818.4 / 1.15 = 1,581.2','IND30S','SEP25C; Q2FS Note 12'),
('2.75% Sep-2032',1581.25,7.2072,138.75,1150,'Derived: Note 12 accreted 1,818.4 / 1.15','IND32','SEP25C; Q2FS Note 12'),
('1.25% Mar-2031',2587.5,5.4579,183.22,1200,'Reported','IND31M','MAR26C; Q2FS Note 12'),
('2.625% Mar-2033',1750,5.5460,180.31,1200,'Reported','IND33','MAR26C; Q2FS Note 12'),
('0.50% Feb-2030',3450,3.1902,313.46,1100,'Reported','IND30F','AUG26C'),
('4.50% Feb-2034',2300,3.0802,324.65,1250,'Reported','IND34','AUG26C')]
cv0=8
for i,(n,pr,rt,cp,acc,tag,k,sk) in enumerate(CV):
    r=cv0+i
    cs.cell(r,1,n); cs.cell(r,2,pr).font=BLUE; cs.cell(r,3,rt).font=BLUE; cs.cell(r,4,cp).font=BLUE
    cs.cell(r,5,f'=IF({P}>D{r},"ITM","OTM")')
    cs.cell(r,6,f'=B{r}*1000*C{r}')
    cs.cell(r,7,f'=IF(E{r}="ITM",F{r},0)')
    cs.cell(r,8,1000).font=BLUE
    cs.cell(r,9,f'=B{r}*1000*MAX(0,C{r}-H{r}/{P})')
    cs.cell(r,10,f'=F{r}*{P}/1000000')
    cs.cell(r,11,acc).font=BLUE
    cs.cell(r,12,tag); cs.cell(r,13,f"{src(k)} ; also {sk}")
rT=cv0+len(CV)
cs.cell(rT,1,'Total converts').font=TOT
for col in 'BFGIJ': cs[f'{col}{rT}']=f'=SUM({col}{cv0}:{col}{rT-1})'; cs[f'{col}{rT}'].font=TOT
cs.cell(rT,12,'Derived'); cs.cell(rT,13,'Sum of rows. Conversion shares ignore make-whole additional shares (only on a Make-Whole Fundamental Change or redemption notice; see MakeWhole_Tables).')
cs.cell(rT+1,1,'NSS = net-share settlement via Combination Settlement: company pays H per $1k in cash, shares only for conversion value above H. Default H = $1,000 (original principal). Set H to column K to model cash = full accreted principal. Default settlement in all indentures is Physical (shares).')
# ---- Waterfall
w0=rT+3
sec(cs,w0,'B. Dilution waterfall at input price',8)
hdr(cs,w0+1,['Step','Shares - base case (ITM converts if-converted)','Shares - all converts if-converted (incl. OTM)','Shares - ITM converts net-share settled (cash = H)','Cumulative (base case)','Tag','Source / derivation','% of basic (base case)'])
W=[]
def wrow(label,b,c=None,d=None,tag='',source='',bold=False,cum=True):
    W.append((label,b,c if c is not None else b,d if d is not None else b,tag,source,bold,cum))
r=w0+2
rows_idx={}
lines=[
 ('iss','Shares issued & outstanding 30-Jun-2026 (Class A 238,400,165 + Class B 33,455,053; excl. 50,185,726 treasury)',271855218,None,None,'Reported',src('Q2PR','Outstanding shares'),False),
 ('exch','+ Aug-2026 exchange shares for $400M 2029s + $400M 2031s (post-period)','=ROUNDDOWN((4403.18*400000+4413.18*400000+400000000*0.02*79/360+400000000*0.03*79/360)/223.9,0)',None,None,'Derived','Ex.4.5 formula: (exchange price x principal + accrued interest 6/5-8/24/26, 30/360 = 79 days) / $223.90 close on 8/19/26. Company PR: "approximately 15.8 million". '+src('EXAG')+' ; '+src('AUG26P'),False),
 ('nvda','+ NVIDIA pre-funded warrant shares (exercise $0.0001; already in BASIC weighted shares; NOT in issued count - counted once here)',21065936,None,None,'Reported',src('Q2FS','Note 14, p.F-31')+' ; '+src('Q2FS','Note 1 EPS policy, p.F-11'),False),
 ('basic','Subtotal: Basic shares (economic, incl. warrant)',None,None,None,'Derived','sum of above',True),
 ('rsu','+ Unvested RSUs (no exercise price -> full count)',6234091,None,None,'Reported',src('Q2FS','Note 13, p.F-29')+' (unvested RSUs 30-Jun-2026; WA grant FV $58.22)',False),
 ('opt','+ Options, treasury stock method: N x (1 - WAEP/P) if P > WAEP',f'=IF({P}>88.65,6740600*(1-88.65/{P}),0)',None,None,'Derived','6,740,600 options, WAEP $88.65: '+src('Q2FS','Note 13, p.F-29')+'. Aggregate-WAEP shortcut (tranche-level strikes not disclosed) -> estimate.',False),
 ('awards','Subtotal: Basic + equity awards',None,None,None,'Derived','sum',True),
]
for i,(n,pr,rt,cp,acc,tag,k,sk) in enumerate(CV):
    rr=cv0+i
    lines.append((f'cv{i}',f'+ {n} convert (conv. price ${cp:,.2f})',f'=G{rr}',f'=F{rr}',f'=IF(E{rr}="ITM",I{rr},0)','Derived',f'Section A row {rr}: remaining orig principal x conv. rate; OTM = 0 in base case',False))
lines += [
 ('postcv','Subtotal: After converts',None,None,None,'Derived','sum',True),
 ('atm','+ ATM program remaining capacity (25,000,000 authorized - 12,729,493 sold in Q2-26; no later sales disclosed)',12270507,None,None,'Reported',src('Q2FS','Note 14, p.F-32 ("12,270,507 Class A shares remained available")'),False),
 ('fd','TOTAL: FULLY DILUTED SHARES',None,None,None,'Derived','sum',True),
]
first=r
for key,label,b,c,d,tag,source,is_sum in lines:
    rows_idx[key]=r
    cs.cell(r,1,label)
    if is_sum:
        # sum from the previous subtotal (or start) to r-1
        prev=[v for kk,v in rows_idx.items() if kk in('basic','awards','postcv') and v<r]
        start=first
        for col in 'BCD':
            cs[f'{col}{r}']=f'=SUM({col}{start}:{col}{r-1})-SUMPRODUCT(({chr(64+6)}{start}:{chr(64+6)}{r-1}="SUBTOTAL")*({col}{start}:{col}{r-1}))'
        # simpler: subtotal rows flagged; recompute below
        for col in range(1,9): cs.cell(r,col).font=TOT
    else:
        cs.cell(r,2,b); cs.cell(r,3,c if c is not None else f'=B{r}'); cs.cell(r,4,d if d is not None else f'=B{r}')
        if isinstance(b,(int,float)): cs.cell(r,2).font=BLUE
    cs.cell(r,6,tag); cs.cell(r,7,source)
    r+=1
# rewrite subtotal formulas cleanly: each subtotal = previous subtotal + items in between
subs=['basic','awards','postcv','fd']
prevrow=None
for s_ in subs:
    rr=rows_idx[s_]
    lo=(prevrow+1) if prevrow else first
    for col in 'BCD':
        cs[f'{col}{rr}']=(f'={col}{prevrow}+SUM({col}{lo}:{col}{rr-1})' if prevrow else f'=SUM({col}{lo}:{col}{rr-1})')
    prevrow=rr
# cumulative and % columns
b_row=rows_idx['basic']
for key,rr in rows_idx.items():
    cs[f'E{rr}']=(f'=B{rr}' if (key in subs or rr==first) else f'=E{rr-1}+B{rr}')
    cs[f'H{rr}']=f'=B{rr}/$B${b_row}'
    cs[f'H{rr}'].number_format='0.0%'
for rr in range(first,r):
    for col in 'BCDE': cs[f'{col}{rr}'].number_format='#,##0'
fd=rows_idx['fd']
# Summary
s0=r+1
sec(cs,s0,'C. Summary at input price',8)
summ=[
 ('Basic shares (economic)',f'=B{b_row}',f'=C{b_row}',f'=D{b_row}'),
 ('Fully diluted shares',f'=B{fd}',f'=C{fd}',f'=D{fd}'),
 ('Dilution vs basic (shares)',f'=B{fd}-B{b_row}',f'=C{fd}-C{b_row}',f'=D{fd}-D{b_row}'),
 ('Dilution vs basic (%)',f'=B{fd}/B{b_row}-1',f'=C{fd}/C{b_row}-1',f'=D{fd}/D{b_row}-1'),
 ('FD excl. ATM remaining',f'=B{fd}-B{rows_idx["atm"]}',f'=C{fd}-C{rows_idx["atm"]}',f'=D{fd}-D{rows_idx["atm"]}'),
 ('Equity value at input price, basic ($M)',f'=B{b_row}*{P}/1000000',f'=C{b_row}*{P}/1000000',f'=D{b_row}*{P}/1000000'),
 ('Equity value at input price, fully diluted ($M)',f'=B{fd}*{P}/1000000',f'=C{fd}*{P}/1000000',f'=D{fd}*{P}/1000000'),
 ('Memo: option exercise proceeds if ITM ($M)',f'=IF({P}>88.65,6740600*88.65/1000000,0)',None,None),
 ('Memo: ATM gross proceeds if fully sold at input price ($M, before commissions)',f'=B{rows_idx["atm"]}*{P}/1000000',None,None),
 ('Memo: converts ITM, original principal ($M) retired on conversion',f'=SUMPRODUCT((E{cv0}:E{rT-1}="ITM")*B{cv0}:B{rT-1})',None,None),
 ('Memo: NSS cash outlay on ITM converts at H ($M)',f'=SUMPRODUCT((E{cv0}:E{rT-1}="ITM")*B{cv0}:B{rT-1}*H{cv0}:H{rT-1})/1000',None,None),
]
hdr(cs,s0+1,['Metric','Base case (ITM if-converted)','All converts if-converted','ITM NSS','','Tag','Notes'])
rr=s0+2
for lab,b,c,d in summ:
    cs.cell(rr,1,lab); cs.cell(rr,2,b)
    if c: cs.cell(rr,3,c)
    if d: cs.cell(rr,4,d)
    cs.cell(rr,6,'Derived')
    fmt='0.0%' if '(%)' in lab else ('#,##0.0' if '$M' in lab else '#,##0')
    for col in 'BCD': cs[f'{col}{rr}'].number_format=fmt
    rr+=1
cs.cell(rr,1,'Not in waterfall (disclosed but excluded): (i) 2,743,655 unvested Eigen AI restricted shares are legally issued and assumed inside the 271.9M issued count (excluded from weighted EPS counts) - not added again; (ii) make-whole additional shares (event-driven); (iii) awards granted after 30-Jun-2026 and any ATM sales after 30-Jun-2026 - not disclosed; (iv) Inferize acquisition (6-K 10/1/26) - "The terms of the transaction were not disclosed"; (v) AGM record-date Class A count 238,402,543 vs 238,400,165 at 30-Jun (+2,378) - immaterial, not added; (vi) no capped calls exist to offset convert dilution.')
cs.cell(rr+1,1,'Cross-check (Derived): Q2FS Note 1 reports 65,416,804 weighted if-converted shares excluded as anti-dilutive in Q2-26 for the six pre-Aug series. At initial rates on pre-exchange principal: 2029 487.5 x 19.4363 = 9.475M; 2031 9.475M; Sep-30 11.396M; Sep-32 11.396M; Mar-31 14.122M; Mar-33 9.706M; total 65.57M (within 0.2% of the reported figure) -> supports the remaining-principal inputs.')
rr+=3
# ---- Debt & financing register
d0=rr
sec(cs,d0,'D. Debt & financing register (30-Jun-2026 balance sheet + post-period)',13)
hdr(cs,d0+1,['Instrument','Date','Original principal ($M)','Remaining orig principal ($M)','Coupon / rate','Maturity','Accreted principal at maturity 30-Jun ($M)','Carrying 30-Jun ($M)','Fair value 30-Jun ($M)','Security / ranking','Prepay / earliest call','Tag','Source'])
DEBT=[('2.00% conv. notes Jun-2029','2025-06-05',500,87.5,'2.00%','2029-06-05',587.5,504.1,2928.7,'Senior unsecured','Soft call from 2026-12-10','Reported; remaining Derived','IND29'),
('3.00% conv. notes Jun-2031','2025-06-05',500,87.5,'3.00%','2031-06-05',612.5,498.4,2908.4,'Senior unsecured','Soft call from 2027-06-10','Reported; remaining Derived','IND31'),
('1.00% conv. notes Sep-2030','2025-09-15',1581.25,1581.25,'1.00%','2030-09-15',1818.4,1588.2,3555.5,'Senior unsecured','Soft call from 2028-09-20','Derived principal (1,818.4/1.15)','IND30S'),
('2.75% conv. notes Sep-2032','2025-09-15',1581.25,1581.25,'2.75%','2032-09-15',1818.4,1575.0,3488.0,'Senior unsecured','Soft call from 2028-09-20','Derived principal','IND32'),
('1.25% conv. notes Mar-2031','2026-03-20',2587.5,2587.5,'1.25%','2031-03',3105.0,2588.6,4696.5,'Senior unsecured','Soft call from 2029-03-20','Reported','IND31M'),
('2.625% conv. notes Mar-2033','2026-03-20',1750,1750,'2.625%','2033-03',2100.0,1744.7,3246.6,'Senior unsecured','Soft call from 2030-03-20','Reported','IND33'),
('0.50% conv. notes Feb-2030 (post-period)','2026-08-24',3450,3450,'0.50%','2030-02-15','n/a (post-period; 110% accretion = 3,795.0 Derived)','n/a (post-period)','n/a (post-period)','Senior unsecured','Soft call from 2028-02-21','Reported','IND30F'),
('4.50% conv. notes Feb-2034 (post-period)','2026-08-24',2300,2300,'4.50%','2034-02-15','n/a (post-period; 125% accretion = 2,875.0 Derived)','n/a (post-period)','n/a (post-period)','Senior unsecured','Soft call from 2028-08-21','Reported','IND34'),
('Senior secured asset-backed term loan (Nebius Compute II LLC / Oy; MUFG lead)','2026-07-10',775,'drawn amount not disclosed','Term SOFR (0% floor) + 2.50%','2030-10-31','n/a','n/a (post-period)','n/a','Secured by substantially all borrower assets + borrower shares; collateralized by deployed GPU infrastructure and contracted cash flows from an investment-grade customer; DSCR 1.15x + min. liquidity; non-recourse to parent except bad-acts/performance guarantees','Prepayable without premium (breakage only)','Reported','ABL'),
('Debt, current (30-Jun-2026 balance sheet)','2026-06-30','n/a','n/a','not disclosed','within 12 months','n/a',46.7,'not disclosed','not disclosed (no breakdown in FS)','n/a','Reported','Q2PR'),
]
rr=d0+2; dfirst=rr
for row in DEBT:
    for j,v in enumerate(row[:12]): cs.cell(rr,j+1,v)
    cs.cell(rr,13,src(row[12]) + (' ; '+src('Q2FS','Note 12, p.F-28') if row[12].startswith('IND') and 'post' not in row[0] else ''))
    for j in (3,4,7,8,9):
        if isinstance(cs.cell(rr,j).value,(int,float)): cs.cell(rr,j).font=BLUE
    rr+=1
cs.cell(rr,1,'Total converts (pro forma, post Aug-26)').font=TOT
cs[f'C{rr}']=f'=SUM(C{dfirst}:C{dfirst+7})'; cs[f'D{rr}']=f'=SUM(D{dfirst}:D{dfirst+7})'; cs[f'G{rr}']=f'=SUM(G{dfirst}:G{dfirst+5})'; cs[f'H{rr}']=f'=SUM(H{dfirst}:H{dfirst+5})'; cs[f'I{rr}']=f'=SUM(I{dfirst}:I{dfirst+5})'
cs.cell(rr,12,'Derived'); cs.cell(rr,13,'G/H/I = 30-Jun-2026 six series only (Note 12 totals 10,041.8 / 8,499.0 / 20,823.7); C/D include Aug-26 issues; D is after Aug-26 exchange (G/H/I still include the 800M exchanged)')
tot_cv=rr; rr+=1
cs.cell(rr,1,'Cash & equivalents 30-Jun-2026'); cs.cell(rr,8,8042.1).font=BLUE; cs.cell(rr,12,'Reported'); cs.cell(rr,13,src('Q2PR','Balance sheet')); cash=rr; rr+=1
cs.cell(rr,1,'Restricted cash 30-Jun-2026'); cs.cell(rr,8,1056.0).font=BLUE; cs.cell(rr,12,'Reported'); cs.cell(rr,13,src('Q2PR','Balance sheet')); rcash=rr; rr+=1
cs.cell(rr,1,'Aug-2026 convert gross proceeds (post-period)'); cs[f'H{rr}']=f'=C{dfirst+6}+C{dfirst+7}'; cs.cell(rr,12,'Derived'); cs.cell(rr,13,'3,450 + 2,300 gross; net proceeds per AUG26C PR not re-stated here'); augp=rr; rr+=1
cs.cell(rr,1,'Net debt 30-Jun-2026, carrying basis ($M) = debt current + convert carrying - cash (ex restricted)'); cs[f'H{rr}']=f'=46.7+H{tot_cv}-H{cash}'; cs.cell(rr,12,'Derived'); cs.cell(rr,13,'46.7 + 8,499.0 - 8,042.1'); rr+=1
cs.cell(rr,1,'Pro forma remaining orig. convert principal ($M, post Aug-26)'); cs[f'H{rr}']=f'=D{tot_cv}'; cs.cell(rr,12,'Derived'); rr+=2
# ---- Leases & commitments
l0=rr
sec(cs,l0,'E. Lease obligations (ASC 842 - US GAAP filer; all operating leases) and other commitments',13)
hdr(cs,l0+1,['Item','31-Dec-2025','30-Jun-2026','','','','','','','','','Tag','Source'])
LE=[('Operating lease ROU assets',918.8,1855.1),('Operating lease liabilities - current',84.9,114.9),('Operating lease liabilities - non-current',760.5,1510.4),('Total operating lease liabilities',845.4,1625.3),
('Weighted avg remaining lease term (years)',8.6,8.7),('Weighted avg discount rate (%)',6.3,6.3),('Operating lease cost - 1H (Q2FS) / FY (20-F)','not extracted for FY',99.7),
('Leases signed, not yet commenced - undiscounted payments',9759.6,12054.1),('Prepayments on not-yet-commenced leases',271.0,612.4)]
rr=l0+2
for lab,a,b in LE:
    cs.cell(rr,1,lab); cs.cell(rr,2,a); cs.cell(rr,3,b)
    for j in (2,3):
        if isinstance(cs.cell(rr,j).value,(int,float)): cs.cell(rr,j).font=BLUE
    cs.cell(rr,12,'Reported'); cs.cell(rr,13,src('20F25','Note 8, p.F-39/F-40')+' ; '+src('Q2FS','Note 8, p.F-24/F-25')); rr+=1
cs.cell(rr,1,'Finance leases'); cs.cell(rr,2,'none - "all of which are operating leases"'); cs.cell(rr,3,'none'); cs.cell(rr,12,'Reported'); cs.cell(rr,13,src('Q2FS','Note 8, p.F-24')); rr+=2
hdr(cs,rr,['Lease maturity table (undiscounted, $M)','As of 31-Dec-2025','As of 30-Jun-2026','','','','','','','','','Tag','Source']); rr+=1
MT=[('2026 (remainder of 2026 for Jun-26 column)',135.8,91.6),('2027',124.3,235.2),('2028',120.1,234.2),('2029',125.7,241.4),('2030',123.6,242.7),('Thereafter',482.0,1130.0)]
mt0=rr
for lab,a,b in MT:
    cs.cell(rr,1,lab); cs.cell(rr,2,a).font=BLUE; cs.cell(rr,3,b).font=BLUE; cs.cell(rr,12,'Reported'); cs.cell(rr,13,src('20F25','Note 8, p.F-39')+' ; '+src('Q2FS','Note 8, p.F-24')); rr+=1
cs.cell(rr,1,'Total lease payments'); cs[f'B{rr}']=f'=SUM(B{mt0}:B{rr-1})'; cs[f'C{rr}']=f'=SUM(C{mt0}:C{rr-1})'; cs.cell(rr,12,'Derived (ties to reported 1,111.5 / 2,175.1)'); tp=rr; rr+=1
cs.cell(rr,1,'Less imputed interest'); cs.cell(rr,2,-266.1).font=BLUE; cs.cell(rr,3,-549.8).font=BLUE; cs.cell(rr,12,'Reported'); rr+=1
cs.cell(rr,1,'Lease liabilities (PV)'); cs[f'B{rr}']=f'=B{tp}+B{rr-1}'; cs[f'C{rr}']=f'=C{tp}+C{rr-1}'; cs.cell(rr,12,'Derived (ties to 845.4 / 1,625.3)'); rr+=2
cs.cell(rr,1,'Energy & power supply commitments (fixed/minimum guaranteed), as of 30-Jun-2026 ($M)'); cs.cell(rr,12,'Reported'); cs.cell(rr,13,src('Q2FS','Note 11, p.F-26')); rr+=1
PW=[('2026 (remainder)',81.7),('2027',541.8),('2028',542.0),('2029',540.5),('2030',535.8),('Thereafter',3082.3)]
pw0=rr
for lab,v in PW: cs.cell(rr,1,lab); cs.cell(rr,3,v).font=BLUE; cs.cell(rr,12,'Reported'); rr+=1
cs.cell(rr,1,'Total power commitments'); cs[f'C{rr}']=f'=SUM(C{pw0}:C{rr-1})'; cs.cell(rr,12,'Derived'); rr+=1
cs.cell(rr,1,'GPU / equipment vendor financing, sale-leaseback or other asset-backed financing'); cs.cell(rr,3,'not disclosed beyond the $775M secured facility (none found in 20-F FY25, Q1/Q2 FS)'); cs.cell(rr,12,'Not disclosed'); rr+=2
# ---- Equity raises history
e0=rr
sec(cs,e0,'F. Equity issuance / share count history',13)
hdr(cs,e0+1,['Event','Date','Shares','Price ($/sh)','Gross ($M)','Net ($M)','Investors / mechanism','','','','','Tag','Source'])
EQ=[('Divestment consideration shares received into treasury (deemed repurchase)','2024-05-17 / 2024-07-12',162485725,'14.1 WA cost (RUB 1,251.8)','not disclosed in $','n/a','67,632,122 at first closing + 94,853,603 at second; held in treasury','Reported',src('20F24','p.79; Note 13 p.F-36')),
('Private placement (from treasury)','2024-12-02',33333334,21.00,700,667.5,'Accel, NVIDIA, Orbis-managed accounts + others; Accel observer seat','Reported',src('PP24')+' ; '+src('20F24','Note 13 p.F-36')),
('Conversion of Jun-25 notes (one holder)','2025 (H2)',485906,'51.45 (conv.)','25.0 principal','n/a','242,953 shares each for $12.5M of 2029s and 2031s','Reported',src('20F25','p.F-48')),
('Underwritten follow-on (concurrent with Sep-25 converts)','2025-09-10 pricing / 09-15 close',10810811,92.50,1000.0,'not disclosed','Public offering','Reported',src('SEP25C')),
('Follow-on greenshoe (exercised in full)','2025-09-22',1621621,92.50,150.0,'not disclosed','Underwriters option; total Class A offering gross ~$1.15B','Reported (gross derived 1,621,621 x 92.50)',src('SEP25G')),
('ATM program established (25,000,000 shares)','2025-11-12','25,000,000 authorized','market','n/a','n/a','Equity distribution agreement; sales began May 2026','Reported',src('Q2FS','Note 14, p.F-32')),
('NVIDIA pre-funded warrant','2026-03-11',21065936,'$0.0001 exercise; implied $94.94 purchase price (Derived 2,000/21.07M)',2000.0,'not disclosed','Securities purchase agreement; no expiry; HSR exercise blocker; 13G 7/20/26 shows 22,256,412 sh beneficial (9.3%)','Reported; implied price Derived',src('NVDA')+' ; '+src('Q2FS','Note 14, p.F-31')+' ; '+src('NVDA13G')),
('ATM sales (from treasury)','Q2 2026 (May-Jun)',12729493,223.60,'2,846.3 (Derived 12,729,493 x 223.60)',2811.1,'At-the-market; 12,270,507 remain','Reported; gross Derived',src('Q2FS','Note 14, p.F-32')),
('Eigen AI acquisition - consideration shares','2026-06-10','~1.4 million',211.68,298.7,'n/a','Purchase consideration','Reported',src('Q2FS','Note 3, p.F-16')),
('Eigen AI - employee restricted shares (4-yr vesting)','2026-06-10','~2.4 million per Note 3; 2,743,655 unvested per Note 1','n/a','508.1 FV (post-combination SBC)','n/a','Excluded from weighted counts; participating securities','Reported (two figures as disclosed)',src('Q2FS','Note 3 p.F-16; Note 1 p.F-11')),
('Exchange of $400M 2029s + $400M 2031s for shares','2026-08-24','~15.8 million (Derived 15.77M)','223.90 reference close',800.0,'n/a','Privately negotiated exchange; ~1.2% premium to parity','Reported / Derived',src('AUG26P')+' ; '+src('EXAG')),
('Treasury shares remaining 30-Jun-2026','2026-06-30',50185726,'n/a','n/a','n/a','Available to fund ATM/awards without new issuance','Reported',src('Q2PR','Outstanding shares')),
('Buybacks after divestiture','2024-2026','none disclosed','n/a','n/a','n/a','AGM 8/25/26 renewed 18-month authority to repurchase up to 20% and 5-yr authority to issue up to 20%','Reported',src('AGM'))]
rr=e0+2
for row in EQ:
    for j,v in enumerate(row[:7]): cs.cell(rr,j+1,v)
    cs.cell(rr,12,row[7]); cs.cell(rr,13,row[8])
    for j in (3,4,5,6):
        if isinstance(cs.cell(rr,j).value,(int,float)): cs.cell(rr,j).font=BLUE
    rr+=1
cs.column_dimensions['A'].width=62
for col,wd in zip('BCDEFGHIJKLM',[22,22,22,20,16,40,18,16,30,18,26,80]): cs.column_dimensions[col].width=wd
CAPROWS=dict(fd=fd,basic=b_row,first=first,cv0=cv0,rT=rT,rows=rows_idx)
