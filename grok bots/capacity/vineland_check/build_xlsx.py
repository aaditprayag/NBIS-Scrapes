import openpyxl, datetime as dt, re
from openpyxl.styles import Font, PatternFill, Alignment
wb=openpyxl.Workbook(); ws=wb.active; ws.title='Inputs'
B=Font(bold=True); Y=PatternFill('solid',fgColor='FFF2CC')
hdr=['Key','Value','Unit','Tag','Basis / formula','Source URL','Source date']
ws.append(hdr)
rows=[
('order_date',dt.date(2026,9,22),'date','Reported','NJDEP AONOCAPA dated Sep 22 2026 (Facility 75894)','https://dep.nj.gov/wp-content/uploads/newsrel/2026/dataone75894-aonocapa.pdf','2026-09-22'),
('receipt_date',dt.date(2026,9,22),'date','Estimated','Receipt date not disclosed; assume same-day (email/hand) receipt. Change to 2026-09-25 for mail receipt.','https://morristownminute.org/1-07-million-penalty-vineland-data-center/','2026-09-22'),
('deadline_45d','={receipt_date}+45','date','Derived','receipt + 45 calendar days (Nov 6 if receipt Sep 22; Nov 9 if Sep 25)','https://www.enr.com/articles/63694-nj-fines-data-center-1m-over-123-mw-of-unpermitted-generation','2026-09-23'),
('hearing_req_deadline','={receipt_date}+20','date','Derived','receipt + 20 days; penalty final day 21 if no hearing request','https://www.enr.com/articles/63694-nj-fines-data-center-1m-over-123-mw-of-unpermitted-generation','2026-09-23'),
('outage_start','={deadline_45d}+1','date','Derived','first day after the 45-day window','—','—'),
('penalty',1072000,'$','Reported','civil administrative penalty','https://www.nj.gov/governor/news/2026/20260922a.shtml','2026-09-22'),
('genset_count',62,'units','Reported','Caterpillar G3516H','https://www.enr.com/articles/63694-nj-fines-data-center-1m-over-123-mw-of-unpermitted-generation','2026-09-23'),
('genset_kw',1982,'kW each','Reported','nameplate','https://www.enr.com/articles/63694-nj-fines-data-center-1m-over-123-mw-of-unpermitted-generation','2026-09-23'),
('genset_mw_gross','={genset_count}*{genset_kw}/1000','MW gross','Derived','62 x 1.982 = 122.9 MW','—','—'),
('pue',1.2,'x','Estimated','team assumption (nebius_capacity.xlsx Vineland high formula uses /1.2)','nebius_capacity.xlsx','2026-10'),
('grid_mw_gross',5,'MW gross','Estimated','team figure: ~5 MW utility feed (nebius_capacity.xlsx Vineland high formula)','nebius_capacity.xlsx','2026-10'),
('it_ceiling',107,'MW IT','Estimated','team ceiling ~107 MW IT = (5+122.9)/1.2','nebius_capacity.xlsx','2026-10'),
('it_energized_at_deadline','=MIN({it_ceiling},4*{tranche_it_mw})','MW IT','Estimated','4 tranches delivered (T1 Nov-25, T2 Feb-26, T3-T4 Jul-Aug-26) x equal-split tranche size, capped at ceiling','https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf','2026-08-12'),
('grid_it','={grid_mw_gross}/{pue}','MW IT','Derived','grid feed / PUE','—','—'),
('genset_fed_it','=MAX(0,{it_energized_at_deadline}-{grid_it})','MW IT','Derived','energized IT minus grid-fed IT','—','—'),
('tranche_it_mw','=240/9','MW IT','Estimated','equal split of 240 MW IT over 9 tranches (sizes not disclosed)','nebius_capacity.xlsx (Vineland tranches tab)','2026-10'),
('rev_per_mw_yr',11.6,'$M/MW/yr','Estimated','team: $17.4B / 5 yrs / 300 MW (per GROSS MW)','https://www.sec.gov/ (Nebius 6-K, Sep 2025 Microsoft agreement)','2025-09-08'),
('rev_per_mw_it_yr','=17392.9/5/240','$M/MW IT/yr','Derived','sensitivity: same TCV per IT MW (240 MW IT) = $14.5M','nebius_capacity.xlsx (Vineland tranches tab)','2026-10'),
('usd_per_mwh_check','={rev_per_mw_yr}*1000000/8760','$/MWh','Derived','cross-check vs team $1,324/MWh (identical by construction)','—','—'),
('partial_curtail_pct',0.5,'%','Estimated','consent-order style cap on genset output (no permit exists, so "permitted level" is undefined); flex 25-75%','—','—'),
('nox_tons_to_date',140,'short tons','Reported','approx NOx emitted to date per Floodlight analysis of DataOne letter','https://floodlightnews.org/vineland-nj-data-center-air-pollution/','2026-10-01'),
('engine_hours_to_date',46000,'engine-hours','Reported','>46,000 run-hours per DataOne letter (Floodlight)','https://floodlightnews.org/vineland-nj-data-center-air-pollution/','2026-10-01'),
('nox_lb_per_engine_hr','={nox_tons_to_date}*2000/{engine_hours_to_date}','lb/engine-hr','Derived','140 t x 2000 / 46,000 h','—','—'),
('minor_source_nox_tpy',25,'tpy','Reported','NJ major-source NOx threshold (ozone nonattainment) cited by Floodlight','https://floodlightnews.org/vineland-nj-data-center-air-pollution/','2026-10-01'),
('engines_at_minor_cap','={minor_source_nox_tpy}*2000/{nox_lb_per_engine_hr}/8760','engines continuous','Derived','25 tpy / (lb per engine-hr x 8760)','—','—'),
('it_at_minor_cap','={engines_at_minor_cap}*{genset_kw}/1000/{pue}','MW IT','Derived','engines x 1.982 MW / PUE','—','—'),
('fy26_end',dt.date(2026,12,31),'date','Reported','calendar FY','—','—'),
('q1_27_end',dt.date(2027,3,31),'date','Reported','—','—','—'),
]
ROW={r[0]:i+2 for i,r in enumerate(rows)}
def F(x):
    if isinstance(x,str): return re.sub(r'\{(\w+)\}',lambda m:'$B$%d'%ROW[m.group(1)],x)
    return x
for r in rows: ws.append([F(v) for v in r])
for c in ws[1]: c.font=B
for r in range(2,ws.max_row+1):
    if isinstance(ws.cell(r,2).value,dt.date) or (isinstance(ws.cell(r,2).value,str) and ws.cell(r,1).value in('deadline_45d','hearing_req_deadline','outage_start')):
        ws.cell(r,2).number_format='yyyy-mm-dd'
    if ws.cell(r,4).value=='Estimated': ws.cell(r,2).fill=Y
for col,w in zip('ABCDEFG',[26,14,18,11,70,70,12]): ws.column_dimensions[col].width=w
# Scenarios
s=wb.create_sheet('Scenarios')
h=['Scenario','MW IT offline','Outage months','Outage start','Outage end','Days offline total','Days in Q4/FY26','Days in FY27','Q4\'26 rev at risk $M','FY26 rev at risk $M','FY27 rev at risk $M','Total $M','Total $M @ $14.5M/MW IT (sens.)','Tranches affected','Prob (Estimated)','Tag','Notes']
s.append(h)
sc=[
('1. No shutdown (DataOne files by deadline; DEP lets gensets run pending permit / consent order)','=0',0,'No physical impact; T1-T4 bill; risk shifts to permit conditions and Bloom timing (T5-T9)',0.55,'Floodlight Oct 1: DEP allows operation during 45 days; DataOne says it will apply'),
('2a. Partial: gensets capped at assumed % (consent-order cap) until Bloom/grid','=Inputs!{genset_fed_it}*Inputs!{partial_curtail_pct}',2,'T1-T4 partly (~half of genset-fed load); SLA credits likely',0.20,'Curtail % is Estimated (Inputs!B20); 2-month duration assumed'),
('2b. Partial: curtailed to a minor-source (25 tpy NOx) "permitted" level until Bloom/grid','=MAX(0,Inputs!{genset_fed_it}-Inputs!{it_at_minor_cap})',2,'T1-T4 almost fully (only ~1 engine-equivalent allowed)',0.05,'Shows that a literal "permitted level" ≈ full shutdown'),
('3a. Full shutdown of genset-fed load, 1 month','=Inputs!{genset_fed_it}',1,'T1-T4 (all delivered tranches except ~4 MW IT on grid); possible SLA/termination rights',0.10,'Bloom/grid restores within 1 month'),
('3b. Full shutdown, 2 months','=Inputs!{genset_fed_it}',2,'T1-T4; T5-T6 start likely slips',0.06,''),
('3c. Full shutdown, 3 months','=Inputs!{genset_fed_it}',3,'T1-T4; T5-T9 schedule at risk',0.04,''),
]
for i,(n,mw,m,tr,p,note) in enumerate(sc,start=2):
    s.cell(i,1,n); s.cell(i,2,F(mw)); s.cell(i,3,m)
    s.cell(i,4,F('=Inputs!{outage_start}')); s.cell(i,5,f'=IF(C{i}=0,D{i},EDATE(D{i},C{i})-1)')
    s.cell(i,6,f'=IF(C{i}=0,0,E{i}-D{i}+1)')
    s.cell(i,7,F(f'=IF(C{i}=0,0,MIN(E{i},Inputs!{{fy26_end}})-D{i}+1)'))
    s.cell(i,8,f'=F{i}-G{i}')
    s.cell(i,9,F(f'=B{i}*Inputs!{{rev_per_mw_yr}}*G{i}/365'))
    s.cell(i,10,f'=I{i}')
    s.cell(i,11,F(f'=B{i}*Inputs!{{rev_per_mw_yr}}*H{i}/365'))
    s.cell(i,12,f'=J{i}+K{i}')
    s.cell(i,13,F(f'=B{i}*Inputs!{{rev_per_mw_it_yr}}*F{i}/365'))
    s.cell(i,14,tr); s.cell(i,15,p); s.cell(i,16,'Derived (inputs partly Estimated)'); s.cell(i,17,note)
    for c in (4,5): s.cell(i,c).number_format='yyyy-mm-dd'
    for c in (2,9,10,11,12,13): s.cell(i,c).number_format='0.0'
    s.cell(i,15).number_format='0%'
r=len(sc)+2
s.cell(r,1,'Probability-weighted (Estimated weights)').font=B
for col in 'IJKL':
    s[f'{col}{r}']=f'=SUMPRODUCT({col}2:{col}{r-1},$O$2:$O${r-1})'; s[f'{col}{r}'].number_format='0.0'
s[f'O{r}']=f'=SUM(O2:O{r-1})'; s[f'O{r}'].number_format='0%'
s.cell(r+2,1,'Formulas: MW offline x $11.6M/MW/yr x days/365. Q4 days = MIN(end, 2026-12-31) - start + 1; FY27 days = total - Q4 days. Outage end = EDATE(start, months) - 1.').font=Font(italic=True)
s.cell(r+3,1,'Revenue at risk assumes Microsoft does not pay for undelivered capacity (take-or-pay applies to delivered capacity; SLA credits / termination rights per 20-F not quantified). Not modelled: penalty ($1.07M, DataOne-level), economic-benefit penalty, Bloom-delay knock-on to T5-T9.').font=Font(italic=True)
s.cell(r+4,1,'Context: team base Q4\'26 Vineland energized avg 81.8 MW IT -> Q4\'26 Vineland revenue ≈ 81.8 x 11.6/4 = $237M (Derived).').font=Font(italic=True)
for col,w in zip('ABCDEFGHIJKLMNOPQ',[70,12,9,11,11,9,9,9,11,11,11,10,13,55,9,22,55]): s.column_dimensions[col].width=w
for c in s[1]: c.font=B; c.alignment=Alignment(wrap_text=True)
# Sources
src=wb.create_sheet('Sources')
src.append(['Item','Finding','Tag','URL','Date','Confidence'])
for row in [
('NJDEP order','AONOCAPA dated Sep 22 2026; 62 unpermitted CAT G3516H; $1,072,000; 45 days from receipt to submit and obtain permits/certificates or cease','Reported','https://dep.nj.gov/wp-content/uploads/newsrel/2026/dataone75894-aonocapa.pdf (Incapsula-blocked; text via search snippets) ; https://www.nj.gov/governor/news/2026/20260922a.shtml','2026-09-22','High'),
('Receipt date','Not disclosed','—','https://morristownminute.org/1-07-million-penalty-vineland-data-center/','2026-09-22','—'),
('45-day deadline','Nov 6 2026 (receipt Sep 22) to ~Nov 9 (receipt Sep 25); before Nebius Q3 report Nov 10','Derived','https://www.enr.com/articles/63694-nj-fines-data-center-1m-over-123-mw-of-unpermitted-generation','2026-09-23','Medium'),
('DataOne response','"we will apply for the air permits for the temporary generators"','Reported','https://whyy.org/articles/vineland-data-center-fined/','2026-09-23','High'),
('DEP lets gensets run in window; genset ramp 16 (Oct-25) -> 32 (Feb-26) -> 62 (Jun-26); >46,000 run-hours; ~140 t NOx','','Reported','https://floodlightnews.org/vineland-nj-data-center-air-pollution/','2026-10-01','Medium-High'),
('Air-permit public notice/comment','None found for gensets or Bloom; earlier Bergen-engine application withdrawn May 27 2026; NJDEP BOSS public-notice page bot-blocked','Reported (absence)','https://dep.nj.gov/boss/public-notices/','2026-10-04','Medium'),
('Bloom energized','No evidence; stop-construction order Aug 10 2026; no public record of lift; S2 10 m / Landsat 100 m cannot resolve units','—','https://hntrbrk.com/breaking-news/vineland','2026-08-13','—'),
('Planning Board','Next meeting Oct 14 2026 6:30pm; agenda not posted as of Oct 4; latest minutes posted Aug 12 (no DataOne item)','Reported','https://www.vinelandcity.org/event/planning-board-meeting-60/','2026-10-04','High'),
('City Council','Oct 13 2026 agenda: no DataOne/data-center/Bloom items; Oct 27 no files yet','Reported','https://vinelandnj.api.civicclerk.com/v1/Events','2026-10-04','High'),
('Tranche dates','T1 Nov-25, T2 Feb-26 (20-F); two tranches in month before Aug 12 2026 (Q2 letter) -> Jul-Aug, not June','Reported','https://assets.nebius.com/assets/a6ecfd85-a6cb-4967-8ef7-9a25bd261f9c/SHLQ226.pdf','2026-08-12','Medium-High'),
]: src.append(list(row))
for col,w in zip('ABCDEF',[30,90,16,90,12,12]): src.column_dimensions[col].width=w
for c in src[1]: c.font=B
wb.save('/workspace/nebius/capacity/vineland_shutdown.xlsx')
