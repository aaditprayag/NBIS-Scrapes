# SOTP tab. Executed inside build.py.
S.update({
'TOLOKA6K':('6-K','2025-05-07','0001104659-25-045373','tm2514304d1_ex99-1.htm','Toloka strategic investment led by Bezos Expeditions PR (Ex.99.1)'),
'AVRIDE6K':('6-K','2025-10-22','0001104659-25-101520','tm2529289d1_ex99-1.htm','Avride up to $375M Uber/Nebius investment PR (Ex.99.1)'),
})
WEB={'CH_D':'Reuters 2026-01-16 "ClickHouse valued at $15 billion" https://www.reuters.com/technology/database-management-firm-clickhouse-valued-15-billion-amid-ai-boom-2026-01-16/ ; ClickHouse blog 2026-01-16 https://clickhouse.com/blog/clickhouse-raises-400-million-series-d-acquires-langfuse-launches-postgres',
'CH_C':'TechCrunch 2026-01-16 ("up from its $6.35 billion valuation last May") https://techcrunch.com/2026/01/16/snowflake-databricks-challenger-clickhouse-hits-15b-valuation/',
'CH_EXT':'ClickHouse blog 2026-09-22 (Series D "extended to include" JPM Private Capital, BDT & MSD, Craft, 20VC; Scarpelli to board) https://clickhouse.com/blog/clickhouse-appoints-mike-scarpelli-to-board-of-directors',
'CH_ARR':'TechCrunch 2026-05-27 (ARR >$250M; IPO "within the next few years") https://techcrunch.com/2026/05/27/clickhouse-triples-annualized-revenue-to-250m-charting-a-path-toward-an-ipo/',
'TOL':'Reuters 2025-05-07 ("Bezos Expeditions is leading a $72-million investment") https://www.reuters.com/business/amazons-bezos-leads-new-investment-ai-data-company-toloka-2025-05-07/',
'AV_CAP':'Caplight (third-party estimate, not a transaction price) https://www.caplight.com/company/avride',
'AV_HY':'Reuters 2025-03-05 (Hyundai partnership; no investment disclosed) https://www.reuters.com/business/autos-transportation/self-driving-technology-startup-avride-partners-with-hyundai-expand-robotaxi-2025-03-05/'}
sp=wb.create_sheet('SOTP')
sp['A1']='Sum-of-the-parts inputs: non-core businesses & equity stakes ($M). Blue = input (sourced). Filing figures Reported; public-press marks labelled. Estimated values show their method.'
sp['A1'].font=TITLE
hdr(sp,3,['Asset','Nebius ownership','Ownership source / note','Accounting (6/30/26)','Carrying value 31-Dec-25','Carrying value 30-Jun-26','Latest transaction / mark','Mark date','Implied 100% equity value','Implied Nebius share (Estimated)','Tag','Filing sources','Public / third-party sources'])
r=4
# ClickHouse inputs
sp.cell(r,1,'ClickHouse Inc. (minority stake; spun off Sep-2021)')
sp.cell(r,2,0.28).font=BLUE; sp.cell(r,2).number_format='0%'
sp.cell(r,3,'"Approximately 28% stake on a fully diluted basis, upon exercise of penny warrants" (Oct-24 investor presentation); "We hold a 28% ownership stake" (FY24 20-F p.37). FY25 20-F and 2026 letters only say "significant minority stake" (% no longer disclosed). Plus pre-funded Series C warrants bought for $50.0M (May-2025). Dilution from Series C/D/extension not disclosed -> 28% is a STALE upper bound.')
sp.cell(r,4,'ASC 321 measurement alternative (no significant influence since 2022); remeasured on observable rounds')
sp.cell(r,5,737.1).font=BLUE; sp.cell(r,6,1517.7).font=BLUE
sp.cell(r,7,'Series D: $400M at ~$15B (Jan-2026; extended Sep-2026, same round). Prior: Series C ~$6.35B (May-2025); Series B $2B (Oct-2021). NBIS fair-value inputs: $159.59/share Series D price, 57.5% vol, 2-yr liquidity, 22.5% DLOM (Q2FS); Series C $72.50/share (FY25 20-F).')
sp.cell(r,8,'2026-01-16')
sp.cell(r,9,15000).font=BLUE
sp[f'J{r}']=f'=B{r}*I{r}'
sp.cell(r,11,'Carrying: Reported. Valuation: Reported in 20-F ("valuation of approximately $15 billion"). Nebius share: Estimated = stale 28% x $15B headline (upper bound, before dilution, liquidation preferences and DLOM)')
sp.cell(r,12,src('20F25','p.52 (Item 5) & subsequent events note')+' ; '+src('Q2FS','Note 5 p.F-20/21; Note 6 p.F-23')+' ; '+src('20F24','p.37')+' ; '+src('PRES24','p.4 footnote'))
sp.cell(r,13,WEB['CH_D']+' ; '+WEB['CH_C']+' ; '+WEB['CH_EXT']+' ; '+WEB['CH_ARR'])
ch=r; r+=1
# Toloka
sp.cell(r,1,'Toloka Group, Inc. (deconsolidated 2-May-2025)')
sp.cell(r,2,0.81).font=BLUE; sp.cell(r,2).number_format='0%'
sp.cell(r,3,'"The Group holds 81% of the economic interest and 49% of the voting interest in Toloka on an outstanding basis" (FY25 20-F Note 5)')
sp.cell(r,4,'Common: equity method (carrying reduced to nil by losses at 6/30/26); preferred: ASC 321 measurement alternative')
sp.cell(r,5,'101.7 (common 4.7 + preferred 97.0)'); sp.cell(r,6,86.5).font=BLUE
sp.cell(r,7,'Bezos Expeditions-led round (with Mikhail Parakhin), May-2025: $72M per Reuters (amount not given in Nebius 6-K). Retained stake remeasured to $130.6M fair value at deconsolidation (common 33.6 + preferred 97.0; OPM, 80% vol, 24% WACC).')
sp.cell(r,8,'2025-05-07')
sp[f'I{r}']=f'=130.6/B{r}'
sp[f'J{r}']='=130.6'
sp.cell(r,11,'Reported (FV 130.6, carrying 86.5); implied 100% value Derived = 130.6 / 81% (rough; OPM-allocated, not a headline post-money)')
sp.cell(r,12,src('20F25','Note 3; Note 5 p.F-35/37; Note 6 p.F-37')+' ; '+src('Q2FS','Note 5 p.F-20/21')+' ; '+src('TOLOKA6K'))
sp.cell(r,13,WEB['TOL'])
tol=r; r+=1
# Avride
sp.cell(r,1,'Avride (autonomous vehicles & delivery robots; Avride Holding Inc.)')
sp.cell(r,2,'consolidated subsidiary; ~100% of common before SAFE conversion & Avride ESOP (exact % not disclosed)')
sp.cell(r,3,'Avride ESOP authorizes up to 7,926,674 Avride Holding shares (total share count not disclosed). SAFEs to SMB Holding Corporation: $100.0M (Oct-2025) + $75.0M Milestone SAFE (Jun-2026); convert at 80% of next qualifying round price. NBIS filings do not state SMB\'s affiliation.')
sp.cell(r,4,'Consolidated segment; SAFE liabilities at fair value (Level 3)')
sp.cell(r,5,'not separately disclosed (consolidated)'); sp.cell(r,6,'not separately disclosed (consolidated)')
sp.cell(r,7,'Oct-2025: "strategic investments and other commitments of up to $375 million" from Uber and Nebius (split not disclosed). SAFE liabilities FV 6/30/26: $101.8M + $76.3M = $178.1M (claims senior to Nebius common in Avride). Hyundai (Mar-2025): development/fleet partnership, no investment disclosed. No priced round / valuation disclosed by Nebius.')
sp.cell(r,8,'2025-10-22')
sp.cell(r,9,'not disclosed (Caplight third-party estimate ~$2.75B - not a transaction price, low confidence)')
sp.cell(r,10,'not disclosed')
sp.cell(r,11,'Reported (transaction terms); valuation not disclosed')
sp.cell(r,12,src('AVRIDE6K')+' ; '+src('20F25','p.13 risk factor (Uber up to $375M), p.41 (Hyundai), other accrued liabilities note')+' ; '+src('Q2FS','Note 6 p.F-22; Note 15 segments'))
sp.cell(r,13,WEB['AV_CAP']+' ; '+WEB['AV_HY'])
av=r; r+=1
# TripleTen
sp.cell(r,1,'TripleTen (edtech)')
sp.cell(r,2,'100% (consolidated subsidiary)'); sp.cell(r,3,'Wholly owned; no external investors disclosed')
sp.cell(r,4,'Consolidated segment'); sp.cell(r,5,'not separately disclosed (consolidated)'); sp.cell(r,6,'not separately disclosed (consolidated)')
sp.cell(r,7,'No transaction or valuation ever disclosed. Third-party blog estimates ($320M-$800M) exist but are not company data and are not used.')
sp.cell(r,8,'n/a'); sp.cell(r,9,'not disclosed'); sp.cell(r,10,'not disclosed')
sp.cell(r,11,'Not disclosed')
sp.cell(r,12,src('20F25','segment note; Item 15 material weakness re TripleTen revenue recognition')+' ; '+src('Q2FS','Note 15'))
sp.cell(r,13,'n/a'); tt=r; r+=1
# VC fund & other
sp.cell(r,1,'Venture capital fund (equity method)'); sp.cell(r,2,'not disclosed'); sp.cell(r,3,'Fund not named'); sp.cell(r,4,'Equity method')
sp.cell(r,5,6.4).font=BLUE; sp.cell(r,6,13.8).font=BLUE; sp.cell(r,7,'+7.4 change in fair value Q2-26'); sp.cell(r,8,'2026-06-30'); sp.cell(r,9,'not disclosed'); sp[f'J{r}']=f'=F{r}'
sp.cell(r,11,'Reported (carrying used as value)'); sp.cell(r,12,src('Q2FS','Note 5 p.F-20')); sp.cell(r,13,'n/a'); vc=r; r+=1
sp.cell(r,1,'Other non-marketable equity securities'); sp.cell(r,2,'not disclosed'); sp.cell(r,3,'Not named'); sp.cell(r,4,'ASC 321')
sp.cell(r,5,2.5).font=BLUE; sp.cell(r,6,2.5).font=BLUE; sp.cell(r,7,'n/a'); sp.cell(r,8,'n/a'); sp.cell(r,9,'not disclosed'); sp[f'J{r}']=f'=F{r}'
sp.cell(r,11,'Reported'); sp.cell(r,12,src('Q2FS','Note 5 p.F-21')); sp.cell(r,13,'n/a'); oth=r; r+=1
sp.cell(r,1,'Total (stakes with a value; Avride & TripleTen excluded)').font=TOT
sp[f'F{r}']=f'=F{ch}+F{tol}+F{vc}+F{oth}'; sp[f'J{r}']=f'=J{ch}+J{tol}+J{vc}+J{oth}'
sp.cell(r,11,'Derived'); sp.cell(r,12,'Q2FS Note 5 totals: non-marketable 1,606.7 + equity method 13.8 = 1,620.5'); tot=r; r+=2
# per-share helper
sp.cell(r,1,'Per NBIS share (uses CapStructure fully diluted, base case)')
sp[f'F{r}']=f"=F{tot}*1000000/CapStructure!B{CAPROWS['fd']}"; sp[f'J{r}']=f"=J{tot}*1000000/CapStructure!B{CAPROWS['fd']}"
sp.cell(r,11,'Derived'); sp[f'F{r}'].number_format='$0.00'; sp[f'J{r}'].number_format='$0.00'; r+=2
# Segment operating data for valuing consolidated non-core
hdr(sp,r,['Consolidated non-core segment data ($M)','FY2024','FY2025','Q1 2026','Q2 2026','1H 2026','','','','','Tag','Source','']); r+=1
SEG=[('Avride revenue',0.3,1.3,0.9,1.0,'=D{r}+E{r}'),('Avride adj. EBITDA',-67.0,-82.7,-34.1,-40.1,'=D{r}+E{r}'),('TripleTen revenue',28.8,54.1,11.6,10.0,'=D{r}+E{r}'),('TripleTen adj. EBITDA',-30.8,-41.2,-10.4,-9.4,'=D{r}+E{r}')]
for lab,a,b,c,d,f in SEG:
    sp.cell(r,1,lab)
    for j,v in zip((2,3,4,5),(a,b,c,d)): sp.cell(r,j,v).font=BLUE
    sp[f'F{r}']=f.format(r=r); sp.cell(r,11,'Reported (1H Derived)'); sp.cell(r,12,src('20F25','segment note')+' ; '+src('Q1FS','segment note')+' ; '+src('Q2FS','Note 15')); r+=1
sp.cell(r,1,'Operating KPIs: Avride >60,000 commercial rides on Uber in Dallas; robodelivery in Arlington VA & Miami via Uber Eats (Q2-26 letter). TripleTen revenue +10% YoY in Q1-26 (Q1-26 letter); Q2-26 revenue 10.0 vs 12.3 in Q2-25 (-19%, Derived).'); sp.cell(r,12,src('Q2LTR','Other businesses section')+' ; '+src('Q1LTR','Other businesses section')); r+=2
sp.cell(r,1,'Notes: (1) Nebius filings reported gains on ClickHouse revaluation of $597.4M (FY25, Series C) and $780.6M (Q1-26, Series D) - non-cash, flowed through net income. (2) ClickHouse carrying value ($1,517.7M) already reflects the OPM allocation and a 22.5% DLOM, so it sits well below stake x headline valuation. (3) Toloka preferred carrying fell 97.0 -> 86.5 as equity-method losses were applied after common was written to nil. (4) The Avride SAFEs ($178.1M FV) would dilute Nebius at conversion (80% of the next round price).')
for col,wd in zip('ABCDEFGHIJKLM',[40,16,60,30,18,16,60,12,22,20,40,70,70]): sp.column_dimensions[col].width=wd
for row in sp.iter_rows(min_row=3):
    for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
for rr in (ch,tol,vc,oth,tot):
    for col in 'EFIJ': sp[f'{col}{rr}'].number_format='#,##0.0'
