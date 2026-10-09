# Guidance_Log tab (raw, unscored). Executed inside build.py.
import sys as _sys, re as _re, csv as _csv
_sys.path.insert(0,'/workspace/nebius/filings/altdata/scripts')
import importlib as _il
_gd=_il.import_module('guidance_data'); _so=_il.import_module('sources')
AS=_so.S
def _local(k):
    if k.startswith('FOOL_Q'): return '/workspace/nebius/filings/altdata/transcripts/transcript_q'+k[6]+'-20'+k[8:10]+'_fool.txt'
    m=_re.match(r'CALL_Q(\d)_(\d\d)',k)
    if m: return f'/workspace/nebius/filings/altdata/transcripts/transcript_q{m.group(1)}-20{m.group(2)}.txt'
    m=_re.match(r'Q(\d)_(\d\d)_SHL',k)
    if m: return f'/workspace/nebius/filings/altdata/mined_text/Q{m.group(1)}-{m.group(2)}_letter.txt'
    url=AS[k][0]; return '/workspace/nebius/filings/edgar/'+url.rsplit('/',1)[1]+'.txt'
def _norm(s): return _re.sub(r'[\s\u00ad]+',' ',s).replace('’',"'").replace('–','-').replace('—','-').strip().lower()
def _check(k,q):
    try: t=_norm(open(_local(k),encoding='utf-8',errors='ignore').read())
    except Exception: return 'source text not local'
    parts=[p for p in q.split(' ... ') if p.strip()]
    return 'verified' if all(_norm(p) in t for p in parts) else 'NOT FOUND - check'
def _chan(k):
    return 'Earnings call' if k.startswith(('CALL','FOOL')) else ('Shareholder letter' if 'SHL' in k else ('Investor presentation' if 'PRES' in k else ('20-F' if '20F' in k else 'Press release / 6-K')))
EXTRA=[ # (date, src, category, metric, figure, change vs prior, quote, speaker/note)
('2024-10-31','CALL_Q3_24','ARR','ARR Dec-2025 (original range)','above midpoint of $500M-1B','Reaffirmed with upside framing',"In that scenario where we're able to reallocate, we estimate that we'll be able to deliver above the midpoint of the $500 million-$1 billion ARR guidance range that we gave for a year in 2025.",'call'),
('2025-02-20','CALL_Q4_24','GPU','GPU deployments 2025','"tens of thousands"','Call statement; superseded by FY24 20-F actual (~30,000 GPUs at 31-Mar-2025)','We deployed the first several tens of thousands of GPUs and waiting for the next generation with more tens of thousands of GPUs deployed, which will be deployed this year.','call'),
('2025-04-30','FY24_20F','GPU','GPUs deployed at 31-Mar-2025','~30,000 (primarily H200)','Last company GPU count (actual); none thereafter','As of March 31, 2025, we have deployed approximately 30,000 GPUs across our global data center footprint, primarily NVIDIA H200s.','20-F'),
('2025-05-20','CALL_Q1_25','Revenue','Group revenue FY2025','$500-700M','Reiterated',"We're also reiterating our overall revenue guidance for the group, which is in the range of $500-$700 million.",'call'),
('2025-05-20','Q1_25_SHL','Margin','Medium-term adj. EBIT margin','20-30%','NEW long-term target; explicitly tied to 4-year depreciation','In the medium term, our base case expectations are to achieve billions of dollars in revenue with adjusted EBIT margins in the 20-30% range, assuming a conservative depreciation schedule of four years.','letter'),
('2025-05-20','CALL_Q1_25','Margin','Medium-term EBIT margin','20-30%','Same target on call','we anticipate medium-term EBIT margins to range in the sort of 20%-30% range.','call'),
('2025-05-20','CALL_Q1_25','MW','Capacity deployed 2025','>100 MW','Wording "deployed" (vs "in operations" in Feb-25)',"I expect that operationally, we'll have over 100 MW of capacity deployed this year.",'call'),
('2025-08-07','CALL_Q2_25','Revenue','Core business revenue FY2025','$400-600M','Maintained; perimeter now ex-Toloka','For our core business revenue, we are maintaining our guidance of $400 million- $600 million.','call'),
('2025-08-07','CALL_Q2_25','Revenue','Toloka revenue FY2025 guidance','$50-70M','Removed from group guidance after deconsolidation','This excludes the 2025 revenue guidance of $50 million- $70 million we previously gave for Toloka.','call'),
('2025-08-07','CALL_Q2_25','Capex','Capex FY2025','~$2B','Maintained','Finally, we are maintaining our CapEx guidance of around $2 billion in 2025.','call'),
('2025-08-07','CALL_Q2_25','EBITDA','Group adj. EBITDA positive','by end of 2025','Wording moved from "second half of 2025" (May-25) to "by the end of the year"','At the group level, we anticipate turning adjusted EBITDA positive by the end of the year.','call'),
('2025-08-07','Q2_25_SHL','MW','Power 2026','1 GW','First 2026 power target (type of MW unspecified)','We are aggressively scaling capacity and targeting 1 GW of power in 2026.','letter'),
('2025-08-07','CALL_Q2_25','MW','Finland additional MW 2025','+50 MW','New site-level target','In Finland, we expect to have an additional 50 MW in operations this year, just like we spoke earlier.','call'),
('2025-11-11','Q3_25_SHL','ARR','ARR Dec-2026','$7-9B','NEW; framed as conditional on financing ("Given the scale of the financing that we expect to put in place")','we believe that we can achieve annualized run-rate revenue (ARR)1 of $7 billion to $9 billion by the end of 2026.','letter'),
('2025-11-11','CALL_Q3_25','ARR','ARR Dec-2026 booked share','>50% booked','Backlog coverage claim','gives us the confidence that we can achieve the $7 billion-$9 billion of ARR, of which more than half is already booked.','call'),
('2025-11-11','CALL_Q3_25','Financing','Funding mix','corporate debt + asset-backed + equity','NEW financing framework (asset-backed first delivered Jul-26, $775M)','we will be utilizing at least three sources: corporate debt, asset-backed financing, and equity.','call'),
('2026-02-12','Q4_25_SHL','ARR','ARR Dec-2026','$7-9B','Reaffirmed; conditional financing language dropped','We are on track to end the year with ARR1 of $7 billion to $9 billion.','letter'),
('2026-02-12','CALL_Q4_25','EBITDA','Group adj. EBITDA margin FY2026','~40%','NEW','On adjusted EBITDA, we expect group adjusted EBITDA margin to be approximately 40% for 2026.','call'),
('2026-02-12','CALL_Q4_25','Margin','Medium-term EBIT margin','20-30%, "potential to go higher"','Reaffirmed WITHOUT the 4-year-depreciation caveat, same call that announced 4->5 yr life change','We believe the return on these investments are attractive and will remain committed to our medium-term EBIT target of 20%-30%, with the potential to go higher.','call'),
('2026-02-12','CALL_Q4_25','Pipeline','Q1-26 pipeline creation','>$4B','NEW metric (one-off)','The pipeline creation trajectory in Q1 is on track to exceed $4 billion','call'),
('2026-05-13','CALL_Q1_26','All','FY2026 ARR / revenue / margin','$7-9B / $3.0-3.4B / ~40%','Reiterated (capex raised separately)','As such, we are reiterating our full year 2026 guidance for annualized run rate revenue of $7 billion-$9 billion, group revenue of between $3 billion and $3.4 billion, and group-adjusted EBITDA margin of around 40%.','call'),
('2026-05-13','CALL_Q1_26','EBITDA','Quarterly margin path 2026','non-linear','NEW qualifier','We anticipate a nonlinear quarterly adjusted EBITDA margin progression during 2026.','call'),
('2026-05-13','CALL_Q1_26','Utilization','Utilization','"sell out" (qualitative)','Qualitative only - no % ever given','On utilization, we continue to sell out our capacity, and we expect this to be the case for the foreseeable future due to strong market demand and our healthy pipeline.','call'),
('2026-05-13','Q1_26_SHL','MW','Connected power YE2026','800MW-1GW','Reaffirmed','We continue to expect 800MW to 1GW of connected power by year-end.','letter'),
('2026-05-13','Q1_26_SHL','MW','Owned capacity','3 GW across five sites','NEW','In total, our owned facilities will deliver 3 GW of capacity across five sites','letter'),
('2026-08-12','CALL_Q2_26','All','FY2026 ARR / revenue / margin / capex','$7-9B / $3.0-3.4B / ~40% / $20-25B','Reiterated (unchanged despite contracted power raise to 5 GW)','We continue to expect annualized run- rate revenue of $7 billion-$9 billion, group revenue of between $3 billion and $3.4 billion, group adjusted EBITDA margin of approximately 40%, and capital expenditures of $20 billion-$25 billion.','call'),
('2026-08-12','FOOL_Q2_26','MW','Connected power YE2026 (800MW-1GW) vs active','800MW-1GW connected by YE26 (reaffirmed); active "throughout the first half of 2027"; connected-to-revenue lag "a few months"','restated orally, dropped from letter. Not in Q2-26 shareholder letter or PR (no "800"/"connected" in tm2622968d1_ex99-1/-2); restated by Korolenko (Chief Infrastructure & Product Officer) only in answer to Cantor Q. Note: "a few months" is the lag from connected power (data centers) to revenue, per transcript wording','Yes, we still -- well, we still expect to meet this guidance from 800 megawatts to 1 gigawatt of connected power this year. ... You have to commission the data center, build the network, build the clusters, deploy the platform, then onboarding the customers, and then the revenue generation starts. So that takes a few months. ... And in terms of our guidance from 800 megawatts to 1 gigawatt, yes, that\'s more -- that\'s true about the connected, but I would think that about being active in -- throughout the first half of 2027.','call (Andrey Korolenko; Motley Fool transcript, 3rd party)'),
('2026-08-12','CALL_Q2_26','Pricing','New deal economics','>$20M revenue/MW; <2-yr payback','NEW unit-economics disclosure (replaces any GPU-hour pricing)','You should expect the deals we closed in this quarter with more than $20 million per megawatt deals and less than two-year payback period to start coming online from late Q4 and onwards.','call'),
('2026-08-12','Q2_26_SHL','Pricing','ACV per MW','$20-25M','NEW','annual contract value of $20-25 million per MW','letter'),
('2026-08-12','CALL_Q2_26','Prepayments','Customer prepayments 2026','>$9B','Reaffirmed on call',"In total, customers' prepayments will bring in more than $9 billion of up-front funding this year, directly reducing the capital we need from debt and equity.",'call'),
]
gl=wb.create_sheet('Guidance_Log')
gl['A1']='Guidance & wording log (raw, chronological, UNSCORED). Hit-rate scoring is in /workspace/nebius/filings/altdata/ (Guidance_HitRate); this log shares the same quote base (altdata/scripts/guidance_data.py, verified verbatim) plus additional reaffirmations/wording changes added here.'
gl['A1'].font=TITLE
hdr(gl,2,['Date said','Category','Metric','Figure','Change vs prior / wording note','Exact quote','Channel','Speaker/notes','Source key','Source URL','Quote check (local text)','Origin'])
_glr=[]
for d in _gd.G:
    _glr.append((d['said'],d['cat'],d['metric'],d['target'],(d.get('note') or '')[:300],d['quote'],_chan(d['src']),'',d['src'],AS[d['src']][0],'verified by altdata/verify_quotes.py','altdata G '+d['id']))
for w in _gd.W:
    _glr.append((w[0],'Wording/definition','Metric wording',w[2],'Definition / wording change',w[3],_chan(w[1]),'',w[1],AS[w[1]][0],'verified by altdata/verify_quotes.py','altdata W'))
for e in EXTRA:
    _glr.append((e[0],e[2],e[3],e[4],e[5],e[6],_chan(e[1]),e[7],e[1],AS[e[1]][0],_check(e[1],e[6]),'Filings worker (new)'))
_glr.sort(key=lambda r:(r[0],r[1]))
for r in _glr: gl.append(list(r))
GL_N=len(_glr); GL_BAD=[r for r in _glr if r[10].startswith('NOT')]
r0=gl.max_row+2
gl.cell(r0,1,'What they stopped disclosing or never disclosed').font=Font(bold=True,size=12)
hdr(gl,r0+1,['Item','Status','Last disclosed figure (date)','Evidence / exact wording','','','','','Source key','Source URL','',''])
GAPS=[
('Utilization %','NEVER disclosed (qualitative only)','none - "peak utilization", "sold out" (Q4-25 letter; Q1-26 call)','Q3-24 PR footnote: revenue potential "Based on assumed utilization and pricing as per the company\'s medium-term financial plan"; Q4-25 letter: "We operated at peak utilization during the quarter and were sold out of capacity in Q4"','Q4_25_SHL'),
('GPU pricing ($/GPU-hour)','NEVER disclosed; replaced by $/MW revenue metrics in Q2-26','n/a; Q2-26: $12M/MW 2026 base, >$20M/MW Q2 deals, >$40M short-term deals, ACV $20-25M/MW','Q2-26 letter: "annual contract value of $20-25 million per MW"','Q2_26_SHL'),
('GPU counts (installed / deployed)','STOPPED after early 2025','">20,000 GPUs" by YE24 target (Oct-24 call); ">22,000 NVIDIA Blackwell GPUs" plan (Dec-24/Feb-25 PR); "tens of thousands" (Feb-25 call); ~30,000 deployed at 31-Mar-2025 (FY24 20-F, last actual count)','No installed-GPU count after the FY24 20-F (as of 31-Mar-2025: ~30,000 GPUs, primarily H200) in any 2025-26 letter, 20-F or interim FS reviewed; Oct-24 presentation measured capacity "In Nvidia H100 equivalent"','Q4_24_PR'),
('Active MW','STOPPED after YE25','~170 MW active at 31-Dec-2025 (Q4-25 letter)','Q1-26 letter re-defines active power ("power being consumed by installed, operational IT equipment and available for revenue generation") but gives no figure; Q2-26 letter none; Q2-26 call pushes "active" of the 800MW-1GW into 1H27','Q1_26_SHL'),
('Connected MW actual','Target given; YE25 actual not stated','45 MW connected at 31-Mar-2025 (FY24 20-F); 220 MW YE25 target (Aug-25)','Q4-25 letter reports ~170 MW active vs 100 MW target but no YE25 connected figure vs the 220 MW target','FY24_20F'),
('Customer names / mix','Only contract-level 6-Ks (Microsoft, Meta); mix anonymized','Customer A-E % of revenue (Q2-26: C 24%, D 21%, E 14%); ">40 managed clients" (Q3-24 PR) - last customer count','No customer count or industry mix since 2024; >10% customers unnamed in FS notes','Q2_26_FS'),
('Useful life rationale','Thin','4 yrs through FY25 ("conservative depreciation schedule of four years", Q1-25 letter); 5 yrs from 1-Jan-2026','Only rationale: "based on usage patterns and current utilization commitments" (Q2-26 letter p.10). No retirement / re-lease / residual-value data. Medium-term 20-30% EBIT target was first framed on 4-yr depreciation, then reaffirmed in Feb-26 without that caveat','Q2_26_SHL'),
('Capex guidance in filings','Call-only','FY26 $20-25B (Q1-26 call, reiterated Q2-26 call)','Capex guidance not in the 2026 shareholder letters or 6-K PR text reviewed; only on calls','CALL_Q2_26'),
('Segment guidance (Avride, TripleTen)','NEVER guided','n/a','No revenue/EBITDA guidance for Avride or TripleTen in any letter/call reviewed','Q2_26_SHL'),
('RPO duration / ARR by contract type','Partial','RPO $37.5B (Q2-26); ARR definition = AI cloud month revenue x12','Split of ARR between long-term contracts (Microsoft/Meta) and on-demand not disclosed','Q2_26_FS'),
('Microsoft/Meta MW or GPU per contract','NEVER disclosed','TCV only (MSFT up to $19.4B; Meta $2.9B + $12B + $15B backstop)','See MW_Disclosures sheet','FY25_20F'),
]
for g in GAPS:
    gl.append([g[0],g[1],g[2],g[3],'','','','',g[4],AS[g[4]][0],'',''])
for col,wd in zip('ABCDEFGHIJKL',[12,14,30,22,40,80,16,14,14,50,18,18]): gl.column_dimensions[col].width=wd
for row in gl.iter_rows(min_row=3):
    for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
# share raw log with altdata worker
with open('/workspace/nebius/filings/altdata/raw/guidance_log_raw_from_filings.csv','w',newline='') as f:
    w=_csv.writer(f); w.writerow(['date','category','metric','figure','change_note','quote','channel','speaker','src_key','url','quote_check','origin'])
    for r in _glr: w.writerow(r)
print('Guidance rows',GL_N,'bad quotes',len(GL_BAD),[b[2] for b in GL_BAD])
