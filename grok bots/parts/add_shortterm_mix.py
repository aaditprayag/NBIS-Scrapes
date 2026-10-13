#!/usr/bin/env python3
"""Append/replace the 'ShortTerm_Mix' sheet in contract_economics.xlsx (other sheets untouched).
Called at the end of build_contract_economics.py so a rebuild keeps the sheet. Built 2026-10-04."""
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
P='/workspace/nebius/parts/_rebuild/contract_economics.xlsx'  # 10/8 lock guard
NB_CALL='https://stockanalysis.com/stocks/nbis/transcripts/655490-q2-2026/'
NB_LET='https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-2.htm'
NB_FS='https://www.sec.gov/Archives/edgar/data/1513845/000110465926094844/nbis-20260812xex99d2.htm'
NB_Q1C='https://stockanalysis.com/stocks/nbis/transcripts/555850-q1-2026/'
NB_Q1L='https://www.sec.gov/Archives/edgar/data/1513845/000110465926059872/tm2614392d1_ex99-2.htm'
NB_Q4C='https://stockanalysis.com/stocks/nbis/transcripts/402284-q4-2025/'
NB_Q424='https://stockanalysis.com/stocks/nbis/transcripts/246397-q4-2024/'
NB_Q324='https://stockanalysis.com/stocks/nbis/transcripts/216350-q3-2024/'
NB_PR='https://nebius.com/prices'
CW_S1='https://www.sec.gov/Archives/edgar/data/1769628/000119312525044231/d899798ds1.htm'
CW_10K='https://www.sec.gov/Archives/edgar/data/1769628/000176962826000104/crwv-20251231.htm'
CW_Q3='https://www.sec.gov/Archives/edgar/data/1769628/000176962825000062/crwv-20250930.htm'
CW_Q1='https://www.sec.gov/Archives/edgar/data/1769628/000176962826000222/crwv-20260331.htm'
CW_Q2='https://www.sec.gov/Archives/edgar/data/1769628/000176962826000366/crwv-20260630.htm'
CW_CALL='https://s205.q4cdn.com/133937190/files/doc_financials/2026/q2/CRWV-US-CORRECTED-TRANSCRIPT-CoreWeave-Q2-2026-Earnings-Call-11August2026.pdf'
IR_10K='https://www.sec.gov/Archives/edgar/data/1878848/000187884826000052/iren-20260630.htm'
IR_PR='https://www.sec.gov/Archives/edgar/data/1878848/000187884826000051/irenreportsfy26results.htm'
IR_JUL='https://www.sec.gov/Archives/edgar/data/1878848/000114036126028871/ef20078253_ex99-1.htm'
IR_CALL='https://stockanalysis.com/stocks/iren/transcripts/662889-q4-2026/'
NS_S1='https://www.sec.gov/Archives/edgar/data/2110365/000119312526395475/ck0002110365-20260918.htm'
SA='https://newsletter.semianalysis.com/p/the-great-gpu-shortage-rental-capacity'
SA_IDX='https://gpu-index.semianalysis.com/'
SA_CMX='https://newsletter.semianalysis.com/p/the-gpu-cloud-clustermax-rating-system-how-to-rent-gpus'
H=['Company','Metric','Value','Period','Basis (revenue vs capacity)','Tag','Source URL','Page / section','Confidence']
R=[
# ---- Nebius
['Nebius','Share of 2027 capacity on deals of 6 months or less','Not disclosed. Mgmt: "deliberately allocating a portion of capacity for short-term and immediate needs"; "we could sell today our entire 2027 capacity on these [mid-term] terms... We are not doing this"','Q2-26 (Aug 12 2026)','Capacity','Reported (qualitative); % not disclosed',NB_CALL,'Prepared remarks (Volozh, three deal types); Q&A Tyler Radke/Citi answer (Boroditsky)','High (quote) / n.a. (%)'],
['Nebius','Short-term deal tenor and price','Tenor "typically for up to six months" (call) / "typically between three and six months" (letter); "3- to 6-month engagements" for dedicated GB300 clusters; negotiating $40-50M per MW "and sometimes above"; letter: Q3 short-term deals ">$40M" per MW','Q2-26 / early Q3-26','Price per MW (capacity)','Reported',NB_LET+' ; '+NB_CALL,'Letter p.2 KPI panel ("Q3 short-term capacity deals >$40M") and p.3 "A contract mix built to capture a rising market"; call Q&A Gili/Boroditsky on auction + short-term scale training','High'],
['Nebius','Number of short-term deals signed','1 ("One such deal has been signed in Q3"; capacity expected live Q4-26). Deal size/MW not disclosed','Q3-26','Count','Reported',NB_LET,'Letter p.3 "Shorter-duration contracts" bullet; p.4 "In Q3, we closed our first short-term capacity deal"','High'],
['Nebius','Status of short-term program','"is a new effort" - i.e. no historical base for a % of capacity','Q2-26','Qualitative','Reported',NB_LET,'Letter p.3 contract-mix bullets','High'],
['Nebius','Capacity auction (mid-term, not short-term)','First auction pilot for a mid-term contract cleared "15% above the highest price we ever charged before" for Blackwell. NOT a <=6-month deal - do not count in the short-term share','Q2/Q3-26','Price signal','Reported',NB_CALL,'Prepared remarks (Volozh) "auction for the midterm contract"','High'],
['Nebius','Core mid-term deal economics (the alternative use of the same MW)','Mid-term 1-3 yr contracts at $20-25M ACV per MW; Q2 deals avg >$20M/MW; 50-60% capex prepaid; 4 landmark deals avg >$1B TCV','Q2-26','Price per MW (capacity)','Reported',NB_LET,'Letter p.2 KPI panel; p.3-4 "Landmark customer agreements"','High'],
['Nebius','Price premium of short-term vs mid-term per MW','~2.0x (Derived: $45M midpoint / $22.5M midpoint = 2.0x; range $40/25 = 1.6x to $50/20 = 2.5x)','Q2-26','Price per MW (capacity)','Derived',NB_LET,'Letter p.2-3 (inputs)','Medium (per-MW bases may differ; letter does not state MW basis)'],
['Nebius','Self-serve / on-demand capacity retained','"We do retain a portion of capacity for self-service"; "vast majority of capacity coming online over the next several quarters to 12 months is already under contract or earmarked" - % not disclosed','Q1-26 (May 13 2026)','Capacity','Reported (qualitative)',NB_Q1C,'Q&A answer by Marc Boroditsky (sold-out / future-selling question)','High (quote)'],
['Nebius','On-demand plus reserve flexibility','"We proactively manage our ability to support on-demand capacity while also offering reserve contracts"','Q1-26','Qualitative','Reported',NB_Q1L,'Letter, "Go-to-market remains agile" section','High'],
['Nebius','Trend in contract duration','Avg contract duration of new cloud customers +50% in 2025; Q4-25 had ~2x as many >12-month transactions as Q3-25; Hopper renewals "off of short reserve agreements" renewing at 12 months or longer','FY25 / Q4-25','Qualitative (direction: lengthening)','Reported',NB_Q4C,'Prepared remarks (Volozh, avg duration +50%); Q&A (Boroditsky, deal trends)','High'],
['Nebius','Historical mix commentary','"some short-term on-demand contracts..."; expected Blackwell to bring "more of a mix of short-term as well as long-term"; Q3-24: "a lot of on-demand customers", 1-2 yr H200 contracts in pipeline','Q3-24, Q4-24','Qualitative','Reported',NB_Q424+' ; '+NB_Q324,'Q4-24 Q&A (Chernin, customer mix); Q3-24 Q&A (Chernin, Blackwell contracts)','High (quote)'],
['Nebius','On-demand + Token Factory share of revenue (team estimate)','~4% base (range 1.3-9%) of Q2-26 revenue of $582.3M (~$23M)','Q2-26','Revenue','Estimated (team; method: residual after inferred hyperscaler take-or-pay 45% and reserved plug, benchmarked to CoreWeave ~2% on-demand + OpenRouter signal)','/workspace/nebius/shared/findings_log.xlsx (Demand bot row, 2026-10-04 18:28 ET)','Findings log','Low-Medium'],
['Nebius','On-demand share converted to capacity basis','~2.6% of capacity (Derived: on-demand priced ~1.54x commit [commit "up to 35% below on-demand" -> 1/0.65]; capacity share = (4%/1.54)/((4%/1.54)+96%) = 2.6%; higher if on-demand utilization <100%)','Q2-26','Capacity','Derived (from Estimated input)',NB_PR+' ; findings_log','Prices page (commit discount up to 35%); team 4% estimate','Low'],
['Nebius','Contracted backlog duration (RPO)','RPO $37,490.6M at 6/30/26: 36% in next 24 months, 40% months 25-48, rest after. RPO EXCLUDES obligations with original duration <=1 year, so <=6-month deals can never show up here','6/30/26','Revenue (backlog)','Reported',NB_FS,'Interim FS, revenue note "Remaining performance obligations"','High'],
['Nebius','Large anchors absorbing 2027 MW','Microsoft 5-yr (~$17.4B committed; all tranches due to date delivered); Meta 5-yr $27B: $12B dedicated capacity from early 2027 + $15B Nebius-optional capacity','Q1-26 / Q2-26','Capacity (context)','Reported',NB_Q1C+' ; '+NB_LET,'Q1-26 call Q&A (Meta structure); Q2-26 letter p.5 (Microsoft tranches)','High'],
# ---- CoreWeave
['CoreWeave','Committed-contract share of revenue','20% (2022), 88% (2023), 96% (2024)','FY22-FY24','Revenue','Reported',CW_S1,'S-1 Risk Factors ("committed contracts accounted for...") and Note: Revenue recognition','High'],
['CoreWeave','Committed-contract share of revenue','Over 98% (FY25); 96% (FY24); 88% (FY23)','FY25','Revenue','Reported',CW_10K,'10-K Risk Factors (take-or-pay vs pay-as-you-go) and Notes to FS - Revenue recognition','High'],
['CoreWeave','Committed-contract share of revenue','98% (3Q25) vs 96% (3Q24); 98%/95% 9M','3Q25','Revenue','Reported',CW_Q3,'10-Q Note: Revenue; Risk Factors','High'],
['CoreWeave','Committed-contract share of revenue','98% (1Q26 and 1Q25)','1Q26','Revenue','Reported',CW_Q1,'10-Q Note: Revenue recognition','High'],
['CoreWeave','Committed-contract share of revenue','98% (2Q26 and 2Q25; 1H26 and 1H25)','2Q26','Revenue','Reported',CW_Q2,'10-Q Note: Revenue recognition; Risk Factors','High'],
['CoreWeave','On-demand (pay-as-you-go) share of revenue','<=2% (Derived: 100% - 98% committed); 4% FY24 (100-96)','FY25 / 2Q26','Revenue','Derived',CW_10K+' ; '+CW_Q2,'As above','High'],
['CoreWeave','Committed contract duration','Typically 2-5 years; weighted-avg ~4 years (12/31/24) -> ~5 years (12/31/25)','FY24 / FY25','Duration','Reported',CW_S1+' ; '+CW_10K,'S-1 Business/MD&A ("typically two to five years"; "weighted-average contract duration of approximately four years"); 10-K Business ("approximately five years")','High'],
['CoreWeave','Shorter-term contracts','Mgmt "really, really aggressive" about selling shorter-term contracts "to extract additional margin"; DDTL 5.5 first financing to include shorter-duration contracts; enterprise wants 2-3 yr terms. Volume/share not disclosed. CAVEAT: shorter committed contracts sit inside the 98% "committed" bucket, so 98% does NOT mean only 2% is <=6 months','2Q26 call (Aug 11 2026)','Qualitative','Reported',CW_CALL,'Q&A (Intrator on DDTL 5.5 / term mix; Agrawal on renewals)','High (quote) / n.a. (%)'],
['CoreWeave','Share of capacity/revenue on <=6-month deals','Not disclosed','2024-2Q26','n.a.','Not disclosed',CW_Q2,'-','n.a.'],
# ---- IREN
['IREN','Contract tenor range','"a variety of contracts with existing customers with terms ranging from month to month up to five years"; primarily multi-year reserved capacity; on-demand to "complement over time"','FY26 (10-K, Aug 27 2026)','Duration','Reported',IR_10K,'Item 1 Business (AI Cloud Services) and Item 1A Risk Factors','High'],
['IREN','Weighted-average contract term','~4 years (weighted by contribution to ARR)','Jul 20 2026','Duration (ARR-weighted)','Reported',IR_JUL,'Press release (Ex.99.1) body + footnote 4','High'],
['IREN','Share of ARR under contract','~85% of >$4bn ARR target under contract','Jul 20 2026','Revenue (ARR)','Reported',IR_JUL,'Ex.99.1 first paragraph','High'],
['IREN','Recent deal tenor; spot/on-demand stance','Recent deals "3-to-5-year deals, not two-year, not spot capacity"; "we are not sitting on capacity to time a spot price"; on-demand/shorter-term = possible future "dabble"','FY26 call (Aug 27 2026)','Qualitative (current short-term share ~nil)','Reported',IR_CALL,'Prepared remarks (Roberts, pricing); Q&A Knoblauch/Cantor (Roberts)','High'],
['IREN','Short-term deal pricing','Not disclosed; "Shorter duration contracts and Mirantis managed services provide further revenue upside". 3-yr contracts >$20M/IT MW, discussions ~$25M; 3-yr pricing +125% since Nov, 5-yr +70%','FY26','Price per MW (capacity)','Reported',IR_PR+' ; '+IR_CALL,'PR highlights "Contracted pricing increasing"; call prepared remarks','High'],
['IREN','Share of AI cloud revenue/capacity that is hourly or <=6 months','Not disclosed (AI Cloud revenue FY26 $128.8M; no split)','FY26','n.a.','Not disclosed',IR_10K,'MD&A AI Cloud Services revenue','n.a.'],
# ---- Nscale
['Nscale','Committed vs on-demand','Committed take-or-pay = "vast majority" of revenue FY24-FY25; on-demand "not currently a material component of our revenue model"','FY24-1H26','Revenue','Reported (qualitative)',NS_S1,'S-1 Risk Factors (take-or-pay model); Business "On-Demand Usage Arrangements"','High'],
['Nscale','Weighted-average contract life','~5.7 years; committed TCV $2.6B active / $103.4B active and contracted (8/31/26)','Aug 31 2026','Duration (TCV)','Reported',NS_S1,'S-1 Prospectus Summary / Business "committed contracts"','High'],
['Nscale','NVIDIA reserved cluster terms','Initial service terms 12-48 months, $1.2B aggregate (shortest disclosed tenor = 12 months)','Jul 2025 MSA','Duration','Reported',NS_S1,'S-1 Business / Material contracts (NVIDIA MSA)','High'],
# ---- Lambda / Crusoe / APLD
['Lambda','On-demand vs reserved share','Not disclosed - no S-1/F-1 on EDGAR as of 2026-10-04 (EDGAR company search + full-text search returned no hits). SemiAnalysis names Lambda/Runpod as focused on on-demand/spot capacity (secondary, qualitative)','2026','n.a.','Not disclosed','https://www.sec.gov/cgi-bin/browse-edgar?company=lambda&type=S-1 ; '+SA,'EDGAR search; SemiAnalysis "Short-term Rental" section','Medium'],
['Crusoe','On-demand vs reserved share','Not disclosed (private). Offers on-demand, spot (via sales) and reserved instance agreements','2026','n.a.','Not disclosed','https://docs.crusoecloud.com/usage-billing/spot.md ; https://www.crusoe.ai/cloud/pricing','Docs: Spot; Reservations overview','Medium'],
['Applied Digital','Lease tenor (colocation landlord, not GPU cloud)','~15-year take-or-pay leases with CoreWeave for 400 MW critical IT (Polaris Forge 1): 0% short-term by construction - low relevance as a comp for this metric','FY25-FY26','Duration','Reported','https://www.globenewswire.com/news-release/2025/07/30/3124427/0/en/Applied-Digital-Reports-Fiscal-Fourth-Quarter-and-Full-Year-2025-Results.html','Fiscal Q4 FY25 release, business highlights','Medium'],
# ---- Industry (secondary)
['Industry (SemiAnalysis) - SECONDARY','Where rental volume transacts','"most of the GPU rental market is transacted on a long-term basis with contracts of at least 6mths and longer"; large neoclouds prefer >=1-yr, 2-3 yr preferred, 5-yr offtakes; short-term (on-demand/spot/<3-mo) "in many cases represents residual capacity". No % share published publicly','Apr 2 2026','Market value/volume (qualitative)','Reported (secondary)',SA,'"Introducing the SemiAnalysis 1Y H100 contract price index" bullets; "GPU Rental Market Structure Today"','Medium'],
['Industry (SemiAnalysis) - SECONDARY','Short-term availability','"On-Demand GPU rental capacity is sold out across all GPU types"; all capacity online until Aug-Sep 2026 already booked','Early 2026','Qualitative','Reported (secondary)',SA,'Intro; "Surge Pricing Comes to the GPU Rental Market"','Medium'],
['Industry (SemiAnalysis) - SECONDARY','H100 term structure','On-demand index ~$2.80-2.83 (Jan-Mar 2026), sold out Apr 2026; 1-yr contract $2.10-2.70/GPU-hr (Apr 2026)','Apr 2026','Price (not share)','Reported (secondary)',SA_IDX,'"Detailed Pricing History" table','Medium'],
['Industry (SemiAnalysis) - SECONDARY','Provider tenor posture','CoreWeave "rarely accept short-term rentals"; Nebius and Crusoe "offer competitive terms for short-term rentals"','Mar 2025','Qualitative','Reported (secondary)',SA_CMX,'ClusterMAX: provider notes','Medium'],
['Industry','Share of neocloud revenue/capacity on <=6-month deals','Not disclosed by any public source found; SemiAnalysis full term structure is paywalled (institutional)','2026','n.a.','Not disclosed',SA,'-','n.a.'],
# ---- Recommendation
['Nebius (recommendation)','Base share of 2027 capacity on deals of 6 months or less','5% of avg 2027 active IT MW (low test 2%, high test 10%; 25% = stress outside evidence). Method: peer ceiling (CRWV on-demand <=2% rev, IREN ~nil, Nscale immaterial) + Nebius own on-demand ~2.6% of capacity (derived) + incremental premium 3-6-mo program ("a new effort", 1 deal signed) of ~2-3%; anchors (MSFT, Meta $12B dedicated from early 2027) take a large slice of 2027 MW','FY27','Capacity','Estimated',NB_LET+' ; '+CW_Q2+' ; '+IR_CALL+' ; '+NS_S1,'See rows above','Low-Medium'],
['Nebius (recommendation)','Revenue share implied by capacity share','5% capacity -> ~9.7% of AI-cloud revenue on those MW (Derived: 5%x$45M / (5%x$45M + 95%x$22M) = 2.25/23.15); 10% -> 18.5%; 2% -> 4.1%','FY27','Revenue (derived from capacity)','Derived',NB_LET,'Uses $45M and $22M midpoints from letter p.2-3','Medium'],
['Nebius (recommendation)','Occupancy haircut for short-term MW','Apply ~80-90% occupancy to the short-term MW (3-6-mo deals leave re-letting gaps between deals); $40-50M/MW is the price while a deal runs, not an annual yield. Effective yield at 85%: ~$38M/MW (Derived: 0.85 x $45M)','FY27','Price per MW (capacity)','Estimated (assumption; no disclosure)',NB_LET,'Analyst assumption','Low'],
]
wb=load_workbook(P)
if 'ShortTerm_Mix' in wb.sheetnames: del wb['ShortTerm_Mix']
ws=wb.create_sheet('ShortTerm_Mix')
ws.append(H)
for r in R: ws.append(r)
for c in ws[1]:
    c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F4E78'); c.alignment=Alignment(wrap_text=True,vertical='top')
for w,col in zip([24,40,80,18,22,24,60,48,18],'ABCDEFGHI'): ws.column_dimensions[col].width=w
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
    if str(row[0].value).startswith('Nebius (recommendation)'):
        for c in row: c.font=Font(bold=True)
ws.freeze_panes='A2'
wb.save(P); print('ShortTerm_Mix rows', len(R), 'sheets', wb.sheetnames)
