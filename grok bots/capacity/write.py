import pickle,csv
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment,PatternFill
cols,R,K,T,S=pickle.load(open('/workspace/nebius/capacity/data.pkl','rb'))
# YE2026 energization build (INFERENCE)
Y=[ # site, low, base, high, basis, source
("Mäntsälä #1",75,75,75,"75 MW expansion completed early 2026 (Reported built)",S['lapp_pdf']),
("Paris (Equinix PA10)",5,5,5,"MW never disclosed; 5 MW = third-party DC Atlas (Estimated, low confidence)","https://dcatlas.io/en/explore/facilities/nebius-equinix-paris"),
("Keflavik (Verne)",10,10,10,"10 MW live since Mar 2025 (Reported)",S['nj_ice']),
("Kansas City (Patmos)",5,20,40,"5 MW phase 1 reported; phase 2 MW n/d; 40 MW = stated max. Base 20 = Estimated midpoint",S['dcd_kc']),
("Vineland (DataOne)",200,200,300,"Q2'25: 100 MW in 2025 + 100 MW in 2026 = 200 MW (Reported plan); 300 MW = site max; Q2'26 call: Vineland part of 2026 connected power. Risk: Bloom switch + site-plan hearing",S['q2_25']),
("London Longcross (Ark)",16,19,19,"16 MW live + 3 MW in delivery (Ark, Jun 2026); LP02 36 MW timing n/d (excluded)",S['uk_ark_dcn']),
("Harlow (Kao)",0,0,4.4,"22 MW total; bulk in KLON-03 from early 2027; 4.4 MW in KLON-02 is a third-party inference","https://northwiseproject.com/research/nebius-harlow-data-center"),
("UK 3rd new site",0,0,0,"not disclosed (counted 0)",S['uk_pr']),
("Modi'in (Mega Or)",8,8,8,"8 MW live since Oct 2025 (Reported)",S['il_80']),
("Masmiyya (Mega DC)",0,22,22,"22 MW, delivery Q3 2026 (landlord disclosure)",S['il_80']),
("Beit Shemesh (Mega DC)",0,29,58,"58 MW staged Q3'26–Q1'27; base = half by YE26 (Estimated: linear over 3 quarters ≈ 2/3, rounded down to 1/2 for slippage)",S['il_80']),
("Madrid-Getafe (Merlin Edged)",9,18,18,"18 MW; phase 1 live Sep 2026, phases 2-3 'coming weeks'",S['es_li']),
("Tallinn Hüüru (Greenergy)",0,22,22,"22 MW first phase 'later in 2026'",S['dcd_ee']),
("Béthune (Azur)",0,120,120,"Nebius CCO: ~half (120 MW) operational by end-2026",S['reut_beth']),
("Birmingham BHM01",0,50,100,"Nebius page: initial-phase energization Q4 2026; Q1'26 letter: operational 2027; phase 1 ~100 MW press-reported; base = 50% (Estimated)",S['bham']),
("Independence MO",0,0,0,"IPL initial power 2H 2026 but MW n/d; letter: operational 2027 → counted 0 (upside not quantifiable)",S['indep']),
("Minnesota colo",0,0,0,"MW not disclosed (region live) → counted 0; upside",S['docs']),
("Oklahoma colo",0,0,0,"MW not disclosed → counted 0; upside",S['f20']),
("France 2nd colo",0,0,0,"not disclosed → 0; upside",S['f20']),
("Lappeenranta / Mäntsälä #2 / Pennsylvania / AIB",0,0,0,"First capacity 2027+ (Reported) / AIB date n/d",S['lapp_pdf']),
]
tot=[round(sum(r[i] for r in Y),1) for i in (1,2,3)]
out='/workspace/nebius/capacity/'
with open(out+'site_list.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(cols); w.writerows(R)
wb=Workbook(); ws=wb.active; ws.title='Sites'
hdr=Font(bold=True,color='FFFFFF'); fill=PatternFill('solid',fgColor='1F3864')
def sheet(ws,head,rows,widths=None):
    ws.append(head)
    for c in ws[1]: c.font=hdr; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='top')
    for r in rows: ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
    for i,_ in enumerate(head):
        ws.column_dimensions[ws.cell(1,i+1).column_letter].width=(widths[i] if widths else 30)
    ws.freeze_panes='C2'
sheet(ws,cols,R,[4,30,28,10,9,9,16,30,30,30,30,16,60,40,35,22,80])
ws2=wb.create_sheet('Site x Quarter KPIs'); sheet(ws2,['Entity','Quarter','Metric','Value','Tag','Method / note','Source'],K,[22,9,22,50,22,60,70])
# wide pivot
Q=["Q1'24","Q2'24","Q3'24","Q4'24","Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26"]
ents=[];[ents.append((e,m)) for e,q,m,*_ in K if (e,m) not in ents]
piv=[]
for e,m in ents:
    row=[e,m]
    for q in Q:
        v=[f"{r[3]} [{r[4]}]" for r in K if r[0]==e and r[2]==m and r[1]==q]
        row.append(v[0] if v else 'not disclosed')
    piv.append(row)
ws3=wb.create_sheet('KPI pivot'); sheet(ws3,['Entity','Metric']+Q,piv,[22,22]+[30]*10)
ws4=wb.create_sheet('YE2026 energized vs guide'); sheet(ws4,['Site','Low MW','Base MW','High MW','Basis (INFERENCE unless Reported)','Source'],[list(r) for r in Y]+[["TOTAL identified",*tot,"Guide: 800 MW–1 GW CONNECTED by YE2026 (not active). Gap to 800: low %.0f / base %.0f / high %.0f MW"%tuple(800-t for t in tot),""]],[34,10,10,10,80,70])
ws5=wb.create_sheet('Company targets history'); sheet(ws5,['Date','Target / disclosure','Source'],T,[12,100,90])
wb.save(out+'site_list.xlsx')
pickle.dump((Y,tot),open(out+'ye26.pkl','wb'))
print(tot)
