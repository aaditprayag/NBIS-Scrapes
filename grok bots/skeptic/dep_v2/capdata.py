from openpyxl import load_workbook
wb=load_workbook('/workspace/nebius/capacity/nebius_capacity.xlsx',data_only=True)
q=wb['Quarterly MW']; hdr=[c.value for c in q[1]]
cols=[hdr.index(f"Q{n} {y}") for y in range(2026,2031) for n in range(1,5)][1:]   # Q2'26..Q4'30 (19)
OWN=('Independence','Butler','Birmingham','Lappeenranta','Mantsala')
def num(v):
    try: return float(v)
    except: return 0.0
tot={};it={};own={};vin={}
for r in q.iter_rows(min_row=2,values_only=True):
    site,blk,basis=r[0],r[1],r[2]
    if not site or not blk: continue
    for sc,key in (('low','Energized MW low'),('high','Energized MW high'),('base','Energized MW base')):
        if site=='COMPANY' and blk==f'Identified sites - energized' : pass
        if str(blk).startswith(key) and site!='COMPANY':
            vals=[num(r[c]) for c in cols]
            tot[sc]=[a+b for a,b in zip(tot.get(sc,[0]*19),vals)]
            if basis=='IT': it[sc]=[a+b for a,b in zip(it.get(sc,[0]*19),vals)]
            if str(site).startswith(OWN): own[sc]=[a+b for a,b in zip(own.get(sc,[0]*19),vals)]
            if str(site).startswith('Vineland'): vin[sc]=vals
# also Q2'26 values for lag
c2=hdr.index("Q2 2026")
q2={}
for r in q.iter_rows(min_row=2,values_only=True):
    if r[0] and r[0]!='COMPANY' and r[1] and str(r[1]).startswith('Energized MW'):
        for sc in('low','high','base'):
            if str(r[1]).startswith(f'Energized MW {sc}'): q2[sc]=q2.get(sc,0)+num(r[c2])
CAP=dict(tot=tot,it=it,own=own,vin=vin,q2=q2)
if __name__=="__main__":
    for sc in('low','base','high'): print(sc,'YE26',tot[sc][2],'YE27',tot[sc][6],'IT ye27',it[sc][5],'own ye27',own[sc][5],'vin',vin[sc][:6],'q2',q2[sc])
