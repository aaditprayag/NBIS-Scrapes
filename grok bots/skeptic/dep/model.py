# Python mirror of 03_depreciation.xlsx logic (used for checks, charts, md)
Q=[f"Q{q} {y}" for y in range(2026,2031) for q in range(1,5)][2:]  # Q3'26..Q4'30
N=len(Q); FY=[int(q[-4:]) for q in Q]
A=dict(lives=[4,5,6],gpu=0.79,ngl=12,lag=2,ex_srv=5197.7,ex_age=4,ex_oth=1115.4,niu=7836.4,niu_q3=0.5,
 cx26=23207,h1cx=8130.3,cx27=29135,cx28=38885,cxg29=0.0,cxg30=0.0,
 eb28=13978,rv28=22983,w28=[.22,.24,.26,.28],int_q=78.3,inc_rate=0.07,start_cash=14497.1,
 imp_pct=0.0,cpmw=30.0,yov=None,imp_idx=6,wacc=0.11,g=0.03,tax=0.258,ev=69620,shares=287.66,netcash=226.4,px=242.81)
rv=[904.1,1452,2175,2646,3304,4052]; eb=[354.7,670.7,1277,1581,1883,2372]
cx=[(A['cx26']-A['h1cx'])/2]*2+[A['cx27']/4]*4+[A['cx28']/4]*4+[A['cx28']*(1+A['cxg29'])/4]*4+[A['cx28']*(1+A['cxg29'])*(1+A['cxg30'])/4]*4
gpuin=[];ngin=[]
for t in range(N):
    base=A['niu']*(A['niu_q3'] if t==0 else (1-A['niu_q3']) if t==1 else 0)
    lagc=cx[t-A['lag']] if t>=A['lag'] else 0
    gpuin.append((base+lagc)*A['gpu']); ngin.append((base+lagc)*(1-A['gpu']))
cumg=[A['ex_srv']+sum(gpuin[:t+1]) for t in range(N)]
for t in range(6,10):
    rv.append(A['rv28']*A['w28'][t-6]); eb.append(A['eb28']*A['w28'][t-6])
avg28=sum(cumg[6:10])/4; yld_c=A['eb28']/avg28; yld=A['yov'] if A.get('yov') else yld_c; mrg=A['eb28']/A['rv28']
for t in range(10,N):
    eb.append(yld/4*cumg[t-1]); rv.append(eb[-1]/mrg)   # EBITDA = FY28 yield x prior-qtr gross in service
def dep(L):
    R=dict(ex=[],new=[],ng=[],imp=[])
    remQ=4*L-A['ex_age']
    for t in range(N):
        R['ex'].append(A['ex_srv']/(4*L)*(1-A['imp_pct'] if t>=A['imp_idx'] else 1) if t<remQ else 0)
        R['new'].append((sum(gpuin[v] for v in range(N) if v<t and v>t-4*L)+0.5*gpuin[t]+(0.5*gpuin[t-4*L] if t-4*L>=0 else 0))/(4*L))  # half-quarter convention
        R['ng'].append(A['ex_oth']/(4*A['ngl'])+(sum(ngin[:t])+0.5*ngin[t])/(4*A['ngl']))
        R['imp'].append(A['imp_pct']*A['ex_srv']*max(0,remQ-t)/(4*L) if t==A['imp_idx'] else 0)
    R['da']=[R['ex'][t]+R['new'][t]+R['ng'][t] for t in range(N)]
    R['ebit']=[eb[t]-R['da'][t]-R['imp'][t] for t in range(N)]
    return R
D={L:dep(L) for L in A['lives']}
fcf=[];cum=0;inc=[]
for t in range(N):
    need=max(0,-cum-A['start_cash']); inc.append(need*A['inc_rate']/4)
    f=eb[t]-cx[t]-A['int_q']-inc[t]; fcf.append(f); cum+=f
cumf=[sum(fcf[:t+1]) for t in range(N)]
need=[max(0,-c-A['start_cash']) for c in cumf]
df=[(1+A['wacc'])**(-t/4) for t in range(N)]
pvf=sum(f*d for f,d in zip(fcf,df))
E30=sum(eb[14:18]); G30=cumg[-1]
rdcf={}
for L in A['lives']:
    R=D[L]; M=4*(R['ex'][-1]+R['new'][-1]+R['ng'][-1])  # steady-state replacement = annualized D&A at that life
    Gact=4*L*(R['ex'][-1]+R['new'][-1])
    fT=(E30*(1+A['g'])-M*(1+A['g']))*(1-A['tax'])+A['tax']*0 - A['g']*Gact
    TV=fT/(A['wacc']-A['g']); pvtv=TV*df[-1]; evm=pvf+pvtv
    Ereq=(( (A['ev']-pvf)/df[-1]*(A['wacc']-A['g']) + A['g']*Gact)/((1-A['tax'])*(1+A['g'])))+M
    yreq=Ereq/Gact; yact=E30/Gact
    rdcf[L]=dict(yreq=yreq,yact=yact,mw_req=yreq*A['cpmw']*A['gpu']/mrg,M=M,Gact=Gact,fT=fT,TV=TV,pvtv=pvtv,evm=evm,ps=(evm+A['netcash'])/A['shares'],Ereq=Ereq,mult=TV/E30)
def fy(arr,y): return sum(a for a,f in zip(arr,FY) if f==y)
H1={'eb':365.7,'rv':981.3,'da':{4:471.7+86.1,5:471.7,6:471.7-344.4/6},'cx':8130.3}
if __name__=="__main__":
    print('yield FY28 EBITDA/avg GPU gross',round(yld_c,3),'avg28',round(avg28))
    for y in range(2026,2031):
        print(y,'rev',round(fy(rv,y)+(981.3 if y==2026 else 0)),'EBITDA',round(fy(eb,y)+(365.7 if y==2026 else 0)),'capex',round(fy(cx,y)+(8130.3 if y==2026 else 0)),'FCF(H2 only 26)',round(fy(fcf,y)),
              ' '.join(f"D&A{L}={round(fy(D[L]['da'],y)+(H1['da'][L] if y==2026 else 0))} EBIT{L}={round(fy(D[L]['ebit'],y)+(365.7-H1['da'][L] if y==2026 else 0))}" for L in A['lives']))
    print('Q3 26 D&A by life',[round(D[L]['da'][0],1) for L in A['lives']])
    print('pvFCF',round(pvf),'E30',round(E30),'G30',round(G30),'peak need',round(max(need)))
    for L,r in rdcf.items(): print(L,{k:round(v,1) for k,v in r.items()})
