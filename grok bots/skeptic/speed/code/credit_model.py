"""Study 3: convert-implied credit spread / implied vol, and Merton equity-implied spreads.
Simplified model (documented in README): convert = straight bond (coupons on original principal + accreted principal at maturity,
discounted at UST+spread, semi-annual) + European call on conversion shares with strike = accreted principal (Black-Scholes,
no dividends, r = UST at matching tenor). Ignores issuer call, holder puts, make-whole, borrow cost, and credit risk on the option.
Output is a model estimate, NOT a trading price."""
import math, json
N=lambda x:0.5*(1+math.erf(x/math.sqrt(2)))
def bs_call(S,K,T,r,v):
    d1=(math.log(S/K)+(r+0.5*v*v)*T)/(v*math.sqrt(T)); d2=d1-v*math.sqrt(T)
    return S*N(d1)-K*math.exp(-r*T)*N(d2)
def bond(cpn,acc,T,y):  # per 100 original; semiannual coupons on original principal; y = annual, s.a. compounding
    n=max(1,round(T*2)); pv=0
    for k in range(1,n+1):
        t=T-(n-k)*0.5; pv+=cpn/2/(1+y/2)**(2*t)
    return pv+acc/(1+y/2)**(2*T)
def conv(S,n_sh,cpn,acc,T,r,s,v):  # n_sh shares per 100 original
    return bond(cpn,acc,T,r+s)+n_sh*bs_call(S,acc/n_sh,T,math.log(1+r/2)*2,v)
def solve(f,lo,hi,target,it=200):
    for _ in range(it):
        mid=(lo+hi)/2
        if (f(mid)-target)*(f(lo)-target)<=0: hi=mid
        else: lo=mid
    return (lo+hi)/2
def interp(curve,T):
    ks=sorted(curve); 
    for a,b in zip(ks,ks[1:]):
        if a<=T<=b: return curve[a]+(curve[b]-curve[a])*(T-a)/(b-a)
    return curve[ks[-1]] if T>ks[-1] else curve[ks[0]]
# UST par curves, Treasury daily yield curve CSV (data/ust_2026.csv)
UST={'2026-08-19':{2:4.19,3:4.25,5:4.35,7:4.48,10:4.65},'2026-09-17':{2:4.67,3:4.75,5:4.78,7:4.86,10:4.94},'2026-10-02':{2:4.83,3:4.96,5:5.06,7:5.17,10:5.28}}
NOTES={ # pricing date, S0, shares per 100 original, coupon, accretion, maturity tenor from pricing (yrs)
 'NBIS 0.50% 2030':dict(d='2026-08-19',S=223.90,n=0.31902,c=0.50,a=110,T=(2030+45/365)-(2026+231/365)),
 'NBIS 4.50% 2034':dict(d='2026-08-19',S=223.90,n=0.30802,c=4.50,a=125,T=(2034+45/365)-(2026+231/365)),
 'CRWV 2.875% 2033':dict(d='2026-09-17',S=79.88,n=1.02194,c=2.875,a=100,T=(2033+90/365)-(2026+259/365)),
}
out={'grid':[],'notes':{}}
for k,x in NOTES.items():
    r=interp(UST[x['d']],x['T'])/100
    ytm_nooption=solve(lambda y:bond(x['c'],x['a'],x['T'],y),0.0,0.5,100)  # yield if never converted, at par
    rec=dict(x,r=r,ytm_at_par=ytm_nooption,spread_if_no_option_bp=(ytm_nooption-r)*1e4)
    rec['implied_spread_bp_by_vol']={v:round(1e4*solve(lambda s:conv(x['S'],x['n'],x['c'],x['a'],x['T'],r,s,v/100),-0.05,0.6,100),0) for v in (30,40,50,60,70)}
    rec['implied_vol_by_spread']={s:round(100*solve(lambda v:conv(x['S'],x['n'],x['c'],x['a'],x['T'],r,s/1e4,v),0.01,2.0,100),1) for s in (200,400,600,836)}
    out['notes'][k]=rec
# Mark-to-model 2026-10-02 (S: NBIS $242.81 close 10/2 per 01_priced_in.md; CRWV $89.62 close 10/2 per parts/comps.csv)
S2={'NBIS':242.81,'CRWV':89.62}
for k,x in NOTES.items():
    T2=x['T']-((2026+275/365)-(2026+(231 if 'NBIS' in k else 259)/365)); r2=interp(UST['2026-10-02'],T2)/100
    S=S2[k[:4]]; m={}
    for v in (40,50,60):
        for s in (400,600,836):
            m[f'vol{v}_spr{s}']=round(conv(S,x['n'],x['c'],x['a'],T2,r2,s/1e4,v/100),1)
    out['notes'][k]['mtm_2026-10-02']=m; out['notes'][k]['parity_2026-10-02']=round(x['n']*S,1)
# Merton (equity-implied) spreads, 5y horizon
def merton(E,sE,D,r,T=5):
    # solve V, sV from E=BS(V,D) and sE*E = N(d1)*sV*V
    V=E+D; sV=sE*E/V
    for _ in range(500):
        d1=(math.log(V/D)+(r+0.5*sV**2)*T)/(sV*math.sqrt(T)); d2=d1-sV*math.sqrt(T)
        Ec=V*N(d1)-D*math.exp(-r*T)*N(d2)
        V=V-(Ec-E)/max(N(d1),1e-6); sV=sE*E/(N(d1)*V)
    d1=(math.log(V/D)+(r+0.5*sV**2)*T)/(sV*math.sqrt(T)); d2=d1-sV*math.sqrt(T)
    s=-math.log(N(d2)+V/(D*math.exp(-r*T))*N(-d1))/T
    return dict(V=V,sV=sV,PD_rn=N(-d2),spread_bp=s*1e4)
r5=math.log(1+0.0506/2)*2
out['merton']={
 'CRWV':dict(E=49428.71,sE=0.71,D=39268,src='E: 551.54M sh x $89.62 (parts/comps.csv, 10/2 close); D: pro forma debt $39,268M ex-leases (parts/comps.csv); sE: IV30 ~70-72% (opti-view, 10/2)'),
 'NBIS':dict(E=74960,sE=0.746,D=14270.7,src='E: 308.72M pro forma sh x $242.81 (01_priced_in.md); D: pro forma debt $14,270.7M (01_priced_in.md); sE: IV30 74.6% (alphaquery, 10/2-3)')}
for k,v in out['merton'].items(): v.update(merton(v['E'],v['sE'],v['D'],r5)); v.update({f'spread_bp_sE{int(s*100)}':merton(v['E'],s,v['D'],r5)['spread_bp'] for s in (0.6,0.9,1.1)})
json.dump(out,open('/workspace/nebius/skeptic/speed/data/credit_model_out.json','w'),indent=1,default=str)
print(json.dumps(out,indent=1,default=str))
