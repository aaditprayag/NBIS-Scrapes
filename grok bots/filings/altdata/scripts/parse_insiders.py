"""Parse Form 4 / 4A / 144 XMLs in raw/insiders -> scripts/insiders.json"""
import glob,json,re,os,xml.etree.ElementTree as ET
B='https://www.sec.gov/Archives/edgar/data/1513845/'
def strip(t):
    for e in t.iter():
        if '}' in e.tag: e.tag=e.tag.split('}')[1]
    return t
def v(e,p):
    x=e.find(p); return (x.text or '').strip() if x is not None and x.text else ''
L={a:(d,f,doc) for d,f,a,doc in json.load(open('raw/insider_filings_list.json'))}
f4=[];f144=[]
for fn in sorted(glob.glob('raw/insiders/*.xml')):
    form,fd,acc=os.path.basename(fn)[:-4].split('_')
    r=strip(ET.parse(fn).getroot())
    url=B+acc.replace('-','')+'/'+L[acc][2].split('/')[-1]
    idx=B+acc.replace('-','')+'/'+acc+'-index.htm'
    if form.startswith('4'):
        owners=[v(o,'reportingOwnerId/rptOwnerName') for o in r.findall('reportingOwner')]
        rel=r.find('reportingOwner/reportingOwnerRelationship')
        title=v(rel,'officerTitle') or ('Director' if v(rel,'isDirector') in('1','true') else '') or ('10% owner' if v(rel,'isTenPercentOwner') in('1','true') else '')
        plan=v(r,'aff10b5One')
        fns={f.get('id'):(f.text or '').strip() for f in r.findall('footnotes/footnote')}
        for t in r.findall('nonDerivativeTable/nonDerivativeTransaction'):
            ids=[x.get('id') for x in t.iter('footnoteId')]
            notes=' '.join(fns.get(i,'') for i in dict.fromkeys(ids))
            code=v(t,'transactionCoding/transactionCode')
            sh=float(v(t,'transactionAmounts/transactionShares/value') or 0)
            px=v(t,'transactionAmounts/transactionPricePerShare/value')
            px=float(px) if px else None
            nl=notes.lower()
            if 'withholding' in nl or 'sell-to-cover' in nl or 'cover' in nl and 'tax' in nl: kind='RSU tax sell-to-cover'
            elif plan in('1','true') or '10b5-1' in nl: kind='10b5-1 plan'
            else: kind='discretionary/other'
            f4.append(dict(form=form,filed=fd,acc=acc,owner='; '.join(owners),title=title,date=v(t,'transactionDate/value'),
              code=code,ad=v(t,'transactionAmounts/transactionAcquiredDisposedCode/value'),shares=sh,price=px,
              value=round(sh*px,2) if px else None,after=v(t,'postTransactionAmounts/sharesOwnedFollowingTransaction/value'),
              plan_flag=plan,kind=kind,notes=notes[:400],url=url,index=idx))
    else:
        fdat=r.find('formData')
        si=fdat.find('securitiesInformation')
        rem=v(fdat,'remarks')
        pads=[x.text for x in fdat.iter('planAdoptionDate')]
        f144.append(dict(form='144',filed=fd,acc=acc,seller=v(fdat,'issuerInfo/nameOfPersonForWhoseAccountTheSecuritiesAreToBeSold'),
          rel='/'.join(x.text for x in fdat.iter('relationshipToIssuer')),
          shares=float(v(si,'noOfUnitsSold') or 0),aggval=float(v(si,'aggregateMarketValue') or 0),
          approx_date=v(si,'approxSaleDate'),broker=v(si,'brokerOrMarketmakerDetails/name'),
          acquired=v(fdat,'securitiesToBeSold/natureOfAcquisitionTransaction'),
          plan_adopt=';'.join(p for p in pads if p),plan10b51=bool(pads) or '10B5' in rem.upper().replace('-',''),
          remarks=rem[:300],url=url,index=idx))
json.dump(dict(f4=f4,f144=f144),open('scripts/insiders.json','w'),indent=1)
print(len(f4),len(f144))
from collections import Counter
print(Counter((x['code'],x['kind']) for x in f4))
for x in f144: print(x['filed'],x['seller'][:22],x['rel'][:30],int(x['shares']),int(x['aggval']),x['approx_date'],x['plan_adopt'],x['acquired'][:30])
