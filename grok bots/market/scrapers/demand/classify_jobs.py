#!/usr/bin/env python3
"""Rule-based function + geography/site mapping for a jobs_<date>.csv snapshot (from jobs_snapshot.py or Wayback parses).
Functions: Engineering/R&D, Sales & GTM, DC ops/facilities/construction, Hardware/supply chain, G&A, Other.
Usage: classify_jobs.py jobs_YYYY-MM-DD.csv -> jobs_classified_YYYY-MM-DD.csv"""
import csv, re, sys, os
FUNC=[  # (function, regex on title+department) evaluated in order
 ('DC ops / facilities / construction', r'data ?cent(er|re)|facilit|construction|electrical|mechanical engineer|critical (facilit|environment)|(?<!reliability )site (lead|manager|operations)|site manager|physical security|security officer|commissioning|power (engineer|systems)|energy|real estate|site selection|substation|hvac|technician|dcops|dc ops|data hall|rack|cabling|fiber|smart hands|campus|EHS|health and safety|environmental health'),
 ('Hardware / supply chain / capacity', r'hardware|supply chain|procurement|sourcing|logistic|capacity (plan|operations)|asset manag|inventory|HW Automation|component|npi|manufactur|warehouse'),
 ('Sales & GTM', r'sales|\bfield (enablement|planning)|\bSDR\b|solution (specialist|manager)|solutions manager|account (executive|manager|director)|business development|go[- ]to[- ]market|\bgtm\b|solutions? architect|sales engineer|customer success|customer experience|forward deployed|partner|alliances|channel|marketing|developer relations|devrel|brand|communications|events|pre-?sales|revenue operations|pricing|growth|demand gen|commercial|deal desk|technical account'),
 ('G&A', r'financ|account(ant|ing)|tax|treasury|legal|counsel|paralegal|compliance|audit|internal control|people|talent|recruit|hr\b|human resources|hrbp|c&b|compensation|payroll|mobility|immigration|office manager|executive assistant|workplace|corporate development|investor relations|fp&a|controller|procure-to-pay|it support|it infrastructure|enterprise technolog|ceo|cso office|strategy|chief of staff|admin|risk|insurance|government|policy|public affairs'),
 ('Engineering / R&D / product', r'program manager|project manager|delivery|engineer|developer|software|backend|frontend|sre|site reliability|devops|infrastructure|platform|kubernetes|network|security|ml\b|machine learning|research|scientist|data|product (manager|owner|designer|lead)|product|design|architect|qa|test|analyst|technical writer|robotics|ai\b|llm|inference|gpu|system'),
]
def func(title,dept):
    s=f'{title} | {dept}'
    if re.search(r'recruit|talent acquisition',title,re.I): return 'G&A'
    if re.search(r'marketing',title,re.I): return 'Sales & GTM'
    if re.search(r'site reliability|\bSRE\b',title,re.I): return 'Engineering / R&D / product'
    for f,rx in FUNC:
        tgt=title if f.startswith('DC') else s   # DC rule on title only (Lambda's dept 'Data Center Business' is the whole company)
        if f.startswith('DC') and re.search(r'Data Center (Design|Construction)|Physical Security',dept,re.I): tgt=s
        if re.search(rx,tgt,re.I): return f
    return 'Other'
SITES=[ # (site label, regex on location string) -> maps to Nebius announced sites (capacity/site_list.csv)
 ('Vineland, NJ (DataOne)', r'New Jersey|Vineland'),
 ('Kansas City / Independence, MO', r'Kansas City|Independence, Missouri|Missouri'),
 ('Mäntsälä, Finland', r'M[äa]nts[äa]l[äa]'),
 ('Lappeenranta, Finland', r'Lappeenranta'),
 ('Finland (other/unspecified)', r'Finland'),
 ('Paris / Béthune, France', r'B[ée]thune|Paris|France|Hauts-de-France'),
 ('UK (London/Longcross/Harlow/Newport)', r'United Kingdom|London|Newport|Harlow|Surrey|England'),
 ('Israel (Modi\'in / Masmiyya / Beit Shemesh / Tel Aviv)', r'Israel|Tel Aviv|Masmi|Beit Shemesh|Modi'),
 ('Iceland (Keflavik)', r'Iceland|Keflav'),
 ('Spain (Madrid)', r'Spain|Madrid'),
 ('Estonia (Tallinn)', r'Estonia|Tallinn'),
 ('Alabama (Birmingham)', r'Alabama|Birmingham'),
 ('Pennsylvania', r'Pennsylvania|Philadelphia'),
 ('Oklahoma', r'Oklahoma'),
 ('Minnesota', r'Minnesota'),
]
def site(loc):
    hits=[lab for lab,rx in SITES if re.search(rx,loc,re.I)]
    # keep Finland(other) only if no specific Finnish site
    if any('Finland' in h and 'other' not in h for h in hits): hits=[h for h in hits if 'other' not in h]
    return '; '.join(hits) if hits else ''
def country(loc):
    parts=[p.strip() for p in re.split(r';',loc) if p.strip()]
    out=[]
    for p in parts:
        if re.search(r'Remote - Europe',p): out.append('Remote-Europe')
        elif re.search(r'Remote - United States|United States|, (CA|NY|NJ|TX|WA|PA|IL|GA|NV|MO|ND|NC|WY|DC|D\.C\.)\b|Washington, D\.C\.',p): out.append('United States')
        elif re.search(r'Remote',p): out.append('Remote-other')
        else: out.append(p.split(',')[-1].strip())
    return '; '.join(dict.fromkeys(out))
if __name__=='__main__':
    f=sys.argv[1]; rows=list(csv.DictReader(open(f)))
    out=f.replace('jobs_','jobs_classified_')
    with open(out,'w',newline='') as fh:
        w=csv.writer(fh); w.writerow(list(rows[0].keys())+['function','country','nebius_site_(DC/hardware roles only)'])
        for r in rows:
            loc=r['location']+('; '+r['offices'] if r.get('offices') else '')
            fn=func(r['title'],r['department'])
            st=site(loc+' '+r['title']) if (r['company']=='Nebius' and fn in('DC ops / facilities / construction','Hardware / supply chain / capacity')) else ''
            w.writerow(list(r.values())+[fn,country(r['location']),st])
    print(out)
