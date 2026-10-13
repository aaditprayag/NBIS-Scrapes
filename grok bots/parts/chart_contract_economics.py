import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import json
fm.fontManager.addfont('/usr/share/fonts/truetype/sand-box/google/Inter/Inter-VariableFont_opsz,wght.ttf')
plt.rcParams['font.family']='DejaVu Sans'
N={n['ID']:n for n in json.load(open('/workspace/nebius/parts/contract_economics_norm.json'))}
NEB='#C0392B'; PEER='#2E5C8A'; EST='#9AA5B1'; RUB='#7D5BA6'

fig,(a1,a2)=plt.subplots(1,2,figsize=(16,8.2),gridspec_kw={'width_ratios':[1.15,1]})
# ---- panel 1: $M per IT MW-yr
rows=[('Microsoft–Nebius (GB300)\nIT MW from 100k GPUs',N['N1']['pit'],NEB,'',None),
      ('Microsoft–Nscale all SOWs\n(GB300 + Vera Rubin)',N['P2']['pit'],PEER,'',None),
      ('Microsoft–Nscale/Aker Narvik\n(GB300)',N['P2a']['pit'],PEER,'',None),
      ('Microsoft–IREN (GB300)\n200 MW IT disclosed',N['P1']['pit'],PEER,'',None),
      ('Meta–Nebius #1\n(MW est., low conf.)',N['N2']['pit'],NEB,'//',None),
      ('Anthropic–Nscale\n(Vera Rubin, 460 MW IT)',N['O1']['pit'],RUB,'',None),
      ('OpenAI–Oracle (4.5 GW,\nbasis unknown; /1.2 if gross)',N['O2']['pmw'],EST,'//',N['O2']['pit'])]
y=list(range(len(rows)))[::-1]
for yi,(lab,v,c,h,hi) in zip(y,rows):
    a1.barh(yi,v,color=c if not h else 'white',edgecolor=c,hatch=h,height=0.62,lw=1.5)
    txt=f'${v:.1f}M' + (f'–{hi:.1f}M' if hi else '')
    if hi: a1.plot([v,hi],[yi,yi],color=c,lw=3)
    a1.text((hi or v)+0.3,yi,txt,va='center',fontsize=11,fontweight='bold',color='#222')
# Nebius MSFT alternative bases
yn=y[0]
for val,lab in [(N['N1']['pmw'],'300 MW site basis\n(team $1,324/MW-hr)'),(17392.9/5/240,'240 MW IT')]:
    a1.plot([val],[yn],marker='D',color='white',mec='black',ms=9,zorder=5)
a1.text(0.3,yn-0.47,'◆ $11.6M on 300-MW site basis (team $1,324/MW-hr)   ◆ $14.5M on ~240 MW IT',fontsize=8.5,color='#444',va='center')
a1.set_yticks(y); a1.set_yticklabels([r[0] for r in rows],fontsize=10.5)
for x,lab,col in [(12,'Nebius "2026 base"\n~$12M/MW',NEB),(20,'Nebius Q2-26 deals >$20M/MW\nIREN new 3-yr >$20M/IT MW','#2D7D46')]:
    a1.axvline(x,ls='--',color=col,lw=1.3); a1.text(x+0.2,-0.95,lab,fontsize=9,color=col,va='bottom')
a1.set_xlim(0,24); a1.set_ylim(-1.1,len(rows)-0.3)
a1.set_xlabel('$M of contract revenue per IT MW per year (TCV / term / IT MW)',fontsize=11)
a1.set_title('Contract $ per IT MW-yr',fontsize=13,fontweight='bold',loc='left')
a1.grid(axis='x',alpha=0.25); [a1.spines[s].set_visible(False) for s in ('top','right')]

# ---- panel 2: $/GPU-hr
rows2=[('Microsoft–Nebius (GB300)',N['N1']['gh'],NEB,N['N1o']['gh']),
       ('Microsoft–Nscale all SOWs (mix)',N['P2']['gh'],PEER,None),
       ('Microsoft–Narvik (GB300)',N['P2a']['gh'],PEER,None),
       ('Microsoft–IREN (GB300, est.)',N['P1']['gh'],PEER,9700e6/(72000*8760*5)),
       ('Anthropic–Nscale (Vera Rubin)',N['O1']['gh'],RUB,None),
       ('Anthropic–xAI (chip n.d.)',N['O3']['gh'],EST,None)]
y2=list(range(len(rows2)))[::-1]
for yi,(lab,v,c,hi) in zip(y2,rows2):
    a2.barh(yi,v,color=c,height=0.62)
    if hi: a2.plot([v,hi],[yi,yi],color='black',lw=2.5)
    a2.text((hi or v)+0.15,yi,'\\$'+f'{v:.2f}'+((f'–{hi:.2f}'+(r' (w/ \$19.4B option)' if c==NEB else ' (GPU/MW range)')) if hi else ''),va='center',fontsize=11,fontweight='bold')
a2.set_yticks(y2); a2.set_yticklabels([r[0] for r in rows2],fontsize=10.5)
for x,lab,col,yy,ha in [(3.70,'Astra GB300\n3-yr lease $3.70','#555',-0.95,'left'),(9.50,'Nebius B300\nOD $9.50',NEB,-0.95,'right'),(10.50,'CoreWeave\nGB200 list\n$10.50','#2D7D46',2.6,'left')]:
    a2.axvline(x,ls='--',color=col,lw=1.2); a2.text(x+(0.1 if ha=='left' else -0.1),yy,lab,fontsize=8.5,color=col,va='bottom',ha=ha)
a2.set_xlim(0,12); a2.set_ylim(-1.1,len(rows2)-0.3)
a2.set_xlabel('Implied $ per GPU-hour = TCV / (GPUs × 8,760 × years)',fontsize=11)
a2.set_title('Implied $ per GPU-hr',fontsize=13,fontweight='bold',loc='left')
a2.grid(axis='x',alpha=0.25); [a2.spines[s].set_visible(False) for s in ('top','right')]

fig.suptitle('Microsoft pays Nebius more per GB300 than its other neocloud anchors — but every anchor sits far below 2026 market pricing',
             fontsize=15,fontweight='bold',x=0.01,ha='left',y=0.985)
fig.text(0.01,0.935,'Hyperscaler / anchor GPU-capacity contracts, normalized to IT MW (GB300 491 GPUs/IT MW, Vera Rubin 320) and to GPU-hours. Red = Nebius. Hatched = estimated/low confidence. Black whiskers = upside or range.',fontsize=10.5,color='#444',ha='left')
fig.text(0.01,0.012,'Sources: Nebius 20-F FY25 Note 1 p.F-17/F-18, Microsoft & Meta 6-Ks, Q2-26 shareholder letter p.3-4; IREN 8-K & FY26 results; Nscale S-1 & Microsoft press release; Aker-Nscale Narvik release; Bloomberg via DCD (">100k GB300" Nebius); Bloomberg (Anthropic–Nscale 6 yr/460 MW);\n'
         'WSJ (Oracle–OpenAI); SpaceX 424B4 via CCIR (xAI); list prices nebius.com/prices, coreweave.com/pricing, SAVRN (Astra). Implied rates assume ratable billing over full term.\nNebius $/GPU-hr uses a press GPU floor (">100k") so is a ceiling; Meta–Nebius #2 and all CoreWeave orders omitted (MW/GPUs not disclosed). Built 2026-10-04 ET.',
         fontsize=7.6,color='#555',ha='left')
plt.tight_layout(rect=[0,0.07,1,0.92])
out='/workspace/nebius/market/charts/contract_usd_per_mw_yr.png'
plt.savefig(out,dpi=180); print(out)
