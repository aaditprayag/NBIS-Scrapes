p='build_final_billed.py'; s=open(p).read()
def R(a,b):
    global s
    assert a in s,(a[:90]); s=s.replace(a,b,1)
R("VF=f'Quarterly_IT!$B${t0+4}';","""qi.cell(t0+10,1,'memo: SUPERSEDED Oct 8 — Vineland FY27 HIGH on energized-lagged basis'); qi.cell(t0+10,2,f'=(J{_v}+2*K{_v}+2*L{_v}+2*M{_v}+N{_v})/8'); qi.cell(t0+10,3,'D (superseded)')
VOLDH=f'Quarterly_IT!$B${t0+10}'
VF=f'Quarterly_IT!$B${t0+4}';""")
R("""        ft.cell(r,11,f'=(Quarterly_IT!J{q}""","""        ft.cell(r,11,f'={VF}') if s.startswith('Vineland') else ft.cell(r,11,f'=(Quarterly_IT!J{q}""")
R("'FY27 avg billed IT HIGH (D; lock 645.4, Vineland energized-lagged — see checks)'","'FY27 avg billed IT HIGH (D; Vineland at billed tranches, base Bloom timing)'")
R("('FY27 avg billed IT HIGH',645.4,'D (team lock ~645; NOT strictly tranche-consistent — see memo rows)','K','model v0.9d Judge_Slide B26 (A27 memo 646.9)'),",
  "('FY27 avg billed IT HIGH',663.8,'D (team lock Oct 8 ~664; Vineland billed-tranche basis, base Bloom timing)','K','Oct 8 re-cut: other sites high path (model method) 476.8 + Vineland tranche 187.0'),")
R("""('MEMO — FY27 HIGH strictly consistent: other sites high + Vineland 187.0 (base tranche timing)',f'=K{T}-K{VROW}+{VF}','D — differs from 645.4 lock by +18.4'),""",
  """('MEMO — SUPERSEDED Oct 8: FY27 avg billed HIGH 645.4 (Vineland energized-lagged in high)',f'=K{T}-K{VROW}+{VOLDH}','D (superseded)'),""")
R("""("MEMO — FY27 HIGH with Vineland BULL tranche timing""","""("MEMO — BULL Bloom timing: FY27 HIGH with Vineland BULL tranche timing""")
R("""r=r+len(memo)
extra=[""","""r=r+len(memo)+1
ft.cell(r,1,'MW NEEDED AT $16M (D; Street FY27 and $16M are blue inputs)').font=BOLD
ft.cell(r+1,1,'Street FY27 revenue consensus ($M)'); c=ft.cell(r+1,2,12300); c.font=BLUE; c.fill=FILL; ft.cell(r+1,3,'E (consensus, per Filings/Modeling)')
ft.cell(r+2,1,'Comps price ($M per billed IT MW-yr)'); c=ft.cell(r+2,2,16); c.font=BLUE; c.fill=FILL; ft.cell(r+2,3,'E (top of disclosed comps band $9.7–16.7M)')
ft.cell(r+3,1,'FY27 avg billed MW needed = Street ÷ $16M'); ft.cell(r+3,2,f'=B{r+1}/B{r+2}'); ft.cell(r+3,3,'D')
for j,h in enumerate(['Path','FY27 avg billed MW','Gap MW (needed − path)','Needed ÷ path'],1): ft.cell(r+4,j,h).font=BOLD
paths=[('Base (tranche basis)',f'=J{T}'),('High (tranche basis, base Bloom timing)',f'=K{T}'),('Bull Bloom timing (memo)',f'=K{T}-K{VROW}+{VBULL}')]
for k,(n,fml) in enumerate(paths):
    rr=r+5+k; ft.cell(rr,1,n); ft.cell(rr,2,fml); ft.cell(rr,3,f'=$B${r+3}-B{rr}'); ft.cell(rr,4,f'=$B${r+3}/B{rr}'); ft.cell(rr,4).number_format='0.00"x"'
ft.cell(r+8,1,"NOTE (per Grok): the model engine (Contracts tab: MSFT 2027 revenue ~$2,643.8M ≈ ~182 MW-yr; billing MW avg 185.6) is the source of truth. A19's 187.0 overstates it by ~1–5 MW (T3–T4 ramp + Q1'27 genset cap), so the final pair may land a few MW lower once Modeling links A19 to Contracts.").font=Font(italic=True,color='C00000')
r=r+10
extra=[""")
R("""HIGH lock 645.4 is NOT strictly consistent: same tranche basis = 663.8 (base timing) or 696.8 (bull timing); A27 memo 646.9 uses quarter-average of quarter-end tranches (Vineland 170).""",
  """HIGH now also on tranche basis (Oct 8 team decision): 663.8 (base Bloom timing; 645.4 superseded; 696.8 bull-timing memo; A27 memo 646.9 used quarter-avg of quarter-end tranches).""")
R("""('FY27 avg billed HIGH','~645 (645.4)','645.4 (lock kept)','No — flagged','Strictly consistent tranche basis = 663.8 (base timing) / 696.8 (bull); A27 memo 646.9'),""",
  """('FY27 avg billed HIGH','~645 (645.4)','663.8 (663.80)','YES (team decision Oct 8)','Vineland in high moved to billed-tranche basis (168.5 → 187.0, base Bloom timing); 645.4 superseded memo; 696.8 bull memo'),
('MW needed at $16M vs paths','gap vs high 645 ≈ 123','768.75 needed: gap 232 (base, 1.43x) / 105 (high, 1.16x) / 72 (bull)','Derived change','Street $12.3B (E) ÷ $16M (E)'),""")
R("FY27 avg billed 536.5 base (Oct 8 re-cut, Vineland at billed tranches; supersedes 475.0) / 645.4 high (lock kept; strictly consistent 663.8)",
  "FY27 avg billed 536.5 base / 663.8 high (Oct 8 re-cut, Vineland at billed tranches in both; supersede 475.0 / 645.4)")
R("Other sites (and Vineland HIGH) = model v0.9d method","Vineland HIGH = same 187.0 (base Bloom timing). Other sites = model v0.9d method")
R("re-cut v2: FY27 base on billed-tranche basis","re-cut v3: FY27 base and high on billed-tranche basis")
R("(re-cut: FY27 base on billed-tranche basis)","(re-cut v3: FY27 base and high on billed-tranche basis)")
open(p,'w').write(s); print('patched')
