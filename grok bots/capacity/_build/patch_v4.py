p='build_final_billed.py'; s=open(p).read()
def R(a,b):
    global s
    assert a in s,(a[:90]); s=s.replace(a,b,1)
R("VF=f'Quarterly_IT!$B${t0+4}';","""qi.cell(t0+12,1,'Vineland FY27 billed MW-yr — ENGINE-LINKED (v0.9e lock): engine Microsoft FY27 revenue ÷ $/IT MW-yr — USED IN FINAL TABLE').font=BOLD
c=qi.cell(t0+12,2,182.329216759459); c.font=BLUE; c.fill=FILL; qi.cell(t0+12,3,'D (static copy of /workspace/nebius/model/nebius_model.xlsx v0.9e Judge_Slide B19 = Drivers!B33; Contracts MS_Rev FY27)')
qi.cell(t0+13,1,'memo: engine Microsoft FY27 revenue $M (Contracts tab, capped; T3–T4 ramp + Q1\\'27 genset cap)'); c=qi.cell(t0+13,2,2643.76); c.font=BLUE; c.fill=FILL; qi.cell(t0+13,3,'D (v0.9e Contracts; ~$14.5M/IT MW-yr)')
qi.cell(t0+14,1,'memo: engine-basis BULL Bloom timing ≈ engine + (bull − base tranche-months) = +33.3 MW-yr for one-quarter-earlier T5–T9 starts'); qi.cell(t0+14,2,f'=B{t0+12}+(B{t0+6}-B{t0+5})'); qi.cell(t0+14,3,'D (approx.; assumes engine ramp/cap unchanged)')
VENG=f'Quarterly_IT!$B${t0+12}'; VENGB=f'Quarterly_IT!$B${t0+14}'
VF=f'Quarterly_IT!$B${t0+4}';""")
R("ft.cell(r,10,f'={VF}' if s.startswith('Vineland')","ft.cell(r,10,f'={VENG}' if s.startswith('Vineland')")
R("ft.cell(r,11,f'={VF}') if s.startswith('Vineland')","ft.cell(r,11,f'={VENG}') if s.startswith('Vineland')")
R("'FY27 avg billed IT BASE (D; Vineland at billed tranches)','FY27 avg billed IT HIGH (D; Vineland at billed tranches, base Bloom timing)'",
  "'FY27 avg billed IT BASE (D; Vineland engine-linked 182.3)','FY27 avg billed IT HIGH (D; Vineland engine-linked 182.3)'")
R("('FY27 avg billed IT BASE',536.48,'D (team lock Oct 8 ~536; Vineland billed-tranche basis)','J','Oct 8 re-cut: non-Vineland model method 349.5 + Vineland tranche basis 187.0 (Judge_Slide A19)'),",
  "('FY27 avg billed IT BASE',531.84,'D (team lock v0.9e, engine-linked)','J','nebius_model.xlsx v0.9e Judge_Slide B25 (=Drivers!B34) 531.841; non-Vineland 349.5 + Vineland engine 182.3'),")
R("('FY27 avg billed IT HIGH',663.8,'D (team lock Oct 8 ~664; Vineland billed-tranche basis, base Bloom timing)','K','Oct 8 re-cut: other sites high path (model method) 476.8 + Vineland tranche 187.0'),",
  "('FY27 avg billed IT HIGH',659.16,'D (team lock v0.9e, engine-linked)','K','nebius_model.xlsx v0.9e Judge_Slide B26 (=Drivers!C34) 659.157; non-Vineland high 476.8 + Vineland engine 182.3'),")
R("""memo=[(""","""memo=[('MEMO — superseded Oct 8 ~9:50 ET: A19 typed $2,711M, pre engine link — FY27 BASE 536.48',f'=J{T}-J{VROW}+{VF}','D (superseded)'),
('MEMO — superseded Oct 8 ~9:50 ET: A19 typed $2,711M, pre engine link — FY27 HIGH 663.80',f'=K{T}-K{VROW}+{VF}','D (superseded)'),
(""")
R("""("MEMO — BULL Bloom timing: FY27 HIGH with Vineland BULL tranche timing (T5–T6 Q1'27, T7–T9 Q2'27 = 220.0)",f'=K{T}-K{VROW}+{VBULL}','D'),""",
  """("MEMO — BULL Bloom timing (engine basis): FY27 HIGH with T5–T6 from Q1'27, T7–T9 from Q2'27 (Vineland ≈ 182.3 + 33.3 = 215.7)",f'=K{T}-K{VROW}+{VENGB}','D (approx.)'),
("MEMO — BULL Bloom timing on superseded A19-typed basis (Vineland 220.0) — 696.8",f'=K{T}-K{VROW}+{VBULL}','D (superseded basis)'),""")
R("""paths=[('Base (tranche basis)',f'=J{T}'),('High (tranche basis, base Bloom timing)',f'=K{T}'),('Bull Bloom timing (memo)',f'=K{T}-K{VROW}+{VBULL}')]""",
  """paths=[('Base (engine-linked v0.9e)',f'=J{T}'),('High (engine-linked v0.9e)',f'=K{T}'),('Bull Bloom timing (memo, engine basis approx.)',f'=K{T}-K{VROW}+{VENGB}')]""")
R("""ft.cell(r+8,1,"NOTE (per Grok): the model engine (Contracts tab: MSFT 2027 revenue ~$2,643.8M ≈ ~182 MW-yr; billing MW avg 185.6) is the source of truth. A19's 187.0 overstates it by ~1–5 MW (T3–T4 ramp + Q1'27 genset cap), so the final pair may land a few MW lower once Modeling links A19 to Contracts.").font=Font(italic=True,color='C00000')""",
"""ft.cell(r+8,1,"NOTE: v0.9e accepted — A19 now linked to the model engine (Contracts MS_Rev FY27 $2,643.8M ÷ ~$14.5M = 182.3 MW-yr). Final pair 531.8 / 659.2 replaces 536.5 / 663.8 (A19 typed $2,711M).").font=Font(italic=True,color='C00000')
ft.cell(r+9,1,'Non-Microsoft hurdle: Street FY27 less engine Microsoft revenue ÷ non-Vineland billed MW ($M per billed MW)'); c=ft.cell(r+9,2,27.627757898385); c.font=BLUE; c.fill=FILL; ft.cell(r+9,3,'D (Modeling v0.9e Judge_Slide B20 = Drivers!B9; static copy)')
ft.cell(r+10,1,'  check: (Street − engine MSFT revenue) ÷ (FY27 base − Vineland engine MW)'); ft.cell(r+10,2,f'=(B{r+1}-Quarterly_IT!$B${t0+13})/(J{T}-J{VROW})'); ft.cell(r+10,3,'D (cross-check)')""")
R("r=r+10\nextra=[","r=r+12\nextra=[")
R("'Company totals tie to locked anchors (see Final_Table checks): YE26 billed 257.7 base / 284.2 high; FY27 avg billed 536.5 base / 663.8 high (Oct 8 re-cut, Vineland at billed tranches in both; supersede 475.0 / 645.4);",
  "'Company totals tie to locked anchors (see Final_Table checks): YE26 billed 257.7 base / 284.2 high; FY27 avg billed 531.8 base / 659.2 high (v0.9e engine-linked; Vineland 182.3 MW-yr; supersede 536.5 / 663.8, 475.0 / 645.4);")
R("FY27 avg billed: Vineland BASE = billed-tranche basis (107 MW full-year + $1,159.5M/$14.5M = 187.0, Judge_Slide A19). Vineland HIGH = same 187.0 (base Bloom timing).",
  "FY27 avg billed: Vineland (base and high) = engine-linked 182.3 MW-yr (v0.9e Contracts MS_Rev FY27 $2,643.8M ÷ ~$14.5M; Judge_Slide A19).")
R("RESOLVED Oct 8: FY27 revenue and MW now both on billed-tranche basis.","RESOLVED Oct 8 (v0.9e engine-linked): Vineland FY27 = 182.3 MW-yr from the model engine; FY27 = 531.8 / 659.2; $16M needs 768.75 = 1.45x base / 1.17x high; non-Microsoft hurdle ~$27.6M/billed MW. HISTORY: FY27 revenue and MW first moved to billed-tranche basis.")
R("""('FY27 avg billed BASE','~475 (474.96)','536.5 (536.48)','YES (team decision Oct 8)',""","""('FY27 avg billed BASE / HIGH (v0.9e engine-linked)','~475 / ~645','531.8 / 659.2 (531.84 / 659.16)','YES (v0.9e lock Oct 8 ~9:50 ET)','Vineland 182.3 MW-yr engine-linked; 536.5 / 663.8 superseded (A19 typed $2,711M)'),
('FY27 avg billed BASE (interim)','~475 (474.96)','536.5 (536.48) — superseded ~9:50 ET','interim',""")
R("re-cut v3: FY27 base and high on billed-tranche basis","re-cut v4: FY27 engine-linked v0.9e")
R("(re-cut v3: FY27 base and high on billed-tranche basis)","(re-cut v4: FY27 engine-linked v0.9e)")
R("""'Vineland NJ (DataOne)':('leased (build-to-suit)','Microsoft','Med (YE26) / Low-Med (FY27)',"4 of 9 tranches delivered (R); MW/tranche E (26.7 IT). YE26 and FY27 base on billed-tranche basis (106.7 / 187.0);""",
  """'Vineland NJ (DataOne)':('leased (build-to-suit)','Microsoft','Med (YE26) / Low-Med (FY27)',"4 of 9 tranches delivered (R); MW/tranche E (26.7 IT). YE26 billed-tranche 106.7; FY27 engine-linked 182.3 MW-yr (v0.9e);""")
open(p,'w').write(s); print('patched')
