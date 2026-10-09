p='build_final_billed.py'; s=open(p).read()
def R(a,b):
    global s
    assert a in s,(a[:80]); s=s.replace(a,b,1)
R("VB=f'Quarterly_IT!$B${vr+3}'",
"""VB=f'Quarterly_IT!$B${vr+3}'
t0=vr+5
qi.cell(t0,1,'Vineland FY27 avg billed — billed-tranche basis (same as revenue engine / Judge_Slide A19)').font=BOLD
qi.cell(t0+1,1,'T1–T4 IT MW billed full year 2027'); c=qi.cell(t0+1,2,107); c.font=BLUE; c.fill=FILL; qi.cell(t0+1,3,'D (model v0.9d Judge_Slide A19: 107 MW x $14.5M = $1,551.5M; 4 x 26.7 E)')
qi.cell(t0+2,1,"T5–T9 2027 revenue $M (T5–T6 from Q2'27, T7–T9 from Q3'27)"); c=qi.cell(t0+2,2,1159.5); c.font=BLUE; c.fill=FILL; qi.cell(t0+2,3,'D (vineland_watch.md §2; Judge_Slide A19)')
qi.cell(t0+3,1,'Microsoft rate $M per IT MW-yr'); c=qi.cell(t0+3,2,14.5); c.font=BLUE; c.fill=FILL; qi.cell(t0+3,3,'D ($17,392.9M / 5 yrs / 240 IT MW)')
qi.cell(t0+4,1,'Vineland FY27 avg billed IT MW (tranche basis) — USED IN FINAL TABLE'); qi.cell(t0+4,2,f'=B{t0+1}+B{t0+2}/B{t0+3}'); qi.cell(t0+4,3,'D'); qi.cell(t0+4,1).font=BOLD
qi.cell(t0+5,1,'memo: same, from tranche months (4x26.67 full yr + 26.67x(2x9+3x6)/12)'); qi.cell(t0+5,2,f'=4*B{vr+2}+B{vr+2}*(2*9+3*6)/12'); qi.cell(t0+5,3,'D (0.3 MW rounding vs 107)')
qi.cell(t0+6,1,"memo: BULL tranche timing (T5–T6 from Q1'27, T7–T9 from Q2'27)"); qi.cell(t0+6,2,f'=4*B{vr+2}+B{vr+2}*(2*12+3*9)/12'); qi.cell(t0+6,3,'D (tranche track bull starts)')
qi.cell(t0+7,1,'memo: Judge_Slide A27 basis (quarter-average of quarter-end delivered tranches 106.7/106.7/160/240/240)'); qi.cell(t0+7,2,'=((106.667+106.667)/2+(106.667+160)/2+(160+240)/2+(240+240)/2)/4'); qi.cell(t0+7,3,'D')
_v=qrow["Vineland NJ (DataOne)"]
qi.cell(t0+8,1,'memo: SUPERSEDED Oct 8 — Vineland FY27 on energized-lagged basis (did not match revenue engine)'); qi.cell(t0+8,2,f'=(D{_v}+2*E{_v}+2*F{_v}+2*G{_v}+H{_v})/8'); qi.cell(t0+8,3,'D (superseded)')
qi.cell(t0+9,1,"memo: Contracts engine cross-check (model v0.9d Contracts r141/r142): MSFT 2027 billing IT MW quarter avg 185.6; 2027 revenue capped $2,643.8M (T3–T4 ramp, Q1'27 genset cap 102.4)"); c=qi.cell(t0+9,2,185.605); c.font=BLUE; c.fill=FILL; qi.cell(t0+9,3,'D (static copy)')
VF=f'Quarterly_IT!$B${t0+4}'; VBULL=f'Quarterly_IT!$B${t0+6}'; VA27=f'Quarterly_IT!$B${t0+7}'; VOLD=f'Quarterly_IT!$B${t0+8}'""")
R("""            ft.cell(r,8,f'={VB}'); ft.cell(r,9,f'={VB}')""","""            ft.cell(r,8,f'={VB}'); ft.cell(r,9,f'={VB}'); VROW=r""")
R("""        ft.cell(r,10,f'=(Quarterly_IT!D{q}+2*Quarterly_IT!E{q}+2*Quarterly_IT!F{q}+2*Quarterly_IT!G{q}+Quarterly_IT!H{q})/8')""",
"""        ft.cell(r,10,f'={VF}' if s.startswith('Vineland') else f'=(Quarterly_IT!D{q}+2*Quarterly_IT!E{q}+2*Quarterly_IT!F{q}+2*Quarterly_IT!G{q}+Quarterly_IT!H{q})/8')""")
R("'FY27 avg billed IT BASE (D)','FY27 avg billed IT HIGH (D)'","'FY27 avg billed IT BASE (D; Vineland at billed tranches)','FY27 avg billed IT HIGH (D; lock 645.4, Vineland energized-lagged — see checks)'")
R("('FY27 avg billed IT BASE',474.96,'D (team lock ~475)','J','model v0.9d Judge_Slide B25 (=SUMIF(Q_Yr,2027,MW_Avg)/4)'),",
  "('FY27 avg billed IT BASE',536.48,'D (team lock Oct 8 ~536; Vineland billed-tranche basis)','J','Oct 8 re-cut: non-Vineland model method 349.5 + Vineland tranche basis 187.0 (Judge_Slide A19)'),")
R("('FY27 avg billed IT HIGH',645.4,'D (team lock ~645)','K','model v0.9d Judge_Slide B26'),",
  "('FY27 avg billed IT HIGH',645.4,'D (team lock ~645; NOT strictly tranche-consistent — see memo rows)','K','model v0.9d Judge_Slide B26 (A27 memo 646.9)'),")
R("""r=A+2+len(anchors)
extra=[""","""r=A+2+len(anchors)
memo=[('MEMO — SUPERSEDED Oct 8: FY27 avg billed BASE 474.96 (energized-lagged basis, did not match revenue engine)',f'=J{T}-J{VROW}+{VOLD}','D (superseded)'),
('MEMO — FY27 HIGH strictly consistent: other sites high + Vineland 187.0 (base tranche timing)',f'=K{T}-K{VROW}+{VF}','D — differs from 645.4 lock by +18.4'),
("MEMO — FY27 HIGH with Vineland BULL tranche timing (T5–T6 Q1'27, T7–T9 Q2'27 = 220.0)",f'=K{T}-K{VROW}+{VBULL}','D'),
('MEMO — FY27 HIGH per Judge_Slide A27 (Vineland quarter-avg of quarter-end tranches = 170.0)',f'=K{T}-K{VROW}+{VA27}','D (model shows 646.9)'),
('MEMO — Non-Vineland FY27 avg billed BASE: new (col D) / old slide-implied 475 − 187 (col E)',f'=J{T}-J{VROW}','D'),
('MEMO — Street FY27 $12.3B ÷ FY27 avg billed BASE ($M/MW): new (col D) / old (col E)',f'=12300/J{T}','D (consensus $12.3B, E)')]
for k,(n,fml,t) in enumerate(memo):
    ft.cell(r+k,1,n).font=Font(italic=True); ft.cell(r+k,4,fml); ft.cell(r+k,3,t)
ft.cell(r+4,5,f'=474.96-{VF}'); ft.cell(r+4,6,f'=E{r+4}/D{r+4}-1'); ft.cell(r+4,6).number_format='0.0%'; ft.cell(r+4,7,'col F = change in slide-implied non-Vineland $/MW (~$33.3M -> ~$27.4M; Modeling to confirm $)')
ft.cell(r+5,5,'=12300/474.96'); ft.cell(r+5,7,'old basis ~$25.9M -> new ~$22.9M')
r=r+len(memo)
extra=[""")
R('''"2) Vineland FY27 avg billed = 125.5 under the model's lock method (energized lagged 1Q). The judge slide's 'Vineland ~187 of ~475' uses tranche revenue ($1,551.5M + $1,159.5M) ÷ $14.5M, which is a different basis. On a consistent tranche basis the FY27 base total would be ~536 (=475.0 − 125.5 + 187.0, D). The lock stays 475; Modeling should reword the line as '~187 MW-equivalent of revenue', not as a share of 475.",''',
'''"2) RESOLVED Oct 8: FY27 revenue and MW now both on billed-tranche basis. Vineland FY27 = 187.0 (T1–T4 107 MW full year + T5–T9 $1,159.5M ÷ $14.5M), replacing 125.5; FY27 base = 536.5 (was 475.0). Non-Vineland MW rises ~288 → ~349.5, so the slide's implied non-Vineland $/MW falls ~17.6% (~$33.3M → ~$27.4M, Modeling to confirm). HIGH lock 645.4 is NOT strictly consistent: same tranche basis = 663.8 (base timing) or 696.8 (bull timing); A27 memo 646.9 uses quarter-average of quarter-end tranches (Vineland 170). Contracts-engine cross-check: MSFT 2027 billing MW avg 185.6; revenue $2,643.8M (T3–T4 ramp + Q1'27 genset cap) ≈ 182 MW-yr at $14.49M. 26.7 IT/tranche is E: if early tranches are smaller, revenue and MW fall together.",''')
R("""('FY27 avg billed base / high','~475 / ~645','475.0 / 645.4','No','Now split by site (new); totals tie'),""",
"""('FY27 avg billed BASE','~475 (474.96)','536.5 (536.48)','YES (team decision Oct 8)','Vineland moved to billed-tranche basis (125.5 → 187.0) to match revenue engine; 475 kept as superseded memo'),
('FY27 avg billed HIGH','~645 (645.4)','645.4 (lock kept)','No — flagged','Strictly consistent tranche basis = 663.8 (base timing) / 696.8 (bull); A27 memo 646.9'),""")
R("""('NEW flag','—',"Vineland FY27 125.5 (lock method) vs ~187 (judge-slide tranche-revenue basis)",'Flag','Consistent tranche basis would give ~536 FY27 base; lock unchanged')]""",
"""('Street FY27 $/billed MW (base)','~$25.9M','~$22.9M','Derived change','$12.3B consensus (E) ÷ 536.5; Modeling to confirm'),
('Non-Vineland FY27 avg billed base','~288 (slide-implied)','~349.5','Derived change','Slide-implied non-Vineland $/MW ~−17.6%')]""")
R("'Company totals tie to locked anchors (see Final_Table checks): YE26 billed 257.7 base / 284.2 high; FY27 avg billed 475.0 / 645.4; YE26 energized IT 301.7; YE26 connected 941 base / 561 low.',",
  "'Company totals tie to locked anchors (see Final_Table checks): YE26 billed 257.7 base / 284.2 high; FY27 avg billed 536.5 base (Oct 8 re-cut, Vineland at billed tranches; supersedes 475.0) / 645.4 high (lock kept; strictly consistent 663.8); YE26 energized IT 301.7; YE26 connected 941 base / 561 low.',")
R('''"  FY27 avg billed (base/high) = model v0.9d method''','''"  FY27 avg billed: Vineland BASE = billed-tranche basis (107 MW full-year + $1,159.5M/$14.5M = 187.0, Judge_Slide A19). Other sites (and Vineland HIGH) = model v0.9d method''')
R("f'Thu Oct 8 2026 ~9:30 AM ET.","f'Thu Oct 8 2026 ~9:45 AM ET (re-cut: FY27 base on billed-tranche basis).")
R("f'Built: Thu Oct 8 2026 ~9:30 AM ET.","f'Built: Thu Oct 8 2026 ~9:45 AM ET (re-cut v2: FY27 base on billed-tranche basis).")
R("""'Vineland NJ (DataOne)':('leased (build-to-suit)','Microsoft','Med',"4 of 9 tranches delivered (R) but MW/tranche is E (26.7 IT); YE26 billed 106.7 (billed-tranche basis) > energized 81.8; FY27 hinges on Bloom (Case B -$966M @~35%)",""",
"""'Vineland NJ (DataOne)':('leased (build-to-suit)','Microsoft','Med (YE26) / Low-Med (FY27)',"4 of 9 tranches delivered (R); MW/tranche E (26.7 IT). YE26 and FY27 base on billed-tranche basis (106.7 / 187.0); FY27 hinges on Bloom T5–T9 (Case B -$966M @~35%); smaller early tranches cut revenue and MW together",""")
open(p,'w').write(s); print('patched')
