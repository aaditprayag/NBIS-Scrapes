# Nebius vs peers unit-economics memo - Oct 8 2026 refresh (2-page memo + page-3 one-page summary table).
# ~10:00 ET: forward bar / footnote re-pointed to model v0.9e locks (via comparison.xlsx Nebius_Inputs); multiples vs K02b shown to 2 dp.
# Every KPI number is read from peers/comparison.xlsx via common_data.py (Comparison / Ranking / Nebius_Inputs /
# Contracts_$MWyr / Sites); Microsoft gross->IT read from peers/price_or_mw.xlsx Nebius_Check. n/d where undisclosed.
import sys; sys.path.insert(0, "/workspace/nebius/peers/memo")
from common_data import *
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1f,0x4e,0x8c); GREY = RGBColor(0x55,0x55,0x55)
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.left_margin = sec.right_margin = Inches(0.55)
sec.top_margin = Inches(0.45); sec.bottom_margin = Inches(0.4)
st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(8.9)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
st.paragraph_format.space_after = Pt(1.5); st.paragraph_format.space_before = Pt(0)
st.paragraph_format.line_spacing = 1.0

def para(text="", bold_lead=None, size=None, color=None, after=1.5, before=0, italic=False, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
    if align is not None: p.alignment = align
    if bold_lead:
        r = p.add_run(bold_lead); r.bold = True
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = color
    r = p.add_run(text); r.italic = italic
    if size: r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return p

def rich(parts, after=1.5, style=None, size=None, p=None):
    if p is None:
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
    for t in parts:
        txt, fmt_ = (t if isinstance(t, tuple) else (t, ""))
        r = p.add_run(txt); r.bold = "b" in fmt_; r.italic = "i" in fmt_
        if size: r.font.size = Pt(size)
    return p

def bullet(parts, after=1.2):
    p = rich(parts, after=after, style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.16); p.paragraph_format.first_line_indent = Inches(-0.12)
    return p

def head(t, before=4):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(t.upper()); r.bold = True; r.font.size = Pt(9.2); r.font.color.rgb = NAVY
    pPr = p._p.get_or_add_pPr(); b = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    for k,v in (("w:val","single"),("w:sz","4"),("w:space","1"),("w:color","1F4E8C")): bt.set(qn(k),v)
    b.append(bt); pPr.append(b)
    return p

def shade(cell, hexcol):
    tcPr = cell._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
    s.set(qn("w:val"),"clear"); s.set(qn("w:color"),"auto"); s.set(qn("w:fill"),hexcol); tcPr.append(s)

def table(data, widths, fsize=7.0, header_fill="1F4E8C", first_bold=False, align_left_cols=(0,), flagcol=None, nebcol=None, wrap_rows=None):
    t = doc.add_table(rows=len(data), cols=len(widths)); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"; t.autofit = False
    tblPr = t._tbl.tblPr; m = OxmlElement("w:tblCellMar")
    for k,v in (("top",8),("left",40),("bottom",8),("right",40)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"),str(v)); e.set(qn("w:type"),"dxa"); m.append(e)
    tblPr.append(m)
    for j,gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))): gc.set(qn("w:w"), str(int(widths[j]*1440)))
    fc = {"Better":"DCEFE0","Worse":"F8D9D5","In line":"EDEDED"}
    for i,row in enumerate(data):
        for j,val in enumerate(row):
            c = t.cell(i,j); c.width = Inches(widths[j])
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j in align_left_cols else WD_ALIGN_PARAGRAPH.CENTER
            parts = val if isinstance(val, list) else [(str(val), "")]
            for txt, fm_ in parts:
                r = p.add_run(txt); r.font.size = Pt(fsize); r.bold = "b" in fm_; r.italic = "i" in fm_
                if i == 0: r.bold = True; r.font.color.rgb = RGBColor(0xff,0xff,0xff)
            if i == 0: shade(c, header_fill)
            else:
                if nebcol is not None and j == nebcol:
                    shade(c, "E8EEF7")
                    for r in p.runs: r.bold = True
                if flagcol is not None and j == flagcol and str(val) in fc:
                    shade(c, fc[str(val)])
                    for r in p.runs: r.bold = True
    borders = OxmlElement("w:tblBorders")
    for e in ("top","left","bottom","right","insideH","insideV"):
        el = OxmlElement(f"w:{e}"); el.set(qn("w:val"),"single"); el.set(qn("w:sz"),"2"); el.set(qn("w:color"),"BFBFBF"); borders.append(el)
    tblPr.append(borders)
    # repeat header / no row split
    for i,row in enumerate(t.rows):
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
    return t

N = "Nebius"
def v(k,p="Nebius",d=None): return f(k,p,d)
def pct_gap(a, b): return (a/b - 1)*100

# ---------- values used in prose (all read from the workbook) ----------
fd_debt, fd_conv, fd_v09 = ni("N_FD_DEBT"), ni("N_FD_IFCONV"), ni("N_FD_V09D")
fd_sens = ni("D_FD_CONV_SENS")
ps_drop = (1 - fd_debt/fd_conv)*100
fd_mid = FDC["FD_MID"]/1e6; px30, px34, refpx = FDC["PX_2030"], FDC["PX_2034"], FDC["REF_PX"]
ye_lo, ye_hi = ni("D_YE26_LO_BASE"), ni("D_YE26_HI_HIGH")
st_b, st_x, st_h, st_xh = ni("D_STREET_FY27_BASE"), ni("D_STREET_FY27_NONMSFT"), ni("D_STREET_FY27_HIGH"), ni("D_STREET_FY27_NONMSFT_HIGH")
msft_rev, msft_mw = ni("N_MSFT_FY27_REV"), ni("N_MSFT_FY27_MW")
# 2026-10-08 ~23:45 ET: identified-site energized range from the current Capacity lock (Quarterly MW!N124/N125; was 152-194 on v1)
import openpyxl as _ox
_q = _ox.load_workbook("/workspace/nebius/capacity/nebius_capacity.xlsx", data_only=True)["Quarterly MW"]
ID_LO, ID_HI = _q["N124"].value, _q["N125"].value
assert (ID_LO, ID_HI) == (157, 229) and _q["N1"].value == "Q2 2026"
bill_ye_b, bill_ye_h, bill_27_b, bill_27_h = ni("N_BILL_YE26_BASE"), ni("N_BILL_YE26_HIGH"), ni("N_BILL_FY27_BASE"), ni("N_BILL_FY27_HIGH")
k02b_mw = ni("D_K02B_BASE_MW")
msft_house = CON["N1"][22]; assert abs(msft_house - 14.494) < 0.001, msft_house  # $14.49M [E, house] = model Contracts!E12
cmp_ids = ["P1","NEW1","P2","O1","NEW2","O2"]
peer_con = {cid: (CON[cid][22] if isinstance(CON[cid][22],(int,float)) else CON[cid][21]) for cid in cmp_ids}
pc_lo, pc_hi = min(peer_con.values()), max(peer_con.values())
assert (pc_lo, pc_hi) == (BAND[BAND_LO_ID], BAND[BAND_HI_ID])
vin = next(r for r in SITES if r[0]=="Nebius" and r[1] and r[1].startswith("Vineland"))
den = next(r for r in SITES if r[0]=="CoreWeave" and r[1] and r[1].startswith("Denton"))
abi = next(r for r in SITES if r[0]=="Crusoe" and r[1] and r[1].startswith("Abilene"))
cap_int, int_exp = ni("N32"), ni("N31")
k13x = raw("K13",N)/raw("K13","CoreWeave")

# ---------- Header ----------
p = para("Nebius (NBIS) vs. AI-cloud peers: where the unit economics are better or worse, and why", size=12.5, color=NAVY, after=1)
p.runs[0].bold = True
para("To: NBIS research team   |   From: Aadit Prayag   |   Thu Oct 8, 2026 (refresh of the Oct 4 draft)   |   Data: Q2'26 (Jun-26) unless noted; peers at latest available period   |   Source: peers/comparison.xlsx (model v0.9e locks; Sites rebuilt Oct 8 ~11:15 PM ET to match the four peer site appendices)",
     size=7.1, color=GREY, after=3)

head("Bottom line", before=0)
rich([("Clearest beat: revenue density. ", "b"),
      f"Nebius earns ${v('K02b')}M per average billed IT MW (K02b, E) vs. ${v('K02b','CoreWeave')}M at CoreWeave, and ${v('K02')}M per average active MW (K02) vs. a ${fm('K02')}M peer median "
      f"(CoreWeave {v('K02','CoreWeave')}, IREN {v('K02','IREN')}). It also leads on adj. EBITDA per MW (${v('K05')}M vs. ${v('K05','CoreWeave')}M at CoreWeave), realized $/GPU-hr "
      f"(${v('K07')} vs. ${fm('K07')} median, E) and top-customer share ({v('K24')}% vs. {v('K24','CoreWeave')}% CoreWeave, {v('K24','Nscale')}% Nscale). ",
      ("Clearest trail: capital intensity. ", "b"),
      f"Trailing-4Q cash capex per MW added is ${v('K13')}M (E) vs. ${v('K13','CoreWeave')}M at CoreWeave ({k13x:.1f}x) and a ${fm('K13')}M median, and {v('K40')}% of contractual interest is capitalized "
      f"(median {fm('K40',0)}%), because Nebius is pre-funding owned 2027 capacity. Funding cost and backlog coverage are in line "
      f"({v('K28')}% blended debt cost vs. {fm('K28',1)}% median; RPO {v('K23')}x run-rate revenue vs. {fm('K23',1)}x). The Nebius MW denominator is a Capacity estimate, so read the density lead as directional."],
     after=2)
k02b_now = raw('K02b', N)
rich([("Forward bar (billed-IT basis, context only, not scored): ", "b"),
      f"the YE26 run-rate guide implies ${ye_lo:.1f}M per billed IT MW ($7B / {bill_ye_b:.1f} MW base) to ${ye_hi:.1f}M ($9B / {bill_ye_h:.1f} MW high). "
      f"Street FY27 ex-Microsoft is ${st_x:.1f}M per billed MW ((${ni('N_STREET_FY27')/1000:.1f}B - ${msft_rev/1000:.2f}B Microsoft) / ({bill_27_b:.1f} - {msft_mw:.1f} Vineland billed MW-yr), model v0.9e\u00b2); "
      f"that is the figure to hold against peers, whose $/MW mixes carry no $14.5M Microsoft anchor (blended ${st_b:.1f}M is a memo only; high path ${st_h:.1f}M blended / ${st_xh:.1f}M ex-Microsoft). "
      f"Peer contract prices run ${pc_lo:.1f}-{pc_hi:.1f}M per MW-yr (top: Lambda-Anthropic ~${pc_hi:.1f}M, {BAND_TERM:.0f}-yr, press, {BAND_CONF} confidence; page 3). Against today's K02b of ${k02b_now:.2f}M, the bar is roughly {st_x/k02b_now:.2f}x on FY27 ex-Microsoft and {ye_lo/k02b_now:.2f}x on the YE26 base guide "
      f"({ye_hi/k02b_now:.2f}x on the YE26 high path): it needs a much richer exit mix (company cites Q3 short-term deals >$40M/MW), not today's pricing."], after=2)

# ---------- Peer by peer ----------
head("Peer by peer: where Nebius is better / worse (same KPI definitions as the Comparison)")
pb = [
 ["Peer", "Nebius better", "Nebius worse", "Net"],
 ["CoreWeave",
  f"Rev/billed IT MW {v('K02b')} vs {v('K02b','CoreWeave')}; adj. EBITDA/MW {v('K05')} vs {v('K05','CoreWeave')}; top customer {v('K24')}% vs {v('K24','CoreWeave')}%; debt cost {v('K28')}% vs {v('K28','CoreWeave')}%; RPO cover {v('K23')}x vs {v('K23','CoreWeave')}x; life {v('K16',d=0)} vs {v('K16','CoreWeave',0)} yrs (implied {v('K18')} vs {v('K18','CoreWeave')}).",
  f"T4Q capex/MW added {v('K13')} vs {v('K13','CoreWeave')}; net PP&E/MW {v('K14')} vs {v('K14','CoreWeave')}; capitalized interest {v('K40')}% vs {v('K40','CoreWeave')}%; RPO within 24m {v('K22')}% vs {v('K22','CoreWeave')}%; contracted/active {v('K04')}x vs {v('K04','CoreWeave')}x.",
  "Beat: density. Trail: capex/MW."],
 ["IREN",
  f"Rev/avg MW {v('K02')} vs {v('K02','IREN')} (IREN Sep-25, E); $/GPU-hr {v('K07')} vs {v('K07','IREN')}; incremental rev/new MW {v('K43')} vs {v('K43','IREN')}; rev/GPU hardware {v('K44')}x vs {v('K44','IREN')}x; debt/RPO {v('K39')} vs {v('K39','IREN')}; implied life {v('K18')} vs {v('K18','IREN')}.",
  f"Debt cost {v('K28')}% vs {v('K28','IREN')}% (convert-heavy blend; IREN's Aug-26 GPU financing is 9.0%, iren_kpis.md); RPO cover {v('K23')}x vs {v('K23','IREN')}x; RPO within 24m {v('K22')}% vs {v('K22','IREN',1)}%. Capex/MW and concentration n/d.",
  "Beat: rev/MW. Trail: debt cost."],
 ["Oracle OCI",
  f"Incremental rev/new MW {v('K43')} vs {v('K43','Oracle OCI')} (E); RPO cover {v('K23')}x vs {v('K23','Oracle OCI')}x; within 24m {v('K22')}% vs {v('K22','Oracle OCI')}% (E); life {v('K16',d=0)} vs {v('K16','Oracle OCI',0)} yrs (implied {v('K18')} vs {v('K18','Oracle OCI')}).",
  f"T4Q capex/MW {v('K13')} vs {v('K13','Oracle OCI')} (E; Oracle capex 100% to OCI = upper bound); {v('K27')} customers >=10% vs {v('K27','Oracle OCI')}; debt cost {v('K28')}% vs {v('K28','Oracle OCI')}%; debt/RPO {v('K39')} vs {v('K39','Oracle OCI')}. Rev/MW n/d (no installed MW).",
  "Beat: backlog cover. Trail: capex/MW, concentration."],
 ["Nscale",
  f"Rev/period-end MW {v('K01')} vs {v('K01','Nscale')}; adj. EBITDA/MW {v('K05')} vs {v('K05','Nscale')}; $/GPU-hr {v('K07')} vs {v('K07','Nscale')} (E); top customer {v('K24')}% vs {v('K24','Nscale')}%; within 24m {v('K22')}% vs {v('K22','Nscale')}%; implied life {v('K18')} vs {v('K18','Nscale')}.",
  f"Gross in-service PP&E/MW {v('K15')} vs {v('K15','Nscale')}; net PP&E/MW {v('K14')} vs {v('K14','Nscale')}; capitalized interest {v('K40')}% vs {v('K40','Nscale',0)}%; RPO cover {v('K23')}x vs {v('K23','Nscale',0)}x (Nscale pre-ramp). Debt cost: Nscale {v('K28','Nscale')}% is legacy debt (excluded); marginal ~6.25% (E).",
  "Beat: EBITDA/MW. Trail: PP&E/MW."],
 ["Lambda",
  f"Rev/period-end MW {v('K01')} vs {v('K01','Lambda')} (Lambda Q2'25, press inputs); $/GPU-hr {v('K07')} vs {v('K07','Lambda')} (E); rev/GPU {v('K06')} vs {v('K06','Lambda')} $K/GPU-yr.",
  f"Price realization {v('K09')}% vs {v('K09','Lambda')}% of list (E); top >=10% customers combined {v('K26')}% vs {v('K26','Lambda')}% (Lambda Amazon+Microsoft, P). Capex, life, debt cost, RPO n/d (Lambda DDTL {v('K29','Lambda')}% R is facility-level; Nebius K29 n/d).",
  "Beat: rev/MW. Trail: realization (directional)."],
 ["Crusoe",
  f"Rev/period-end MW {v('K01')} vs {v('K01','Crusoe')} (E; Crusoe FY26E revenue on ~1 GW GROSS MW, landlord+cloud mix: not like-for-like).",
  f"Contracted/active {v('K04')}x vs {v('K04','Crusoe')}x; evidenced-active hyperscale site: Abilene ~{abi[7]:.0f} MW delivered ({abi[6]:.0f}-{abi[8]:.0f}; D on R, Oracle Sep 10; IT vs gross basis unstated) vs Vineland {vin[6]:.0f}-{vin[8]:.0f} MW (base {vin[7]:.0f}). Net PP&E/MW {v('K14','Crusoe')} is shell+power only (excluded). Most KPIs n/d.",
  "Beat: rev/MW (weak comp). Trail: execution."],
]
table(pb, [0.78, 3.05, 2.88, 0.69], fsize=6.9, align_left_cols=(0,1,2,3), nebcol=None)
para(f"$M/MW-yr unless noted; capex and PP&E in $M/MW. Nebius MW = Capacity estimate (E). K02b = revenue per average BILLED IT MW (Nebius {k02b_mw:.1f} IT MW; CoreWeave K02b = K02, active ~ billed); other peers publish no billed-MW series (n/d). "
     f"Flags and medians: comparison.xlsx Ranking; full scorecard on page 3. \u00b2FY27 billed IT MW = team model v0.9e (source of truth; Drivers B34/C34): {bill_27_b:.1f} base / {bill_27_h:.1f} high, engine-linked (non-Vineland path + Vineland {msft_mw:.1f} billed MW-yr = Microsoft FY27 revenue ${msft_rev:,.1f}M / $14.5M). "
     f"Today's K02b ({k02b_now:.2f}) is blended, i.e. includes Microsoft revenue.", size=6.6, color=GREY, after=1)

# ---------- Figure ----------
pf = doc.add_paragraph(); pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.paragraph_format.space_before = Pt(2); pf.paragraph_format.space_after = Pt(0)
pf.add_run().add_picture("/workspace/nebius/peers/memo/fig_unit_economics.png", width=Inches(7.3))
pf.add_run().add_break(WD_BREAK.PAGE)

# ---------- Page 2: why ----------
head("Why Nebius is better", before=0)
bullet([("Mix and pricing. ","b"),
        f"Realized ${v('K07')}/GPU-hr (E) against a fleet-weighted list of ${v('K08')} is a {v('K09')}% price-realization index, vs. {v('K09','CoreWeave')}% at CoreWeave (list ${v('K08','CoreWeave')}) and {v('K09','Lambda')}% at Lambda (E, Q2'25). "
        "CoreWeave's book is dominated by long-dated reserved hyperscaler pricing; more of Nebius's revenue is near list (our inference; the $/GPU-hr also includes non-GPU-hour revenue). "
        "Q2'26 'landmark' deals averaged $20-25M ACV per MW (company, MW basis unstated) vs. a 2026 base of ~$12M/MW."])
bullet([("Young, Blackwell-heavy fleet; little dilution yet. ","b"),
        f"Revenue per $ of GPU hardware is {v('K44')}x vs. {v('K44','CoreWeave')}x at CoreWeave and {v('K44','IREN')}x at IREN; incremental revenue per new MW is ${v('K43')}M vs. a ${fm('K43')}M median. "
        f"Timing flatters Nebius: energized {ni('N16'):.1f} \u2192 {ni('N17'):.0f} MW in Q2'26 (Estimated, Capacity base case), not a growth signal: +{ni('N17')-ni('N16'):.1f} MW is small next to the gap between the 800 MW-1 GW YE26 connected target (company) and our {bill_ye_b:.1f} MW YE26 billed IT path (E). CoreWeave's new MW bill with a lag."])
bullet([("Owned sites show up in EBITDA per MW. ","b"),
        f"Adj. EBITDA per average MW ${v('K05')}M (E) vs. ${v('K05','CoreWeave')}M at CoreWeave, which leases shells at ~${CON['F4'][21]:.1f}-{CON['F5'][21]:.1f}M per MW-yr (Applied Digital, Galaxy facility leases) inside operating cost. Nebius's owned capacity moves that cost into capex and depreciation."])
bullet([("Conservative depreciation and broader customer base. ","b"),
        f"5-year GPU/server life (from 4 in 2026) vs. 6 at CoreWeave and Oracle; implied life {v('K18')} yrs vs. {fm('K18')} median (CoreWeave {v('K18','CoreWeave')} after the Oct 8 K18 fix). "
        f"Top customer {v('K24')}% (then {v('K25',d=0)}%, {ni('N27'):.0f}%; {v('K26')}% combined) vs. CoreWeave {v('K24','CoreWeave')}%/{v('K25','CoreWeave',0)}% ({v('K26','CoreWeave')}% combined)."], after=1)

head("Why Nebius is worse")
bullet([(f"Capex per MW added ${v('K13')}M vs. ${fm('K13')}M median. ","b"),
        f"Mostly timing: $7.8B of PP&E was not yet in use at Jun-26 (2027 capacity), and Nebius buys land and builds shells and substations where CoreWeave leases. "
        f"On CoreWeave's in-service definition (K12) Nebius FY25 is ${v('K12')}M/MW vs. ${v('K12','CoreWeave')}M: in line, so the gap is pre-funding, not overpaying. "
        f"Gross in-service PP&E per MW is still higher (${v('K15')}M vs. ${fm('K15')}M median): owned infrastructure plus costlier GPUs per MW."])
bullet([(f"Capitalized interest {v('K40')}% of contractual interest (median {fm('K40',0)}%). ","b"),
        f"${cap_int:.1f}M capitalized vs. ${int_exp:.1f}M reported Q2'26 interest, so P&L interest understates the cost of carry by about a third; it returns as depreciation when those assets go live."])
bullet([("Execution. ","b"),
        f"Vineland has ~{vin[7]:.0f} MW (range {vin[6]:.0f}-{vin[8]:.0f}) energized of {vin[5]:.0f} MW contracted (Yellow/Red): 62 rented gensets with a ~107 MW IT ceiling, a $1.07M NJDEP fine with a permit deadline ~early/mid Nov, and stop-construction orders not publicly lifted. "
        f"Contracted/active is {v('K04')}x vs. {v('K04','CoreWeave')}x at CoreWeave; only CoreWeave Denton (~{den[7]:.0f} MW) and Crusoe Abilene (~{abi[7]:.0f} MW delivered per Oracle Sep 10, D on R; range {abi[6]:.0f}-{abi[8]:.0f}; IT vs gross basis unstated) show hundreds of evidenced-active MW."], after=1)

head("In line: funding and backlog")
para(f"Blended debt cost {v('K28')}% vs. {fm('K28',1)}% median (CoreWeave {v('K28','CoreWeave')}%, Oracle {v('K28','Oracle OCI')}%, IREN {v('K28','IREN')}%); implied cash cost {v('K36')}% vs. {fm('K36',1)}%. "
     "Marginal secured debt costs more across the group (CoreWeave ~9-10%, IREN 9.0%, Lambda DDTL 6.78%), so owned-site funding will likely cost more than the blend. "
     f"RPO is {v('K23')}x run-rate revenue ({fm('K23',1)}x median) with {v('K22')}% due within 24 months ({fm('K22',0)}%); debt/RPO {v('K39')}x vs. {fm('K39')}x. "
     f"The Microsoft anchor works out to about ${msft_house:.2f}M per IT MW-yr on our house estimate ($17.4B over five years on ~240 IT MW, model Contracts!E12; neither company states the MW), vs. IREN-Microsoft ${peer_con['P1']:.1f}M, "
     f"Nscale-Microsoft ${peer_con['P2']:.1f}M, Nscale-Anthropic ${peer_con['O1']:.1f}M and Lambda-Anthropic ${peer_con['NEW2']:.1f}M (band top; P, {BAND_CONF}): mid-pack, so today's density lead comes from realized mix, not a richer anchor.", after=1)

head("Share-count convention (debt vs. if-converted)")
para(f"Primary convention at today's price: the new Feb-2030 (conv. ${px30:.2f}) and Feb-2034 (conv. ${px34:.2f}) notes count as DEBT while NBIS trades below their conversion prices, so FD = {fd_debt:.2f}M with the $5.75B face in debt "
     f"(at the ${refpx:.2f} 10/7 close both series are out of the money). If-converted (price >= ${px34:.2f}): FD = {fd_conv:.2f}M with ALL convert debt removed from EV; between ${px30:.2f} and ${px34:.2f} only the 2030s convert ({fd_mid:.2f}M). "
     f"Spread +{fd_sens:.2f}% shares (per-share values ~{ps_drop:.1f}% lower at constant equity value); model v0.9e solves each scenario under the treatment consistent with its own PT. Never mix share and debt sides. {fd_v09:.2f}M (v0.9d, 6/30 pre-exchange) is historical only. No KPI in this memo is per-share, so nothing above moves.", after=1)

head("What would change the view / what to watch")
bullet([("Nebius Q3'26 print (~Nov 10, est.; not announced by the company): ","b"),
        f"a reported energized or billing MW figure would replace our (E) denominators; at the top of the {ID_LO:.0f}-{ID_HI:.0f} MW identified-site range, period-end revenue/MW drops to ~${ni('N01')*4/ID_HI:.1f}M (from ${v('K01')}M). "
        "Also watch capex vs. in-service additions, capitalized interest and concentration. As the YE26 build lands, revenue/MW will dilute the way CoreWeave's did."])
bullet([("Vineland permits: ","b"), "NJDEP genset permit (~early/mid Nov), lifting of the Aug 6/10 stop-construction orders, Bloom energization."])
bullet([("Peers: ","b"), "Nscale S-1/A (IPO talk ~Oct 15, press, unconfirmed) with active MW; whether CoreWeave's ~70% July price increase lifts revenue/MW off $8.2M; IREN Horizons 2-4; Abilene buildings 7-8 (618 of ~824 MW delivered) and the Texas data-center permit pause (TCEQ report Oct 19); any Oracle installed-MW disclosure (site appendices)."], after=1)

head("Data-quality note")
para(f"Nebius MW is a Capacity & Power estimate (mixed IT/gross; {ni('N17'):.0f} MW base, identified sites {ID_LO:.0f}-{ID_HI:.0f}); Nebius GPUs (~{ni('N20')/1000:.0f}k) are modeled, so every per-MW and per-GPU Nebius figure is (E). "
     "K02b uses Capacity billing IT MW, not energized or gross MW. Lambda and Crusoe are press-based, mixed-vintage points; Oracle has no installed-MW figure and capex is 100% allocated to OCI (upper bound); "
     "IREN K02 is Sep-25. Most medians rest on one to four peers. CoreWeave K18 now reads 6.47 yrs (pack bug 16.67 fixed Oct 8). Tags: R reported, P press, D derived, E estimated. "
     "Sources: peers/comparison.xlsx; crwv/iren/oracle_kpis.md; private_peers.md; sites/site_verification.md + appendix_*.pdf (Oct 8); capacity/site_verification.md; filings/fd_share_bridge.md; model/nebius_model.xlsx v0.9e (Drivers, Judge_Slide).",
     size=7.3, color=GREY, after=0)

# ---------- Page 3: one-page summary table ----------
pb_ = doc.paragraphs[-1]; pb_.add_run().add_break(WD_BREAK.PAGE)
p = para("One-page summary: Nebius vs. each peer (comparison.xlsx Comparison / Ranking, Q2'26 unless noted)", size=10.5, color=NAVY, after=1)
p.runs[0].bold = True
para("Thu Oct 8, 2026  |  Flag = Ranking tab (Nebius vs. comparable-peer median; >10% gap = Better/Worse)  |  value (tag)  |  n/d = not disclosed / not computable on the same definition", size=6.8, color=GREY, after=2)

def cellv(k, p_):
    x = raw(k, p_)
    if x is None:
        s = ROWS[k][COL[p_]]; return "n/d" if s in (None,"n/d") else str(s)
    d = SUM_DEC.get(k, 1)
    if k == "K16" or (k == "K22" and abs(x-round(x))>1e-9): d = 0 if abs(x-round(x))<1e-9 else 1
    s = f"{x:,.{d}f} ({stag(tag(k,p_))})"
    if excluded(k, p_): s += "\u2020"
    return s
def medv(k):
    m = med(k); d = 2 if k == "K16" else SUM_DEC.get(k,1)
    return f"{m:,.{d}f}" if isinstance(m,(int,float)) else "n/d"

hdr = ["KPI (unit)", "Nebius", "CoreWeave", "IREN", "Oracle OCI", "Nscale", "Lambda", "Crusoe", "Peer median", "Flag"]
srows = [hdr]
lastg = None
for lab, k, u, g in SUMMARY:
    srows.append([[(g + ": ", "b"), (f"{lab} [{u}]", "")]] + [cellv(k, p_) for p_ in PEERS] + [medv(k), flag(k)])
table(srows, [1.95, 0.64, 0.64, 0.6, 0.62, 0.6, 0.5, 0.5, 0.5, 0.48], fsize=7.2, nebcol=1, flagcol=9)
para(f"\u00b9K02b = revenue per average BILLED IT MW (Capacity billing path; Nebius {k02b_mw:.1f} IT MW avg Q1-Q2'26), not energized or gross MW; CoreWeave K02b = K02 (active ~ billed); other peers publish no billed-MW series. "
     f"K02 uses average energized MW (Nebius {ni('N16'):.1f}/{ni('N17'):.0f} MW, Capacity estimate). \u2020Shown but excluded from the median (Nscale K28 = legacy $137M debt; marginal secured ~6.25% E). "
     "Periods: IREN K02 Sep-25; Oracle May-26 (FQ4 FY26), K18 Aug-26; Nscale H1'26; Lambda/Crusoe mostly n/d (press-only private peers). Nscale K23 is pre-ramp (tiny H1'26 revenue vs. $56.4B RPO). "
     "Lower is better for capex/MW, cost of debt and concentration; shorter life = more conservative (Better). K12 (CoreWeave in-service definition) is Nebius FY25.",
     size=6.2, color=GREY, after=3)

# Contracts table
p = para("Hyperscaler / AI-lab compute contracts: implied price per MW (comparison.xlsx Contracts_$MWyr; TCV / term / MW)", size=8.6, color=NAVY, after=1)
p.runs[0].bold = True
con, con_nd, con_fac = contract_rows()
def num(x, d=1, pre=""):
    return f"{pre}{x:,.{d}f}" if isinstance(x,(int,float)) else ("n/d" if x in (None,"","n/d","n/a") else str(x))
crow = [["Provider", "Customer", "Signed", "TCV ($M)", "Tag", "Term (yrs)", "MW stated / IT MW", "$M per MW-yr (stated MW)", "$M per IT MW-yr"]]
for r in con:
    tcv = num(r["tcv"], 0)
    mw = f"{num(r['mw'],0)} / {num(r['itmw'],0)}"
    stated = num(r["stated"], 1); it = num(r["it"], 1)
    if r["id"] == "NAGG1":
        stated = "20-25 (ACV/MW, R)"; it = "basis n/d"
    if r["id"] == "N1":
        it = f"{CON['N1'][22]:.2f} (E, house; model Contracts!E12)"; mw = "n/d / 240 (E)"
    if r["id"] == "N2" and isinstance(r["it"], (int,float)):
        it = f"{it} (MW E, low conf.)"
    if r["id"] == "NEW16":
        stated += " (naive LB)"
    signed = "n/d" if r["signed"] in (None,"") else str(r["signed"]).replace(" (press)", " (P)").replace(" (confirmed)", "")
    tagt = str(r["tcvtag"] or "n/d").replace("R (up to, contingent)", "R (up to; contingent)")
    if r["id"] == BAND_HI_ID: tagt = f"{tagt} ({BAND_CONF}; band top)"
    crow.append([r["provider"], r["customer"], signed, tcv, tagt, num(r["term"], 1) if isinstance(r["term"], (int,float)) else num(r["term"]), mw, stated, it])
table(crow, [0.95, 1.45, 0.78, 0.6, 0.72, 0.42, 0.72, 0.68, 1.08], fsize=6.8, align_left_cols=(0,1,8))
nd_list = "; ".join(f"{r['provider']}-{r['customer']}" for r in con_nd).replace("Crusoe (JV) / Oracle-Oracle for OpenAI (Abilene Stargate)", "Crusoe JV / Oracle-OpenAI (Abilene Stargate)")
fac_list = "; ".join(f"{r['provider']}-{r['customer'].split(' (')[0]} {r['stated']:.2f}" for r in con_fac if isinstance(r['stated'],(int,float)))
fac_nd = [r for r in con_fac if not isinstance(r['stated'],(int,float))]
para("NON-COMPUTE facility leases (shell + power only, NOT comparable to compute $/MW): " + fac_list
     + (f"; {', '.join(r['provider']+'-'+r['customer'] for r in fac_nd)} n/d" if fac_nd else "") + " ($M per MW-yr, stated MW).", size=6.2, color=GREY, after=1)
para("No $/MW computable (MW, term or value undisclosed): " + nd_list + ".", size=6.0, color=GREY, after=1)
para("Tags: R reported, P press-reported, D derived, E estimated; '(up to)' = maximum TCV. Microsoft-Nebius: no MW stated by either company; shown only as the house "
     f"${CON['N1'][22]:.2f}M per IT MW-yr (E; $17.4B / 5 yrs / ~240 IT MW est., model Contracts!E12). Nebius Meta #2 and backstop tie no MW (n/d); backstop is contingent. Nebius MW is a Capacity estimate throughout. "
     f"Share count: {fd_debt:.2f}M debt convention at the ${refpx:.2f} 10/7 close (if-converted {fd_conv:.2f}M above ${px34:.2f}, all convert debt out of EV; {fd_mid:.2f}M between ${px30:.2f} and ${px34:.2f}); never mix share and debt sides. "
     "Nebius Q3'26 results ~Nov 10 (est.), not announced. " + BAND_NOTE, size=6.0, color=GREY, after=0)

doc.save("/workspace/nebius/peers/memo/nebius_vs_peers_unit_economics.docx")
print("saved")
