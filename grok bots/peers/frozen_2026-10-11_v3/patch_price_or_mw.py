#!/usr/bin/env python3
# 2026-10-08 ET (Nebius Peer KPIs): reproducible patch of peers/price_or_mw.xlsx onto the Nebius FY27 billed-MW locks.
# Source of truth for the Oct-5 content is price_or_mw_pre_oct8_backup.xlsx (build_price_or_mw.py is stale and guarded: it no longer
# runs and predates the 10-05 hand corrections), so this script always starts from that backup -> idempotent.
# ~10:00 ET v0.9e: locks now come from the TEAM MODEL v0.9e (peers/nebius_locks.py model_locks(): Drivers B34/C34 FY27 billed MW
# 531.84 / 659.16, B32 Microsoft FY27 revenue $2,643.8M, B33 Vineland billed MW-yr 182.33; Judge_Slide B34/B35 YE26). Replaces the
# 08:44 Capacity tranche patch (536.48 / 663.8, $2,711M on 186.97 MW; backup price_or_mw_pre_v09e_backup.xlsx). Superseded memo row 37 dropped.
# Run: python3 patch_price_or_mw.py && python3 ../model/recalc.py $PWD/price_or_mw.xlsx && cp /tmp/rc/price_or_mw.xlsx . && python3 build_price_or_mw_chart.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import nebius_locks as NL
P = "/workspace/nebius/peers"
SRC, OUT = f"{P}/price_or_mw_pre_oct8_backup.xlsx", f"{P}/price_or_mw.xlsx"
BOLD = Font(bold=True); BLUE = Font(color="0000CC"); KEY = PatternFill("solid", fgColor="FFF2CC"); HDR = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")
ML = NL.model_locks(); V = lambda k: ML[k]["value"]; C = lambda k: ML[k]["cell"]
B, H, VIN, MS = V("FY27_BASE"), V("FY27_HIGH"), V("VIN_MW"), V("MSFT_REV")
assert abs(B - 531.8) < 0.1 and abs(H - 659.2) < 0.1 and abs(MS - 2643.8) < 0.05 and abs(V("YE26_BASE") - 257.7) < 0.05 and abs(V("YE26_HIGH") - 284.2) < 0.05, ML
assert abs(V("STREET_FY27") - 12300) < 0.5
MOD = "model/nebius_model.xlsx"
BAND_NOTE = "CONTRACT COMP BAND (separate metric: TCV / term / MW, not revenue per billed MW): $9.7-16.7M per MW-yr - low IREN-Microsoft $9.7M; top Lambda-Anthropic ~$16.7M ($35B / 6-yr / ~350 MW; press inputs, IT vs gross unstated; Medium confidence) - comparison.xlsx Contracts_$MWyr NEW2 (2026-10-08; was $9.7-16.2M, Nscale-Anthropic top)."
wb = openpyxl.load_workbook(SRC)
mp, ip, nc, rs, rd = wb["MW_Paths"], wb["Implied_Price"], wb["Nebius_Check"], wb["Revenue_Sources"], wb["README"]
assert mp.max_row == 28 and rs.max_row == 18 and ip["A5"].value.startswith("Nebius STREET") and ip["A18"].value is None and ip["A24"].value is None and ip["A32"].value is None
assert mp["E9"].value == 257.7 or str(mp["E9"].value).startswith("="), mp["E9"].value

# ---- MW_Paths r30-36: model v0.9e locks + Microsoft carve-out
mp.cell(30, 1, "NEBIUS FY27 avg billed IT MW - MODEL v0.9e (engine-linked: non-Vineland path + Vineland billed MW-yr = engine Microsoft revenue / $14.5M)").font = BOLD
rows = [
 (31, "FY27 avg billed IT MW - BASE (model v0.9e) - STREET DENOMINATOR", B, "D", f"{MOD} {C('FY27_BASE')} = {C('NONVIN_BASE')} {V('NONVIN_BASE'):.2f} + {C('VIN_MW')} {VIN:.2f}"),
 (32, "FY27 avg billed IT MW - HIGH (model v0.9e)", H, "D", f"{MOD} {C('FY27_HIGH')} = {C('NONVIN_HIGH')} {V('NONVIN_HIGH'):.2f} + {VIN:.2f}"),
 (33, "Vineland / Microsoft FY27 billed MW-yr (revenue basis = Microsoft revenue / $14.5M)", VIN, "D", f"{MOD} {C('VIN_MW')} (alt Contracts row 141 billing-MW avg {V('MS_BILLMW_ALT'):.1f}, memo)"),
 (34, "Microsoft (Vineland) FY27 revenue $M in the model engine (Contracts MS_Rev)", MS, "D", f"{MOD} {C('MSFT_REV')} = SUMIF(Q_Yr,2027,MS_Rev), Contracts row 145"),
 (35, "NON-Microsoft FY27 avg billed IT MW - BASE", "=J31-J33", "D", f"base less Vineland billed MW-yr (= {MOD} {C('NONVIN_BASE')})"),
 (36, "NON-Microsoft FY27 avg billed IT MW - HIGH", "=J32-J33", "D", f"high less Vineland (= {C('NONVIN_HIGH')}; Vineland identical on both paths)"),
]
for r, a, v, t, s in rows:
    mp.cell(r, 1, a); c = mp.cell(r, 10, v); c.number_format = "#,##0.00"
    if isinstance(v, float): c.font = BLUE
    mp.cell(r, 11, t); mp.cell(r, 12, s).alignment = WRAP
for r in (31, 35): mp.cell(r, 10).fill = KEY
mp["A6"] = str(mp["A6"].value).replace(" - STREET DENOMINATOR", " - HISTORY ONLY (v0.9d energized-lag path; not used - see row 31)")

# ---- Revenue_Sources r19-20
rs.cell(19, 1, "Nebius"); rs.cell(19, 2, "FY27 Microsoft (Vineland) revenue - model v0.9e engine"); rs.cell(19, 3, "=MW_Paths!J34"); rs.cell(19, 4, "D")
rs.cell(19, 5, "CY2027"); rs.cell(19, 6, MOD); rs.cell(19, 7, f"{C('MSFT_REV')} (Contracts MS_Rev row 145: T1-T4 + ramp + Q1'27 genset cap + T5-T9 Bloom starts)")
rs.cell(20, 1, "Nebius"); rs.cell(20, 2, "Street FY27 ex-Microsoft revenue"); rs.cell(20, 3, "=C4-C19"); rs.cell(20, 4, "E/D")
rs.cell(20, 5, "CY2027"); rs.cell(20, 6, "S1 less row 19"); rs.cell(20, 7, "Peer comparator numerator (peers have no $14.5M Microsoft anchor)")
for r in (19, 20): rs.cell(r, 3).number_format = "#,##0.0"

# ---- Nebius_Check
nonms = (V("STREET_FY27") - MS) / (B - VIN); blend = V("STREET_FY27") / B
nc["A1"] = f"Verify Nebius Street $12.3B / FY27 avg billed IT MW (model v0.9e {B:.1f}) -> ~${blend:.1f}M blended; NON-Microsoft ~${nonms:.1f}M (peer comparator); YE26 guide $25-32M. Align $14.5M vs $17.1M Microsoft readings."
nc["A5"] = "FY27 avg billed IT MW - BASE, model v0.9e (STREET DENOMINATOR)"; nc["B5"] = "=MW_Paths!J31"
nc["D5"] = f"{MOD} {C('FY27_BASE')} = {B:.2f} (engine-linked)"
nc["D6"] = f"Blended (incl. Microsoft $14.5M anchor) - memo; = {MOD} {C('BLENDED')} {V('BLENDED'):.4f}"
nc["A9"] = "model v0.9e Judge_Slide / Drivers printed Street blended $/MW (Drivers B10 = Judge_Slide B16)"; nc["B9"] = V("BLENDED"); nc["B9"].font = BLUE
nc["D9"] = f"{MOD} {C('BLENDED')}"
nc["A10"] = "Delta vs model v0.9e (should be ~0)"
nc.cell(32, 1, "NON-MICROSOFT Street $/billed IT MW (PEER COMPARATOR - peers carry no $14.5M Microsoft anchor)").font = BOLD
nchk = [
 (33, "Street FY27 revenue ($M)", "=B4", "E", f"S1 (= {MOD} {C('STREET_FY27')})"),
 (34, "less Microsoft (Vineland) FY27 revenue ($M)", "=MW_Paths!J34", "D", f"Revenue_Sources r19 = {C('MSFT_REV')}"),
 (35, "FY27 avg billed IT MW - BASE (model v0.9e)", "=MW_Paths!J31", "D", C("FY27_BASE")),
 (36, "less Vineland/Microsoft FY27 billed MW-yr", "=MW_Paths!J33", "D", C("VIN_MW")),
 (37, "NON-MSFT $M per billed IT MW - BASE (HEADLINE peer comparator)", "=(B33-B34)/(B35-B36)", "D", f"calc: (12,300 - {MS:,.1f}) / ({B:.2f} - {VIN:.2f}) = (12,300 - {MS:,.1f}) / {B-VIN:.2f}"),
 (38, "Grok Bot figure (Oct 8, v0.9e message)", 27.6, "memo", "for reconciliation only (rounded)"),
 (39, "Delta derived - Grok Bot (should be ~0)", "=B37-B38", "D", ""),
 (40, "Street / FY27 avg billed HIGH (model v0.9e) - blended", "=B33/MW_Paths!J32", "D", f"$12.3B / {H:.2f} (high path)"),
 (41, "NON-MSFT $M per billed IT MW - HIGH", "=(B33-B34)/MW_Paths!J36", "D", "Microsoft carve-out identical on the high path"),
 (42, "Model v0.9e printed non-MSFT hurdle (Drivers B9)", V("HURDLE_XMSFT"), "D", f"{MOD} {C('HURDLE_XMSFT')} (Judge_Slide B20)"),
 (43, "Delta derived - model (should be 0)", "=B37-B42", "D", ""),
 (44, "MW needed at $16M comps = Street / $16M", "=B33/16", "D", f"{MOD} {C('MW_NEEDED_16')}"),
 (45, "Coverage needed at $16M - BASE (x)", "=B44/B35", "D", f"{MOD} {C('COV16_BASE')} = {V('COV16_BASE'):.4f}"),
 (46, "Coverage needed at $16M - HIGH (x)", "=B44/MW_Paths!J32", "D", f"{MOD} {C('COV16_HIGH')} = {V('COV16_HIGH'):.4f}"),
]
for r, a, v, t, s in nchk:
    nc.cell(r, 1, a); c = nc.cell(r, 2, v); c.number_format = "#,##0.00"
    if isinstance(v, float): c.font = BLUE
    nc.cell(r, 3, t); nc.cell(r, 4, s).alignment = WRAP
nc["B37"].fill = KEY; nc["B37"].font = BOLD

# ---- Implied_Price
BM = "Nebius_Check!$B$37"     # peer benchmark = NON-MSFT Street $/billed MW
def setrow(r, a, rev, rt, mw, mt, note, key=False):
    ip.cell(r, 1, a); ip.cell(r, 2, rev).number_format = "#,##0"; ip.cell(r, 3, rt); ip.cell(r, 4, mw).number_format = "#,##0.0"; ip.cell(r, 5, mt)
    ip.cell(r, 6, f"=B{r}/D{r}").number_format = "0.00"; ip.cell(r, 10, note).alignment = WRAP
    ip.cell(r, 9, f'=IF(AND(F{r}>=13,F{r}<=18),"IN $13-18 band",IF(F{r}>18,"ABOVE (rich like/above Nebius Street)","BELOW $13-18"))')
    for c in range(1, 10): ip.cell(r, c).fill = KEY if key else PatternFill()
setrow(5, "Nebius STREET blended (consensus / FY27 avg billed, model v0.9e)", "=Revenue_Sources!C4", "E", "=MW_Paths!J31", "D",
       f"MEMO (blended, incl. Microsoft $14.5M anchor): $12.3B / {B:.2f} ({MOD} {C('FY27_BASE')}) = ~${blend:.1f}M. Use row 18 (ex-MSFT) against peers.")
setrow(18, "Nebius STREET NON-MICROSOFT (peer comparator)", "=Revenue_Sources!C20", "E/D", "=MW_Paths!J35", "D",
       f"THE PEER BENCHMARK (model v0.9e): ($12.3B - ${MS/1000:.4f}B Microsoft) / ({B:.1f} - {VIN:.1f} Vineland billed MW-yr) = ~${nonms:.1f}M (= Drivers B9). Peers' $/MW has no $14.5M Microsoft anchor.", key=True)
setrow(24, "Nebius STREET blended / FY27 avg billed HIGH (model v0.9e)", "=Revenue_Sources!C4", "E", "=MW_Paths!J32", "D",
       f"High path: $12.3B / {H:.2f} ({MOD} {C('FY27_HIGH')}). Ex-MSFT high = Nebius_Check B41.")
for r in list(range(5, 19)) + [24]:
    if ip.cell(r, 6).value:
        ip.cell(r, 7, f"=F{r}-{BM}").number_format = "0.00"; ip.cell(r, 8, f"=F{r}/{BM}-1").number_format = "0.0%"
ip["G4"] = f"vs Nebius Street ex-MSFT (~${nonms:.1f}M)"
ip["A1"] = f"PRICE-OR-MW TEST: FY27 revenue (or run-rate guide) ÷ average billed IT MW → implied $M per billed IT MW-yr. Compare peers to Nebius Street EX-MICROSOFT ~${nonms:.1f}M (blended Street ~${blend:.1f}M on model v0.9e {B:.1f} MW is memo)."
ip["F26"] = f"vs Nebius ex-MSFT ~${nonms:.1f}M"
ip["A27"] = "Nebius Street (blended memo)"
for r in range(27, 32): ip.cell(r, 6, f"=D{r}-{BM}").number_format = "+0.0;-0.0"
ip.cell(32, 1, "Nebius Street ex-MSFT (PEER BENCHMARK)"); ip.cell(32, 2, "=B18/1000").number_format = "0.00"; ip.cell(32, 3, "=D18").number_format = "#,##0"
ip.cell(32, 4, "=F18").number_format = "0.0"; ip.cell(32, 5, "E/D / D"); ip.cell(32, 6, f"=D32-{BM}").number_format = "+0.0;-0.0"
for c in range(1, 7): ip.cell(32, c).fill = KEY
ip["A34"] = (f"Peers come out in the ~$11-16M per billed IT MW band on base paths. The like-for-like Nebius figure is Street EX-MICROSOFT ~${nonms:.1f}M "
    f"(($12.3B - ${MS/1000:.2f}B) / ({B:.1f} - {VIN:.1f}) billed IT MW, model v0.9e) - ~{nonms/16:.1f}-{nonms/11.3:.1f}x the peer band. "
    f"Blended Street is ~${blend:.1f}M (memo; pulled down by the $14.5M Microsoft anchor). "
    "CoreWeave Street / base MW ≈ $11-13M; IREN YE26 ARR/IT MW = $13.3M (R/R) and FY27 ARR-path proxy ≈ $14-18M; Oracle ≈ $14M (AI/non-AI and gross-MW caveats). "
    f"High path: $12.3B / {H:.1f} = ${12300/H:.1f}M blended, ex-MSFT ${(12300-MS)/(H-VIN):.1f}M. At $16M comps the Street needs {V('MW_NEEDED_16'):.0f} MW = {V('COV16_BASE'):.2f}x base / {V('COV16_HIGH'):.2f}x high. "
    "YE26 Nebius guide ($7-9B / ~258-284 billing MW) implies $25-35M. Confidence MEDIUM: CoreWeave YE27 MW Estimated; IREN lacks AI GAAP consensus; Oracle perimeter impure; Vineland MW per tranche is E. "
    + BAND_NOTE + " The $13-18 band in col I is this test's slide hypothesis band for revenue per billed MW, not the contract band.")
# ---- Oct-5 text that still described the ~475 energized-lag denominator (v0.9e: relabel; row 6 stays as clearly labelled history)
src = wb["Sources"]
rd["A17"] = ("CORRECTION / BASIS: All Nebius $/MW in this pack are on BILLED IT MW (K02b basis). YE26 guide: low=$7B/MW_Paths!E9(257.7); high=$9B/MW_Paths!E10(284.2 Judge_Slide B35 / Capacity HIGH billing). "
             f"Street FY27 uses model v0.9e FY27 avg billed IT MW {B:.1f} (Drivers B34; MW_Paths row 31), not energized.")
nc["A30"] = (str(nc["A30"].value).split(" Street FY27 at $26/MW")[0] + f" Street FY27 is ~${blend:.1f}M blended / ~${nonms:.1f}M ex-Microsoft on BILLED IT MW (model v0.9e {B:.1f}), not energized.")
src["B3"] = (f"Nebius FY27 avg billed IT MW {B:.2f} = model v0.9e Drivers B34 (engine-linked: non-Vineland path {V('NONVIN_BASE'):.1f} + Vineland billed MW-yr {VIN:.1f}). "
             "Capacity 'Billing MW (IT) - base' quarter-end FY27 avg = 593.9 (memo). History: the Oct-5 pack used the model energized-lag path (MW_Paths row 6).")
src["C3"] = "S2 model/nebius_model.xlsx Drivers rows 29-39 / Judge_Slide rows 16, 25-29; capacity/nebius_capacity.xlsx Quarterly MW rows 136/147"
src["D15"] = f"model/nebius_model.xlsx Judge_Slide B35 (v0.9e); FY27 high path {H:.1f} = Drivers C34"
mp["L6"] = "HISTORY ONLY (v0.9d model KPIs row BILLING, energized-lag); not used since Oct 8; Q4'26 open=212.74"
# README
rd["A1"] = "price_or_mw.xlsx - Nebius Peer KPIs - PRICE-OR-MW TEST - built 2026-10-05 ET; CORRECTED 2026-10-05 (YE26 high ÷ 284.2); RE-BASED 2026-10-08 ~10:00 ET to model v0.9e locks (patch_price_or_mw.py). Recalc: 0 errors."
rd["A5"] = (f"NEBIUS BENCHMARK (model v0.9e): Street FY27 $12.3B (S1) / FY27 avg billed IT MW {B:.2f} ({MOD} {C('FY27_BASE')}) = ~${blend:.1f}M blended (memo). "
            f"PEER COMPARATOR = ex-Microsoft ~${nonms:.1f}M (($12.3B - ${MS:,.1f}M) / ({B:.2f} - {VIN:.2f})). High path {H:.2f}: ${12300/H:.1f}M blended. "
            "YE26 ARR guide $7-9B / YE26 billing 258-284 MW = $25-35M.")
rd["A7"] = f"HEADLINE RESULT: Peers land in ~$11-16M on base paths - NOT as rich as Nebius Street ex-Microsoft (~${nonms:.1f}M). Confidence MEDIUM."
rd.cell(19, 1, f"OCT 8 ~10:00 ET RE-BASE TO MODEL v0.9e (source of truth): FY27 avg billed IT MW BASE {B:.2f} / HIGH {H:.2f} ({C('FY27_BASE')} / {C('FY27_HIGH')}); Microsoft FY27 revenue ${MS:,.1f}M ({C('MSFT_REV')}) "
               f"on Vineland {VIN:.2f} billed MW-yr ({C('VIN_MW')}); YE26 unchanged 257.7 / 284.2. Peer 'vs' columns benchmark ex-MSFT (Nebius_Check B37 = model Drivers B9). "
               "Oct-5 content preserved in price_or_mw_pre_oct8_backup.xlsx; build_price_or_mw.py is stale (guarded) - use patch_price_or_mw.py.").alignment = WRAP
rd.cell(20, 1, "HISTORY (superseded, not inputs): 08:44 ET patch used Capacity tranche locks 536.48 / 663.8 and Microsoft $2,711M on 186.97 MW (ex-MSFT 27.44); file kept as price_or_mw_pre_v09e_backup.xlsx.").alignment = WRAP
rd.cell(21, 1, "2026-10-08 PM: " + BAND_NOTE).alignment = WRAP
# 2026-10-08 ~23:40 ET energized-MW switch (coordinator): Nebius energized MW = current Capacity lock Quarterly MW!M130/N130 181.5 / 193
#   (was 171.5 / 173 from outdated nebius_capacity_v1.xlsx). Only the K02 cross-check row and its text move; billed-MW locks untouched.
import openpyxl as _ox
_q = _ox.load_workbook("/workspace/nebius/capacity/nebius_capacity.xlsx", data_only=True)["Quarterly MW"]
E1, E2 = _q["M130"].value, _q["N130"].value
assert (E1, E2) == (181.5, 193) and str(_q["B130"].value).startswith("Total incl. plug")
k02 = 574.9 * 4 / ((E1 + E2) / 2)
nc["B27"] = f"=574.9*4/AVERAGE({E1},{E2})"
nc["D27"] = f"Comparison K02 = {k02:.2f} E (energized {E1} / {E2}, Capacity base case Quarterly MW!M130/N130; was 13.35 on v1 171.5 / 173)"
nc["A30"] = str(nc["A30"].value).replace("K02 (energized) = 13.35;", f"K02 (energized) = {k02:.2f};")
assert f"{k02:.2f}" in nc["A30"].value
src["B14"] = str(src["B14"].value).replace("energized MW 171.5/173 (Capacity)", f"energized MW {E1}/{E2:.0f} (Capacity base case, E; Quarterly MW!M130/N130)")
assert "171.5" not in src["B14"].value
# 2026-10-09 ~11:12 ET N1 ruling (coordinator): Microsoft-Nebius is "$14.49M per IT MW [E, house] (Contracts!E12)" only.
#   Nebius_Check rows 16-26 rebuilt around model v0.9e Contracts!E12; 11.60 (300 MW site size) / 17.08 (GPU-derived) kept as HISTORY text only.
#   No lock cell (B5/B6/B9/B37/B40-B46, Implied_Price, MW_Paths) is touched.
_m = _ox.load_workbook("/workspace/nebius/model/nebius_model.xlsx", data_only=True)
E12, MSFT_MW, F111 = _m["Contracts"]["E12"].value, _m["Inputs"]["F216"].value, _m["Inputs"]["F111"].value
assert abs(E12 - 17392.9 / (5 * 240)) < 1e-6 and MSFT_MW == 240 and F111 == 14.5, (E12, MSFT_MW, F111)
for r in range(16, 27):
    for c in range(1, 5): nc.cell(r, c).value = None
n1 = [
 (16, "MICROSOFT-NEBIUS $ PER IT MW-YR (ruling Oct 9 2026): $14.49M per IT MW [E, house] (model Contracts!E12) only", None, None, None),
 (17, "Neither Nebius nor Microsoft states a contract MW; the current figure uses the house estimate of 240 IT MW (model Inputs!F216 MSFT_MW). 300 MW is Vineland's site size.", None, None, None),
 (18, "TCV committed ($M)", 17392.9, "R", "S12 (6-K 2025-09-08; = model Inputs!F215 MSFT_TCV)"),
 (19, "Term (years)", 5, "R", "6-K 2025-09-08"),
 (20, "ACV = TCV/term ($M/yr)", "=B18/B19", "D", None),
 (21, "IT MW - house estimate (model v0.9e Inputs!F216 MSFT_MW)", MSFT_MW, "E", "model/nebius_model.xlsx Inputs!F216; not stated by Nebius/Microsoft"),
 (22, "\u2192 $M per IT MW-yr [E, house] = model Contracts!E12 (CURRENT)", "=B20/B21", "E", f"${E12:.2f}M per IT MW [E, house]; agrees with the Drivers 14.5 input (model Inputs!F111)"),
 (23, "Model v0.9e Contracts!E12 printed value", E12, "D", "model/nebius_model.xlsx Contracts!E12 = MSFT_TCV / (5 x MSFT_MW)"),
 (24, "Delta derived - model (should be 0)", "=B22-B23", "D", None),
 (25, "HISTORY (superseded Oct 9 2026, not current values): 11.60 = ACV / 300 MW (Vineland site size, not a stated contract MW); 17.08 = ACV / GPU-derived 203.6 IT MW (100,000 GB300 / 491; retired). Do not quote.", None, None, None),
]
for r, a, v, t, d in n1:
    nc.cell(r, 1, a).alignment = WRAP
    if v is not None:
        c = nc.cell(r, 2, v); c.number_format = "#,##0.00"
        if isinstance(v, (int, float)) and r in (18, 19, 21, 23): c.font = BLUE
    if t: nc.cell(r, 3, t)
    if d: nc.cell(r, 4, d).alignment = WRAP
nc["A16"].font = BOLD; nc["B22"].fill = KEY; nc["B22"].font = BOLD
nc["A25"].font = Font(italic=True, color="808080")
nc["A1"] = nc["A1"].value.replace("Align $14.5M vs $17.1M Microsoft readings.", "Microsoft-Nebius = $14.49M per IT MW [E, house] (model Contracts!E12) only.")
_a30 = nc["A30"].value; _old = "quote Microsoft-Nebius as '$11.6 on stated 300 MW | ~$14.5 on Capacity gross\u2192IT read | $17.1 on GPU-derived 203.7 IT MW'."
assert _old in _a30, _a30
nc["A30"] = _a30.replace(_old, "quote Microsoft-Nebius only as '$14.49M per IT MW [E, house] (Contracts!E12)' (11.6 / 17.1 superseded, history only).")
rd["A12"] = "Microsoft-Nebius: $14.49M per IT MW [E, house] = model Contracts!E12 only (ruling Oct 9 2026; Nebius_Check rows 16-24). Earlier 11.6 (300 MW site size) / 17.1 (GPU-derived) readings are superseded - history only (Nebius_Check row 25)."
src["B13"] = "Nebius-Microsoft ACV: $17.4B / 5 yr = $3.48B/yr; on the house ~240 IT MW estimate (model Inputs!F216) = $14.49M per IT MW-yr [E, house] = model Contracts!E12. HISTORY (superseded Oct 9 2026): 11.60 on the 300 MW site size and 17.08 on GPU-derived 203.7 IT MW are not current values."
wb.save(OUT); print("saved", OUT, f"blended {blend:.4f} nonMSFT {nonms:.4f} high {12300/H:.4f} nonMSFT_high {(12300-MS)/(H-VIN):.4f}")
