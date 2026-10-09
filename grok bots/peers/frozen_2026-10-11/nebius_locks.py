# Shared read-only loaders for Nebius locks used by peers/ builds (Nebius Peer KPIs).
# 2026-10-08 ~10:00 ET (v0.9e): the TEAM MODEL is now the source of truth for the billed-MW locks and the Microsoft carve-out.
#   model_locks() reads model/nebius_model.xlsx (Cover A2 must say v0.9e): Drivers B34/C34 FY27 billed MW (engine-linked), B32 Microsoft
#   (Vineland) FY27 revenue (Contracts MS_Rev), B33 Vineland billed MW-yr (= B32 / $14.5M), B30/C30 non-Vineland MW, B9 ex-MSFT hurdle,
#   B10 blended, B37-B39 $16M coverage; Judge_Slide B33/C33 YE26 guide, B34/B35 YE26 billed MW, B72 Street FY27; Inputs r473-475 FD;
#   Drivers horizon table r59-65 PTs. Rows are located by LABEL and every value carries the cell it came from.
# capacity_locks() (capacity/final_billed_vs_connected.xlsx) is kept for HISTORY/AUDIT only (Oct 8 AM tranche locks 536.48/663.8,
#   superseded by v0.9e); no build uses it as an input any more.
# fd_conventions() still parses filings/fd_share_bridge.md (thresholds + the mid tier); builds cross-check it against the model's FD.
# Never writes to model/, capacity/ or filings/.
import re, openpyxl
W = "/workspace/nebius"
CAPX = f"{W}/capacity/final_billed_vs_connected.xlsx"
FDMD = f"{W}/filings/fd_share_bridge.md"
MODEL = f"{W}/model/nebius_model.xlsx"

def _find(ws, col, pred):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and pred(v.strip()):
            return r
    raise KeyError(f"{ws.title}: label not found")

def capacity_locks():
    wf = openpyxl.load_workbook(CAPX); wv = openpyxl.load_workbook(CAPX, data_only=True)
    ft, fv = wf["Final_Table"], wv["Final_Table"]; qt, qv = wf["Quarterly_IT"], wv["Quarterly_IT"]
    out = {}
    for key, lab in [("FY27_BASE", "FY27 avg billed IT BASE"), ("FY27_HIGH", "FY27 avg billed IT HIGH"),
                     ("YE26_BASE", "YE26 billed IT BASE"), ("YE26_HIGH", "YE26 billed IT HIGH")]:
        r = _find(ft, 1, lambda s, lab=lab: s == lab)
        locked, sitesum = fv.cell(r, 2).value, fv.cell(r, 4).value
        assert isinstance(locked, (int, float)) and isinstance(sitesum, (int, float)) and abs(locked - sitesum) < 0.5, (key, locked, sitesum)
        out[key] = dict(value=float(locked), sitesum=float(sitesum), cell=f"Final_Table!B{r}", sumcell=f"Final_Table!D{r}",
                        tag=str(ft.cell(r, 3).value), src=str(ft.cell(r, 7).value))
    r = _find(ft, 1, lambda s: s.startswith("MEMO") and "SUPERSEDED" in s and "BASE 474.96" in s)
    out["FY27_BASE_SUPERSEDED"] = dict(value=float(fv.cell(r, 4).value), cell=f"Final_Table!D{r}")
    try:
        r = _find(ft, 1, lambda s: s.startswith("MEMO") and "SUPERSEDED" in s and "HIGH 645.4" in s)
        out["FY27_HIGH_SUPERSEDED"] = dict(value=float(fv.cell(r, 4).value), cell=f"Final_Table!D{r}")
    except KeyError:
        out["FY27_HIGH_SUPERSEDED"] = None
    q = {}
    for key, pred in [("T14_MW", lambda s: s.startswith("T1–T4 IT MW billed full year 2027")),
                      ("T59_REV", lambda s: s.startswith("T5–T9 2027 revenue")),
                      ("MSFT_RATE", lambda s: s.startswith("Microsoft rate $M per IT MW-yr")),
                      ("VIN_MW", lambda s: s.startswith("Vineland FY27 avg billed IT MW (tranche basis)"))]:
        r = _find(qt, 1, pred)
        q[key] = dict(value=float(qv.cell(r, 2).value), cell=f"Quarterly_IT!B{r}", label=str(qt.cell(r, 1).value))
    q["MSFT_REV"] = q["T14_MW"]["value"] * q["MSFT_RATE"]["value"] + q["T59_REV"]["value"]
    out["VINELAND"] = q
    return out

def _row(ws, pred, col=1):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and pred(v.strip()):
            return r
    raise KeyError(f"{ws.title}: label not found")

def model_locks():
    """v0.9e locks from the team model (cached values; the model is saved recalculated). Returns {key: dict(value, cell, label)}."""
    wf = openpyxl.load_workbook(MODEL); wv = openpyxl.load_workbook(MODEL, data_only=True)
    ver = str(wv["Cover"]["A2"].value)
    assert ver.startswith("VERSION v0.9e"), ver[:60]
    dv, df, jv, iv = wv["Drivers"], wf["Drivers"], wv["Judge_Slide"], wv["Inputs"]
    out = {"VERSION": dict(value="v0.9e", cell="Cover!A2", label=ver[:120])}
    def put(key, sheet, ws, r, col, lab=None):
        v = ws.cell(r, col).value
        assert isinstance(v, (int, float)), (key, sheet, r, col, v)
        out[key] = dict(value=float(v), cell=f"{sheet}!{openpyxl.utils.get_column_letter(col)}{r}", label=str(lab or ws.cell(r, 1).value)[:160])
    r = _row(dv, lambda s: s.startswith("FY27 BILLED MW - ENGINE-LINKED")); put("FY27_BASE", "Drivers", dv, r, 2); put("FY27_HIGH", "Drivers", dv, r, 3)
    r = _row(dv, lambda s: s.startswith("Microsoft (Vineland) FY27 revenue in the engine")); put("MSFT_REV", "Drivers", dv, r, 2)
    r = _row(dv, lambda s: s.startswith("Vineland billed MW-yr on the revenue basis")); put("VIN_MW", "Drivers", dv, r, 2)
    r = _row(dv, lambda s: s.startswith("less Vineland inside that path")); put("NONVIN_BASE", "Drivers", dv, r, 2); put("NONVIN_HIGH", "Drivers", dv, r, 3)
    r = _row(dv, lambda s: s.startswith("NON-MICROSOFT hurdle")); put("HURDLE_XMSFT", "Drivers", dv, r, 2)
    r = _row(dv, lambda s: s.startswith("Blended STREET")); put("BLENDED", "Drivers", dv, r, 2)
    r = _row(dv, lambda s: s.startswith("MW NEEDED at $16M")); put("MW_NEEDED_16", "Drivers", dv, r, 2)
    r = _row(dv, lambda s: s.startswith("multiple needed / engine-linked")); put("COV16_BASE", "Drivers", dv, r, 2); put("COV16_HIGH", "Drivers", dv, r, 3)
    r = _row(dv, lambda s: s.startswith("memo: MS billing MW avg FY27")); put("MS_BILLMW_ALT", "Drivers", dv, r, 2)
    r = _row(dv, lambda s: s.startswith("memo: Capacity locked pair")); put("CAP_FY27_BASE", "Drivers", dv, r, 2); put("CAP_FY27_HIGH", "Drivers", dv, r, 3)
    r = _row(jv, lambda s: s.startswith("YE26 annualized run-rate guide")); put("GUIDE_LO", "Judge_Slide", jv, r, 2); put("GUIDE_HI", "Judge_Slide", jv, r, 3)
    r = _row(jv, lambda s: s.startswith("YE26 billing IT MW - Capacity BASE")); put("YE26_BASE", "Judge_Slide", jv, r, 2)
    r = _row(jv, lambda s: s.startswith("YE26 billing IT MW - Capacity HIGH")); put("YE26_HIGH", "Judge_Slide", jv, r, 2)
    r = _row(jv, lambda s: s.startswith("Street FY27 revenue ($M)")); put("STREET_FY27", "Judge_Slide", jv, r, 2)
    r = _row(iv, lambda s: s.startswith("FD shares - new notes as DEBT")); put("FD_DEBT", "Inputs", iv, r, 6)
    r = _row(iv, lambda s: s.startswith("FD shares - new notes IF-CONVERTED")); put("FD_IFCONV", "Inputs", iv, r, 6)
    r = _row(iv, lambda s: s.startswith("FD shares used by the selected treatment")); put("FD_USED", "Inputs", iv, r, 6)
    # PT horizon table (12m PT = column D); keep the scenario label so a $0 is never shown without its reason
    hr = _row(dv, lambda s: s == "Scenario"); pts = {}
    for rr in range(hr + 1, hr + 12):
        lab = dv.cell(rr, 1).value
        if not isinstance(lab, str) or not isinstance(dv.cell(rr, 4).value, (int, float)): continue
        pts[lab] = dict(pt12=float(dv.cell(rr, 4).value), ev=float(dv.cell(rr, 2).value), treat=dv.cell(rr, 7).value, fd=dv.cell(rr, 8).value,
                        need=dv.cell(rr, 9).value, cell=f"Drivers!D{rr}")
    out["PT"] = pts
    # sanity: the model's own formulas must close
    m = out
    assert abs(m["FY27_BASE"]["value"] - (m["NONVIN_BASE"]["value"] + m["VIN_MW"]["value"])) < 0.01
    assert abs(m["HURDLE_XMSFT"]["value"] - (m["STREET_FY27"]["value"] - m["MSFT_REV"]["value"]) / m["NONVIN_BASE"]["value"]) < 1e-6
    assert abs(m["BLENDED"]["value"] - m["STREET_FY27"]["value"] / m["FY27_BASE"]["value"]) < 1e-6
    return out

def fd_conventions():
    lines = open(FDMD, encoding="utf-8").read().splitlines()
    def grab(pred, rx):
        i = next(i for i, l in enumerate(lines, 1) if pred(l))
        m = re.search(rx, lines[i - 1]); assert m, (i, lines[i - 1])
        return i, m
    i1, m1 = grab(lambda l: l.startswith("| **FD1**"), r"\*\*([\d,]+)\*\*\s*\|\s*\*\*D")
    i0, m0 = grab(lambda l: l.startswith("| **FD0**"), r"\*\*([\d,]+)\*\*\s*\|\s*\*\*D")
    ia, ma = grab(lambda l: "Alternative convention" in l, r"FD = \*\*([\d,]+)\*\*")
    i9, m9 = grab(lambda l: l.startswith("| 9 |") and "Feb-2030" in l, r"\|\s*([\d,]{6,})\s*\|.*\(\$([\d.]+)\)")
    i10, m10 = grab(lambda l: l.startswith("| 10 |") and "Feb-2034" in l, r"\|\s*([\d,]{6,})\s*\|.*\(\$([\d.]+)\)")
    ip, mp = grab(lambda l: l.startswith("**Moneyness at $"), r"Moneyness at \$([\d.]+)\*\* \(([^)]*)\)")
    n = lambda s: int(s.replace(",", ""))
    d = dict(FD_IFCONV=n(m1.group(1)), FD_IFCONV_LINE=i1, FD_630=n(m0.group(1)), FD_630_LINE=i0,
             FD_DEBT=n(ma.group(1)), FD_DEBT_LINE=ia,
             SH_2030=n(m9.group(1)), PX_2030=float(m9.group(2)), LINE_2030=i9,
             SH_2034=n(m10.group(1)), PX_2034=float(m10.group(2)), LINE_2034=i10,
             REF_PX=float(mp.group(1)), REF_PX_NOTE=mp.group(2), REF_PX_LINE=ip)
    d["FD_MID"] = d["FD_DEBT"] + d["SH_2030"]                     # 2030s in the money, 2034s out
    assert d["FD_MID"] + d["SH_2034"] == d["FD_IFCONV"], d       # bridge must close
    return d

if __name__ == "__main__":
    import json
    print(json.dumps(model_locks(), indent=1, default=str)); print(fd_conventions())
