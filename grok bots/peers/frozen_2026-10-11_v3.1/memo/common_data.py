# Shared loader: every number in the memo / summary table is read from peers/comparison.xlsx
# (plus the Microsoft gross->IT read from peers/price_or_mw.xlsx Nebius_Check). No hard-coded KPI values.
import openpyxl, re
CMP = "/workspace/nebius/peers/comparison.xlsx"
wb = openpyxl.load_workbook(CMP, data_only=True)
_c = wb["Comparison"]
HDR = [c.value for c in _c[3]]
ROWS = {r[0].value: [c.value for c in r] for r in _c.iter_rows(min_row=4) if r[0].value and str(r[0].value).startswith("K")}
PEERS = ["Nebius","CoreWeave","IREN","Oracle OCI","Nscale","Lambda","Crusoe"]
SHORT = {"Nebius":"Nebius","CoreWeave":"CoreWeave","IREN":"IREN","Oracle OCI":"Oracle OCI","Nscale":"Nscale","Lambda":"Lambda","Crusoe":"Crusoe"}
COL = {p: HDR.index(p) for p in PEERS}
MED = HDR.index("Peer median (formula; comparable peers)")
FLAG = HDR.index("Nebius better / worse")
PCOM = HDR.index("Peer basis comment")
RANK = {r[1]: {"rank": r[0], "flag": r[3], "neb": r[4], "med": r[5], "gap": r[6]}
        for r in wb["Ranking"].iter_rows(min_row=3, values_only=True) if r[1]}
NI = {r[0]: r for r in wb["Nebius_Inputs"].iter_rows(min_row=3, values_only=True) if r[0]}
CON = {r[0]: r for r in wb["Contracts_$MWyr"].iter_rows(min_row=3, values_only=True) if r[0]}
CON_HDR = [c.value for c in wb["Contracts_$MWyr"][2]]
SITES = [r for r in wb["Sites"].iter_rows(values_only=True)]
_pom = openpyxl.load_workbook("/workspace/nebius/peers/price_or_mw.xlsx", data_only=True)["Nebius_Check"]
MSFT_GROSS_READ = next(r[1] for r in _pom.iter_rows(values_only=True) if r[0] and str(r[0]).startswith("\u2192 $M per IT MW-yr [E, house] = model Contracts!E12"))  # 2026-10-09 N1 ruling: house $14.49M (unused in memo outputs)

def raw(k, p):
    x = ROWS[k][COL[p]]
    return x if isinstance(x, (int, float)) else None
def tag(k, p):
    t = ROWS[k][COL[p] + 1]
    return t
def asof(k, p):
    return ROWS[k][COL[p] + 2]
def med(k):
    return RANK[k]["med"] if k in RANK else ROWS[k][MED]
def flag(k):
    return RANK[k]["flag"] if k in RANK else ROWS[k][FLAG]
def excluded(k, p):
    s = ROWS[k][PCOM] or ""
    return bool(re.search(r"Excluded from median:.*\b" + re.escape(p.split()[0]) + r"\b", s))

TAGMAP = {"Reported":"R","Derived":"D","Estimated":"E","R":"R","D":"D","E":"E","P":"P",
          "D (inputs P)":"D/P","E (inputs P/R)":"E/P","E (inputs P)":"E/P"}
def stag(t):
    if t is None: return ""
    return TAGMAP.get(t, t)
def f(k, p, d=None):
    """value only, formatted"""
    x = raw(k, p)
    if x is None:
        v = ROWS[k][COL[p]]
        return "n/d" if v in (None, "n/d") else str(v)
    if d is None: d = DEC.get(k, 2)
    return f"{x:,.{d}f}"
def ft(k, p, d=None):
    """value + short tag"""
    s = f(k, p, d)
    if s in ("n/d", "n.m."): return s
    return f"{s} ({stag(tag(k,p))})"
def fm(k, d=None):
    if d is None: d = DEC.get(k, 2)
    m = med(k)
    return f"{m:,.{d}f}" if isinstance(m, (int, float)) else "n/d"
def ni(key): return NI[key][2]
DEC = {"K01":2,"K02":2,"K02b":2,"K05":2,"K06":1,"K07":2,"K09":0,"K12":1,"K13":1,"K14":1,"K15":1,"K16":1,"K17":1,"K18":2,
       "K22":0,"K23":1,"K24":0,"K26":0,"K27":0,"K28":1,"K36":1,"K39":2,"K40":0,"K43":1,"K44":2,"K04":1,"K29":2}

# One-page summary table spec: (row label, KPI id, unit, group)
SUMMARY = [
    ("Revenue per avg BILLED IT MW (K02b)\u00b9", "K02b", "$M/MW-yr", "Revenue per MW"),
    ("Revenue per avg active MW (K02)", "K02", "$M/MW-yr", "Revenue per MW"),
    ("T4Q cash capex / T4Q new active MW (K13)", "K13", "$M/MW", "Capex per MW"),
    ("In-service additions / new MW, CRWV def. (K12)", "K12", "$M/MW", "Capex per MW"),
    ("Useful life, tech equipment (K16)", "K16", "years", "Depreciation life"),
    ("Implied depreciable life = 1/yield (K18)", "K18", "years", "Depreciation life"),
    ("Weighted eff. cost of debt (K28)", "K28", "%", "Cost of funding"),
    ("Implied cash cost of debt (K36)", "K36", "%", "Cost of funding"),
    ("RPO / annualized revenue (K23)", "K23", "years", "Backlog coverage"),
    ("RPO % recognized within 24m (K22)", "K22", "%", "Backlog coverage"),
    ("Largest customer % of revenue (K24)", "K24", "%", "Customer concentration"),
]
SUM_DEC = {"K16":1,"K18":2,"K13":1,"K12":1,"K28":1,"K36":1,"K23":1,"K22":0,"K24":0,"K02":2,"K02b":2}

# Contract rows for the page-3 table: compute contracts with a computable $/MW-yr, plus all Nebius rows.
def contract_rows():
    out, nd, fac = [], [], []
    for cid, r in CON.items():
        if cid is None or not isinstance(r[1], str) or r[1] not in ("Reused","New","Reused (KPI)"): continue
        cat = r[4] or ""
        rec = dict(id=cid, provider=r[2], customer=r[3], cat=cat, signed=r[5], tcv=r[6], tcvtag=r[7], term=r[8],
                   mw=r[9], mwbasis=r[10], itmw=r[11], itmeth=r[12], stated=r[21], it=r[22], gpuhr=r[23], src=r[24], notes=r[25])
        if cat.startswith("Facility lease"):
            fac.append(rec); continue
        isnum = lambda x: isinstance(x, (int, float))
        if isnum(rec["stated"]) or isnum(rec["it"]) or rec["provider"] == "Nebius":
            out.append(rec)
        else:
            nd.append(rec)
    return out, nd, fac

# ---- Comp band (2026-10-08 ET coordinator correction): compute-contract $/MW-yr band over the memo comparison set.
BAND_IDS = ["P1", "NEW1", "P2", "O1", "NEW2", "O2"]
BAND = {cid: (CON[cid][22] if isinstance(CON[cid][22], (int, float)) else CON[cid][21]) for cid in BAND_IDS}
BAND_LO_ID, BAND_HI_ID = min(BAND, key=BAND.get), max(BAND, key=BAND.get)
assert BAND_LO_ID == "P1" and BAND_HI_ID == "NEW2" and abs(BAND["NEW2"] - 16.67) < 0.01 and abs(BAND["P1"] - 9.7) < 0.05, ("comp band drift", BAND)
BAND_TERM = CON["NEW2"][8]; assert BAND_TERM == 6
BAND_CONF = "Medium"
BAND_NOTE = (f"Comp band ${BAND[BAND_LO_ID]:.1f}-{BAND[BAND_HI_ID]:.1f}M per MW-yr: low IREN-Microsoft ${BAND['P1']:.1f}M (R, IT MW); top Lambda-Anthropic ~${BAND['NEW2']:.1f}M "
             f"($35B / {BAND_TERM:.0f}-yr term / ~350 MW; press inputs, IT vs gross MW unstated; {BAND_CONF} confidence).")

# ---- Oct 8 lock refresh (~10:00 ET: model v0.9e = source of truth): share-count convention + FY27 billed MW (read live) ----
sys_path_added = __import__("sys").path.insert(0, "/workspace/nebius/peers")
import nebius_locks as _L
FDC = _L.fd_conventions()          # thresholds / share counts from filings/fd_share_bridge.md
assert abs(FDC["FD_DEBT"]/1e6 - NI["N_FD_DEBT"][2]) < 1e-6 and abs(FDC["FD_IFCONV"]/1e6 - NI["N_FD_IFCONV"][2]) < 1e-6, "FD mismatch workbook vs fd_share_bridge"
ML = _L.model_locks()              # model/nebius_model.xlsx v0.9e (Drivers / Judge_Slide / Inputs)
for _k, _n in (("FY27_BASE", "N_BILL_FY27_BASE"), ("FY27_HIGH", "N_BILL_FY27_HIGH"), ("MSFT_REV", "N_MSFT_FY27_REV"), ("VIN_MW", "N_MSFT_FY27_MW"),
               ("YE26_BASE", "N_BILL_YE26_BASE"), ("YE26_HIGH", "N_BILL_YE26_HIGH"), ("FD_DEBT", "N_FD_DEBT"), ("FD_IFCONV", "N_FD_IFCONV")):
    assert abs(ML[_k]["value"] - NI[_n][2]) < 1e-6, ("comparison.xlsx not on model v0.9e - rebuild", _k, ML[_k]["value"], NI[_n][2])
assert abs(ML["HURDLE_XMSFT"]["value"] - NI["D_STREET_FY27_NONMSFT"][2]) < 1e-6 and abs(ML["BLENDED"]["value"] - NI["D_STREET_FY27_BASE"][2]) < 1e-6, "hurdle mismatch"
