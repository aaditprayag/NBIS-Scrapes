import openpyxl, shutil, os, sys
sys.path.insert(0, '/workspace/nebius/model')
from recalc import recalc
from openpyxl.styles import Font, PatternFill
SRC = "/workspace/nebius/model/nebius_model.xlsx"
wb0 = openpyxl.load_workbook(SRC)
DN = wb0.defined_names
def loc(nm):
    sh, rg = DN[nm].attr_text.split("!"); return sh.strip("'"), rg.replace("$", "").split(":")[0]
inputs = wb0["Inputs"]
def inrow(nm):  # live cell is F{r}; base D{r}
    sh, c = loc(nm); return int(c[1:])
V07C = {"OffBSWt": 0.5, "T59Align": 0, "PTBasis": 2}
V07B = {**V07C, "BridgeFix": 0, "RubinMW": 18, "VinIdleOn": 0, "FP_Resize": 0, "TVROResOn": 0}
NEWR = {**V07B, "GB_Racks": 6.8, "CPM_Lease_VR": 41, "CPM_Own_VR": 56, "LeasePrepayOn": 0, "ND_Px": 22e6 / (491 * 8760), "PrepayMode": 1}
V07 = {**NEWR, "CapIntFix": 0, "OwnOpexOn": 0, "CPM_SP": 11.25, "ShellMW": 4, "PowerMW": 6, "GPU_Resid": 0, "CapV": 1, "RackOn": 0, "CVDebtBasis": 1}
V06 = {**V07, "RubinMode": 0, "STGBOnly": 0, "CPM_Lease_VR": 32, "CPM_Own_VR": 43, "LoanOn": 1, "FP_On": 0, "CVDebtBasis": 2, "AVBurnOn": 0}
V05 = {**V06, "MYPathMode": 0, "ND_GenOn": 0, "Rubin_Prem": 500 / 320, "CH_Lo": 1520, "TO_Lo": 86.5, "TO_Hi": 86.5, "AV_Hi": 175, "Other_Val": 0, "AVPt": 0.5}
V04 = {**V05, "STShare": 0.10, "STMode": 0, "PrepayByCust": 0, "PrepayDecayOn": 0, "DRTermOn": 0, "DRSchedMode": 0, "DR_AmortQ": 12}
NEW0 = {**V04, "NewDealOn": 0, "PrepayBase": 0, "GPUAnchorOn": 0}
OLD3 = {**NEW0, "GPULead": 3, "DCLead": 2, "PxPathMode": 0, "PxChg_B": -0.10, "PxChg_N": -0.05, "PxCycleOn": 0, "TGMode": 0, "Rubin_Prem": 1.5 / 1.22, "CLT1Rent": 2.25 * 1.25}
V01 = {**OLD3, "FixMgmt": 0, "TermBase": 0, "TVFloor": 0, "OM_MW": 0, "RepriceOn": 0, "RenewIdx": 13, "GPULead": 1}
BRIDGE = [("v0.1 (as reported)", V01),
          ("+ fleet ratio forecast-only / capex on driver MW (FixMgmt)", {**V01, "FixMgmt": 1}),
          ("+ terminal on Q4'30 exit incl. Q1'31 MW (include both)", {**V01, "FixMgmt": 1, "TermBase": 1}),
          ("+ TV run-off floor", {**V01, "FixMgmt": 1, "TermBase": 1, "TVFloor": 1}),
          ("+ owned-site O&M", {**OLD3, "GPULead": 1, "TermBase": 1, "RepriceOn": 0, "RenewIdx": 13}),
          ("+ H2'26 repricing (= v0.2)", {**OLD3, "GPULead": 1, "TermBase": 1}),
          ("+ exclude both Q4'30 spend & Q1'31 MW", {**OLD3, "GPULead": 1}),
          ("+ GPU capex lead 3 qtrs (= v0.3)", {**OLD3}),
          ("+ split capex lead: shell/power 3, GPUs 2 qtrs", {**OLD3, "GPULead": 2, "DCLead": 3}),
          ("+ price path: post-Rubin cut 22% + -6%/yr from year 2", {**NEW0, "TGMode": 0, "Rubin_Prem": 1.5 / 1.22, "CLT1Rent": 2.25 * 1.25}),
          ("+ Rubin premium at per-MW parity (1.56x GB300)", {**NEW0, "TGMode": 0, "CLT1Rent": 2.25 * 1.25}),
          ("+ CLT1 rent $2.2M/IT MW-yr", {**NEW0, "TGMode": 0}),
          ("+ terminal g from fundamentals", {**NEW0}),
          ("+ new-deal pricing: 2027+ MW at $22M multi-year + 10% short-term $45M", {**V04, "PrepayBase": 0, "GPUAnchorOn": 0}),
          ("+ prepayments on new-build capex only (38.5%)", {**V04, "GPUAnchorOn": 0}),
          ("+ GPU history anchor ~30k at 3/31/25 (= v0.4)", {**V04}),
          ("+ closing deferred revenue run-off deducted from EV (cash double-count fix)", {**V04, "DRTermOn": 1}),
          ("+ prepayments by customer: signed MSFT/Meta 0% new, new hyperscaler 30%", {**V04, "DRTermOn": 1, "PrepayByCust": 1}),
          ("+ prepayment decay: other deals 38.5% -> 22.5% for 2028+ signings", {**V05, "STShare": 0.10, "STMode": 0, "DRSchedMode": 0, "DR_AmortQ": 12}),
          ("+ DR release on Reported 6/30 split; new prepayments straight-line 4 yrs", {**V05, "STShare": 0.10, "STMode": 0}),
          ("+ short-term = 5% of TOTAL active MW at $45M x 85% occupancy (= v0.5)", {**V05}),
          ("+ new-deal $/MW path fix: flat nominal within generation to 2028, then -3%/yr (no per-chip ageing)", {**V05, "MYPathMode": 1}),
          ("+ Rubin new-deal pricing 1.2x GB300 per GPU-hr (generation step-down -22% $/MW)", {k: v for k, v in V05.items() if k not in ("MYPathMode", "ND_GenOn", "Rubin_Prem")}),
          ("+ stake values per Filings Note 5; Avride $0 in base (= v0.6)", {**V06}),
          ("+ $775M asset-backed facility $0 drawn (undisclosed)", {**V06, "LoanOn": 0}),
          ("+ existing converts OTM at solved price as debt at ORIGINAL principal ($13.4B) not accreted", {**V06, "LoanOn": 0, "CVDebtBasis": 1}),
          ("+ Rubin revenue direct $18M/IT MW (no IT/gross toggle); short-term $45M tranche GB300-only", {**V06, "LoanOn": 0, "CVDebtBasis": 1, "RubinMode": 1, "STGBOnly": 1}),
          ("+ Rubin capex LEASED $41M (was $32M) - only ~50 IT MW leased Rubin in 2027-30 (AIB CLT1)", {**V07, "FP_On": 0, "AVBurnOn": 0, "CPM_Own_VR": 43}),
          ("+ Rubin capex OWNED $56M (was $43M) - ~2,048 IT MW owned Rubin 2027-30 (IT/GPU slice 31.75 -> 44.75, replaced every 5 yrs)", {**V07, "FP_On": 0, "AVBurnOn": 0}),
          ("+ Avride double-count rule (base: Avride $0, burn kept - no change)", {**V07, "FP_On": 0}),
          ("+ 2027 funding plan: ATM 12.27M sh + 2 x $6.4B converts; post-funding per-share (= v0.7)", {**V07}),
          ("FIX 5: capitalised interest out of UFCF tax / terminal D&A (financing no longer moves EV)", {**V07, "CapIntFix": 1}),
          ("FIX 2: owned-site opex $2.25M/IT MW-yr all-in, no rent (P&L O&M was $0.75M + power ~$0.66M)", {**V07, "CapIntFix": 1, "OwnOpexOn": 1}),
          ("FIX 3a: data hall $12.5M/IT MW (was 11.25) split 30% shell 20y / 70% power-cooling 10y; servers+network only on 5-yr cycle", {**V07, "CapIntFix": 1, "OwnOpexOn": 1, "CPM_SP": 12.5, "ShellMW": 3.75, "PowerMW": 8.75}),
          ("FIX 3b: retired GPUs 15% residual at refresh (after tax), forecast + terminal", {**V07, "CapIntFix": 1, "OwnOpexOn": 1, "CPM_SP": 12.5, "ShellMW": 3.75, "PowerMW": 8.75, "GPU_Resid": 0.15}),
          ("+ Capacity v2 (YE26 ~302 / YE27 ~1,082 IT MW; new sites with own capex; Basis-tab IT factors)", {**NEWR, "RackOn": 0, "CVDebtBasis": 1}),
          ("+ GB300 capex from rack price: 6.8 x $4.5M + $6.1M = $36.7M leased / $49.2M owned (was 41 / 53)", {**NEWR, "CVDebtBasis": 1}),
          ("+ converts left in debt at CARRYING value (Note 12 amortized cost / Aug notes at face, accreted)", {**NEWR}),
          ("+ Capacity capex v2: GB300 racks 491/72 = 6.82 -> $36.8M", {**NEWR, "GB_Racks": 491 / 72}),
          ("+ Rubin capex v2: leased $41.1M (was 41), owned $52.6M (was 56)", {**NEWR, "GB_Racks": 491 / 72, "CPM_Lease_VR": 41.1, "CPM_Own_VR": 52.6}),
          ("+ lease prepayments $1.6M per new leased IT MW (upfront, 10-yr amortisation)", {**V07B, "CPM_Lease_VR": 41.1, "ND_Px": 22e6 / (491 * 8760), "PrepayMode": 1}),
          ("+ GB300 2027+ new-deal price $16M/IT MW from contracts (was $22M)", {**V07B, "CPM_Lease_VR": 41.1, "PrepayMode": 1}),
          ("+ new prepayments at Q2'26 pace $1.2B/qtr through Q4'28 (was % of capex) (= v0.7b)", {**V07B, "CPM_Lease_VR": 41.1}),
          ("v0.7c 1: asset floor on the SHARED EV-to-equity bridge (floor only; DCF unchanged)", {"TVROResOn": 0, "RubinMW": 18, "VinIdleOn": 0, "FP_Resize": 0, "CPM_Lease_VR": 41.1}),
          ("v0.7c 2: Rubin_Prem 1.2x input removed -> derived output (no output effect, verified)", {"TVROResOn": 0, "RubinMW": 18, "VinIdleOn": 0, "FP_Resize": 0, "CPM_Lease_VR": 41.1}),
          ("v0.7c 3: colo fit-out excluded from leased Rubin as for GB300: $41.1M -> $40.1M", {"TVROResOn": 0, "RubinMW": 18, "VinIdleOn": 0, "FP_Resize": 0}),
          ("v0.7c 4: Rubin price $16.2M/IT MW (bull 20 / bear 14.5; was 18/20/16)", {"TVROResOn": 0, "VinIdleOn": 0, "FP_Resize": 0}),
          ("v0.7c 5: Bloom-slip / Vineland-loss MW idle, not re-sold (no base effect: slip off)", {"TVROResOn": 0, "FP_Resize": 0}),
          ("v0.7c 6: 3-yr re-pricing option added (off in base)", {"TVROResOn": 0, "FP_Resize": 0}),
          ("v0.7c 7: run-off terminal value keeps 15% after-tax GPU residual as exit-fleet cohorts retire", {"FP_Resize": 0}),
          ("v0.7c 8: funding re-sized: ATM -> converts -> equity top-up to $1.0B every quarter to Q4'30 (= v0.7c)", {})]
BRIDGE = [(k, {**V07C, **v}) for k, v in BRIDGE] + [
          ("v0.7d 1: plan shares added to Sh_Basic / FD at the headline basis + authority check (no value change)", {"OffBSWt": 0.5, "T59Align": 0, "PTBasis": 2}),
          ("v0.7d 2: off-B/S commitments 50% -> 0% (same rent / power already in EBITDA; EV unchanged)", {"T59Align": 0, "PTBasis": 2}),
          ("v0.7d 3: Vineland T5-T9 billing at Capacity v2 start quarters, no extra ramp (FY27 $447M -> $1,160M)", {"PTBasis": 2}),
          ("v0.7d 4: HEADLINE PT self-consistent - all new issuance at the solved PT (= v0.7d)", {}),
          ("v0.8: stop-building scenario + Meta backstop toggle added, both OFF in base; floor excludes backstop (no base change) (= v0.8)", {}),
          ("v0.8b: funding-stack toggles (prepay / ABS / uncontracted price) added OFF; backstop built at $52.6M owned Rubin when stop-building (no base change) (= v0.8b)", {}),
          ("v0.8c: equity value per current share before issuance + reconciliation; authority ~130M (no value change) (= v0.8c)", {})]
BRIDGE = [(k, {"CpnChg": 0, "SC_NoPrem": 0, **v}) for k, v in BRIDGE] + [
          ("v0.8d 1: split uncontracted pricing (leased GB300 / owned Rubin) + renewal at 5-yr rate incl. terminal - toggles OFF in base (no base change)", {"CpnChg": 0, "SC_NoPrem": 0}),
          ("v0.8d 2: convert cash coupons in window charged against equity value (after tax)", {"SC_NoPrem": 0}),
          ("v0.8d 3: R5 off - new shares issued at today's value (SC_Disc = Ke/(1+Ke)), not the 12-month PT (= v0.8d)", {})]
GRID = [(st, b, d) for d in (1, 0) for b in (1, 2) for st in (0.0, 0.02, 0.05, 0.10)]
VARS = [("Base (live)", {}), ("Bull", {"Scen": 1}), ("Bear", {"Scen": 3}), ("Gross", {"NDBasis": 2}),
        ("STOP: base contracted only / stop building", {"StopBuild": 1}), ("STOP: bull", {"Scen": 1, "StopBuild": 1}), ("STOP: bear", {"Scen": 3, "StopBuild": 1}), ("STOP: gross", {"NDBasis": 2, "StopBuild": 1}),
        ("STOP: base, commitments 0% (all avoided / sublet)", {"StopBuild": 1, "StopCommitWt": 0}), ("STOP: base, commitments 100% of cut share", {"StopBuild": 1, "StopCommitWt": 1}),
        ("STOP: base WACC 9%", {"StopBuild": 1, "WACC": 0.09}), ("STOP: base WACC 10%", {"StopBuild": 1, "WACC": 0.10}),
        ("STOP: base + Meta backstop 150 MW", {"StopBuild": 1, "BackstopMW": 150}), ("STOP: base + Meta backstop 190 MW", {"StopBuild": 1, "BackstopMW": 190}),
        ("BACKSTOP: base 150 MW at $16M", {"BackstopMW": 150}), ("BACKSTOP: base 190 MW at $16M", {"BackstopMW": 190}), ("BACKSTOP: bear 190 MW at $14.5M", {"Scen": 3, "BackstopMW": 190}),
        ("WACC 9% BEAR", {"Scen": 3, "WACC": 0.09}), ("WACC 10% BEAR", {"Scen": 3, "WACC": 0.10}), ("WACC 11% BEAR", {"Scen": 3, "WACC": 0.11}),
        ("D7: base OffBSWt 50% (v0.7c)", {"OffBSWt": 0.5}), ("D7: base T5-T9 v0.7c ramp", {"T59Align": 0}), ("D7: base spot headline", {"PTBasis": 2}), ("D7: base path headline", {"PTBasis": 3}),
        ("D7: bull spot headline", {"Scen": 1, "PTBasis": 2}), ("D7: bear spot headline", {"Scen": 3, "PTBasis": 2}),
        ("FLOOR: base pre-tax DCF (TaxRate 0)", {"TaxRate": 0}), ("FLOOR: base GB300 new deals $18M", {"ND_Px": 18e6 / (491 * 8760)}), ("FLOOR: base GB300 new deals $20M", {"ND_Px": 20e6 / (491 * 8760)}),
        ("GPU life 4 yrs (all vintages)", {"LifeOverride": 4}), ("GPU life 5 yrs", {"LifeOverride": 5}), ("GPU life 6 yrs", {"LifeOverride": 6}),
        ("Base w/ mgmt connected guide MW (MWSel=3)", {"MWSel": 3}), ("STRESS: energized->billing lag 2 qtrs", {"EnLag": 2}), ("Capex = consensus (CapexMode=1)", {"CapexMode": 1}),
        ("Price-cycle leg, BULL cut 6.25% (base MW path, 50% renewal FY27)", {"PxCycleOn": 1, "PxCyc": 0.0625}),
        ("Price-cycle leg, BASE cut 22% (base MW path, 50% renewal FY27)", {"PxCycleOn": 1, "PxCyc": 0.22}),
        ("Price-cycle leg, BEAR cut 31% (base MW path, 50% renewal FY27)", {"PxCycleOn": 1, "PxCyc": 0.31}),
        ("Base w/ PISLag = 0 (dep at energization)", {"PISLag": 0}), ("Base w/ Microsoft = 24% customer", {"TopCust": 2}),
        ("Capex method 1 (GPUs x $/GPU)", {"CapexMethod": 1}), ("Exit-multiple blend 50% (old TV weighting)", {"DCFWeight": 0.5}),
        ("MEMO: include-both terminal (TermBase=1)", {"TermBase": 1}),
        ("Sens: PD + SG&A held at FY28 % of revenue (OpexHold=1)", {"OpexHold": 1}),
        ("Sens: base replacement capex at price path (ReplPxOn=1)", {"ReplPxOn": 1}),
        ("Sens: uniform capex lead 3 qtrs (v0.3 timing)", {"GPULead": 3, "DCLead": 3}),
        ("Sens: GPU lead 1 qtr (shell/power 3)", {"GPULead": 1}),
        ("Sens: Rubin revenue $14.5M/IT MW (bear)", {"RubinMW": 14.5}), ("Sens: Rubin revenue $16.2M/IT MW (base)", {"RubinMW": 16.2}), ("Sens: Rubin revenue $18M/IT MW (v0.7b)", {"RubinMW": 18}), ("Sens: Rubin revenue $20M/IT MW (bull)", {"RubinMW": 20}),
        ("Sens: Rubin capex $33.3M/IT MW leased (bull)", {"CPM_Lease_VR": 33.3}), ("Sens: Rubin capex $41.1M leased (incl. fit-out, v0.7b)", {"CPM_Lease_VR": 41.1}), ("Sens: Rubin capex $48.2M leased (bear)", {"CPM_Lease_VR": 48.2}),
        ("Sens: $775M facility fully drawn", {"LoanOn": 1}),
        ("Sens: Rubin OWNED capex $45.8M (bull)", {"CPM_Own_VR": 45.8}), ("Sens: Rubin OWNED capex $56M (v0.7)", {"CPM_Own_VR": 56}), ("Sens: Rubin OWNED capex $60.7M (bear)", {"CPM_Own_VR": 60.7}),
        ("Sens: Rubin capex bull both (33.3 leased / 45.8 owned)", {"CPM_Lease_VR": 33.3, "CPM_Own_VR": 45.8}), ("Sens: Rubin capex bear both (48.2 / 60.7)", {"CPM_Lease_VR": 48.2, "CPM_Own_VR": 60.7}),
        ("Sens: Rubin capex leased $32M + owned $43M (v0.6 both)", {"CPM_Lease_VR": 32, "CPM_Own_VR": 43}),
        ("Sens: GPU residual 0% (bear)", {"GPU_Resid": 0}), ("Sens: GPU residual 10%", {"GPU_Resid": 0.10}), ("Sens: GPU residual 25% (~alt: old fleet earns 2 more yrs at 30% net of power)", {"GPU_Resid": 0.25}),
        ("Sens: Capacity v1 (CapV 1)", {"CapV": 1}),
        ("Sens: converts in debt at original principal (CVDebtBasis 1)", {"CVDebtBasis": 1}),
        ("BEAR - pre-funding, converts at principal", {"Scen": 3, "FP_On": 0, "CVDebtBasis": 1}),
        ("BASE - pre-funding (FP_On 0)", {"FP_On": 0}), ("BULL - pre-funding", {"Scen": 1, "FP_On": 0}), ("BEAR - pre-funding", {"Scen": 3, "FP_On": 0}),
        ("BASE gross - pre-funding", {"NDBasis": 2, "FP_On": 0}),
        ("BEAR ATM at 30% below $242.81", {"Scen": 3, "FP_ATMDisc": 0.30}), ("BEAR ATM at 50% below", {"Scen": 3, "FP_ATMDisc": 0.50}),
        ("BULL Avride alt: credit 50% of $2.75B AND add back PV of burn", {"Scen": 1, "AVPt": 0.5, "AVBurnOn": 1}),
        ("BULL Avride v0.6 (credit value, keep burn = double count)", {"Scen": 1, "AVPt": 0.5, "AVBurnOn": 0}),
        ("Sens: converts net-share settled (DilMethod 3)", {"DilMethod": 3}), ("GROSS - converts net-share settled", {"NDBasis": 2, "DilMethod": 3}),
        ("Sens: converts debt at accreted (CVDebtBasis 2)", {"CVDebtBasis": 2}),
        ("RACK $3.8M IT", {"GB_RackPx": 3.8}), ("RACK $3.8M GROSS", {"GB_RackPx": 3.8, "NDBasis": 2}),
        ("RACK $4.0M IT", {"GB_RackPx": 4.0}), ("RACK $4.0M GROSS", {"GB_RackPx": 4.0, "NDBasis": 2}),
        ("RACK $4.5M IT", {"GB_RackPx": 4.5}), ("RACK $4.5M GROSS", {"GB_RackPx": 4.5, "NDBasis": 2}),
        ("RACK $5.0M IT", {"GB_RackPx": 5.0}), ("RACK $5.0M GROSS", {"GB_RackPx": 5.0, "NDBasis": 2}),
        ("RACK $5.5M IT", {"GB_RackPx": 5.5}), ("RACK $5.5M GROSS", {"GB_RackPx": 5.5, "NDBasis": 2}),

        ("STRESS: one-off 31% cut to new-deal $/MW once Rubin ships", {"MYStressOn": 1}),
        ("BULL - gross reading", {"Scen": 1, "NDBasis": 2}), ("BEAR - gross reading", {"Scen": 3, "NDBasis": 2}),
        ("Sens: peer adj. EBITDA margin 62% from FY28 (PeerMarginOn)", {"PeerMarginOn": 1}),
        ("Sens: terminal g = 0 floor (TermRR = 0)", {"TermRR": 0}),
        ("Sens: MW basis company ~190 (history + forward)", {"HistMWBasis": 2}), ("Sens: MW basis MISMATCH (190 history, 173 forward)", {"HistMWBasis": 3}),
        ("Sens: Q3'26 from June exit run-rate (JuneExitOn)", {"JuneExitOn": 1}),
        ("Sens: Vineland shutdown prob-weighted (VinShutSel=7)", {"VinShutSel": 7}),
        ("Sens: Vineland full 3 months (VinShutSel=6)", {"VinShutSel": 6}),
        ("Sens: Bloom slips 2 quarters (T5-T9 +2 qtrs)", {"BloomSlipOn": 1}),
        ("Sens: Vineland prob-weighted + Bloom slips 2 qtrs", {"VinShutSel": 7, "BloomSlipOn": 1}),
        ("Sens: prepayments OFF", {"PrepayOn": 0}),
        ("Capacity proposed path (MWSel=4)", {"MWSel": 4}), ("Capacity proposed path, prepayments OFF", {"MWSel": 4, "PrepayOn": 0}),
        ("Sens: new-deal pricing OFF (GPU-hour engine for new MW)", {"NewDealOn": 0}),
        ("ND: bear $14.5M/IT MW ($3.37/GPU-hr)", {"ND_Px": 14.5e6 / (491 * 8760)}), ("ND: mix-shift $18M/IT MW ($4.18)", {"ND_Px": 18e6 / (491 * 8760)}), ("ND: $20M/IT MW ($4.65)", {"ND_Px": 20e6 / (491 * 8760)}),
        ("ND: v0.7 $22M/IT MW ($5.12) - sensitivity", {"ND_Px": 22e6 / (491 * 8760)}), ("ND: $22M per GROSS MW = $26.4M/IT MW ($6.14) - sensitivity", {"ND_Px": 22e6 / (491 * 8760), "NDBasis": 2}),
        ("PREPAY: guidance case $4.5B H2'26 then Q2 pace", {"PrepayMode": 3}), ("PREPAY: v0.7 % of new-build capex", {"PrepayMode": 1}),
        ("PREPAY: base, pre-funding", {"FP_On": 0}), ("PREPAY: guidance, pre-funding", {"PrepayMode": 3, "FP_On": 0}), ("PREPAY: OFF, pre-funding", {"PrepayOn": 0, "FP_On": 0}),
        ("PREPAY: lease prepayments OFF", {"LeasePrepayOn": 0}),
        ("DRREF: base with 25% refundable placeholder", {"DRRefPct": 0.25}),
        ("VIN: Bloom slip 2 qtrs, v0.7b treatment (slipped MW re-sold)", {"BloomSlipOn": 1, "VinIdleOn": 0}), ("VIN: Bloom slip 2 qtrs, fixed (slipped MW idle)", {"BloomSlipOn": 1}),
        ("VIN: base + T1-T4 SLA credits only", {"VinT14Loss": 1, "VinLossOn": 0}), ("VIN: base + T1-T4 SLA + permanent loss (redeploy after 2 qtrs)", {"VinT14Loss": 1}),
        ("VIN: BEAR + T1-T4 SLA credits only", {"Scen": 3, "VinT14Loss": 1, "VinLossOn": 0}), ("VIN: BEAR + T1-T4 SLA + permanent loss (redeploy after 2 qtrs)", {"Scen": 3, "VinT14Loss": 1}),
        ("VIN: BEAR + T1-T4 SLA + loss + Bloom slip", {"Scen": 3, "VinT14Loss": 1, "BloomSlipOn": 1}),
        ("3YR: BEAR + 3-yr deals (40%) re-price at renewal", {"Scen": 3, "ThreeYrOn": 1}), ("3YR: BASE + 40% 3-yr deals re-price", {"ThreeYrOn": 1, "ThreeYrSh": 0.4}),
        ("WACC 9% BASE", {"WACC": 0.09}), ("WACC 10% BASE", {"WACC": 0.10}), ("WACC 11% BASE", {"WACC": 0.11}),
        ("WACC 9% GROSS", {"WACC": 0.09, "NDBasis": 2}), ("WACC 10% GROSS", {"WACC": 0.10, "NDBasis": 2}), ("WACC 11% GROSS", {"WACC": 0.11, "NDBasis": 2}),
        ("WACC 9% BULL", {"Scen": 1, "WACC": 0.09}), ("WACC 10% BULL", {"Scen": 1, "WACC": 0.10}), ("WACC 11% BULL", {"Scen": 1, "WACC": 0.11}),
        ("FUND: base v0.7b plan (not re-sized)", {"FP_Resize": 0}), ("FUND: bull v0.7b plan", {"Scen": 1, "FP_Resize": 0}), ("FUND: bear v0.7b plan", {"Scen": 3, "FP_Resize": 0}),
        ("FUND: base top-up only to YE27", {"FP_TopEnd": 16}), ("FUND: bear top-up only to YE27", {"Scen": 3, "FP_TopEnd": 16}),
        ("RESID: run-off TV without residual (v0.7b)", {"TVROResOn": 0}),
        ("RESID: bull alt 2 more yrs re-lease at 30% of $/MW net of opex (~7.7% of GPU capex)", {"GPU_Resid": 0.077}), ("DRREF: bear without placeholder", {"Scen": 3, "DRRefPct": 0}),
        ("FLOOR: short-term share 0%", {"STShare": 0}), ("Sens: short-term share 2% of total MW", {"STShare": 0.02}), ("Sens: short-term share 10% of total MW", {"STShare": 0.10}),
        ("HEADLINE: unsold new MW at $18M/IT MW (contracted backlog at existing pricing)", {"MYRateU": 18}), ("HEADLINE: unsold new MW at $22M/IT MW", {"MYRateU": 22}), ("HEADLINE: unsold new MW at $25M/IT MW", {"MYRateU": 25}),
        ("EXTREME (unsupported): short-term share 25%", {"STShare": 0.25}),
        ("Sens: new-deal basis per GROSS MW (= $26.4M per IT MW)", {"NDBasis": 2}),
        ("Sens: prepayment decay OFF (other deals 38.5% flat)", {"PrepayDecayOn": 0}), ("Sens: new hyperscaler prepay 20%", {"HSPre": 0.20}), ("Sens: new hyperscaler prepay 40%", {"HSPre": 0.40}),
        ("Sens: DR run-off deduction OFF (v0.4 cash treatment)", {"DRTermOn": 0}),
        ("Sens: GPU anchor OFF (Capacity file history)", {"GPUAnchorOn": 0}),
        ("Sens: AIB CLT1 prepayment $250M", {"CLT1Prepay": 250}), ("Sens: AIB CLT1 prepayment $500M", {"CLT1Prepay": 500}),
        ("Sens: Inferize $130M Q4'26 (press-reported)", {"InferizeOn": 1}),
        ("BEAR + power loosens 2028 (-15%/yr)", {"Scen": 3, "PwrLooseOn": 1}),
        ("BULL - v0.3 settings", {**NEW0, "Scen": 1, "GPULead": 3, "DCLead": 2, "PxPathMode": 0, "PxChg_B": 0.0, "PxChg_N": -0.05 / 3, "TGMode": 0, "Rubin_Prem": 1.5 / 1.22, "CLT1Rent": 2.0 * 1.2}),
        ("BEAR - v0.3 settings", {**NEW0, "Scen": 3, "GPULead": 3, "DCLead": 2, "PxPathMode": 0, "PxChg_B": -0.25, "PxChg_N": -0.20, "TGMode": 0, "Rubin_Prem": 1.5 / 1.22, "CLT1Rent": 2.5 * 1.35})] + [("BRIDGE " + k, v) for k, v in BRIDGE] + [(f"GRID {st}|{b}|{d}", {"STShare": st, "NDBasis": b, "PrepayDecayOn": d}) for st, b, d in GRID]
MET = [("FY26 revenue", "=SUMIF(Q_Yr,2026,PL_Rev)"), ("FY27 revenue", "=SUMIF(Q_Yr,2027,PL_Rev)"), ("FY28 revenue", "=SUMIF(Q_Yr,2028,PL_Rev)"),
       ("FY27 adj. EBITDA", "=SUMIF(Q_Yr,2027,PL_EBITDA)"), ("FY27 EBIT", "=SUMIF(Q_Yr,2027,PL_EBIT)"), ("FY27 D&A", "=SUMIF(Q_Yr,2027,PL_DA)"),
       ("DCF EV", "=V_EV"), ("FD value/share", "=V_PS"), ("12m PT", "=V_PT"), ("Liquidity <0 qtr", "=Liq_NegQ"), ("YE27 liquidity", "=INDEX(Liq,1,16)"), ("FY27 capex", "=SUMIF(Q_Yr,2027,Capex)"),
       ("FY28 adj. EBITDA", "=SUMIF(Q_Yr,2028,PL_EBITDA)"), ("YE28 liquidity", "=INDEX(Liq,1,20)"), ("FY26 adj. EBITDA", "=SUMIF(Q_Yr,2026,PL_EBITDA)"),
       ("FY28 EBIT", "=SUMIF(Q_Yr,2028,PL_EBIT)"), ("FY26 capex", "=SUMIF(Q_Yr,2026,Capex)"), ("FY28 capex", "=SUMIF(Q_Yr,2028,Capex)"),
       ("FY26 EBIT", "=SUMIF(Q_Yr,2026,PL_EBIT)"), ("Equity value FD", "=V_EQ"), ("Q2'26 util", "=UtilQ2"), ("TV PV", "=V_PVTV"), ("Term capex", "=V_CXT"),
       ("6m value", "=V_PT6"), ("3y value", "=V_PT36"), ("GB300 rev/IT MW gross", "=RM_GB300"), ("GB300 net leased", "=RM_GB300_L"), ("GB300 net owned", "=RM_GB300_O"), ("Rubin rev/IT MW", "=RM_Rubin"),
       ("Q3'26 rev", "=INDEX(PL_Rev,1,11)"), ("Terminal g", "=V_g"), ("FY30 margin", "=SUMIF(Q_Yr,2030,PL_EBITDA)/SUMIF(Q_Yr,2030,PL_Rev)"), ("YE26 liquidity", "=INDEX(Liq,1,12)"), ("FY28 rev", "=SUMIF(Q_Yr,2028,PL_Rev)"),
       ("New-MW blended $/IT MW-yr", "=RM_ND"), ("Vineland loss FY26-27", "=INDEX(VinLost,1,12)+INDEX(VinLost,1,13)"), ("Min liquidity", "=MIN(Liq)"), ("YE27 active MW", "=INDEX(MW_Act,1,16)"), ("Total new prepayments", "=SUM(CF_Prepay)"), ("DR run-off adj", "=V_DRAdj"), ("Sold share new MW", "=RPO_Sold"), ("Post-funding $/sh", "=V_PSF"), ("Post-funding FD sh", "=V_SHF"), ("Pre-funding FD sh", "=V_SH"), ("Min liq to YE27", "=MIN(INDEX(Liq,1,11):INDEX(Liq,1,16))"), ("ATM shares", "=FP_ATMSh"), ("Avride add-back", "=V_AVAdd"), ("SOTP", "=V_SOTP"), ("FD sh full-conv", "=V_SHFull"), ("FD sh NSS", "=V_SHNSS"),
       ("Asset floor EV", "=PF_EV"), ("Stop-build commitments in debt", "=V_ODEBT"), ("Funding gap", "=X_Gap"), ("Eq value / current share pre-issuance", "=X_PS0"), ("YE27 non-contract MW", "=INDEX(MW_Act,1,16)-INDEX(A_MW,1,16)-INDEX(A_PostMW,1,16)"), ("Capex Q3'26-Q4'30", "=SUMPRODUCT((Q_Idx>=11)*(Q_Idx<=28)*Capex)"), ("Floor $/sh cons", "=PF_PS1"), ("Floor $/sh mark", "=PF_PS2"), ("Floor $/sh high", "=PF_PS3"),
       ("NPVc leased Rubin", "=PB_NPVc_LR"), ("NPVc owned Rubin", "=PB_NPVc_OR"), ("NPVat leased Rubin", "=PB_NPVat_LR"), ("NPVat owned Rubin", "=PB_NPVat_OR"), ("NPV pre leased Rubin", "=PB_NPV_LR"), ("NPV pre owned Rubin", "=PB_NPV_OR"),
       ("PX GB IT", "=PX_GB"), ("PX GB gross", "=PX_GBg"), ("PX VR", "=PX_VR"), ("DX FY27", "=DX_FY27"), ("DX model FY27", "=DX_Model27"), ("FY27 AI rev", "=SUMIF(Q_Yr,2027,AIRev)"),
       ("Post FD full-conv", "=V_SHFFull"), ("YE26 energized", "=INDEX(MW_En,1,12)"), ("YE27 energized", "=INDEX(MW_En,1,16)"), ("Floor flag", "=PF_Flag"), ("PB RU IT", "=PB_RU_IT"), ("IRR leased Rubin", "=PB_RU_IRR"), ("IRR owned Rubin", "=PB_IRR_OR"),
       ("CV debt at solved (pre)", "=V_CVDEBT"), ("CV debt post", "=V_CVDEBTF"), ("Issuance need", "=V_SHFFull-V_SHFull+SUM(CF_EqSh)"), ("Refundable DR placeholder at V_Q", "=DRRefPct*INDEX(CF_DRBal,1,V_Q)"), ("Lease prepayments total", "=SUM(CapLPre)"), ("New prepayments H2'26-FY28", "=SUM(INDEX(CF_Prepay,1,11):INDEX(CF_Prepay,1,20))"), ("Floor $/sh SHARED (headline)", "=PF_PSHead"), ("Floor PT shared", "=PF_PTH"), ("FY27 Microsoft rev", "=SUMIF(Q_Yr,2027,MS_Rev)"),
       ("NPVc leased GB300", "=Payback_Floor!$P$28"), ("NPVat leased GB300", "=Payback_Floor!$O$28"), ("Top-up equity $M", "=SUM(FP_EqTop)"), ("Top-up shares M", "=SUM(FP_EqTopSh)"),
       ("Post-funding new shares M", "=V_FPSh"), ("Min liq all qtrs", "=MIN(INDEX(Liq,1,11):INDEX(Liq,1,28))"), ("Rubin/GB300 GPU-hr x", "=PX_Ratio"), ("Rubin/GB300 capex x", "=PX_CapRatio"),
       ("PT self-consistent", "=V_PTSC"), ("PT spot memo", "=V_PTSpot"), ("PT path memo", "=V_PTPath"), ("Value/sh SC", "=V_PSSC"), ("Issue px SC", "=V_PISC"), ("New sh headline M", "=V_NewShHead"),
       ("Convert reserve headline M", "=V_CVResHead"), ("Authority need M", "=V_AuthNeed"), ("New sh SC M", "=V_NewShSC"), ("New sh path M", "=V_NewShPath"), ("Floor $/sh SC", "=PF_PSSC"),
       ("PF PV pre-tax", "=PF_PV"), ("PF PV after-tax", "=PF_PVat"), ("PF capex", "=PF_Cap"), ("SC re-funding $", "=V_XRSC"), ("SC FD sh", "=V_SHSC"), ("Spot new sh M", "=V_FPSh"), ("PT SC conv", "=V_CONVSC"), ("Path end px", "=V_PEndPath")]
res = {}
for k, (lab, ch) in enumerate(VARS):
    wb = openpyxl.load_workbook(SRC)
    ws = wb["Inputs"]
    for nm, v in ch.items():
        if nm == "Scen": ws["D4"] = v
        else:
            ws[f"D{inrow(nm)}"] = v
            sc = ch.get("Scen", 2)
            if sc != 2 and isinstance(ws[f"{'C' if sc == 1 else 'E'}{inrow(nm)}"].value, (int, float)): ws[f"{'C' if sc == 1 else 'E'}{inrow(nm)}"] = v
    sx = wb.create_sheet("_snap")
    for i, (m, f) in enumerate(MET): sx[f"A{i+1}"] = f
    p = f"/tmp/snap/var{k}.xlsx"; os.makedirs("/tmp/snap", exist_ok=True); wb.save(p)
    out = recalc(p, "/tmp/snap/out")
    r = openpyxl.load_workbook(out, data_only=True)["_snap"]
    res[lab] = [r[f"A{i+1}"].value for i in range(len(MET))]
    print(lab, [round(x, 1) if isinstance(x, (int, float)) else x for x in res[lab]])
# write static tables
wb = openpyxl.load_workbook(SRC)
se = wb["Sensitivity"]
BOLD = Font(bold=True, size=10); HDR = Font(bold=True, color="FFFFFF", size=10); HF = PatternFill("solid", fgColor="1F3864"); IT = Font(italic=True, size=9, color="595959")
r0 = 12
se[f"A{r0-1}"] = "2) Scenario / toggle snapshot - STATIC values from LibreOffice recalcs at build time (re-run snap.py after changing inputs). $M except per-share."; se[f"A{r0-1}"].font = IT
se.cell(r0, 1, "Variant").font = HDR; se.cell(r0, 1).fill = HF
for j, (m, f) in enumerate(MET):
    c = se.cell(r0, 2 + j, m); c.font = HDR; c.fill = HF
for i, (lab, ch) in enumerate(VARS):
    se.cell(r0 + 1 + i, 1, lab)
    for j, v in enumerate(res[lab]):
        c = se.cell(r0 + 1 + i, 2 + j, v); c.number_format = "0.00" if j in (7, 8) else "#,##0;(#,##0)"
rr = r0 + len(VARS) + 3
se[f"A{rr}"] = "3) FY27 EBIT at GPU life 4 / 5 / 6 yrs vs Skeptic"; se[f"A{rr}"].font = BOLD
for j, t in enumerate(["", "4 yrs", "5 yrs", "6 yrs"]):
    if t: se.cell(rr + 1, 1 + j, t).font = BOLD
se.cell(rr + 2, 1, "Model FY27 EBIT ($M, static)"); se.cell(rr + 3, 1, "Skeptic FY27 EBIT ($M)"); se.cell(rr + 4, 1, "Model - Skeptic")
for j, (lab, sk) in enumerate([("GPU life 4 yrs (all vintages)", "SkEBIT4"), ("GPU life 5 yrs", "SkEBIT5"), ("GPU life 6 yrs", "SkEBIT6")]):
    se.cell(rr + 2, 2 + j, res[lab][4]).number_format = "#,##0;(#,##0)"
    se.cell(rr + 3, 2 + j, f"={sk}").number_format = "#,##0"
    cl = "BCD"[j]; se.cell(rr + 4, 2 + j, f"={cl}{rr+2}-{cl}{rr+3}").number_format = "#,##0;(#,##0)"
for j in range(10): se.column_dimensions["BCDEFGHIJK"[j]].width = 14
ou = wb["Outputs"]
lab2row = {ou.cell(r, 1).value: r for r in range(1, ou.max_row + 1)}
for lab, var in [("FY27 revenue - low (model: bear)", "Bear"), ("FY27 revenue - high (model: bull)", "Bull"), ("FY27 revenue - STRESS (2-qtr energized lag)", "STRESS: energized->billing lag 2 qtrs")]:
    ou.cell(lab2row[lab], 2, res[var][1]).number_format = "#,##0"
r_ = ou.max_row + 2
ou.cell(r_, 1, "Funding runway: capex toggle side by side (STATIC, snap.py)").font = BOLD
for j, t in enumerate(["", "Path-consistent (CapexMode 2, method 3)", "Consensus (CapexMode 1)"]):
    if t: ou.cell(r_ + 1, 1 + j, t).font = BOLD
ou.cell(r_ + 2, 1, "First quarter liquidity < 0"); ou.cell(r_ + 3, 1, "YE27 liquidity ($M)"); ou.cell(r_ + 4, 1, "FY27 capex ($M)")
for j, var in enumerate(["Base (live)", "Capex = consensus (CapexMode=1)"]):
    ou.cell(r_ + 2, 2 + j, res[var][9]); ou.cell(r_ + 3, 2 + j, res[var][10]).number_format = "#,##0;(#,##0)"; ou.cell(r_ + 4, 2 + j, res[var][11]).number_format = "#,##0"
r_ = r_ + 6
B_ = res["Base (live)"]
ou.cell(r_, 1, "Price-cycle leg alone (Blackwell cut after Rubin; 50% reserved renewal in FY27; base MW path) vs base, STATIC ($M)").font = BOLD
for j, t in enumerate(["", "FY27 revenue", "FY27 adj. EBITDA", "FY28 revenue", "FY28 adj. EBITDA", "Skeptic 03 v3.4 FY27 rev", "Skeptic 03 v3.4 FY28 rev"]):
    if t: ou.cell(r_ + 1, 1 + j, t).font = BOLD
for i, (lab, sk) in enumerate([("Price-cycle leg, BULL cut 6.25% (base MW path, 50% renewal FY27)", (-77, -370)), ("Price-cycle leg, BASE cut 22% (base MW path, 50% renewal FY27)", (-271, -1290)), ("Price-cycle leg, BEAR cut 31% (base MW path, 50% renewal FY27)", (-383, -1810))]):
    v = res[lab]; ou.cell(r_ + 2 + i, 1, lab)
    for j, (a, b) in enumerate([(v[1], B_[1]), (v[3], B_[3]), (v[2], B_[2]), (v[12], B_[12])]):
        ou.cell(r_ + 2 + i, 2 + j, a - b).number_format = "#,##0;(#,##0)"
    ou.cell(r_ + 2 + i, 6, sk[0]); ou.cell(r_ + 2 + i, 7, sk[1])
r_ = r_ + 7
ou.cell(r_, 1, "Comparison vs Skeptic v3.2 (skeptic_03_depreciation.xlsx). NOTE: Skeptic's path run uses CONSENSUS EBITDA; ours is bottom-up. Model values STATIC from snap.py").font = BOLD
for j, t in enumerate(["", "Model zero qtr", "Skeptic zero qtr", "Model YE27 ($B)", "Skeptic YE27 ($B)", "Model YE28 ($B)", "Skeptic YE28 ($B)"]):
    if t: ou.cell(r_ + 1, 1 + j, t).font = BOLD
rows = [("Path-tied capex - base", "Base (live)", "=Sk3_ZeroQ", "=Sk3_YE27", "=Sk3_YE28"), ("Low (model: bear scenario)", "Bear", "=Sk3L_ZeroQ", "=Sk3L_YE27", "=Sk3L_YE28"),
        ("High (model: bull scenario)", "Bull", "=Sk3H_ZeroQ", "=Sk3H_YE27", "=Sk3H_YE28"), ("Consensus capex", "Capex = consensus (CapexMode=1)", "=Sk3C_ZeroQ", "=Sk3C_YE27", None)]
for i, (lab, var, z, y7, y8) in enumerate(rows):
    v = res[var]; rr = r_ + 2 + i; ou.cell(rr, 1, lab); ou.cell(rr, 2, v[9]); ou.cell(rr, 3, z)
    ou.cell(rr, 4, v[10] / 1000).number_format = "0.0"; ou.cell(rr, 5, y7).number_format = "0.0"
    ou.cell(rr, 6, v[13] / 1000).number_format = "0.0"
    if y8: ou.cell(rr, 7, y8).number_format = "0.0"
rr = r_ + 7
ou.cell(rr, 1, "DCF EV at GPU life ($B): model vs Skeptic v3.2 (market EV ~$74.7B)").font = BOLD
for j, (lab, sk) in enumerate([("GPU life 4 yrs (all vintages)", "=Sk3_DCF4"), ("GPU life 5 yrs", "=Sk3_DCF5"), ("GPU life 6 yrs", "=Sk3_DCF6")]):
    ou.cell(rr + 1, 2 + j, ["4 yrs", "5 yrs", "6 yrs"][j]).font = BOLD
    ou.cell(rr + 2, 2 + j, res[lab][6] / 1000).number_format = "0.0"; ou.cell(rr + 3, 2 + j, sk).number_format = "0.0"
ou.cell(rr + 2, 1, "Model DCF EV"); ou.cell(rr + 3, 1, "Skeptic v3.2 DCF EV"); ou.cell(rr + 4, 1, "Market EV"); ou.cell(rr + 4, 2, "=V_MKTEV/1000").number_format = "0.0"
rr = rr + 6
ou.cell(rr, 1, "v0.3 -> v0.4 by scenario (full bridge on Bridge tab), STATIC").font = BOLD
for j, t in enumerate(["", "DCF EV $M", "FD equity $M", "FD $/share", "12m PT", "Q2'26 util", "FY27 rev", "FY27 EBITDA"]):
    if t: ou.cell(rr + 1, 1 + j, t).font = BOLD
for i, lab in enumerate(["BRIDGE + GPU capex lead 3 qtrs (= v0.3)", "Base (live)", "BULL - v0.3 settings", "Bull", "BEAR - v0.3 settings", "Bear", "MEMO: include-both terminal (TermBase=1)", "Base w/ mgmt connected guide MW (MWSel=3)"]):
    v = res[lab]; ou.cell(rr + 2 + i, 1, lab)
    for j, x in enumerate([v[6], v[19], v[7], v[8], v[20], v[1], v[3]]):
        c = ou.cell(rr + 2 + i, 2 + j, x); c.number_format = "0.0%" if j == 4 else ("0.00" if j in (2, 3) else "#,##0;(#,##0)")
va = wb["Valuation"]
def put(nm, val):
    sh, c = loc(nm); wb[sh][c] = val
put("V_EVMemo", res["MEMO: include-both terminal (TermBase=1)"][6]); put("V_PSMemo", res["MEMO: include-both terminal (TermBase=1)"][7])
for nm, lab, j in [("V_EVBaseS", "Base (live)", 6), ("V_EVBullS", "Bull", 6), ("V_EVBearS", "Bear", 6), ("V_UtBaseS", "Base (live)", 20), ("V_UtBullS", "Bull", 20), ("V_UtBearS", "Bear", 20)]:
    put(nm, res[lab][j])
br = wb["Bridge"]
br["A2"] = "Base scenario. Each row = cumulative settings recomputed in the CURRENT model (bug fixes such as the renewal-formula bracket and scenario-invariant history are in every row). $B and $/share are FD value/share."
hd = ["Step", "DCF EV $B", "Change $B", "FD $/share", "Change $/sh", "12m PT $", "FY26 rev $M", "FY27 rev $M", "FY28 rev $M", "FY27 EBITDA $M", "YE27 liquidity $M", "First neg. qtr"]
for j, t in enumerate(hd):
    c = br.cell(4, 1 + j, t); c.font = HDR; c.fill = HF
prev = None
for i, (k, _) in enumerate(BRIDGE):
    v = res["BRIDGE " + k]; r = 5 + i
    br.cell(r, 1, k)
    vals = [v[6] / 1000, (v[6] - prev[6]) / 1000 if prev else None, v[7], (v[7] - prev[7]) if prev else None, v[8], v[0], v[1], v[2], v[3], v[10], v[9]]
    for j, x in enumerate(vals):
        c = br.cell(r, 2 + j, x); c.number_format = "0.0;(0.0)" if j in (0, 1) else ("0.00;(0.00)" if j in (2, 3, 4) else "#,##0;(#,##0)")
    if k.startswith("v0.1") or "(= v0.2)" in k or "(= v0.3)" in k or "(= v0.4)" in k or "(= v0.5)" in k or "(= v0.6)" in k or "(= v0.7)" in k or "(= v0.7b)" in k or "(= v0.7c)" in k or "(= v0.7d)" in k or "(= v0.8)" in k or "(= v0.8b)" in k or "(= v0.8c)" in k or "(= v0.8d)" in k:
        for j in range(12): br.cell(r, 1 + j).font = BOLD
    prev = v
br.column_dimensions["A"].width = 58
for j in range(11): br.column_dimensions["BCDEFGHIJKL"[j]].width = 13
r_ = 5 + len(BRIDGE) + 2
br.cell(r_, 1, "Horizon per-share values (FD value/share rolled at Ke; excludes future funding dilution - see Funding tab)").font = BOLD
for j, t in enumerate(["Scenario", "Today", "6-month", "12-month PT", "3-year"]):
    c = br.cell(r_ + 1, 1 + j, t); c.font = HDR; c.fill = HF
for i, lab in enumerate(["Base (live)", "Bull", "Bear"]):
    v = res[lab]; br.cell(r_ + 2 + i, 1, lab)
    for j, x in enumerate([v[7], v[23], v[8], v[24]]): br.cell(r_ + 2 + i, 2 + j, x).number_format = "0.00"
r_ = r_ + 7
br.cell(r_, 1, "New-MW pricing vs IREN >$20M/IT MW-yr ($M per IT MW-yr). GB300 GPU-hour engine (legacy fleet) at Oct-1 list x RealRes x Util_Res; new 2027+ MW at new-deal blend (col G)").font = BOLD
for j, t in enumerate(["Scenario", "GB300 gross", "net leased (-$2.2M rent)", "net owned (-O&M)", "Rubin gross", "IREN gross / leased / owned", "New-deal blend gross"]):
    c = br.cell(r_ + 1, 1 + j, t); c.font = HDR; c.fill = HF
for i, lab in enumerate(["Base (live)", "Bull", "Bear"]):
    v = res[lab]; br.cell(r_ + 2 + i, 1, lab)
    for j, x in enumerate([v[25], v[26], v[27], v[28]]): br.cell(r_ + 2 + i, 2 + j, x).number_format = "0.0"
    br.cell(r_ + 2 + i, 6, "20.0 / 17.8 / 19.25"); br.cell(r_ + 2 + i, 7, v[34]).number_format = "0.0"
r_ = r_ + 7
br.cell(r_, 1, "v0.5 BASE RANGE: 12m PT ($/share) by short-term share (of TOTAL active MW) x new-deal $/MW basis, with / without prepayment decay (STATIC). Green = base; yellow = supported tests; orange = gross-MW basis (alternative reading of slide); red = equity floored at 0").font = BOLD
GF = {"base": PatternFill("solid", fgColor="C6EFCE"), "test": PatternFill("solid", fgColor="FFEB9C"), "gross": PatternFill("solid", fgColor="F8CBAD"), "zero": PatternFill("solid", fgColor="FF9999")}
for bi, d in enumerate((1, 0)):
    r0_ = r_ + 1 + bi * 6
    br.cell(r0_, 1, "Prepayment decay ON (base)" if d == 1 else "Prepayment decay OFF (other deals 38.5% flat)").font = BOLD
    for j, st in enumerate((0.0, 0.02, 0.05, 0.10)):
        c = br.cell(r0_, 2 + j, f"ST share {st:.0%}" + (" (floor)" if st == 0 else "")); c.font = HDR; c.fill = HF
    for j, t_ in enumerate(["12m PT", "EV $B", "YE27 liq $B", "1st neg qtr"]):
        pass
    for i, b in enumerate((1, 2)):
        br.cell(r0_ + 1 + i, 1, "IT basis $22M/IT MW" if b == 1 else "Gross-MW basis = $26.4M/IT MW")
        for j, st in enumerate((0.0, 0.02, 0.05, 0.10)):
            v = res[f"GRID {st}|{b}|{d}"]
            c = br.cell(r0_ + 1 + i, 2 + j, v[8]); c.number_format = "0.0"
            k_ = "zero" if (v[8] or 0) <= 0 else ("base" if (st == 0.05 and b == 1 and d == 1) else ("gross" if b == 2 else "test"))
            c.fill = GF[k_]
            br.cell(r0_ + 3 + i, 1, ("  EV $B / YE27 liq $B / 1st neg - " + ("IT" if b == 1 else "gross")))
            br.cell(r0_ + 3 + i, 2 + j, f"{v[6]/1000:.1f} / {v[10]/1000:.1f} / {v[9]}")
wb.save(SRC)
print("written")
