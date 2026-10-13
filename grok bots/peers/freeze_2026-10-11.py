#!/usr/bin/env python3
# Freeze the Sunday (Oct 11 2026) peer deliverables: hash + list them in peers/FREEZE.md and make READ-ONLY copies under
# peers/frozen_2026-10-11/ (same relative paths). Originals are not touched. Re-running refuses if the frozen dir exists.
import os, hashlib, shutil, stat, datetime, sys
P = "/workspace/nebius/peers"; V31 = "--v3.1" in sys.argv; V3 = "--v3" in sys.argv and not V31; V2 = "--v2" in sys.argv and not V3; FZ1 = f"{P}/frozen_2026-10-11"; FZ2 = f"{P}/frozen_2026-10-11_v2"
FZ = f"{P}/frozen_2026-10-11_v3.1" if V31 else f"{P}/frozen_2026-10-11_v3" if V3 else (FZ2 if V2 else FZ1)
# v3 mode (2026-10-09 ~4:40 PM): new read-only copies in frozen_2026-10-11_v3/ + RECONCILE.md (v2 copy + 'resolved in v3' note); v1 and v2 dirs never touched.
# v2 mode (2026-10-09 PM): same deliverable list, new read-only copies in frozen_2026-10-11_v2/; frozen_2026-10-11/ (v1) is never touched.
L_NEB = "MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16"
L_BAND = "comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M"
L_K13 = "K13 capex/MW 77.9 vs peer median 33.18"
GROUPS = [
 ("Memo (3 pages)", ["memo/nebius_vs_peers_unit_economics.docx", "memo/nebius_vs_peers_unit_economics.pdf", "memo/fig_unit_economics.png",
                     "memo/build_memo.py", "memo/build_figs.py", "memo/common_data.py"], f"{L_NEB}; {L_BAND}; {L_K13}; Q3 date ~Nov 10 (est.)"),
 ("Summary table", ["memo/nebius_vs_peers_summary_table.xlsx", "memo/nebius_vs_peers_summary_table.png", "memo/build_summary_table.py"], f"{L_NEB}; {L_BAND}; {L_K13}"),
 ("Comparison", ["comparison.xlsx", "build_comparison.py", "build_charts.py", "nebius_locks.py"] + [f"charts/{c}" for c in (
     "active_vs_contracted_mw.png", "capex_per_mw.png", "contracts_usd_per_mw_yr.png", "cost_of_debt.png", "depreciation_yield_useful_life.png", "nebius_billing_lock.png",
     "rev_per_active_mw.png", "rev_per_gpu_hr.png", "rpo_to_annualized_revenue.png", "sites_active_vs_announced_mw.png", "top_customer_pct.png")],
  f"{L_NEB}; {L_BAND}; {L_K13}; FD 358.56M debt / 376.65M if-converted (model Inputs)"),
 ("price_or_mw", ["price_or_mw.xlsx", "charts/price_or_mw.png", "patch_price_or_mw.py", "build_price_or_mw_chart.py"],
  "Nebius_Check B37 = 27.6278 ex-MSFT hurdle; B6 23.1272; B22 14.494 (MSFT house, Contracts!E12; 11.60/17.08 HISTORY only); billed YE26/FY27 locks; band $9.7–16.7M"),
 ("Benchmark", ["benchmark/peer_benchmark.xlsx", "benchmark/peer_benchmark.docx", "benchmark/peer_benchmark.pdf", "benchmark/peer_benchmark.png", "benchmark/build_peer_benchmark.py"],
  f"{L_NEB}; {L_BAND}; YE26 guide hurdles 27.16/31.67"),
 ("Site appendices", sorted([f"sites/{f}" for f in os.listdir(f"{P}/sites") if f.startswith("appendix_") and f.split(".")[-1] in ("md", "docx", "pdf")])
  + ["sites/peer_site_appendices.pdf", "sites/sites.xlsx", "sites/site_verification.md", "sites/build_site_appendices.py", "sites/site_data.py"]
  + sorted(f"sites/charts/{f}" for f in os.listdir(f"{P}/sites/charts") if f.endswith(".png")),
  "site active-base MW: Denton ~260 IT, Childress 75 (50–100), Abilene 618 (421–618), Port Washington 0, KC 12, Monarch/Kvandal 0; energized 181.5→193 (Nebius context)"),
 ("Final appendix pack (this build)", ["final_appendix/peer_kpi_appendix.xlsx", "final_appendix/peer_kpi_appendix.docx", "final_appendix/peer_kpi_appendix.pdf",
     "final_appendix/chart_usd_per_mw_yr.png", "final_appendix/chart_contracted_vs_live_mw.png", "final_appendix/chart_capex_per_mw.png", "final_appendix/build_peer_kpi_appendix.py"],
  f"ALL: {L_NEB}; {L_BAND}; {L_K13}; asserted at build + cell-for-cell vs comparison.xlsx / peer_benchmark.xlsx / price_or_mw.xlsx / sites"),
]
if os.path.exists(FZ): sys.exit(f"{FZ} exists - refusing to overwrite a freeze")
def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""): h.update(b)
    return h.hexdigest()
rows, sums = [], []
for g, files, locks in GROUPS:
    for rel in files:
        src = f"{P}/{rel}"; assert os.path.isfile(src), src
        st = os.stat(src); h = sha(src)
        rows.append((g, rel, datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M ET"), st.st_size, h))
        dst = f"{FZ}/{rel}"; os.makedirs(os.path.dirname(dst), exist_ok=True); shutil.copy2(src, dst)
        assert sha(dst) == h, rel; os.chmod(dst, 0o444); sums.append(f"{h}  {rel}")
now = datetime.datetime.now().strftime("%a %b %d, %Y %H:%M ET")
if V31:
    FZ3 = f"{P}/frozen_2026-10-11_v3"
    v3 = {l.split("  ", 1)[1]: l.split("  ", 1)[0] for l in open(f"{FZ3}/SHA256SUMS").read().splitlines() if l.strip()}
    changed = [r for r in rows if v3.get(r[1]) != r[4]]
    assert [r[1] for r in changed] == ["benchmark/peer_benchmark.xlsx"], [r[1] for r in changed]
    rec = open(f"{FZ3}/RECONCILE.md", encoding="utf-8").read().rstrip("\n") + "\n\n" + open(f"{P}/reconcile_v3_1_note.md", encoding="utf-8").read()
    open(f"{FZ}/RECONCILE.md", "w", encoding="utf-8").write(rec); os.chmod(f"{FZ}/RECONCILE.md", 0o444); rec_h = sha(f"{FZ}/RECONCILE.md")
    sums.append(f"{rec_h}  RECONCILE.md")
    old = open(f"{P}/FREEZE.md", encoding="utf-8").read().split("\n", 1)[1]
    old = old.replace("## CURRENT: v3 - frozen", "## SUPERSEDED: v3 - frozen", 1)
    md = ["# FREEZE - peer deliverables for Sunday Oct 11, 2026 finals", "",
          f"## CURRENT: v3.1 - FINAL for Sunday - frozen {now} (supersedes v3, v2, v1)", "",
          f"{len(rows)} files in {len(GROUPS)} groups ({sum(r[3] for r in rows):,} bytes) + RECONCILE.md; {len(changed)} changed vs v3. Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11_v3.1/` + `SHA256SUMS` (all {len(rows)} files AND RECONCILE.md). "
          "v1, v2, v3 dirs left untouched (superseded). Verify: `cd peers/frozen_2026-10-11_v3.1 && sha256sum -c SHA256SUMS`.", "",
          "### What changed in v3.1 and why (coordinator, Oct 9 ~4:44 PM ET; tags only)", "",
          "peer_benchmark.xlsx rebuilt so Comp_Band picks up GPU Market's 4:38 PM ET weakest-input tags (market/contract_economics.xlsx Print_Table; RECONCILE deck-match fixes): "
          "D16 'R / D' -> 'R/D'; D19 'D' -> 'D, term P'; D20 'D on E' -> 'E, house'; D21 'E' -> 'P (IT read E)'; D22 'R / D' -> 'R/D, term P'; D23 'R (secondary) / D' -> 'P'. "
          "No other cell (values or formulas, all sheets) differs vs v3; 84 formulas, recalc 0 errors; no value, median, band or lock moved.",
          "peer_benchmark.pdf / .png / .docx: rebuilt output was NOT shipped - kept the v3 bytes. PNG rebuilt byte-identical; PDF differed in 94 bytes, all in CreationDate, trailer /ID and /DocChecksum (extracted text and 150-dpi render pixels identical); docx differed only in zip container metadata (every member part byte-identical).",
          "Re-checks: benchmark build asserts, appendix build (22 lock checks + cell-for-cell vs comparison / peer_benchmark / price_or_mw; outputs content-identical, v3 bytes restored), nebius_locks.py - all pass. Pages 1/2/3.", "",
          "### Changed files (v3 -> v3.1 SHA-256)", "", "| Path (peers/) | v3 SHA-256 | v3.1 SHA-256 | v3.1 size |", "|---|---|---|---:|"]
    md += [f"| `{r[1]}` | `{v3.get(r[1])[:16]}...` | `{r[4][:16]}...` | {r[3]:,} |" for r in changed]
    md += [f"| `RECONCILE.md` (v3 note + v3.1 line) | `{sha(f'{FZ3}/RECONCILE.md')[:16]}...` | `{rec_h[:16]}...` | {os.path.getsize(f'{FZ}/RECONCILE.md'):,} |", ""]
    md += ["### Locks (unchanged; re-asserted)", "", f"{L_NEB}; {L_BAND}; {L_K13}; B37 27.6278; Meta #1 $11.52M [E]; Oracle live 2,050; IREN ≤6.5x, Oracle ratio n/d, K04 median 6.0x.", ""]
    for g, files, locks in GROUPS:
        md += [f"### v3.1 - {g}", "", "| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |", "|---|---|---:|---|---|"]
        md += [f"| `{r[1]}` | {r[2]} | {r[3]:,} | `{r[4]}` | {'CHANGED' if v3.get(r[1]) != r[4] else 'same'} |" for r in rows if r[0] == g]; md.append("")
    md += ["---", "", old]
    open(f"{P}/FREEZE.md", "w", encoding="utf-8").write("\n".join(md))
    open(f"{FZ}/SHA256SUMS", "w").write("\n".join(sums) + "\n"); shutil.copy2(f"{P}/FREEZE.md", f"{FZ}/FREEZE.md")
    for f in ("SHA256SUMS", "FREEZE.md"): os.chmod(f"{FZ}/{f}", 0o444)
    for dp, dn, fn in sorted(os.walk(FZ), key=lambda x: -len(x[0])): os.chmod(dp, 0o555)
    print(len(rows), "files frozen (v3.1) + RECONCILE.md;", len(changed), "changed vs v3:", [r[1] for r in changed]); sys.exit(0)
if V3:
    v2 = {l.split("  ", 1)[1]: l.split("  ", 1)[0] for l in open(f"{FZ2}/SHA256SUMS").read().splitlines() if l.strip()}
    changed = [r for r in rows if v2.get(r[1]) != r[4]]
    rec = open(f"{FZ2}/RECONCILE.md", encoding="utf-8").read().rstrip("\n") + "\n\n" + open(f"{P}/reconcile_v3_note.md", encoding="utf-8").read()
    open(f"{FZ}/RECONCILE.md", "w", encoding="utf-8").write(rec); os.chmod(f"{FZ}/RECONCILE.md", 0o444); rec_h = sha(f"{FZ}/RECONCILE.md")
    sums.append(f"{rec_h}  RECONCILE.md")
    old = open(f"{P}/FREEZE.md", encoding="utf-8").read().split("\n", 1)[1]
    old = old.replace("## CURRENT: v2 - frozen", "## SUPERSEDED: v2 - frozen", 1).replace("## SUPERSEDED: v1", "### SUPERSEDED: v1", 1)
    md = ["# FREEZE - peer deliverables for Sunday Oct 11, 2026 finals", "",
          f"## CURRENT: v3 - frozen {now} (supersedes v2 and v1)", "",
          f"{len(rows)} files in {len(GROUPS)} groups ({sum(r[3] for r in rows):,} bytes) + RECONCILE.md; {len(changed)} changed vs v2. Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11_v3/` + `SHA256SUMS` (covers all {len(rows)} files AND RECONCILE.md). "
          "v1 (`frozen_2026-10-11/`) and v2 (`frozen_2026-10-11_v2/`) are left untouched (superseded). Verify: `cd peers/frozen_2026-10-11_v3 && sha256sum -c SHA256SUMS`.", "",
          "### What changed in v3 and why (coordinator, Oct 9 ~4:36 PM ET; RECONCILE.md items 1-2 + side note)", "",
          "1. **IREN contracted/live** -> \"≤6.5x [D]\" (upper bound; not in any median). Live ~40 MW [R, Jun 30, 2026, FY26 10-K], company-wide, excludes Horizon 1 (50 IT MW, accepted Aug 13); Childress ~75 MW is one site, not used. Appendix table/xlsx (U7 tag 'D (upper bound: <=)'), footnote 8, chart_contracted_vs_live_mw.png label; comparison K04 IREN value stays n/d with the wording as a cell note (J8) and in Chart_Data A35.",
          "2. **Oracle contracted/live** -> n/d (was >4.9x in the appendix). Contracted shown as \">10 GW [R, FQ1 FY27 call Sep 10]\" (appendix O8 text '>10 GW'; chart label '>10 GW', bar drawn at a 10,000 floor); footnote 7 = RECONCILE wording; comparison K04 Oracle stays n/d + cell note (M8) + Chart_Data A35; charts/active_vs_contracted_mw.png labels '>10 GW'.",
          "3. **comparison.xlsx Meta #1 note**: Contracts_$MWyr M6 'E (LOW): unsourced secondary ~50 MW' -> 'E: ~50 IT MW team estimate (model Contracts!D23:D24 = 48.0 MW; Capacity tranches list n/d)'; README note A52 reworded to match. The memo summary table copies that field (Contracts_$MWyr L6) and was rebuilt (1 cell).",
          "**Metric / median moves: none.** K04 contracted/live median stays 6.0x (CRWV 2.47 / Nscale 16.91 / Crusoe 6.0 - IREN and Oracle were never in the set); K01 median, K13 33.18, K14 unchanged; band $9.7–16.7M unchanged; all locks re-asserted.",
          "Not rebuilt / unchanged vs v2: memo docx/pdf/fig (no Meta #1 note, no IREN/Oracle ratio), summary-table PNG, price_or_mw, site appendices, benchmark (shows no IREN/Oracle ratio; its rebuild re-passed all checks with identical PDF/PNG - v2 bytes restored, see note below).",
          "Note: GPU Market updated market/contract_economics.xlsx and the band slide at 4:38 PM ET (the RECONCILE tag fixes). A benchmark rebuild would mirror 6 Comp_Band tag cells (D16, D19-D23) from it; not taken in v3 (outside the approved scope) - coordinator call.", "",
          "### Changed files (v2 -> v3 SHA-256)", "", "| Path (peers/) | v2 SHA-256 | v3 SHA-256 | v3 size |", "|---|---|---|---:|"]
    md += [f"| `{r[1]}` | `{v2.get(r[1], 'new')[:16]}...` | `{r[4][:16]}...` | {r[3]:,} |" for r in changed]
    md += [f"| `RECONCILE.md` (v2 note + 'resolved in v3') | `{sha(f'{FZ2}/RECONCILE.md')[:16]}...` | `{rec_h[:16]}...` | {os.path.getsize(f'{FZ}/RECONCILE.md'):,} |", ""]
    md += ["### Locks (unchanged; re-asserted by every build)", "", f"{L_NEB}; {L_BAND}; {L_K13}; ex-MSFT hurdle price_or_mw B37 27.6278; Nebius-Meta #1 $11.52M [E] = Print_Table #3; Oracle live 2,050 = oracle_kpis Inputs!P48; NEW in v3: IREN ratio 6.5 shown as ≤ (upper bound), Oracle ratio n/d, K04 median 6.0x.", ""]
    for g, files, locks in GROUPS:
        md += [f"### v3 - {g}", "", "| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |", "|---|---|---:|---|---|"]
        md += [f"| `{r[1]}` | {r[2]} | {r[3]:,} | `{r[4]}` | {'CHANGED' if v2.get(r[1]) != r[4] else 'same'} |" for r in rows if r[0] == g]; md.append("")
    md += ["---", "", old]
    open(f"{P}/FREEZE.md", "w", encoding="utf-8").write("\n".join(md))
    open(f"{FZ}/SHA256SUMS", "w").write("\n".join(sums) + "\n"); shutil.copy2(f"{P}/FREEZE.md", f"{FZ}/FREEZE.md")
    for f in ("SHA256SUMS", "FREEZE.md"): os.chmod(f"{FZ}/{f}", 0o444)
    for dp, dn, fn in sorted(os.walk(FZ), key=lambda x: -len(x[0])): os.chmod(dp, 0o555)
    print(len(rows), "files frozen (v3) + RECONCILE.md;", len(changed), "changed vs v2:", [r[1] for r in changed]); sys.exit(0)
if V2:
    v1 = {l.split("  ", 1)[1]: l.split("  ", 1)[0] for l in open(f"{FZ1}/SHA256SUMS").read().splitlines() if l.strip()}
    v1md = open(f"{FZ1}/FREEZE.md", encoding="utf-8").read().split("\n", 1)[1]
    changed = [r for r in rows if v1.get(r[1]) != r[4]]
    md = ["# FREEZE - peer deliverables for Sunday Oct 11, 2026 finals", "",
          f"## CURRENT: v2 - frozen {now} (supersedes v1)", "",
          f"{len(rows)} files in {len(GROUPS)} groups ({sum(r[3] for r in rows):,} bytes); {len(changed)} changed vs v1. Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11_v2/` + `SHA256SUMS`. "
          "v1 copies in `peers/frozen_2026-10-11/` are left untouched (superseded). Verify: `cd peers/frozen_2026-10-11_v2 && sha256sum -c SHA256SUMS`.", "",
          "### What changed in v2 and why (coordinator rulings, Oct 9 ~2:36 PM ET)", "",
          "1. **Oracle live MW** in comparison.xlsx (build_comparison.py) 1,200 (May-26) -> ~2,050 (Aug-26): cumulative AI MW delivered since Jun-25, tag E, read from peers/oracle/oracle_kpis.xlsx Inputs!P48 (I45E; FQ1 FY27 call) - the same source/tag as peer_benchmark and price_or_mw. Changes only Chart_Data E33 + its basis note A35 and charts/active_vs_contracted_mw.png. No KPI, median, flag or ranking moves: Oracle K04 stays n/d (contracted is a '>10 GW' lower bound), K13 (46.39, FY26 capex / FY26 delivered MW) is period-matched and not fed by this cell; the 6.0x contracted/live median (CRWV/Nscale/Crusoe) and the 33.18 capex/MW median are unchanged.",
          "2. **Nebius-Meta #1**: kept at $11.52M [E] in comparison.xlsx; peer_benchmark.* and peer_kpi_appendix.* now show \"$11.52M [E, MW is a team estimate of ~50 IT MW]\" (was n/d), sourced to GPU Market market/contract_economics.xlsx Print_Table #3 and the final band slide market/charts/contract_band_slide_final.png, and to Capacity (capacity/nebius_capacity.xlsx 'Capacity tranches' G3:H3 = n/d - Capacity sizes no Meta #1 MW; the team MW estimate is model Contracts!D23+D24 = 48.0 IT MW). One of the eight band prints; band ends stay $9.7M / $16.7M (asserted on both the memo set and the eight prints). Benchmark footnotes set 6.0 -> 5.8 pt to hold 1 page.",
          "3-4. IREN live MW vintage (~40 Jun-30 vs Childress ~75 Oct) and Oracle contracted '>10 GW' lower bound: unchanged, footnoted (appendix fn 8 and 7).",
          "Memo, summary table, price_or_mw and site appendices: inputs unchanged (memo reads no Chart_Data cell; memo already shows Meta #1 11.5 [MW E]) -> not rebuilt; hashes identical to v1.", "",
          "### Changed files (v1 -> v2 SHA-256)", "", "| Path (peers/) | v1 SHA-256 | v2 SHA-256 | v2 size |", "|---|---|---|---:|"]
    md += [f"| `{r[1]}` | `{v1.get(r[1], 'new')[:16]}...` | `{r[4][:16]}...` | {r[3]:,} |" for r in changed] + [""]
    md += ["### Locks (unchanged from v1; re-asserted by every build)", "",
           "MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M (Drivers!B9 = price_or_mw B37 27.6278); $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D] (Drivers!B10); clear-spot $24.51M / $22.68M; turn-positive $17.05M / $16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; K13 77.9 vs median 33.18; band $9.7–16.7M; Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; NEW in v2: Nebius-Meta #1 $11.52M [E] = Print_Table #3; Oracle live 2,050 = oracle_kpis Inputs!P48.", ""]
    for g, files, locks in GROUPS:
        md += [f"### v2 - {g}", "", "| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |", "|---|---|---:|---|---|"]
        md += [f"| `{r[1]}` | {r[2]} | {r[3]:,} | `{r[4]}` | {'CHANGED' if v1.get(r[1]) != r[4] else 'same'} |" for r in rows if r[0] == g]; md.append("")
    md += ["---", "", "## SUPERSEDED: v1 (kept for audit; copies in peers/frozen_2026-10-11/, untouched)", "", v1md.replace("\n## ", "\n### v1 - ")]
    open(f"{P}/FREEZE.md", "w", encoding="utf-8").write("\n".join(md))
    open(f"{FZ}/SHA256SUMS", "w").write("\n".join(sums) + "\n"); shutil.copy2(f"{P}/FREEZE.md", f"{FZ}/FREEZE.md")
    for f in ("SHA256SUMS", "FREEZE.md"): os.chmod(f"{FZ}/{f}", 0o444)
    for dp, dn, fn in sorted(os.walk(FZ), key=lambda x: -len(x[0])): os.chmod(dp, 0o555)
    print(len(rows), "files frozen (v2);", len(changed), "changed vs v1:", [r[1] for r in changed]); sys.exit(0)
md = [f"# FREEZE - peer deliverables for Sunday Oct 11, 2026 finals", "",
      f"Frozen {now} by the Nebius Peer KPIs workstream. {len(rows)} files in {len(GROUPS)} deliverable groups; total {sum(r[3] for r in rows):,} bytes.",
      f"Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11/` (same relative paths) + `SHA256SUMS` there. Originals left in place, unchanged.",
      "Verify: `cd /workspace/nebius/peers/frozen_2026-10-11 && sha256sum -c SHA256SUMS` (frozen copies) or run it from `peers/` to check the originals have not moved.",
      "Model: model/nebius_model.xlsx v0.9e (read-only, not copied: owned by the model workstream). Nebius Q3'26 results date: ~Nov 10 (est.), not announced.", "",
      "## Locks every Sunday deliverable relies on (model v0.9e)", "",
      "| Lock | Value | Source cell |", "|---|---|---|",
      "| Nebius-Microsoft $ per IT MW [E, house], only figure (11.60 / 17.08 HISTORY only) | $14.49M | model Contracts!E12 = 17,392.9 / (5 x 240) |",
      "| Ex-MSFT hurdle (peer comparator) | $27.63M (27.6278) | model Drivers!B9; price_or_mw Nebius_Check!B37 |",
      "| Memo: Street FY27 ÷ our 531.8 billed MW, all customers [D] | $23.13M | model Drivers!B10 |",
      "| Clear-spot base / high | $24.51M / $22.68M | Drivers!B6 / Debate_Grid!A18 |",
      "| Turn-positive base / high | $17.05M / $16.70M | Debate_Grid!A13 / A10 |",
      "| Energized MW Q1 → Q2'26 [E] | 181.5 → 193 | capacity Quarterly MW!M130 / N130 |",
      "| Billed IT MW YE26 base / high | 257.7 / 284.2 | Judge_Slide!B34 / B35 |",
      "| FY27 avg billed IT MW base / high | 531.84 / 659.16 | Drivers!B34 / C34 |",
      "| K13 capex/MW Nebius vs peer median | 77.9 vs 33.18 | comparison.xlsx Comparison K13 |",
      "| Comp band (Nebius excluded) | $9.7–16.7M | comparison.xlsx Contracts_$MWyr P1 / NEW2 |",
      "| Only disclosed-MW print ≥ $20M | Bitdeer A201 ≤ $22.4M (ceiling) | peer_benchmark.xlsx Contracts BD-A201 |", ""]
for g, files, locks in GROUPS:
    md += [f"## {g}", "", f"Relies on: {locks}", "", "| Path (peers/) | mtime | Size (bytes) | SHA-256 |", "|---|---|---:|---|"]
    md += [f"| `{r[1]}` | {r[2]} | {r[3]:,} | `{r[4]}` |" for r in rows if r[0] == g]; md.append("")
md += ["## Known differences between the frozen files (documented, not edited)", "",
       "1. Oracle live MW: comparison.xlsx Chart_Data / active_vs_contracted_mw.png use 1,200 (cumulative AI MW delivered, May-26 vintage, matching its K13 period); peer_benchmark, price_or_mw and the final appendix use ~2,050 (same series, Aug-26). Appendix footnote 7.",
       "2. Nebius-Meta #1: comparison.xlsx Contracts_$MWyr N2 (and Chart_Data / contracts chart) carry 11.52 on an unsourced ~50 MW (E, low confidence); peer_benchmark and the appendix show n/d (filings state no MW). Not a band input. Appendix footnote 11.",
       "3. IREN live MW: company ~40 MW AI operating at Jun-30 (comparison, benchmark) vs Childress site ~75 IT active in Oct (site appendix) - vintage difference; IREN contracted/live 6.5x is computed only in the appendix (comparison K04 = n/d) and kept out of the median.",
       "4. Oracle contracted: comparison Chart_Data shows 10,000; benchmark says 'n/d as filed; >10 GW secured' - same lower bound; the appendix shows '>10,000 [R, lower bound]'.", ""]
open(f"{P}/FREEZE.md", "w", encoding="utf-8").write("\n".join(md))
open(f"{FZ}/SHA256SUMS", "w").write("\n".join(sums) + "\n")
shutil.copy2(f"{P}/FREEZE.md", f"{FZ}/FREEZE.md")
for f in ("SHA256SUMS", "FREEZE.md"): os.chmod(f"{FZ}/{f}", 0o444)
for dp, dn, fn in sorted(os.walk(FZ), key=lambda x: -len(x[0])): os.chmod(dp, 0o555)
print(len(rows), "files frozen;", sum(r[3] for r in rows), "bytes")
