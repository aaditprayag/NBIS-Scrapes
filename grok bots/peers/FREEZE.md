# FREEZE - peer deliverables for Sunday Oct 11, 2026 finals

## CURRENT: v3.1 - FINAL for Sunday - frozen Fri Oct 09, 2026 16:46 ET (supersedes v3, v2, v1)

61 files in 7 groups (7,568,607 bytes) + RECONCILE.md; 1 changed vs v3. Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11_v3.1/` + `SHA256SUMS` (all 61 files AND RECONCILE.md). v1, v2, v3 dirs left untouched (superseded). Verify: `cd peers/frozen_2026-10-11_v3.1 && sha256sum -c SHA256SUMS`.

### What changed in v3.1 and why (coordinator, Oct 9 ~4:44 PM ET; tags only)

peer_benchmark.xlsx rebuilt so Comp_Band picks up GPU Market's 4:38 PM ET weakest-input tags (market/contract_economics.xlsx Print_Table; RECONCILE deck-match fixes): D16 'R / D' -> 'R/D'; D19 'D' -> 'D, term P'; D20 'D on E' -> 'E, house'; D21 'E' -> 'P (IT read E)'; D22 'R / D' -> 'R/D, term P'; D23 'R (secondary) / D' -> 'P'. No other cell (values or formulas, all sheets) differs vs v3; 84 formulas, recalc 0 errors; no value, median, band or lock moved.
peer_benchmark.pdf / .png / .docx: rebuilt output was NOT shipped - kept the v3 bytes. PNG rebuilt byte-identical; PDF differed in 94 bytes, all in CreationDate, trailer /ID and /DocChecksum (extracted text and 150-dpi render pixels identical); docx differed only in zip container metadata (every member part byte-identical).
Re-checks: benchmark build asserts, appendix build (22 lock checks + cell-for-cell vs comparison / peer_benchmark / price_or_mw; outputs content-identical, v3 bytes restored), nebius_locks.py - all pass. Pages 1/2/3.

### Changed files (v3 -> v3.1 SHA-256)

| Path (peers/) | v3 SHA-256 | v3.1 SHA-256 | v3.1 size |
|---|---|---|---:|
| `benchmark/peer_benchmark.xlsx` | `283e9ae05a272a5f...` | `3444d38fe34a15c9...` | 39,690 |
| `RECONCILE.md` (v3 note + v3.1 line) | `e8aa76e2c901129c...` | `523c4f63f88b911a...` | 6,810 |

### Locks (unchanged; re-asserted)

MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18; B37 27.6278; Meta #1 $11.52M [E]; Oracle live 2,050; IREN ≤6.5x, Oracle ratio n/d, K04 median 6.0x.

### v3.1 - Memo (3 pages)

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `memo/nebius_vs_peers_unit_economics.docx` | 2026-10-09 11:10 ET | 252,519 | `d30915104359484252c7912ed177387240b6cf56efef605a8c1e067c9bc19ee4` | same |
| `memo/nebius_vs_peers_unit_economics.pdf` | 2026-10-09 11:10 ET | 386,638 | `81677fb88b5f8280fa65f6664263b5831b4e28a3f65ec6272724dac6a89b839d` | same |
| `memo/fig_unit_economics.png` | 2026-10-09 11:10 ET | 233,859 | `eae8ff4b7a97b4246568e829d9e6637f628b0e1eb9f8702903bda677f437848c` | same |
| `memo/build_memo.py` | 2026-10-09 11:08 ET | 28,236 | `7d15746a1944b52d0c67f05bb75c39c97a13b7d12289c6b73f90a8884899968a` | same |
| `memo/build_figs.py` | 2026-10-09 11:09 ET | 5,177 | `86090bb8b5954e4c1cf2cdd3044c3589b65fad0a94d15cbae3c5d8a58c38ed0c` | same |
| `memo/common_data.py` | 2026-10-09 11:12 ET | 7,179 | `af47c5ed275ffd14fcb9c9ffbb2ea328fca1eb4484431aacb57266fe2b93718c` | same |

### v3.1 - Summary table

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `memo/nebius_vs_peers_summary_table.xlsx` | 2026-10-09 16:38 ET | 18,160 | `bfe4e5f5296b7d4f99f80974229f73e05e66b8aa14a68188aabeb882aa1f3a5b` | same |
| `memo/nebius_vs_peers_summary_table.png` | 2026-10-09 11:10 ET | 631,795 | `a628eaba5ffdf883231bbf9b9f19a19e6bb386a8a778d405bb50ddd285c1c170` | same |
| `memo/build_summary_table.py` | 2026-10-09 11:08 ET | 6,975 | `42876bddb61ea26d486b6b6a252defd4355bf3ef17942dc5d32bbaf35e13b851` | same |

### v3.1 - Comparison

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `comparison.xlsx` | 2026-10-09 16:38 ET | 95,201 | `8da67dd81ed904b4d960bfd802a26f3affb4e645571bc744664f61c611038237` | same |
| `build_comparison.py` | 2026-10-09 16:37 ET | 97,043 | `38b3edaf045e5b157296814727ae992ede3cce369a635fce59ef1d6f3edd6ba9` | same |
| `build_charts.py` | 2026-10-09 16:38 ET | 14,205 | `2797ae2ca265fcc59740bbdee6a8d076fc65f42095bb6234716247027ba5f21d` | same |
| `nebius_locks.py` | 2026-10-08 19:06 ET | 9,332 | `d64303e86b186bc040779f61d730b057227d300e1358af8588c6dea82c584840` | same |
| `charts/active_vs_contracted_mw.png` | 2026-10-09 16:38 ET | 143,666 | `8d82d913d87ff2e8d2fc487fc99b57da9a8ed60281f730e7d6a2418baa9752cc` | same |
| `charts/capex_per_mw.png` | 2026-10-09 16:38 ET | 119,091 | `130b0389019761e34485beec1e076b6e7a3ba902571a11152ccf5a36d929ed13` | same |
| `charts/contracts_usd_per_mw_yr.png` | 2026-10-09 16:38 ET | 210,725 | `e485cd1437d97abf79825ce4ee2d214f3af1ba87f33c6bdbdaed84721c0e367a` | same |
| `charts/cost_of_debt.png` | 2026-10-09 16:38 ET | 100,705 | `959d96b485cfa1866901769db2e70d0b1e836486d98a196f075c8153834b0d34` | same |
| `charts/depreciation_yield_useful_life.png` | 2026-10-09 16:38 ET | 102,098 | `c12f31b2b0a5139d0da3557be0fd8b510fce4a895f6f03471deef0c4e077dd50` | same |
| `charts/nebius_billing_lock.png` | 2026-10-09 16:38 ET | 193,296 | `65503e16139400a042096cef0223c8f981891f41b7f228cd989182c0366e6271` | same |
| `charts/rev_per_active_mw.png` | 2026-10-09 16:38 ET | 160,247 | `3925851967bcd1bc9681819cbe46d1483a1d28149a5519647a1947744489d4ce` | same |
| `charts/rev_per_gpu_hr.png` | 2026-10-09 16:38 ET | 99,326 | `9a4085abbdba873ff426fb082ef6b8519558a11dc9c8aeefe2b32fb56c48f560` | same |
| `charts/rpo_to_annualized_revenue.png` | 2026-10-09 16:38 ET | 69,765 | `14f6546caa01b997b8b626ce1fca6030c650e2fd9d9be8e44b3375221ff7ad56` | same |
| `charts/sites_active_vs_announced_mw.png` | 2026-10-09 16:38 ET | 178,348 | `174b76fec988044efd92c3f90f5172c7410f1f273cac6ded50cce3f83790df63` | same |
| `charts/top_customer_pct.png` | 2026-10-09 16:38 ET | 95,729 | `f2f7c6b4a502df1ca7bc50658303fb2eaa811f37b27506799f9823b782be8384` | same |

### v3.1 - price_or_mw

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `price_or_mw.xlsx` | 2026-10-09 11:12 ET | 27,999 | `3cca6431d3eeb25dcb0781e28b4ed0f38fad48e1855de1729a324e074ce2e5e8` | same |
| `charts/price_or_mw.png` | 2026-10-09 11:12 ET | 194,907 | `070189b92d44feca9f5538e770b08f661189201f26ef5e92263f45e1d8ecc8c0` | same |
| `patch_price_or_mw.py` | 2026-10-09 11:12 ET | 18,755 | `1522642db4c9ea90b48e1811028c678d9a9b8546114640850a8c92e7b506cbde` | same |
| `build_price_or_mw_chart.py` | 2026-10-09 11:12 ET | 3,545 | `2eff7494e144b3deab7a8d3ccde6c35e98d17abf6c994268af610b562cb38330` | same |

### v3.1 - Benchmark

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `benchmark/peer_benchmark.xlsx` | 2026-10-09 16:45 ET | 39,690 | `3444d38fe34a15c9f8e48c99fca485e7a7d82c640c28aa3b4ead039ec25d7cf8` | CHANGED |
| `benchmark/peer_benchmark.docx` | 2026-10-09 14:42 ET | 43,716 | `a59d660aa97790c0ad415f9837f98434634f3cf2dd541649756f2f5270379707` | same |
| `benchmark/peer_benchmark.pdf` | 2026-10-09 14:42 ET | 86,336 | `4bf926ab93f6b1e9116394f2f54f12e9eff6e1497a525f7119ff604ea1ea38e3` | same |
| `benchmark/peer_benchmark.png` | 2026-10-09 14:42 ET | 1,021,355 | `b47bd6c5573c937d32565d3444f86aec4137b48c9c4e0ae5187f7401f010dd8a` | same |
| `benchmark/build_peer_benchmark.py` | 2026-10-09 14:42 ET | 66,323 | `74bf1b8f500b1a54d5afe4f2cb177af09748045a5e5b3eafe5ad4218b9513bde` | same |

### v3.1 - Site appendices

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `sites/appendix_coreweave_denton.docx` | 2026-10-08 19:06 ET | 71,624 | `5a97d13a57b3ae08c2015b3f793ccaa985966d25a5015f9d3d571d7bae80ed16` | same |
| `sites/appendix_coreweave_denton.md` | 2026-10-08 19:06 ET | 6,720 | `5d2cc3126aa4e37d513dd471a2aef4af60cbc9f5fe257bda8cb02e697fdbdc72` | same |
| `sites/appendix_coreweave_denton.pdf` | 2026-10-08 19:06 ET | 114,498 | `84631855d58ae3112b60a0785bc6b6ff5d1530074e0c7373a139c663bb4ff91a` | same |
| `sites/appendix_crusoe_abilene.docx` | 2026-10-08 19:06 ET | 74,709 | `e10eb22ccf70d19d7d417e21bccf2d71b7c1d7801493f7550a9c01a677724c9a` | same |
| `sites/appendix_crusoe_abilene.md` | 2026-10-08 19:06 ET | 7,515 | `5a75ffddb37452f246ecbe75acc69842a3e4b67b64db36211893b7d029a6d383` | same |
| `sites/appendix_crusoe_abilene.pdf` | 2026-10-08 19:06 ET | 116,120 | `d6159bc7191db1ace314cde239ee50f6b2f15d9d09f1b6de08bb6a6c6548e72d` | same |
| `sites/appendix_iren_childress.docx` | 2026-10-08 19:06 ET | 71,251 | `c922aade55d67b6e746b8fa9fd4109373b6fbb570ec1d88cbb45643e0e0f495f` | same |
| `sites/appendix_iren_childress.md` | 2026-10-08 19:06 ET | 6,866 | `457a8685639577b559f038e8a10de648b670893935637f3d8d95f873b7705aa9` | same |
| `sites/appendix_iren_childress.pdf` | 2026-10-08 19:06 ET | 112,647 | `641067c1cbf9927969b2d42d40e46d954160a6518381709e90f21449e3776cc4` | same |
| `sites/appendix_oracle_port_washington.docx` | 2026-10-08 19:06 ET | 68,991 | `81741b4813bd4c5d35c9b5aff2ac9eedff9c31cd5068f1f5da895d418a7ade91` | same |
| `sites/appendix_oracle_port_washington.md` | 2026-10-08 19:06 ET | 6,433 | `ca50e60291fa3644a2d11068d91a398b91c5d477ed06ad930a149ba741240a43` | same |
| `sites/appendix_oracle_port_washington.pdf` | 2026-10-08 19:06 ET | 110,443 | `ea8ea8a92b78ab6f1df8446ae686a6db30b319fb6a6991b5f360706cfcde7c39` | same |
| `sites/peer_site_appendices.pdf` | 2026-10-08 19:06 ET | 281,084 | `546103829b4a5c7851373d05029b2775e969aa22d31501d2cab433c6b40f6ee0` | same |
| `sites/sites.xlsx` | 2026-10-08 23:12 ET | 62,723 | `49f9069a9e9d0dc1c916bfa8865fc049efc63b77a0504b975960c536805ee833` | same |
| `sites/site_verification.md` | 2026-10-08 23:12 ET | 28,470 | `46bcec41af24a744c60662ad6d46283da64bc883b84b1bc4edcb3d2b2ec52be2` | same |
| `sites/build_site_appendices.py` | 2026-10-08 19:06 ET | 9,826 | `c26a99ff4cefbfcf1f74d11d4501b7dc0a5ee23d87a9602392916a4d9ce0b55a` | same |
| `sites/site_data.py` | 2026-10-08 19:06 ET | 25,146 | `b78d97394ac3f091ff5b1bc275f99498d21f8750462983edb26ba99980b2fa80` | same |
| `sites/charts/coreweave_denton_mw.png` | 2026-10-08 19:06 ET | 35,242 | `ce8ff543c59c8fd46fffe2c67942f90867d111bd0c56a955ff2f10d54bcfef3e` | same |
| `sites/charts/crusoe_abilene_mw.png` | 2026-10-08 19:06 ET | 37,895 | `ab1b37d57273c0c20b75854594c00bbfbfb9add70dd505b72b5eab0983710912` | same |
| `sites/charts/iren_childress_mw.png` | 2026-10-08 19:06 ET | 34,477 | `e859642c8938788f9b0b043ff437f80a3af82f4d795e050b062c08a2fff612e3` | same |
| `sites/charts/oracle_port_washington_mw.png` | 2026-10-08 19:06 ET | 32,618 | `7dd2bc8b5774748af038b09fe302373b54666c0ce3fd5d20bd9d3f721efa86e8` | same |

### v3.1 - Final appendix pack (this build)

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v3 |
|---|---|---:|---|---|
| `final_appendix/peer_kpi_appendix.xlsx` | 2026-10-09 16:39 ET | 20,315 | `dbe368b104e40054a78cd07d0271c5d91ebd2f56b83ea1e9dc3b10a0f324ec85` | same |
| `final_appendix/peer_kpi_appendix.docx` | 2026-10-09 16:39 ET | 469,187 | `0bbab551931b39ce0e3f9a076864e5c0a6faa3fe0bdf2204065f0b1f1096176a` | same |
| `final_appendix/peer_kpi_appendix.pdf` | 2026-10-09 16:39 ET | 467,276 | `5f081b1f5fe358181dd3cd3c501fa3f07b6314d064d70fda9eedf6a4034bb1c5` | same |
| `final_appendix/chart_usd_per_mw_yr.png` | 2026-10-09 16:39 ET | 250,891 | `0c550f2ec8b76fa1518d77b98b66b22c5cfd5229baf98855b2ba3c0c9126725b` | same |
| `final_appendix/chart_contracted_vs_live_mw.png` | 2026-10-09 16:39 ET | 135,744 | `9ae1408660aa95386f69bd1d14bcfff016e960d59864af8ac5aab4c02a2cfc4c` | same |
| `final_appendix/chart_capex_per_mw.png` | 2026-10-09 16:39 ET | 93,189 | `54fa95105066769168ce40e83655f00f40961db5bc1221a8a11e21e431329a2e` | same |
| `final_appendix/build_peer_kpi_appendix.py` | 2026-10-09 16:39 ET | 54,766 | `6f4eda2954a1e5895e216848607c7b7b9b2b5658edd73d53c5f7846f1e694a31` | same |

---


## SUPERSEDED: v3 - frozen Fri Oct 09, 2026 16:41 ET (supersedes v2 and v1)

61 files in 7 groups (7,568,594 bytes) + RECONCILE.md; 10 changed vs v2. Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11_v3/` + `SHA256SUMS` (covers all 61 files AND RECONCILE.md). v1 (`frozen_2026-10-11/`) and v2 (`frozen_2026-10-11_v2/`) are left untouched (superseded). Verify: `cd peers/frozen_2026-10-11_v3 && sha256sum -c SHA256SUMS`.

### What changed in v3 and why (coordinator, Oct 9 ~4:36 PM ET; RECONCILE.md items 1-2 + side note)

1. **IREN contracted/live** -> "≤6.5x [D]" (upper bound; not in any median). Live ~40 MW [R, Jun 30, 2026, FY26 10-K], company-wide, excludes Horizon 1 (50 IT MW, accepted Aug 13); Childress ~75 MW is one site, not used. Appendix table/xlsx (U7 tag 'D (upper bound: <=)'), footnote 8, chart_contracted_vs_live_mw.png label; comparison K04 IREN value stays n/d with the wording as a cell note (J8) and in Chart_Data A35.
2. **Oracle contracted/live** -> n/d (was >4.9x in the appendix). Contracted shown as ">10 GW [R, FQ1 FY27 call Sep 10]" (appendix O8 text '>10 GW'; chart label '>10 GW', bar drawn at a 10,000 floor); footnote 7 = RECONCILE wording; comparison K04 Oracle stays n/d + cell note (M8) + Chart_Data A35; charts/active_vs_contracted_mw.png labels '>10 GW'.
3. **comparison.xlsx Meta #1 note**: Contracts_$MWyr M6 'E (LOW): unsourced secondary ~50 MW' -> 'E: ~50 IT MW team estimate (model Contracts!D23:D24 = 48.0 MW; Capacity tranches list n/d)'; README note A52 reworded to match. The memo summary table copies that field (Contracts_$MWyr L6) and was rebuilt (1 cell).
**Metric / median moves: none.** K04 contracted/live median stays 6.0x (CRWV 2.47 / Nscale 16.91 / Crusoe 6.0 - IREN and Oracle were never in the set); K01 median, K13 33.18, K14 unchanged; band $9.7–16.7M unchanged; all locks re-asserted.
Not rebuilt / unchanged vs v2: memo docx/pdf/fig (no Meta #1 note, no IREN/Oracle ratio), summary-table PNG, price_or_mw, site appendices, benchmark (shows no IREN/Oracle ratio; its rebuild re-passed all checks with identical PDF/PNG - v2 bytes restored, see note below).
Note: GPU Market updated market/contract_economics.xlsx and the band slide at 4:38 PM ET (the RECONCILE tag fixes). A benchmark rebuild would mirror 6 Comp_Band tag cells (D16, D19-D23) from it; not taken in v3 (outside the approved scope) - coordinator call.

### Changed files (v2 -> v3 SHA-256)

| Path (peers/) | v2 SHA-256 | v3 SHA-256 | v3 size |
|---|---|---|---:|
| `memo/nebius_vs_peers_summary_table.xlsx` | `b321a0f1a0c1f79b...` | `bfe4e5f5296b7d4f...` | 18,160 |
| `comparison.xlsx` | `b1c535b0fd3fc382...` | `8da67dd81ed904b4...` | 95,201 |
| `build_comparison.py` | `e3edb2ed15916364...` | `38b3edaf045e5b15...` | 97,043 |
| `build_charts.py` | `3f605119f4a58d95...` | `2797ae2ca265fcc5...` | 14,205 |
| `charts/active_vs_contracted_mw.png` | `07dbba3e753e5dca...` | `8d82d913d87ff2e8...` | 143,666 |
| `final_appendix/peer_kpi_appendix.xlsx` | `a8e009efaf31b634...` | `dbe368b104e40054...` | 20,315 |
| `final_appendix/peer_kpi_appendix.docx` | `ec15b27534e610a2...` | `0bbab551931b39ce...` | 469,187 |
| `final_appendix/peer_kpi_appendix.pdf` | `2c56abbf921c87f3...` | `5f081b1f5fe35818...` | 467,276 |
| `final_appendix/chart_contracted_vs_live_mw.png` | `f999b4644f8f05ce...` | `9ae1408660aa9538...` | 135,744 |
| `final_appendix/build_peer_kpi_appendix.py` | `1c6718ac52a2f009...` | `6f4eda2954a1e589...` | 54,766 |
| `RECONCILE.md` (v2 note + 'resolved in v3') | `99fb62b62832924b...` | `e8aa76e2c901129c...` | 6,469 |

### Locks (unchanged; re-asserted by every build)

MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18; ex-MSFT hurdle price_or_mw B37 27.6278; Nebius-Meta #1 $11.52M [E] = Print_Table #3; Oracle live 2,050 = oracle_kpis Inputs!P48; NEW in v3: IREN ratio 6.5 shown as ≤ (upper bound), Oracle ratio n/d, K04 median 6.0x.

### v3 - Memo (3 pages)

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `memo/nebius_vs_peers_unit_economics.docx` | 2026-10-09 11:10 ET | 252,519 | `d30915104359484252c7912ed177387240b6cf56efef605a8c1e067c9bc19ee4` | same |
| `memo/nebius_vs_peers_unit_economics.pdf` | 2026-10-09 11:10 ET | 386,638 | `81677fb88b5f8280fa65f6664263b5831b4e28a3f65ec6272724dac6a89b839d` | same |
| `memo/fig_unit_economics.png` | 2026-10-09 11:10 ET | 233,859 | `eae8ff4b7a97b4246568e829d9e6637f628b0e1eb9f8702903bda677f437848c` | same |
| `memo/build_memo.py` | 2026-10-09 11:08 ET | 28,236 | `7d15746a1944b52d0c67f05bb75c39c97a13b7d12289c6b73f90a8884899968a` | same |
| `memo/build_figs.py` | 2026-10-09 11:09 ET | 5,177 | `86090bb8b5954e4c1cf2cdd3044c3589b65fad0a94d15cbae3c5d8a58c38ed0c` | same |
| `memo/common_data.py` | 2026-10-09 11:12 ET | 7,179 | `af47c5ed275ffd14fcb9c9ffbb2ea328fca1eb4484431aacb57266fe2b93718c` | same |

### v3 - Summary table

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `memo/nebius_vs_peers_summary_table.xlsx` | 2026-10-09 16:38 ET | 18,160 | `bfe4e5f5296b7d4f99f80974229f73e05e66b8aa14a68188aabeb882aa1f3a5b` | CHANGED |
| `memo/nebius_vs_peers_summary_table.png` | 2026-10-09 11:10 ET | 631,795 | `a628eaba5ffdf883231bbf9b9f19a19e6bb386a8a778d405bb50ddd285c1c170` | same |
| `memo/build_summary_table.py` | 2026-10-09 11:08 ET | 6,975 | `42876bddb61ea26d486b6b6a252defd4355bf3ef17942dc5d32bbaf35e13b851` | same |

### v3 - Comparison

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `comparison.xlsx` | 2026-10-09 16:38 ET | 95,201 | `8da67dd81ed904b4d960bfd802a26f3affb4e645571bc744664f61c611038237` | CHANGED |
| `build_comparison.py` | 2026-10-09 16:37 ET | 97,043 | `38b3edaf045e5b157296814727ae992ede3cce369a635fce59ef1d6f3edd6ba9` | CHANGED |
| `build_charts.py` | 2026-10-09 16:38 ET | 14,205 | `2797ae2ca265fcc59740bbdee6a8d076fc65f42095bb6234716247027ba5f21d` | CHANGED |
| `nebius_locks.py` | 2026-10-08 19:06 ET | 9,332 | `d64303e86b186bc040779f61d730b057227d300e1358af8588c6dea82c584840` | same |
| `charts/active_vs_contracted_mw.png` | 2026-10-09 16:38 ET | 143,666 | `8d82d913d87ff2e8d2fc487fc99b57da9a8ed60281f730e7d6a2418baa9752cc` | CHANGED |
| `charts/capex_per_mw.png` | 2026-10-09 16:38 ET | 119,091 | `130b0389019761e34485beec1e076b6e7a3ba902571a11152ccf5a36d929ed13` | same |
| `charts/contracts_usd_per_mw_yr.png` | 2026-10-09 16:38 ET | 210,725 | `e485cd1437d97abf79825ce4ee2d214f3af1ba87f33c6bdbdaed84721c0e367a` | same |
| `charts/cost_of_debt.png` | 2026-10-09 16:38 ET | 100,705 | `959d96b485cfa1866901769db2e70d0b1e836486d98a196f075c8153834b0d34` | same |
| `charts/depreciation_yield_useful_life.png` | 2026-10-09 16:38 ET | 102,098 | `c12f31b2b0a5139d0da3557be0fd8b510fce4a895f6f03471deef0c4e077dd50` | same |
| `charts/nebius_billing_lock.png` | 2026-10-09 16:38 ET | 193,296 | `65503e16139400a042096cef0223c8f981891f41b7f228cd989182c0366e6271` | same |
| `charts/rev_per_active_mw.png` | 2026-10-09 16:38 ET | 160,247 | `3925851967bcd1bc9681819cbe46d1483a1d28149a5519647a1947744489d4ce` | same |
| `charts/rev_per_gpu_hr.png` | 2026-10-09 16:38 ET | 99,326 | `9a4085abbdba873ff426fb082ef6b8519558a11dc9c8aeefe2b32fb56c48f560` | same |
| `charts/rpo_to_annualized_revenue.png` | 2026-10-09 16:38 ET | 69,765 | `14f6546caa01b997b8b626ce1fca6030c650e2fd9d9be8e44b3375221ff7ad56` | same |
| `charts/sites_active_vs_announced_mw.png` | 2026-10-09 16:38 ET | 178,348 | `174b76fec988044efd92c3f90f5172c7410f1f273cac6ded50cce3f83790df63` | same |
| `charts/top_customer_pct.png` | 2026-10-09 16:38 ET | 95,729 | `f2f7c6b4a502df1ca7bc50658303fb2eaa811f37b27506799f9823b782be8384` | same |

### v3 - price_or_mw

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `price_or_mw.xlsx` | 2026-10-09 11:12 ET | 27,999 | `3cca6431d3eeb25dcb0781e28b4ed0f38fad48e1855de1729a324e074ce2e5e8` | same |
| `charts/price_or_mw.png` | 2026-10-09 11:12 ET | 194,907 | `070189b92d44feca9f5538e770b08f661189201f26ef5e92263f45e1d8ecc8c0` | same |
| `patch_price_or_mw.py` | 2026-10-09 11:12 ET | 18,755 | `1522642db4c9ea90b48e1811028c678d9a9b8546114640850a8c92e7b506cbde` | same |
| `build_price_or_mw_chart.py` | 2026-10-09 11:12 ET | 3,545 | `2eff7494e144b3deab7a8d3ccde6c35e98d17abf6c994268af610b562cb38330` | same |

### v3 - Benchmark

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `benchmark/peer_benchmark.xlsx` | 2026-10-09 14:42 ET | 39,677 | `283e9ae05a272a5f033b56ce29fe6fda582602db892531a4353971e3c009b8b6` | same |
| `benchmark/peer_benchmark.docx` | 2026-10-09 14:42 ET | 43,716 | `a59d660aa97790c0ad415f9837f98434634f3cf2dd541649756f2f5270379707` | same |
| `benchmark/peer_benchmark.pdf` | 2026-10-09 14:42 ET | 86,336 | `4bf926ab93f6b1e9116394f2f54f12e9eff6e1497a525f7119ff604ea1ea38e3` | same |
| `benchmark/peer_benchmark.png` | 2026-10-09 14:42 ET | 1,021,355 | `b47bd6c5573c937d32565d3444f86aec4137b48c9c4e0ae5187f7401f010dd8a` | same |
| `benchmark/build_peer_benchmark.py` | 2026-10-09 14:42 ET | 66,323 | `74bf1b8f500b1a54d5afe4f2cb177af09748045a5e5b3eafe5ad4218b9513bde` | same |

### v3 - Site appendices

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `sites/appendix_coreweave_denton.docx` | 2026-10-08 19:06 ET | 71,624 | `5a97d13a57b3ae08c2015b3f793ccaa985966d25a5015f9d3d571d7bae80ed16` | same |
| `sites/appendix_coreweave_denton.md` | 2026-10-08 19:06 ET | 6,720 | `5d2cc3126aa4e37d513dd471a2aef4af60cbc9f5fe257bda8cb02e697fdbdc72` | same |
| `sites/appendix_coreweave_denton.pdf` | 2026-10-08 19:06 ET | 114,498 | `84631855d58ae3112b60a0785bc6b6ff5d1530074e0c7373a139c663bb4ff91a` | same |
| `sites/appendix_crusoe_abilene.docx` | 2026-10-08 19:06 ET | 74,709 | `e10eb22ccf70d19d7d417e21bccf2d71b7c1d7801493f7550a9c01a677724c9a` | same |
| `sites/appendix_crusoe_abilene.md` | 2026-10-08 19:06 ET | 7,515 | `5a75ffddb37452f246ecbe75acc69842a3e4b67b64db36211893b7d029a6d383` | same |
| `sites/appendix_crusoe_abilene.pdf` | 2026-10-08 19:06 ET | 116,120 | `d6159bc7191db1ace314cde239ee50f6b2f15d9d09f1b6de08bb6a6c6548e72d` | same |
| `sites/appendix_iren_childress.docx` | 2026-10-08 19:06 ET | 71,251 | `c922aade55d67b6e746b8fa9fd4109373b6fbb570ec1d88cbb45643e0e0f495f` | same |
| `sites/appendix_iren_childress.md` | 2026-10-08 19:06 ET | 6,866 | `457a8685639577b559f038e8a10de648b670893935637f3d8d95f873b7705aa9` | same |
| `sites/appendix_iren_childress.pdf` | 2026-10-08 19:06 ET | 112,647 | `641067c1cbf9927969b2d42d40e46d954160a6518381709e90f21449e3776cc4` | same |
| `sites/appendix_oracle_port_washington.docx` | 2026-10-08 19:06 ET | 68,991 | `81741b4813bd4c5d35c9b5aff2ac9eedff9c31cd5068f1f5da895d418a7ade91` | same |
| `sites/appendix_oracle_port_washington.md` | 2026-10-08 19:06 ET | 6,433 | `ca50e60291fa3644a2d11068d91a398b91c5d477ed06ad930a149ba741240a43` | same |
| `sites/appendix_oracle_port_washington.pdf` | 2026-10-08 19:06 ET | 110,443 | `ea8ea8a92b78ab6f1df8446ae686a6db30b319fb6a6991b5f360706cfcde7c39` | same |
| `sites/peer_site_appendices.pdf` | 2026-10-08 19:06 ET | 281,084 | `546103829b4a5c7851373d05029b2775e969aa22d31501d2cab433c6b40f6ee0` | same |
| `sites/sites.xlsx` | 2026-10-08 23:12 ET | 62,723 | `49f9069a9e9d0dc1c916bfa8865fc049efc63b77a0504b975960c536805ee833` | same |
| `sites/site_verification.md` | 2026-10-08 23:12 ET | 28,470 | `46bcec41af24a744c60662ad6d46283da64bc883b84b1bc4edcb3d2b2ec52be2` | same |
| `sites/build_site_appendices.py` | 2026-10-08 19:06 ET | 9,826 | `c26a99ff4cefbfcf1f74d11d4501b7dc0a5ee23d87a9602392916a4d9ce0b55a` | same |
| `sites/site_data.py` | 2026-10-08 19:06 ET | 25,146 | `b78d97394ac3f091ff5b1bc275f99498d21f8750462983edb26ba99980b2fa80` | same |
| `sites/charts/coreweave_denton_mw.png` | 2026-10-08 19:06 ET | 35,242 | `ce8ff543c59c8fd46fffe2c67942f90867d111bd0c56a955ff2f10d54bcfef3e` | same |
| `sites/charts/crusoe_abilene_mw.png` | 2026-10-08 19:06 ET | 37,895 | `ab1b37d57273c0c20b75854594c00bbfbfb9add70dd505b72b5eab0983710912` | same |
| `sites/charts/iren_childress_mw.png` | 2026-10-08 19:06 ET | 34,477 | `e859642c8938788f9b0b043ff437f80a3af82f4d795e050b062c08a2fff612e3` | same |
| `sites/charts/oracle_port_washington_mw.png` | 2026-10-08 19:06 ET | 32,618 | `7dd2bc8b5774748af038b09fe302373b54666c0ce3fd5d20bd9d3f721efa86e8` | same |

### v3 - Final appendix pack (this build)

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v2 |
|---|---|---:|---|---|
| `final_appendix/peer_kpi_appendix.xlsx` | 2026-10-09 16:39 ET | 20,315 | `dbe368b104e40054a78cd07d0271c5d91ebd2f56b83ea1e9dc3b10a0f324ec85` | CHANGED |
| `final_appendix/peer_kpi_appendix.docx` | 2026-10-09 16:39 ET | 469,187 | `0bbab551931b39ce0e3f9a076864e5c0a6faa3fe0bdf2204065f0b1f1096176a` | CHANGED |
| `final_appendix/peer_kpi_appendix.pdf` | 2026-10-09 16:39 ET | 467,276 | `5f081b1f5fe358181dd3cd3c501fa3f07b6314d064d70fda9eedf6a4034bb1c5` | CHANGED |
| `final_appendix/chart_usd_per_mw_yr.png` | 2026-10-09 16:39 ET | 250,891 | `0c550f2ec8b76fa1518d77b98b66b22c5cfd5229baf98855b2ba3c0c9126725b` | same |
| `final_appendix/chart_contracted_vs_live_mw.png` | 2026-10-09 16:39 ET | 135,744 | `9ae1408660aa95386f69bd1d14bcfff016e960d59864af8ac5aab4c02a2cfc4c` | CHANGED |
| `final_appendix/chart_capex_per_mw.png` | 2026-10-09 16:39 ET | 93,189 | `54fa95105066769168ce40e83655f00f40961db5bc1221a8a11e21e431329a2e` | same |
| `final_appendix/build_peer_kpi_appendix.py` | 2026-10-09 16:39 ET | 54,766 | `6f4eda2954a1e5895e216848607c7b7b9b2b5658edd73d53c5f7846f1e694a31` | CHANGED |

---


## SUPERSEDED: v2 - frozen Fri Oct 09, 2026 14:44 ET (supersedes v1)

61 files in 7 groups (7,529,647 bytes); 13 changed vs v1. Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11_v2/` + `SHA256SUMS`. v1 copies in `peers/frozen_2026-10-11/` are left untouched (superseded). Verify: `cd peers/frozen_2026-10-11_v2 && sha256sum -c SHA256SUMS`.

### What changed in v2 and why (coordinator rulings, Oct 9 ~2:36 PM ET)

1. **Oracle live MW** in comparison.xlsx (build_comparison.py) 1,200 (May-26) -> ~2,050 (Aug-26): cumulative AI MW delivered since Jun-25, tag E, read from peers/oracle/oracle_kpis.xlsx Inputs!P48 (I45E; FQ1 FY27 call) - the same source/tag as peer_benchmark and price_or_mw. Changes only Chart_Data E33 + its basis note A35 and charts/active_vs_contracted_mw.png. No KPI, median, flag or ranking moves: Oracle K04 stays n/d (contracted is a '>10 GW' lower bound), K13 (46.39, FY26 capex / FY26 delivered MW) is period-matched and not fed by this cell; the 6.0x contracted/live median (CRWV/Nscale/Crusoe) and the 33.18 capex/MW median are unchanged.
2. **Nebius-Meta #1**: kept at $11.52M [E] in comparison.xlsx; peer_benchmark.* and peer_kpi_appendix.* now show "$11.52M [E, MW is a team estimate of ~50 IT MW]" (was n/d), sourced to GPU Market market/contract_economics.xlsx Print_Table #3 and the final band slide market/charts/contract_band_slide_final.png, and to Capacity (capacity/nebius_capacity.xlsx 'Capacity tranches' G3:H3 = n/d - Capacity sizes no Meta #1 MW; the team MW estimate is model Contracts!D23+D24 = 48.0 IT MW). One of the eight band prints; band ends stay $9.7M / $16.7M (asserted on both the memo set and the eight prints). Benchmark footnotes set 6.0 -> 5.8 pt to hold 1 page.
3-4. IREN live MW vintage (~40 Jun-30 vs Childress ~75 Oct) and Oracle contracted '>10 GW' lower bound: unchanged, footnoted (appendix fn 8 and 7).
Memo, summary table, price_or_mw and site appendices: inputs unchanged (memo reads no Chart_Data cell; memo already shows Meta #1 11.5 [MW E]) -> not rebuilt; hashes identical to v1.

### Changed files (v1 -> v2 SHA-256)

| Path (peers/) | v1 SHA-256 | v2 SHA-256 | v2 size |
|---|---|---|---:|
| `comparison.xlsx` | `849ac1803f0e6d55...` | `b1c535b0fd3fc382...` | 93,145 |
| `build_comparison.py` | `2a8fdeb0bdceb102...` | `e3edb2ed15916364...` | 95,579 |
| `charts/active_vs_contracted_mw.png` | `ec2a6619595d8665...` | `07dbba3e753e5dca...` | 117,641 |
| `benchmark/peer_benchmark.xlsx` | `4b605144edac2c51...` | `283e9ae05a272a5f...` | 39,677 |
| `benchmark/peer_benchmark.docx` | `a6b17f92c2775cac...` | `a59d660aa97790c0...` | 43,716 |
| `benchmark/peer_benchmark.pdf` | `8684bf521477fa13...` | `4bf926ab93f6b1e9...` | 86,336 |
| `benchmark/peer_benchmark.png` | `c2364c72d99b8ff2...` | `b47bd6c5573c937d...` | 1,021,355 |
| `benchmark/build_peer_benchmark.py` | `e44da74fff2e1a5c...` | `74bf1b8f500b1a54...` | 66,323 |
| `final_appendix/peer_kpi_appendix.xlsx` | `e718a53c73436705...` | `a8e009efaf31b634...` | 20,242 |
| `final_appendix/peer_kpi_appendix.docx` | `629d6fbc6c7dacf9...` | `ec15b27534e610a2...` | 466,001 |
| `final_appendix/peer_kpi_appendix.pdf` | `5983c7c78770f941...` | `2c56abbf921c87f3...` | 463,217 |
| `final_appendix/chart_usd_per_mw_yr.png` | `39eda5f1655d5ddb...` | `0c550f2ec8b76fa1...` | 250,891 |
| `final_appendix/build_peer_kpi_appendix.py` | `e2b5e20224f54fe8...` | `1c6718ac52a2f009...` | 53,588 |

### Locks (unchanged from v1; re-asserted by every build)

MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M (Drivers!B9 = price_or_mw B37 27.6278); $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D] (Drivers!B10); clear-spot $24.51M / $22.68M; turn-positive $17.05M / $16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; K13 77.9 vs median 33.18; band $9.7–16.7M; Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; NEW in v2: Nebius-Meta #1 $11.52M [E] = Print_Table #3; Oracle live 2,050 = oracle_kpis Inputs!P48.

### v2 - Memo (3 pages)

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `memo/nebius_vs_peers_unit_economics.docx` | 2026-10-09 11:10 ET | 252,519 | `d30915104359484252c7912ed177387240b6cf56efef605a8c1e067c9bc19ee4` | same |
| `memo/nebius_vs_peers_unit_economics.pdf` | 2026-10-09 11:10 ET | 386,638 | `81677fb88b5f8280fa65f6664263b5831b4e28a3f65ec6272724dac6a89b839d` | same |
| `memo/fig_unit_economics.png` | 2026-10-09 11:10 ET | 233,859 | `eae8ff4b7a97b4246568e829d9e6637f628b0e1eb9f8702903bda677f437848c` | same |
| `memo/build_memo.py` | 2026-10-09 11:08 ET | 28,236 | `7d15746a1944b52d0c67f05bb75c39c97a13b7d12289c6b73f90a8884899968a` | same |
| `memo/build_figs.py` | 2026-10-09 11:09 ET | 5,177 | `86090bb8b5954e4c1cf2cdd3044c3589b65fad0a94d15cbae3c5d8a58c38ed0c` | same |
| `memo/common_data.py` | 2026-10-09 11:12 ET | 7,179 | `af47c5ed275ffd14fcb9c9ffbb2ea328fca1eb4484431aacb57266fe2b93718c` | same |

### v2 - Summary table

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `memo/nebius_vs_peers_summary_table.xlsx` | 2026-10-09 11:10 ET | 20,117 | `b321a0f1a0c1f79b64de46c7f2759f44618d70c6bc2a920930948ef9864053bb` | same |
| `memo/nebius_vs_peers_summary_table.png` | 2026-10-09 11:10 ET | 631,795 | `a628eaba5ffdf883231bbf9b9f19a19e6bb386a8a778d405bb50ddd285c1c170` | same |
| `memo/build_summary_table.py` | 2026-10-09 11:08 ET | 6,975 | `42876bddb61ea26d486b6b6a252defd4355bf3ef17942dc5d32bbaf35e13b851` | same |

### v2 - Comparison

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `comparison.xlsx` | 2026-10-09 14:39 ET | 93,145 | `b1c535b0fd3fc38254e3d6215df5256a27cda54de323e1d352450bc5a8a4c915` | CHANGED |
| `build_comparison.py` | 2026-10-09 14:39 ET | 95,579 | `e3edb2ed15916364837251c7276e2ee6be291714cf13da3d4d1f1afd528d2392` | CHANGED |
| `build_charts.py` | 2026-10-09 11:09 ET | 14,029 | `3f605119f4a58d954c2b31f49ae0d23f0b20fd149cb44862d180822cad1b65f1` | same |
| `nebius_locks.py` | 2026-10-08 19:06 ET | 9,332 | `d64303e86b186bc040779f61d730b057227d300e1358af8588c6dea82c584840` | same |
| `charts/active_vs_contracted_mw.png` | 2026-10-09 14:39 ET | 117,641 | `07dbba3e753e5dca1e2ee6982ff140677f88e6ff3be97c3d1695d2d9ce46eba3` | CHANGED |
| `charts/capex_per_mw.png` | 2026-10-09 14:39 ET | 119,091 | `130b0389019761e34485beec1e076b6e7a3ba902571a11152ccf5a36d929ed13` | same |
| `charts/contracts_usd_per_mw_yr.png` | 2026-10-09 14:39 ET | 210,725 | `e485cd1437d97abf79825ce4ee2d214f3af1ba87f33c6bdbdaed84721c0e367a` | same |
| `charts/cost_of_debt.png` | 2026-10-09 14:39 ET | 100,705 | `959d96b485cfa1866901769db2e70d0b1e836486d98a196f075c8153834b0d34` | same |
| `charts/depreciation_yield_useful_life.png` | 2026-10-09 14:39 ET | 102,098 | `c12f31b2b0a5139d0da3557be0fd8b510fce4a895f6f03471deef0c4e077dd50` | same |
| `charts/nebius_billing_lock.png` | 2026-10-09 14:39 ET | 193,296 | `65503e16139400a042096cef0223c8f981891f41b7f228cd989182c0366e6271` | same |
| `charts/rev_per_active_mw.png` | 2026-10-09 14:39 ET | 160,247 | `3925851967bcd1bc9681819cbe46d1483a1d28149a5519647a1947744489d4ce` | same |
| `charts/rev_per_gpu_hr.png` | 2026-10-09 14:39 ET | 99,326 | `9a4085abbdba873ff426fb082ef6b8519558a11dc9c8aeefe2b32fb56c48f560` | same |
| `charts/rpo_to_annualized_revenue.png` | 2026-10-09 14:39 ET | 69,765 | `14f6546caa01b997b8b626ce1fca6030c650e2fd9d9be8e44b3375221ff7ad56` | same |
| `charts/sites_active_vs_announced_mw.png` | 2026-10-09 14:39 ET | 178,348 | `174b76fec988044efd92c3f90f5172c7410f1f273cac6ded50cce3f83790df63` | same |
| `charts/top_customer_pct.png` | 2026-10-09 14:39 ET | 95,729 | `f2f7c6b4a502df1ca7bc50658303fb2eaa811f37b27506799f9823b782be8384` | same |

### v2 - price_or_mw

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `price_or_mw.xlsx` | 2026-10-09 11:12 ET | 27,999 | `3cca6431d3eeb25dcb0781e28b4ed0f38fad48e1855de1729a324e074ce2e5e8` | same |
| `charts/price_or_mw.png` | 2026-10-09 11:12 ET | 194,907 | `070189b92d44feca9f5538e770b08f661189201f26ef5e92263f45e1d8ecc8c0` | same |
| `patch_price_or_mw.py` | 2026-10-09 11:12 ET | 18,755 | `1522642db4c9ea90b48e1811028c678d9a9b8546114640850a8c92e7b506cbde` | same |
| `build_price_or_mw_chart.py` | 2026-10-09 11:12 ET | 3,545 | `2eff7494e144b3deab7a8d3ccde6c35e98d17abf6c994268af610b562cb38330` | same |

### v2 - Benchmark

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `benchmark/peer_benchmark.xlsx` | 2026-10-09 14:42 ET | 39,677 | `283e9ae05a272a5f033b56ce29fe6fda582602db892531a4353971e3c009b8b6` | CHANGED |
| `benchmark/peer_benchmark.docx` | 2026-10-09 14:42 ET | 43,716 | `a59d660aa97790c0ad415f9837f98434634f3cf2dd541649756f2f5270379707` | CHANGED |
| `benchmark/peer_benchmark.pdf` | 2026-10-09 14:42 ET | 86,336 | `4bf926ab93f6b1e9116394f2f54f12e9eff6e1497a525f7119ff604ea1ea38e3` | CHANGED |
| `benchmark/peer_benchmark.png` | 2026-10-09 14:42 ET | 1,021,355 | `b47bd6c5573c937d32565d3444f86aec4137b48c9c4e0ae5187f7401f010dd8a` | CHANGED |
| `benchmark/build_peer_benchmark.py` | 2026-10-09 14:42 ET | 66,323 | `74bf1b8f500b1a54d5afe4f2cb177af09748045a5e5b3eafe5ad4218b9513bde` | CHANGED |

### v2 - Site appendices

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `sites/appendix_coreweave_denton.docx` | 2026-10-08 19:06 ET | 71,624 | `5a97d13a57b3ae08c2015b3f793ccaa985966d25a5015f9d3d571d7bae80ed16` | same |
| `sites/appendix_coreweave_denton.md` | 2026-10-08 19:06 ET | 6,720 | `5d2cc3126aa4e37d513dd471a2aef4af60cbc9f5fe257bda8cb02e697fdbdc72` | same |
| `sites/appendix_coreweave_denton.pdf` | 2026-10-08 19:06 ET | 114,498 | `84631855d58ae3112b60a0785bc6b6ff5d1530074e0c7373a139c663bb4ff91a` | same |
| `sites/appendix_crusoe_abilene.docx` | 2026-10-08 19:06 ET | 74,709 | `e10eb22ccf70d19d7d417e21bccf2d71b7c1d7801493f7550a9c01a677724c9a` | same |
| `sites/appendix_crusoe_abilene.md` | 2026-10-08 19:06 ET | 7,515 | `5a75ffddb37452f246ecbe75acc69842a3e4b67b64db36211893b7d029a6d383` | same |
| `sites/appendix_crusoe_abilene.pdf` | 2026-10-08 19:06 ET | 116,120 | `d6159bc7191db1ace314cde239ee50f6b2f15d9d09f1b6de08bb6a6c6548e72d` | same |
| `sites/appendix_iren_childress.docx` | 2026-10-08 19:06 ET | 71,251 | `c922aade55d67b6e746b8fa9fd4109373b6fbb570ec1d88cbb45643e0e0f495f` | same |
| `sites/appendix_iren_childress.md` | 2026-10-08 19:06 ET | 6,866 | `457a8685639577b559f038e8a10de648b670893935637f3d8d95f873b7705aa9` | same |
| `sites/appendix_iren_childress.pdf` | 2026-10-08 19:06 ET | 112,647 | `641067c1cbf9927969b2d42d40e46d954160a6518381709e90f21449e3776cc4` | same |
| `sites/appendix_oracle_port_washington.docx` | 2026-10-08 19:06 ET | 68,991 | `81741b4813bd4c5d35c9b5aff2ac9eedff9c31cd5068f1f5da895d418a7ade91` | same |
| `sites/appendix_oracle_port_washington.md` | 2026-10-08 19:06 ET | 6,433 | `ca50e60291fa3644a2d11068d91a398b91c5d477ed06ad930a149ba741240a43` | same |
| `sites/appendix_oracle_port_washington.pdf` | 2026-10-08 19:06 ET | 110,443 | `ea8ea8a92b78ab6f1df8446ae686a6db30b319fb6a6991b5f360706cfcde7c39` | same |
| `sites/peer_site_appendices.pdf` | 2026-10-08 19:06 ET | 281,084 | `546103829b4a5c7851373d05029b2775e969aa22d31501d2cab433c6b40f6ee0` | same |
| `sites/sites.xlsx` | 2026-10-08 23:12 ET | 62,723 | `49f9069a9e9d0dc1c916bfa8865fc049efc63b77a0504b975960c536805ee833` | same |
| `sites/site_verification.md` | 2026-10-08 23:12 ET | 28,470 | `46bcec41af24a744c60662ad6d46283da64bc883b84b1bc4edcb3d2b2ec52be2` | same |
| `sites/build_site_appendices.py` | 2026-10-08 19:06 ET | 9,826 | `c26a99ff4cefbfcf1f74d11d4501b7dc0a5ee23d87a9602392916a4d9ce0b55a` | same |
| `sites/site_data.py` | 2026-10-08 19:06 ET | 25,146 | `b78d97394ac3f091ff5b1bc275f99498d21f8750462983edb26ba99980b2fa80` | same |
| `sites/charts/coreweave_denton_mw.png` | 2026-10-08 19:06 ET | 35,242 | `ce8ff543c59c8fd46fffe2c67942f90867d111bd0c56a955ff2f10d54bcfef3e` | same |
| `sites/charts/crusoe_abilene_mw.png` | 2026-10-08 19:06 ET | 37,895 | `ab1b37d57273c0c20b75854594c00bbfbfb9add70dd505b72b5eab0983710912` | same |
| `sites/charts/iren_childress_mw.png` | 2026-10-08 19:06 ET | 34,477 | `e859642c8938788f9b0b043ff437f80a3af82f4d795e050b062c08a2fff612e3` | same |
| `sites/charts/oracle_port_washington_mw.png` | 2026-10-08 19:06 ET | 32,618 | `7dd2bc8b5774748af038b09fe302373b54666c0ce3fd5d20bd9d3f721efa86e8` | same |

### v2 - Final appendix pack (this build)

| Path (peers/) | mtime | Size (bytes) | SHA-256 | vs v1 |
|---|---|---:|---|---|
| `final_appendix/peer_kpi_appendix.xlsx` | 2026-10-09 14:43 ET | 20,242 | `a8e009efaf31b634fcdc8ebfe5ed8b83313ff5fda5fb82687e4531462964a593` | CHANGED |
| `final_appendix/peer_kpi_appendix.docx` | 2026-10-09 14:43 ET | 466,001 | `ec15b27534e610a21f5c8c267f5c90571fa90a796b341fb83a3ad22cf0007695` | CHANGED |
| `final_appendix/peer_kpi_appendix.pdf` | 2026-10-09 14:43 ET | 463,217 | `2c56abbf921c87f3cce280a3aa8d8d06e9bc797abc5da68609203e239cb056e1` | CHANGED |
| `final_appendix/chart_usd_per_mw_yr.png` | 2026-10-09 14:43 ET | 250,891 | `0c550f2ec8b76fa1518d77b98b66b22c5cfd5229baf98855b2ba3c0c9126725b` | CHANGED |
| `final_appendix/chart_contracted_vs_live_mw.png` | 2026-10-09 14:43 ET | 133,057 | `f999b4644f8f05ce4b718a9ea33f528630fd0b86c96de596be8a2be1b66fa86d` | same |
| `final_appendix/chart_capex_per_mw.png` | 2026-10-09 14:43 ET | 93,189 | `54fa95105066769168ce40e83655f00f40961db5bc1221a8a11e21e431329a2e` | same |
| `final_appendix/build_peer_kpi_appendix.py` | 2026-10-09 14:43 ET | 53,588 | `1c6718ac52a2f00975580b66eef05cdf4f95935fa87b80d3abf880e0426a284d` | CHANGED |

### v2 addendum (Fri Oct 9, 2026 4:35 PM ET)

- Added `peers/frozen_2026-10-11_v2/RECONCILE.md` (5,498 bytes, chmod 444; SHA-256 `99fb62b62832924b79d46b60c839e6132c48bc52b54ec2bbdbbf8e97a8cd221e`). This is a reconciliation note covering IREN live MW, Oracle contracted MW and the deck-match check (8/15 match; all 7 mismatches are tags). It is deliberately NOT listed in SHA256SUMS. No frozen file was changed: the directory was briefly set to u+w and then back to 555, and `sha256sum -c SHA256SUMS` passed on all 61 entries afterwards.

---

### SUPERSEDED: v1 (kept for audit; copies in peers/frozen_2026-10-11/, untouched)


Frozen Fri Oct 09, 2026 14:35 ET by the Nebius Peer KPIs workstream. 61 files in 7 deliverable groups; total 7,517,000 bytes.
Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11/` (same relative paths) + `SHA256SUMS` there. Originals left in place, unchanged.
Verify: `cd /workspace/nebius/peers/frozen_2026-10-11 && sha256sum -c SHA256SUMS` (frozen copies) or run it from `peers/` to check the originals have not moved.
Model: model/nebius_model.xlsx v0.9e (read-only, not copied: owned by the model workstream). Nebius Q3'26 results date: ~Nov 10 (est.), not announced.

### v1 - Locks every Sunday deliverable relies on (model v0.9e)

| Lock | Value | Source cell |
|---|---|---|
| Nebius-Microsoft $ per IT MW [E, house], only figure (11.60 / 17.08 HISTORY only) | $14.49M | model Contracts!E12 = 17,392.9 / (5 x 240) |
| Ex-MSFT hurdle (peer comparator) | $27.63M (27.6278) | model Drivers!B9; price_or_mw Nebius_Check!B37 |
| Memo: Street FY27 ÷ our 531.8 billed MW, all customers [D] | $23.13M | model Drivers!B10 |
| Clear-spot base / high | $24.51M / $22.68M | Drivers!B6 / Debate_Grid!A18 |
| Turn-positive base / high | $17.05M / $16.70M | Debate_Grid!A13 / A10 |
| Energized MW Q1 → Q2'26 [E] | 181.5 → 193 | capacity Quarterly MW!M130 / N130 |
| Billed IT MW YE26 base / high | 257.7 / 284.2 | Judge_Slide!B34 / B35 |
| FY27 avg billed IT MW base / high | 531.84 / 659.16 | Drivers!B34 / C34 |
| K13 capex/MW Nebius vs peer median | 77.9 vs 33.18 | comparison.xlsx Comparison K13 |
| Comp band (Nebius excluded) | $9.7–16.7M | comparison.xlsx Contracts_$MWyr P1 / NEW2 |
| Only disclosed-MW print ≥ $20M | Bitdeer A201 ≤ $22.4M (ceiling) | peer_benchmark.xlsx Contracts BD-A201 |

### v1 - Memo (3 pages)

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18; Q3 date ~Nov 10 (est.)

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `memo/nebius_vs_peers_unit_economics.docx` | 2026-10-09 11:10 ET | 252,519 | `d30915104359484252c7912ed177387240b6cf56efef605a8c1e067c9bc19ee4` |
| `memo/nebius_vs_peers_unit_economics.pdf` | 2026-10-09 11:10 ET | 386,638 | `81677fb88b5f8280fa65f6664263b5831b4e28a3f65ec6272724dac6a89b839d` |
| `memo/fig_unit_economics.png` | 2026-10-09 11:10 ET | 233,859 | `eae8ff4b7a97b4246568e829d9e6637f628b0e1eb9f8702903bda677f437848c` |
| `memo/build_memo.py` | 2026-10-09 11:08 ET | 28,236 | `7d15746a1944b52d0c67f05bb75c39c97a13b7d12289c6b73f90a8884899968a` |
| `memo/build_figs.py` | 2026-10-09 11:09 ET | 5,177 | `86090bb8b5954e4c1cf2cdd3044c3589b65fad0a94d15cbae3c5d8a58c38ed0c` |
| `memo/common_data.py` | 2026-10-09 11:12 ET | 7,179 | `af47c5ed275ffd14fcb9c9ffbb2ea328fca1eb4484431aacb57266fe2b93718c` |

### v1 - Summary table

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `memo/nebius_vs_peers_summary_table.xlsx` | 2026-10-09 11:10 ET | 20,117 | `b321a0f1a0c1f79b64de46c7f2759f44618d70c6bc2a920930948ef9864053bb` |
| `memo/nebius_vs_peers_summary_table.png` | 2026-10-09 11:10 ET | 631,795 | `a628eaba5ffdf883231bbf9b9f19a19e6bb386a8a778d405bb50ddd285c1c170` |
| `memo/build_summary_table.py` | 2026-10-09 11:08 ET | 6,975 | `42876bddb61ea26d486b6b6a252defd4355bf3ef17942dc5d32bbaf35e13b851` |

### v1 - Comparison

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18; FD 358.56M debt / 376.65M if-converted (model Inputs)

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `comparison.xlsx` | 2026-10-09 11:10 ET | 93,098 | `849ac1803f0e6d557b9864600301a262cf9a653757f0386b17dae7ff6b9f54c8` |
| `build_comparison.py` | 2026-10-09 11:10 ET | 94,858 | `2a8fdeb0bdceb10227c7e0c334f76674bb58aa434ce48b94cae078f08a81f30c` |
| `build_charts.py` | 2026-10-09 11:09 ET | 14,029 | `3f605119f4a58d954c2b31f49ae0d23f0b20fd149cb44862d180822cad1b65f1` |
| `nebius_locks.py` | 2026-10-08 19:06 ET | 9,332 | `d64303e86b186bc040779f61d730b057227d300e1358af8588c6dea82c584840` |
| `charts/active_vs_contracted_mw.png` | 2026-10-09 11:10 ET | 109,274 | `ec2a6619595d8665a1d992ffd033d32a2dcdce07b0ede9b5d0809b8d020e7977` |
| `charts/capex_per_mw.png` | 2026-10-09 11:10 ET | 119,091 | `130b0389019761e34485beec1e076b6e7a3ba902571a11152ccf5a36d929ed13` |
| `charts/contracts_usd_per_mw_yr.png` | 2026-10-09 11:10 ET | 210,725 | `e485cd1437d97abf79825ce4ee2d214f3af1ba87f33c6bdbdaed84721c0e367a` |
| `charts/cost_of_debt.png` | 2026-10-09 11:10 ET | 100,705 | `959d96b485cfa1866901769db2e70d0b1e836486d98a196f075c8153834b0d34` |
| `charts/depreciation_yield_useful_life.png` | 2026-10-09 11:10 ET | 102,098 | `c12f31b2b0a5139d0da3557be0fd8b510fce4a895f6f03471deef0c4e077dd50` |
| `charts/nebius_billing_lock.png` | 2026-10-09 11:10 ET | 193,296 | `65503e16139400a042096cef0223c8f981891f41b7f228cd989182c0366e6271` |
| `charts/rev_per_active_mw.png` | 2026-10-09 11:10 ET | 160,247 | `3925851967bcd1bc9681819cbe46d1483a1d28149a5519647a1947744489d4ce` |
| `charts/rev_per_gpu_hr.png` | 2026-10-09 11:10 ET | 99,326 | `9a4085abbdba873ff426fb082ef6b8519558a11dc9c8aeefe2b32fb56c48f560` |
| `charts/rpo_to_annualized_revenue.png` | 2026-10-09 11:10 ET | 69,765 | `14f6546caa01b997b8b626ce1fca6030c650e2fd9d9be8e44b3375221ff7ad56` |
| `charts/sites_active_vs_announced_mw.png` | 2026-10-09 11:10 ET | 178,348 | `174b76fec988044efd92c3f90f5172c7410f1f273cac6ded50cce3f83790df63` |
| `charts/top_customer_pct.png` | 2026-10-09 11:10 ET | 95,729 | `f2f7c6b4a502df1ca7bc50658303fb2eaa811f37b27506799f9823b782be8384` |

### v1 - price_or_mw

Relies on: Nebius_Check B37 = 27.6278 ex-MSFT hurdle; B6 23.1272; B22 14.494 (MSFT house, Contracts!E12; 11.60/17.08 HISTORY only); billed YE26/FY27 locks; band $9.7–16.7M

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `price_or_mw.xlsx` | 2026-10-09 11:12 ET | 27,999 | `3cca6431d3eeb25dcb0781e28b4ed0f38fad48e1855de1729a324e074ce2e5e8` |
| `charts/price_or_mw.png` | 2026-10-09 11:12 ET | 194,907 | `070189b92d44feca9f5538e770b08f661189201f26ef5e92263f45e1d8ecc8c0` |
| `patch_price_or_mw.py` | 2026-10-09 11:12 ET | 18,755 | `1522642db4c9ea90b48e1811028c678d9a9b8546114640850a8c92e7b506cbde` |
| `build_price_or_mw_chart.py` | 2026-10-09 11:12 ET | 3,545 | `2eff7494e144b3deab7a8d3ccde6c35e98d17abf6c994268af610b562cb38330` |

### v1 - Benchmark

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; YE26 guide hurdles 27.16/31.67

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `benchmark/peer_benchmark.xlsx` | 2026-10-09 11:04 ET | 38,529 | `4b605144edac2c5130f429caaed1fbe67a2acacedb40d4e3146e3bcec24b1f8c` |
| `benchmark/peer_benchmark.docx` | 2026-10-09 11:04 ET | 43,670 | `a6b17f92c2775cac39e135a1d119c29a0087e4941f9591d9d0ae6d7e3e3f386e` |
| `benchmark/peer_benchmark.pdf` | 2026-10-09 11:04 ET | 86,271 | `8684bf521477fa13906d84b82da7c34328b1ad1a2bf8042cafaf743abbf925cc` |
| `benchmark/peer_benchmark.png` | 2026-10-09 11:04 ET | 1,055,734 | `c2364c72d99b8ff21dd230641cfe023e04cdaba3b0270017e92b58de56b3b30b` |
| `benchmark/build_peer_benchmark.py` | 2026-10-09 11:04 ET | 61,544 | `e44da74fff2e1a5cb53e5a342fd94d1108a410d8f48b1414661c9c2968a3e8a9` |

### v1 - Site appendices

Relies on: site active-base MW: Denton ~260 IT, Childress 75 (50–100), Abilene 618 (421–618), Port Washington 0, KC 12, Monarch/Kvandal 0; energized 181.5→193 (Nebius context)

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `sites/appendix_coreweave_denton.docx` | 2026-10-08 19:06 ET | 71,624 | `5a97d13a57b3ae08c2015b3f793ccaa985966d25a5015f9d3d571d7bae80ed16` |
| `sites/appendix_coreweave_denton.md` | 2026-10-08 19:06 ET | 6,720 | `5d2cc3126aa4e37d513dd471a2aef4af60cbc9f5fe257bda8cb02e697fdbdc72` |
| `sites/appendix_coreweave_denton.pdf` | 2026-10-08 19:06 ET | 114,498 | `84631855d58ae3112b60a0785bc6b6ff5d1530074e0c7373a139c663bb4ff91a` |
| `sites/appendix_crusoe_abilene.docx` | 2026-10-08 19:06 ET | 74,709 | `e10eb22ccf70d19d7d417e21bccf2d71b7c1d7801493f7550a9c01a677724c9a` |
| `sites/appendix_crusoe_abilene.md` | 2026-10-08 19:06 ET | 7,515 | `5a75ffddb37452f246ecbe75acc69842a3e4b67b64db36211893b7d029a6d383` |
| `sites/appendix_crusoe_abilene.pdf` | 2026-10-08 19:06 ET | 116,120 | `d6159bc7191db1ace314cde239ee50f6b2f15d9d09f1b6de08bb6a6c6548e72d` |
| `sites/appendix_iren_childress.docx` | 2026-10-08 19:06 ET | 71,251 | `c922aade55d67b6e746b8fa9fd4109373b6fbb570ec1d88cbb45643e0e0f495f` |
| `sites/appendix_iren_childress.md` | 2026-10-08 19:06 ET | 6,866 | `457a8685639577b559f038e8a10de648b670893935637f3d8d95f873b7705aa9` |
| `sites/appendix_iren_childress.pdf` | 2026-10-08 19:06 ET | 112,647 | `641067c1cbf9927969b2d42d40e46d954160a6518381709e90f21449e3776cc4` |
| `sites/appendix_oracle_port_washington.docx` | 2026-10-08 19:06 ET | 68,991 | `81741b4813bd4c5d35c9b5aff2ac9eedff9c31cd5068f1f5da895d418a7ade91` |
| `sites/appendix_oracle_port_washington.md` | 2026-10-08 19:06 ET | 6,433 | `ca50e60291fa3644a2d11068d91a398b91c5d477ed06ad930a149ba741240a43` |
| `sites/appendix_oracle_port_washington.pdf` | 2026-10-08 19:06 ET | 110,443 | `ea8ea8a92b78ab6f1df8446ae686a6db30b319fb6a6991b5f360706cfcde7c39` |
| `sites/peer_site_appendices.pdf` | 2026-10-08 19:06 ET | 281,084 | `546103829b4a5c7851373d05029b2775e969aa22d31501d2cab433c6b40f6ee0` |
| `sites/sites.xlsx` | 2026-10-08 23:12 ET | 62,723 | `49f9069a9e9d0dc1c916bfa8865fc049efc63b77a0504b975960c536805ee833` |
| `sites/site_verification.md` | 2026-10-08 23:12 ET | 28,470 | `46bcec41af24a744c60662ad6d46283da64bc883b84b1bc4edcb3d2b2ec52be2` |
| `sites/build_site_appendices.py` | 2026-10-08 19:06 ET | 9,826 | `c26a99ff4cefbfcf1f74d11d4501b7dc0a5ee23d87a9602392916a4d9ce0b55a` |
| `sites/site_data.py` | 2026-10-08 19:06 ET | 25,146 | `b78d97394ac3f091ff5b1bc275f99498d21f8750462983edb26ba99980b2fa80` |
| `sites/charts/coreweave_denton_mw.png` | 2026-10-08 19:06 ET | 35,242 | `ce8ff543c59c8fd46fffe2c67942f90867d111bd0c56a955ff2f10d54bcfef3e` |
| `sites/charts/crusoe_abilene_mw.png` | 2026-10-08 19:06 ET | 37,895 | `ab1b37d57273c0c20b75854594c00bbfbfb9add70dd505b72b5eab0983710912` |
| `sites/charts/iren_childress_mw.png` | 2026-10-08 19:06 ET | 34,477 | `e859642c8938788f9b0b043ff437f80a3af82f4d795e050b062c08a2fff612e3` |
| `sites/charts/oracle_port_washington_mw.png` | 2026-10-08 19:06 ET | 32,618 | `7dd2bc8b5774748af038b09fe302373b54666c0ce3fd5d20bd9d3f721efa86e8` |

### v1 - Final appendix pack (this build)

Relies on: ALL: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18; asserted at build + cell-for-cell vs comparison.xlsx / peer_benchmark.xlsx / price_or_mw.xlsx / sites

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `final_appendix/peer_kpi_appendix.xlsx` | 2026-10-09 14:34 ET | 19,246 | `e718a53c734367052ccbb087362ca39879a29c6161c499240434a79a56c4eb86` |
| `final_appendix/peer_kpi_appendix.docx` | 2026-10-09 14:34 ET | 456,974 | `629d6fbc6c7dacf99cbbaf50de89ef5492cb5f5b31114c883da380b2585234c1` |
| `final_appendix/peer_kpi_appendix.pdf` | 2026-10-09 14:34 ET | 455,075 | `5983c7c78770f941392a270a8aefa7002a6bfc7bf4b196989099a18cb80ec78d` |
| `final_appendix/chart_usd_per_mw_yr.png` | 2026-10-09 14:34 ET | 240,648 | `39eda5f1655d5ddb222d8913192072ac094d771c54463fd16160a331fee1be8f` |
| `final_appendix/chart_contracted_vs_live_mw.png` | 2026-10-09 14:34 ET | 133,057 | `f999b4644f8f05ce4b718a9ea33f528630fd0b86c96de596be8a2be1b66fa86d` |
| `final_appendix/chart_capex_per_mw.png` | 2026-10-09 14:34 ET | 93,189 | `54fa95105066769168ce40e83655f00f40961db5bc1221a8a11e21e431329a2e` |
| `final_appendix/build_peer_kpi_appendix.py` | 2026-10-09 14:33 ET | 50,143 | `e2b5e20224f54fe8f826bde2dfad56d62bd794f53b282717800775e8d5022653` |

### v1 - Known differences between the frozen files (documented, not edited)

1. Oracle live MW: comparison.xlsx Chart_Data / active_vs_contracted_mw.png use 1,200 (cumulative AI MW delivered, May-26 vintage, matching its K13 period); peer_benchmark, price_or_mw and the final appendix use ~2,050 (same series, Aug-26). Appendix footnote 7.
2. Nebius-Meta #1: comparison.xlsx Contracts_$MWyr N2 (and Chart_Data / contracts chart) carry 11.52 on an unsourced ~50 MW (E, low confidence); peer_benchmark and the appendix show n/d (filings state no MW). Not a band input. Appendix footnote 11.
3. IREN live MW: company ~40 MW AI operating at Jun-30 (comparison, benchmark) vs Childress site ~75 IT active in Oct (site appendix) - vintage difference; IREN contracted/live 6.5x is computed only in the appendix (comparison K04 = n/d) and kept out of the median.
4. Oracle contracted: comparison Chart_Data shows 10,000; benchmark says 'n/d as filed; >10 GW secured' - same lower bound; the appendix shows '>10,000 [R, lower bound]'.
