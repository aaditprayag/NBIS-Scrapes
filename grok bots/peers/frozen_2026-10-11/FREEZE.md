# FREEZE - peer deliverables for Sunday Oct 11, 2026 finals

Frozen Fri Oct 09, 2026 14:35 ET by the Nebius Peer KPIs workstream. 61 files in 7 deliverable groups; total 7,517,000 bytes.
Read-only copies (chmod 444, dirs 555): `peers/frozen_2026-10-11/` (same relative paths) + `SHA256SUMS` there. Originals left in place, unchanged.
Verify: `cd /workspace/nebius/peers/frozen_2026-10-11 && sha256sum -c SHA256SUMS` (frozen copies) or run it from `peers/` to check the originals have not moved.
Model: model/nebius_model.xlsx v0.9e (read-only, not copied: owned by the model workstream). Nebius Q3'26 results date: ~Nov 10 (est.), not announced.

## Locks every Sunday deliverable relies on (model v0.9e)

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

## Memo (3 pages)

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18; Q3 date ~Nov 10 (est.)

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `memo/nebius_vs_peers_unit_economics.docx` | 2026-10-09 11:10 ET | 252,519 | `d30915104359484252c7912ed177387240b6cf56efef605a8c1e067c9bc19ee4` |
| `memo/nebius_vs_peers_unit_economics.pdf` | 2026-10-09 11:10 ET | 386,638 | `81677fb88b5f8280fa65f6664263b5831b4e28a3f65ec6272724dac6a89b839d` |
| `memo/fig_unit_economics.png` | 2026-10-09 11:10 ET | 233,859 | `eae8ff4b7a97b4246568e829d9e6637f628b0e1eb9f8702903bda677f437848c` |
| `memo/build_memo.py` | 2026-10-09 11:08 ET | 28,236 | `7d15746a1944b52d0c67f05bb75c39c97a13b7d12289c6b73f90a8884899968a` |
| `memo/build_figs.py` | 2026-10-09 11:09 ET | 5,177 | `86090bb8b5954e4c1cf2cdd3044c3589b65fad0a94d15cbae3c5d8a58c38ed0c` |
| `memo/common_data.py` | 2026-10-09 11:12 ET | 7,179 | `af47c5ed275ffd14fcb9c9ffbb2ea328fca1eb4484431aacb57266fe2b93718c` |

## Summary table

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; K13 capex/MW 77.9 vs peer median 33.18

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `memo/nebius_vs_peers_summary_table.xlsx` | 2026-10-09 11:10 ET | 20,117 | `b321a0f1a0c1f79b64de46c7f2759f44618d70c6bc2a920930948ef9864053bb` |
| `memo/nebius_vs_peers_summary_table.png` | 2026-10-09 11:10 ET | 631,795 | `a628eaba5ffdf883231bbf9b9f19a19e6bb386a8a778d405bb50ddd285c1c170` |
| `memo/build_summary_table.py` | 2026-10-09 11:08 ET | 6,975 | `42876bddb61ea26d486b6b6a252defd4355bf3ef17942dc5d32bbaf35e13b851` |

## Comparison

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

## price_or_mw

Relies on: Nebius_Check B37 = 27.6278 ex-MSFT hurdle; B6 23.1272; B22 14.494 (MSFT house, Contracts!E12; 11.60/17.08 HISTORY only); billed YE26/FY27 locks; band $9.7–16.7M

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `price_or_mw.xlsx` | 2026-10-09 11:12 ET | 27,999 | `3cca6431d3eeb25dcb0781e28b4ed0f38fad48e1855de1729a324e074ce2e5e8` |
| `charts/price_or_mw.png` | 2026-10-09 11:12 ET | 194,907 | `070189b92d44feca9f5538e770b08f661189201f26ef5e92263f45e1d8ecc8c0` |
| `patch_price_or_mw.py` | 2026-10-09 11:12 ET | 18,755 | `1522642db4c9ea90b48e1811028c678d9a9b8546114640850a8c92e7b506cbde` |
| `build_price_or_mw_chart.py` | 2026-10-09 11:12 ET | 3,545 | `2eff7494e144b3deab7a8d3ccde6c35e98d17abf6c994268af610b562cb38330` |

## Benchmark

Relies on: MSFT $14.49M per IT MW [E, house] (Contracts!E12) only; ex-MSFT hurdle $27.63M; $23.13M = Street FY27 ÷ our 531.8 billed MW, all customers [D]; clear-spot $24.51M/$22.68M; turn-positive $17.05M/$16.70M; energized 181.5→193 [E]; billed YE26 257.7/284.2, FY27 531.84/659.16; comp band $9.7–16.7M (Nebius rows excluded); Bitdeer ≤$22.4M only disclosed-MW print ≥$20M; YE26 guide hurdles 27.16/31.67

| Path (peers/) | mtime | Size (bytes) | SHA-256 |
|---|---|---:|---|
| `benchmark/peer_benchmark.xlsx` | 2026-10-09 11:04 ET | 38,529 | `4b605144edac2c5130f429caaed1fbe67a2acacedb40d4e3146e3bcec24b1f8c` |
| `benchmark/peer_benchmark.docx` | 2026-10-09 11:04 ET | 43,670 | `a6b17f92c2775cac39e135a1d119c29a0087e4941f9591d9d0ae6d7e3e3f386e` |
| `benchmark/peer_benchmark.pdf` | 2026-10-09 11:04 ET | 86,271 | `8684bf521477fa13906d84b82da7c34328b1ad1a2bf8042cafaf743abbf925cc` |
| `benchmark/peer_benchmark.png` | 2026-10-09 11:04 ET | 1,055,734 | `c2364c72d99b8ff21dd230641cfe023e04cdaba3b0270017e92b58de56b3b30b` |
| `benchmark/build_peer_benchmark.py` | 2026-10-09 11:04 ET | 61,544 | `e44da74fff2e1a5cb53e5a342fd94d1108a410d8f48b1414661c9c2968a3e8a9` |

## Site appendices

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

## Final appendix pack (this build)

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

## Known differences between the frozen files (documented, not edited)

1. Oracle live MW: comparison.xlsx Chart_Data / active_vs_contracted_mw.png use 1,200 (cumulative AI MW delivered, May-26 vintage, matching its K13 period); peer_benchmark, price_or_mw and the final appendix use ~2,050 (same series, Aug-26). Appendix footnote 7.
2. Nebius-Meta #1: comparison.xlsx Contracts_$MWyr N2 (and Chart_Data / contracts chart) carry 11.52 on an unsourced ~50 MW (E, low confidence); peer_benchmark and the appendix show n/d (filings state no MW). Not a band input. Appendix footnote 11.
3. IREN live MW: company ~40 MW AI operating at Jun-30 (comparison, benchmark) vs Childress site ~75 IT active in Oct (site appendix) - vintage difference; IREN contracted/live 6.5x is computed only in the appendix (comparison K04 = n/d) and kept out of the median.
4. Oracle contracted: comparison Chart_Data shows 10,000; benchmark says 'n/d as filed; >10 GW secured' - same lower bound; the appendix shows '>10,000 [R, lower bound]'.
