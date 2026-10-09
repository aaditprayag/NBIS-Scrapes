# PROGRESS (stage 2 done 2026-10-04 ~19:00 ET)

## Done
- Stage 1: H100 SXM 80GB (all providers), Nebius discount wording, Nebius 2026 increases dated.
- Stage 2: QA of H200, B200, B300, GB200/GB300 NVL72 (explicit "Contact sales" rows -> No_Public_Price sheet), H100 PCIe/NVL, A100 (40GB/80GB/80GB PCIe, Nov 2020 -> Oct 2026), L40S.
  Sheets: SKU_Nebius_vs_Medians, A100, A100_Decay, A100_Monthly_2020_2026_FFill/Obs, L40S, No_Public_Price, Coverage_by_SKU, Delisted_Series.
  Charts: market/charts/medians_<sku>.png (12 SKUs), a100_decay_2020_2026.png.
- GCP: gpus-pricing iframes 2020-23 (GPU component) + vm-instance-pricing static tables 2025-02..2026-03 (bundled VM).
- Azure: instances.vantage.sh pages + Retail API (ND A100 v4, NC A100 v4).
- AWS: P4 instance-type page (Nov 2020 launch: $32.77 / 8) + vantage p4d/p4de/p5e/p5en/p6-b200.
- RunPod gap Jul 2025-May 2026 filled (gpu-models pages); RunPod 2022-24 (gpu-instance/pricing).
- Vast.ai marketplace ranges Aug 2024-Jun 2025; Nscale confirmed no public prices.
- Forward-fill is delisting-aware: a series stops being carried once >=2 later priced captures of the same page lack it (Delisted_Series sheet). Raw sheets are never filled.
- Nebius L40S: Nebius swapped the Intel/AMD labels between the 2026-09 and 2026-10 pages (same vCPU configs/prices); headline = lowest "from" config ($1.55) throughout.
- time_to_deploy.xlsx fixed for missed nebius.ai snapshots (Nebius H200 first listed 2024-10-06 at $4.85).

## Remaining gaps
1. GCP Jul 2023-Jan 2025 (JS pages); GCP after 2026-03 carried forward as stale.
2. Azure has no intermediate history from the Retail API (current price + effective date only).
3. Vast.ai after Jun 2025 (JS only). Nscale: never published prices.
4. Lambda A100 2021-05..2022-06 pages carried no prices. CoreWeave 2022-24 a-la-carte only.
5. Nebius A100 assumed 80GB (page says "A100 SXM4").
6. Many Blackwell / GB200 / GB300 offers are "contact sales" (see No_Public_Price).
