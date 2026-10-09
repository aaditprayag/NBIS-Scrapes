# Funding capacity note — Class A share authority (Filings & Capital)

**As of:** Mon Oct 5, 2026  
**Audience:** Modeling / CapStructure  
**Public information only.** Tags: Reported / Derived (math shown) / Estimated (method shown).

---

## AGM authority (Items 13–14) — PASSED

**Source:** 6-K results release 2026-08-26, Ex.99.1 `edgar/tm2624077d1_ex99-1.htm.txt` (AGM held 2026-08-25).

| Item | Resolution | For | Against | Abstained |
|------|------------|-----|---------|-----------|
| **13** | Board may issue Class A up to an **additional 20%** of issued share capital **excluding Class C**, for **5 years** from AGM date | **375,374,419** | **23,704,865** | 4,000,909 |
| **14** | Board may **exclude pre-emptive rights** on such Class A issues, **5 years** | **369,944,562** | **29,062,939** | 4,072,692 |

Both **PASSED**. [Reported]

### Eligible votes at AGM [Reported]
- Class A: **238,402,543** shares × 1 vote = 238,402,543 votes  
- Class B: **33,455,053** shares × 10 votes = **334,550,530** votes  
- Voted as a single class

---

## Pool size (~20% authority)

Issued capital excl. Class C at **6/30/26** (Q2 interim Note 14):
- Class A issued: **288,585,891** [Reported]
- Class B issued: **33,455,053** [Reported]
- Sum: **322,040,944** [Derived: 288,585,891 + 33,455,053]

**AGM issuance pool ≈ 20% × 322,040,944 = 64,408,188.8 ≈ 64.4M Class A shares** [Derived]

*(AGM text: “up to an additional 20% of the issued share capital (excluding Class C Shares) … from time to time” — pool refreshes with issued capital; 6/30 snapshot used here.)*

---

## Treasury + headroom

| Component | Shares | Tag | Source / math |
|-----------|--------|-----|----------------|
| Class A issued 6/30/26 | 288,585,891 | Reported | Q2FS Note 14 |
| Class A outstanding 6/30/26 | 238,400,165 | Reported | Q2 PR / BS outstanding shares |
| **Treasury Class A 6/30/26** | **50,185,726** | **Derived** | 288,585,891 − 238,400,165 |
| AGM 20% pool (snapshot) | ~64.4M | Derived | above |
| **Headroom without new vote** | **~114.6M** | **Derived** | 50,185,726 + 64,408,189 = **114,593,915** |
| NVIDIA pre-funded warrant (unexercised) | 21,065,936 | Reported | Q2FS Note 14; exercise **not disclosed** through latest 6-Ks (through 10/1/26 Inferize 6-K) |
| Settlement of warrant | From **treasury or authorized share capital** | Reported | FY25 20-F: “Upon exercise, the Warrant Shares will be delivered from the Group’s treasury shares or issued from its authorized share capital.” |
| **Headroom net of unexercised warrant** | **~93.5M** | **Derived** | 114,593,915 − 21,065,936 = **93,527,979** |

**Modeling shorthand:** ~**94–115M** Class A funding-capacity shares without a new AGM vote (lower end nets the NVIDIA warrant; upper end is treasury + 20% pool).

Authorized Class A remains **500,000,000** [Reported, Note 14] — absolute ceiling is higher; the binding near-term constraint for *board-authorized* issuance *with pre-emption exclusion* is the AGM 20% + treasury path above.

---

## Estimated Class A dissent (if Class B voted all For)

Method: attribute all Against votes to Class A; assume Class B’s **334,550,530** votes were For; Class A participation = total votes cast − Class B votes.

Total votes cast Items 13/14 = For + Against + Abstained = **403,080,193** (same for both items).  
Class A participation ≈ 403,080,193 − 334,550,530 = **68,529,663**.

| Item | Against | Against / Class A participation | Est. Class A dissent |
|------|---------|----------------------------------|----------------------|
| 13 | 23,704,865 | 23,704,865 / 68,529,663 | **~34.6% ≈ 35%** [Estimated] |
| 14 | 29,062,939 | 29,062,939 / 68,529,663 | **~42.4% ≈ 42%** [Estimated] |

Assumptions: Class B voted 100% For; all Against from Class A; abstentions included in participation denominator. Actual Class B split not disclosed.

---

## Related gaps (not in this note’s scope)

- NVIDIA warrant **exercise status** after issuance: not disclosed through latest 6-Ks  
- Drawn amount on **$775M** senior secured facility (signed Jul-2026): not disclosed  
- Post-6/30 ATM sales / Aug-26 exchange share issuance already reduced treasury / increased issued vs 6/30 snapshot — refresh pool/treasury after next interim BS

---

## File pointers

- AGM results: `/workspace/nebius/filings/edgar/tm2624077d1_ex99-1.htm.txt`  
- Q2 share capital / warrant: `/workspace/nebius/filings/edgar/nbis-20260812xex99d2.htm.txt` Note 14  
- NVIDIA settlement language: `/workspace/nebius/filings/edgar/nbis-20251231x20f.htm.txt` (Equity Investment with NVIDIA / Note)  
- Fact base: `/workspace/nebius/filings/nebius_fact_base.md`  
