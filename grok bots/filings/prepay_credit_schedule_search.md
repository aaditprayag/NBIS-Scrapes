# Prepayment credit schedule search — Nebius (NBIS)

**As of:** Mon Oct 5, 2026  
**Scope:** All Nebius SEC filings under `/workspace/nebius/filings/edgar/` and mined transcripts/letters under `/workspace/nebius/filings/altdata/mined_text/`.  
**Question:** Does Nebius disclose **how** customer prepayments are credited against future billings (e.g. back-loaded after month 24, pro rata over the term, front-loaded, months 25–60)?

---

## Conclusion

**Nebius does not disclose the credit schedule** for how customer prepayments are applied / credited against future billings.

Corpus-wide search (regex for credit-against-billings, pro rata over/after, back-load / front-load, months 25–60, after month 24, applied against future invoices) returned **no Nebius disclosure** of a prepayment-to-billing credit schedule. Mentions of “service credits” are **SLA / delivery-delay remedies**, not prepaid-capacity amortization schedules.

---

## What IS disclosed (Nebius)

| Item | Figure / language | Tag | Source |
|------|----------------|-----|--------|
| Q2-deal prepayments vs CapEx | **50–60%** of associated CapEx covered by prepayments; ~70% of deals include prepayments | Reported | Q2’26 shareholder letter (`altdata/mined_text/Q2-26_letter.txt`): “70% of deals including prepayments, covering 50-60% of the associated capex”; slide text “50-60% of capex self-financed … by prepayments”. Also Ex.99.2 6-K 2026-08-12. |
| 2026 prepayment guide | **>$9B** customer prepayments expected in 2026 | Reported | Q2’26 letter: “We expect over $9 billion in customer prepayments in 2026.” / “We expect to receive over $9 billion in customer prepayments in 2026.” |
| Deferred revenue 6/30/26 | **$5,975.2M** total (current **$979.4M** + non-current **$4,995.8M**) | Reported | Q2’26 interim FS `edgar/nbis-20260812xex99d2.htm.txt`, balance sheet + Note (Deferred Revenue), p.F-18 / F-19 |
| Deferred revenue draw (Q2’26) | **$130.8M** of Q2’26 revenue was from the opening deferred balance; **$154.3M** for 1H’26 | Reported | Same note, p.F-18: “Revenue recognized during the three and six months ended June 30, 2026, which was included in the deferred revenue balances at the beginning of each respective period, was $130.8 and $154.3, respectively.” |
| “Amounts that may be refundable” | Deferred revenue includes **amounts that may be refundable** | Reported | Same note: “The Group recognizes deferred revenue when cash is received and before performance obligations are fulfilled, **including amounts that may be refundable**.” |
| Recognition horizon (deferred) | Expected to be recognized as revenue over **one to five calendar years** | Reported | Same note — recognition *period*, not a credit-vs-billing schedule |
| RPO timing (not prepay credit) | RPO **$37,490.6M** at 6/30/26: **36%** within 24 months ending 6/30/28; **40%** months **25–48**; remainder thereafter | Reported | Same Note 1, Remaining performance obligations — this is **revenue recognition timing of RPO**, not how prepayments offset invoices |
| Microsoft upfront | Aggregate upfront payments **~$6,958.1M** of up to **$17,392.9M** TCV; remainder invoiced monthly through Oct-2031 | Reported | FY25 20-F `edgar/nbis-20251231x20f.htm.txt`, Note 1 Significant Customer Agreements, p.F-17: “including aggregate upfront payments of approximately $6,958.1, with the remaining consideration invoiced monthly over the service terms of the respective tranches through October 2031.” |
| Microsoft / Meta “service credits” | Rights to receive **service credits** or terminate tranches for delivery delays / repeated availability failure | Reported | 20-F Material Contracts (Microsoft / Meta 2025) and Note 1 — **SLA credits**, not prepaid-capacity credit schedules |
| Payback (context) | Expected payback on Q2 deals’ CapEx + related opex = **1 year 10 months** (forecast incl. capacity not yet built; revenue-recognition basis excluding prepayment) | Reported | Q2’26 letter footnotes (2)(3) and body |

### Exact quotes (key)

**Deferred revenue / refundable / draw (Q2 interim, Deferred Revenue note, ~p.F-18):**
> “The Group recognizes deferred revenue when cash is received and before performance obligations are fulfilled, including amounts that may be refundable. … As of December 31, 2025 and June 30, 2026, the balance of deferred revenue, including current and non-current portion, was $1,577.5 and $5,975.2, respectively. … Revenue recognized during the three and six months ended June 30, 2026, which was included in the deferred revenue balances at the beginning of each respective period, was $130.8 and $154.3, respectively. … Deferred revenue is expected to be recognized as revenue over the period ranging from one to five calendar years.”

**Microsoft upfront (FY25 20-F Note 1, ~p.F-17):**
> “Microsoft has committed to pay the Group fees under the Microsoft Agreement estimated to be up to $17,392.9, irrespective of actual utilization of the GPU capacity, including aggregate upfront payments of approximately $6,958.1, with the remaining consideration invoiced monthly over the service terms of the respective tranches through October 2031.”

**No schedule language found** for: pro rata credit of prepay after month 24; back-loaded / front-loaded credit; months 25–60 credit window; “credited against future billings” mechanics.

---

## Peer context only (NOT a Nebius fact)

**IREN–Microsoft 8-K (peer):** discloses that **20% prepay is credited pro rata after month 24**.  
Cite **only** as peer-contract structure context for Modeling / comps. **Do not** treat as Nebius’s schedule — Nebius has not disclosed an equivalent mechanic in the searched corpus.

---

## Files checked (priority + sweep)

- `edgar/nbis-20260812xex99d2.htm.txt` (Q2’26 interim — deferred / RPO notes)
- `edgar/nbis-20251231x20f.htm.txt` (FY25 20-F — Microsoft / Meta / deferred)
- `edgar/nbis-20260331xex99d2.htm.txt` (Q1’26 interim)
- `edgar/nbis-20241231x20f.htm.txt` (FY24 20-F)
- `altdata/mined_text/Q2-26_letter.txt`, `Q2-26_call.txt`
- Broader `edgar/*.txt` + `altdata/mined_text/*.txt` pattern sweep (see Conclusion)

---

## Implication for Modeling

Revenue timing of deferred / RPO is disclosed at a coarse horizon (1–5 years; RPO buckets 0–24 / 25–48 / later). **Invoice-level prepayment credit timing is a modeling assumption**, not a filings input — flag as Estimated / scenario unless / until Nebius discloses.
