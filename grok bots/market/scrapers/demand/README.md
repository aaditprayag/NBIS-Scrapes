# Nebius demand signals – scrapers (run date 2026-10-04 ET)

Everything here uses public, unauthenticated endpoints. Raw downloads go in `raw/`. Outputs are dated CSVs in this folder.
The workbook is `/workspace/nebius/parts/demand_signals.xlsx`. Charts go to `/workspace/nebius/market/charts/demand_*.png`.

## Run order
| # | script | what it does | output |
|---|---|---|---|
| 1 | `jobs_snapshot.py` | Pulls live ATS boards: Nebius and CoreWeave from the Greenhouse `boards-api.greenhouse.io/v1/boards/<co>/jobs?content=true`, Lambda from the Ashby `api.ashbyhq.com/posting-api/job-board/lambda` | `raw/*_jobs_<date>.json`, `jobs_<date>.csv` |
| 2 | `classify_jobs.py` | Rule-based function classifier on title plus department. Maps each role to a country, and maps DC and hardware roles to a site (capacity/site_list.csv) | `jobs_classified_<date>.csv` |
| 3 | `wayback_jobs.py [fetch\|parse]` | Gets the Wayback CDX list of ATS API captures, downloads each `id_` capture (gzip-safe) and parses it to rows | `raw/wayback/jobs_<co>/*.json`, `jobs_history.csv` |
| 4 | `tranco_history.py` | Tranco list rank at quarter-ends for nebius.com, nebius.ai, coreweave.com, lambda.ai, lambdalabs.com, crusoe.ai, together.ai, nscale.com | `tranco_<date>.csv` |
| 5 | `openrouter_stats.py` | openrouter.ai/providers (monthly tokens for all providers), the provider/nebius chart data (90-day daily tokens by model), and /api/v1/models prices | `openrouter_*_<date>.csv` |
| 6 | `customer_stories.py` | Customer-story URLs from the nebius.com sitemap, the GPU and usage numbers each page discloses, and the first Wayback capture of each URL | `customer_stories_<date>.csv`, `raw/wayback/stories_first_seen.json` |
| 7 | `github_stats.py` | GitHub search API: repos and stars for the nebius, coreweave and lambdal orgs | `github_<date>.csv` |
| 8 | `mentions.py` | HN Algolia comment and story counts per quarter (Nebius, CoreWeave, Lambda), plus Stack Overflow tag and search counts | `mentions_<date>.csv` |
| 9 | `docs_changelog.py` | Parses docs.nebius.com/changelog.md into entries per week and quarter, and flags self-serve entries (trial, spot, pricing) | `docs_changelog_*.csv` |
| 10 | `build_demand_xlsx.py` | Builds the workbook: one sheet per signal, plus Mix_Estimate, Sources and README | `parts/demand_signals.xlsx` |
| 11 | `demand_charts.py` | Builds the deck PNGs. Run it after step 10, because the mix and customer charts read the workbook | `market/charts/demand_*.png` |

## Tags
- **Observed** means read straight off a public third-party page or API.
- **Reported** means a company filing, letter or call.
- **Derived** means arithmetic on Observed or Reported values.
- **Estimated** means it needs an assumption. The assumption is stated in the note column.

## Known gaps (as of 2026-10-04)
- web.archive.org rate-limits the shared box IP with 429 errors and TLS EOFs. Only 2 of the 14 Nebius ATS captures downloaded (Nov-24 and Dec-24). None of the CoreWeave or Lambda captures downloaded. To fill the history later, rerun `python3 wayback_jobs.py fetch && python3 wayback_jobs.py parse`, then steps 10 and 11. Until then, the `first_published` vintage table in Jobs_History is the proxy for the hiring ramp. It has survivorship bias.
- No public ATS capture exists for Nebius before Nov 2024, so Q1–Q3 2024 job counts are not available.
- Similarweb returned 403, and no Semrush or Ahrefs numbers were visible, so there is no visit data. Tranco gives a rank, not visits, and has no subdomain ranks (studio.nebius.com, tokenfactory.nebius.com).
- OpenRouter shows only 90 days of daily history per provider. It covers only the share of Token Factory traffic that OpenRouter routes.
- Hugging Face does not list Nebius as a live inference provider. The date it was removed was not found.
- Model-catalog size over time was not collected, because the Token Factory model pages are JS-rendered and their Wayback captures were rate-limited.
- No GitHub contributor history (the core API is rate-limited when unauthenticated). No public Discord invite or size.
- The mapping of customers C, D and E to names is an inference, not a disclosure.
