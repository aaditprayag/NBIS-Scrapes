# Contract liabilities / prepayments / commitments by quarter ($M). None = not available/not disclosed.
Q=['Q1-24','Q2-24','Q3-24','Q4-24','Q1-25','Q2-25','Q3-25','Q4-25','Q1-26','Q2-26']
# each row: (metric, {quarter: value}, {quarter: src_key}, note)
NA=None
R=[
('Deferred revenue (contract liability), total',
 {'Q2-24':13.8,'Q3-24':17.4,'Q4-24':16.3,'Q1-25':19.0,'Q2-25':19.3,'Q3-25':15.6,'Q4-25':1577.5,'Q1-26':4778.1,'Q2-26':5975.2},
 {'Q2-24':'Q2_24_PR','Q3-24':'Q3_24_PR','Q4-24':'FY25_20F','Q1-25':'Q1_25_PR','Q2-25':'Q2_25_PR','Q3-25':'Q3_25_PR','Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},
 'Q1-24 n/a: Yandex N.V. pre-divestment consolidated BS, not comparable. Q4-24 shown as restated $16.3M in FY25 20-F (Q4-24 PR printed $16.5M). Pre-Q4-25 balances are small (mostly TripleTen/legacy advances).'),
('  of which current',{'Q4-25':275.5,'Q1-26':685.6,'Q2-26':979.4},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Split first shown at YE25.'),
('  of which non-current',{'Q4-24':0.0,'Q4-25':1302.0,'Q1-26':4092.5,'Q2-26':4995.8},{'Q4-24':'FY25_20F','Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Non-current = recognized beyond 12 months: 83.6% of total at Q2-26.'),
('Change in deferred revenue in operating cash flow (quarter)',
 {'Q4-25':1566.5,'Q1-26':3198.0,'Q2-26':1197.0},{'Q4-25':'Q4_25_PR','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_PR'},'Q4-25 from Q4 column of PR cash flow; FY25 total 1,565.8. Q1-25..Q3-25 immaterial (<$11M YTD).'),
('Advances from strategic customers (cash, period)',{'Q4-25':982.5,'Q1-26':3198.0},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_OFR'},'FY25: "advances received from customers under strategic customer agreements in amount of $982.5 million" (full-year figure, booked in Q4). Q1-26: "advances received under customer agreements of $3,198.0 million". Difference between $982.5M cash and $1,561M deferred increase at FY25 = billed/receivable not yet collected (inference).'),
('Revenue recognized from opening deferred revenue (period)',{'Q4-24':6.9,'Q4-25':16.3,'Q2-26':130.8},{'Q4-24':'FY25_20F','Q4-25':'FY25_20F','Q2-26':'Q2_26_FS'},'Q4-24 and Q4-25 values are FULL-YEAR; Q2-26 is three months ($154.3M six months). Burn of prepayments into revenue starting.'),
('Remaining performance obligations (RPO)',{'Q4-25':21333.0,'Q1-26':33585.3,'Q2-26':37490.6},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Not disclosed before YE25. Q2-26: "fully attributed to the Group’s Nebius AI cloud business".'),
('  RPO % recognized within 24 months',{'Q4-25':28,'Q1-26':29,'Q2-26':36},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'%'),
('  RPO % recognized months 25-48',{'Q4-25':39,'Q1-26':39,'Q2-26':40},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'%'),
('  Implied RPO $ within 24 months',{'Q4-25':round(21333.0*.28,1),'Q1-26':round(33585.3*.29,1),'Q2-26':round(37490.6*.36,1)},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Derived = RPO x disclosed %. Q2-26: ~$13.5B over Jul-26..Jun-28 = ~$6.7B/yr avg.'),
('Deferred revenue / RPO (%)',{'Q4-25':round(1577.5/21333.0*100,1),'Q1-26':round(4778.1/33585.3*100,1),'Q2-26':round(5975.2/37490.6*100,1)},{},'Derived: share of backlog already cash-prepaid/billed.'),
('Microsoft contract: committed fees (TCV, up to)',{'Q4-25':17392.9},{'Q4-25':'FY25_20F'},'"estimated to be up to $17,392.9 ... including aggregate upfront payments of approximately $6,958.1" (FY25 20-F Note 1). 5-yr term, 9 tranches 2025-26, Vineland NJ.'),
('Microsoft contract: aggregate upfront payments',{'Q4-25':6958.1},{'Q4-25':'FY25_20F'},'Aggregate over contract, not cash received to date (timing not disclosed).'),
('Meta deal 1 (Nov-2025) TCV',{'Q4-25':2880.7},{'Q4-25':'FY25_20F'},'2 tranches delivered Nov-25 and Feb-26.'),
('Meta deal 2 (Mar-2026) dedicated TCV / unsold-capacity backstop',{'Q1-26':12000.0},{'Q1-26':'FY25_20F'},'"approximately $12 billion of dedicated compute capacity" + "up to $15 billion" purchase of unsold capacity (total up to ~$27B). Delivery from early 2027.'),
('Purchase commitments: power/utilities & other (total future, undiscounted)',{'Q4-24':29.4,'Q4-25':29.1,'Q2-26':5324.1},{'Q4-24':'FY24_20F','Q4-25':'FY25_20F','Q2-26':'Q2_26_FS'},'Q4-24: 5.9+6.0+6.2+6.4+4.9. Q4-25: power 6.8+7.0+7.2+5.6 + other 0.6+1.0+0.4+0.4+0.1. Q2-26: energy/power supply fixed or min-guaranteed 81.7(2026)+541.8+542.0+540.5+535.8+3,082.3 thereafter. Q1-26 note has no commitments table (n/a). 183x jump in 2 quarters = owned GW-scale sites.'),
('Leases signed, not yet commenced (future undiscounted payments)',{'Q4-24':168.0,'Q4-25':9759.6,'Q1-26':9905.2,'Q2-26':12054.1},{'Q4-24':'FY24_20F','Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Off-balance-sheet until commencement; mostly data centers, commence 2026-27, terms up to 11-12 yrs.'),
('  Prepayments made on not-yet-commenced leases',{'Q4-25':271.0,'Q1-26':453.4,'Q2-26':612.4},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},''),
('GPU / hardware purchase obligations (non-cancellable)',{},{},'NOT DISCLOSED as a quantified commitment in any filing reviewed. Filings reference "property and equipment received in advance from suppliers under deferred payment arrangements" (interest-bearing vendor financing) without a balance.'),
('Acquired P&E/intangibles not yet paid for (supplemental CF, non-cash)',{'Q4-25':1057.7,'Q1-26':430.4,'Q2-26':973.3},{'Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'As labelled in supplemental cash-flow disclosure: Q4-25 value = FY2025; Q1-26 = 3M; Q2-26 = 6M (H1-26). Vendor-deferred capex on top of cash capex (cash capex understates accrual spend).'),
('Restricted cash (end of period, continuing ops)',{'Q1-26':328.7,'Q2-26':1056.0},{'Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Jumped to $1.06B at Q2-26 (from $74.5M at Q2-25); purpose not itemized in the CF statement.'),
('Property & equipment, net',{'Q2-24':328.3,'Q3-24':475.3,'Q4-24':846.7,'Q1-25':1334.1,'Q2-25':1789.4,'Q3-25':3314.4,'Q4-25':5553.3,'Q1-26':7131.7,'Q2-26':13045.2},
 {'Q2-24':'Q2_24_PR','Q3-24':'Q3_24_PR','Q4-24':'FY25_20F','Q1-25':'Q1_25_PR','Q2-25':'Q2_25_PR','Q3-25':'Q3_25_PR','Q4-25':'FY25_20F','Q1-26':'Q1_26_FS','Q2-26':'Q2_26_FS'},'Context for coverage ratios. Q4-24/Q4-25 = 20-F restated values.'),
('Deferred revenue / PP&E net (%)',{q:None for q in []},{},'computed in sheet'),
]
