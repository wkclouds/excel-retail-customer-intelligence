# Methodology

## Population and grain

Every row from both original sheets is profiled: 525,461 rows in `Year 2009-2010`, 541,910 in `Year 2010-2011`. Their date ranges overlap from 1–9 December 2010. Preparation removes an exact cross-sheet import match only when all eight original fields and the within-sheet occurrence number match. This preserves repeated lines within a sheet while keeping one copy of the repeated import. Removed lines are exported to `cross_sheet_overlap_removed.csv` for audit. Source grain is an observed invoice line; the file has no unique line identifier. Invoice references are text. A leading `C` is detected case-insensitively.

An eligible purchase has a valid invoice/date/country, positive quantity, positive unit price and no cancellation prefix. An eligible return/adjustment has valid keys, positive price, nonzero quantity and either negative quantity or a cancellation prefix. Populations are mutually exclusive. All other rows are excluded from monetary calculations. Exactly 1,019,654 purchase rows + 19,165 return rows + 6,029 excluded monetary rows + 22,523 repeated import rows = 1,067,371 original input rows. The analysis population before monetary exclusions is 1,044,848 rows.

## Metrics

- Gross sales: sum of quantity × price across eligible purchase rows.
- Return value: sum of absolute quantity × price across eligible return/adjustment rows. It is reported positive.
- Net sales: gross sales minus return value. These are analytical sales values, not an audited accounting revenue statement.
- Orders: distinct invoice references containing at least one eligible purchase line.
- Customers: distinct non-missing customer IDs among eligible purchases.
- Gross AOV: gross sales / positive purchase orders. Return invoices are excluded from the denominator; returns are excluded from the numerator.
- Repeat customer rate: identified purchasers with at least two distinct eligible purchase invoices / identified purchasers.
- Return ratio: return value / gross sales; not the fraction of orders returned.
- YoY: Jan–Nov 2011 net sales / Jan–Nov 2010 net sales − 1. December 2011 is incomplete.

Source amounts retain full numerical precision; displayed amounts round to two decimals. Aggregate numerical reconciliations allow £0.000001 for floating-point addition. Counts must match exactly.

## Missing values, repetition and anomalies

Missing customer IDs do not invalidate sales. Their valid purchase lines generate £3,102,190.740 gross sales and are excluded only from customer-based analysis. Missing descriptions do not remove known stock codes. The last available description is displayed per stock code; this is a label choice, not an invented product mapping.

Zero and negative prices are excluded from monetary analysis. Negative quantities without `C` are candidate adjustments, not proven consumer returns. A positive-quantity cancellation is still assigned to return value under the explicit prefix rule. This differs from simply summing all signed quantity × price values.

Exact duplicate checks compare all eight original columns before normalization. The source sheets' overlapping import is removed using the occurrence-matched rule above. Remaining repeated lines are retained because the source has no line identifier establishing whether they are errors. Their net amount is reported as sensitivity. `source_profile.json` records each sheet's date range, raw repeated count, cross-sheet overlap removed and remaining repeated count. The cross-sheet rule does not justify deleting all identical-looking purchases.

## RFM

Population: 5,878 identified eligible purchasers. Reference date: 2011-12-10, one day after the last observation. Recency is the number of calendar days from the latest purchase date to this reference date. Frequency is distinct purchase invoices over the whole window. Monetary is eligible gross purchase spend, not net revenue or lifetime value.

Use empirical 20th, 40th, 60th and 80th percentile cut points with pandas' default linear interpolation. A value exactly on a threshold remains in the lower raw band. Frequency and monetary bands receive scores 1–5. Recency scores reverse to 5–1 so a more recent customer has the larger score. Ties remain together; scores need not contain equal counts.

| Metric | P20 | P40 | P60 | P80 |
| --- | ---: | ---: | ---: | ---: |
| Recency days | 20 | 59 | 190 | 411 |
| Distinct invoices | 1 | 2 | 4 | 8 |
| Gross spend GBP | 287.310 | 612.116 | 1,231.758 | 2,916.438 |

Evaluate these segment rules in order; first match wins:

1. Champions: R ≥ 4 and F ≥ 4 and M ≥ 4.
2. At risk: R ≤ 2 and F ≥ 3.
3. Loyal: R ≥ 3 and F ≥ 3.
4. Recent light buyers: R ≥ 4 and F ≤ 2.
5. Hibernating: R ≤ 2.
6. Developing: all remaining customers.

Names are descriptive rule labels. They do not imply churn prediction, causal responsiveness or employment/client outcomes. Country is the customer's latest observed positive purchase country.

## Cohorts

Assign each identified purchaser to the month of their earliest eligible purchase **observed in this file**. Deduplicate customer-month activity, then count distinct cohort members active at each integer month age. The denominator is the original cohort size. Returning in month 3 does not require activity in months 1 and 2. Blank cells represent future unobserved ages; observed zero activity would be numeric zero. The final calendar month is partial for every cohort.

## Product, regional and source tables

Products are grouped by stock code and ranked by gross sales. Pareto shares divide gross code revenue by total gross revenue. The 80% count is the first code count whose cumulative share reaches 80%. Return-only and zero-sales codes stay in the detailed table but are excluded from the positive-sales-code denominator. Stock codes may represent charges or adjustments as well as merchandise.

Country comparisons use source country labels. A customer can appear in more than one country, so country customer counts must not be summed to obtain global unique customers. Each positive invoice is checked to occupy one month and one country, making order counts additive across the country-month table. Customer cohorts use a many-to-one join to a unique customer lookup; no transaction joins multiply line amounts.

## Forecasting and refresh

No forecast is delivered. Two years of history, the bounded customer window and incomplete final month require a carefully evaluated time-based holdout before a forecast would be credible. No synthetic prices, costs, margins or demand are added.

The CSVs and customer scores are prepared snapshots. Rebuilding from the source recomputes them. Excel formulas update selected dashboard aggregates, RFM summaries and cohort rates, but do not rerun Python preparation. The workbook contains no macros or network refresh connections.
