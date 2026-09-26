# Data dictionary

Original headers are preserved in the raw workbook. Preparation normalizes the labels shown below.

| Source column | Prepared name | Meaning and handling |
| --- | --- | --- |
| `Invoice` | `InvoiceNo` | Invoice identifier, treated as text; `C` prefix means cancellation. |
| `StockCode` | `StockCode` | Text product/service code. Not all codes are merchandise. |
| `Description` | `Description` | Item name. Missing labels do not invalidate a code. |
| `Quantity` | `Quantity` | Signed units on the observed line. |
| `InvoiceDate` | `InvoiceDate` | Invoice timestamp with no supplied timezone. |
| `Price` | `UnitPrice` | GBP per unit. Nonpositive values excluded from monetary metrics. |
| `Customer ID` | `CustomerID` | Anonymized identifier stored as text; missing IDs omitted from customer analysis. |
| `Country` | `Country` | Customer country of residence, not a city or delivery region. |

## Prepared table grains

| File | One row represents | Key |
| --- | --- | --- |
| `monthly.csv` | Calendar month | Month |
| `country_month.csv` | Country in a calendar month | Month + Country |
| `regional.csv` | Country | Country |
| `products.csv` | Stock code | StockCode |
| `customers_rfm.csv` | Identified positive purchaser | CustomerID |
| `segments.csv` | RFM rule segment | Segment |

`GrossSales`, `ReturnValue`, `NetSales` and `Monetary` are GBP. `Orders` and `Frequency` count distinct positive purchase invoices. `PurchasedUnits` and `ReturnedUnits` use their respective eligible populations. `IdentifiedCustomers` is a distinct count within the group, not an additive measure. `RevenueShare` is a ratio; `CumulativeShare` follows descending gross sales. `Rank` is one-based and descriptive.

Customer fields `FirstPurchase` and `LastPurchase` are observed timestamps. `RecencyDays` uses the fixed reference date. `R`, `F`, `M` are integer scores 1–5. `Cohort` is first observed purchase month. `Country` in the customer table is the latest positive purchase country.

`analysis.json` is the workbook builder's prepared input, containing the same aggregates plus cohort counts, quality metrics, scoring thresholds and provenance. `source_profile.json` records source-sheet coverage and repeated-row diagnostics. `cross_sheet_overlap_removed.csv` preserves the exact source lines removed from the second sheet's repeated import. These are derived from the original records, not fabricated examples.
