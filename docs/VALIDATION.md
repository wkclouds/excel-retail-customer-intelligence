# Validation record

## Completed automated checks

- ZIP integrity and original workbook SHA-256 recorded.
- Both worksheets read in full: 1,067,371 rows; no row truncation.
- Included purchases + returns + excluded monetary lines + removed cross-sheet overlap reconcile exactly to source rows.
- Positive purchase invoices checked to belong to one month and one country.
- Net sales reconcile across monthly, country, product and country-month aggregates.
- Distinct orders reconcile to the underlying positive invoice population.
- Customer detail and cohort initial populations reconcile to 5,878 identified purchasers.
- RFM scores stay within 1–5; retained cohort counts stay between zero and cohort size.
- Dashboard country/date controls changed and checked against independent preparation totals, then restored.
- Reversed start/end dates yield a visible invalid-range message and zero totals.
- Segment selector changed to At risk, its customer count verified, then restored to Champions.
- Formula error scan found no Excel errors in the authoring engine.
- Reconciliation differences are rounded to six decimal places to exclude floating-point summation noise.
- Every worksheet was rendered from the workbook and visually reviewed. The PNGs are actual rendered workbook views, not mockups.

Detailed preparation and control checks are in [validation.json](validation.json). Export checks are in [export_validation.json](export_validation.json).

## Native Excel checks still required

The laptop registry reported Microsoft Office ProPlus2024Retail, version 16.0.20326.20158, x64. A native Excel session was not available for completing these checks. Successful authoring-engine calculation and exported-file parsing do not prove desktop Excel behavior.

1. Open the workbook in desktop Excel and confirm there is no repair prompt.
2. Use the worksheet tabs listed on Start_Here and inspect each native chart.
3. Change the dashboard country and dates, then restore the defaults. Compare the full-scope gross value to £20,533,741.92.
4. Change the RFM segment selector and apply detail-table filters; confirm their stated scopes.
5. Confirm automatic calculation is enabled, and use Calculate Now if needed.
6. Save a local working copy and reopen it.

No Power Query, PivotTable, PivotChart, slicer or Data Model was created; none is claimed as tested. Refresh All is not a source-data refresh process. Rerun preparation and the builder to regenerate source-derived tables.

## Interpretation checks

Do not compare partial December 2011 to full December 2010. Do not sum distinct country customer counts to produce a global count. Do not interpret cohort first observation as true acquisition. Do not call gross RFM spend profit or predicted lifetime value. Exact repeated imports across overlapping source sheets are removed; other repeated lines remain included and their sensitivity is documented.
