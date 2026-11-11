# Financial Statistics Dashboard Systems

An Excel dashboard project for exploring income sources, financial targets, geographical sales performance, and the sales process.

**Project owner:** Srikanth Kamuni  
**Platform:** Microsoft Excel  
**Workbook:** `Financial Statistics Dashboard Systems.xlsm`

## Overview

This workbook brings financial and sales information into connected dashboard views. It combines source tables, PivotTables, formulas, charts, and year slicers to help users compare results against targets and explore the composition of income and sales.

The workbook includes four presentation sheets: Income Sources, Geographically, Sales Process, and Projects Status. The project-status page contains text-based figures and should be treated as a presentation page until its data connections are verified.

## Business questions

- Which income sources contribute most to total income?
- How does actual income compare with the income target?
- What are the average monthly income and operating profit?
- How much income comes from B2B and B2C activity?
- Which countries contribute most to sales, and how do results compare with targets?
- How are orders distributed across sales channels, payment methods, and delivery types?

## Dashboard features

### Income sources

- Analysis of usage fees, subscriptions, renting, licensing, asset sales, and advertising.
- Income breakdowns within each source.
- Actual income, target income, and achievement percentage.
- Quantity of items by income source.
- Average monthly income and operating profit.
- B2B and B2C income amounts and contributions.

### Geographical sales

- Sales analysis for Egypt, Russia, the USA, the United Kingdom, Canada, and Brazil.
- Country-level sales amounts and percentage contributions.
- Overall sales target achievement.
- Tax summaries displayed on the dashboard.

### Sales process

- Website and branch sales channels.
- Cash and credit-card payment methods.
- Customer registration status.
- Order assembly and sale status fields.
- Paid and refunded order information.
- Download, branch, and shipment delivery types.

### Project status

The Projects Status sheet presents departmental and project summary figures. The inspected figures appear as static drawing text; a verified source-to-dashboard calculation flow has not been established for this page.

## Workbook structure

| Worksheet | Purpose |
| --- | --- |
| Data Tables | Holds the income, country sales, and sales-process source tables. |
| Pivot tables | Aggregates source records and contains formulas that support dashboard metrics. |
| Income Sources | Presents income sources, target achievement, monthly income, operating profit, and B2B/B2C contributions. |
| Geographically | Presents country sales, sales contributions, target achievement, and tax summaries. |
| Sales Process | Presents sales channels, payments, registration, order status, and delivery information. |
| Projects Status | Contains project and departmental presentation figures requiring data-link verification. |

## Source data

The Data Tables sheet contains multiple datasets side by side. Treat each table as a separate source; worksheet row alignment does not establish a relationship between them.

| Dataset | Main fields |
| --- | --- |
| Income | Year, Month, Income sources, Income Breakdowns, Counts, Income, Target Income, operating profit, Marketing Strategies |
| Country sales | Year, Country, Amount, Target |
| Sales process | Order Number, Year, Month, POS, Payment Method, Assembly Stage, Registration Status, Sale Status, Delivery Type, Amount, Target |

Order numbers can repeat in the sales-process data. Confirm what each row represents before interpreting a row count as a count of unique orders or customers.

The workbook does not establish an externally verified data source or whether the records are synthetic. Its figures should be presented as project data, without claims about a real company's results.

## Excel techniques

- **PivotTables:** Aggregate financial and sales data for the dashboard views.
- **Charts:** Present contributions, comparisons, and performance measures.
- **Year slicers:** Filter the associated PivotTables by year.
- **GETPIVOTDATA:** Retrieve PivotTable results for dashboard calculations.
- **VLOOKUP:** Retrieve amounts and contributions from supporting summary ranges.
- **IFERROR:** Handle selected formula errors or unavailable results.
- **Shapes and linked text:** Display dashboard labels and financial summaries.

The file contains a VBA project. Its code and runtime behavior were not tested during the documentation review, so no specific macro automation is claimed here.

## How to use

1. Download `Financial Statistics Dashboard Systems.xlsm` from the repository after the workbook has been uploaded.
2. Open it in desktop Microsoft Excel. Desktop Excel is recommended for reviewing the workbook's slicers, shapes, and VBA features.
3. Start with the Income Sources sheet, then explore Geographically and Sales Process.
4. Use the available year slicer on each view to change its reporting period. Do not assume one slicer controls every dashboard.
5. Review the Pivot tables sheet to trace headline results to their supporting summaries.
6. Review Data Tables to understand the underlying fields and record structure.

Keep the `.xlsm` format when saving a working copy to retain the embedded VBA project. Review the VBA code and trust the file before enabling macros.

## Updating the data

1. Save a separate working copy before making changes.
2. Update the appropriate source table while preserving its headers, field types, and table boundaries.
3. Check that appended records fall within the relevant Excel table or PivotTable source range.
4. Use **Data → Refresh All** to refresh the workbook's summaries.
5. Check slicer selections and compare dashboard totals with the relevant source records.
6. Confirm that charts, formulas, and linked text reflect the refreshed results.

This refresh workflow needs testing in desktop Excel before relying on the workbook for recurring reporting.

## Metric definitions

| Metric | Interpretation |
| --- | --- |
| Total income | Sum of Income for the selected reporting scope. |
| Income target | Sum of Target Income for the same reporting scope. |
| Income achievement | Total income divided by the corresponding income target. |
| Average monthly income | Average of the monthly income totals included in the supporting summary. Check how months with no records are handled. |
| Operating profit | Sum of the source operating profit field; this does not independently validate the underlying accounting calculation. |
| B2B/B2C contribution | Income for each segment divided by total income for the same scope. |
| Country sales contribution | Country Amount divided by total Amount for the same scope. |
| Sales target achievement | Actual sales Amount divided by the corresponding Target. |

Income and sales metrics come from different source tables. Their totals should not be assumed to reconcile or be added together without a defined business relationship. Currency and units should also be confirmed before comparing views.

## Review findings and limitations

The documentation is based on inspection of workbook structure, stored values, formulas, drawing text, and embedded Excel objects. It is not a full accounting audit or a desktop Excel interaction test.

- A stored `#VALUE!` error exists in cell A1 of Geographically. Investigate its cause and any dependencies before release.
- The Projects Status figures appear to be static text rather than verified calculated results.
- VBA execution, navigation buttons, slicer interactions, and refresh behavior have not been tested in desktop Excel.
- Saved dashboard values can reflect a previous filter selection or refresh. Treat them as stored results until refreshed and reconciled.
- Source provenance, tax assumptions, currency consistency, and unique-order counting need confirmation.

## Suggested improvements

- Resolve the geographical-sheet error and validate affected results.
- Connect the project-status page to a documented project dataset.
- Define row grain, currency, and reporting scope for each source table.
- Reconcile headline KPIs against independent source totals.
- Test source expansion, refresh, and slicer connections in desktop Excel.
- Add verified dashboard screenshots and a short walkthrough to the repository.

## Repository description

Excel financial and sales dashboard with income analysis, target tracking, country sales, B2B/B2C contributions, and sales-process views using PivotTables, charts, formulas, and year slicers.

## License

No license is specified in this documentation. Add an appropriate license after confirming ownership and reuse rights for the workbook, data, and visual assets.
