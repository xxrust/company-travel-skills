# Workbook field mapping

## `报销单`

The first sheet is a print-ready copy of the company's paper form. One page has six fixed detail rows. Each row is one travel leg or one linked lodging row. When more than six rows are needed, create another sheet with the same page layout and continue the rows there; do not extend the first page vertically.

| Column | Meaning |
|---|---|
| 起讫日期 | Day fields in the fixed paper-form row |
| 起讫地点 | `A→B`; keep the direction visible |
| 车船费名称 / 金额 | Taxi, train, flight, company car, etc. |
| 住宿费 | Hotel amount linked to the travel leg |
| 出差补助 | Manual company-approved allowance |
| 市内交通费 | Local rides not treated as intercity transport |
| 杂费用途 / 金额 | Other approved expenses |
| 附注 | Source index, dates, missing fields, or route warnings |

The total row uses formulas. Signature cells are deliberately blank. The paper form is configured for landscape A4 printing.

## `发票明细`

One row per invoice or travel document. Keep source filename and page number, invoice/order number, dates, seller, buyer fields, amounts, origin/destination, hotel dates, category, and validation status. Amounts should be numeric cells when extracted with confidence; uncertain amounts stay blank and are explained in `备注`.

## `校验与说明`

This sheet stores the expected company master data, the run-level checklist, and the reminder that the template is copied for each claim.
