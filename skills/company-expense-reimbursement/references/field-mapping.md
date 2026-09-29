# Workbook field mapping

## `报销单`

The first sheet is the print-ready summary. Each row is one travel leg or one linked lodging row.

| Column | Meaning |
|---|---|
| 起讫日期 | Travel date or date range |
| 起讫时间 | Start/end time when known |
| 起讫地点 | `A→B`; keep the direction visible |
| 车船费名称 | Taxi, train, flight, etc. |
| 车船费金额 | Transportation amount |
| 住宿费 | Hotel amount linked to the travel leg |
| 出差补助 | Manual company-approved allowance |
| 市内交通费 | Local rides not treated as intercity transport |
| 杂费用途 / 杂费金额 | Other approved expenses |
| 发票/附件索引 | Row or source reference in `发票明细` |
| 备注 | Missing dates, route warnings, or other review notes |

The total row uses formulas. Signature cells are deliberately blank.

## `发票明细`

One row per invoice or travel document. Keep source filename and page number, invoice/order number, dates, seller, buyer fields, amounts, origin/destination, hotel dates, category, and validation status. Amounts should be numeric cells when extracted with confidence; uncertain amounts stay blank and are explained in `备注`.

## `校验与说明`

This sheet stores the expected company master data, the run-level checklist, and the reminder that the template is copied for each claim.
