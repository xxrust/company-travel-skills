---
name: company-expense-reimbursement
description: "Create and validate company travel reimbursement workbooks from invoice PDFs, invoice links, scans, or travel details. Use MinerU for supplied PDFs/images, load private company and traveler profiles locally, copy the reusable Excel template, and flag invoice-field errors, missing hotel stay dates, and non-closed routes."
---

# Company Expense Reimbursement

Use this skill when the user sends an invoice PDF, an invoice address/link, an invoice scan/photo, or asks to prepare or check a company travel reimbursement. The deliverable is a copied Excel workbook plus a concise validation report.

If an invoice address is a directly downloadable PDF or image, save a local copy in the current case folder and process it like an uploaded file. If it requires a login or cannot be fetched, ask the user to upload the document instead of guessing from the URL.

When invoked from `business-travel-workflow`, load the case's `case.yaml` first. Link each invoice or travel document to a `case_id` and, when possible, a specific `leg_id` or lodging stay. Do not change planned or actual route fields from invoice text alone; return route/date conflicts to the workflow as open issues.

## Required workflow

1. Read [references/company-policy.md](references/company-policy.md) and [references/local-profile.md](references/local-profile.md) before validating invoice fields or travel rules. Company and traveler identity data must come from the local profile, never from this public repository.
2. For every supplied PDF or page image, run the local MinerU workflow from `$mineru-pdf-to-md` first. Preserve the source filename and page number. Do not silently replace MinerU with a generic OCR/text extractor. If the local model is unavailable, report the blocker.
   If MinerU's Markdown has obvious character corruption or misses a field, use the PDF's embedded text layer only as a cross-check, keep the MinerU artifact, and mark the affected field as `需人工复核` when the two sources disagree.
3. Create a new workbook. The script copies the blank public template when no local profile is found, and generates a locally populated copy when a profile is found:

   ```powershell
   python scripts/new_workbook.py --output "path\to\报销单_姓名_YYYYMMDD.xlsx" \
     --profile "C:\Users\<user>\.codex\private\company-profile.yaml" \
     --user-profile "C:\Users\<user>\.codex\private\user-profile.yaml"
   ```

   Keep the template unchanged; each reimbursement gets its own copy.
4. Put one source invoice or travel document per row on `发票明细`. Put the final reimbursement rows on `报销单`. Use the invoice/source index in `报销单` so every amount can be traced back to a source.
5. Preserve uncertainty. Never infer an origin, destination, travel date, hotel stay date, or amount from a filename, invoice issue date, or an example image when the source does not state it. Leave the cell blank and flag it for the user.
6. Before completion, report all validation findings. A field mismatch in the company name, tax number, address, phone, bank, or account is an immediate warning and must not be silently corrected.

## Expense mapping and travel rules

- Classify taxi and other local rides as `市内交通费` unless the user says the route is intercity; classify air tickets or itinerary receipts as `车船费`; classify hotel invoices as `住宿费`.
- A hotel invoice is incomplete for reimbursement unless it contains both check-in and check-out dates. The amount may be entered for review, but the row must be marked `待补入住/离店时间`.
- If travel from A to B happened on a date and the employee stayed in B for consecutive days, put the lodging amount on the same A→B row. If B has multiple hotel invoices, add rows from top to bottom in date order and keep each source index visible.
- Check the route as a sequence. The last destination must return to the initial origin. If the evidence does not form a closed loop, issue `路线未闭环` and show the missing or unmatched leg. Do not manufacture a return leg.
- Leave all signature cells blank. The reimbursement form is intentionally unsigned; the traveler and reviewers sign by hand after printing.

## Output requirements

- `报销单` contains the print-ready summary and formulas for category totals.
- `发票明细` contains extracted fields, source/page traceability, hotel stay fields, route fields, and validation status.
- `校验与说明` contains the locally loaded company master data and the rules used for the run. If no local profile is available, leave identity fields blank and report that configuration is missing.
- State which PDFs/pages were converted, which fields were confirmed, and every unresolved issue. If an invoice contains an address or other company field that differs from the master data, call it out before presenting the workbook as ready.

The template layout and column semantics are documented in [references/field-mapping.md](references/field-mapping.md).
