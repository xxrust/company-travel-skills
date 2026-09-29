---
name: travel-closure
description: "Close a completed business trip by reconciling planned and actual routes, summarizing work results, checking return evidence and open issues, and preparing the case for reimbursement."
---

# Travel Closure

Use this skill when the user asks to finish, summarize, archive, or prepare a business trip for reimbursement. Read [references/closure-checklist.md](references/closure-checklist.md) and start from [assets/出差总结模板.md](assets/出差总结模板.md) unless the company has a required form.

## Workflow

1. Load `case.yaml`, the application, all daily reports, route evidence, and the current `issues.md`.
2. Compare planned and actual dates, route, destinations, work scope, and outcomes. Preserve both values and explain every material deviation.
3. Check route closure. The actual path must return to the initial origin. If the return leg is absent or unsupported, write `路线未闭环`, identify the missing evidence, and keep the case out of `ready_for_reimbursement`.
4. Check that each daily report in the travel window is present or explicitly marked `未提交`/`待补充`.
5. Summarize completed work, measurable results, customer or internal confirmation, follow-up work, and unresolved risks. Do not convert plans into results.
6. Link the closeout document under `04-出差结束/`, add an event, and advance the state only when the checklist permits it. The final reimbursement workbook is created by `company-expense-reimbursement`.

## Boundaries

- Do not fabricate a return trip, meeting result, acceptance, or approval to make the case pass.
- Do not delete unresolved issues after writing the summary. Carry them into the reimbursement or follow-up stage.
- Do not fill signature fields. Any paper signature remains a manual action.
