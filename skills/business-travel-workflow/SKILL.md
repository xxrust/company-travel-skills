---
name: business-travel-workflow
description: "Coordinate a complete company business trip as one case from application through daily reports, trip closure, and reimbursement. Maintain shared case state, route work to the appropriate travel skills, enforce phase gates, and surface unresolved issues."
---

# Business Travel Workflow

This is the user-facing entry point for a complete business trip. Use it when the user asks to create, update, report, close, or reimburse one trip and the work spans more than one phase.

Colleagues use this workflow through natural-language requests. When they ask to find trip invoices, scan enterprise mail, or prepare reimbursement, the agent should invoke the phase skills and local scripts internally. Do not expose PowerShell or Python commands as a prerequisite for ordinary users. Ask only for missing trip dates, destination, or business facts; credentials remain in the local machine profile.

The unit of work is a **差旅案件**. One agent coordinates the case; the phase skills provide focused instructions:

- `travel-application` for the planned trip and approval material.
- `travel-daily-report` for one append-only report per day during the trip.
- `travel-closure` for actual-route reconciliation, results, and closeout.
- `company-expense-reimbursement` for invoice extraction, validation, and the reimbursement workbook.
- `$mineru-pdf-to-md` for the document-extraction step required by the reimbursement skill.

Read these references before creating or changing a case:

- [case-schema.md](references/case-schema.md) — the shared `case.yaml` contract.
- [state-machine.md](references/state-machine.md) — allowed states and transition gates.
- [case-layout.md](references/case-layout.md) — the folder and artifact convention.

## Operating procedure

1. Locate the case folder from the user's path, case ID, or artifact names. If no case exists, create one with:

   ```powershell
   python scripts/new_case.py --output "path\to\差旅案件" --case-id "TRIP-YYYY-001" --person "姓名"
   ```

2. Load `case.yaml` and `issues.md` before acting. Keep planned facts, actual facts, source documents, and unresolved issues separate.
3. Determine the user's requested phase and the current case state. Do not advance the state merely because a document was created; advance it only when the transition gate is satisfied or the user explicitly records the external approval/action.
4. Apply the relevant phase skill. When several outputs are requested, complete them in lifecycle order and update the shared case record after each phase.
5. Record every generated artifact in `case.yaml.documents` or the relevant phase list. Keep source filenames and dates so an amount, route leg, or report statement can be traced.
6. Return the artifact paths, the new case state, and all open issues. A case with unresolved required fields must remain visibly incomplete.

## Invariants

- Do not invent destinations, dates, meetings, outputs, approvals, return legs, or amounts. Use `待补充` or an open issue when evidence is missing.
- Preserve planned and actual routes as separate fields. A change in the actual route does not silently rewrite the application.
- Daily reports are append-only. Correct a prior report with an amendment that points to the original date; do not erase the earlier record.
- The travel path must close back to its starting point before closure or reimbursement is presented as ready. If evidence is missing, write `路线未闭环` with the missing leg.
- Never create electronic signatures or claim that an approval/signature happened without source evidence. Leave signature fields for the user's handwritten or external approval process.
- Reimbursement is a downstream phase. It must link each invoice to a case leg or explicitly mark the link as unresolved.

## Completion language

Use precise states such as `草稿`, `申请已提交`, `申请已批准`, `出差中`, `待收尾`, `待报销`, `已关闭`. Use an `issues.md` entry for `待补充`, `待人工复核`, `路线未闭环`, `发票信息不一致`, or other blockers. Do not describe the case as complete while a required gate is open.
