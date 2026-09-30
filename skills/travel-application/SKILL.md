---
name: travel-application
description: "Prepare or revise a business-trip application with planned purpose, route, dates, budget, approval fields, and traceable supporting documents. Keep planned facts separate from actual trip facts."
---

# Travel Application

Use this skill when the user asks to create, complete, review, or revise an application for a planned business trip. When the request is part of a multi-phase case, load the existing `case.yaml` and write the application artifact under `01-出差申请/`.

Read [references/application-fields.md](references/application-fields.md) before drafting the form. When the company form is available, use [assets/公司出差申请单模板.docx](assets/公司出差申请单模板.docx) as the primary template and preserve its layout. Use [assets/出差申请模板.md](assets/出差申请模板.md) as the fallback when the company form is unavailable.

## Workflow

1. Load the case record and retain its `case_id`. If no case exists, ask the workflow skill to create one or create a case folder before drafting the application.
2. Collect planned purpose, department, traveler, destination and route, planned dates, transport/lodging plan, budget, expected outputs, approver, and attachments. Keep unknown fields marked `待补充`.
3. Distinguish `planned_route` and `actual_route`. Do not update actual travel facts from an application.
4. Produce the requested document format. Preserve the company form when supplied; otherwise use the fallback template. Save the working copy under `01-出差申请/`. Never overwrite the reusable template with a filled application.
5. Record the artifact in `case.yaml.documents` and add an event. Set the case to `applied` only when the user confirms that the application was submitted. Set `approved` only with an approval source or explicit user confirmation.

## Boundaries

- Do not submit an approval request, send email, or sign on the user's behalf unless a separate explicit authorization and tool route exists.
- Do not promise reimbursement eligibility from the application. Eligibility is checked against actual evidence during closure and reimbursement.
- Do not silently replace a user-provided route, date, budget, or approver. Put conflicts in `issues.md`.
