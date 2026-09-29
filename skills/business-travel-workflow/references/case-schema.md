# 差旅案件共享数据结构

`case.yaml` 是案件的事实索引和状态入口。Excel、Markdown、PDF 和图片是输出或证据文件，不能替代案件索引。

## Minimum fields

```yaml
case_id: TRIP-2026-001
status: draft
person: "姓名"
department: "部门"
purpose: "出差目的"
planned_start: "2026-01-01"
planned_end: "2026-01-03"
planned_route: ["起点", "目的地", "起点"]
actual_start: null
actual_end: null
actual_route: []
budget: null
approver: null
application: {}
daily_reports: []
legs: []
documents: []
expenses: []
open_issues: []
events: []
```

## Field rules

- `planned_*` records what was requested or approved. `actual_*` records what evidence shows happened.
- Use ISO dates (`YYYY-MM-DD`) and 24-hour times when known. Preserve the original text in a document or note when the date is ambiguous.
- `planned_route` and `actual_route` are ordered lists. A route leg can be represented in `legs` as:

  ```yaml
  - leg_id: LEG-001
    date: "2026-01-01"
    start: "台州"
    end: "重庆"
    purpose: "现场安装"
    source_ids: [DOC-001]
  ```

- A daily report entry records `date`, `path`, `file`, `status`, and `source_ids`.
- A document entry records `document_id`, `path`, `kind`, `source_date`, and `related_leg_ids`. Invoice-specific fields remain in the reimbursement workbook's `发票明细` sheet.
- An issue records `issue_id`, `severity`, `status`, `description`, `evidence`, and `next_action`. Do not delete an issue; close it with a resolution and evidence.
- An event records `timestamp`, `actor`, `action`, `from_status`, `to_status`, and `evidence`. The event list is append-only.

## Truth hierarchy

1. Original approval, itinerary, invoice, meeting, or other source document.
2. A manually confirmed fact recorded in `case.yaml` with an event or note.
3. A derived total or summary generated from the case records.

If two sources conflict, preserve both, create an issue, and use `待人工复核` until the user resolves it.
