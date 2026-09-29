---
name: travel-daily-report
description: "Create and maintain append-only daily business-trip reports tied to a shared case, with location, work performed, results, evidence, next plan, and issues. Use for in-trip progress reporting."
---

# Travel Daily Report

Use this skill when the user asks to write, update, review, or backfill a daily report during a business trip. Read [references/daily-report-fields.md](references/daily-report-fields.md) and use [assets/每日汇报模板.md](assets/每日汇报模板.md) when no company format is supplied.

## Workflow

1. Load the case's `case.yaml`. Verify that the report date belongs to the planned or actual trip window, or record the reason for an exception.
2. Create exactly one primary report file per calendar date under `02-出差过程/`, named `YYYY-MM-DD_每日汇报.md`. If the date already has a report, append an amendment section instead of overwriting the original facts.
3. Record the actual location and route, work completed, concrete results, contacts or meetings when evidenced, attachments, next plan, and blockers. Separate completed work from plans.
4. If information is missing, write `待补充` and add an issue. Do not invent customer names, measurements, acceptance results, or travel legs.
5. Register the report in `case.yaml.daily_reports` and add an event. A daily report does not by itself mean the trip is complete.

## Minimum quality checks

- The date, author, location, and report period are clear.
- Each stated result has a source, attachment, or user confirmation when it is material.
- The report's route does not silently rewrite the application; actual changes go to `actual_route` and an issue/event.
- A missed day is recorded as `未提交` or `待补充`, not silently skipped.
