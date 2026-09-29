# Local company and traveler profiles

The public skill contains rules and blank templates only. Private values are read at runtime from:

```text
%USERPROFILE%\.codex\private\company-profile.yaml
%USERPROFILE%\.codex\private\user-profile.yaml
```

Override the paths with `CODEX_EXPENSE_PROFILE` and `CODEX_USER_PROFILE`, or pass `--profile` and `--user-profile` to the workbook scripts.

## Company profile example

```yaml
name: ""
tax_id: ""
address: ""
phone: ""
bank_name: ""
bank_account: ""
```

## User profile example

```yaml
name: ""
department: ""
title: ""
employee_id: ""
```

If no profile is found, the generated workbook keeps these fields blank or labels them as local configuration fields. Never infer them from an invoice filename or from a previous case.
