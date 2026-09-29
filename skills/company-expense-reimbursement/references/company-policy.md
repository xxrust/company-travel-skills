# Company reimbursement rules

The expected buyer information is loaded from a local company profile. Do not put a real tax number, bank account, address, phone number, or employee identity in this public skill repository. See [local-profile.md](local-profile.md).

Compare values after trimming whitespace and treating obvious OCR punctuation variants as equivalent, but report any substantive difference.

## Rules supplied by the user

- Transportation is the dominant expense category; support taxi, air ticket/itinerary, and lodging documents.
- A lodging document must show both check-in and check-out dates.
- If A→B occurs on a travel date and lodging is in B for consecutive days, lodging belongs on the A→B line.
- Multiple lodging invoices in B are written as multiple rows from top to bottom in date order.
- The travel path should form a closed loop. If it does not, warn the user.
- Do not fill or print signatures; signatures are handwritten.

## Validation language

Use concrete statuses such as `已核对`, `购方信息不一致`, `待补入住/离店时间`, `路线未闭环`, `金额待确认`, and `来源无法读取`. Do not call a workbook complete when a required hotel date or company-field mismatch remains unresolved.
