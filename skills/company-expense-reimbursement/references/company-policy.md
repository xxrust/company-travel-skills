# Company reimbursement rules

The expected buyer information is loaded from a local company profile. Do not put a real tax number, bank account, address, phone number, or employee identity in this public skill repository. See [local-profile.md](local-profile.md).

Compare values after trimming whitespace and treating obvious OCR punctuation variants as equivalent, but report any substantive difference.

## Rules supplied by the user

- Transportation is the dominant expense category; support taxi, air ticket/itinerary, and lodging documents.
- The paper form's 车船费 name field accepts only `车票`, `机票`, or `船票`. Never write a train number, flight number, route, seat class, or "公司车自驾" there; those details belong in the 发票明细 sheet and the validation report.
- A lodging document must show both check-in and check-out dates.
- If A→B occurs on a travel date and lodging is in B for consecutive days, lodging belongs on the A→B line.
- Multiple lodging invoices in B are written as multiple rows from top to bottom in date order.
- The travel path should form a closed loop. If it does not, warn the user.
- Leave all travel allowance amounts and their subtotal blank for finance. Leave the Chinese uppercase grand total blank for finance as well.
- Deliver the XLSX form for WPS. Verify a PDF exported by WPS with portrait A4, one page per form. Do not print without the user's instruction.
- Do not fill or print signatures; signatures are handwritten.

## Validation language

Use concrete statuses such as `已核对`, `购方信息不一致`, `待补入住/离店时间`, `路线未闭环`, `金额待确认`, and `来源无法读取`. Do not call a workbook complete when a required hotel date or company-field mismatch remains unresolved.
