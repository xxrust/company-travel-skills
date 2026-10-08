"""Amount helpers for the paper reimbursement workbook."""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Iterable


_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_UNITS = ("", "拾", "佰", "仟")
_GROUP_UNITS = ("", "万", "亿", "兆")


def _group_to_upper(group: int) -> str:
    result = []
    zero_pending = False
    for index in range(3, -1, -1):
        digit = (group // (10**index)) % 10
        if digit:
            if zero_pending and result:
                result.append("零")
            result.append(_DIGITS[digit] + _UNITS[index])
            zero_pending = False
        elif result:
            zero_pending = True
    return "".join(result)


def amount_to_chinese_upper(amount: Decimal | int | float | str) -> str:
    """Convert a CNY amount to standard Chinese uppercase currency text."""
    value = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if value < 0:
        return "负" + amount_to_chinese_upper(-value)
    integer, fraction = divmod(int(value * 100), 100)
    groups = []
    group_index = 0
    while integer:
        groups.append((integer % 10000, group_index))
        integer //= 10000
        group_index += 1
    if not groups:
        integer_text = "零"
    else:
        pieces = []
        zero_between = False
        for group, index in reversed(groups):
            if group == 0:
                zero_between = bool(pieces)
                continue
            if zero_between or (pieces and group < 1000):
                pieces.append("零")
            pieces.append(_group_to_upper(group) + _GROUP_UNITS[index])
            zero_between = False
        integer_text = "".join(pieces)
    jiao, fen = divmod(fraction, 10)
    result = "人民币" + integer_text + "元"
    if jiao:
        result += _DIGITS[jiao] + "角"
    elif fen:
        result += "零"
    if fen:
        result += _DIGITS[fen] + "分"
    elif not jiao:
        result += "整"
    return result


def paper_total(ws, rows: Iterable[int] = range(9, 15)) -> Decimal:
    """Sum numeric expense cells, excluding the finance-owned allowance."""
    total = Decimal("0")
    for row in rows:
        for column in (8, 9, 11, 13):
            value = ws.cell(row, column).value
            if isinstance(value, (int, float, Decimal)):
                total += Decimal(str(value))
    return total


def fill_paper_total(ws) -> None:
    """Leave the uppercase total for finance to complete by hand."""
    return None
