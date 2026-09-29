#!/usr/bin/env python
"""Build the reusable Excel template used by the reimbursement skill."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Mapping

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

from reimbursement_profile import load_company_profile, load_user_profile


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "assets" / "company-expense-template.xlsx"

BLUE = "1F4E78"
LIGHT_BLUE = "D9EAF7"
PALE_YELLOW = "FFF2CC"
PALE_RED = "FCE4D6"
WHITE = "FFFFFF"
GRAY = "666666"
thin = Side(style="thin", color="7F8C8D")
border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_range(ws, cell_range: str, *, fill=None, font=None, alignment=None, add_border=True):
    for row in ws[cell_range]:
        for cell in row:
            if fill:
                cell.fill = fill
            if font:
                cell.font = font
            if alignment:
                cell.alignment = alignment
            if add_border:
                cell.border = border


def make_workbook(
    company: Mapping[str, str] | None = None,
    user: Mapping[str, str] | None = None,
) -> Workbook:
    company = dict(company or {})
    user = dict(user or {})
    wb = Workbook()
    ws = wb.active
    ws.title = "报销单"
    detail = wb.create_sheet("发票明细")
    check = wb.create_sheet("校验与说明")

    title_font = Font(name="Microsoft YaHei", size=18, bold=True, color=BLUE)
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color=WHITE)
    body_font = Font(name="Microsoft YaHei", size=10, color="000000")
    small_font = Font(name="Microsoft YaHei", size=9, color=GRAY)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    blue_fill = PatternFill("solid", fgColor=BLUE)
    light_fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    yellow_fill = PatternFill("solid", fgColor=PALE_YELLOW)
    red_fill = PatternFill("solid", fgColor=PALE_RED)

    # Print-ready summary
    ws.merge_cells("A1:L1")
    ws["A1"] = "差旅费报销单"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:L2")
    company_name = company.get("name") or "公司名称（本地配置）"
    ws["A2"] = f"{company_name}｜模板副本（每次报销请复制后填写）"
    ws["A2"].font = small_font
    ws["A2"].alignment = center

    labels = [("A4", "部门"), ("D4", "姓名"), ("F4", "职别"), ("H4", "出差事由"), ("A5", "出差起止日期"), ("F5", "共"), ("H5", "附件单据")]
    for addr, text in labels:
        ws[addr] = text
        ws[addr].font = Font(name="Microsoft YaHei", size=10, bold=True, color=BLUE)
        ws[addr].fill = light_fill
        ws[addr].border = border
        ws[addr].alignment = center
    for merge in ("B4:C4", "E4:E4", "G4:G4", "I4:L4", "B5:E5", "G5:G5", "I5:L5"):
        ws.merge_cells(merge)
    for row in ws["A4:L5"]:
        for cell in row:
            cell.border = border
            cell.font = Font(name="Microsoft YaHei", size=10, bold=bool(cell.font.bold), color=BLUE if cell.column in (1, 4, 6, 8) else "000000")
            cell.alignment = left if cell.column not in (1, 4, 6, 8) else center
    ws["B4"].fill = yellow_fill
    ws["E4"].fill = yellow_fill
    ws["G4"].fill = yellow_fill
    ws["I4"].fill = yellow_fill
    ws["B5"].fill = yellow_fill
    ws["G5"].fill = yellow_fill
    ws["I5"].fill = yellow_fill
    ws["B4"] = user.get("department", "")
    ws["E4"] = user.get("name", "")
    ws["G4"] = user.get("title", "")
    ws["F5"] = "共"
    ws["H5"] = "附件单据"
    ws["G5"] = ""
    ws["I5"] = ""

    headers = ["起讫日期", "起讫时间", "起讫地点", "车船费名称", "车船费金额", "住宿费", "出差补助", "市内交通费", "杂费用途", "杂费金额", "发票/附件索引", "备注"]
    for col, value in enumerate(headers, 1):
        cell = ws.cell(7, col, value)
        cell.fill = blue_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    for row in range(8, 20):
        for col in range(1, 13):
            cell = ws.cell(row, col)
            cell.font = body_font
            cell.alignment = left if col in (3, 4, 9, 11, 12) else center
            cell.border = border
            if col in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12):
                cell.fill = yellow_fill if col not in (5, 6, 7, 8, 10) else PatternFill(fill_type=None)
            if col in (5, 6, 7, 8, 10):
                cell.number_format = '#,##0.00'
    ws.merge_cells("A20:D20")
    ws["A20"] = "合计"
    ws["A20"].font = Font(name="Microsoft YaHei", bold=True, color=BLUE)
    ws["A20"].alignment = center
    for col in range(1, 13):
        ws.cell(20, col).border = border
        ws.cell(20, col).fill = light_fill
    for col in (5, 6, 7, 8, 10):
        letter = chr(64 + col)
        ws.cell(20, col).value = f"=SUM({letter}8:{letter}19)"
        ws.cell(20, col).number_format = '#,##0.00'
        ws.cell(20, col).font = Font(name="Microsoft YaHei", bold=True)
    ws["K20"] = "总计"
    ws["K20"].alignment = center
    ws["L20"] = "=SUM(E20:H20,J20)"
    ws["L20"].number_format = '#,##0.00'
    ws["L20"].font = Font(name="Microsoft YaHei", bold=True)
    ws.merge_cells("A22:C22")
    ws["A22"] = "合计金额（大写，打印后手工填写）"
    ws.merge_cells("D22:I22")
    ws["D22"].fill = yellow_fill
    ws["J22"] = "¥"
    ws["K22"] = ""
    ws.merge_cells("A24:D24")
    ws["A24"] = "单位主管："
    ws.merge_cells("E24:H24")
    ws["E24"] = "复核："
    ws.merge_cells("I24:L24")
    ws["I24"] = "出差人："
    for row in ws["A22:L24"]:
        for cell in row:
            cell.border = border
            cell.font = body_font
            cell.alignment = left
    ws["A22"].font = small_font
    ws["A24"].font = body_font
    ws["E24"].font = body_font
    ws["I24"].font = body_font
    ws.freeze_panes = "A8"
    ws.sheet_view.showGridLines = False
    widths = {"A": 14, "B": 12, "C": 22, "D": 14, "E": 12, "F": 12, "G": 12, "H": 14, "I": 14, "J": 12, "K": 16, "L": 26}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in range(8, 20):
        ws.row_dimensions[r].height = 32
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.35, bottom=0.35, header=0.1, footer=0.1)
    ws.print_area = "A1:L24"

    # Invoice detail sheet
    detail.sheet_view.showGridLines = False
    detail_headers = [
        "来源文件", "页码", "单据类型", "发票号码/订单号", "开票/乘车/入住日期", "结束日期",
        "购方名称", "购方税号", "购方地址", "购方电话", "购方开户行", "购方账号",
        "销方/承运方/酒店", "税率", "不含税金额", "税额", "价税合计", "出发地", "到达地",
        "入住时间", "离开时间", "分类", "可报销?", "校验状态", "备注",
    ]
    for col, value in enumerate(detail_headers, 1):
        cell = detail.cell(1, col, value)
        cell.fill = blue_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    for row in range(2, 102):
        for col in range(1, len(detail_headers) + 1):
            cell = detail.cell(row, col)
            cell.font = body_font
            cell.border = border
            cell.alignment = left if col in (1, 4, 7, 9, 13, 18, 19, 24, 25) else center
            if col in (15, 16, 17):
                cell.number_format = '#,##0.00'
    for col, width in {1: 34, 2: 8, 3: 14, 4: 24, 5: 16, 6: 16, 7: 24, 8: 22, 9: 42, 10: 14, 11: 18, 12: 24, 13: 30, 14: 10, 15: 14, 16: 14, 17: 14, 18: 18, 19: 18, 20: 16, 21: 16, 22: 14, 23: 12, 24: 20, 25: 40}.items():
        detail.column_dimensions[chr(64 + col) if col <= 26 else f"A{col}"].width = width
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = "A1:Y101"
    category_validation = DataValidation(type="list", formula1='"出租车,飞机/行程单,住宿,火车/动车,其他"', allow_blank=True)
    yes_validation = DataValidation(type="list", formula1='"是,否,待确认"', allow_blank=True)
    status_validation = DataValidation(type="list", formula1='"已核对,购方信息不一致,待补入住/离店时间,路线未闭环,金额待确认,来源无法读取"', allow_blank=True)
    detail.add_data_validation(category_validation)
    detail.add_data_validation(yes_validation)
    detail.add_data_validation(status_validation)
    category_validation.add("V2:V101")
    yes_validation.add("W2:W101")
    status_validation.add("X2:X101")
    detail.conditional_formatting.add("X2:X101", FormulaRule(formula=['AND($X2<>"",$X2<>"已核对")'], fill=red_fill))

    # Check and instructions sheet
    check.sheet_view.showGridLines = False
    check.merge_cells("A1:D1")
    check["A1"] = "报销校验与填写说明"
    check["A1"].font = title_font
    check["A1"].alignment = center
    for col, value in enumerate(["字段", "公司标准值", "本次票据值", "结果/备注"], 1):
        cell = check.cell(3, col, value)
        cell.fill = blue_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    company_fields = [
        ("公司名称", company.get("name", "")),
        ("税号", company.get("tax_id", "")),
        ("地址", company.get("address", "")),
        ("电话", company.get("phone", "")),
        ("开户行", company.get("bank_name", "")),
        ("账号", company.get("bank_account", "")),
    ]
    for row, (field, value) in enumerate(company_fields, 4):
        check.cell(row, 1, field)
        check.cell(row, 2, value)
        check.cell(row, 3, "")
        check.cell(row, 4, "待核对（本地配置）" if value else "待填写本地配置")
        for col in range(1, 5):
            check.cell(row, col).border = border
            check.cell(row, col).alignment = left
            check.cell(row, col).font = body_font
            if col in (2, 3):
                check.cell(row, col).fill = yellow_fill
    check.merge_cells("A12:D12")
    check["A12"] = "必须检查"
    check["A12"].fill = blue_fill
    check["A12"].font = header_font
    check["A12"].alignment = center
    rules = [
        "每个 PDF/图片先用本地 MinerU 解析，并在发票明细保留来源文件和页码。",
        "住宿票必须同时有入住和离开日期；缺任一日期就标记待补，不得当作完整凭证。",
        "A→B 出差当天的 B 地住宿费写在 A→B 行；B 地多张住宿票按日期从上到下分行。",
        "路线必须闭环；缺少回程证据时标记路线未闭环，不补写虚构行程。",
        "报销单不打印签名，单位主管、复核、出差人签名留待打印后手写。",
    ]
    for row, rule in enumerate(rules, 13):
        check.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        check.cell(row, 1, f"{row - 12}. {rule}")
        check.cell(row, 1).alignment = left
        check.cell(row, 1).font = body_font
        for col in range(1, 5):
            check.cell(row, col).border = border
    check.column_dimensions["A"].width = 20
    check.column_dimensions["B"].width = 48
    check.column_dimensions["C"].width = 48
    check.column_dimensions["D"].width = 30
    check.freeze_panes = "A4"
    check.page_setup.orientation = "landscape"
    check.page_setup.fitToWidth = 1
    check.sheet_properties.pageSetUpPr.fitToPage = True

    return wb


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a blank or locally populated reimbursement template.")
    parser.add_argument("--output", default=str(OUTPUT), help="Destination .xlsx path")
    parser.add_argument("--profile", help="Local company profile YAML")
    parser.add_argument("--user-profile", help="Local traveler profile YAML")
    args = parser.parse_args()

    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    company, company_path = load_company_profile(args.profile, discover=bool(args.profile))
    user, user_path = load_user_profile(args.user_profile, discover=bool(args.user_profile))
    wb = make_workbook(company, user)
    wb.save(output)
    source_note = f"company_profile={company_path or 'none'} user_profile={user_path or 'none'}"
    print(f"{output}\n{source_note}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
