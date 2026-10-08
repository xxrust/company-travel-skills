#!/usr/bin/env python
"""Build the reusable Excel workbook for the company paper reimbursement form."""

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


def build_reimbursement_page(ws, user: Mapping[str, str], company: Mapping[str, str], page_number: int = 1) -> None:
    """Create one fixed-size page matching the supplied paper form."""
    title_font = Font(name="SimSun", size=20, underline="double", color=BLUE)
    label_font = Font(name="Microsoft YaHei", size=10, color=BLUE)
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color=BLUE)
    body_font = Font(name="Microsoft YaHei", size=10, color="000000")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    line = Side(style="thin", color="7897BA")

    ws.title = "报销单" if page_number == 1 else f"报销单-第{page_number}页"
    ws.sheet_view.showGridLines = False
    widths = [9.0, 22.0, 7.0, 7.0, 8.0, 8.0, 8.0, 6.5, 7.5, 9.0]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width
    heights = {1: 20, 2: 16, 3: 18, 4: 30, 5: 24, 6: 20, 7: 18, 14: 21, 15: 23, 16: 24}
    for row, height in heights.items():
        ws.row_dimensions[row].height = height
    for row in range(8, 14):
        ws.row_dimensions[row].height = 23

    for merge in (
        "A1:G2", "H1:I1", "H2:I2", "A3:B3", "C3:G3", "H3:J3",
        "B4:C4", "G4:J4", "A5:B5", "C5:G5", "H5:J5",
        "A6:A7", "B6:B7", "C6:D6", "E6:E7", "F6:F7", "G6:G7",
        "H6:I6", "J6:J7", "A14:C14", "A15:F15", "G15:I15", "A16:C16", "D16:G16", "H16:J16",
    ):
        ws.merge_cells(merge)

    labels = {
        "A1": "差 旅 费 报 销 单", "H1": "预领款", "H2": "补领或缴还", "A3": "部门：",
        "A4": "姓　名", "D4": "职　别", "F4": "出差\n事由", "A5": "出差起止日期",
        "C5": "自　　　　　　至", "H5": "附单据",
        "A6": "起讫\n日期", "B6": "起讫地点", "C6": "车船费", "C7": "名称", "D7": "金额",
        "E6": "宿费", "F6": "出差\n补助", "G6": "市内\n交通费", "H6": "杂费",
        "H7": "用途", "I7": "金额", "J6": "附注", "A14": "合　　计",
        "A15": "合计金额（大写）", "A16": "单位主管：", "D16": "复　核：", "H16": "出差人：",
    }
    for addr, value in labels.items():
        cell = ws[addr]
        cell.value = value
        cell.font = title_font if addr == "A1" else label_font
        cell.alignment = center if addr in ("A1", "A6", "B6", "C6", "E6", "F6", "G6", "H6", "J6", "A14") else left

    ws["A3"] = "部门：" + user.get("department", "")
    ws["B4"] = user.get("name", "")
    ws["E4"] = user.get("title", "")
    ws["C5"] = "自　　　　　　至　　　　　共　　天"
    ws["H5"] = "附单据　　　　　张"
    ws["J1"] = ""
    ws["J2"] = ""
    ws["C3"] = "年　　月　　日"
    ws["C3"].font = label_font
    ws["C3"].alignment = center
    ws.row_dimensions[14].hidden = True
    ws.row_dimensions[14].height = 0
    ws["G15"] = "￥："

    # The photographed form has a narrow, single date field; give it enough
    # room for a complete month/day while keeping the location field dominant.
    for row in range(6, 8):
        for col in range(1, 11):
            cell = ws.cell(row, col)
            cell.font = header_font
            cell.alignment = center
            cell.fill = PatternFill(fill_type=None)
    for row in range(8, 14):
        for col in range(1, 11):
            cell = ws.cell(row, col)
            cell.font = body_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False) if col == 1 else (center if col in (3, 4, 5, 6, 7, 9) else left)
            if col in (4, 5, 7, 9):
                cell.number_format = "#,##0.00"
            if col == 3:
                cell.font = Font(name="Microsoft YaHei", size=9)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)
    for row in range(6, 15):
        for col in range(1, 11):
            cell = ws.cell(row, col)
            cell.border = Border(left=line, right=line, top=line, bottom=line)
    for col, column in ((4, "D"), (5, "E"), (7, "G"), (9, "I")):
        ws.cell(14, col).value = f'=IF(SUM({column}8:{column}13)=0,"",SUM({column}8:{column}13))'
        ws.cell(14, col).number_format = "#,##0.00"
        ws.cell(14, col).alignment = center
    for row in (3, 4, 5):
        for col in range(1, 11):
            ws.cell(row, col).border = Border(bottom=line)
    ws["A3"].border = Border(bottom=Side(style="dashed",color="7897BA"))
    ws["B3"].border = Border(bottom=Side(style="dashed",color="7897BA"))
    ws["C3"].border = Border()
    ws["H3"].border = Border()
    for row, starts, ends in ((4,(1,2,4,5,6,7),(1,3,4,5,6,10)),(5,(1,3,8),(2,7,10))):
        for col in range(1,11):
            ws.cell(row,col).border = Border(top=line,bottom=line,left=line if col in starts else Side(),right=line if col in ends else Side())
    ws["J3"].border = Border(right=line, top=line, bottom=line)
    ws["A4"].border = Border(left=line, bottom=line)
    ws["J4"].border = Border(right=line, bottom=line)
    ws["A5"].border = Border(left=line, bottom=line)
    ws["J5"].border = Border(right=line, bottom=line)
    for row in (1, 2):
        for col in range(8, 11):
            ws.cell(row, col).border = Border(left=line, right=line, top=line, bottom=line)
    for addr in ("A3", "B4", "E4", "G4", "C5", "H5", "J1", "J2"):
        ws[addr].font = body_font
        ws[addr].alignment = left
    ws["E4"].alignment = Alignment(horizontal="left",vertical="center",shrink_to_fit=True,wrap_text=False)
    ws["C5"].font = Font(name="Microsoft YaHei",size=9)
    ws["C5"].alignment = Alignment(horizontal="center",vertical="center",shrink_to_fit=True,wrap_text=False)
    for col in range(1, 11):
        ws.cell(15, col).border = Border(top=line, bottom=line, left=line if col == 1 else Side(), right=line if col == 10 else Side())
    ws["I15"].border = Border(top=line,bottom=line,right=line)
    ws["J15"].border = Border(top=line,bottom=line,left=line,right=line)
    ws["G15"].alignment = Alignment(horizontal="left",vertical="center",indent=2)
    ws["G15"].font = body_font
    for addr in ("A16", "D16", "H16"):
        ws[addr].font = body_font
        ws[addr].alignment = left

    ws.freeze_panes = "A8"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.25, bottom=0.25, header=0.1, footer=0.1)
    ws.print_area = "A1:J16"
    ws.print_options.horizontalCentered = True


def add_reimbursement_page(wb: Workbook, user: Mapping[str, str], company: Mapping[str, str], page_number: int):
    ws = wb.create_sheet()
    build_reimbursement_page(ws, user, company, page_number)
    return ws


def make_workbook(company: Mapping[str, str] | None = None, user: Mapping[str, str] | None = None) -> Workbook:
    company = dict(company or {})
    user = dict(user or {})
    wb = Workbook()
    ws = wb.active
    build_reimbursement_page(ws, user, company)
    detail = wb.create_sheet("发票明细")
    check = wb.create_sheet("校验与说明")
    title_font = Font(name="Microsoft YaHei", size=18, bold=True, color=BLUE)
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color=WHITE)
    body_font = Font(name="Microsoft YaHei", size=10, color="000000")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    blue_fill = PatternFill("solid", fgColor=BLUE)
    yellow_fill = PatternFill("solid", fgColor=PALE_YELLOW)
    red_fill = PatternFill("solid", fgColor=PALE_RED)

    headers = ["来源文件", "页码", "单据类型", "发票号码/订单号", "开票/乘车/入住日期", "结束日期", "购买方名称", "购买方税号", "购买方地址", "购买方电话", "购买方开户行", "购买方账号", "销方/承运方/酒店", "税率", "不含税金额", "税额", "价税合计", "出发地", "到达地", "入住时间", "离开时间", "分类", "可报销", "校验状态", "备注"]
    for col, value in enumerate(headers, 1):
        cell = detail.cell(1, col, value)
        cell.fill = blue_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    for row in range(2, 102):
        for col in range(1, 26):
            cell = detail.cell(row, col)
            cell.font = body_font
            cell.border = border
            cell.alignment = left if col in (1, 4, 7, 9, 13, 18, 19, 24, 25) else center
            if col in (15, 16, 17):
                cell.number_format = "#,##0.00"
    for col, width in {1: 34, 2: 8, 3: 16, 4: 24, 5: 18, 6: 16, 7: 24, 8: 22, 9: 42, 10: 14, 11: 18, 12: 24, 13: 30, 14: 10, 15: 14, 16: 14, 17: 14, 18: 18, 19: 18, 20: 16, 21: 16, 22: 14, 23: 12, 24: 20, 25: 40}.items():
        detail.column_dimensions[chr(64 + col)].width = width
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = "A1:Y101"
    category = DataValidation(type="list", formula1='"出租车/网约车,飞机/行程单,住宿,火车/动车,其他"', allow_blank=True)
    yes = DataValidation(type="list", formula1='"是,否,待确认"', allow_blank=True)
    status = DataValidation(type="list", formula1='"已核对,购买方信息不一致,待补充入住/离店时间,路线未闭环,金额待确认,来源无法读取"', allow_blank=True)
    for validation, ref in ((category, "V2:V101"), (yes, "W2:W101"), (status, "X2:X101")):
        detail.add_data_validation(validation)
        validation.add(ref)
    detail.conditional_formatting.add("X2:X101", FormulaRule(formula=['AND($X2<>"",$X2<>"已核对")'], fill=red_fill))

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
    company_fields = [("公司名称", company.get("name", "")), ("税号", company.get("tax_id", "")), ("地址", company.get("address", "")), ("电话", company.get("phone", "")), ("开户行", company.get("bank_name", "")), ("账号", company.get("bank_account", ""))]
    for row, (field, value) in enumerate(company_fields, 4):
        for col, item in enumerate((field, value, "", "待核对（本地配置）" if value else "待填写本地配置"), 1):
            cell = check.cell(row, col, item)
            cell.border = border
            cell.alignment = left
            cell.font = body_font
            if col in (2, 3):
                cell.fill = yellow_fill
    check.merge_cells("A12:D12")
    check["A12"] = "必须检查"
    check["A12"].fill = blue_fill
    check["A12"].font = header_font
    check["A12"].alignment = center
    rules = [
        "每个 PDF/图片先用本地 MinerU 解析，并在发票明细保留来源文件和页码。",
        "住宿票必须同时有入住和离店日期；缺任一日期就标记待补充，不得当作完整凭证。",
        "A 到 B 出差当天的住宿费写在 A 到 B 行；B 地多张住宿票按日期从上到下分行。",
        "路线必须闭环；缺少回程证据时标记路线未闭环，不虚构行程。",
        "报销单不打印签名，单位主管、复核、出差人栏留空，打印后手写。",
        "模板字段、填充坐标、纸面表头必须一致；第一条行程必须核对交通方式和去程方向。",
    ]
    for row, rule in enumerate(rules, 13):
        check.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        check.cell(row, 1, f"{row - 12}. {rule}")
        check.cell(row, 1).alignment = left
        check.cell(row, 1).font = body_font
        for col in range(1, 5):
            check.cell(row, col).border = border
    for col, width in {"A": 20, "B": 48, "C": 48, "D": 30}.items():
        check.column_dimensions[col].width = width
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
    make_workbook(company, user).save(output)
    print(f"{output}\ncompany_profile={company_path or 'none'} user_profile={user_path or 'none'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
