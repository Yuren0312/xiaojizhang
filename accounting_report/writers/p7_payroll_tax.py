# -*- coding: utf-8 -*-
"""个税扣缴明细表

以员工维度生成个税扣缴明细 xlsx，按自然人电子税务局（ITS）导入模板格式。

生成一个 xlsx 工作簿，包含 2 个 sheet：
  1. 个税扣缴明细 — 员工维度累计预扣预缴计算
  2. 专项附加扣除汇总 — 按员工汇总六项专项附加扣除

计算逻辑（累计预扣法）：
  累计应纳税所得额 = 累计收入 − 累计减除费用(5000×月数)
                    − 累计专项扣除 − 累计专项附加扣除
                    − 累计其他扣除
  累计应扣缴税额 = 累计应纳税所得额 × 预扣率 − 速算扣除数
  本期应补（退）税额 = 累计应扣缴税额 − 累计已预扣缴税额

数据来源：
  - 员工工资表：【依赖未实现数据源】输入/工资表.xlsx
  - 社保公积金明细：【依赖未实现数据源】社保公积金台账
  - 专项附加扣除：【依赖未实现数据源】个税APP导出数据

⚠ 当前系统无员工工资数据源，生成模板级输出。
   请在输入目录放置工资表（含姓名、税前工资、社保个人部分、公积金个人部分等字段）后重新运行。
   输出文件名加 [模板] 前缀。
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ..models import ZERO, money

# ── 样式常量 ────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_FONT = Font(name="微软雅黑", size=10)
BOLD_FONT = Font(name="微软雅黑", bold=True, size=10)
TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
SUBTITLE_FONT = Font(name="微软雅黑", bold=True, size=11)
SECTION_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SECTION_FONT = Font(name="微软雅黑", bold=True, size=10, color="1F3864")
WARN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
WARN_FONT = Font(name="微软雅黑", size=9, color="996600")
TEMPLATE_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
MONEY_FMT = '#,##0.00'
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)


# ── 个税预扣率表（居民个人工资薪金所得） ─────────────────────────────
# 累计预扣预缴应纳税所得额 → 预扣率(%) / 速算扣除数
TAX_BRACKETS = [
    # (累计应纳税所得额上限, 预扣率, 速算扣除数)
    (Decimal("36000"),     Decimal("0.03"), Decimal("0")),
    (Decimal("144000"),    Decimal("0.10"), Decimal("2520")),
    (Decimal("300000"),    Decimal("0.20"), Decimal("16920")),
    (Decimal("420000"),    Decimal("0.25"), Decimal("31920")),
    (Decimal("660000"),    Decimal("0.30"), Decimal("52920")),
    (Decimal("960000"),    Decimal("0.35"), Decimal("85920")),
    (None,                 Decimal("0.45"), Decimal("181920")),  # 无上限
]

# 每月减除费用
MONTHLY_DEDUCTION = Decimal("5000")


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _write_row(ws, row: int, values: list, money_cols: set | None = None,
               bold: bool = False, fill: PatternFill | None = None,
               font: Font | None = None, align: Alignment | None = None,
               pct_cols: set | None = None) -> None:
    """写一行数据"""
    if money_cols is None:
        money_cols = set()
    if pct_cols is None:
        pct_cols = set()
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or (BOLD_FONT if bold else CELL_FONT)
        cell.alignment = align or Alignment(vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        if fill:
            cell.fill = fill
        if i in money_cols and isinstance(v, (Decimal, int, float)):
            cell.number_format = MONEY_FMT
        if i in pct_cols and isinstance(v, (Decimal, float)):
            cell.number_format = '0.0%'


def _auto_width(ws, min_width: int = 10, max_width: int = 50) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            v = str(cell.value) if cell.value is not None else ""
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            lengths.append(length)
        best = max(lengths) if lengths else min_width
        ws.column_dimensions[col_letter].width = max(min(best + 3, max_width), min_width)


def _calc_tax(cumulative_taxable: Decimal) -> Decimal:
    """按累计预扣率表计算个税（五级/七级超额累进）

    每档计税 = min(剩余应纳税所得额, 本档宽度)；剩余逐档递减，减到 0 为止。
    等价于「全额 × 预扣率 − 速算扣除数」。
    注意：不能用「remaining <= upper」作为本档结束判断——remaining 是从
    上一档之后还没计的部分，若它超过本档宽度，超出部分应进入更高档。
    """
    if cumulative_taxable <= ZERO:
        return ZERO
    remaining = cumulative_taxable
    tax = ZERO
    prev_threshold = ZERO
    for upper, rate, _quick_deduction in TAX_BRACKETS:
        if upper is None:
            bracket_amount = remaining
        else:
            bracket_amount = min(remaining, upper - prev_threshold)
        if bracket_amount > ZERO:
            tax += bracket_amount * rate
        remaining -= bracket_amount
        prev_threshold = upper if upper is not None else prev_threshold
        if remaining <= ZERO:
            break
    return tax


# ── Sheet 1: 个税扣缴明细 ───────────────────────────────────────────

def _payroll_tax_detail(ws, employees: list, year: int, month: int,
                        is_template: bool) -> None:
    """员工维度累计预扣预缴个税明细"""

    ws.merge_cells("A1:N1")
    title = (f"{year}年{month}月 个人所得税扣缴明细表"
             f"{' [模板]' if is_template else ''}")
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    if is_template:
        ws.merge_cells("A2:N2")
        ws.cell(row=2, column=1,
                value="⚠ 此为模板级输出，员工工资数据源缺失（输入/工资表.xlsx）。请补充工资表后重新生成。"
                ).font = WARN_FONT
        ws.cell(row=2, column=1).fill = WARN_FILL

    # ── 基础信息 ──
    info_row = 3 if not is_template else 4
    info_items = [
        ("扣缴义务人", "示例科技有限公司"),
        ("所属期", f"{year}年{month}月"),
        ("申报人数", len(employees) if employees else "【数据缺失】"),
    ]
    for i, (label, val) in enumerate(info_items):
        ws.cell(row=info_row + i, column=1, value=label).font = BOLD_FONT
        ws.cell(row=info_row + i, column=2, value=val).font = CELL_FONT

    header_row = info_row + 5

    # ── 表头（模拟 ITS 导入模板） ──
    headers = [
        "序号", "姓名", "身份证号",
        "本期收入", "本期免税收入",
        "基本养老保险", "基本医疗保险", "失业保险", "住房公积金",
        "累计收入", "累计减除费用\n(5000×月数)",
        "累计专项扣除", "累计专项附加扣除",
        "累计应纳税所得额", "累计应扣缴税额",
        "累计已预扣税额", "本期应补（退）税额",
    ]

    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    _style_header(ws, header_row, len(headers))

    money_cols = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17}
    row = header_row + 1

    if employees:
        for idx, emp in enumerate(employees, start=1):
            name = emp.get("name", "")
            id_card = emp.get("id_card", "")

            # 本期收入（税前工资）
            monthly_income = money(emp.get("monthly_income", ZERO))
            tax_free_income = money(emp.get("tax_free_income", ZERO))

            # 本期专项扣除（社保公积金个人部分）
            pension = money(emp.get("pension", ZERO))          # 基本养老保险
            medical = money(emp.get("medical", ZERO))           # 基本医疗保险
            unemployment = money(emp.get("unemployment", ZERO)) # 失业保险
            housing_fund = money(emp.get("housing_fund", ZERO)) # 住房公积金
            monthly_special_deduct = pension + medical + unemployment + housing_fund

            # 累计数据
            cumulative_income = money(emp.get("cumulative_income",
                                              monthly_income))
            cumulative_deduction = MONTHLY_DEDUCTION * Decimal(month)  # 累计减除费用
            cumulative_special_deduct = money(emp.get(
                "cumulative_special_deduct",
                monthly_special_deduct))
            cumulative_special_additional = money(emp.get(
                "cumulative_special_additional", ZERO))  # 累计专项附加扣除
            cumulative_other_deduct = money(emp.get(
                "cumulative_other_deduct", ZERO))  # 累计其他扣除（企业年金等）

            # 累计应纳税所得额
            cumulative_taxable = (cumulative_income
                                  - cumulative_deduction
                                  - cumulative_special_deduct
                                  - cumulative_special_additional
                                  - cumulative_other_deduct)
            if cumulative_taxable < ZERO:
                cumulative_taxable = ZERO

            # 累计应扣缴税额
            cumulative_tax = _calc_tax(cumulative_taxable)

            # 累计已预扣缴税额
            cumulative_prepaid = money(emp.get("cumulative_prepaid", ZERO))

            # 本期应补（退）税额
            current_tax_due = cumulative_tax - cumulative_prepaid

            _write_row(ws, row, [
                idx, name, id_card,
                float(monthly_income),
                float(tax_free_income),
                float(pension), float(medical),
                float(unemployment), float(housing_fund),
                float(cumulative_income),
                float(cumulative_deduction),
                float(cumulative_special_deduct),
                float(cumulative_special_additional),
                float(cumulative_taxable),
                float(cumulative_tax),
                float(cumulative_prepaid),
                float(current_tax_due),
            ], money_cols)
            row += 1
    else:
        # 模板模式：占位行
        # 【依赖未实现数据源】输入/工资表.xlsx
        _write_row(ws, row, [
            1,
            "【模板】员工姓名",
            "【模板】身份证号",
            0.0, 0.0,
            0.0, 0.0, 0.0, 0.0,
            0.0,
            0.0,
            0.0,
            0.0,
            0.0,
            0.0,
            0.0,
            0.0,
        ], money_cols, fill=TEMPLATE_FILL)

    # ── 补充说明 ──
    row += 2
    notes = [
        "计算逻辑说明：",
        "• 累计应纳税所得额 = 累计收入 − 累计减除费用(5000×月数) − 累计专项扣除 − 累计专项附加扣除",
        "• 累计应扣缴税额 = 累计应纳税所得额 × 预扣率 − 速算扣除数（居民个人工资薪金累计预扣法）",
        "• 本期应补（退）税额 = 累计应扣缴税额 − 累计已预扣缴税额",
        "",
        "数据源依赖：【以下数据源均未实现】",
        "• 基础工资数据：输入/工资表.xlsx（姓名、税前工资、社保个人部分、公积金个人部分）",
        "• 社保公积金：输入/社保公积金台账.xlsx",
        "• 专项附加扣除：自然人电子税务局（ITS）导出数据 → 输入/专项附加扣除.xlsx",
        "• 历史预扣缴记录：【依赖未实现数据源】往期个税申报记录",
        "",
        "预扣率表（居民个人工资薪金所得）：",
        "    累计应纳税所得额 ≤ 36,000   → 3%   速算扣除 0",
        "    36,000 ~ 144,000            → 10%  速算扣除 2,520",
        "    144,000 ~ 300,000           → 20%  速算扣除 16,920",
        "    300,000 ~ 420,000           → 25%  速算扣除 31,920",
        "    420,000 ~ 660,000           → 30%  速算扣除 52,920",
        "    660,000 ~ 960,000           → 35%  速算扣除 85,920",
        "    > 960,000                   → 45%  速算扣除 181,920",
    ]
    for note in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        ws.cell(row=row, column=1, value=note).font = Font(name="微软雅黑", size=9, color="666666")
        row += 1

    _auto_width(ws, min_width=12, max_width=28)


# ── Sheet 2: 专项附加扣除汇总 ───────────────────────────────────────

def _special_deduction_summary(ws, employees: list, is_template: bool) -> None:
    """汇总每个员工的六项专项附加扣除"""

    ws.merge_cells("A1:J1")
    ws.cell(row=1, column=1,
            value=f"专项附加扣除汇总{' [模板]' if is_template else ''}").font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    if is_template:
        ws.merge_cells("A2:J2")
        ws.cell(row=2, column=1,
                value="⚠ 此为模板级输出，专项附加扣除数据源缺失（个税APP导出）。请补充数据后重新生成。"
                ).font = WARN_FONT
        ws.cell(row=2, column=1).fill = WARN_FILL

    header_row = 4 if not is_template else 5
    headers = [
        "序号", "姓名",
        "子女教育", "继续教育", "住房贷款利息",
        "住房租金", "赡养老人", "婴幼儿照护",
        "合计", "备注",
    ]

    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    _style_header(ws, header_row, len(headers))

    money_cols = {3, 4, 5, 6, 7, 8, 9}
    row = header_row + 1

    if employees:
        for idx, emp in enumerate(employees, start=1):
            name = emp.get("name", "")
            child_edu = money(emp.get("child_edu", ZERO))
            continue_edu = money(emp.get("continue_edu", ZERO))
            housing_loan = money(emp.get("housing_loan", ZERO))
            housing_rent = money(emp.get("housing_rent", ZERO))
            elderly_care = money(emp.get("elderly_care", ZERO))
            infant_care = money(emp.get("infant_care", ZERO))

            # 住房贷款利息和住房租金互斥，不能同时享受
            total = (child_edu + continue_edu
                    + max(housing_loan, housing_rent)
                    + elderly_care + infant_care)

            note_parts = []
            if housing_loan > ZERO and housing_rent > ZERO:
                note_parts.append("住房贷款利息与租金互斥，已取较大值")
            note = "；".join(note_parts) if note_parts else ""

            _write_row(ws, row, [
                idx, name,
                float(child_edu), float(continue_edu),
                float(housing_loan), float(housing_rent),
                float(elderly_care), float(infant_care),
                float(total), note,
            ], money_cols)
            row += 1
    else:
        # 模板占位行
        placeholders = ["子女教育\n(1000/2000/孩)", "继续教育\n(400/月)", "住房贷款利息\n(1000/月)",
                       "住房租金\n(800-1500/月)", "赡养老人\n(2000-3000/月)", "婴幼儿照护\n(1000-2000/月)"]
        _write_row(ws, row, [
            1, "【模板】员工姓名",
            *placeholders,
            0.0, "【依赖未实现数据源】个税APP导出"
        ], money_cols, fill=TEMPLATE_FILL)

    # ── 补充说明 ──
    row += 2
    notes = [
        "专项附加扣除标准（每月）：",
        "• 子女教育：每个子女 2,000 元（3岁至博士）",
        "• 继续教育：学历(学位) 400 元 / 职业资格 3,600 元/年",
        "• 住房贷款利息：首套房 1,000 元/月（夫妻一方扣除）",
        "• 住房租金：800/1,100/1,500 元/月（视城市等级）—— 与住房贷款利息互斥",
        "• 赡养老人：2,000-3,000 元/月（视是否为独生子女）",
        "• 3岁以下婴幼儿照护：每个婴幼儿 2,000 元/月",
        "",
        "数据源：自然人电子税务局（ITS）→ 专项附加扣除信息采集 → 导出",
        "【依赖未实现数据源】输入/专项附加扣除.xlsx",
    ]
    for note in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        ws.cell(row=row, column=1, value=note).font = Font(name="微软雅黑", size=9, color="666666")
        row += 1

    _auto_width(ws, min_width=12, max_width=30)


# ── 主函数 ──────────────────────────────────────────────────────────────

def generate(
    year: int,
    month: int,
    employees: list | None = None,  # 【依赖未实现数据源】员工工资数据
    out_path: Path | str | None = None,
) -> tuple[Path, bool]:
    """生成个税扣缴明细表 xlsx。

    Args:
        year: 年份
        month: 月份
        employees: 员工列表。每项为 dict:
            {
                "name": str,                   # 姓名
                "id_card": str,                # 身份证号
                "monthly_income": Decimal,      # 本期税前工资
                "tax_free_income": Decimal,    # 本期免税收入
                "pension": Decimal,             # 基本养老保险（个人部分）
                "medical": Decimal,             # 基本医疗保险（个人部分）
                "unemployment": Decimal,        # 失业保险（个人部分）
                "housing_fund": Decimal,        # 住房公积金（个人部分）
                "cumulative_income": Decimal,   # 累计收入（含本期）
                "cumulative_special_deduct": Decimal,  # 累计专项扣除
                "cumulative_special_additional": Decimal,  # 累计专项附加扣除
                "cumulative_other_deduct": Decimal,       # 累计其他扣除（如企业年金）
                "cumulative_prepaid": Decimal,  # 累计已预扣缴税额
                # Sheet 2 专项附加扣除明细：
                "child_edu": Decimal,           # 子女教育
                "continue_edu": Decimal,        # 继续教育
                "housing_loan": Decimal,        # 住房贷款利息
                "housing_rent": Decimal,        # 住房租金
                "elderly_care": Decimal,        # 赡养老人
                "infant_care": Decimal,         # 婴幼儿照护
            }
            如果为 None，生成模板输出。
        out_path: 输出路径。为 None 时自动生成。

    Returns:
        (out_path, is_template): 输出路径 + 是否为模板
    """
    if out_path is None:
        out_path = Path(".")
    out_path = Path(out_path)
    is_template = employees is None or len(employees) == 0

    if out_path.is_dir():
        prefix = f"{year:04d}-{month:02d}"
        fname = f"{prefix} {'[模板] ' if is_template else ''}个税扣缴明细表.xlsx"
        out_path = out_path / fname

    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Sheet 1: 个税扣缴明细
    ws1 = wb.active
    ws1.title = "个税扣缴明细"
    _payroll_tax_detail(ws1, employees or [], year, month, is_template)

    # Sheet 2: 专项附加扣除汇总
    ws2 = wb.create_sheet("专项附加扣除汇总")
    _special_deduction_summary(ws2, employees or [], is_template)

    wb.save(out_path)
    return out_path, is_template
