# -*- coding: utf-8 -*-
"""所得税预缴底稿

生成一个 xlsx 工作簿，包含 2 个 sheet：
  1. 企业所得税月（季）度预缴申报表（A类）— 适用于有限公司
  2. 经营所得个人所得税预缴表 — 适用于个体户/个人独资企业

自动识别企业类型并选择对应底稿格式。

数据来源：
  - 利润表：营业收入、营业成本、利润总额
  - 小微企业优惠政策：年应纳税所得额 ≤ 300 万元
  - 资产总额 / 从业人数：【依赖未实现数据源】资产台账 + 员工名册

⚠ 当前为示例科技有限公司（有限公司），默认生成 A类 企业所得税底稿。
   个体户模式需在调用时传入 entity_type="individual"。

⚠ 资产总额/从业人数数据源缺失时，生成模板级输出并在文件名加 [模板] 前缀。
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ..models import ReportSet, ZERO, money

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


# ── 企业所得税政策常量 ───────────────────────────────────────────────
# 基本税率 25%
CORP_TAX_RATE = Decimal("0.25")

# 小微企业优惠政策（截至2027年底）
# 年应纳税所得额 ≤ 300 万：减按 25% 计入应纳税所得额，按 20% 税率
# 实际税率 = 25% × 20% = 5%
SMALL_PROFIT_RATE = Decimal("0.05")   # 实际税率 5%（= 25%×20%）
SMALL_PROFIT_THRESHOLD = Decimal("3000000")  # 年应纳税所得额 300 万

# 小微企业三要素门槛
SMALL_PROFIT_ASSET_MAX = Decimal("50000000")   # 资产总额 ≤ 5000 万
SMALL_PROFIT_EMPLOYEE_MAX = 300                # 从业人数 ≤ 300 人
SMALL_PROFIT_REVENUE_MAX = Decimal("50000000") # 年应纳税收入 ≤ 5000 万（非必要条件，但常见）

# 个体户经营所得个人所得税税率表（五级超额累进）
INDIVIDUAL_TAX_BRACKETS = [
    # (年应纳税所得额上限, 税率, 速算扣除数)
    (Decimal("30000"),     Decimal("0.05"), Decimal("0")),
    (Decimal("90000"),     Decimal("0.10"), Decimal("1500")),
    (Decimal("300000"),    Decimal("0.20"), Decimal("10500")),
    (Decimal("500000"),    Decimal("0.30"), Decimal("40500")),
    (None,                 Decimal("0.35"), Decimal("65500")),  # 无上限
]


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _write_row(ws, row: int, values: list, money_cols: set | None = None,
               bold: bool = False, fill: PatternFill | None = None,
               font: Font | None = None, align: Alignment | None = None) -> None:
    """写一行数据"""
    if money_cols is None:
        money_cols = set()
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or (BOLD_FONT if bold else CELL_FONT)
        cell.alignment = align or Alignment(vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        if fill:
            cell.fill = fill
        if i in money_cols and isinstance(v, (Decimal, int, float)):
            cell.number_format = MONEY_FMT


def _auto_width(ws, min_width: int = 10, max_width: int = 55) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            v = str(cell.value) if cell.value is not None else ""
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            lengths.append(length)
        best = max(lengths) if lengths else min_width
        ws.column_dimensions[col_letter].width = max(min(best + 3, max_width), min_width)


def _calc_individual_tax(annual_taxable_income: Decimal) -> Decimal:
    """个体户经营所得个税计算（五级超额累进）

    每档计税 = min(剩余所得, 本档宽度)，逐档递减（与 p7._calc_tax 同修复：
    不能用 remaining<=upper 判断，否则跨档时少算税）。
    """
    if annual_taxable_income <= ZERO:
        return ZERO
    remaining = annual_taxable_income
    total_tax = ZERO
    prev_bracket = ZERO
    for upper, rate, _quick_deduction in INDIVIDUAL_TAX_BRACKETS:
        if upper is None:
            bracket_amount = remaining
        else:
            bracket_amount = min(remaining, upper - prev_bracket)
        if bracket_amount > ZERO:
            total_tax += bracket_amount * rate
        remaining -= bracket_amount
        prev_bracket = upper if upper is not None else prev_bracket
        if remaining <= ZERO:
            break
    return total_tax


# ── Sheet 1: 企业所得税月（季）度预缴申报表（A类） ─────────────────

def _corp_income_tax_a(ws, report_set, prior_report_set,
                       year: int, month: int, quarter: int,
                       asset_total: Decimal | None,
                       employee_count: int | None,
                       is_template: bool) -> None:
    """企业所得税 A 类预缴申报底稿"""

    # 数据提取
    if report_set is not None and hasattr(report_set, 'is_ytd'):
        revenue_ytd = report_set.is_ytd.get(1, ZERO)     # 营业收入 累计
        cost_ytd = report_set.is_ytd.get(2, ZERO)         # 营业成本 累计
        profit_total_ytd = report_set.is_ytd.get(30, ZERO)  # 利润总额 累计
    else:
        revenue_ytd = ZERO
        cost_ytd = ZERO
        profit_total_ytd = ZERO

    # 小微企业判定
    # 【依赖未实现数据源】资产总额、从业人数 来自资产台账 + 员工名册
    if asset_total is None:
        asset_total = ZERO  # 缺失时标记为模板
    if employee_count is None:
        employee_count = 0

    is_small_profit = (
        profit_total_ytd <= SMALL_PROFIT_THRESHOLD
        and (asset_total <= SMALL_PROFIT_ASSET_MAX or asset_total == ZERO)
        and (employee_count <= SMALL_PROFIT_EMPLOYEE_MAX or employee_count == 0)
    )

    # 应纳税额计算
    if is_small_profit and profit_total_ytd > ZERO:
        tax_payable = profit_total_ytd * SMALL_PROFIT_RATE
        effective_rate = SMALL_PROFIT_RATE
    else:
        tax_payable = profit_total_ytd * CORP_TAX_RATE
        effective_rate = CORP_TAX_RATE

    # 已预缴税额：【依赖未实现数据源】前期预缴台账
    tax_prepaid = ZERO

    # 本期应补（退）
    tax_due = tax_payable - tax_prepaid

    # ── 标题 ──
    ws.merge_cells("A1:F1")
    title = (f"{year}年{quarter}季度 企业所得税月（季）度预缴纳税申报表（A类）"
             f"{' [模板]' if is_template else ''}")
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    if is_template:
        ws.merge_cells("A2:F2")
        ws.cell(row=2, column=1,
                value="⚠ 此为模板级输出，资产总额/从业人数数据源缺失。请补充资产台账和员工名册后重新生成。"
                ).font = WARN_FONT
        ws.cell(row=2, column=1).fill = WARN_FILL

    # ── 基础信息 ──
    info_row = 3 if not is_template else 4
    info_items = [
        ("纳税人识别号", "【依赖未实现数据源】纳税人基础信息"),
        ("纳税人名称", "示例科技有限公司"),
        ("预缴方式", "按照实际利润额预缴"),
        ("企业类型", "有限公司"),
        ("所属时期", f"{year}-{month:02d}"),
    ]
    for i, (label, val) in enumerate(info_items):
        ws.cell(row=info_row + i, column=1, value=label).font = BOLD_FONT
        ws.cell(row=info_row + i, column=2, value=val).font = CELL_FONT

    header_row = info_row + 6

    # ── 预缴税款计算 ──
    headers = ["行次", "项　　目", "本年累计金额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    _style_header(ws, header_row, len(headers))

    money_cols = {3}
    row = header_row + 1

    # 主表行
    lines = [
        (1,  "一、营业收入", revenue_ytd),
        (2,  "　　减：营业成本", cost_ytd),
        (3,  "　　减：税金及附加", report_set.is_ytd.get(3, ZERO) if report_set else ZERO),
        (4,  "　利润总额（1−2−3）", profit_total_ytd),
        (5,  "　加：纳税调增额", ZERO),  # 【依赖未实现数据源】纳税调整台账
        (6,  "　减：纳税调减额", ZERO),  # 【依赖未实现数据源】纳税调整台账
        (7,  "　减：弥补以前年度亏损", ZERO),  # 【依赖未实现数据源】亏损弥补台账
        (8,  "　实际利润额（4+5−6−7）", profit_total_ytd),
        (9,  "　税率（25%）", f"{float(CORP_TAX_RATE)*100:.0f}%"),
        (10, "　应纳所得税额（8×9）", profit_total_ytd * CORP_TAX_RATE),
    ]

    for line_no, name, value in lines:
        _write_row(ws, row, [line_no, name, float(value) if isinstance(value, Decimal) else value], money_cols)
        row += 1

    # ── 减免所得税额 ──
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="减免所得税优惠明细").font = SECTION_FONT
    ws.cell(row=row, column=1).fill = SECTION_FILL
    row += 1

    # 小微企业减免
    if is_small_profit and profit_total_ytd > ZERO:
        tax_full = profit_total_ytd * CORP_TAX_RATE
        tax_reduced = profit_total_ytd * SMALL_PROFIT_RATE
        small_profit_exemption = tax_full - tax_reduced
    else:
        small_profit_exemption = ZERO

    exemption_lines = [
        (11, "符合条件的小型微利企业减免所得税额",
         small_profit_exemption),
        (12, "其他减免所得税额", ZERO),  # 【依赖未实现数据源】税收优惠台账
    ]

    for line_no, name, value in exemption_lines:
        _write_row(ws, row, [line_no, name, float(value)], money_cols)
        row += 1

    # ── 最终税额 ──
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="本期应补（退）所得税额").font = SECTION_FONT
    ws.cell(row=row, column=1).fill = SECTION_FILL
    row += 1

    final_lines = [
        (13, "减：减免所得税额（11+12）", small_profit_exemption),
        (14, "减：已预缴所得税额", tax_prepaid),  # 【依赖未实现数据源】预缴台账
        (15, "减：以前年度多缴在本期抵缴", ZERO),  # 【依赖未实现数据源】汇算清缴数据
        (16, "本期应补（退）所得税额（10−13−14−15）", tax_due),
    ]

    for line_no, name, value in final_lines:
        _write_row(ws, row, [line_no, name, float(value)], money_cols,
                   bold=(line_no == 16))
        row += 1

    # ── 小微企业条件判定 ──
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="小微企业判定条件").font = SUBTITLE_FONT
    row += 1

    cond_headers = ["条件", "标准", "实际值", "判定"]
    for i, h in enumerate(cond_headers, start=1):
        ws.cell(row=row, column=i, value=h)
    _style_header(ws, row, len(cond_headers))
    row += 1

    conditions = [
        ("年应纳税所得额 ≤ 300 万",
         "≤ 3,000,000",
         float(profit_total_ytd),
         "✓" if profit_total_ytd <= SMALL_PROFIT_THRESHOLD else "✗"),
        ("资产总额 ≤ 5000 万",
         "≤ 50,000,000",
         float(asset_total) if asset_total > ZERO else "【数据缺失】",
         ("✓" if asset_total <= SMALL_PROFIT_ASSET_MAX else "✗") if asset_total > ZERO else "? 待补录"),
        ("从业人数 ≤ 300 人",
         "≤ 300",
         employee_count if employee_count > 0 else "【数据缺失】",
         ("✓" if employee_count <= SMALL_PROFIT_EMPLOYEE_MAX else "✗") if employee_count > 0 else "? 待补录"),
    ]

    for cond, standard, actual, result in conditions:
        _write_row(ws, row, [cond, standard, actual, result], money_cols,
                   font=WARN_FONT if "数据缺失" in str(actual) else None,
                   fill=WARN_FILL if "数据缺失" in str(actual) else None)
        row += 1

    # 小微企业判定结果
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    if is_small_profit:
        status_text = (f"✅ 符合小微企业条件 → 实际税负 {float(SMALL_PROFIT_RATE)*100:.0f}% "
                       f"→ 应纳所得税额 {float(tax_payable):,.2f} 元")
    else:
        status_text = "❌ 不符合小微企业条件（或数据不完整无法判定）→ 适用 25% 法定税率"
    ws.cell(row=row, column=1, value=status_text).font = SECTION_FONT

    # ── 数据源说明 ──
    row += 2
    notes = [
        "数据源说明：",
        "• 营业收入、营业成本、利润总额：来自利润表本年累计数",
        "• 纳税调整项目：【依赖未实现数据源】纳税调整台账",
        "• 资产总额：资产台账（净值合计，输入/企业信息/资产台账.xlsx）",
        "• 从业人数：员工名册（输入/企业信息/员工名册.xlsx）",
        "• 已预缴税额：【依赖未实现数据源】前期预缴台账",
    ]
    for note in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1, value=note).font = Font(name="微软雅黑", size=9, color="666666")
        row += 1

    _auto_width(ws)


# ── Sheet 2: 经营所得个人所得税预缴表（个体户） ─────────────────────

def _individual_income_tax(ws, report_set, year: int, month: int, quarter: int,
                           is_template: bool) -> None:
    """个体户经营所得个税预缴底稿"""

    if report_set is not None and hasattr(report_set, 'is_ytd'):
        revenue_ytd = report_set.is_ytd.get(1, ZERO)
        cost_ytd = report_set.is_ytd.get(2, ZERO)
        expense_total = sum(report_set.is_ytd.get(line, ZERO) for line in [3, 11, 14, 18])
        profit_ytd = report_set.is_ytd.get(30, ZERO)
    else:
        revenue_ytd = ZERO
        cost_ytd = ZERO
        expense_total = ZERO
        profit_ytd = ZERO

    # 业主减除费用：【依赖未实现数据源】业主个人基础信息
    owner_deduction = Decimal("5000") * Decimal(month)  # 每月5000
    # 专项附加扣除：【依赖未实现数据源】业主专项附加扣除信息
    special_deduction = ZERO

    taxable_income = profit_ytd - owner_deduction - special_deduction
    if taxable_income < ZERO:
        taxable_income = ZERO

    annual_estimated_taxable = taxable_income  # 年化估计
    tax_payable = _calc_individual_tax(annual_estimated_taxable)

    # 已预缴：【依赖未实现数据源】前期预缴台账
    tax_prepaid = ZERO
    tax_due = tax_payable - tax_prepaid

    # ── 标题 ──
    ws.merge_cells("A1:E1")
    ws.cell(row=1, column=1,
            value=f"{year}年{quarter}季度 经营所得个人所得税预缴申报表{' [模板]' if is_template else ''}"
            ).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    info_row = 3
    info_items = [
        ("纳税人识别号", "【依赖未实现数据源】业主身份证号"),
        ("业主姓名", "【依赖未实现数据源】业主基础信息"),
        ("被投资单位名称", "示例科技有限公司"),
        ("征收方式", "查账征收"),
        ("所属时期", f"{year}-{month:02d}"),
    ]
    for i, (label, val) in enumerate(info_items):
        ws.cell(row=info_row + i, column=1, value=label).font = BOLD_FONT
        ws.cell(row=info_row + i, column=2, value=val).font = CELL_FONT

    header_row = info_row + 6
    headers = ["行次", "项　　目", "金额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    _style_header(ws, header_row, len(headers))

    money_cols = {3}
    row = header_row + 1

    lines = [
        (1,  "一、收入总额", revenue_ytd),
        (2,  "　　减：成本", cost_ytd),
        (3,  "　　减：费用", expense_total),
        (4,  "二、利润总额（1−2−3）", profit_ytd),
        (5,  "　　减：纳税调减额", ZERO),
        (6,  "　　加：纳税调增额", ZERO),
        (7,  "三、纳税调整后所得（4−5+6）", profit_ytd),
        (8,  "　　减：弥补以前年度亏损", ZERO),
        (9,  "　　减：业主减除费用", owner_deduction),
        (10, "　　减：专项附加扣除", special_deduction),
        (11, "四、应纳税所得额（7−8−9−10）", taxable_income),
    ]

    for line_no, name, value in lines:
        _write_row(ws, row, [line_no, name, float(value)], money_cols)
        row += 1

    # 税率表
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="适用税率表（经营所得五级超额累进）").font = SUBTITLE_FONT
    row += 1

    rate_headers = ["级数", "全年应纳税所得额", "税率", "速算扣除数"]
    for i, h in enumerate(rate_headers, start=1):
        ws.cell(row=row, column=i, value=h)
    _style_header(ws, row, len(rate_headers))
    row += 1

    rate_rows = [
        (1, "不超过 30,000 元", "5%", 0),
        (2, "30,000 ~ 90,000 元", "10%", 1500),
        (3, "90,000 ~ 300,000 元", "20%", 10500),
        (4, "300,000 ~ 500,000 元", "30%", 40500),
        (5, "超过 500,000 元", "35%", 65500),
    ]

    for level, bracket, rate, deduction in rate_rows:
        _write_row(ws, row, [level, bracket, rate, deduction], {4})
        row += 1

    # 税额计算
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="应纳税额计算").font = SUBTITLE_FONT
    row += 1

    tax_lines = [
        ("应纳税额", tax_payable),
        ("减：已预缴税额", tax_prepaid),
        ("本期应补（退）税额", tax_due),
    ]

    for name, value in tax_lines:
        _write_row(ws, row, ["", name, float(value)], money_cols, bold=(name == "本期应补（退）税额"))
        row += 1

    _auto_width(ws)


# ── 主函数 ──────────────────────────────────────────────────────────────

def generate(
    report_set,            # 当期 ReportSet
    year: int,
    month: int,
    entity_type: str = "corp",   # "corp" = 有限公司 / "individual" = 个体户
    prior_report_set=None,       # 上期 ReportSet（本年累计计算需要）
    asset_total: Decimal | None = None,    # 【依赖未实现数据源】资产总额
    employee_count: int | None = None,     # 【依赖未实现数据源】从业人数
    out_path: Path | str | None = None,
) -> tuple[Path, bool]:
    """生成所得税预缴底稿 xlsx。

    Args:
        report_set: 当前期报表
        year: 年份
        month: 月份
        entity_type: "corp"（有限公司，A类）或 "individual"（个体户，经营所得个税）
        prior_report_set: 上期报表
        asset_total: 资产总额期末数（万元或元）
        employee_count: 季度平均从业人数
        out_path: 输出路径

    Returns:
        (out_path, is_template): 输出路径 + 是否为模板输出
    """
    if out_path is None:
        out_path = Path(".")
    out_path = Path(out_path)
    quarter = (month - 1) // 3 + 1

    # 判定是否模板输出
    is_template = (asset_total is None or employee_count is None)
    if is_template:
        prefix = f"{year:04d}-{month:02d}"
        entity_label = "企业所得税" if entity_type == "corp" else "经营所得个税"
        fname = f"{prefix} [模板] {entity_label}预缴底稿.xlsx"
    else:
        prefix = f"{year:04d}-{month:02d}"
        entity_label = "企业所得税" if entity_type == "corp" else "经营所得个税"
        fname = f"{prefix} {entity_label}预缴底稿.xlsx"

    if out_path.is_dir():
        out_path = out_path / fname
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    if entity_type == "corp":
        ws1 = wb.active
        ws1.title = "企业所得税A类预缴"
        _corp_income_tax_a(ws1, report_set, prior_report_set,
                          year, month, quarter, asset_total, employee_count,
                          is_template)
    else:
        ws1 = wb.active
        ws1.title = "经营所得个税预缴"
        _individual_income_tax(ws1, report_set, year, month, quarter,
                              is_template)

    wb.save(out_path)
    return out_path, is_template
