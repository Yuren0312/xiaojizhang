# -*- coding: utf-8 -*-
"""增值税申报底稿（小规模纳税人）

生成一个 xlsx 工作簿，包含 3 个 sheet：
  1. 增值税申报主表 — 按电子税务局格式的申报底稿
  2. 附加税费计算 — 城建税/教育费附加/地方教育附加自动计算
  3. 发票与收入明细 — 发票数据与收入对应关系

数据来源：
  - 收入数据：利润表 营业收入（行次1）
  - 发票数据：发票台账（可选，已支持自动采集）——
      输入/{YYYY-MM}/收入发票/invoice_ledger.xlsx、输入/企业信息/发票台账.xlsx，
      或直接扫描输入目录下的发票 PDF 自动建台账（parsers/invoice_pdf.py）
  - 营业科目发生额：账务系统
  - 减免税优惠：小规模纳税人现行政策

⚠ 如果发票台账不存在，开票收入/无票收入无法拆分，统一归入「无票收入」，
   并在输出文件名加 [模板] 前缀。此时申报底稿仅供参考，不可直接用于申报。
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
PCT_FMT = '0.00%'
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)


# ── 小规模纳税人税收政策常量 ─────────────────────────────────────────────
# 征收率：小规模纳税人适用 3% 减按 1%（2023-2027 延续政策）
# 免税门槛：月销售额 ≤ 10 万元（季度 30 万）免征增值税
VAT_RATE = Decimal("0.03")        # 法定征收率 3%
VAT_REDUCED_RATE = Decimal("0.01")  # 优惠征收率 1%
TAX_FREE_THRESHOLD = Decimal("100000")  # 月销售额免税门槛

# 附加税费率（以实际缴纳增值税为基数）
URBAN_MAINTENANCE_TAX_RATE = Decimal("0.07")   # 城建税 7%（市区）
EDUCATION_SURCHARGE_RATE = Decimal("0.03")      # 教育费附加 3%
LOCAL_EDUCATION_SURCHARGE_RATE = Decimal("0.02")  # 地方教育附加 2%
# 小规模纳税人减半征收附加税费（六税两费减免）
HALF_RATE = Decimal("0.5")


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _write_row(ws, row: int, values: list, money_cols: set | None = None,
               bold: bool = False, fill: PatternFill | None = None,
               font: Font | None = None, pct_cols: set | None = None,
               skip_cols: set | None = None) -> None:
    """写一行数据，金额列用数字格式，百分比列用百分比格式

    skip_cols: 要跳过的列号集合（用于避免写入已合并的单元格）
    """
    if money_cols is None:
        money_cols = set()
    if pct_cols is None:
        pct_cols = set()
    if skip_cols is None:
        skip_cols = set()
    for i, v in enumerate(values, start=1):
        if i in skip_cols:
            continue
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or (BOLD_FONT if bold else CELL_FONT)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        if fill:
            cell.fill = fill
        if i in money_cols and isinstance(v, (Decimal, int, float)):
            cell.number_format = MONEY_FMT
        if i in pct_cols and isinstance(v, (Decimal, float)):
            bill_number_format = '0.00%'
            cell.number_format = bill_number_format


def _auto_width(ws, min_width: int = 10, max_width: int = 55) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            v = str(cell.value) if cell.value is not None else ""
            # 中文字符算2宽度
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            lengths.append(length)
        best = max(lengths) if lengths else min_width
        ws.column_dimensions[col_letter].width = max(min(best + 3, max_width), min_width)


# ── Sheet 1: 增值税申报主表 ─────────────────────────────────────────────

def _vat_main_table(ws, report_set, year: int, month: int,
                    invoice_data: dict | None,
                    is_template: bool,
                    adjust=None) -> None:
    """生成增值税申报主表底稿（模拟电子税务局格式）"""

    # 数据提取
    if report_set is not None and hasattr(report_set, 'is_month'):
        revenue = report_set.is_month.get(1, ZERO)  # 营业收入 行次1
    else:
        revenue = ZERO

    # 发票数据：[依赖未实现数据源] 合同台账/发票台账
    # 当前无发票台账，无法区分开票收入与无票收入
    if invoice_data:
        invoiced_revenue = money(invoice_data.get("invoiced", ZERO))
        uninvoiced_revenue = money(invoice_data.get("uninvoiced", ZERO))
        tax_exempt_revenue = money(invoice_data.get("tax_exempt", ZERO))
    else:
        invoiced_revenue = ZERO
        uninvoiced_revenue = revenue  # 全部归入无票收入
        tax_exempt_revenue = ZERO

    total_taxable = invoiced_revenue + uninvoiced_revenue

    # 判定是否适用免税（月销售额 ≤ 10 万）
    is_tax_free = total_taxable <= TAX_FREE_THRESHOLD and total_taxable > ZERO

    # 应纳税额计算（税额 = 基数 × 征收率，必须量化到分，避免多位小数写入申报表）
    if is_tax_free:
        vat_payable = ZERO
        effective_rate = ZERO
    else:
        vat_payable = money(total_taxable * VAT_REDUCED_RATE)
        effective_rate = VAT_REDUCED_RATE

    # 已缴税额：来自「税务调整」页录入（预缴税款台账格式不统一，不做解析）
    vat_prepaid = money(getattr(adjust, "vat_prepaid", ZERO))

    # 本期应补（退）
    vat_due = vat_payable - vat_prepaid

    # ── 标题行 ──
    ws.merge_cells("A1:G1")
    title = (f"{year}年{month}月 增值税及附加税费申报表（小规模纳税人适用）"
             f"{' [模板]' if is_template else ''}")
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    if is_template:
        ws.merge_cells("A2:G2")
        ws.cell(row=2, column=1,
                value="⚠ 此为模板级输出，发票台账数据源缺失，开票/无票收入已合并。请补充发票台账后重新生成。"
                ).font = WARN_FONT
        ws.cell(row=2, column=1).fill = WARN_FILL

    # ── 基础信息 ──
    info_row = 3 if not is_template else 4
    info_items = [
        ("纳税人识别号", "【依赖未实现数据源】纳税人基础信息"),
        ("纳税人名称", "示例科技有限公司"),
        ("所属行业", "【依赖未实现数据源】纳税人基础信息"),
        ("纳税期限", "月"),
    ]
    for i, (label, val) in enumerate(info_items):
        ws.cell(row=info_row + i, column=1, value=label).font = BOLD_FONT
        ws.cell(row=info_row + i, column=2, value=val).font = CELL_FONT

    header_row = info_row + 5

    # ── 表头 ──
    headers = ["项目", "栏次", "计税依据\n（不含税）", "征收率", "本期应纳税额",
               "本期已缴税额", "本期应补（退）税额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    _style_header(ws, header_row, len(headers))

    money_cols = {3, 5, 6, 7}
    pct_cols = {4}
    row = header_row + 1

    # ── 一、计税依据 ──
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="一、计税依据").font = SECTION_FONT
    ws.cell(row=row, column=1).fill = SECTION_FILL
    row += 1

    # 第1栏：应征增值税不含税销售额
    _write_row(ws, row, [
        "（一）应征增值税不含税销售额",
        1,
        float(total_taxable),
        None,
        None, None, None,
    ], money_cols, pct_cols=pct_cols)
    row += 1

    # 第2栏：其中：开具增值税专用发票不含税销售额
    # 发票台账 → 专票不含税金额（未提供台账时为 0，该列留空）
    special_invoice_amount = ZERO
    if invoice_data:
        special_invoice_amount = money(invoice_data.get("special_invoice", ZERO))
    _write_row(ws, row, [
        "　其中：开具增值税专用发票销售额",
        2,
        float(special_invoice_amount),
        None,
        None, None, None,
    ], money_cols, pct_cols=pct_cols)
    row += 1

    # 第3栏：开具普通发票不含税销售额
    # 发票台账 → 普票不含税金额（未提供台账时为 0，该列留空）
    _write_row(ws, row, [
        "　其中：开具普通发票销售额",
        3,
        float(invoiced_revenue),
        None,
        None, None, None,
    ], money_cols, pct_cols=pct_cols)
    row += 1

    # ── 二、税款计算 ──
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="二、税款计算").font = SECTION_FONT
    ws.cell(row=row, column=1).fill = SECTION_FILL
    row += 1

    # 第4栏：本期应纳税额
    _write_row(ws, row, [
        "本期应纳税额",
        4,
        float(total_taxable),
        f"{float(VAT_REDUCED_RATE)*100:.0f}%",
        float(vat_payable),
        None, None,
    ], money_cols, pct_cols=pct_cols)
    row += 1

    # 第5栏：应纳税额减征额（法定3%减按1%优惠）
    tax_reduction = money(total_taxable * (VAT_RATE - VAT_REDUCED_RATE)) if not is_tax_free else ZERO
    _write_row(ws, row, [
        "本期应纳税额减征额（减按1%）",
        5,
        None,
        None,
        float(tax_reduction),
        None, None,
    ], money_cols, pct_cols=pct_cols)
    row += 1

    # 第6栏：免税销售额
    _write_row(ws, row, [
        "（二）免税销售额",
        6,
        float(tax_exempt_revenue) if not is_tax_free else float(total_taxable),
        "免税",
        ZERO,
        None, None,
    ], money_cols, pct_cols=pct_cols)

    if is_tax_free:
        ws.cell(row=row, column=5, value="月销售额≤10万，免征").font = WARN_FONT
    row += 1

    # 第7栏：本期已缴税额
    _write_row(ws, row, [
        "本期已缴税额",
        7,
        None, None, None,
        float(vat_prepaid),
        None,
    ], money_cols, pct_cols=pct_cols)
    row += 1

    # ── 合计行 ──
    row += 1
    _write_row(ws, row, [
        "本期应补（退）税额",
        "",
        None, None, None, None,
        float(vat_due),
    ], money_cols, pct_cols=pct_cols, bold=True)
    ws.cell(row=row, column=1).fill = SECTION_FILL

    # ── 数据源说明 ──
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    notes = [
        f"数据源说明：",
        f"• 营业收入（行1）= {float(revenue):,.2f} 元，来自利润表",
        f"• 开票收入拆分：发票台账自动解析（销项 {float(invoiced_revenue):,.2f} / 无票收入 {float(uninvoiced_revenue):,.2f}）",
        f"• 进项发票（购方=本公司）不计入销售额，明细见「发票与收入明细」",
        f"• 征收率：小规模纳税人适用 3% 减按 1%（2023-2027 延续政策）",
        f"• 免税判定：月销售额 {'≤' if is_tax_free else '>'} 10 万元 → {'免税' if is_tax_free else '征税'}",
    ]
    if is_template:
        notes.append("• ⚠ 当前为模板输出，缺少发票数据，不可直接用于申报")
    for note in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        ws.cell(row=row, column=1, value=note).font = Font(name="微软雅黑", size=9, color="666666")
        row += 1

    _auto_width(ws)


# ── Sheet 2: 附加税费计算 ───────────────────────────────────────────────

def _surtax_calculation(ws, vat_payable: Decimal, is_tax_free: bool,
                        is_template: bool) -> None:
    """城建税 / 教育费附加 / 地方教育附加自动计算"""

    ws.merge_cells("A1:E1")
    ws.cell(row=1, column=1,
            value=f"附加税费计算表{' [模板]' if is_template else ''}").font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    # 小规模纳税人减半征收（六税两费减半）
    actual_rate_multiplier = HALF_RATE  # 减半征收

    urban_tax = money(vat_payable * URBAN_MAINTENANCE_TAX_RATE * actual_rate_multiplier)
    edu_surcharge = money(vat_payable * EDUCATION_SURCHARGE_RATE * actual_rate_multiplier)
    local_edu_surcharge = money(vat_payable * LOCAL_EDUCATION_SURCHARGE_RATE * actual_rate_multiplier)
    total_surtax = money(urban_tax + edu_surcharge + local_edu_surcharge)

    headers = ["税种", "计税依据（增值税）", "法定税率", "减免后实际税率", "应纳税额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    _style_header(ws, 3, len(headers))

    money_cols = {2, 5}
    pct_cols = {3, 4}

    surtax_items = [
        ("城市维护建设税", URBAN_MAINTENANCE_TAX_RATE, urban_tax),
        ("教育费附加", EDUCATION_SURCHARGE_RATE, edu_surcharge),
        ("地方教育附加", LOCAL_EDUCATION_SURCHARGE_RATE, local_edu_surcharge),
    ]

    row = 4
    for name, legal_rate, tax_amount in surtax_items:
        effective_rate = legal_rate * actual_rate_multiplier
        _write_row(ws, row, [
            name,
            float(vat_payable),
            f"{float(legal_rate)*100:.0f}%",
            f"{float(effective_rate)*100:.1f}%",
            float(tax_amount),
        ], money_cols, pct_cols=pct_cols)
        row += 1

    # 合计行
    _write_row(ws, row, [
        "合　　计",
        float(vat_payable),
        "",
        "",
        float(total_surtax),
    ], money_cols, pct_cols=pct_cols, bold=True, fill=SECTION_FILL)

    # 说明行
    row += 2
    notes = [
        "说明：",
        f"• 附加税费 = 实际缴纳增值税 × 适用税率 × {'50%（减半）' if actual_rate_multiplier == HALF_RATE else '100%'}",
        "• 小规模纳税人享受六税两费减半征收优惠（截至2027年底）",
        "• 城市维护建设税 7%（市区）/ 5%（县城、镇）/ 1%（其他）",
        "• 数据自动填入利润表「税金及附加」行次3",
    ]
    for note in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.cell(row=row, column=1, value=note).font = Font(name="微软雅黑", size=9, color="666666")
        row += 1

    # 特殊说明
    if is_tax_free:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.cell(row=row, column=1,
                value="✅ 本期增值税免税，附加税费无需缴纳。").font = Font(
                    name="微软雅黑", size=10, color="338833", bold=True)

    _auto_width(ws)


# ── Sheet 3: 发票与收入明细 ─────────────────────────────────────────────

def _invoice_detail(ws, report_set, invoice_data: dict | None,
                    is_template: bool, year: int, month: int) -> None:
    """发票开具情况与营业收入的对应关系"""

    ws.merge_cells("A1:H1")
    ws.cell(row=1, column=1,
            value=f"发票与收入对应明细{' [模板]' if is_template else ''}").font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    if report_set is not None and hasattr(report_set, 'is_month'):
        revenue = report_set.is_month.get(1, ZERO)
    else:
        revenue = ZERO

    headers = ["序号", "发票代码", "发票号码", "开票日期", "购方名称",
               "货物/服务名称", "不含税金额", "税额", "方向"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    _style_header(ws, 3, len(headers))

    money_cols = {7, 8}

    # 发票台账（自动解析发票 PDF 生成）：进项（购方=本公司）不参与销售额
    if invoice_data and "details" in invoice_data:
        for idx, inv in enumerate(invoice_data["details"], start=1):
            _write_row(ws, 3 + idx, [
                idx,
                inv.get("code", ""),
                inv.get("number", ""),
                inv.get("date", ""),
                inv.get("buyer", ""),
                inv.get("service", ""),
                float(money(inv.get("amount", ZERO))),
                float(money(inv.get("tax", ZERO))),
                inv.get("direction", ""),
            ], money_cols)
        detail_row = 3 + len(invoice_data["details"]) + 1
    else:
        # 模板模式：占位行
        _write_row(ws, 4, [
            1,
            "【模板】请补充发票台账",
            "【模板】输入/发票台账.xlsx",
            f"{year}-{month:02d}",
            "【模板】购方名称",
            "【模板】服务名称",
            float(revenue) if revenue > ZERO else 0.0,
            float(money(revenue * VAT_REDUCED_RATE)) if revenue > ZERO else 0.0,
            "销项",
        ], money_cols, fill=TEMPLATE_FILL)
        detail_row = 5

    # 汇总行（进项=购方为本公司的费用票，不参与销售额申报）
    detail_row += 1
    if invoice_data and "details" in invoice_data:
        details = invoice_data["details"]
        # 销项合计只统计 direction == "销项"：方向「待确认」（购方未识别）的
        # 发票在主表 invoiced 里被排除，这里若用 != "进项" 会把待确认票也算进
        # 销项合计，造成 sheet3 与主表「参与销售额」口径不一致（白班12 修复）
        sale_amount = sum(money(d.get("amount", ZERO)) for d in details
                          if d.get("direction") == "销项")
        sale_tax = sum(money(d.get("tax", ZERO)) for d in details
                       if d.get("direction") == "销项")
        input_amount = sum(money(d.get("amount", ZERO)) for d in details
                           if d.get("direction") == "进项")
        input_tax = sum(money(d.get("tax", ZERO)) for d in details
                        if d.get("direction") == "进项")
    else:
        sale_amount = sale_tax = input_amount = input_tax = ZERO

    ws.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=6)
    _write_row(ws, detail_row, [
        "销项开票合计（参与销售额）", "", "", "", "", "",
        float(sale_amount),
        float(sale_tax),
        "",
    ], money_cols, bold=True, skip_cols={2, 3, 4, 5, 6})
    detail_row += 1

    ws.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=6)
    _write_row(ws, detail_row, [
        "进项发票合计（费用票，不参与销售额）", "", "", "", "", "",
        float(input_amount),
        float(input_tax),
        "",
    ], money_cols, bold=True, fill=SECTION_FILL, skip_cols={2, 3, 4, 5, 6})
    detail_row += 1

    total_invoiced = sale_amount

    # 无票收入：优先取发票台账解析结果；缺失时按 ① − ② 推算（此时 ④ 自然为 0）
    if invoice_data and "uninvoiced" in invoice_data:
        uninvoiced_revenue = money(invoice_data.get("uninvoiced", ZERO))
    else:
        uninvoiced_revenue = revenue - total_invoiced

    # 收入核对
    detail_row += 1
    ws.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=8)
    ws.cell(row=detail_row, column=1, value="收入核对").font = SUBTITLE_FONT
    detail_row += 1

    # ④ 差异 = ① − (② + ③无票收入)。补差后 ① = ② + ③ + 免税收入，
    # 差异自然等于免税收入——这是正常差异（免税票不计入销项发票合计），
    # 不应提示"需核查"；若差异为 0 也正常（无免税收入时）。白班12 修复。
    check_items = [
        ("① 利润表营业收入", float(revenue), ""),
        ("② 销项发票不含税金额合计", float(total_invoiced), ""),
        ("③ 无票收入", float(uninvoiced_revenue),
         "发票台账自动解析" if not is_template else "未提供发票台账，无法准确拆分"),
        ("④ 差异（① − ② − ③）",
         float(revenue - total_invoiced - uninvoiced_revenue),
         "应为免税收入金额（免税票不计入销项合计，属正常差异）" if not is_template else ""),
    ]

    for item, val, note in check_items:
        _write_row(ws, detail_row, [item, val, note], {2})
        detail_row += 1

    # 补充说明
    detail_row += 1
    ws.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=8)
    notes_text = [
        "数据源说明：",
        "• 发票明细：输入/发票台账.xlsx（自动解析发票 PDF 生成，或用户手工维护）",
        "• 进项发票（购方=本公司）不参与销售额申报，单独列示（费用票）",
        "• 方向「待确认」的发票（购方未识别）不计入销售额，请在台账中核对",
    ]
    if is_template:
        notes_text.append("• ⚠ 当前为模板输出，请补充发票台账后重新运行以生成完整申报底稿")
    for note in notes_text:
        ws.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=8)
        ws.cell(row=detail_row, column=1, value=note).font = Font(name="微软雅黑", size=9, color="666666")
        detail_row += 1

    _auto_width(ws)


# ── 主函数 ──────────────────────────────────────────────────────────────

def generate(
    report_set,         # ReportSet 或 None
    year: int,
    month: int,
    invoice_data: dict | None = None,  # 发票台账解析结果（可选；未提供时按无票处理）
    adjust=None,                       # 税务调整项（取其中的 vat_prepaid）
    out_path: Path | str | None = None,
) -> tuple[Path, bool]:
    """生成增值税申报底稿 xlsx。

    Args:
        report_set: 计算好的报表（从中提取营业收入等）
        year: 年份
        month: 月份
        invoice_data: 发票台账数据。如为 None，生成模板级输出。
            dict 格式:
            {
                "invoiced": Decimal,       # 开票不含税收入
                "uninvoiced": Decimal,     # 无票收入
                "tax_exempt": Decimal,     # 免税收入
                "special_invoice": Decimal, # 专票不含税
                "details": [               # 发票明細
                    {
                        "code": "发票代码",
                        "number": "发票号码",
                        "date": "开票日期",
                        "buyer": "购方",
                        "service": "服务",
                        "amount": Decimal,  # 不含税
                        "tax": Decimal,     # 税额
                    },
                    ...
                ],
            }
        out_path: 输出路径。为 None 时自动生成。

    Returns:
        (out_path, is_template): 输出文件路径 + 是否为模板
    """
    if out_path is None:
        out_path = Path(".")

    out_path = Path(out_path)
    if out_path.is_dir():
        prefix = f"{year:04d}-{month:02d}"
        # 无发票数据 → [模板] 前缀
        if invoice_data is None:
            fname = f"{prefix} [模板] 增值税申报底稿.xlsx"
        else:
            fname = f"{prefix} 增值税申报底稿.xlsx"
        out_path = out_path / fname

    out_path.parent.mkdir(parents=True, exist_ok=True)
    is_template = invoice_data is None

    wb = Workbook()

    # 数据准备
    if report_set is not None and hasattr(report_set, 'is_month'):
        revenue = report_set.is_month.get(1, ZERO)
    else:
        revenue = ZERO

    if invoice_data:
        # 无票收入补差：invoice_ledger 的 uninvoiced 是预留 0，若利润表营业收入
        # 大于开票收入 + 免税收入，差额就是无票收入。不补差时主表计税依据只含
        # 开票收入，与 accrue_taxes（按利润表收入计提附加税）口径不一致，且
        # 用户看到"收入 15 万、申报底稿只算 9 万"会误以为漏票（白班12 修复）。
        # 注意：这里就地修正 dict 副本，让 _vat_main_table / _invoice_detail
        # 读到一致的 uninvoiced，避免三处各自读原始键得到 0。
        invoice_data = dict(invoice_data)
        invoiced = money(invoice_data.get("invoiced", ZERO))
        uninvoiced = money(invoice_data.get("uninvoiced", ZERO))
        tax_exempt = money(invoice_data.get("tax_exempt", ZERO))
        if uninvoiced == ZERO and revenue > invoiced + tax_exempt:
            uninvoiced = max(revenue - invoiced - tax_exempt, ZERO)
            invoice_data["uninvoiced"] = uninvoiced
        total_taxable = invoiced + uninvoiced
    else:
        total_taxable = revenue

    is_tax_free = total_taxable <= TAX_FREE_THRESHOLD and total_taxable > ZERO
    if is_tax_free:
        vat_payable = ZERO
    else:
        vat_payable = money(total_taxable * VAT_REDUCED_RATE)

    # Sheet 1: 增值税申报主表
    ws1 = wb.active
    ws1.title = "增值税申报主表"
    _vat_main_table(ws1, report_set, year, month, invoice_data, is_template,
                    adjust=adjust)

    # Sheet 2: 附加税费计算
    ws2 = wb.create_sheet("附加税费计算")
    _surtax_calculation(ws2, vat_payable, is_tax_free, is_template)

    # Sheet 3: 发票与收入明细
    ws3 = wb.create_sheet("发票与收入明细")
    _invoice_detail(ws3, report_set, invoice_data, is_template, year, month)

    wb.save(out_path)
    return out_path, is_template
