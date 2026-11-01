# -*- coding: utf-8 -*-
"""记账凭证导出（体验官白班轮新增 · 输出 8）

补上产品闭环缺失的一环：明细账/报表都有，唯独没有「记账凭证」。
从 JournalEntry 生成标准记账凭证 Excel：
  - 按 (日期, 摘要) 自然分组 → 一张凭证
  - 每张凭证：记字 N 号，借/贷两行（或组内多行），含凭证小计
  - 表头与底部携带 制单人 / 审核人（JournalEntry.prepared_by / reviewer，
    模型已支持，不写死人名，由上层接入真实成员）

输出：{year}-{month} 记账凭证.xlsx（Sheet「记账凭证」）
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ..models import ZERO

TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
CELL_FONT = Font(name="微软雅黑", size=10)
BOLD_FONT = Font(name="微软雅黑", bold=True, size=10)
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = '#,##0.00'


def generate(
    entries: list,
    year: int,
    month: int,
    out_path: Path | str | None = None,
) -> Path:
    """生成记账凭证 xlsx。

    Args:
        entries: JournalEntry 列表（已按月过滤）。
        year/month: 账期。
        out_path: 输出路径（目录或文件）。
    """
    out_path = Path(out_path) if out_path else Path(".")
    if out_path.is_dir():
        out_path = out_path / f"{year:04d}-{month:02d} 记账凭证.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "记账凭证"

    # ── 标题 ──
    ws.merge_cells("A1:E1")
    c = ws.cell(1, 1, f"{year}年{month}月 记账凭证（共 {len(entries)} 笔分录）")
    c.font = TITLE_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── 表头 ──
    headers = ["凭证字号", "日期", "摘要", "科目名称", "借方金额", "贷方金额"]
    for j, h in enumerate(headers, 1):
        cell = ws.cell(3, j, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    # ── 按 (日期, 摘要) 分组 ──
    # 先按日期升序排，避免凭证号随输入顺序错乱（银行流水跨文件/业务事项混排时）
    groups: dict[tuple, list] = defaultdict(list)
    order: list[tuple] = []
    for e in sorted(
        entries,
        key=lambda x: (str(getattr(x, "date", "")), str(getattr(x, "summary", ""))),
    ):
        key = (str(getattr(e, "date", "")), str(getattr(e, "summary", "")))
        if key not in groups:
            order.append(key)
        groups[key].append(e)

    row = 4
    voucher_no = 0
    total_debit = ZERO
    total_credit = ZERO

    for key in order:
        voucher_no += 1
        date, summary = key
        group = groups[key]

        # 按科目归并借贷发生额：一借多贷/多借一贷/多借多贷的凭证会被拆成多个
        # 借贷对（每对带共享科目的全额），逐条输出会把共享科目金额摊错——如
        # 「借银行存款300 / 贷A100 / 贷B200」会渲染成 借100/贷100 + 借200/贷200，
        # 读起来像两笔错误的凭证（白班14 修复）。按科目汇总后输出标准记账凭证行：
        #   借 银行存款_中国工商银行 300 / 贷 短期投资_理财A 100 / 贷 …B 200
        debit_legs: dict[str, Decimal] = defaultdict(Decimal)
        credit_legs: dict[str, Decimal] = defaultdict(Decimal)
        for e in group:
            amount = getattr(e, "amount", ZERO) or ZERO
            dr = str(getattr(e, "debit", "")).strip()
            cr = str(getattr(e, "credit", "")).strip()
            if dr:
                debit_legs[dr] += amount
            if cr:
                credit_legs[cr] += amount

        # 标准排版：借方科目行在前，贷方科目行在后；金额只落在对应列
        legs = ([(acct, amt, True) for acct, amt in debit_legs.items()]
                + [(acct, amt, False) for acct, amt in credit_legs.items()])
        for acct, amt, is_debit in legs:
            if is_debit:
                total_debit += amt
            else:
                total_credit += amt

            ws.cell(row, 1, f"记字 {voucher_no} 号").border = BORDER
            ws.cell(row, 2, date).border = BORDER
            ws.cell(row, 3, summary).border = BORDER
            ws.cell(row, 4, acct).border = BORDER
            dc = ws.cell(row, 5)
            dc.value = float(amt) if is_debit else None
            dc.number_format = MONEY_FMT
            dc.border = BORDER
            cc = ws.cell(row, 6)
            cc.value = float(amt) if not is_debit else None
            cc.number_format = MONEY_FMT
            cc.border = BORDER
            for col in range(1, 7):
                ws.cell(row, col).font = CELL_FONT
                ws.cell(row, col).alignment = Alignment(vertical="center")
            row += 1

        # 凭证小计行（按归并后的借贷合计；若组内借贷不平则标红提示——审计文档
        # 宁可暴露问题，不出一张看着平衡实际有误的凭证）
        subtotal_d = sum(debit_legs.values())
        subtotal_c = sum(credit_legs.values())
        balanced = subtotal_d == subtotal_c
        ws.cell(row, 3, "本凭证合计" + ("" if balanced else " ⚠ 借贷不平！")).font = BOLD_FONT
        ws.cell(row, 3).border = BORDER
        if not balanced:
            ws.cell(row, 3).font = Font(name="微软雅黑", bold=True, size=10, color="C62828")
        sd = ws.cell(row, 5, float(subtotal_d))
        sd.number_format = MONEY_FMT
        sd.font = BOLD_FONT
        sd.border = BORDER
        sc = ws.cell(row, 6, float(subtotal_c))
        sc.number_format = MONEY_FMT
        sc.font = BOLD_FONT
        sc.border = BORDER
        for col in (1, 2, 4):
            ws.cell(row, col).border = BORDER
        row += 1

    # ── 合计行 ──
    ws.cell(row, 3, "合计").font = BOLD_FONT
    ws.cell(row, 3).border = BORDER
    td = ws.cell(row, 5, float(total_debit))
    td.number_format = MONEY_FMT
    td.font = BOLD_FONT
    td.border = BORDER
    tc = ws.cell(row, 6, float(total_credit))
    tc.number_format = MONEY_FMT
    tc.font = BOLD_FONT
    tc.border = BORDER
    for col in (1, 2, 4):
        ws.cell(row, col).border = BORDER
    row += 2

    # ── 制证/审核信息（取自分录字段，不写死人名） ──
    prepared = {str(getattr(e, "prepared_by", "")) for e in entries} - {""}
    reviewer = {str(getattr(e, "reviewer", "")) for e in entries} - {""}
    ws.cell(row, 1, f"制证人：{'、'.join(prepared) if prepared else '待接入成员'}").font = CELL_FONT
    ws.cell(row, 4, f"审核人：{'、'.join(reviewer) if reviewer else '待审核'}").font = CELL_FONT

    # ── 列宽 ──
    for col, w in zip("ABCDEF", (14, 12, 26, 36, 14, 14)):
        ws.column_dimensions[col].width = w

    wb.save(out_path)
    return out_path
