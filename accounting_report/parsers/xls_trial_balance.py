# -*- coding: utf-8 -*-
"""科目余额表 .xls → TrialBalance 条目列表

Sheet「科目余额表」，第3行表头：
科目编码 / 科目名称 / 期初余额借方 / 期初余额贷方 / 本期发生额借方 /
本期发生额贷方 / 本年累计借方 / 本年累计贷方 / 期末余额借方 / 期末余额贷方
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import List

import re
import xlrd

from ..models import money, ZERO


@dataclass
class TrialBalance:
    """科目余额表一行"""

    account_code: str      # 科目编码（如 1001）
    account_name: str      # 科目名称（如 库存现金）
    open_debit: Decimal    # 期初余额借方
    open_credit: Decimal   # 期初余额贷方
    period_debit: Decimal  # 本期发生额借方
    period_credit: Decimal # 本期发生额贷方
    ytd_debit: Decimal     # 本年累计借方
    ytd_credit: Decimal    # 本年累计贷方
    end_debit: Decimal     # 期末余额借方
    end_credit: Decimal    # 期末余额贷方

    @property
    def open_balance(self) -> Decimal:
        """期初余额（借方为正，贷方为负）"""
        return self.open_debit - self.open_credit

    @property
    def end_balance(self) -> Decimal:
        """期末余额（借方为正，贷方为负）"""
        return self.end_debit - self.end_credit


def _cell_str(row, col: int, sheet, book: xlrd.Book) -> str:
    """读取单元格字符串"""
    cell = row[col]
    try:
        ctype = cell.ctype
    except AttributeError:
        ctype = xlrd.XL_CELL_EMPTY

    if ctype == xlrd.XL_CELL_EMPTY or cell.value is None:
        return ""
    if ctype == xlrd.XL_CELL_DATE:
        dt = xlrd.xldate_as_datetime(cell.value, book.datemode)
        return dt.strftime("%Y-%m-%d")
    if ctype == xlrd.XL_CELL_NUMBER:
        s = str(cell.value)
        if s.endswith(".0") and len(s) > 2:
            s = s[:-2]
        return s
    return str(cell.value).strip()


def _clean_cell(raw: str) -> str:
    """清洗单元格：去掉换行、多余空格"""
    if not raw:
        return ""
    return raw.replace("\r", "").replace("\n", " ").strip()


def _clean_account_code(raw: str) -> str:
    """科目编码：只取首个纯数字段（丢弃换行后的重复编码/杂串）

    例如 "560206\\n560206    2025" → "560206"
    """
    if not raw:
        return ""
    for seg in raw.replace("\r", "").split("\n"):
        m = re.match(r"^\s*(\d+(?:\.\d+)?)\b", seg)
        if m:
            return m.group(1)
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\b", raw)
    return m.group(1) if m else _clean_cell(raw)


def _clean_account_name(raw: str) -> str:
    """科目名称：逐段取第一个「去除行首编码后非空」的段

    例如 "业务招待\\n招待" → "业务招待"；"560206 业务招待\\n2025" → "业务招待"
    """
    if not raw:
        return ""
    for seg in raw.replace("\r", "").split("\n"):
        seg = seg.strip()
        if not seg:
            continue
        seg2 = re.sub(r"^\d+(?:\.\d+)?\s*", "", seg).strip()
        if seg2:
            return seg2
    return _clean_cell(raw)


def parse(file_path: str | Path) -> List[TrialBalance]:
    """解析 .xls 科目余额表"""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"科目余额表文件不存在：{file_path}")

    book = xlrd.open_workbook(str(file_path))

    # 尝试找到科目余额表 sheet
    sheet = None
    for name in ("科目余额表", "试算平衡表", "余额表"):
        try:
            sheet = book.sheet_by_name(name)
            break
        except xlrd.XLRDError:
            continue

    if sheet is None:
        # 用第一个 sheet
        sheet = book.sheet_by_index(0)

    if sheet.nrows < 3:
        raise ValueError(f"{file_path.name} 的科目余额表至少需要3行（表头在第3行）")

    # 科目余额表的表头通常跨两行：
    #   R3: 科目编码 | 科目名称 | 期初余额 | (合并) | 本期发生额 | (合并) | 本年累计发生额 | (合并) | 期末余额 | (合并)
    #   R4:          |          | 借方 | 贷方 | 借方 | 贷方 | 借方 | 贷方 | 借方 | 贷方
    # 合并两行构建完整表头关键词
    header_row = sheet.row(2)
    sub_row = sheet.row(3) if sheet.nrows > 3 else None
    headers = [_clean_cell(_cell_str(header_row, i, sheet, book))
               for i in range(len(header_row))]
    sub_headers = [_clean_cell(_cell_str(sub_row, i, sheet, book))
                   for i in range(len(sub_row))] if sub_row else []

    def _find_col(*keywords: str) -> int:
        for kw in keywords:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
            # 也尝试在合并行中匹配：大类 + 子表头
            for i, h in enumerate(headers):
                sub = sub_headers[i] if i < len(sub_headers) else ""
                combined = h + sub
                if kw in combined:
                    return i
        return -1

    col_code = _find_col("科目编", "科目代码", "编码")
    col_name = _find_col("科目名", "科目")
    col_open_d = _find_col("期初", "借方") if "期初" in str(headers) else -1
    # 如果通用匹配失败，用固定列的语义：两行表头时列位置是固定的
    if col_open_d < 0:
        # 基于大类名定位: 期初余额借方=col2, 期初余额贷方=col3
        for i, h in enumerate(headers):
            if "期初" in h:
                col_open_d = i
                col_open_c = i + 1 if i + 1 < len(headers) else -1
                break
    if col_open_d < 0:
        col_open_d = 2; col_open_c = 3
    else:
        col_open_c = col_open_d + 1

    # 同上: 本期发生额借方/贷方
    col_period_d = -1
    for i, h in enumerate(headers):
        if "本期" in h:
            col_period_d = i; col_period_c = i + 1
            break
    if col_period_d < 0:
        col_period_d = 4; col_period_c = 5
    else:
        col_period_c = col_period_d + 1

    # 本年累计
    col_ytd_d = -1
    for i, h in enumerate(headers):
        if "本年累计" in h or "累计" in h:
            col_ytd_d = i; col_ytd_c = i + 1
            break
    if col_ytd_d < 0:
        col_ytd_d = 6; col_ytd_c = 7
    else:
        col_ytd_c = col_ytd_d + 1

    # 期末余额
    col_end_d = -1
    for i, h in enumerate(headers):
        if "期末" in h:
            col_end_d = i; col_end_c = i + 1
            break
    if col_end_d < 0:
        col_end_d = 8; col_end_c = 9
    else:
        col_end_c = col_end_d + 1

    missing = []
    if col_code < 0:
        missing.append("科目编码")
    if col_name < 0:
        missing.append("科目名称")
    if missing:
        raise ValueError(
            f"{file_path.name} 科目余额表缺少列：{missing}；现有表头: {headers}"
        )

    entries: List[TrialBalance] = []
    for row_idx in range(3, sheet.nrows):
        row = sheet.row(row_idx)
        # 跳过全空行
        all_empty = True
        for c in row:
            v = str(c.value).strip() if c.value is not None else ""
            if v:
                all_empty = False
                break
        if all_empty:
            continue

        code_raw = _cell_str(row, col_code, sheet, book)
        if not code_raw:
            continue

        name_raw = _cell_str(row, col_name, sheet, book)

        def _num(col: int) -> Decimal:
            if col < 0:
                return ZERO
            return money(_cell_str(row, col, sheet, book))

        entries.append(TrialBalance(
            account_code=_clean_account_code(code_raw),
            account_name=_clean_account_name(name_raw),
            open_debit=_num(col_open_d),
            open_credit=_num(col_open_c),
            period_debit=_num(col_period_d),
            period_credit=_num(col_period_c),
            ytd_debit=_num(col_ytd_d),
            ytd_credit=_num(col_ytd_c),
            end_debit=_num(col_end_d),
            end_credit=_num(col_end_c),
        ))

    if not entries:
        raise ValueError(f"{file_path.name} 的科目余额表未解析出任何条目")
    return entries
