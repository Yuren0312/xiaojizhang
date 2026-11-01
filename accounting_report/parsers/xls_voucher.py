# -*- coding: utf-8 -*-
"""示例科技有限公司 .xls 凭证文件 → VoucherEntry 列表

数据结构：Sheet「凭证列表」，第2行是表头（日期/凭证字号/摘要/科目/借方金额/贷方金额）。
同一凭证字号可能跨多行（一证多分录），按顺序读取后归组即可。
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
class VoucherEntry:
    """一条凭证明细"""

    date: str             # YYYY-MM-DD
    voucher_no: str       # 凭证字号（记-1 / 收-5 等）
    summary: str          # 摘要
    account_code: str     # 科目编码（如 560303）
    account_name: str     # 科目名称（如 财务费用_银行手续费）
    debit: Decimal        # 借方金额
    credit: Decimal       # 贷方金额

    @property
    def amount(self) -> Decimal:
        """净额：借方正数，贷方负数，平账时为零"""
        return self.debit - self.credit


def _cell_str(row, col: int, sheet, book: xlrd.Book) -> str:
    """读取单元格字符串，处理 None / 数字 / 日期"""
    cell = row[col]
    ctype = sheet.cell_type(row._Row__idx, col) if hasattr(row, '_Row__idx') else cell.ctype
    # Try ctype attribute first
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
        # 整数可能是科目编码
        s = str(cell.value)
        if s.endswith(".0") and len(s) > 2:
            s = s[:-2]
        return s
    return str(cell.value).strip()


def _clean_account_name(raw: str) -> str:
    """科目名称可能含换行或多余编号，逐段取第一个「编码 名称」的名称部分。

    例如 "560206 管理费用_招待\\n560206 2025" → "管理费用_招待"
         "560206\\n560206 2025" → 无名称段，兜底去编码后返回
    """
    if not raw:
        return ""
    for seg in raw.replace("\r", "").split("\n"):
        seg = seg.strip()
        if not seg:
            continue
        parts = seg.split(None, 1)
        if len(parts) == 2 and parts[0].replace(".", "").isdigit():
            return parts[1].strip()
    for seg in raw.replace("\r", "").split("\n"):
        seg2 = re.sub(r"^\d+(?:\.\d+)?\s*", "", seg.strip()).strip()
        if seg2:
            return seg2
    return raw.strip()


def parse(file_path: str | Path) -> List[VoucherEntry]:
    """解析凭证文件（.xls 用 xlrd，.xlsx 用 openpyxl），返回 VoucherEntry 列表"""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"凭证文件不存在：{file_path}")

    if file_path.suffix.lower() == ".xlsx":
        return _parse_xlsx(file_path)
    return _parse_xls(file_path)


def _parse_xlsx(file_path: Path) -> List[VoucherEntry]:
    """解析 .xlsx 凭证（openpyxl；xlrd 不支持 xlsx，此前 xlsx 凭证必解析失败）"""
    from openpyxl import load_workbook
    wb = load_workbook(str(file_path), data_only=True)
    ws = wb["凭证列表"] if "凭证列表" in wb.sheetnames else wb.active

    def _val(cell) -> str:
        v = cell.value
        if v is None:
            return ""
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        return str(v).strip()

    rows = list(ws.iter_rows())
    if len(rows) < 2:
        raise ValueError(f"{file_path.name} 的「凭证列表」sheet 至少需要2行（表头在第2行）")
    headers = [_val(c) for c in rows[1]]

    def _find(*names: str) -> int | None:
        for name in names:
            for i, h in enumerate(headers):
                if name in h:
                    return i
        return None

    col_date = _find("日期")
    col_voucher = _find("凭证字号")
    col_summary = _find("摘要")
    col_account = _find("科目")
    col_debit = _find("借方")
    col_credit = _find("贷方")
    missing = [n for n, c in (("日期", col_date), ("凭证字号", col_voucher),
                              ("摘要", col_summary), ("科目", col_account),
                              ("借方", col_debit), ("贷方", col_credit)) if c is None]
    if missing:
        raise ValueError(f"凭证列表缺少列 {missing}，现有表头: {headers}")

    entries: List[VoucherEntry] = []
    carry_date = carry_voucher = carry_summary = ""
    for row in rows[2:]:
        raw_account = _val(row[col_account])
        if not raw_account:
            continue
        date_str = _val(row[col_date])
        if date_str:
            carry_date = date_str
            carry_voucher = _val(row[col_voucher])
            carry_summary = _val(row[col_summary])
        else:
            date_str = carry_date
        voucher_no = _val(row[col_voucher]) or carry_voucher
        summary = _val(row[col_summary]) or carry_summary
        debit = money(_val(row[col_debit]))
        credit = money(_val(row[col_credit]))
        code_and_name = raw_account.replace("\r", "").split("\n")[0].strip()
        tokens = code_and_name.split(None, 1)
        if len(tokens) >= 2 and tokens[0].replace(".", "").isdigit():
            account_code, account_name = tokens[0], tokens[1].strip()
        else:
            account_code, account_name = "", code_and_name
        entries.append(VoucherEntry(
            date=date_str, voucher_no=voucher_no, summary=summary,
            account_code=account_code, account_name=account_name,
            debit=debit, credit=credit,
        ))
    wb.close()
    if not entries:
        raise ValueError(f"{file_path.name} 的凭证列表未解析出任何分录")
    return entries


def _parse_xls(file_path: Path) -> List[VoucherEntry]:
    """解析 .xls 凭证（xlrd）"""

    book = xlrd.open_workbook(str(file_path))
    sheet = book.sheet_by_name("凭证列表")

    # 第2行（0-indexed row 1）是表头
    if sheet.nrows < 2:
        raise ValueError(f"{file_path.name} 的「凭证列表」sheet 至少需要2行（表头在第2行）")

    header_row = sheet.row(1)
    headers = [str(c.value).strip() if c.value is not None else "" for c in header_row]

    # 定位列索引
    try:
        col_date = headers.index("日期")
    except ValueError:
        # 尝试模糊匹配
        col_date = next((i for i, h in enumerate(headers) if "日期" in h), None)
        if col_date is None:
            raise ValueError(f"凭证列表缺少「日期」列，现有表头: {headers}")

    col_voucher = next((i for i, h in enumerate(headers) if "凭证字号" in h), None)
    if col_voucher is None:
        raise ValueError(f"凭证列表缺少「凭证字号」列，现有表头: {headers}")

    col_summary = next((i for i, h in enumerate(headers) if "摘要" in h and "科目" not in h), None)
    if col_summary is None:
        raise ValueError(f"凭证列表缺少「摘要」列，现有表头: {headers}")

    col_account = next((i for i, h in enumerate(headers) if "科目" in h), None)
    if col_account is None:
        raise ValueError(f"凭证列表缺少「科目」列，现有表头: {headers}")

    col_debit = next((i for i, h in enumerate(headers) if "借方" in h), None)
    if col_debit is None:
        raise ValueError(f"凭证列表缺少「借方金额」列，现有表头: {headers}")

    col_credit = next((i for i, h in enumerate(headers) if "贷方" in h), None)
    if col_credit is None:
        raise ValueError(f"凭证列表缺少「贷方金额」列，现有表头: {headers}")

    entries: List[VoucherEntry] = []
    # 财务软件导出的凭证，同一凭证多行分录只有首行填日期+凭证字号，
    # 后续行日期和凭证字号为空，需要从上一行继承。
    carry_date: str = ""
    carry_voucher_no: str = ""
    carry_summary: str = ""
    # 从第3行（索引2）开始读数据
    for row_idx in range(2, sheet.nrows):
        row = sheet.row(row_idx)
        # 跳过完全空行
        if all(c.value is None or str(c.value).strip() == "" for c in row):
            continue

        raw_account = _cell_str(row, col_account, sheet, book)
        # 科目列为空的行跳过（可能是标题行或空行）
        if not raw_account:
            continue

        date_str = _cell_str(row, col_date, sheet, book)
        if date_str:
            carry_date = date_str
            carry_voucher_no = _cell_str(row, col_voucher, sheet, book)
            carry_summary = _cell_str(row, col_summary, sheet, book)
        else:
            # 继承上一行的日期/凭证字号/摘要（多行分录）
            date_str = carry_date
            voucher_no_str = carry_voucher_no
            summary_str = carry_summary

        voucher_no = _cell_str(row, col_voucher, sheet, book) or carry_voucher_no
        summary = _cell_str(row, col_summary, sheet, book) or carry_summary
        debit = money(_cell_str(row, col_debit, sheet, book))
        credit = money(_cell_str(row, col_credit, sheet, book))

        # 解析科目编码和科目名称
        # 格式可能是 "560303 财务费用_银行手续费" 或换行分隔
        if raw_account:
            # 按空白分割第一个词当编码，其余当名称
            account_parts = _clean_account_name(raw_account)
            code_and_name = raw_account.replace("\r", "").split("\n")[0].strip()
            # 再拆分编码和名称
            tokens = code_and_name.split(None, 1)
            if len(tokens) >= 2 and tokens[0].replace(".", "").isdigit():
                account_code = tokens[0]
                account_name = tokens[1].strip()
            else:
                account_code = ""
                account_name = code_and_name
        else:
            account_code = ""
            account_name = ""

        entries.append(VoucherEntry(
            date=date_str,
            voucher_no=voucher_no,
            summary=summary,
            account_code=account_code,
            account_name=account_name,
            debit=debit,
            credit=credit,
        ))

    if not entries:
        raise ValueError(f"{file_path.name} 的凭证列表未解析出任何分录")
    return entries
