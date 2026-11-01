# -*- coding: utf-8 -*-
"""明细账 .xlsx → 按科目分组的交易列表

明细账文件包含多个科目区域，每个区域以标题行「XXXX 科目名 明细账」开头。
第4行是表头：日期 / 凭证号 / 摘要 / 借方 / 贷方 / 方向 / 余额。
需要识别科目区域边界，输出按科目分组的交易列表。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Dict, List

import openpyxl

from ..models import money, ZERO


# 科目区域标题行匹配：「科目编码 科目名 明细账」或「科目名 明细账」
SECTION_RE = re.compile(r"^(.+?)\s*明细账\s*$")

# 表头特征关键词（第4行）
HEADER_KEYWORDS = ["日期", "凭证号", "摘要", "借方", "贷方", "方向", "余额"]


@dataclass
class LedgerEntry:
    """明细账一行"""

    date: str             # YYYY-MM-DD
    voucher_no: str       # 凭证号
    summary: str          # 摘要
    debit: Decimal        # 借方金额
    credit: Decimal       # 贷方金额
    direction: str        # 方向（借/贷/平）
    balance: Decimal      # 余额


@dataclass
class AccountLedger:
    """一个科目的明细账"""

    account_code: str     # 科目编码
    account_name: str     # 科目名称
    entries: List[LedgerEntry] = field(default_factory=list)


def _norm_date(cell_value) -> str:
    """统一日期为 YYYY-MM-DD"""
    if cell_value is None:
        return ""
    if hasattr(cell_value, "strftime"):
        return cell_value.strftime("%Y-%m-%d")
    s = str(cell_value).strip()
    # 可能是纯数字（Excel 日期序列号）
    return s


def _parse_section_title(text: str) -> tuple[str, str] | None:
    """从标题行解析科目编码和科目名称。

    "560303 财务费用_银行手续费 明细账" → ("560303", "财务费用_银行手续费")
    返回 None 表示不是科目标题行。
    """
    text = (text or "").strip()
    m = SECTION_RE.match(text)
    if not m:
        return None

    prefix = m.group(1).strip()
    # 尝试拆分编码和名称
    parts = prefix.split(None, 1)  # 按空白分割一次
    if len(parts) >= 2:
        first = parts[0]
        second = parts[1]
        # 如果第一部分是纯数字（可能含点），则是编码
        if first.replace(".", "").isdigit():
            return (first, second)
        else:
            # 没有编码，全是名称
            return ("", prefix)
    else:
        return ("", prefix)


def _is_header_row(row_values: list) -> bool:
    """判断某行是否为明细账表头（日期 / 凭证号 / 摘要 / 借方 / 贷方 / 方向 / 余额）"""
    texts = [str(v).strip() if v is not None else "" for v in row_values]
    if not texts or all(t == "" for t in texts):
        return False
    matched = sum(1 for kw in HEADER_KEYWORDS if any(kw in t for t in texts))
    return matched >= 5  # 宽松匹配：只要5个关键词有匹配即认为表头


def _is_empty_row(row_values: list) -> bool:
    """判断是否为空行"""
    return all(v is None or str(v).strip() == "" for v in row_values)


def parse(file_path: str | Path) -> Dict[str, AccountLedger]:
    """解析 .xlsx 明细账文件（遍历所有 sheet，每个 sheet 是一个科目的明细账），
    返回按科目名称索引的 AccountLedger 字典。

    result["财务费用_银行手续费"] → AccountLedger(entries=[...])
    同时保留 account_code 属性。
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"明细账文件不存在：{file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    result: Dict[str, AccountLedger] = {}

    for ws in wb.worksheets:
        all_rows: list[list] = []
        for row in ws.iter_rows(values_only=True):
            all_rows.append(list(row))
        if len(all_rows) < 4:
            continue

        # 扫描所有行，找出科目区域（标题行应以「明细账」结尾）
        sections: list[dict] = []
        i = 0
        while i < len(all_rows):
            row = all_rows[i]
            title_text = str(row[0]).strip() if row[0] is not None else ""
            section_info = _parse_section_title(title_text)
            if section_info is not None:
                header_idx = i + 1
                for offset in range(1, min(4, len(all_rows) - i)):
                    check_row = all_rows[i + offset]
                    if _is_header_row(check_row):
                        header_idx = i + offset
                        break
                if header_idx >= len(all_rows):
                    i += 1
                    continue
                end_idx = len(all_rows)
                for j in range(header_idx + 1, len(all_rows)):
                    test_row = all_rows[j]
                    test_text = str(test_row[0]).strip() if test_row[0] is not None else ""
                    if _parse_section_title(test_text) is not None:
                        end_idx = j
                        break
                sections.append({
                    "title_idx": i,
                    "header_idx": header_idx,
                    "data_start": header_idx + 1,
                    "end_idx": end_idx,
                    "code": section_info[0],
                    "name": section_info[1],
                })
                i = end_idx
            else:
                i += 1

        if not sections:
            continue

        for sec in sections:
            ledger = AccountLedger(
                account_code=sec["code"],
                account_name=sec["name"],
            )

            header_values = [str(v).strip() if v is not None else ""
                             for v in all_rows[sec["header_idx"]]]

            col_date = next((i for i, h in enumerate(header_values) if "日期" in h), 0)
            col_voucher = next((i for i, h in enumerate(header_values) if "凭证" in h or "字号" in h), 1)
            col_summary = next((i for i, h in enumerate(header_values) if "摘要" in h), 2)
            col_debit = next((i for i, h in enumerate(header_values) if "借方" in h), 3)
            col_credit = next((i for i, h in enumerate(header_values) if "贷方" in h), 4)
            col_dir = next((i for i, h in enumerate(header_values) if "方向" in h), 5)
            col_balance = next((i for i, h in enumerate(header_values) if "余额" in h), 6)

            for row_idx in range(sec["data_start"], sec["end_idx"]):
                row = all_rows[row_idx]
                if _is_empty_row(row):
                    continue

                def _get(col: int):
                    if col < len(row):
                        return row[col]
                    return None

                date_val = _get(col_date)
                date_str = _norm_date(date_val)
                if not date_str:
                    continue

                voucher = str(_get(col_voucher) or "").strip()
                summary = str(_get(col_summary) or "").strip()
                # 跳过科目汇总行（非交易）：期初余额/本期合计/本年累计/承前页/过次页/上年结转
                if any(kw in summary for kw in ("期初余额", "本期合计", "本年累计",
                                                 "承前页", "过次页", "上年结转", "承上年",
                                                 "期末余额", "结转下年", "结转下期")):
                    continue
                debit = money(_get(col_debit))
                credit = money(_get(col_credit))
                direction = str(_get(col_dir) or "").strip()
                balance = money(_get(col_balance))

                ledger.entries.append(LedgerEntry(
                    date=date_str,
                    voucher_no=voucher,
                    summary=summary,
                    debit=debit,
                    credit=credit,
                    direction=direction,
                    balance=balance,
                ))

            result[sec["name"]] = ledger

    total_entries = sum(len(l.entries) for l in result.values())
    if total_entries == 0:
        raise ValueError(f"{file_path.name} 的明细账未解析出任何交易条目")

    return result


def parse_as_flat_list(file_path: str | Path) -> List[LedgerEntry]:
    """解析明细账，返回扁平化交易列表（不按科目分组）。

    每条 LedgerEntry 不包含科目信息；如需科目归属，请用 parse()。
    """
    grouped = parse(file_path)
    all_entries: List[LedgerEntry] = []
    for ledger in grouped.values():
        all_entries.extend(ledger.entries)
    return all_entries
