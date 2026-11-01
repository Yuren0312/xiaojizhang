# -*- coding: utf-8 -*-
"""银行对账报告生成器

输入：银行流水（icbc_statement.BankTxn）+ 凭证分录 + 明细账
输出：包含对账明细/未达账项/余额调节表/费用统计/往来汇总的 xlsx
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from ..models import JournalEntry, money, ZERO
from ..parsers.icbc_statement import BankTxn


@dataclass
class ReconItem:
    """一笔对账条目"""
    date: str
    summary: str
    counterparty: str
    bank_debit: Decimal = ZERO    # 银行借方（银行视角支出）
    bank_credit: Decimal = ZERO   # 银行贷方（银行视角收入）
    book_debit: Decimal = ZERO    # 账面借方（企业借方）
    book_credit: Decimal = ZERO   # 账面贷方（企业贷方）
    matched: bool = False
    note: str = ""


@dataclass
class BankReconciliation:
    """银行对账完整结果"""
    items: list[ReconItem] = field(default_factory=list)
    unreached_bank: list[ReconItem] = field(default_factory=list)  # 银行已记企业未记
    unreached_book: list[ReconItem] = field(default_factory=list)  # 企业已记银行未记
    bank_balance: Decimal = ZERO     # 银行对账单余额
    book_balance: Decimal = ZERO     # 企业账面余额
    adjusted_balance: Decimal = ZERO  # 调节后余额


def build_reconciliation(
    bank_txns: list[BankTxn],
    entries: list[JournalEntry],
    ledger_bank_txns: list[dict] | None = None,
) -> BankReconciliation:
    """构建银行对账数据

    bank_txns: 银行对账单解析出的流水
    entries: 记账分录（从中筛选涉及货币资金的）
    ledger_bank_txns: 明细账中银行存款科目的明细（可选）
    """
    result = BankReconciliation()

    # 提取账面银行存款变动
    cash_entries = [e for e in entries if "货币资金" in (e.debit, e.credit)
                    or "银行存款" in (e.debit, e.credit)]

    used_bank: set[int] = set()
    used_book: set[int] = set()

    for ei, entry in enumerate(cash_entries):
        key = f"{entry.date}|{float(entry.amount):.2f}"
        # 找同金额同日期的银行流水
        candidates = [
            (bi, t) for bi, t in enumerate(bank_txns)
            if t.date == entry.date and abs(float(t.inflow + t.outflow) - float(entry.amount)) < 0.02
        ]
        matched = False
        for bi, txn in candidates:
            if bi not in used_bank:
                used_bank.add(bi)
                used_book.add(ei)
                item = ReconItem(
                    date=entry.date, summary=entry.summary,
                    counterparty=getattr(txn, 'counterparty', ''),
                    bank_debit=txn.debit, bank_credit=txn.credit,
                    book_debit=entry.amount if entry.debit in ("货币资金","银行存款") else ZERO,
                    book_credit=entry.amount if entry.credit in ("货币资金","银行存款") else ZERO,
                    matched=True,
                )
                result.items.append(item)
                matched = True
                break
        if not matched:
            used_book.add(ei)
            item = ReconItem(
                date=entry.date, summary=entry.summary,
                counterparty="",
                book_debit=entry.amount if entry.debit in ("货币资金","银行存款") else ZERO,
                book_credit=entry.amount if entry.credit in ("货币资金","银行存款") else ZERO,
                matched=False, note="企业已记，银行未记"
            )
            result.items.append(item)
            # 企业已记、银行未记 → 企业侧未达（unreached_book），
            # 不能塞进 unreached_bank（银行已记、企业未记），否则调节表加错侧
            result.unreached_book.append(item)

    # 未匹配的银行流水
    for bi, txn in enumerate(bank_txns):
        if bi not in used_bank:
            item = ReconItem(
                date=txn.date, summary=txn.summary,
                counterparty=txn.counterparty,
                bank_debit=txn.debit, bank_credit=txn.credit,
                matched=False, note="银行已记，企业未记"
            )
            result.items.append(item)
            result.unreached_bank.append(item)

    # 计算余额
    if bank_txns:
        result.bank_balance = bank_txns[-1].balance
    result.book_balance = sum(
        (e.amount for e in entries if e.debit in ("货币资金","银行存款")), ZERO
    ) - sum(
        (e.amount for e in entries if e.credit in ("货币资金","银行存款")), ZERO
    )

    # 调节后余额 = 银行余额 + 企业已记银行未记收入 - 企业已记银行未记支出
    result.adjusted_balance = result.bank_balance
    for item in result.unreached_bank:
        result.adjusted_balance += item.book_debit - item.book_credit

    return result
