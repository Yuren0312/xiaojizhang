# -*- coding: utf-8 -*-
"""银行流水 → 分录

对账单是「银行视角」：贷方发生额 = 钱进企业账户，借方发生额 = 钱出。
转成企业账面分录时方向要反过来：
    进账 → 借 货币资金 / 贷 对方科目
    出账 → 借 对方科目 / 贷 货币资金

对方科目由 rules/mapping.yaml 的关键词规则决定。匹配不上不猜，直接报错并列清单。
"""
from __future__ import annotations

from ..models import JournalEntry, ZERO
from ..parsers.icbc_statement import BankTxn
from ..rules import bank_mapping


class UnmappedTransaction(ValueError):
    """有流水匹配不到科目"""


def _match_account(txn: BankTxn) -> tuple[str, str] | None:
    """返回 (科目, 命中的关键词)"""
    haystack = f"{txn.summary} {txn.counterparty}"
    for rule in bank_mapping()["bank_rules"]:
        if rule["match"] in haystack:
            return rule["account"], rule["match"]
    return None


def from_bank(txns: list[BankTxn]) -> list[JournalEntry]:
    cash = bank_mapping()["cash_account"]
    entries: list[JournalEntry] = []
    unmapped: list[BankTxn] = []

    for txn in txns:
        hit = _match_account(txn)
        if hit is None:
            unmapped.append(txn)
            continue
        account, _ = hit

        if txn.inflow > ZERO:
            entries.append(JournalEntry(
                date=txn.date, summary=txn.summary,
                debit=cash, credit=account, amount=txn.inflow,
                source=txn.source,
            ))
        if txn.outflow > ZERO:
            entries.append(JournalEntry(
                date=txn.date, summary=txn.summary,
                debit=account, credit=cash, amount=txn.outflow,
                source=txn.source,
            ))

    if unmapped:
        detail = "\n  ".join(
            f"{t.date}  {t.summary}  对方:{t.counterparty}  "
            f"借{t.debit} 贷{t.credit}"
            for t in unmapped
        )
        raise UnmappedTransaction(
            "以下银行流水无法归类到科目，请在 accounting_report/rules/mapping.yaml "
            f"补充 bank_rules 后重跑：\n  {detail}"
        )

    return entries
