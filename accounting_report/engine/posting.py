# -*- coding: utf-8 -*-
"""分录 → 科目本期发生额 + 现金流量条目

发生额一律记「自然方向」：资产/费用借方为正，负债/权益/收入贷方为正。
这样报表取数不用再判方向，直接加。
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal

from ..models import Account, JournalEntry, ZERO
from ..rules import resolve_account


@dataclass
class CashItem:
    """一条现金流量表条目（金额恒为正，行次决定流入还是流出）"""

    line: int
    amount: Decimal
    entry: JournalEntry


@dataclass
class Posting:
    movements: dict[str, Decimal] = field(default_factory=lambda: defaultdict(lambda: ZERO))
    cash_items: list[CashItem] = field(default_factory=list)

    def by_bs_line(self) -> dict[int, Decimal]:
        return self._by(lambda a: a.bs)

    def by_bs_sub_line(self) -> dict[int, Decimal]:
        return self._by(lambda a: a.bs_sub)

    def by_is_line(self) -> dict[int, Decimal]:
        return self._by(lambda a: a.is_)

    def by_is_sub_line(self) -> dict[int, Decimal]:
        return self._by(lambda a: a.is_sub)

    def _by(self, pick) -> dict[int, Decimal]:
        out: dict[int, Decimal] = defaultdict(lambda: ZERO)
        for name, delta in self.movements.items():
            line = pick(resolve_account(name))
            if line is not None:
                out[line] += delta
        return dict(out)

    def cash_by_line(self) -> dict[int, Decimal]:
        out: dict[int, Decimal] = defaultdict(lambda: ZERO)
        for item in self.cash_items:
            out[item.line] += item.amount
        return dict(out)


def _delta(account: Account, amount: Decimal, on_debit: bool) -> Decimal:
    """记在借方还是贷方 → 该科目自然方向上的增减"""
    if account.natural_debit:
        return amount if on_debit else -amount
    return -amount if on_debit else amount


def post(entries: list[JournalEntry]) -> Posting:
    result = Posting()

    for entry in entries:
        dr = resolve_account(entry.debit)
        cr = resolve_account(entry.credit)

        result.movements[dr.name] += _delta(dr, entry.amount, on_debit=True)
        result.movements[cr.name] += _delta(cr, entry.amount, on_debit=False)

        # 现金流量：一借一贷里恰好一方是现金，才产生现金流
        if dr.cash and cr.cash:
            continue  # 现金内部划转（如提现），不影响现金总额
        if not dr.cash and not cr.cash:
            continue  # 不涉及现金（如挂账费用）

        if dr.cash:
            counterpart, line = cr, (entry.cf_line or cr.cf_in)
            direction = "流入"
        else:
            counterpart, line = dr, (entry.cf_line or dr.cf_out)
            direction = "流出"

        if line is None:
            raise ValueError(
                f"科目「{counterpart.name}」没有定义现金{direction}的归类行次"
                f"（accounts.yaml 的 cf_in / cf_out）。"
                f"涉及分录：{entry.date} {entry.summary} {entry.amount}"
            )
        result.cash_items.append(CashItem(line=line, amount=entry.amount, entry=entry))

    return result
