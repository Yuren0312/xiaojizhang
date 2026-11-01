# -*- coding: utf-8 -*-
"""数据模型

金额一律用 Decimal，量化到分。用 float 会出现 0.1 + 0.2 = 0.30000000000000004
这种误差，报表要求元至角分、逐格零差异，容不下。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Iterable

CENT = Decimal("0.01")
ZERO = Decimal("0.00")


def money(value) -> Decimal:
    """任意输入 → 量化到分的 Decimal"""
    if value is None or value == "":
        return ZERO
    if isinstance(value, Decimal):
        d = value
    elif isinstance(value, str):
        s = value.strip().replace(",", "").replace("，", "")
        if s in ("", "-", "--"):
            return ZERO
        d = Decimal(s)
    else:
        # float 先转 str，避免把二进制误差带进来
        d = Decimal(str(value))
    return d.quantize(CENT, rounding=ROUND_HALF_UP)


def total(values: Iterable) -> Decimal:
    acc = ZERO
    for v in values:
        acc += money(v)
    return acc


@dataclass(frozen=True)
class Account:
    """科目定义"""

    name: str
    type: str  # asset | liability | equity | revenue | expense
    bs: int | None = None        # 资产负债表行次
    bs_sub: int | None = None    # 资产负债表「其中」行次
    is_: int | None = None       # 利润表行次
    is_sub: int | None = None    # 利润表「其中」行次
    cash: bool = False           # 是否现金及现金等价物
    cf_in: int | None = None     # 作为现金流入的对方科目时，归到现金流量表第几行
    cf_out: int | None = None    # 作为现金流出的对方科目时，归到第几行

    @property
    def natural_debit(self) -> bool:
        """自然余额在借方（资产、费用）"""
        return self.type in ("asset", "expense")


@dataclass
class JournalEntry:
    """一笔分录：借一科目、贷一科目"""

    date: str          # YYYY-MM-DD
    summary: str
    debit: str         # 借方科目名
    credit: str        # 贷方科目名
    amount: Decimal
    source: str = ""   # 来源单据，便于追溯
    cf_line: int | None = None  # 显式指定现金流量表行次（覆盖科目默认映射）
    # 制证/审核（导出凭证时使用；默认不写死人名，由上层接入真实成员）
    prepared_by: str = "待接入成员"
    reviewer: str = "待审核"

    def __post_init__(self):
        self.amount = money(self.amount)


@dataclass
class PriorPeriod:
    """上期报表取数结果"""

    year: int
    month: int
    # 资产负债表：行次 → 金额
    bs_end: dict[int, Decimal] = field(default_factory=dict)    # 期末余额（= 本期期初）
    bs_year_start: dict[int, Decimal] = field(default_factory=dict)  # 年初余额
    # 利润表 / 现金流量表：行次 → 本年累计
    is_ytd: dict[int, Decimal] = field(default_factory=dict)
    cf_ytd: dict[int, Decimal] = field(default_factory=dict)
    # 季报：行次 → {1..4 季度 → 金额}
    is_quarters: dict[int, dict[int, Decimal]] = field(default_factory=dict)
    cf_quarters: dict[int, dict[int, Decimal]] = field(default_factory=dict)


@dataclass
class ReportSet:
    """一期算好的五张表。键为行次，值为 Decimal。"""

    year: int
    month: int
    bs_end: dict[int, Decimal] = field(default_factory=dict)
    bs_year_start: dict[int, Decimal] = field(default_factory=dict)
    is_month: dict[int, Decimal] = field(default_factory=dict)
    is_ytd: dict[int, Decimal] = field(default_factory=dict)
    cf_month: dict[int, Decimal] = field(default_factory=dict)
    cf_ytd: dict[int, Decimal] = field(default_factory=dict)
    is_quarters: dict[int, dict[int, Decimal]] = field(default_factory=dict)
    cf_quarters: dict[int, dict[int, Decimal]] = field(default_factory=dict)
    entries: list[JournalEntry] = field(default_factory=list)

    @property
    def quarter(self) -> int:
        return (self.month - 1) // 3 + 1
