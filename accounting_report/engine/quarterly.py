# -*- coding: utf-8 -*-
"""月报 → 两张季报

当前季度那一列 = 上月报表同季度值 + 本月发生额；
跨季度时（本月是 1/4/7/10 月）该列从本月重新起算，不带上季度的数。
其余季度列原样继承；跨年则全部清零。

现金流量表季报的「期初现金余额」是存量不是流量，不能累加：
本季期初 = 本季首月的期初；期末 = 期初 + 本季净增加。
"""
from __future__ import annotations

from decimal import Decimal

from ..models import PriorPeriod, ReportSet, ZERO
from ..rules import cash_flow_rules, income_statement_rules
from .compute import CF_LINES, IS_LINES, apply_totals, _r2


def _quarter_of(month: int) -> int:
    return (month - 1) // 3 + 1


def _spread(
    prior_quarters: dict[int, dict[int, Decimal]],
    month_values: dict[int, Decimal],
    n_lines: int,
    rules: dict,
    quarter: int,
    same_quarter: bool,
    same_year: bool,
    stock_lines: frozenset[int] = frozenset(),
) -> dict[int, dict[int, Decimal]]:
    totals = {int(k) for k in rules["totals"]}
    out: dict[int, dict[int, Decimal]] = {}

    for line in range(1, n_lines + 1):
        if line in totals:
            continue
        cols: dict[int, Decimal] = {}
        for q in range(1, 5):
            inherited = ZERO
            if same_year:
                inherited = prior_quarters.get(line, {}).get(q, ZERO)
            if q != quarter:
                cols[q] = inherited
            elif line in stock_lines:
                # 存量：本季期初。同季度沿用，跨季度则取本月期初
                cols[q] = inherited if same_quarter else month_values[line]
            else:
                cols[q] = _r2((inherited if same_quarter else ZERO) + month_values[line])
        out[line] = cols

    # 逐列套合计公式。结果写到新字典里——边算边往 out 里塞合计行的话，
    # 第一季度算完就会污染后面几个季度的取数。
    computed: dict[int, dict[int, Decimal]] = {}
    for q in range(1, 5):
        column = apply_totals({line: cols[q] for line, cols in out.items()}, rules)
        for line, value in column.items():
            computed.setdefault(line, {})[q] = value

    return {line: computed[line] for line in sorted(computed)}


def build(report: ReportSet, prior: PriorPeriod) -> None:
    """就地填充 report.is_quarters / report.cf_quarters"""
    quarter = report.quarter
    same_year = prior.year == report.year
    same_quarter = same_year and _quarter_of(prior.month) == quarter

    report.is_quarters = _spread(
        prior.is_quarters, report.is_month, IS_LINES,
        income_statement_rules(), quarter, same_quarter, same_year,
    )

    cf_rules = cash_flow_rules()
    report.cf_quarters = _spread(
        prior.cf_quarters, report.cf_month, CF_LINES,
        cf_rules, quarter, same_quarter, same_year,
        stock_lines=frozenset({cf_rules["opening_cash_line"]}),
    )
