# -*- coding: utf-8 -*-
"""勾稽校验

任何一条不平就抛异常——宁可不出表，也不出一张看着像样、实际不平的表。
"""
from __future__ import annotations

from decimal import Decimal

from ..models import JournalEntry, PriorPeriod, ReportSet, ZERO
from ..rules import balance_sheet_rules, cash_flow_rules, income_statement_rules
from .compute import _r2, _retained_earnings_extra


class ReconciliationError(ValueError):
    """勾稽关系不成立"""


def _require(ok: bool, message: str, problems: list[str]) -> None:
    if not ok:
        problems.append(message)


def run(report: ReportSet, prior: PriorPeriod,
        entries: list[JournalEntry]) -> list[str]:
    """跑全部校验，返回通过的项目清单；有任何一条不过就抛 ReconciliationError"""
    bs = balance_sheet_rules()
    inc = income_statement_rules()
    cf = cash_flow_rules()

    problems: list[str] = []
    passed: list[str] = []

    # 1. 每笔分录借贷相等（结构上一借一贷，这里校验金额非负且非零）
    bad = [e for e in entries if e.amount <= ZERO]
    _require(not bad, f"存在金额非正的分录：{[e.summary for e in bad]}", problems)
    if not bad:
        passed.append(f"分录借贷平衡（{len(entries)} 笔）")

    # 2. 资产合计 = 负债和所有者权益合计
    assets = report.bs_end[30]
    liab_equity = report.bs_end[53]
    _require(assets == liab_equity,
             f"资产负债表不平：资产合计 {assets} ≠ 负债和所有者权益合计 {liab_equity}",
             problems)
    if assets == liab_equity:
        passed.append(f"资产负债表平衡（{assets}）")

    # 年初栏同样要平
    ys_a, ys_le = report.bs_year_start[30], report.bs_year_start[53]
    _require(ys_a == ys_le,
             f"资产负债表年初栏不平：{ys_a} ≠ {ys_le}", problems)
    if ys_a == ys_le:
        passed.append(f"资产负债表年初栏平衡（{ys_a}）")

    # 3. 未分配利润期末 = 期初 + 本月净利润 + 非结转权益性变动（分红/利润分配等）
    #    （与 compute._balance_sheet 口径一致：3103/3104 结转 posting 不进本行，
    #     但股东分红等真实权益业务必须计入，否则分红月整期报表不平）
    re_line = bs["retained_earnings_line"]
    extra_re = _retained_earnings_extra(entries, re_line)
    expected = _r2(prior.bs_end.get(re_line, ZERO)
                   + report.is_month[inc["net_profit_line"]] + extra_re)
    actual = report.bs_end[re_line]
    _require(expected == actual,
             f"未分配利润勾稽不符：期初 {prior.bs_end.get(re_line, ZERO)} + 本月净利 "
             f"{report.is_month[inc['net_profit_line']]} + 权益变动 {extra_re} = {expected}，"
             f"表上是 {actual}",
             problems)
    if expected == actual:
        passed.append(f"未分配利润勾稽（{actual}）")

    # 4. 现金流量表期末现金 = 资产负债表货币资金
    closing = report.cf_month[cf["closing_cash_line"]]
    cash_bs = report.bs_end[bs["cash_line"]]
    _require(closing == cash_bs,
             f"现金勾稽不符：现金流量表期末现金 {closing} ≠ 资产负债表货币资金 {cash_bs}",
             problems)
    if closing == cash_bs:
        passed.append(f"现金流量表与货币资金勾稽（{cash_bs}）")

    # 本年累计栏的期末现金也应一致
    closing_ytd = report.cf_ytd[cf["closing_cash_line"]]
    _require(closing_ytd == cash_bs,
             f"现金流量表本年累计期末现金 {closing_ytd} ≠ 货币资金 {cash_bs}", problems)
    if closing_ytd == cash_bs:
        passed.append("现金流量表本年累计期末现金一致")

    # 5. 季报本季度值 = 上月同季度值 + 本月值（对流量行）
    _check_quarters(report, prior, problems, passed)

    if problems:
        raise ReconciliationError(
            "报表勾稽校验未通过，已中止出表：\n  - " + "\n  - ".join(problems)
        )
    return passed


def _check_quarters(report: ReportSet, prior: PriorPeriod,
                    problems: list[str], passed: list[str]) -> None:
    quarter = report.quarter
    same_year = prior.year == report.year
    same_quarter = same_year and (prior.month - 1) // 3 + 1 == quarter

    inc_totals = {int(k) for k in income_statement_rules()["totals"]}
    cf_rules = cash_flow_rules()
    cf_totals = {int(k) for k in cf_rules["totals"]}
    opening = cf_rules["opening_cash_line"]

    mismatched: list[str] = []

    for line, cols in report.is_quarters.items():
        if line in inc_totals:
            continue
        base = prior.is_quarters.get(line, {}).get(quarter, ZERO) if same_quarter else ZERO
        if cols[quarter] != base + report.is_month[line]:
            mismatched.append(f"利润表季报第 {line} 行")

    for line, cols in report.cf_quarters.items():
        if line in cf_totals or line == opening:
            continue
        base = prior.cf_quarters.get(line, {}).get(quarter, ZERO) if same_quarter else ZERO
        if cols[quarter] != base + report.cf_month[line]:
            mismatched.append(f"现金流量表季报第 {line} 行")

    _require(not mismatched,
             f"季报与月报衔接不符：{mismatched}", problems)
    if not mismatched:
        passed.append(f"季报与月报衔接（第 {quarter} 季度）")
