# -*- coding: utf-8 -*-
"""科目余额表期初 → 上期期末快照（prior 兜底）回归测试

背景：run_full 单月跑（无上月报表）时，prior 只能退化成零余额，
资产负债表期末就变成「本期净变化」，银行对账的账面余额跟着错。
科目余额表本身带期初余额列，是这条链路上唯一还在手上的期初来源。

覆盖：
  - 期初非零 → bs_end 按行次落地，合计行按公式补齐
  - 父子科目去重（1002 与 100201 都是货币资金，不能算两遍）
  - 期初全 0（新设账套）→ 返回 None，由调用方回退零余额
  - 未登记科目 / 无 bs 行次的科目 → 跳过且不抛异常
"""
from __future__ import annotations

import sys
from decimal import Decimal as D
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from accounting_report.engine.compute import prior_from_trial_balance  # noqa: E402
from accounting_report.parsers.xls_trial_balance import TrialBalance  # noqa: E402
from accounting_report.rules import balance_sheet_rules  # noqa: E402

CASH_LINE = balance_sheet_rules()["cash_line"]          # 1
ASSET_TOTAL = 30
LIAB_TOTAL = 41
LE_TOTAL = 53


def _tb(code: str, name: str, open_debit=D("0"), open_credit=D("0"),
        end_debit=D("0"), end_credit=D("0")) -> TrialBalance:
    """只关心期初/期末两列的简化构造器"""
    return TrialBalance(
        account_code=code,
        account_name=name,
        open_debit=D(open_debit),
        open_credit=D(open_credit),
        period_debit=D("0"),
        period_credit=D("0"),
        ytd_debit=D("0"),
        ytd_credit=D("0"),
        end_debit=D(end_debit),
        end_credit=D(end_credit),
    )


def test_opening_balance_becomes_prior_bs_end():
    """银行存款期初 1000 → 货币资金行 = 1000，而不是 0"""
    prior = prior_from_trial_balance([_tb("1002", "银行存款", D("1000"))], 2026, 6)
    assert prior is not None
    assert prior.bs_end[CASH_LINE] == D("1000.00")
    # 上期年月要回退一期，供同一会计年度判断用
    assert (prior.year, prior.month) == (2026, 5)


def test_cross_year_rolls_back_to_december():
    prior = prior_from_trial_balance([_tb("1002", "银行存款", D("500"))], 2026, 1)
    assert prior is not None
    assert (prior.year, prior.month) == (2025, 12)


def test_parent_child_not_double_counted():
    """1002 与其下级 100201 是同一笔 1000，只能算一次"""
    rows = [
        _tb("1002", "银行存款", D("1000")),
        _tb("100201", "银行存款_中国工商银行", D("1000")),
    ]
    prior = prior_from_trial_balance(rows, 2026, 6)
    assert prior is not None
    assert prior.bs_end[CASH_LINE] == D("1000.00")


def test_asset_and_liability_totals_applied():
    """合计行由公式重算：资产合计含货币资金，负债+权益 = 资产（借贷必平）"""
    rows = [
        _tb("1002", "银行存款", D("1000")),
        _tb("应付账款", "应付账款", D("0"), D("600")),
        _tb("实收资本", "实收资本", D("0"), D("400")),
    ]
    prior = prior_from_trial_balance(rows, 2026, 6)
    assert prior is not None
    assert prior.bs_end[CASH_LINE] == D("1000.00")
    assert prior.bs_end[ASSET_TOTAL] >= D("1000.00")
    # 资产 = 负债 + 所有者权益（科目余额表本身平衡时）
    assert prior.bs_end[ASSET_TOTAL] == prior.bs_end[LE_TOTAL]


def test_all_zero_opening_returns_none():
    """期初全 0（新设账套）→ None，让调用方走零余额分支"""
    rows = [_tb("1002", "银行存款", D("0"), D("0"), end_debit=D("9999"))]
    assert prior_from_trial_balance(rows, 2026, 6) is None


def test_empty_input_returns_none():
    assert prior_from_trial_balance([], 2026, 6) is None
    assert prior_from_trial_balance(None, 2026, 6) is None


def test_unknown_account_skipped_without_raising():
    """未登记科目跳过；已登记科目的期初仍要落地"""
    rows = [
        _tb("1002", "银行存款", D("800")),
        _tb("9999", "某个还没登记的科目", D("123")),
    ]
    prior = prior_from_trial_balance(rows, 2026, 6)
    assert prior is not None
    assert prior.bs_end[CASH_LINE] == D("800.00")
