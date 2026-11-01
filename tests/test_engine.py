# -*- coding: utf-8 -*-
"""引擎核心逻辑测试（全部使用虚构数字，不依赖任何真实账套）

覆盖四条容易出错的规则：
  1. 存量列  资产负债表期末 = 期初 + 本期发生额
  2. 流量列  利润表/现金流量表本月 = 本期发生额，与上月本月值无关
  3. 累计列  本年累计 = 上月累计 + 本月
  4. 季报    当季 = 上月同季 + 本月；跨季重新起算；跨年全部清零
"""
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from accounting_report.engine import checks, compute, quarterly  # noqa: E402
from accounting_report.models import JournalEntry, PriorPeriod, money  # noqa: E402


def D(x) -> Decimal:
    return money(x)


def make_prior(year: int = 2026, month: int = 1) -> PriorPeriod:
    """一个自洽的虚构上期：资产 1000 = 负债 200 + 权益 800"""
    prior = PriorPeriod(year=year, month=month)
    prior.bs_end = {i: D(0) for i in range(1, 54)}
    prior.bs_end.update({1: D(1000), 15: D(1000), 30: D(1000),
                         33: D(200), 41: D(200), 47: D(200),
                         48: D(1000), 51: D(-200), 52: D(800), 53: D(1000)})
    prior.bs_year_start = {i: D(0) for i in range(1, 54)}
    prior.bs_year_start.update({1: D(1000), 15: D(1000), 30: D(1000),
                                48: D(1000), 52: D(1000), 53: D(1000)})

    prior.is_ytd = {i: D(0) for i in range(1, 33)}
    prior.is_ytd.update({2: D(200), 21: D(-200), 30: D(-200), 32: D(-200)})

    prior.cf_ytd = {i: D(0) for i in range(1, 23)}
    prior.cf_ytd.update({21: D(1000), 22: D(1000)})

    prior.is_quarters = {i: {q: D(0) for q in range(1, 5)} for i in range(1, 33)}
    for line, value in ((2, 200), (21, -200), (30, -200), (32, -200)):
        prior.is_quarters[line][1] = D(value)

    prior.cf_quarters = {i: {q: D(0) for q in range(1, 5)} for i in range(1, 23)}
    prior.cf_quarters[21][1] = D(1000)
    prior.cf_quarters[22][1] = D(1000)
    return prior


def make_entries(month_prefix: str = "2026-02") -> list[JournalEntry]:
    """三笔业务：收款 500、付管理费 100、挂账成本 50（第三笔不动钱）"""
    return [
        JournalEntry(f"{month_prefix}-05", "收到货款", "货币资金", "营业收入", D(500)),
        JournalEntry(f"{month_prefix}-10", "付办公费", "管理费用", "货币资金", D(100)),
        JournalEntry(f"{month_prefix}-28", "挂账成本", "营业成本", "应付账款", D(50)),
    ]


@pytest.fixture
def report():
    prior = make_prior()
    r = compute.build(prior, make_entries(), 2026, 2)
    quarterly.build(r, prior)
    return prior, r


def test_balance_sheet_is_stock(report):
    """期末 = 期初 + 发生额；未分配利润要吸收本月净利"""
    _, r = report
    assert r.bs_end[1] == D(1400)      # 货币资金 1000 + 500 - 100
    assert r.bs_end[33] == D(250)      # 应付账款 200 + 50
    assert r.bs_end[51] == D(150)      # 未分配利润 -200 + 本月净利 350
    assert r.bs_end[30] == r.bs_end[53] == D(1400)


def test_income_statement_month_is_pure_flow(report):
    """本月列只反映本月发生额，绝不掺上月的本月值"""
    _, r = report
    assert r.is_month[1] == D(500)     # 营业收入
    assert r.is_month[2] == D(50)      # 营业成本（只有本月挂账那 50，不含上期 200）
    assert r.is_month[14] == D(100)    # 管理费用
    assert r.is_month[32] == D(350)    # 净利润 500 - 50 - 100


def test_income_statement_ytd_accumulates(report):
    """本年累计 = 上月累计 + 本月"""
    _, r = report
    assert r.is_ytd[2] == D(250)       # 200 + 50
    assert r.is_ytd[32] == D(150)      # 上期 -200 + 本月 350


def test_cash_flow_classification(report):
    """现金流按对方科目归类；挂账那笔不产生现金流"""
    _, r = report
    assert r.cf_month[1] == D(500)     # 销售收现（对方是营业收入）
    assert r.cf_month[6] == D(100)     # 支付其他经营现金（对方是管理费用）
    assert r.cf_month[7] == D(400)     # 经营净额
    assert r.cf_month[21] == D(1000)   # 期初现金 = 上月末货币资金
    assert r.cf_month[22] == D(1400)   # 期末现金
    assert r.cf_month[22] == r.bs_end[1]


def test_quarters_accumulate_within_quarter(report):
    """2 月与上期 1 月同属一季度，当季列继续累加"""
    _, r = report
    assert r.is_quarters[2][1] == D(250)     # 营业成本 Q1 = 200 + 50
    assert r.is_quarters[32][1] == D(150)
    assert r.cf_quarters[1][1] == D(500)
    assert r.cf_quarters[21][1] == D(1000)   # 期初现金是存量，同季沿用不累加
    assert r.cf_quarters[22][1] == D(1400)


def test_new_quarter_restarts():
    """4 月是二季度首月，当季列从本月重新起算，不带一季度的数"""
    prior = make_prior(2026, 3)
    prior.is_quarters[2][1] = D(500)          # 一季度营业成本 500
    prior.cf_quarters[21][2] = D(0)
    r = compute.build(prior, make_entries("2026-04"), 2026, 4)
    quarterly.build(r, prior)

    assert r.is_quarters[2][1] == D(500)      # 一季度原样继承
    assert r.is_quarters[2][2] == D(50)       # 二季度只有本月的 50
    assert r.cf_quarters[21][2] == D(1000)    # 本季期初 = 本月期初，不是 0
    assert r.cf_quarters[22][2] == D(1400)


def test_new_year_resets():
    """跨年：本年累计清零、季报清零、上年期末结转为新年度年初"""
    prior = make_prior(2025, 12)
    r = compute.build(prior, make_entries("2026-01"), 2026, 1)
    quarterly.build(r, prior)

    assert r.is_ytd[2] == D(50)               # 不再含 2025 年的 200
    assert r.is_ytd[32] == D(350)
    assert r.bs_year_start[1] == D(1000)      # 新年度年初 = 上年末
    assert r.bs_year_start[33] == D(200)
    assert r.is_quarters[2][1] == D(50)
    assert r.cf_ytd[21] == D(1000)            # 年初现金 = 上年末现金


def test_checks_pass(report):
    prior, r = report
    passed = checks.run(r, prior, make_entries())
    assert len(passed) >= 6


def test_checks_catch_broken_balance(report):
    """人为破坏平衡，校验必须拦住"""
    prior, r = report
    r.bs_end[30] = D(9999)
    with pytest.raises(checks.ReconciliationError, match="资产负债表不平"):
        checks.run(r, prior, make_entries())
