# -*- coding: utf-8 -*-
"""发生额 + 上期数 → 三张月报

三条取数规则，按列的性质分开，不混用：
    资产负债表·期末余额  = 期初余额 + 本期发生额        （存量）
    利润表/现金流量表·本月 = 本期发生额直接汇总           （流量，与上月本月值无关）
    利润表/现金流量表·累计 = 上月累计 + 本月              （累计）

合计行一律由公式重算，不从上期继承，避免上期的舍入误差滚雪球。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from ..models import JournalEntry, PriorPeriod, ReportSet, ZERO
from ..rules import (
    balance_sheet_rules,
    cash_flow_rules,
    income_statement_rules,
    resolve_account,
)
from .posting import Posting, post

BS_LINES = 53
IS_LINES = 32
CF_LINES = 22

CENT = Decimal('0.01')


def _r2(v: Decimal) -> Decimal:
    """累计/存量列统一舍入到分，避免未舍入浮点尾巴跨期传播
    （OPT-02：现金流量表累计列期末与 BS 货币资金差 6.12 的根因）"""
    return v.quantize(CENT)


def apply_totals(values: dict[int, Decimal], rules: dict) -> dict[int, Decimal]:
    """按合计公式补齐合计行（自动处理合计依赖合计，如 资产合计 = 流动 + 非流动）"""
    out = dict(values)
    pending = {int(k): v for k, v in rules["totals"].items()}

    while pending:
        progressed = False
        for line, spec in list(pending.items()):
            deps = list(spec.get("add") or []) + list(spec.get("sub") or [])
            if any(d in pending for d in deps):
                continue  # 依赖的合计还没算出来，下一轮再说
            value = ZERO
            for d in spec.get("add") or []:
                value += out.get(d, ZERO)
            for d in spec.get("sub") or []:
                value -= out.get(d, ZERO)
            out[line] = value
            del pending[line]
            progressed = True
        if not progressed:
            raise ValueError(f"合计公式存在循环依赖：{sorted(pending)}")

    return out


def _fill(values: dict[int, Decimal], n: int) -> dict[int, Decimal]:
    return {line: values.get(line, ZERO) for line in range(1, n + 1)}


def prior_from_trial_balance(trial_balances, year: int, month: int) -> PriorPeriod | None:
    """用科目余额表的「期初余额」列构造上期期末快照。

    用途：未找到上月报表时的最后一级兜底。科目余额表本身带期初余额，
    是比报表 PDF 更可靠的期初来源（PDF 只有本期列、版式还多变）；
    不接这条兜底时，单月补跑/首月导入会退化成「期初全 0」，
    资产负债表期末只剩本期净变化，银行对账的账面余额也跟着错。

    口径与过账一致：自然方向为正（资产/费用借方为正，负债/权益/收入贷方为正）。
    父子科目去重：有下级科目的编码（如 1002，其下有 100201）整行跳过，
    否则「银行存款 + 中国工商银行」会把同一笔余额算两遍。

    期初全为 0（新设账套或表里没有期初列）时返回 None，由调用方回退到零余额。
    """
    rows = [t for t in (trial_balances or []) if getattr(t, "account_code", None)]
    if not rows:
        return None

    codes = {str(t.account_code) for t in rows}
    parents = {c for c in codes if any(o != c and o.startswith(c) for o in codes)}

    rules = balance_sheet_rules()
    subs = set(rules["sub_lines"])

    by_line: dict[int, Decimal] = defaultdict(lambda: ZERO)
    nonzero = False
    for t in rows:
        if str(t.account_code) in parents:
            continue  # 父科目由下级明细代表，避免重复累计
        try:
            acct = resolve_account(t.account_name)
        except KeyError:
            continue  # 未登记的科目不进报表，行为与过账一致
        if acct.bs is None:
            continue
        # 与 _balance_sheet 取数口径一致：明细行只认 bs_sub
        if acct.bs in subs:
            if acct.bs_sub is None:
                continue
            line = acct.bs_sub
        else:
            line = acct.bs
        balance = t.open_debit - t.open_credit
        if not acct.natural_debit:
            balance = -balance
        if balance:
            nonzero = True
        by_line[line] = _r2(by_line[line] + balance)

    if not nonzero:
        return None

    bs_end = apply_totals(_fill(by_line, BS_LINES), rules)
    py, pm = (year, month - 1) if month > 1 else (year - 1, 12)
    return PriorPeriod(year=py, month=pm, bs_end=bs_end)


def _is_carryover(entry: JournalEntry, dr, cr) -> bool:
    """判定结转/结清损益凭证（不进利润表业务发生额）

    结构判定：借贷任一方是「未分配利润」（本年利润/利润分配经别名归入），
    且另一方是损益科目（收入/费用）——只有期间损益结转才长这样。
    不用摘要关键词：真实业务摘要可能含「结转」（如"结转云服务费到应付"），
    按关键词会误排除导致利润表漏数；真结转凭证（借本年利润/贷损益类）
    无论摘要怎么写，结构上都绕不开未分配利润。
    """
    if dr.name == "未分配利润":
        return cr.type in ("revenue", "expense")
    if cr.name == "未分配利润":
        return dr.type in ("revenue", "expense")
    return False


def _income_flow(entries: list[JournalEntry]):
    """利润表·业务发生额（排除期间损益结转凭证）

    结转凭证（期末结转/年终结转）把损益科目余额清零，若按 posting 净值取数，
    利润表会全部取到 0。正确口径：只统计业务分录并按科目方向：
      借方 费用 → 增加；借方 收入 → 冲减（红字）
      贷方 收入 → 增加；贷方 费用 → 冲减（如利息收入冲减财务费用）
    """
    flow: dict[int, Decimal] = defaultdict(lambda: ZERO)
    flow_sub: dict[int, Decimal] = defaultdict(lambda: ZERO)
    for e in entries:
        try:
            dr = resolve_account(e.debit)
            cr = resolve_account(e.credit)
        except KeyError:
            continue
        if _is_carryover(e, dr, cr):
            continue
        # 借方
        if dr.type == "expense":
            flow[dr.is_] += e.amount
            if dr.is_sub:
                flow_sub[dr.is_sub] += e.amount
        elif dr.type == "revenue":
            flow[dr.is_] -= e.amount
            if dr.is_sub:
                flow_sub[dr.is_sub] -= e.amount
        # 贷方
        if cr.type == "revenue":
            flow[cr.is_] += e.amount
            if cr.is_sub:
                flow_sub[cr.is_sub] += e.amount
        elif cr.type == "expense":
            flow[cr.is_] -= e.amount
            if cr.is_sub:
                flow_sub[cr.is_sub] -= e.amount
    return flow, flow_sub


def _income_statement(prior: PriorPeriod, posting: Posting, same_year: bool,
                      entries: list[JournalEntry]):
    rules = income_statement_rules()
    subs = set(rules["sub_lines"])
    totals = {int(k) for k in rules["totals"]}

    # 利润表为流量表：本月 = 业务发生额（排除结转），不能用 posting 净值
    flow, flow_sub = _income_flow(entries)

    month: dict[int, Decimal] = {}
    for line in range(1, IS_LINES + 1):
        if line in totals:
            continue
        month[line] = flow_sub.get(line, ZERO) if line in subs else flow.get(line, ZERO)
    month = apply_totals(month, rules)

    ytd: dict[int, Decimal] = {}
    for line in range(1, IS_LINES + 1):
        if line in totals:
            continue
        base = prior.is_ytd.get(line, ZERO) if same_year else ZERO
        ytd[line] = _r2(base + month[line])
    ytd = apply_totals(ytd, rules)

    return _fill(month, IS_LINES), _fill(ytd, IS_LINES)


def _retained_earnings_extra(entries: list[JournalEntry], re_line: int) -> Decimal:
    """非结转凭证中涉及未分配利润（re_line）的净变动。

    结转凭证（借本年利润 贷损益 / 借损益 贷本年利润）的变动已隐含在净利润里，
    不能重复计入；但真实权益业务——股东分红、利润分配、以前年度损益调整——
    借/贷未分配利润会影响所有者权益总额，若不记入会导致资产减、权益不减，
    资产负债表永远不平（白班10：分红业务整期报表中止出表的根因）。
    年终结转（借本年利润 贷利润分配）两侧都是 re_line，净变动 0，天然无害。
    """
    extra = ZERO
    for e in entries:
        try:
            dr = resolve_account(e.debit)
            cr = resolve_account(e.credit)
        except KeyError:
            continue
        if _is_carryover(e, dr, cr):
            continue
        if dr.bs == re_line:
            extra -= e.amount
        if cr.bs == re_line:
            extra += e.amount
    return extra


def _balance_sheet(prior: PriorPeriod, posting: Posting, net_profit: Decimal,
                   same_year: bool, entries: list[JournalEntry]):
    rules = balance_sheet_rules()
    subs = set(rules["sub_lines"])
    totals = {int(k) for k in rules["totals"]}
    re_line = rules["retained_earnings_line"]

    mv = posting.by_bs_line()
    mv_sub = posting.by_bs_sub_line()

    end: dict[int, Decimal] = {}
    for line in range(1, BS_LINES + 1):
        if line in totals or line == re_line:
            continue
        delta = mv_sub.get(line, ZERO) if line in subs else mv.get(line, ZERO)
        end[line] = _r2(prior.bs_end.get(line, ZERO) + delta)
    # 未分配利润 = 期初 + 本月净利润（利润表口径） + 非结转权益性变动（分红等）。
    # 注意：3103/3104 在 aliases 中映射到未分配利润行，结转凭证若按 posting 净值
    # 再进本行会与净利润双重计算，故 re_line 排除 posting，仅补非结转权益业务。
    extra_re = _retained_earnings_extra(entries, re_line)
    end[re_line] = _r2(prior.bs_end.get(re_line, ZERO) + net_profit + extra_re)
    end = apply_totals(end, rules)

    # 年初余额：同一会计年度内原样延续；跨年则上年期末结转为新年度年初
    year_start: dict[int, Decimal] = {}
    for line in range(1, BS_LINES + 1):
        if line in totals:
            continue
        year_start[line] = (prior.bs_year_start.get(line, ZERO) if same_year
                            else prior.bs_end.get(line, ZERO))
    year_start = apply_totals(year_start, rules)

    return _fill(end, BS_LINES), _fill(year_start, BS_LINES)


def _cash_flow(prior: PriorPeriod, posting: Posting, same_year: bool):
    rules = cash_flow_rules()
    totals = {int(k) for k in rules["totals"]}
    opening = rules["opening_cash_line"]
    cash_line = balance_sheet_rules()["cash_line"]

    flows = posting.cash_by_line()

    month: dict[int, Decimal] = {}
    for line in range(1, CF_LINES + 1):
        if line in totals or line == opening:
            continue
        month[line] = flows.get(line, ZERO)
    # 本月期初现金 = 上月末货币资金
    month[opening] = prior.bs_end.get(cash_line, ZERO)
    month = apply_totals(month, rules)

    ytd: dict[int, Decimal] = {}
    for line in range(1, CF_LINES + 1):
        if line in totals or line == opening:
            continue
        base = prior.cf_ytd.get(line, ZERO) if same_year else ZERO
        ytd[line] = _r2(base + month[line])
    # 本年期初现金 = 年初货币资金；跨年时即上年末余额
    ytd[opening] = _r2(prior.cf_ytd.get(opening, ZERO) if same_year
                       else prior.bs_end.get(cash_line, ZERO))
    ytd = apply_totals(ytd, rules)
    # 修复：closing line 22 必须 = opening + net（都用已舍入值），
    # 不能由 apply_totals 用未舍入 month[22] 重新算（会导致累计列 vs BS 差 1.36/6.12）
    net = ytd.get(20, ZERO)
    ytd[22] = _r2(ytd[opening] + net)

    return _fill(month, CF_LINES), _fill(ytd, CF_LINES)


def build(prior: PriorPeriod, entries: list[JournalEntry],
          year: int, month_no: int) -> ReportSet:
    posting = post(entries)
    same_year = prior.year == year

    is_month, is_ytd = _income_statement(prior, posting, same_year, entries)
    net_profit = is_month[income_statement_rules()["net_profit_line"]]
    bs_end, bs_year_start = _balance_sheet(prior, posting, net_profit, same_year, entries)
    cf_month, cf_ytd = _cash_flow(prior, posting, same_year)

    return ReportSet(
        year=year, month=month_no,
        bs_end=bs_end, bs_year_start=bs_year_start,
        is_month=is_month, is_ytd=is_ytd,
        cf_month=cf_month, cf_ytd=cf_ytd,
        entries=list(entries),
    )
