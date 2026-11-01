# -*- coding: utf-8 -*-
"""会计业务分析引擎 — 输出4：分析简报

基于全量数据生成：
1. 资金收支结构（经营/投资/筹资现金流拆分）
2. 成本费用分析
3. 异常业务检测
4. 风险提示与建议
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from ..models import JournalEntry, ReportSet, money, ZERO
from ..parsers.icbc_statement import BankTxn


@dataclass
class CashFlowStructure:
    """资金收支结构"""
    operating_inflow: Decimal = ZERO
    operating_outflow: Decimal = ZERO
    investing_inflow: Decimal = ZERO
    investing_outflow: Decimal = ZERO
    financing_inflow: Decimal = ZERO
    financing_outflow: Decimal = ZERO

    @property
    def operating_net(self) -> Decimal:
        return self.operating_inflow - self.operating_outflow

    @property
    def investing_net(self) -> Decimal:
        return self.investing_inflow - self.investing_outflow

    @property
    def financing_net(self) -> Decimal:
        return self.financing_inflow - self.financing_outflow

    @property
    def total_inflow(self) -> Decimal:
        return self.operating_inflow + self.investing_inflow + self.financing_inflow

    @property
    def total_outflow(self) -> Decimal:
        return self.operating_outflow + self.investing_outflow + self.financing_outflow


@dataclass
class ExpenseBreakdown:
    """费用分类统计"""
    category: str
    amount: Decimal = ZERO
    percentage: Decimal = ZERO
    months: dict[str, Decimal] = field(default_factory=dict)


@dataclass
class Anomaly:
    """异常业务条目"""
    category: str          # 异常类别
    severity: str          # high / medium / low
    date: str = ""
    description: str = ""
    amount: Decimal = ZERO
    suggestion: str = ""


@dataclass
class AnalysisResult:
    """完整的分析简报"""
    cash_structure: CashFlowStructure = field(default_factory=CashFlowStructure)
    expenses: list[ExpenseBreakdown] = field(default_factory=list)
    anomalies: list[Anomaly] = field(default_factory=list)
    risk_items: list[str] = field(default_factory=list)
    suggestions: list[str] = field(default_factory=list)
    summary: str = ""


# 异常检测阈值
LARGE_TRANSACTION_THRESHOLD = Decimal("50000")  # 大额交易阈值 5万
SMALL_TEST_AMOUNT_MAX = Decimal("1.00")          # 小额认证款上限
ROUND_AMOUNTS = {Decimal("10000"), Decimal("100000"), Decimal("50000")}


def _classify_cash_flow(entries: list[JournalEntry]) -> CashFlowStructure:
    """按经营/投资/筹资分类现金流

    口径与输出3现金流量表完全同源：直接复用 posting.cash_by_line()
    （accounts.yaml 的 cf_in/cf_out 权威分类），按行次区间归入三大活动。
    此前用关键词匹配会与现金流量表漂移——股东分红（cf_out:18 筹资流出）
    被误归经营流出、现金内部划转被误计等（白班11 修复）。

    行次区间不再硬编码，而是从 cash_flow.yaml 的 totals 公式推导
    （7=经营净额、13=投资净额、19=筹资净额，add/sub 各自的行次即该活动
    的流入/流出集合），与现金流量表永久同源——改公式行次时分类自动跟随，
    不会出现两表再漂移（白班12 增强）。
    """
    from ..engine.posting import post
    from ..rules import cash_flow_rules

    result = CashFlowStructure()
    lines = post(entries).cash_by_line()
    rules = cash_flow_rules()
    totals = {int(k): v for k, v in rules["totals"].items()}
    inflow_lines = set(rules["inflow_lines"])

    def _activity(activity_total: int):
        spec = totals[activity_total]
        add = {int(d) for d in spec.get("add") or []}
        sub = {int(d) for d in spec.get("sub") or []}
        both = add | sub
        return (both & inflow_lines), (both - inflow_lines)

    operating_in, operating_out = _activity(7)
    investing_in, investing_out = _activity(13)
    financing_in, financing_out = _activity(19)

    for line, amount in lines.items():
        if line in operating_in:
            result.operating_inflow += amount
        elif line in operating_out:
            result.operating_outflow += amount
        elif line in investing_in:
            result.investing_inflow += amount
        elif line in investing_out:
            result.investing_outflow += amount
        elif line in financing_in:
            result.financing_inflow += amount
        elif line in financing_out:
            result.financing_outflow += amount
        elif line == rules["net_increase_line"] or line == rules["closing_cash_line"]:
            continue  # 汇总/期末行不计入三大活动

    return result


def _analyze_expenses(entries: list[JournalEntry]) -> list[ExpenseBreakdown]:
    """分析费用结构"""
    by_category: dict[str, ExpenseBreakdown] = {}

    expense_keywords = {
        "营业成本": "营业成本/云服务费",
        "办公费": "办公费",
        "业务招待": "业务招待费",
        "福利费": "福利费",
        "服务费": "其他服务费",
        "银行手续费": "银行手续费",
        "利息": "利息净支出",
    }

    for e in entries:
        matched = False
        for kw, cat in expense_keywords.items():
            if kw in e.debit or kw in e.credit:
                if cat not in by_category:
                    by_category[cat] = ExpenseBreakdown(category=cat)
                by_category[cat].amount += e.amount
                month_key = e.date[:7]
                by_category[cat].months[month_key] = (
                    by_category[cat].months.get(month_key, ZERO) + e.amount
                )
                matched = True
                break

    # 计算占比
    total_expense = sum(br.amount for br in by_category.values())
    if total_expense > ZERO:
        for br in by_category.values():
            br.percentage = (br.amount / total_expense * Decimal("100")).quantize(
                Decimal("0.01"))

    return sorted(by_category.values(), key=lambda x: float(x.amount), reverse=True)


def _detect_anomalies(
    bank_txns: list[BankTxn],
    entries: list[JournalEntry],
    trial_balances: list | None = None,
) -> list[Anomaly]:
    """检测异常业务"""
    anomalies: list[Anomaly] = []

    # 1. 小额平台认证款（<1元的入账）
    for txn in bank_txns:
        if ZERO < txn.inflow <= SMALL_TEST_AMOUNT_MAX and "认证" in txn.summary:
            anomalies.append(Anomaly(
                category="小额平台认证款",
                severity="low",
                date=txn.date,
                description=f"{txn.summary}（对方：{txn.counterparty}）",
                amount=txn.inflow,
                suggestion="属于第三方平台商户认证测试打款，金额微小，不影响报表。建议确认是否已完成认证。"
            ))

    # 2. 大额理财交易
    for e in entries:
        if "短期投资" in (e.debit, e.credit) and e.amount >= LARGE_TRANSACTION_THRESHOLD:
            direction = "购入" if e.debit == "短期投资" else "赎回"
            anomalies.append(Anomaly(
                category="大额理财交易",
                severity="medium" if e.amount < Decimal("100000") else "high",
                date=e.date,
                description=f"{direction}理财 {e.amount}元：{e.summary}",
                amount=e.amount,
                suggestion="请确认理财产品的风险等级与公司风险承受能力匹配。大额申赎需关注资金流动性。"
            ))

    # 3. 大额整数交易（可能是往来款/借款）
    for e in entries:
        if e.amount >= Decimal("10000") and e.amount % Decimal("10000") == ZERO:
            if any(kw in (e.debit + e.credit + e.summary) for kw in ["借款", "往来", "股东"]):
                anomalies.append(Anomaly(
                    category="大额整数往来款",
                    severity="medium",
                    date=e.date,
                    description=f"{e.summary} {e.amount}元",
                    amount=e.amount,
                    suggestion="大额整数往来款需确认是否有借款协议，避免税务风险。"
                ))

    # 4. 股东/个人往来（通用关键词；不硬编码具体客户姓名，保证引擎可复用）
    for e in entries:
        if any(kw in (e.debit + e.credit + e.summary) for kw in ["股东", "个人"]):
            if e.amount >= Decimal("10000"):
                anomalies.append(Anomaly(
                    category="股东/个人大额往来",
                    severity="high",
                    date=e.date,
                    description=f"{e.summary}：{e.debit}/{e.credit} {e.amount}元",
                    amount=e.amount,
                    suggestion="股东借款需在年底前归还，否则可能被视为分红征收个人所得税。"
                ))

    # 5. 跨月冲正检测（同一摘要、相近金额、不同月份）
    by_summary: dict[str, list[JournalEntry]] = defaultdict(list)
    for e in entries:
        key = e.summary.strip()
        if key:
            by_summary[key].append(e)

    for summary, group in by_summary.items():
        if len(group) >= 2:
            months = {g.date[:7] for g in group}
            if len(months) >= 2:
                amounts = [g.amount for g in group]
                # 检查是否有金额相近但方向相反
                for i in range(len(group)):
                    for j in range(i + 1, len(group)):
                        if abs(float(group[i].amount) - float(group[j].amount)) < 0.02:
                            if (group[i].debit == group[j].credit and
                                group[i].credit == group[j].debit):
                                anomalies.append(Anomaly(
                                    category="跨月冲正",
                                    severity="medium",
                                    date=f"{group[i].date} / {group[j].date}",
                                    description=f"{summary}（{group[i].amount}元）",
                                    amount=group[i].amount,
                                    suggestion="跨月冲正可能影响月度报表准确性，请确认业务合理性。"
                                ))

    return anomalies


def _generate_risks(anomalies: list[Anomaly],
                     cash: CashFlowStructure,
                     entries: list[JournalEntry],
                     report: ReportSet | None = None) -> tuple[list[str], list[str]]:
    """生成风险提示与建议"""
    risks: list[str] = []
    suggestions: list[str] = []

    # 现金流风险
    if cash.operating_net < ZERO:
        risks.append(f"经营活动现金净流出 {abs(cash.operating_net)} 元，主营业务尚未产生正向现金流。")
        suggestions.append("建议加快业务拓展，增加营业收入，改善经营现金流。")

    if cash.financing_inflow > ZERO and cash.operating_inflow == ZERO:
        risks.append("公司运营资金主要依赖外部筹资，缺乏自主经营收入。")
        suggestions.append("建议制定明确的业务收入计划，逐步减少对外部资金的依赖。")

    # 理财风险
    investing_total = cash.investing_inflow + cash.investing_outflow
    if investing_total > LARGE_TRANSACTION_THRESHOLD * 2:
        risks.append(f"投资活动资金流动较大（总额 {investing_total} 元），需关注投资风险。")
        suggestions.append("建议定期评估理财产品收益与风险，确保资金安全性和流动性。")

    # 异常风险
    high_severity = [a for a in anomalies if a.severity == "high"]
    if high_severity:
        risks.append(f"存在 {len(high_severity)} 项高风险异常，请重点核查。")

    # 费用结构
    expense_entries = [e for e in entries if any(
        kw in e.debit for kw in ["营业成本", "管理费用", "财务费用", "销售费用"]
    )]
    total_expense = sum((e.amount for e in expense_entries), ZERO)
    cost_entries = [e for e in entries if "营业成本" in e.debit]
    total_cost = sum((e.amount for e in cost_entries), ZERO)
    if total_expense > ZERO and total_cost > ZERO:
        cost_ratio = float(total_cost / total_expense * Decimal("100"))
        if cost_ratio > 90:
            risks.append(f"营业成本占总费用 {cost_ratio:.1f}%，费用结构较单一。")

    # 往来款风险
    payable_entries = [e for e in entries if any(
        kw in (e.debit + e.credit) for kw in ["其他应付款", "应付账款"]
    )]
    total_payable = sum((e.amount for e in payable_entries), ZERO)
    if total_payable > Decimal("50000"):
        risks.append(f"应付/其他应付挂账金额较大（{total_payable} 元），需关注偿债能力。")
        suggestions.append("建议制定付款计划，及时结算应付账款，维护供应商关系。")

    # 亏损风险：直接用利润表净利润行（is_month 口径），
    # 不能靠「本年利润」科目关键词——业务分录从不直接挂该科目，
    # 靠关键词统计在纯费用月（亏损）必然漏报（白班11 修复）
    net_profit = ZERO
    if report is not None:
        from ..rules import income_statement_rules
        net_profit = report.is_month.get(income_statement_rules()["net_profit_line"], ZERO)
    if net_profit < ZERO:
        risks.append(f"本期净利润为负（{net_profit} 元），企业处于亏损状态。")
        suggestions.append("建议进行成本费用分析，优化支出结构，尽快实现盈亏平衡。")

    if not risks:
        risks.append("未发现重大财务风险。")

    suggestions.insert(0, "建议按月编制并审阅银行余额调节表，确保账实相符。")
    suggestions.append("建议定期（每季度）进行全面的财务分析，及时发现问题并调整经营策略。")

    return risks, suggestions


def analyze(
    bank_txns: list[BankTxn],
    entries: list[JournalEntry],
    report: ReportSet | None = None,
    trial_balances: list | None = None,
) -> AnalysisResult:
    """执行全面分析"""
    result = AnalysisResult()

    # 现金流结构
    result.cash_structure = _classify_cash_flow(entries)

    # 费用分析
    result.expenses = _analyze_expenses(entries)

    # 异常检测
    result.anomalies = _detect_anomalies(bank_txns, entries, trial_balances)

    # 风险与建议（传入 report 供净利润口径判断亏损，避免关键词法漏报）
    result.risk_items, result.suggestions = _generate_risks(
        result.anomalies, result.cash_structure, entries, report=report
    )

    # 生成摘要
    cs = result.cash_structure
    result.summary = (
        f"企业资金总流入 {cs.total_inflow} 元，总流出 {cs.total_outflow} 元。\n"
        f"其中：经营活动净现金流 {cs.operating_net} 元，"
        f"投资活动净现金流 {cs.investing_net} 元，"
        f"筹资活动净现金流 {cs.financing_net} 元。\n"
        f"共识别 {len(result.anomalies)} 项异常业务（高 {sum(1 for a in result.anomalies if a.severity == 'high')} 项、"
        f"中 {sum(1 for a in result.anomalies if a.severity == 'medium')} 项、"
        f"低 {sum(1 for a in result.anomalies if a.severity == 'low')} 项）。"
    )

    return result
