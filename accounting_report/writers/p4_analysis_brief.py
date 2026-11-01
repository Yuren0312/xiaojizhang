# -*- coding: utf-8 -*-
"""会计业务分析简报

生成一个 xlsx 工作簿，包含 4 个 sheet：
  1. 资金收支结构 — 经营活动/投资活动/筹资活动现金流拆分、占比
  2. 成本费用分析 — 月度费用趋势、费用占比数据
  3. 异常业务提示 — 自动检测异常交易
  4. 风险与建议 — 基于数据自动生成的风险提示和改进建议
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Font, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ..models import JournalEntry, ReportSet, ZERO, money
from ..parsers.icbc_statement import BankTxn

# ── 样式常量 ────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_FONT = Font(name="微软雅黑", size=10)
BOLD_FONT = Font(name="微软雅黑", bold=True, size=10)
TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
SUBTITLE_FONT = Font(name="微软雅黑", bold=True, size=11)
SECTION_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SECTION_FONT = Font(name="微软雅黑", bold=True, size=10, color="1F3864")
WARN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
ALERT_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
MONEY_FMT = '#,##0.00'
PCT_FMT = '0.00%'


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _write_row(ws, row: int, values: list, money_cols: set | None = None,
               bold: bool = False, fill: PatternFill | None = None,
               font_color: str | None = None) -> None:
    if money_cols is None:
        money_cols = set()
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        if font_color:
            cell.font = Font(name="微软雅黑", bold=bold, size=10, color=font_color)
        else:
            cell.font = BOLD_FONT if bold else CELL_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if fill:
            cell.fill = fill
        if i in money_cols and isinstance(v, (Decimal, int, float)):
            cell.number_format = MONEY_FMT


def _auto_width(ws, min_width: int = 8, max_width: int = 50) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            v = str(cell.value) if cell.value is not None else ""
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            lengths.append(length)
        best = max(lengths) if lengths else min_width
        ws.column_dimensions[col_letter].width = max(min(best + 2, max_width), min_width)


# ── Sheet 1: 资金收支结构 ───────────────────────────────────────────────

def _cash_flow_structure(ws, report_set, bank_txns: list) -> None:
    """经营活动/投资活动/筹资活动现金流拆分、占比"""

    # 安全处理 report_set 为 None
    if report_set is None or not hasattr(report_set, 'cf_month'):
        operating_in, operating_out, investing_in, investing_out, financing_in, financing_out = ZERO, ZERO, ZERO, ZERO, ZERO, ZERO
    else:
        cf = report_set.cf_month
        # 行次区间从 cash_flow.yaml totals 公式推导（7/13/19 净额行），
        # 与 analysis._classify_cash_flow 同源：改公式行次时此处自动跟随，
        # 避免 sheet 与现金流量表/分析简报再漂移（白班12 增强）
        from ..rules import cash_flow_rules
        rules = cash_flow_rules()
        totals = {int(k): v for k, v in rules["totals"].items()}
        inflow_lines = set(rules["inflow_lines"])

        def _activity(activity_total: int):
            spec = totals[activity_total]
            add = {int(d) for d in spec.get("add") or []}
            sub = {int(d) for d in spec.get("sub") or []}
            both = add | sub
            return (both & inflow_lines), (both - inflow_lines)

        operating_in_lines, operating_out_lines = _activity(7)
        investing_in_lines, investing_out_lines = _activity(13)
        financing_in_lines, financing_out_lines = _activity(19)
        operating_in = sum(cf.get(line, ZERO) for line in operating_in_lines)
        operating_out = sum(cf.get(line, ZERO) for line in operating_out_lines)
        investing_in = sum(cf.get(line, ZERO) for line in investing_in_lines)
        investing_out = sum(cf.get(line, ZERO) for line in investing_out_lines)
        financing_in = sum(cf.get(line, ZERO) for line in financing_in_lines)
        financing_out = sum(cf.get(line, ZERO) for line in financing_out_lines)

    # 如果没有现金流量表数据，从银行流水推算
    if operating_in == ZERO and operating_out == ZERO:
        operating_in, operating_out = _cash_from_txns(bank_txns,
            lambda t: "收入" in t.summary or "认证" in t.summary or "利息" in t.counterparty or "财付通" in t.counterparty)
        investing_in, investing_out = _cash_from_txns(bank_txns,
            lambda t: "理财" in t.summary or "投资" in t.summary or "结构性存款" in t.summary)
        # 筹资活动：股东借款/增资等
        financing_in, financing_out = _cash_from_txns(bank_txns,
            lambda t: "借款" in t.summary or "资本" in t.summary or "增资" in t.summary)

    total_in = operating_in + investing_in + financing_in
    total_out = operating_out + investing_out + financing_out
    net_flow = total_in - total_out

    headers = ["项目", "现金流入", "现金流出", "净流量", "流入占比", "流出占比"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {2, 3, 4}
    row = 3

    sections = [
        ("经营活动", operating_in, operating_out),
        ("投资活动", investing_in, investing_out),
        ("筹资活动", financing_in, financing_out),
    ]

    for name, inflow, outflow in sections:
        net = inflow - outflow
        in_pct = float(inflow / total_in) if total_in > ZERO else 0.0
        out_pct = float(outflow / total_out) if total_out > ZERO else 0.0
        _write_row(ws, row, [
            name,
            float(inflow),
            float(outflow),
            float(net),
            in_pct,
            out_pct,
        ], money_cols)
        row += 1

    # 合计行
    in_pct_total = 1.0 if total_in > ZERO else 0.0
    out_pct_total = 1.0 if total_out > ZERO else 0.0
    _write_row(ws, row, [
        "合计",
        float(total_in),
        float(total_out),
        float(net_flow),
        in_pct_total,
        out_pct_total,
    ], money_cols, bold=True)
    row += 1

    # 分析结论
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    conclusions = []
    if operating_in > operating_out:
        conclusions.append("经营活动现金净流入，经营造血能力正常")
    else:
        conclusions.append("经营活动现金净流出，需关注经营现金流状况")
    if investing_out > ZERO:
        conclusions.append("有投资活动支出（理财/资产购置）")
    if financing_in > ZERO:
        conclusions.append("有筹资活动现金流入")
    ws.cell(row=row, column=1, value="📊 分析结论：" + "；".join(conclusions)).font = SECTION_FONT

    _auto_width(ws)


def _cash_from_txns(txns: list, filter_fn) -> tuple[Decimal, Decimal]:
    """从银行流水中按过滤条件汇总流入流出"""
    inflow = ZERO
    outflow = ZERO
    for txn in txns:
        if filter_fn(txn):
            inflow += txn.credit
            outflow += txn.debit
    return inflow, outflow


# ── Sheet 2: 成本费用分析 ───────────────────────────────────────────────

def _cost_expense_analysis(ws, report_set,
                           voucher_entries: list,
                           trial_balances: list) -> None:
    """月度费用趋势、费用占比数据"""

    # 安全处理 report_set 为 None
    if report_set is None or not hasattr(report_set, 'is_month'):
        inc = {}
    else:
        inc = report_set.is_month
    # 利润表行次 2=营业成本, 3=税金及附加, 11=销售费用, 14=管理费用, 18=财务费用
    cost_lines = {
        "营业成本": 2,
        "税金及附加": 3,
        "销售费用": 11,
        "管理费用": 14,
        "财务费用": 18,
    }

    items = []
    total_cost = ZERO
    for name, line in cost_lines.items():
        val = inc.get(line, ZERO)
        items.append((name, val))
        if val > ZERO:
            total_cost += val  # 负数费用（利息收入冲减）不稀释费用池

    # 加所得税费用
    income_tax = inc.get(31, ZERO)
    if income_tax > ZERO:
        items.append(("所得税费用", income_tax))
        total_cost += income_tax

    # 写费用构成表
    headers = ["费用项目", "本月金额", "占总费用比例", "备注"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {2}
    row = 3

    for name, val in items:
        pct = float(val / total_cost) if total_cost > ZERO else 0.0
        note = _cost_note(name, val)
        _write_row(ws, row, [name, float(val), pct, note], money_cols)
        row += 1

    row += 1
    _write_row(ws, row, ["费用合计", float(total_cost), 1.0, ""], money_cols, bold=True)

    # 费用分析区域
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row=row, column=1, value="📊 费用结构简析").font = SUBTITLE_FONT
    row += 1

    analyses = _build_cost_analysis(items, total_cost)
    for item in analyses:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1, value=item).font = CELL_FONT
        row += 1

    _auto_width(ws)


def _cost_note(name: str, val: Decimal) -> str:
    """生成费用项的备注说明"""
    notes = {
        "营业成本": "云服务、技术服务费等主营业务成本",
        "税金及附加": "城建税、教育费附加等",
        "销售费用": "广告宣传费、市场推广费",
        "管理费用": "办公费、薪酬、折旧等",
        "财务费用": "银行手续费、利息收支净额",
        "所得税费用": "当期所得税",
    }
    return notes.get(name, "")


def _build_cost_analysis(items: list, total_cost: Decimal) -> list[str]:
    """生成费用分析建议"""
    analyses = []
    if total_cost == ZERO:
        analyses.append("本期无费用发生。")
        return analyses

    biggest = max(items, key=lambda x: x[1])
    if biggest[1] > ZERO:
        pct = float(biggest[1] / total_cost) * 100
        analyses.append(f"• {biggest[0]}占费用总额 {pct:.1f}%，为最大费用项目。")

    # 管理费用占比分析
    mgmt = next((v for n, v in items if n == "管理费用"), ZERO)
    if total_cost > ZERO and mgmt / total_cost > Decimal("0.5"):
        analyses.append("• 管理费用占比较高（>50%），建议关注费用管控。")

    # 财务费用分析
    fin = next((v for n, v in items if n == "财务费用"), ZERO)
    if fin < ZERO:
        analyses.append("• 财务费用为负（利息收入 > 手续费），反映现金管理良好。")
    elif fin > ZERO:
        analyses.append("• 存在财务费用支出，关注银行手续费水平。")

    return analyses


# ── Sheet 3: 异常业务提示 ───────────────────────────────────────────────

def _anomaly_detection(ws, bank_txns: list, voucher_entries: list) -> None:
    """自动检测异常交易"""
    headers = ["检测项", "详情", "金额", "风险等级", "建议措施"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {3}
    row = 3
    anomalies: list[dict] = []

    # 1. 小额平台认证款检测
    for txn in bank_txns:
        text = f"{txn.summary} {txn.counterparty}"
        if "认证" in text or "验证" in text or "小微" in text:
            amt = float(txn.credit) if txn.credit > ZERO else float(txn.debit)
            if amt < 1.0:
                anomalies.append({
                    "item": "小额平台认证款",
                    "detail": f"{txn.date} {txn.counterparty}: {txn.summary}",
                    "amount": amt,
                    "level": "低",
                    "suggestion": "平台认证收款，非经营性收入，金额极小，确认是否为正常商户认证流程",
                })

    # 2. 大额整数交易检测
    for txn in bank_txns:
        for amt_field in [txn.debit, txn.credit]:
            if amt_field <= ZERO:
                continue
            amt = float(amt_field)
            if amt >= 10000 and amt == int(amt):
                text = f"{txn.summary} {txn.counterparty}"
                # 排除已知的正常大额整付（如理财申购、定期存款）
                if any(kw in text for kw in ["理财", "定期", "存款"]):
                    continue
                anomalies.append({
                    "item": "大额整数交易",
                    "detail": f"{txn.date} {txn.counterparty}: {txn.summary}",
                    "amount": amt,
                    "level": "中",
                    "suggestion": "大额整数交易，建议核实业务实质并保留完整合同/审批记录",
                })

    # 3. 跨月冲正检测（同一对方、相近金额、方向相反）
    for i, txn1 in enumerate(bank_txns):
        for txn2 in bank_txns[i + 1:]:
            if txn1.counterparty != txn2.counterparty:
                continue
            # 检查方向相反 + 金额相近
            d1, c1 = float(txn1.debit), float(txn1.credit)
            d2, c2 = float(txn2.debit), float(txn2.credit)
            if d1 > 0 and c2 > 0 and abs(d1 - c2) < 0.02:
                # 跨月检查
                if txn1.date[:7] != txn2.date[:7]:
                    anomalies.append({
                        "item": "跨月冲正",
                        "detail": f"付: {txn1.date} {txn1.counterparty} {txn1.summary}, "
                                 f"收: {txn2.date} {txn2.counterparty} {txn2.summary}",
                        "amount": d1,
                        "level": "中",
                        "suggestion": "跨月回款，建议核实交易背景，确认是否为退款/冲正/测试资金往来",
                    })

    # 4. 大额理财申赎检测
    for txn in bank_txns:
        text = f"{txn.summary} {txn.counterparty}"
        if "理财" in text or "结构性存款" in text or "基金" in text:
            for amt_field, direction in [(txn.debit, "申购"), (txn.credit, "赎回")]:
                if amt_field <= ZERO:
                    continue
                amt = float(amt_field)
                if amt >= 50000:  # 5万以上
                    anomalies.append({
                        "item": "大额理财申赎",
                        "detail": f"{txn.date} {direction}: {txn.summary}",
                        "amount": amt,
                        "level": "低",
                        "suggestion": f"大额{direction}，确认已履行内部审批流程，关注资金安全和收益率",
                    })

    # 5. 股东大额往来检测
    for txn in bank_txns:
        cp = txn.counterparty
        if not cp or cp == "（空）":
            continue
        # 检测个人姓名（中文2-3字 + 不含公司/有限/股份等企业后缀）
        is_personal = (
            len(cp) <= 4 and
            all('\u4e00' <= c <= '\u9fff' for c in cp) and
            not any(kw in cp for kw in ["公司", "有限", "股份", "银行", "基金"])
        )
        if is_personal:
            for amt_field in [txn.debit, txn.credit]:
                if amt_field <= ZERO:
                    continue
                amt = float(amt_field)
                if amt >= 10000:
                    anomalies.append({
                        "item": "股东/个人大额往来",
                        "detail": f"{txn.date} 对方:{cp}: {txn.summary}",
                        "amount": amt,
                        "level": "高",
                        "suggestion": "个人账户大额往来，如为股东借款建议签订借款协议并约定利率和还款期限，防范税务风险",
                    })

    # 去重
    seen = set()
    unique_anomalies = []
    for a in anomalies:
        key = (a["item"], a["detail"], a["amount"])
        if key not in seen:
            seen.add(key)
            unique_anomalies.append(a)

    if not unique_anomalies:
        _write_row(ws, row, ["", "✓ 未检测到异常业务提示", "", "", ""], money_cols)
    else:
        # 按风险等级排序
        level_order = {"高": 0, "中": 1, "低": 2}
        unique_anomalies.sort(key=lambda a: level_order.get(a["level"], 99))

        for a in unique_anomalies:
            fill = ALERT_FILL if a["level"] == "高" else (WARN_FILL if a["level"] == "中" else None)
            font_color = "CC0000" if a["level"] == "高" else ("996600" if a["level"] == "中" else None)
            _write_row(ws, row, [
                a["item"], a["detail"],
                a["amount"] if a["amount"] else "",
                a["level"], a["suggestion"],
            ], money_cols, fill=fill, font_color=font_color)
            row += 1

    _auto_width(ws, max_width=60)


# ── Sheet 4: 风险与建议 ─────────────────────────────────────────────────

def _risk_and_advice(ws, report_set, bank_txns: list,
                     voucher_entries: list, trial_balances: list) -> None:
    """基于数据自动生成的风险提示和改进建议"""

    headers = ["类别", "风险点 / 建议", "严重程度", "参考依据"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    items: list[dict] = []
    cf = {}
    inc = {}
    bs = {}
    if report_set is not None and hasattr(report_set, 'cf_month'):
        cf = report_set.cf_month
    if report_set is not None and hasattr(report_set, 'is_month'):
        inc = report_set.is_month
    if report_set is not None and hasattr(report_set, 'bs_end'):
        bs = report_set.bs_end

    # ── 现金流风险 ──
    operating_net = cf.get(7, ZERO)  # 经营活动现金流量净额（合计行），避免与流入流出重复
    if operating_net < ZERO:
        items.append({
            "cat": "现金流风险",
            "text": f"经营活动现金净流出 {float(operating_net):,.2f} 元，经营活动无法覆盖日常支出，依赖投资回收或外部融资维持运转",
            "level": "🔴 高",
            "basis": "现金流量表-经营活动净额",
        })
    else:
        items.append({
            "cat": "现金流",
            "text": f"经营活动现金净流入 {float(operating_net):,.2f} 元，经营造血正常",
            "level": "🟢 正常",
            "basis": "现金流量表-经营活动净额",
        })

    # ── 盈利风险 ──
    revenue = inc.get(1, ZERO)
    cost = inc.get(2, ZERO)
    net_profit = inc.get(32, ZERO) if 32 in inc else sum(
        inc.get(line, ZERO) for line in [1, 20, 22]
    ) - sum(inc.get(line, ZERO) for line in [2, 3, 11, 14, 18, 24, 31])

    if revenue == ZERO:
        # 无收入 ≠ 必然的持续经营风险（白班19 修正）：投入期/回款节奏期
        # 单月无收入常见。用「本年累计收入」区分节奏月与持续无收入：
        #   - 本年累计有收入 → 仅提示关注连续性（中）；
        #   - 本年累计也无收入 → 提示投入期特征 + 现金流关注（中），
        #     不轻判「🔴高 持续经营风险」——对筹备期公司是误报、误导
        #     银行/税务等非财务读者。
        ytd_map = getattr(report_set, 'is_ytd', None) or {}
        ytd_rev = ytd_map.get(1, ZERO) if ytd_map else ZERO
        ytd_net = ytd_map.get(32, ZERO) if ytd_map else ZERO
        cash = bs.get(1, ZERO)
        if ytd_rev > ZERO:
            basis = f"利润表-本月营业收入=0、本年累计={float(ytd_rev):,.2f}"
            text = (f"本期无营业收入（本年累计 {float(ytd_rev):,.2f} 元），"
                    "收入确认存在月份节奏，属正常结算周期特征；若连续两期以上无收入，"
                    "建议核查在手订单与回款计划")
            level = "🟡 中"
        else:
            basis = f"利润表-本年累计营业收入=0（累计净利润={float(ytd_net):,.2f}）"
            text = (f"本期无营业收入，今年以来累计收入也为 0（累计净利润 "
                    f"{float(ytd_net):,.2f} 元），处于业务投入期特征明显；"
                    "建议重点管理现金流，落实收入落地计划后再评估盈利模型")
            level = "🟡 中"
        # 现金储备极低（不足 5 万）时收紧提示，避免弱化真实风险
        if cash <= ZERO:
            level = "🔴 高"
            text += "（账面货币资金已为 0，现金流告急，需立即补充）"
        items.append({
            "cat": "盈利能力",
            "text": text,
            "level": level,
            "basis": basis,
        })
    elif net_profit < ZERO:
        loss_rate = float(abs(net_profit) / revenue) * 100 if revenue > ZERO else 0
        items.append({
            "cat": "盈利能力",
            "text": f"本期净利润为负（{float(net_profit):,.2f} 元），亏损率 {loss_rate:.1f}%，需控制成本费用",
            "level": "🟡 中" if loss_rate < 30 else "🔴 高",
            "basis": f"利润表-净利润={float(net_profit):,.2f}",
        })

    # ── 资金结构风险 ──
    cash = bs.get(1, ZERO)  # 货币资金
    total_assets = bs.get(30, ZERO)  # 资产合计
    if total_assets > ZERO:
        cash_ratio = float(cash / total_assets)
        if cash_ratio > 0.8:
            items.append({
                "cat": "资金结构",
                "text": f"货币资金占资产总额 {cash_ratio*100:.1f}%，资金闲置比例较高，可考虑适当配置短期低风险理财提高收益",
                "level": "🟡 中",
                "basis": f"货币资金/资产总计={cash_ratio*100:.1f}%",
            })

    # ── 负债风险 ──
    total_liabilities = bs.get(47, ZERO) if 47 in bs else sum(
        bs.get(line, ZERO) for line in range(31, 46)
    )
    if total_assets > ZERO and total_liabilities > ZERO:
        debt_ratio = float(total_liabilities / total_assets)
        if debt_ratio > 0.6:
            items.append({
                "cat": "偿债能力",
                "text": f"资产负债率 {debt_ratio*100:.1f}%，超过60%，负债水平偏高",
                "level": "🟡 中",
                "basis": f"资产负债率={debt_ratio*100:.1f}%",
            })

    # ── 税务风险 ──
    if revenue > ZERO:
        tax = inc.get(3, ZERO)
        # 免税月（月销售额 ≤ 10 万免征增值税 → 不计提附加税）属正常，
        # 不得误报"漏报附加税"；仅当收入超门槛且有税负义务却无计提才提示
        if tax == ZERO and revenue > Decimal("100000"):
            items.append({
                "cat": "税务风险",
                "text": "营业收入超过 10 万却无税金及附加，需确认是否漏报城建税及教育费附加",
                "level": "🟡 中",
                "basis": "利润表-税金及附加为0但营业收入>10万",
            })

    # ── 费用合理性 ──
    entertainment = ZERO  # 业务招待费
    for entry in voucher_entries:
        if isinstance(entry, dict):
            dr = entry.get("debit", entry.get("借方科目", ""))
        else:
            dr = str(getattr(entry, 'debit', ''))
        if "业务招待费" in dr or "招待" in dr:
            if isinstance(entry, dict):
                entertainment -= money(entry.get("amount", entry.get("金额", ZERO)))
            else:
                entertainment -= money(getattr(entry, 'amount', ZERO))

    if entertainment < ZERO and revenue > ZERO:
        ent_ratio = float(abs(entertainment) / revenue) * 100
        if ent_ratio > 5:
            items.append({
                "cat": "费用合理性",
                "text": f"业务招待费占营业收入 {ent_ratio:.1f}%，超过 5% 警戒线，有企业所得税纳税调增风险",
                "level": "🟡 中",
                "basis": f"业务招待费/营业收入={ent_ratio:.1f}%",
            })

    # ── 理财管理建议 ──
    has_investment = any(
        "理财" in (txn.summary + txn.counterparty)
        for txn in bank_txns
    )
    if has_investment:
        items.append({
            "cat": "理财管理",
            "text": "存在短期理财操作，建议：① 建立理财台账记录每笔申赎及收益；② 关注理财产品风险等级；③ 确保闲置资金收益率不低于银行活期利率",
            "level": "🟢 建议",
            "basis": "银行流水-理财交易记录",
        })

    # ── 建议项 ──
    items.append({
        "cat": "会计核算",
        "text": "建议月末及时取得银行对账单并编制银行存款余额调节表，确保账实相符",
        "level": "🟢 建议",
        "basis": "最佳实践-银行对账",
    })

    items.append({
        "cat": "档案管理",
        "text": "建议按月份整理归档：银行回单、对账单、发票、合同等原始凭证，满足税务稽查留存要求（至少保存10年）",
        "level": "🟢 建议",
        "basis": "《会计档案管理办法》",
    })

    row = 3
    for item in items:
        fill = ALERT_FILL if "🔴" in item["level"] else (WARN_FILL if "🟡" in item["level"] else None)
        font_color = "CC0000" if "🔴" in item["level"] else ("996600" if "🟡" in item["level"] else "338833")
        _write_row(ws, row, [
            item["cat"], item["text"], item["level"], item["basis"],
        ], fill=fill, font_color=font_color)
        row += 1

    _auto_width(ws, max_width=80)


# ── Sheet 5: 本月损益结构 ─────────────────────────────────────────────

def _profit_structure(ws, report_set) -> None:
    """本月损益结构：收入=100%，成本/费用/净利润构成占比"""
    ws.merge_cells("A1:E1")
    ws.cell(row=1, column=1, value="本月损益结构（收入=100%）").font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    headers = ["项目", "本月金额", "占收入比", "说明", ""]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(3, i, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    is_m = getattr(report_set, "is_month", {}) or {}
    revenue = money(is_m.get(1, ZERO))
    cost = money(is_m.get(2, ZERO))
    admin = money(is_m.get(14, ZERO))
    finance = money(is_m.get(18, ZERO))
    net_profit = money(is_m.get(32, ZERO))

    rows = [
        ("一、营业收入", revenue, "100%", "本期收入基准"),
        ("　　减：营业成本", cost, "", ""),
        ("　　减：管理费用", admin, "", ""),
        ("　　减：财务费用", finance, "", "负数=利息净收入冲减"),
        ("二、净利润", net_profit, "", ""),
    ]
    row = 4
    for name, val, note1, note2 in rows:
        ws.cell(row, 1, name).font = CELL_FONT
        c2 = ws.cell(row, 2, float(val))
        c2.number_format = '#,##0.00'
        c2.font = CELL_FONT
        pct = ""
        if revenue > ZERO and (name.startswith("　　") or name.startswith("二、")):
            pct = f"{float(val) / float(revenue) * 100:.1f}%"
        c3 = ws.cell(row, 3, pct)
        c3.font = BOLD_FONT if name.startswith(("一、", "二、")) else CELL_FONT
        ws.cell(row, 4, note2).font = Font(name="微软雅黑", size=9, color="666666")
        row += 1

    # 点评
    row += 1
    notes = []
    if revenue > ZERO:
        net_margin = float(net_profit) / float(revenue) * 100
        cost_pct = float(cost) / float(revenue) * 100
        notes.append(f"• 净利率 {net_margin:.1f}%，营业成本占收入 {cost_pct:.1f}%")
        if net_margin < 0:
            notes.append("• 本期亏损，建议关注收入规模与成本控制")
        elif net_margin < 10:
            notes.append("• 净利率偏低（<10%），可评估提价或控费空间")
    else:
        notes.append("• 本期无营业收入（可能为费用报销期或新设企业）")
    for n in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.cell(row, 1, n).font = Font(name="微软雅黑", size=10, color="1F3864")
        row += 1

    for col, w in zip("ABCDE", (24, 16, 14, 34, 6)):
        ws.column_dimensions[col].width = w


# ── 主函数 ──────────────────────────────────────────────────────────────

def write_analysis_brief(
    report_set,           # ReportSet 从 compute.build 来的
    bank_txns: list,      # BankTxn 列表
    voucher_entries: list, # 分录列表
    trial_balances: list,  # 科目余额条目
    out_path: Path,
) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Sheet 1: 资金收支结构
    ws1 = wb.active
    ws1.title = "资金收支结构"
    ws1.cell(row=1, column=1, value="资金收支结构分析").font = TITLE_FONT
    ws1.cell(row=1, column=5, value="单位：元").font = Font(name="微软雅黑", size=9, color="666666")
    _cash_flow_structure(ws1, report_set, bank_txns)

    # Sheet 2: 成本费用分析
    ws2 = wb.create_sheet("成本费用分析")
    ws2.cell(row=1, column=1, value="成本费用分析").font = TITLE_FONT
    _cost_expense_analysis(ws2, report_set, voucher_entries, trial_balances)

    # Sheet 3: 异常业务提示
    ws3 = wb.create_sheet("异常业务提示")
    ws3.cell(row=1, column=1, value="异常业务提示").font = TITLE_FONT
    _anomaly_detection(ws3, bank_txns, voucher_entries)

    # Sheet 4: 风险与建议
    ws4 = wb.create_sheet("风险与建议")
    ws4.cell(row=1, column=1, value="风险与建议").font = TITLE_FONT
    _risk_and_advice(ws4, report_set, bank_txns, voucher_entries, trial_balances)

    # Sheet 5: 本月损益结构
    ws5 = wb.create_sheet("本月损益结构")
    _profit_structure(ws5, report_set)

    wb.save(out_path)
    return out_path
