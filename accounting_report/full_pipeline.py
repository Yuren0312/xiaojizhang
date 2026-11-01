# -*- coding: utf-8 -*-
"""全流程管线：输入全部原始资料 → 输出 1/2/3/4 四类会计成果

Query 层：解析银行流水、凭证、明细账、科目余额表
WorkBody 层：分录生成、报表计算、对账、分析
输出层：写入四个 xlsx 文件
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from . import config
from .engine import analysis, checks, compute, journalize, quarterly, reconciliation
from .models import JournalEntry, PriorPeriod, ReportSet, ZERO, money
from .parsers import (
    business_events,
    icbc_receipt,
    icbc_statement,
    prior_report_pdf,
    prior_report_xlsx,
    xls_trial_balance,
    xls_voucher,
    xlsx_ledger,
)
from .pipeline import Inputs, Result as ReportResult, collect_entries, cross_check_receipts, discover, load_prior
from .writers import xlsx_writer


# 尝试导入新 writer（可能尚在构建中）
try:
    from .writers import p1_bank_reconciliation
except ImportError:
    p1_bank_reconciliation = None

try:
    from .writers import p2_account_summary
except ImportError:
    p2_account_summary = None

try:
    from .writers import p4_analysis_brief
except ImportError:
    p4_analysis_brief = None

try:
    from .writers import p5_vat_declaration
except ImportError:
    p5_vat_declaration = None

try:
    from .writers import p6_income_tax_declaration
except ImportError:
    p6_income_tax_declaration = None

try:
    from .writers import p7_payroll_tax
except ImportError:
    p7_payroll_tax = None


@dataclass
class QueryResult:
    """Query 层数据清洗后的标准化数据集"""
    bank_txns: list = field(default_factory=list)
    voucher_entries: list = field(default_factory=list)
    trial_balances: list = field(default_factory=list)
    ledger_bank_txns: list = field(default_factory=list)
    ledger_all: list = field(default_factory=list)
    business_events: list[JournalEntry] = field(default_factory=list)
    journal_entries: list[JournalEntry] = field(default_factory=list)
    prior: PriorReport | None = None
    report: ReportSet | None = None
    warnings: list[str] = field(default_factory=list)


@dataclass
class FullResult:
    """完整的 1/2/3/4 输出结果"""
    query: QueryResult = field(default_factory=QueryResult)
    report_result: ReportResult | None = None
    reconciliation_data: reconciliation.BankReconciliation | None = None
    analysis_data: analysis.AnalysisResult | None = None
    # 输出文件路径
    output3_report: Path | None = None       # 标准报表
    output1_reconciliation: Path | None = None  # 银行对账
    output2_summary: Path | None = None       # 核算汇总
    output4_analysis: Path | None = None      # 分析简报
    output5_vat: Path | None = None           # 增值税申报底稿
    output6_income_tax: Path | None = None    # 所得税预缴底稿
    output7_payroll_tax: Path | None = None   # 个税扣缴明细
    output8_vouchers: Path | None = None      # 记账凭证
    output9_trial_balance: Path | None = None  # 试算平衡表
    prior: PriorPeriod | None = None          # 本期使用的上期数据（月链串联用）
    report: ReportSet | None = None           # 本期报表
    errors: list[str] = field(default_factory=list)


def run_query(
    year: int, month: int,
    input_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
) -> QueryResult:
    """Query 模块：解析所有原始资料，清洗为标准数据集"""
    input_dir = Path(input_dir) if input_dir else config.INPUT_DIR
    output_dir = Path(output_dir) if output_dir else config.OUTPUT_DIR
    month_prefix = f"{year:04d}-{month:02d}"
    result = QueryResult()

    # 1. 解析银行对账单 PDF（多版式兼容，失败不阻断）
    #    递归搜索子目录（输入/银行对账单/、输入/业务回单/ 等）
    bank_parsed = False
    for pdf in sorted(input_dir.rglob("*.pdf")):
        try:
            kind = _classify_pdf(pdf)
            if kind == "对账单":
                txns = icbc_statement.parse(pdf)
                result.bank_txns.extend(txns)
                bank_parsed = True
        except Exception as e:
            result.warnings.append(f"解析对账单 {pdf.name} 失败（版式可能已变，将用凭证 fallback）：{e}")

    # 1b. 按月过滤银行流水（仅保留当月交易），并跨文件去重
    #     （输入目录常有多份同名/跨期重复 PDF，同一笔交易会被解析多次）
    result.bank_txns = [
        t for t in result.bank_txns
        if str(getattr(t, 'date', ''))[:7] == month_prefix[:7]
    ]
    from .pipeline import _dedup_bank_txns
    result.bank_txns = _dedup_bank_txns(result.bank_txns)

    # 2. 解析凭证 xls（递归搜索子目录）
    voucher_files = list(input_dir.rglob("*凭证*.xls")) + list(input_dir.rglob("*凭证*.xlsx"))
    for f in voucher_files:
        try:
            entries = xls_voucher.parse(f)
            result.voucher_entries.extend(entries)
        except Exception as e:
            result.warnings.append(f"解析凭证 {f.name} 失败：{e}")

    # 2b. 银行流水 fallback：从凭证构造 BankTxn
    if not bank_parsed and result.voucher_entries:
        try:
            result.bank_txns = _build_bank_txns_from_vouchers(
                result.voucher_entries, month_prefix
            )
            if result.bank_txns:
                result.warnings.append(
                    f"已从凭证构造 {len(result.bank_txns)} 笔银行流水作为 fallback"
                    "（余额从 0 起算，仅供参考；请补充银行对账单以获得准确余额）"
                )
        except Exception as e:
            result.warnings.append(f"从凭证构造银行流水失败：{e}")

    # 3. 解析科目余额表 xls（递归搜索子目录）
    tb_files = list(input_dir.rglob("*科目余额*.xls")) + list(input_dir.rglob("*科目余额*.xlsx"))
    for f in tb_files:
        try:
            balances = xls_trial_balance.parse(f)
            result.trial_balances.extend(balances)
        except Exception as e:
            result.warnings.append(f"解析科目余额表 {f.name} 失败：{e}")

    # 4. 解析明细账 xlsx（递归搜索子目录）
    #     ⚠ 明细账导出常为多月累计文件（如「2025年11账期至2026年6账期 明细账」），
    #     与银行流水/凭证一致地按账期过滤：核算汇总表的理财/往来/费用明细只应反映
    #     当月交易，否则每期报告都带着全年的未来日期流水（白班20 发现：
    #     1月报告出现 2-6 月交易、7 期明细完全相同）。
    ledger_files = list(input_dir.rglob("*明细账*.xlsx")) + list(input_dir.rglob("*明细账*.xls"))
    for f in ledger_files:
        try:
            grouped = xlsx_ledger.parse(f)
            all_entries = []
            for acct_name, ledger in grouped.items():
                for e in ledger.entries:
                    # 无日期/日期异常的明细行保留并告警由解析层负责；这里仅按月裁剪
                    if str(getattr(e, "date", "") or "")[:7] != month_prefix:
                        continue
                    d = {
                        "account_code": ledger.account_code,
                        "account_name": acct_name,
                        "date": e.date,
                        "voucher_no": e.voucher_no,
                        "summary": e.summary,
                        "debit": e.debit,
                        "credit": e.credit,
                        "direction": e.direction,
                        "balance": e.balance,
                    }
                    all_entries.append(d)
                    result.ledger_all.append(d)
            result.ledger_bank_txns = [
                d for d in all_entries
                if "银行存款" in str(d.get("account_name", ""))
            ]
        except Exception as e:
            result.warnings.append(f"解析明细账 {f.name} 失败：{e}")

    # 5. 解析业务事项
    events_file = input_dir / "业务事项.xlsx"
    if events_file.exists():
        try:
            result.business_events = business_events.parse_xlsx(events_file)
        except Exception as e:
            result.warnings.append(f"解析业务事项失败：{e}")

    # 6. 生成分录：按凭证字号分组，借贷自然配对
    # 财务软件导出的凭证每一行要么在借方要么在贷方，同组内借=贷。
    entries: list[JournalEntry] = []
    from collections import defaultdict
    groups: dict[str, list] = defaultdict(list)
    for ve in result.voucher_entries:
        ve_date = str(getattr(ve, 'date', ''))[:10]
        if ve_date.startswith(month_prefix):
            # 凭证字号为空的行：按 (日期, 摘要) 归组，保持「无字号凭证」
            # 多行借贷在同一张凭证内（若按行独立成组，借贷会被拆散，
            # 单侧为空直接整张丢弃）。字号非空时优先按字号分组。
            vno = str(getattr(ve, 'voucher_no', '') or '').strip()
            if vno:
                key = f"{ve_date}#{vno}"
            else:
                summary = str(getattr(ve, 'summary', '') or '').strip()
                key = f"{ve_date}#__nosum__{summary}" if summary else f"{ve_date}#__empty__"
            groups[key].append(ve)

    for vn, group in groups.items():
        ve_date = str(getattr(group[0], 'date', ''))[:10]
        ve_summary = str(getattr(group[0], 'summary', ''))
        debit_lines = []   # (account, amount)
        credit_lines = []  # (account, amount)
        for ve in group:
            acct = str(getattr(ve, 'account_name', '') or getattr(ve, 'account_code', '')).strip()
            d = money(getattr(ve, 'debit', 0))
            c = money(getattr(ve, 'credit', 0))
            # 财务软件的利息收入等科目可能用借方负数表示贷方
            if d > ZERO:
                debit_lines.append((acct, d))
            elif d < -ZERO:
                credit_lines.append((acct, abs(d)))
            if c > ZERO:
                credit_lines.append((acct, c))
            elif c < -ZERO:
                debit_lines.append((acct, abs(c)))

        # 组级借贷平衡校验：借贷总额不等（借≠贷）的凭证，一律不入账并告警。
        # 必须在配对前校验——否则「借100贷90」会被一对一分支当作合法凭证，
        # 按借方金额 100 生成一借一贷分录，报表凭空多 10 元。
        dr_total = sum(a for _, a in debit_lines)
        cr_total = sum(a for _, a in credit_lines)
        if abs(dr_total - cr_total) >= money('0.01'):
            result.warnings.append(
                f"凭证「{vn}」({ve_date}) 借贷不平衡（借 {dr_total} ≠ 贷 {cr_total}），"
                f"该张凭证未入账，请核查源数据"
            )
            continue
        if not debit_lines or not credit_lines:
            result.warnings.append(
                f"凭证「{vn}」({ve_date}) 单侧为空（借 {len(debit_lines)} 行 / "
                f"贷 {len(credit_lines)} 行），未入账，请核查源数据"
            )
            continue

        # 一对一：直接配对
        if len(debit_lines) == 1 and len(credit_lines) == 1:
            entries.append(JournalEntry(
                date=ve_date, summary=ve_summary,
                debit=debit_lines[0][0], credit=credit_lines[0][0],
                amount=debit_lines[0][1],
                source=f"凭证:{vn}"
            ))
        # 一借多贷：借方科目 × 每笔贷方
        elif len(debit_lines) == 1 and len(credit_lines) > 1:
            for cr_acct, cr_amt in credit_lines:
                entries.append(JournalEntry(
                    date=ve_date, summary=ve_summary,
                    debit=debit_lines[0][0], credit=cr_acct,
                    amount=cr_amt,
                    source=f"凭证:{vn}"
                ))
        # 多借一贷：每笔借方 × 贷方科目
        elif len(debit_lines) > 1 and len(credit_lines) == 1:
            for dr_acct, dr_amt in debit_lines:
                entries.append(JournalEntry(
                    date=ve_date, summary=ve_summary,
                    debit=dr_acct, credit=credit_lines[0][0],
                    amount=dr_amt,
                    source=f"凭证:{vn}"
                ))
        # 多借多贷：按金额降序配对（期末结转等场景）
        else:
            # Collect all debits and credits by account
            dr_by_acct = defaultdict(lambda: ZERO)
            cr_by_acct = defaultdict(lambda: ZERO)
            for acct, amt in debit_lines:
                dr_by_acct[acct] += amt
            for acct, amt in credit_lines:
                cr_by_acct[acct] += amt
            # Remove self-matching (same account on both sides = internal transfer)
            common = set(dr_by_acct.keys()) & set(cr_by_acct.keys())
            for acct in common:
                d = dr_by_acct[acct]
                c = cr_by_acct[acct]
                if d == c:
                    del dr_by_acct[acct]
                    del cr_by_acct[acct]
                elif d > c:
                    dr_by_acct[acct] = d - c
                    del cr_by_acct[acct]
                else:
                    cr_by_acct[acct] = c - d
                    del dr_by_acct[acct]
            # Pair remaining: each debit account against each credit account
            # 使用简单策略：第一笔借方配第一笔贷方，以此类推
            dr_items = list(dr_by_acct.items())
            cr_items = list(cr_by_acct.items())
            dr_total = sum(a for _, a in dr_items)
            cr_total = sum(a for _, a in cr_items)
            if dr_items and cr_items and abs(dr_total - cr_total) < money('0.01'):
                # 按金额排序后配对
                dr_sorted = sorted(dr_items, key=lambda x: x[1], reverse=True)
                cr_sorted = sorted(cr_items, key=lambda x: x[1], reverse=True)
                di, ci = 0, 0
                dr_rem = dr_sorted[0][1]
                cr_rem = cr_sorted[0][1]
                while di < len(dr_sorted) and ci < len(cr_sorted):
                    matched = min(dr_rem, cr_rem)
                    if matched > ZERO:
                        entries.append(JournalEntry(
                            date=ve_date, summary=ve_summary,
                            debit=dr_sorted[di][0], credit=cr_sorted[ci][0],
                            amount=matched,
                            source=f"凭证:{vn}"
                        ))
                    dr_rem -= matched
                    cr_rem -= matched
                    if dr_rem < money('0.01'):
                        di += 1
                        if di < len(dr_sorted):
                            dr_rem = dr_sorted[di][1]
                    if cr_rem < money('0.01'):
                        ci += 1
                        if ci < len(cr_sorted):
                            cr_rem = cr_sorted[ci][1]
            else:
                # 配对失败：借贷总额不等或单侧为空——必须暴露给用户，不能静默丢整张凭证
                result.warnings.append(
                    f"凭证「{vn}」({ve_date}) 借贷不平衡（借 {dr_total} ≠ 贷 {cr_total}），"
                    f"该张凭证未入账，请核查源数据"
                )

    events_check = input_dir / "业务事项.xlsx"
    if events_check.exists():
        entries.extend(
            e for e in result.business_events if e.date.startswith(month_prefix)
        )

    # 6c. 前端「补充说明」文本提醒：web-ui 会把无票收入/现金交易说明等落到
    #     输入/{period}/_{module}_补充说明.txt，但引擎只认结构化「业务事项.xlsx」
    #     入账，不解析自由文本。必须显式提示，避免用户以为填了说明就会入账
    #     却拿到未含该业务的报表（文案承诺与实际行为脱节的体验坑）。
    note_files = sorted(
        input_dir.glob(f"{month_prefix}/*补充说明*.txt")
    ) + sorted(
        input_dir.glob(f"*{month_prefix}*补充说明*.txt")
    )
    if not note_files:
        # 兼容补丁：_补充说明 可能直接放在账期目录（upload-text 落盘位置）
        note_files = sorted(
            (input_dir / month_prefix).glob("_*_补充说明.txt")
        )
    if note_files:
        names = "、".join(f.name for f in note_files[:5])
        result.warnings.append(
            f"检测到补充说明文本（{names}）：当前引擎仅按结构化「业务事项.xlsx」"
            "生成分录，自由文本说明不会自动入账。如需入账，请将业务写入"
            "「输入/业务事项.xlsx」（日期/摘要/借方科目/贷方科目/金额）。"
        )

    result.journal_entries = entries

    # 6d. 当月凭证无数据但银行有流水 → 从银行流水构造分录
    if not result.journal_entries and result.bank_txns:
        try:
            entries_from_bank = _build_journal_entries_from_bank(
                result.bank_txns, month_prefix, result.warnings
            )
            if entries_from_bank:
                result.journal_entries = entries_from_bank
                result.warnings.append(
                    f"当月凭证无数据，已从银行流水构造 {len(entries_from_bank)} 笔分录"
                )
        except Exception as e:
            result.warnings.append(f"从银行流水构造分录失败：{e}")

    return result


def _build_bank_txns_from_vouchers(
    voucher_entries: list, month_prefix: str,
    opening_balance: Decimal | None = None,
) -> list:
    """从凭证条目中提取银行存款科目变动，构造 BankTxn 列表。

    凭证视角 vs 银行视角的方向映射：
    - 凭证借方=银行存款增加 → 银行贷方（credit，账户收入）
    - 凭证贷方=银行存款减少 → 银行借方（debit，账户支出）

    opening_balance：期初余额。从凭证只能推出本月发生额，余额列需要上月期末；
    不传时从 0 起算（余额列仅供参考，调用方应提示）。
    """
    txns: list = []
    running_balance = money(opening_balance) if opening_balance is not None else ZERO
    for ve in voucher_entries:
        ve_date = str(getattr(ve, 'date', ''))[:10]
        if not ve_date.startswith(month_prefix):
            continue
        acct_name = str(getattr(ve, 'account_name', '') or getattr(ve, 'account_code', ''))
        if "银行存款" not in acct_name:
            continue
        d = money(getattr(ve, 'debit', 0))
        c = money(getattr(ve, 'credit', 0))
        # 凭证借方（银行存款增加）→ 银行贷方
        # 凭证贷方（银行存款减少）→ 银行借方
        if d > ZERO:
            running_balance += d
            bank_debit = ZERO
            bank_credit = d
        elif c > ZERO:
            running_balance -= c
            bank_debit = c
            bank_credit = ZERO
        else:
            continue
        txns.append(icbc_statement.BankTxn(
            date=ve_date,
            summary=str(getattr(ve, 'summary', '')),
            counterparty="",
            debit=bank_debit,
            credit=bank_credit,
            balance=running_balance,
            source="凭证构造",
        ))
    return txns


def _build_journal_entries_from_bank(
    bank_txns: list, month_prefix: str,
    warnings: list | None = None,
) -> list[JournalEntry]:
    """当月凭证无数据时，从银行流水构造 JournalEntry。

    银行借方 = 账户支出 → 借对方科目 贷银行存款
    银行贷方 = 账户收入 → 借银行存款 贷对方科目
    摘要走 bank_rules 映射，匹配不上归营业外支出(支出)/营业外收入(收入)并发出警告。
    """
    from .rules import bank_mapping

    mapping_rules = bank_mapping()["bank_rules"]
    cash = bank_mapping()["cash_account"]

    def _match(summary: str, counterparty: str) -> str | None:
        haystack = f"{summary} {counterparty}"
        for rule in mapping_rules:
            if rule["match"] in haystack:
                return rule["account"]
        return None

    def _warn(msg: str):
        if warnings is not None:
            warnings.append(msg)

    entries: list[JournalEntry] = []
    unmatched_count = 0
    for t in bank_txns:
        t_date = str(getattr(t, 'date', ''))[:10]
        if not t_date.startswith(month_prefix):
            continue
        summary = str(getattr(t, 'summary', '')).strip()
        counterparty = str(getattr(t, 'counterparty', '')).strip()
        # BankTxn: debit=支出, credit=收入
        inflow = money(getattr(t, 'credit', 0))
        outflow = money(getattr(t, 'debit', 0))

        account = _match(summary, counterparty)
        if account is None:
            unmatched_count += 1
            # 支出归入"营业外支出"，收入归入"营业外收入"
            if outflow > ZERO:
                account = "营业外支出"
            else:
                account = "营业外收入"
            _warn(
                f'银行流水未匹配到科目（归入"{account}"）：'
                f'{t_date} {summary} {counterparty} '
                f"{'支出' if outflow > ZERO else '收入'}"
                f'{outflow if outflow > ZERO else inflow}——'
                f'请在 rules/mapping.yaml 添加关键词映射'
            )

        if outflow > ZERO:
            entries.append(JournalEntry(
                date=t_date, summary=summary,
                debit=account, credit=cash, amount=outflow,
                source="银行流水",
            ))
        if inflow > ZERO:
            entries.append(JournalEntry(
                date=t_date, summary=summary,
                debit=cash, credit=account, amount=inflow,
                source="银行流水",
            ))
    if unmatched_count > 0:
        _warn(f'共有 {unmatched_count} 笔银行流水未匹配到科目，已归入营业外收支')
    return entries


def accrue_taxes(entries: list[JournalEntry], year: int, month: int,
                 vat_rate: str = "0.01") -> list[JournalEntry]:
    """按本期营业收入计提附加税（小规模纳税人口径，与输出5 申报底稿一致）

    - 免税门槛：月销售额 ≤ 100,000 免征增值税 → 不产生附加税计提（2023-2027 政策）
    - 增值税估算 = 不含税收入 × 征收率（默认 1%：3% 减按 1% 优惠，2023-2027 延续）
    - 附加税 = 增值税 × 12% × 50%（六税两费减半：城建税 7% + 教育费附加 3% + 地方教育附加 2%）
    - 无收入或免征则返回空列表（不产生空计提凭证）
    返回新生成的计提凭证（借:税金及附加 / 贷:应交税费-各税种，一借一贷多行）。
    """
    if month < 1 or month > 12:
        return []
    prefix = f"{year:04d}-{month:02d}"
    from .rules import resolve_account  # noqa: PLC0415 - 局部导入避免循环依赖
    revenue = ZERO
    for e in entries:
        # 科目名做别名归一（"主营业务收入"→"营业收入"）：业务事项手填别名时
        # 也计入收入，与利润表取数口径一致，避免超 10 万却不计提附加税。
        try:
            credit_name = resolve_account(e.credit).name
        except KeyError:
            credit_name = e.credit
        # 借方冲减收入（销售退回/折让红字）也按净额口径扣除：与利润表
        # _income_flow「借方收入冲减、贷方收入增加」一致，避免月销 15 万 +
        # 退货 6 万（净 9 万 ≤ 10 万免征）时仍按 15 万计提附加税。
        if e.date.startswith(prefix):
            if credit_name == "营业收入":
                revenue += e.amount
            else:
                try:
                    debit_name = resolve_account(e.debit).name
                except KeyError:
                    debit_name = e.debit
                if debit_name == "营业收入":
                    revenue -= e.amount
    if revenue <= ZERO:
        return []
    # 月销售额 ≤ 10 万免征增值税 → 无附加税（与输出5 的 TAX_FREE_THRESHOLD 口径一致）
    if revenue <= Decimal("100000"):
        return []
    rate = money(vat_rate)
    vat = (revenue * rate).quantize(money("0.01"))
    surcharge = (vat * money("0.12") * money("0.5")).quantize(money("0.01"))  # 六税两费减半
    if surcharge <= ZERO:
        return []
    # 月末最后一天
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    date = f"{year:04d}-{month:02d}-{last_day:02d}"
    city = (surcharge * money("7") / money("12")).quantize(money("0.01"))
    edu = (surcharge * money("3") / money("12")).quantize(money("0.01"))
    local = (surcharge - city - edu).quantize(money("0.01"))
    return [
        JournalEntry(date=date, summary="计提附加税（自动）", debit="税金及附加",
                     credit="应交税费-城市维护建设税", amount=city, source="税费计提-自动"),
        JournalEntry(date=date, summary="计提附加税（自动）", debit="税金及附加",
                     credit="应交税费-教育费附加", amount=edu, source="税费计提-自动"),
        JournalEntry(date=date, summary="计提附加税（自动）", debit="税金及附加",
                     credit="应交税费-地方教育附加", amount=local, source="税费计提-自动"),
    ]


def _tb_snapshot_end_totals(trial_balances: list) -> tuple:
    """科目余额表快照期末合计（父子科目去重，与 p2 科目余额汇总同口径）→ (借, 贷)

    编码为其他编码严格前缀的行（如 100201 之于 1002）是汇总行的重复明细，
    仅保留父行避免重复累加；返回去重后的期末借方/贷方合计。
    """
    codes = [str(getattr(t, "account_code", "") or "").strip()
             for t in trial_balances or []]
    keep = [t for t, c in zip(trial_balances or [], codes)
            if not any(c.startswith(o) and len(c) > len(o)
                       for o in codes if o and o != c)]
    end_d = sum(money(getattr(t, "end_debit", ZERO)) for t in keep)
    end_c = sum(money(getattr(t, "end_credit", ZERO)) for t in keep)
    return end_d, end_c


def run_workbody(
    query: QueryResult,
    year: int, month: int,
    input_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
    prior: PriorPeriod | None = None,
) -> tuple[ReportSet | None, PriorPeriod | None]:
    """WorkBody 模块：执行会计计算

    prior 参数：外部传入的上期报表（月链运行时由调用方串联传入），
    优先于 discover/load_prior 自动查找——避免输出目录旧文件污染 prior 链
    （OPT-02：2025-12 期初被误填 99,368.42 导致报表翻倍的根因）。
    """
    input_dir = Path(input_dir) if input_dir else config.INPUT_DIR
    output_dir = Path(output_dir) if output_dir else config.OUTPUT_DIR

    # 外部传入的 prior 优先；否则加载上期报表
    if prior is None:
        try:
            found = discover(input_dir, output_dir, year, month)
            prior = load_prior(found)
        except (FileNotFoundError, ValueError) as e:
            query.warnings.append(f"discover 未找到上月报表：{e}")

    if prior is None:
        prior = _find_prior_from_report_pdfs(year, month, query.warnings)

    if prior is None:
        py, pm = (year, month - 1) if month > 1 else (year - 1, 12)
        prior = PriorPeriod(year=py, month=pm)
        query.warnings.append(
            "未找到上月报表，使用零余额作为期初。"
            "首月报表期初余额可能不准确。"
        )

    # 无分录但找到了上月报表 → 标记但不阻断
    if not query.journal_entries:
        query.warnings.append("当月无分录数据，报表将仅反映上期余额。")
        return None, prior

    # 科目合法性预检：未定义科目会在过账（posting.post）时抛 KeyError，
    # 整期报表失败且错误信息不友好。这里先收集完整清单并给出明确指引，
    # 避免用户看到半截堆栈无从下手（如业务事项手填错别字、新科目未登记）。
    from .rules import resolve_account as _resolve
    unknown_accounts: list[str] = []
    for e in query.journal_entries:
        for side in (e.debit, e.credit):
            try:
                _resolve(side)
            except KeyError:
                if side not in unknown_accounts:
                    unknown_accounts.append(side)
    if unknown_accounts:
        raise ValueError(
            "以下科目未在科目表中定义，无法过账："
            f"{'、'.join(unknown_accounts)}。"
            "请检查业务事项/凭证是否写错科目名，"
            "或在 accounting_report/rules/accounts.yaml 中补充科目定义后重跑。"
        )

    # 税金计提：本期有营业收入时，自动计提附加税（借税金及附加 / 贷应交税费）
    try:
        accruals = accrue_taxes(query.journal_entries, year, month)
        if accruals:
            total_surcharge = sum(e.amount for e in accruals)
            query.journal_entries.extend(accruals)
            query.warnings.append(
                f"已自动计提附加税 {money(total_surcharge):.2f} 元"
                "（增值税估算 1% × 附加 12% × 减半，小规模口径，与输出5一致；"
                "月销≤10万免征不计提，可在 mapping.yaml/accrue_taxes 调整）"
            )
    except Exception as e:  # noqa: BLE001 - 计提失败不应阻断出表
        query.warnings.append(f"税金计提跳过（{e}）")

    # 计算报表
    report = compute.build(prior, query.journal_entries, year, month)
    quarterly.build(report, prior)

    # 科目余额表快照账期一致性提示：科目余额表常为「截至某月末」快照文件，
    # 若用户拿 6 月快照跑 1-5 月账期（或多月目录重复使用），核算汇总表第1页
    # 会与本期报表口径不一致。此处把快照期末总额（按 p2 第1页同口径父子去重）
    # 与本期报表对比，不一致时明确告警，不静默出一张口径打架的汇总表（白班20）。
    if query.trial_balances:
        try:
            from decimal import Decimal as _Dec
            tb_end_d, tb_end_c = _tb_snapshot_end_totals(query.trial_balances)
            bs_assets = _Dec(money(report.bs_end.get(30, ZERO)))
            bs_le = _Dec(money(report.bs_end.get(53, ZERO)))
            if abs(tb_end_d - bs_assets) > _Dec("0.02") or abs(tb_end_c - bs_le) > _Dec("0.02"):
                query.warnings.append(
                    f"⚠️ 科目余额表期末合计（借 {money(tb_end_d):,.2f} / "
                    f"贷 {money(tb_end_c):,.2f}）与本期报表资产合计 "
                    f"{money(bs_assets):,.2f} / 负债和所有者权益合计 "
                    f"{money(bs_le):,.2f} 不一致——科目余额表快照可能不是 "
                    f"{year:04d}-{month:02d} 账期（核算汇总表第1页按输入文件原样转述，"
                    f"请核对输入文件账期，或上传当月科目余额表后重跑）。"
                )
        except Exception:  # noqa: BLE001 - 一致性提示失败不阻断出表
            pass

    return report, prior


def run_chain(
    periods: list[tuple[int, int]],
    input_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
    entity_type: str = "corp",
    initial_prior: PriorPeriod | None = None,
) -> list[FullResult]:

    """连续多期运行：每期直接用上期 build 输出构造下期 prior，
    不走 discover 自动加载（避免输出目录旧文件污染 prior 链）。

    首期（无 prior 时）回退到 run_workbody 自动发现；
    若自动发现也失败则用零余额期初。

    initial_prior：可选，强制首期使用指定 prior（如零余额期初），
    彻底绕开 discover 对输入目录展示报表 PDF 的误识别。
    """
    from .models import PriorPeriod as PP

    results: list[FullResult] = []
    # 与 run_full 对齐：None 目录回退到 config 默认（修复不传目录时 _emit_outputs 崩 NoneType）
    input_dir = Path(input_dir) if input_dir else config.INPUT_DIR
    output_dir = Path(output_dir) if output_dir else config.OUTPUT_DIR
    prior: PriorPeriod | None = initial_prior
    for year, month in periods:
        query = run_query(year, month, input_dir, output_dir)
        result = FullResult()
        result.query = query
        try:
            report, new_prior = run_workbody(
                query, year, month, input_dir, output_dir, prior=prior
            )
        except Exception as e:
            result.errors.append(f"报表计算失败：{e}")
            report, new_prior = None, None

        if report is None and new_prior is None:
            result.errors.append("报表计算失败且未找到上月报表。")
            results.append(result)
            continue

        result.report = report
        result.prior = new_prior
        # 用本月 build 结果构造下月 prior（关键：保持链不污染）
        next_prior: PriorPeriod | None = prior
        if report is not None:
            next_prior = PP(
                year=year, month=month,
                bs_end=report.bs_end,
                bs_year_start=report.bs_year_start,
                is_ytd=report.is_ytd,
                cf_ytd=report.cf_ytd,
                is_quarters=report.is_quarters,
                cf_quarters=report.cf_quarters,
            )
        prior = next_prior

        # 本期输出必须用本期 prior（new_prior），不能用已更新的下期 prior：
        # 否则 output9 试算平衡表期初 = 本期期末（双重累计）、勾稽 checks 用错期初，
        # 导致校验恒失败（OPT：run_chain prior 链污染）
        _emit_outputs(result, query, report, new_prior, year, month,
                     input_dir, output_dir, entity_type)
        results.append(result)
    return results


def _cleanup_stale_outputs(output_dir: Path, year: int, month: int, label_kw: str) -> None:
    """清理同账期的旧版输出，保证每类输出只保留最新一份。

    场景：同一账期先跑过模板版（缺发票/工资/资产数据时文件名带 [模板]），
    之后补齐数据再跑生成正式版 → 模板版残留，输出目录同类型文件成对出现，
    结果页列表重复、用户不知该下载哪个。生成新文件前把同账期
    「{prefix}*{label_kw}*.xlsx」旧文件（含 [模板] 版）一并删除。
    仅清理引擎输出，绝不触碰输入目录。
    """
    prefix = f"{year:04d}-{month:02d}"
    for f in output_dir.glob(f"{prefix}*{label_kw}*.xlsx"):
        try:
            f.unlink()
        except OSError:
            pass


def _emit_outputs(
    result: FullResult, query: QueryResult,
    report: ReportSet | None, prior: PriorPeriod | None,
    year: int, month: int,
    input_dir: Path, output_dir: Path, entity_type: str,
    progress=None,
) -> None:
    """输出 1-8（从 run_full 抽出的公共部分，供 run_chain 复用）"""
    def _report(pct: int, msg: str) -> None:
        if progress is not None:
            try:
                progress(pct, msg)
            except Exception:
                pass

    from .writers import p1_bank_reconciliation, p2_account_summary
    from .writers import p4_analysis_brief, p5_vat_declaration
    from .writers import p6_income_tax_declaration, p7_payroll_tax
    from .engine import analysis

    if report is None and prior is None:
        result.errors.append("报表计算失败且未找到上月报表。")
        return

    # === 输出3：标准财务报表（有分录才出） ===
    if report is not None and query.journal_entries:
        out3 = output_dir / f"{year:04d}-{month:02d} 会计报表.xlsx"
        try:
            from .engine import checks as check_mod
            passed = check_mod.run(report, prior, query.journal_entries)
        except check_mod.ReconciliationError as e:
            # 勾稽不过：必须把原因暴露给用户，不能静默吞掉出"看似正常"的表
            query.warnings.append(f"⚠️ 勾稽校验未通过：{e}")
            passed = []
        except Exception as e:  # noqa: BLE001
            query.warnings.append(f"勾稽校验异常：{e}")
            passed = []

        try:
            from .writers import xlsx_writer
            xlsx_writer.write(report, out3)
            result.output3_report = out3
        except Exception as e:
            result.errors.append(f"输出标准报表失败：{e}")

    # === 输出1：银行对账报告（无条件生成） ===
    if p1_bank_reconciliation:
        try:
            out1 = output_dir / f"{year:04d}-{month:02d} 银行对账报告.xlsx"
            book_bank_entries = [
                je for je in query.journal_entries
                if '银行存款' in str(je.debit) or '银行存款' in str(je.credit)
                or '货币资金' in str(je.debit) or '货币资金' in str(je.credit)
            ]
            # 账面余额：直接用本期报表货币资金期末余额（已与明细账验证一致），
            # 不可从分录累加（多借多贷会重复计数）
            opening_cash = report.bs_end.get(1) if report is not None else (
                prior.bs_end.get(1) if prior is not None else None
            )
            p1_bank_reconciliation.write_bank_reconciliation(
                query.bank_txns, book_bank_entries, out1, opening_cash=opening_cash
            )
            result.output1_reconciliation = out1
        except Exception as e:
            result.errors.append(f"输出银行对账报告失败：{e}")

    # === 输出2：全账务核算汇总表 ===
    if p2_account_summary:
        _report(66, "正在生成核算汇总表...")
        try:
            out2 = output_dir / f"{year:04d}-{month:02d} 核算汇总表.xlsx"
            p2_account_summary.write_account_summary(
                query.trial_balances, query.ledger_all, out2,
                period_end=f"{year:04d}-{month:02d}",
            )
            result.output2_summary = out2
        except Exception as e:
            result.errors.append(f"输出核算汇总表失败：{e}")

    # === 输出4：会计业务分析简报（需要报表） ===
    if p4_analysis_brief and report is not None:
        _report(74, "正在生成分析简报...")
        try:
            analysis_data = analysis.analyze(
                query.bank_txns, query.journal_entries, report, query.trial_balances
            )
            result.analysis_data = analysis_data
            out4 = output_dir / f"{year:04d}-{month:02d} 分析简报.xlsx"
            p4_analysis_brief.write_analysis_brief(
                report, query.bank_txns, query.voucher_entries,
                query.trial_balances, out4
            )
            result.output4_analysis = out4
        except Exception as e:
            result.errors.append(f"输出分析简报失败：{e}")

    # === 输出5：增值税申报底稿（需要报表数据） ===
    if p5_vat_declaration and report is not None:
        _report(80, "正在生成增值税申报底稿...")
        try:
            _cleanup_stale_outputs(output_dir, year, month, "增值税")
            invoice_data = _try_load_invoice_ledger(input_dir, year, month)
            out5, vat_is_template = p5_vat_declaration.generate(
                report_set=report,
                year=year, month=month,
                invoice_data=invoice_data,
                out_path=output_dir,
            )
            result.output5_vat = out5
            if vat_is_template:
                query.warnings.append(
                    "输出5: 增值税申报底稿为模板输出（缺少发票台账数据源）"
                )
        except Exception as e:
            result.errors.append(f"输出增值税申报底稿失败：{e}")

    # === 输出6：所得税预缴底稿（需要报表数据） ===
    if p6_income_tax_declaration and report is not None:
        _report(86, "正在生成企业所得税预缴底稿...")
        try:
            asset_total, employee_count = _try_load_asset_employee(input_dir)
            _cleanup_stale_outputs(output_dir, year, month, "预缴底稿")
            out6, inc_tax_is_template = p6_income_tax_declaration.generate(
                report_set=report,
                year=year, month=month,
                entity_type=entity_type,
                asset_total=asset_total,
                employee_count=employee_count,
                out_path=output_dir,
            )
            result.output6_income_tax = out6
            if inc_tax_is_template:
                query.warnings.append(
                    "输出6: 所得税预缴底稿为模板输出（缺少资产台账/员工名册数据源）"
                )
        except Exception as e:
            result.errors.append(f"输出所得税预缴底稿失败：{e}")

    # === 输出7：个税扣缴明细表（独立于报表） ===
    if p7_payroll_tax:
        _report(92, "正在生成个税扣缴明细表...")
        try:
            employees = _try_load_payroll(input_dir, year, month)
            _cleanup_stale_outputs(output_dir, year, month, "个税")
            out7, pr_tax_is_template = p7_payroll_tax.generate(
                year=year, month=month,
                employees=employees,
                out_path=output_dir,
            )
            result.output7_payroll_tax = out7
            if pr_tax_is_template:
                query.warnings.append(
                    "输出7: 个税扣缴明细表为模板输出（缺少员工工资数据源）"
                )
        except Exception as e:
            result.errors.append(f"输出个税扣缴明细表失败：{e}")

    # === 输出8：记账凭证（从分录生成，产品闭环补齐） ===
    _report(96, "正在生成记账凭证...")
    if query.journal_entries:
        try:
            from .writers import p8_vouchers
            out8 = p8_vouchers.generate(
                query.journal_entries, year, month, out_path=output_dir
            )
            result.output8_vouchers = out8
        except Exception as e:
            result.errors.append(f"输出记账凭证失败：{e}")

    # === 输出9：试算平衡表（会计核心文件，从分录汇总生成） ===
    # 门控用「有分录」而非「有科目余额表」：借贷发生额直接来自分录，
    # 科目余额表仅补充科目编码/名称。只传银行对账单+业务事项时也应出表。
    if query.journal_entries:
        try:
            from .writers import p9_trial_balance
            out9 = p9_trial_balance.generate(
                query.journal_entries, query.trial_balances, prior, year, month,
                out_path=output_dir,
            )
            result.output9_trial_balance = out9
        except Exception as e:
            result.errors.append(f"输出试算平衡表失败：{e}")

    _report(100, "全部输出完成")
    return


def run_full(
    year: int, month: int,
    input_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
    entity_type: str = "corp",  # "corp" | "individual"
    progress=None,  # Callable[[int, str], None]：进度回调（0-100 + 阶段文案）
) -> FullResult:
    """运行完整管线：Query → WorkBody → 输出 1/2/3/4"""
    input_dir = Path(input_dir) if input_dir else config.INPUT_DIR
    output_dir = Path(output_dir) if output_dir else config.OUTPUT_DIR
    result = FullResult()

    def _report(pct: int, msg: str) -> None:
        if progress is not None:
            try:
                progress(pct, msg)
            except Exception:
                pass

    # === Query ===
    _report(5, "正在解析输入资料（对账单/凭证/明细账/发票）...")
    query = run_query(year, month, input_dir, output_dir)
    result.query = query
    _report(25, "正在生成会计分录...")

    # === WorkBody ===
    report, prior = run_workbody(query, year, month, input_dir, output_dir)
    _report(45, "正在过账并编制报表...")

    if report is None and prior is None:
        result.errors.append("报表计算失败且未找到上月报表。")
        return result

    # 关键：把本期待算出的报表/上期数据挂到 FullResult。
    # （此前漏赋值，result.report 恒为 None——输出文件虽正常生成，
    #  但调用方拿不到报表对象，如 API 的 analysis_data 回退与链跑拼接。
    #  run_chain 已赋值，run_full 对齐。白班11 修复）
    result.report = report
    result.prior = prior

    # === 输出3：标准财务报表（有分录才出） ===
    if report is not None and query.journal_entries:
        _report(50, "正在校验勾稽关系并输出标准报表...")
        out3 = output_dir / f"{year:04d}-{month:02d} 会计报表.xlsx"
        try:
            from .engine import checks as check_mod
            passed = check_mod.run(report, prior, query.journal_entries)
        except check_mod.ReconciliationError as e:
            # 勾稽不过：暴露原因到 warnings，不让用户拿到一张"看似正常"的不平表
            query.warnings.append(f"⚠️ 勾稽校验未通过：{e}")
            passed = []
        except Exception as e:  # noqa: BLE001
            query.warnings.append(f"勾稽校验异常：{e}")
            passed = []

        try:
            xlsx_writer.write(report, out3)
            result.output3_report = out3
        except Exception as e:
            result.errors.append(f"输出标准报表失败：{e}")

    # === 输出1：银行对账报告（无条件生成） ===
    if p1_bank_reconciliation:
        try:
            _report(58, "正在生成银行对账报告...")
            out1 = output_dir / f"{year:04d}-{month:02d} 银行对账报告.xlsx"
            # 从分录中提取银行科目的交易作为账面数据
            book_bank_entries = [
                je for je in query.journal_entries
                if '银行存款' in str(je.debit) or '银行存款' in str(je.credit)
                or '货币资金' in str(je.debit) or '货币资金' in str(je.credit)
            ]
            p1_bank_reconciliation.write_bank_reconciliation(
                query.bank_txns, book_bank_entries, out1,
                opening_cash=report.bs_end.get(1) if report is not None else (
                    prior.bs_end.get(1) if prior else None)
            )
            result.output1_reconciliation = out1
        except Exception as e:
            result.errors.append(f"输出银行对账报告失败：{e}")

    # === 输出2：全账务核算汇总表 ===
    if p2_account_summary:
        _report(66, "正在生成核算汇总表...")
        try:
            out2 = output_dir / f"{year:04d}-{month:02d} 核算汇总表.xlsx"
            p2_account_summary.write_account_summary(
                query.trial_balances, query.ledger_all, out2,
                period_end=f"{year:04d}-{month:02d}",
            )
            result.output2_summary = out2
        except Exception as e:
            result.errors.append(f"输出核算汇总表失败：{e}")

    # === 输出4：会计业务分析简报（需要报表） ===
    if p4_analysis_brief and report is not None:
        _report(74, "正在生成分析简报...")
        try:
            analysis_data = analysis.analyze(
                query.bank_txns, query.journal_entries, report, query.trial_balances
            )
            result.analysis_data = analysis_data
            out4 = output_dir / f"{year:04d}-{month:02d} 分析简报.xlsx"
            p4_analysis_brief.write_analysis_brief(
                report, query.bank_txns, query.voucher_entries,
                query.trial_balances, out4
            )
            result.output4_analysis = out4
        except Exception as e:
            result.errors.append(f"输出分析简报失败：{e}")

    # === 输出5：增值税申报底稿（需要报表数据） ===
    if p5_vat_declaration and report is not None:
        _report(80, "正在生成增值税申报底稿...")
        try:
            _cleanup_stale_outputs(output_dir, year, month, "增值税")
            # 尝试读取发票台账
            invoice_data = _try_load_invoice_ledger(input_dir, year, month)
            out5, vat_is_template = p5_vat_declaration.generate(
                report_set=report,
                year=year, month=month,
                invoice_data=invoice_data,
                out_path=output_dir,
            )
            result.output5_vat = out5
            if vat_is_template:
                query.warnings.append(
                    "输出5: 增值税申报底稿为模板输出（缺少发票台账数据源）"
                )
        except Exception as e:
            result.errors.append(f"输出增值税申报底稿失败：{e}")

    # === 输出6：所得税预缴底稿（需要报表数据） ===
    if p6_income_tax_declaration and report is not None:
        _report(86, "正在生成企业所得税预缴底稿...")
        try:
            asset_total, employee_count = _try_load_asset_employee(input_dir)
            _cleanup_stale_outputs(output_dir, year, month, "预缴底稿")
            out6, inc_tax_is_template = p6_income_tax_declaration.generate(
                report_set=report,
                year=year, month=month,
                entity_type=entity_type,
                asset_total=asset_total,
                employee_count=employee_count,
                out_path=output_dir,
            )
            result.output6_income_tax = out6
            if inc_tax_is_template:
                query.warnings.append(
                    "输出6: 所得税预缴底稿为模板输出（缺少资产台账/员工名册数据源）"
                )
        except Exception as e:
            result.errors.append(f"输出所得税预缴底稿失败：{e}")

    # === 输出7：个税扣缴明细表（独立于报表） ===
    if p7_payroll_tax:
        _report(92, "正在生成个税扣缴明细表...")
        try:
            # 尝试读取工资表
            employees = _try_load_payroll(input_dir, year, month)
            _cleanup_stale_outputs(output_dir, year, month, "个税")
            out7, pr_tax_is_template = p7_payroll_tax.generate(
                year=year, month=month,
                employees=employees,
                out_path=output_dir,
            )
            result.output7_payroll_tax = out7
            if pr_tax_is_template:
                query.warnings.append(
                    "输出7: 个税扣缴明细表为模板输出（缺少员工工资数据源）"
                )
        except Exception as e:
            result.errors.append(f"输出个税扣缴明细表失败：{e}")

    # === 输出8：记账凭证（从分录生成，产品闭环补齐） ===
    if query.journal_entries:
        _report(96, "正在生成记账凭证...")
        try:
            from .writers import p8_vouchers
            out8 = p8_vouchers.generate(
                query.journal_entries, year, month, out_path=output_dir
            )
            result.output8_vouchers = out8
        except Exception as e:
            result.errors.append(f"输出记账凭证失败：{e}")

    # === 输出9：试算平衡表（会计核心文件，从分录汇总生成） ===
    # 门控用「有分录」而非「有科目余额表」：借贷发生额直接来自分录，
    # 科目余额表仅补充科目编码/名称。只传银行对账单+业务事项时也应出表。
    if query.journal_entries:
        try:
            from .writers import p9_trial_balance
            out9 = p9_trial_balance.generate(
                query.journal_entries, query.trial_balances, prior, year, month,
                out_path=output_dir,
            )
            result.output9_trial_balance = out9
        except Exception as e:
            result.errors.append(f"输出试算平衡表失败：{e}")

    _report(100, "全部输出完成")
    return result


# ---- 工具函数 ----

def _try_load_invoice_ledger(input_dir: Path, year: int, month: int) -> dict | None:
    """尝试从输入目录读取发票台账。

    查找路径：
      输入/{year}-{month}/收入发票/invoice_ledger.xlsx
      输入/企业信息/发票台账.xlsx（全局通用）

    Returns:
        发票汇总 dict 或 None（文件不存在时）
    """
    from .parsers.invoice_ledger import parse_invoice_ledger

    period_dir = f"{year:04d}-{month:02d}"
    candidates = [
        input_dir / period_dir / "收入发票" / "invoice_ledger.xlsx",
        input_dir / period_dir / "收入发票" / "发票台账.xlsx",
        input_dir / "企业信息" / "发票台账.xlsx",
    ]
    for p in candidates:
        if p.exists():
            try:
                return parse_invoice_ledger(p)
            except Exception:
                pass
    # 搜索整个输入目录
    for p in sorted(input_dir.rglob("*发票台账*.xlsx")):
        try:
            return parse_invoice_ledger(p)
        except Exception:
            continue
    for p in sorted(input_dir.rglob("*invoice_ledger*.xlsx")):
        try:
            return parse_invoice_ledger(p)
        except Exception:
            continue
    # 兜底：没有手工台账时，自动扫描发票 PDF 生成台账（体验：用户上传了发票 PDF
    # 就不该再提示"缺发票台账数据源"）
    try:
        from .parsers.invoice_pdf import parse_invoice_pdfs, write_ledger
        import tempfile
        pdfs = [p for p in sorted(input_dir.rglob("*.pdf"))
                if "发票" in p.parent.name or "invoice" in p.name.lower()
                or ("发票" in p.name or "dzfp" in p.name.lower() or "电子" in p.name)]
        # 更宽松：扫描所有含发票特征的 PDF（排除报表/回单/对账单）
        if not pdfs:
            pdfs = [p for p in sorted(input_dir.rglob("*.pdf"))
                    if "回单" not in p.name and "对账单" not in p.name
                    and "报表" not in p.name and "展示" not in p.name]
        invoices = parse_invoice_pdfs(pdfs) if pdfs else []
        if invoices:
            tmp = Path(tempfile.mkdtemp(prefix="inv_")) / "发票台账.xlsx"
            write_ledger(invoices, tmp)
            return parse_invoice_ledger(tmp)
    except Exception:
        pass
    return None


def _try_load_asset_employee(input_dir: Path) -> tuple:
    """尝试从输入目录读取资产台账 + 员工名册（静态配置级）。

    查找路径：
      输入/企业信息/asset_ledger.xlsx 或 资产台账.xlsx
      输入/企业信息/employee_roster.xlsx 或 员工名册.xlsx

    Returns:
        (asset_total: Decimal | None, employee_count: int | None)
    """
    from .parsers.asset_employee_ledger import parse_asset_ledger, parse_employee_roster

    company_dir = input_dir / "企业信息"

    # 资产台账
    asset_total = None
    for name in ("asset_ledger.xlsx", "资产台账.xlsx", "固定资产台账.xlsx"):
        p = company_dir / name
        if p.exists():
            try:
                ad = parse_asset_ledger(p)
                asset_total = ad.get("total_assets")
            except Exception:
                pass
            break

    # 也搜全目录
    if asset_total is None:
        for p in sorted(input_dir.rglob("*资产台账*.xlsx")):
            try:
                ad = parse_asset_ledger(p)
                asset_total = ad.get("total_assets")
            except Exception:
                continue
            break

    # 员工名册
    employee_count = None
    for name in ("employee_roster.xlsx", "员工名册.xlsx", "花名册.xlsx"):
        p = company_dir / name
        if p.exists():
            try:
                rd = parse_employee_roster(p)
                employee_count = rd.get("count")
            except Exception:
                pass
            break

    if employee_count is None:
        for p in sorted(input_dir.rglob("*员工名册*.xlsx")):
            try:
                rd = parse_employee_roster(p)
                employee_count = rd.get("count")
            except Exception:
                continue
            break

    return asset_total, employee_count


def _try_load_payroll(input_dir: Path, year: int, month: int) -> list | None:
    """尝试从输入目录读取当月工资表。

    查找路径：
      输入/{year}-{month}/工资社保/payroll.xlsx 或 工资表.xlsx

    Returns:
        员工 dict 列表（匹配 p7.generate employees 格式）或 None
    """
    from .parsers.payroll_ledger import employees_to_p7_dict, parse_payroll_ledger

    period_dir = f"{year:04d}-{month:02d}"
    candidates = [
        input_dir / period_dir / "工资社保" / "payroll.xlsx",
        input_dir / period_dir / "工资社保" / "工资表.xlsx",
        input_dir / period_dir / "工资表.xlsx",
    ]
    for p in candidates:
        if p.exists():
            try:
                emps = parse_payroll_ledger(p)
                return employees_to_p7_dict(emps) if emps else None
            except Exception:
                pass

    # 全目录搜索
    for p in sorted(input_dir.rglob("*工资表*.xlsx")):
        try:
            emps = parse_payroll_ledger(p)
            if emps:
                return employees_to_p7_dict(emps)
        except Exception:
            continue

    return None


def _find_prior_from_report_pdfs(
    year: int, month: int, warnings: list
) -> PriorPeriod | None:
    """从上期报表 PDF 中找到目标月份的上期报表。

    查找顺序：
    1. config.PRIOR_REPORT_DIR（如果配置了且存在）
    2. 输入目录下的 上期报表/ 子目录
    PDF 按第一页资产负债表日期自动匹配期间。
    """
    import fitz as _fitz
    import re as _re

    # 候选目录：config 配置 → 输入目录子目录
    candidate_dirs = []
    if config.PRIOR_REPORT_DIR and Path(config.PRIOR_REPORT_DIR).exists():
        candidate_dirs.append(Path(config.PRIOR_REPORT_DIR))
    # 也检查输入目录下的常见子目录名
    input_parent = config.INPUT_DIR
    for sub in ("上期报表", "prior_reports", "上月报表"):
        d = input_parent / sub
        if d.exists():
            candidate_dirs.append(d)

    if not candidate_dirs:
        return None

    # 目标上期
    py, pm = (year, month - 1) if month > 1 else (year - 1, 12)
    target_suffix = f"{py}-{pm:02d}"

    for search_dir in candidate_dirs:
        for pdf in sorted(search_dir.glob("*.pdf")):
            try:
                doc = _fitz.open(str(pdf))
                head = doc[0].get_text()[:500]
                doc.close()
            except Exception:
                continue
            # 找日期行：如 "2025-12-31" 或 "2026-01-31"
            for line in head.split("\n"):
                m = _re.search(r"(20\d{2})-(\d{2})-(\d{2})", line)
                if m:
                    pdf_period = f"{m.group(1)}-{m.group(2)}"
                    if pdf_period == target_suffix:
                        try:
                            return prior_report_pdf.parse(pdf)
                        except Exception as e:
                            warnings.append(f"解析上期报表 {pdf.name} 失败：{e}")
                            return None
                    break
    return None


def _classify_pdf(path: Path) -> str:
    """按内容判断 PDF 类型"""
    try:
        import fitz
    except ImportError:
        return "未知"
    try:
        doc = fitz.open(str(path))
        head = doc[0].get_text()[:400]
        doc.close()
    except Exception:
        return "未知"

    if "对账单" in head:
        return "对账单"
    if "回单" in head or "业务回单" in head:
        return "回单"
    if "资产负债表" in head:
        return "上月报表"
    return "未知"
