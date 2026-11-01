# -*- coding: utf-8 -*-
"""端到端流程：读输入 → 生成分录 → 算五张表 → 校验 → 写 xlsx"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import fitz

from .config import INPUT_DIR, OUTPUT_DIR
from .engine import checks, compute, journalize, quarterly
from .models import JournalEntry, PriorPeriod, ReportSet
from .parsers import (business_events, icbc_receipt, icbc_statement,
                      prior_report_pdf, prior_report_xlsx)
from .writers import xlsx_writer

BUSINESS_EVENTS_FILE = "业务事项.xlsx"
INVOICE_CSV_FILE = "票据记录.csv"


def _dedup_bank_txns(txns: list) -> list:
    """按 (日期, 流入, 流出, 对方户名, 摘要) 去重银行流水，保留字段最完整的一条。

    输入目录中同名/重复 PDF（如对账单1.pdf 和 对账单1(1).pdf）
    会导致同一笔交易被解析多次，去重避免重复入账导致余额偏大。

    去重键含对方户名与摘要两道区分：
      - 仅按 (日期, 流入, 流出) 会把同日同额的两笔**不同**交易误合并——
        真实账套中同一天购买两笔同额理财、两笔同额第三方收款（不同平台）
        都发生过（白班13 实测：2026-02-05 两笔 10000 支出被误并）。
      - 对方户名相同仍不够：工行对账单同一天的理财申购（24G0007C 与
        24GS2288 等不同产品）对方户名同是「法人代销理财现金管理产品」、
        金额同额——仅靠 (日期,金额,对方户名) 仍会把真实的两笔并成一笔，
        余额列连续性因此断裂、对账报告调节失衡（白班19 实测：2026-02 余
        额差 20000）。摘要（产品单号）是同额交易的最后区分符。
    跨 PDF 重复解析的同一笔，其摘要文本与对方户名文本一致，仍能被正确合并。
    """
    import re
    def _source_priority(src: str) -> int:
        return len(re.findall(r'\(\d+\)', src or ''))

    def _norm(s: str) -> str:
        return (s or '').replace(' ', '').replace('\u3000', '')

    groups: dict[tuple, list] = {}
    for t in txns:
        key = (t.date, str(t.inflow), str(t.outflow),
               _norm(t.counterparty), _norm(getattr(t, 'summary', '')))
        groups.setdefault(key, []).append(t)

    result = []
    for key, group in groups.items():
        if len(group) == 1:
            result.append(group[0])
        else:
            def _richness(t) -> tuple:
                return (len(_norm(t.counterparty)), len(_norm(getattr(t, 'summary', ''))),
                        -_source_priority(getattr(t, 'source', '')))
            best = max(group, key=_richness)
            result.append(best)
    return result


@dataclass
class Inputs:
    prior_report: Path | None = None       # 上月 5 表 PDF
    prior_xlsx: Path | None = None         # 上月本程序输出
    statements: list[Path] = field(default_factory=list)
    receipts: list[Path] = field(default_factory=list)
    events: Path | None = None
    invoices: Path | None = None


def _classify_pdf(path: Path) -> str:
    """按内容判断 PDF 是哪一类，不依赖文件名"""
    try:
        doc = fitz.open(str(path))
        head = doc[0].get_text()[:300]
        pages = len(doc)
        doc.close()
    except Exception:
        return "未知"

    if "对账单" in head:
        return "对账单"
    if "业务回单" in head or "回单编号" in head:
        return "回单"
    if head.lstrip().startswith("资产负债表") and pages >= 3:
        return "上月报表"
    return "未知"


def discover(input_dir: Path, output_dir: Path, year: int, month: int) -> Inputs:
    found = Inputs()

    for pdf in sorted(input_dir.rglob("*.pdf")):
        kind = _classify_pdf(pdf)
        if kind == "对账单":
            found.statements.append(pdf)
        elif kind == "回单":
            found.receipts.append(pdf)
        elif kind == "上月报表" and found.prior_report is None:
            # 过滤演示/展示样张：优先取真实月份报表，避免 (1)(2) 重名样张
            if "展示" in pdf.name or "样例" in pdf.name or "示例" in pdf.name:
                continue
            found.prior_report = pdf

    events = input_dir / BUSINESS_EVENTS_FILE
    if events.exists():
        found.events = events
    invoices = input_dir / INVOICE_CSV_FILE
    if invoices.exists():
        found.invoices = invoices

    # 上月本程序的输出（滚动出表时优先用它）
    py, pm = (year, month - 1) if month > 1 else (year - 1, 12)
    candidate = output_dir / f"{py:04d}-{pm:02d} 会计报表.xlsx"
    if candidate.exists():
        found.prior_xlsx = candidate

    return found


def load_prior(found: Inputs) -> PriorPeriod:
    if found.prior_xlsx is not None:
        return prior_report_xlsx.parse(found.prior_xlsx)
    if found.prior_report is not None:
        return prior_report_pdf.parse(found.prior_report)
    raise FileNotFoundError(
        "找不到上月报表：既没有上月的输出 xlsx，输入目录里也没有 5 表 PDF"
    )


def collect_entries(found: Inputs, year: int, month: int) -> list[JournalEntry]:
    month_prefix = f"{year:04d}-{month:02d}"
    entries: list[JournalEntry] = []

    # 银行流水（跨文件去重：输入目录常有多份同名/跨期重复 PDF，
    # 同一笔交易会被解析多次导致重复入账）
    txns = []
    for pdf in found.statements:
        txns.extend(icbc_statement.parse(pdf))
    txns = [t for t in txns if t.date.startswith(month_prefix)]
    txns = _dedup_bank_txns(txns)
    entries.extend(journalize.from_bank(txns))

    # 业务事项（挂账费用、计提等不走银行的业务）
    if found.events:
        entries.extend(
            e for e in business_events.parse_xlsx(found.events)
            if e.date.startswith(month_prefix)
        )

    # 票据 OCR 产出
    if found.invoices:
        entries.extend(business_events.parse_invoice_csv(found.invoices, month_prefix))

    return entries


def cross_check_receipts(found: Inputs, entries: list[JournalEntry],
                         year: int, month: int) -> list[str]:
    """回单与分录交叉核对，返回提示信息（不阻断出表）

    只核对本月的回单——输入目录里往往还堆着往月的单据，不过滤会天天报假警。
    """
    month_prefix = f"{year:04d}-{month:02d}"
    notes: list[str] = []
    amounts = [e.amount for e in entries]

    for pdf in found.receipts:
        r = icbc_receipt.parse(pdf)
        if not r.date.startswith(month_prefix):
            continue
        if r.amount in amounts:
            notes.append(f"回单核对通过：{r.date} {r.summary} {r.amount}")
        else:
            notes.append(
                f"⚠ 回单 {r.date} {r.summary} {r.amount} 在本月分录中找不到对应金额，"
                f"可能漏记了一笔"
            )
    return notes


@dataclass
class Result:
    report: ReportSet
    out_path: Path
    entries: list[JournalEntry]
    passed: list[str]
    notes: list[str]
    inputs: Inputs


def run(year: int, month: int,
        input_dir: Path | str = INPUT_DIR,
        output_dir: Path | str = OUTPUT_DIR,
        out_path: Path | str | None = None) -> Result:
    input_dir, output_dir = Path(input_dir), Path(output_dir)

    found = discover(input_dir, output_dir, year, month)
    prior = load_prior(found)

    expected_prior = (year, month - 1) if month > 1 else (year - 1, 12)
    if (prior.year, prior.month) != expected_prior:
        raise ValueError(
            f"上月报表期间是 {prior.year}-{prior.month:02d}，"
            f"但要出 {year}-{month:02d} 的表，应当以 "
            f"{expected_prior[0]}-{expected_prior[1]:02d} 为上期"
        )

    entries = collect_entries(found, year, month)
    notes = cross_check_receipts(found, entries, year, month)

    report = compute.build(prior, entries, year, month)
    quarterly.build(report, prior)
    passed = checks.run(report, prior, entries)

    if out_path is None:
        out_path = output_dir / f"{year:04d}-{month:02d} 会计报表.xlsx"
    written = xlsx_writer.write(report, out_path)

    return Result(report=report, out_path=written, entries=entries,
                  passed=passed, notes=notes, inputs=found)
