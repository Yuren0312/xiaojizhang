# -*- coding: utf-8 -*-
"""命令行入口

    python -m accounting_report --month 2026-05
    python -m accounting_report --month 2026-06 --out 输出/6月.xlsx
    python -m accounting_report --full 2026-02          # 九类输出全流程
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import INPUT_DIR, OUTPUT_DIR
from .engine.checks import ReconciliationError
from .engine.journalize import UnmappedTransaction
from .pipeline import run
from .full_pipeline import run_full


def _parse_month(text: str) -> tuple[int, int]:
    parts = text.replace("/", "-").split("-")
    if len(parts) != 2:
        raise argparse.ArgumentTypeError(f"月份格式应为 YYYY-MM，收到 {text!r}")
    year, month = int(parts[0]), int(parts[1])
    if not 1 <= month <= 12:
        raise argparse.ArgumentTypeError(f"月份超出范围：{month}")
    return year, month


def _init(input_dir: Path, output_dir: Path) -> int:
    """首次使用：建目录 + 生成空的业务事项表"""
    from .parsers.business_events import COLUMNS, write_template

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"输入目录  {input_dir}")
    print(f"输出目录  {output_dir}")

    events = input_dir / "业务事项.xlsx"
    if events.exists():
        print(f"业务事项表已存在，未覆盖：{events}")
    else:
        write_template(events)
        print(f"已生成业务事项表：{events}")
        print(f"  列：{' / '.join(COLUMNS)}")

    print("\n接下来：")
    print(f"  1. 把本月工行对账单 PDF 放进 {input_dir}")
    print(f"  2. 首次还需放入上月 5 表 PDF（之后程序自动读上月输出）")
    print(f"  3. 不走银行的业务（挂账费用、计提等）填进 业务事项.xlsx")
    print(f"  4. 双击 运行.bat，或 python -m accounting_report --full YYYY-MM")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="accounting_report", description="会计AI工具：根据全套财务资料自动输出九类账务文件")
    parser.add_argument("--month", type=_parse_month, help="要出表的月份，如 2026-05（标准报表模式）")
    parser.add_argument("--full", type=_parse_month, dest="full_month",
                        help="全流程模式：报表+对账+核算汇总+分析简报+税务底稿+记账凭证+试算平衡表，如 2026-02")
    parser.add_argument("--input", default=INPUT_DIR, type=Path, help="输入目录")
    parser.add_argument("--output", default=OUTPUT_DIR, type=Path, help="输出目录")
    parser.add_argument("--out", default=None, type=Path, help="输出文件路径（覆盖默认命名）")
    parser.add_argument("--init", action="store_true",
                        help="初始化：建好输入/输出目录和一张空的业务事项表")
    args = parser.parse_args(argv)

    if args.init:
        return _init(args.input, args.output)

    # === 全流程多输出模式 ===
    if args.full_month:
        year, month = args.full_month
        if not args.input.exists():
            print(f"\n输入目录不存在：{args.input}", file=sys.stderr)
            return 1
        print(f"\n=== 全流程多输出模式：{year}-{month:02d} ===\n")
        try:
            result = run_full(year, month, args.input, args.output)
        except Exception as e:
            print(f"\n运行失败：{e}\n", file=sys.stderr)
            return 1

        print(f"Query 层：")
        print(f"  分录    {len(result.query.journal_entries)} 笔")
        print(f"  银行流水 {len(result.query.bank_txns)} 笔")
        print(f"  科目余额 {len(result.query.trial_balances)} 条")
        if result.query.warnings:
            for w in result.query.warnings[:3]:
                print(f"  ⚠ {w}")

        print(f"\n输出文件：")
        for label, path in [
            ("输出1 银行对账报告", result.output1_reconciliation),
            ("输出2 核算汇总表", result.output2_summary),
            ("输出3 标准财务报表", result.output3_report),
            ("输出4 分析简报", result.output4_analysis),
            ("输出5 增值税申报底稿", result.output5_vat),
            ("输出6 所得税预缴底稿", result.output6_income_tax),
            ("输出7 个税扣缴明细", result.output7_payroll_tax),
            ("输出8 记账凭证", result.output8_vouchers),
            ("输出9 试算平衡表", result.output9_trial_balance),
        ]:
            status = f"✅ {path}" if path else "❌ 未生成"
            print(f"  {label}: {status}")

        if result.errors:
            print(f"\n错误：")
            for e in result.errors:
                print(f"  ❌ {e}")
            return 1

        print(f"\n完成 ✅\n")
        return 0

    # === 原标准报表模式 ===
    if args.month is None:
        parser.error("要出表就给 --month（标准报表）或 --full（全流程），如 --full 2026-02；首次使用请先跑 --init")

    year, month = args.month

    if not args.input.exists():
        print(f"\n输入目录不存在：{args.input}\n先运行 python -m accounting_report --init\n",
              file=sys.stderr)
        return 1

    try:
        result = run(year, month, args.input, args.output, args.out)
    except (UnmappedTransaction, ReconciliationError, FileNotFoundError, ValueError) as e:
        print(f"\n出表失败：{e}\n", file=sys.stderr)
        return 1

    src = result.inputs
    print(f"\n=== {year}-{month:02d} 会计报表 ===\n")
    print("输入：")
    print(f"  上期取数  {src.prior_xlsx or src.prior_report}")
    for p in src.statements:
        print(f"  对账单    {p.name}")
    for p in src.receipts:
        print(f"  业务回单  {p.name}")
    if src.events:
        print(f"  业务事项  {src.events.name}")
    if src.invoices:
        print(f"  票据记录  {src.invoices.name}")

    print(f"\n本月分录 {len(result.entries)} 笔：")
    for e in result.entries:
        print(f"  {e.date}  借 {e.debit} / 贷 {e.credit}  {e.amount}   {e.summary}")

    if result.notes:
        print("\n单据核对：")
        for note in result.notes:
            print(f"  {note}")

    print("\n勾稽校验：")
    for item in result.passed:
        print(f"  ✓ {item}")

    r = result.report
    print("\n关键结果：")
    print(f"  货币资金期末      {r.bs_end[1]}")
    print(f"  资产合计          {r.bs_end[30]}")
    print(f"  负债和所有者权益  {r.bs_end[53]}")
    print(f"  本月净利润        {r.is_month[32]}")
    print(f"  本年累计净利润    {r.is_ytd[32]}")

    print(f"\n已输出 → {result.out_path}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
