# -*- coding: utf-8 -*-
"""双击运行的入口（由 运行.bat 调起）

交互式选择模式：
  1. 标准报表（--month）→ 出资产负债表/利润表/现金流量表
  2. 全流程（--full）→ 同时出银行对账/核算汇总/分析简报 + 标准报表
批量调用请直接用 cli.py 命令行。
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from accounting_report.cli import main  # noqa: E402


def default_month() -> str:
    """默认出上个月的表"""
    today = date.today()
    year, month = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
    return f"{year:04d}-{month:02d}"


def detect_input_files() -> bool:
    """检测输入 文件夹下是否有财务资料"""
    input_dir = Path(__file__).resolve().parent / "输入"
    if not input_dir.exists():
        return False
    # 检查是否有凭证/科目余额表/对账单等
    indicators = list(input_dir.glob("*凭证*.*")) + \
                 list(input_dir.glob("*科目余额*.*")) + \
                 list(input_dir.glob("*明细账*.*")) + \
                 list(input_dir.glob("*对账单*.*"))
    return len(indicators) > 0


def run() -> int:
    print("=" * 52)
    print("  会计报表自动计算")
    print("=" * 52)

    suggested = default_month()
    answer = input(f"\n要出哪个月的表？直接回车 = {suggested}\n> ").strip()
    month = answer or suggested

    has_input = detect_input_files()
    if has_input:
        print("\n检测到 输入/ 文件夹下有财务资料，可运行全流程模式。")
        print("  1 = 标准报表（资产负债表+利润表+现金流量表）")
        print("  2 = 全流程四输出（对账报告+核算汇总+会计报表+分析简报）[推荐]")
        mode = input("选哪个？直接回车 = 2\n> ").strip()
        use_full = mode != "1"
    else:
        use_full = False

    args = ["--full" if use_full else "--month", month]
    try:
        code = main(args)
    except SystemExit as e:
        code = int(e.code or 1)

    if code != 0:
        print("\n出表没成功。常见原因和处理办法：\n")
        print("  1. 提示「无法归类到科目」")
        print("     → 说明有笔银行流水程序不认识。打开")
        print("       accounting_report\\rules\\mapping.yaml，")
        print("       照着里面已有的格式，在 bank_rules 下面加一条，再重跑。\n")
        print("  2. 提示「找不到上月报表」")
        print("     → 输入 文件夹里要有上月的 5 表 PDF，")
        print("       或者 输出 文件夹里要有上月本程序生成的 xlsx。\n")
        print("  3. 提示「未定义的科目」")
        print("     → 业务事项.xlsx 里的科目名写错了，")
        print("       正确的科目名见 accounting_report\\rules\\accounts.yaml。\n")
        print("  4. 提示「勾稽校验未通过」")
        print("     → 这是程序在拦截错误结果，别忽略。")
        print("       把完整提示发给开发，不要手工改 xlsx 凑数。\n")

    return code


if __name__ == "__main__":
    raise SystemExit(run())
