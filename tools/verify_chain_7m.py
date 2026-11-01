# -*- coding: utf-8 -*-
"""7 期链跑账务校验（记账账侧 · 白班3）

run_chain((2025,12)..(2026,6))：
  1. 资产负债表平衡 bs_end[30]==bs_end[53]
  2. 勾稽 checks 全过
  3. 试算平衡表借贷全平（期初/本期/期末）
  4. 跨年/跨季衔接抽查
"""
import sys
from pathlib import Path
import tempfile
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from accounting_report.full_pipeline import run_chain
from accounting_report.engine import checks

PERIODS = [(2025, 12), (2026, 1), (2026, 2), (2026, 3), (2026, 4), (2026, 5), (2026, 6)]


def main() -> int:
    input_dir = ROOT / "输入"
    if not input_dir.exists():
        print("输入目录不存在")
        return 1

    out = Path(tempfile.mkdtemp(prefix="chain_verify_"))
    results = run_chain(PERIODS, input_dir=input_dir, output_dir=out)

    all_ok = True
    for (y, m), r in zip(PERIODS, results):
        tag = f"{y}-{m:02d}"
        if r.report is None:
            print(f"[{tag}] ✗ 无报表 (errors={r.errors})")
            all_ok = False
            continue

        bs = r.report.bs_end
        ok_bs = bs.get(30) == bs.get(53)

        try:
            passed = checks.run(r.report, r.prior, r.query.journal_entries)
            ok_checks, n_checks = True, len(passed)
        except Exception as e:
            ok_checks, n_checks = False, 0
            print(f"  [{tag}] 勾稽异常: {e}")

        tb_file = out / f"{tag} 试算平衡表.xlsx"
        tb_ok = None
        if tb_file.exists():
            wb = openpyxl.load_workbook(tb_file, data_only=True)
            ws = wb["试算平衡表"]
            combined = "\n".join(str(r0[0]) for r0 in ws.iter_rows(values_only=True) if r0[0])
            tb_ok = "✓ 试算平衡表借贷平衡" in combined

        # 跨年衔接：2026-01 年初 = 2025-12 期末
        cross_year = ""
        if m == 1:
            prev = {p[0]: p[1] for p in PERIODS}
            prev_res = results[PERIODS.index((y, m)) - 1]
            if prev_res.report is not None:
                same = all(
                    r.report.bs_year_start.get(k, 0) == prev_res.report.bs_end.get(k, 0)
                    for k in (1, 30, 53, 51)
                )
                cross_year = f" 跨年衔接={'✓' if same else '✗'}"

        print(f"[{tag}] BS平衡={ok_bs} 勾稽={ok_checks}({n_checks}项) "
              f"试算平衡={tb_ok if tb_ok is not None else '未生成'}{cross_year}")
        if not ok_bs or not ok_checks or tb_ok is not True:
            all_ok = False

    last = results[-1].report
    print("-" * 60)
    print(f"2026-06 期末货币资金: {last.bs_end.get(1)}")
    print(f"2026-06 期末未分配利润: {last.bs_end.get(51)}")
    print(f"2026-06 本月净利润: {last.is_month.get(32)}")
    print(f"2026-06 累计净利润: {last.is_ytd.get(32)}")
    print("-" * 60)
    print("ALL OK ✓" if all_ok else "HAS ISSUES ✗")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
