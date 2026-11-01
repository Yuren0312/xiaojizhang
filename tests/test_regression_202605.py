# -*- coding: utf-8 -*-
"""回归基线：用 4 月输入算 5 月，必须逐格复现「输出 - 5月会计报表.xlsx」

这是本项目的硬验收线。5 月成品的勾稽关系已逐条验算成立，可信；
只要程序算出来跟它一格不差，就说明取数与算法都对。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from accounting_report.config import SHEETS  # noqa: E402
from accounting_report.pipeline import run  # noqa: E402

BASELINE = ROOT / "输出 - 5月会计报表.xlsx"
INPUT_DIR = ROOT / "输入"
YEAR, MONTH = 2026, 5

# 真实账套数据不进版本库（含银行账号、回单验证码、公司完整财务状况）。
# 公开仓库里跑测试时这些文件不存在，本组测试整体跳过；
# 引擎逻辑本身由 test_engine.py 用虚构数字全覆盖，不依赖真实数据。
pytestmark = pytest.mark.skipif(
    not BASELINE.exists() or not any(INPUT_DIR.glob("*.pdf")),
    reason="本机没有真实账套数据（输出 - 5月会计报表.xlsx / 输入/*.pdf），跳过回归比对",
)


def normalize(value):
    """None 与空串等价；浮点按分四舍五入，消除二进制表示差异"""
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    return str(value).strip()


@pytest.fixture(scope="module")
def generated(tmp_path_factory):
    out = tmp_path_factory.mktemp("report") / "regression.xlsx"
    # 显式不使用上月输出 xlsx，强制走 PDF 冷启动这条路
    result = run(YEAR, MONTH,
                 input_dir=ROOT / "输入",
                 output_dir=tmp_path_factory.mktemp("empty"),
                 out_path=out)
    return result


def test_checks_all_passed(generated):
    assert generated.passed, "勾稽校验一条都没跑"
    # checks.run 不通过会直接抛异常，能走到这里说明全过
    assert any("资产负债表平衡" in p for p in generated.passed)


def test_entries(generated):
    """当期两笔业务：一笔银行进账走利息、一笔挂账费用不动钱"""
    entries = sorted(generated.entries, key=lambda e: e.amount)
    assert len(entries) == 2
    assert (entries[0].debit, entries[0].credit) == ("货币资金", "财务费用-利息")
    assert (entries[1].debit, entries[1].credit) == ("营业成本", "应付账款")


@pytest.mark.parametrize("sheet", SHEETS)
def test_sheet_matches_baseline(generated, sheet):
    expected_wb = openpyxl.load_workbook(BASELINE)
    actual_wb = openpyxl.load_workbook(generated.out_path)

    exp, act = expected_wb[sheet], actual_wb[sheet]
    rows = max(exp.max_row, act.max_row)
    cols = max(exp.max_column, act.max_column)

    diffs = []
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            e = normalize(exp.cell(r, c).value)
            a = normalize(act.cell(r, c).value)
            if e != a:
                coord = exp.cell(r, c).coordinate
                diffs.append(f"{sheet}!{coord}: 期望 {e!r}，实际 {a!r}")

    assert not diffs, f"{sheet} 有 {len(diffs)} 处不一致：\n  " + "\n  ".join(diffs[:25])
