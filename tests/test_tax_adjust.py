# -*- coding: utf-8 -*-
"""税务调整项（纳税调整/弥补亏损/减免/已预缴）回归测试。

背景
----
这三类台账（预缴税款台账、纳税调整台账、税收优惠台账）没有统一格式，
改为会计在前端「税务调整」页录入，落到
``输入/<账期>/_tax-adjust_补充说明.txt``（或 ``税务调整.json``）。
本测试覆盖：
  1. 文本/JSON 解析的容错（别名、千分位、括号负数、注释、未知项目）
  2. 调整项正确代入 p6 各行次，且勾稽关系成立（行10−13−14−15 = 行16）
  3. 未提供调整项时行为与改造前完全一致（不能因为加功能改变既有结果）
  4. p5 的已缴增值税额生效
"""
from __future__ import annotations

import sys
import tempfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from accounting_report import tax_adjust  # noqa: E402
from accounting_report.models import ReportSet, ZERO, money  # noqa: E402
from accounting_report.tax_adjust import TaxAdjust  # noqa: E402
from accounting_report.writers import (  # noqa: E402
    p5_vat_declaration as p5,
    p6_income_tax_declaration as p6,
)

CENT = Decimal("0.01")


def D(x) -> Decimal:
    return money(x)


def _make_report(month: int = 8, is_month=None, is_ytd=None) -> ReportSet:
    return ReportSet(
        year=2026, month=month,
        bs_end={}, bs_year_start={},
        is_month=is_month or {}, is_ytd=is_ytd or {},
        cf_month={}, cf_ytd={}, entries=[],
    )


def _line(ws, line_no):
    rows = [r for r in ws.iter_rows(values_only=True) if r and r[0] == line_no]
    assert rows, f"未找到行次 {line_no}"
    return rows[0]


def _paid_row(ws):
    """p5 主表「本期已缴税额」行（第 7 栏，金额在第 6 列）"""
    rows = [r for r in ws.iter_rows(values_only=True)
            if r and r[0] == "本期已缴税额"]
    assert rows, "应存在「本期已缴税额」行"
    return rows[0]


def _gen_p6(rep, tmp, adjust=None):
    out, _ = p6.generate(
        rep, 2026, 8, entity_type="corp",
        asset_total=D("1000000"), employee_count=50,
        adjust=adjust, out_path=tmp,
    )
    return openpyxl.load_workbook(out, data_only=True)["企业所得税A类预缴"]


# ── 解析 ────────────────────────────────────────────────────────────────

def test_parse_text_aliases_and_tolerance():
    """别名、全角冒号、千分位、¥、括号负数、注释都要能认"""
    txt = """
# 2026年09月
纳税调增额：12,000
纳税调减额 3000
弥补以前年度亏损=0
已预缴所得税额 ¥45,000.00
已缴增值税额：(1,200)
其他减免所得税额（吸纳重点群体就业）800
"""
    vals, problems = tax_adjust.parse_text(txt)
    assert vals["increase"] == Decimal("12000")
    assert vals["decrease"] == Decimal("3000")
    assert vals["loss_offset"] == Decimal("0")
    assert vals["income_tax_prepaid"] == Decimal("45000.00")
    assert vals["vat_prepaid"] == Decimal("-1200"), "括号应识别为负数"
    assert vals["other_exemption"] == Decimal("800")
    assert problems == []


def test_parse_text_reports_problems_instead_of_silent_drop():
    """认不出的行必须报出来，不能静默丢 —— 丢一行就是少一笔调增"""
    vals, problems = tax_adjust.parse_text(
        "纳税调增额 100\n这是一行乱写的\n不存在的项目 999\n"
    )
    assert vals["increase"] == Decimal("100")
    assert len(problems) == 2


def test_ambiguous_names_map_to_correct_field():
    """「已缴增值税额」/「已缴所得税额」不能串到对方去"""
    vat, _ = tax_adjust.parse_text("已缴增值税额 500")
    inc, _ = tax_adjust.parse_text("已缴所得税额 800")
    assert vat["vat_prepaid"] == Decimal("500")
    assert "vat_prepaid" not in inc
    assert inc["income_tax_prepaid"] == Decimal("800")


def test_load_prefers_json_over_text():
    tmp = Path(tempfile.mkdtemp())
    (tmp / tax_adjust.TEXT_FILENAME).write_text("纳税调增额 100", encoding="utf-8")
    (tmp / tax_adjust.JSON_FILENAME).write_text(
        '{"纳税调增额": "200"}', encoding="utf-8"
    )
    adj = tax_adjust.load(tmp)
    assert adj.increase == Decimal("200")
    assert adj.source == tax_adjust.JSON_FILENAME


def test_load_missing_file_is_empty_not_crash():
    tmp = Path(tempfile.mkdtemp())
    adj = tax_adjust.load(tmp)
    assert adj.is_empty and adj.increase == ZERO
    assert adj.taxable_adjustment == ZERO


# ── p6 代入 ─────────────────────────────────────────────────────────────

def test_p6_no_adjust_keeps_legacy_numbers():
    """未录入调整项时，结果必须与改造前一致（行5/6/7 为 0，行16 = 利润×5%）"""
    rep = _make_report(is_ytd={1: D("100000"), 30: D("12345.67")})
    ws = _gen_p6(rep, Path(tempfile.mkdtemp()))
    assert _line(ws, 5)[2] == 0
    assert _line(ws, 6)[2] == 0
    assert _line(ws, 7)[2] == 0
    assert abs(_line(ws, 8)[2] - 12345.67) < 0.001
    assert abs(_line(ws, 16)[2] - 617.28) < 0.001


def test_p6_adjust_flows_into_taxable_income_and_tax():
    """调增 5000、调减 1000、弥补亏损 2000 → 应纳税所得额 = 利润 + 2000"""
    rep = _make_report(is_ytd={1: D("100000"), 30: D("100000")})
    adj = TaxAdjust(increase=D("5000"), decrease=D("1000"), loss_offset=D("2000"))
    ws = _gen_p6(rep, Path(tempfile.mkdtemp()), adjust=adj)

    assert abs(_line(ws, 5)[2] - 5000) < 0.001
    assert abs(_line(ws, 6)[2] - 1000) < 0.001
    assert abs(_line(ws, 7)[2] - 2000) < 0.001
    # 行8 = 100000 + 5000 − 1000 − 2000 = 102000
    assert abs(_line(ws, 8)[2] - 102000) < 0.001
    # 行10 = 102000 × 25% = 25500
    assert abs(_line(ws, 10)[2] - 25500) < 0.001


def test_p6_prepaid_and_other_exemption_reduce_due():
    """已预缴 + 其他减免都要从应补退里扣掉，且勾稽关系成立"""
    rep = _make_report(is_ytd={1: D("100000"), 30: D("100000")})
    adj = TaxAdjust(other_exemption=D("1000"), income_tax_prepaid=D("20000"))
    ws = _gen_p6(rep, Path(tempfile.mkdtemp()), adjust=adj)

    line10 = _line(ws, 10)[2]   # 25000
    line13 = _line(ws, 13)[2]   # 小微减免 20000 + 其他 1000
    line14 = _line(ws, 14)[2]   # 20000
    line15 = _line(ws, 15)[2]   # 0
    line16 = _line(ws, 16)[2]

    assert abs(line14 - 20000) < 0.001
    assert abs(line16 - (line10 - line13 - line14 - line15)) < 0.001, "行16 必须等于 10−13−14−15"


def test_p6_exemption_not_deducted_twice():
    """小微已按 5% 计税，行13 再扣一次减免就等于扣两遍 —— 必须靠行10 反算"""
    rep = _make_report(is_ytd={1: D("100000"), 30: D("100000")})
    ws = _gen_p6(rep, Path(tempfile.mkdtemp()))
    # 100000 × 5% = 5000
    assert abs(_line(ws, 16)[2] - 5000) < 0.001


def test_p6_loss_with_prepaid_does_not_emit_negative_refund():
    """亏损月即使有预缴，也不能出负数税额（会被税局判异常申报）"""
    rep = _make_report(is_ytd={1: D("100000"), 30: D("-50000")})
    adj = TaxAdjust(income_tax_prepaid=D("30000"))
    ws = _gen_p6(rep, Path(tempfile.mkdtemp()), adjust=adj)
    assert _line(ws, 10)[2] == 0
    assert _line(ws, 16)[2] == 0


def test_p6_small_profit_judged_on_taxable_income_not_profit():
    """调增后突破 300 万，就不再符合小微 —— 判定基数必须是应纳税所得额"""
    rep = _make_report(is_ytd={1: D("5000000"), 30: D("2990000")})
    adj = TaxAdjust(increase=D("50000"))   # 299 万 + 5 万 = 304 万 > 300 万
    ws = _gen_p6(rep, Path(tempfile.mkdtemp()), adjust=adj)
    line10 = _line(ws, 10)[2]
    line13 = _line(ws, 13)[2]
    # 不再享受小微：行13 为 0，行16 = 3040000 × 25%
    assert line13 == 0
    assert abs(line10 - 760000) < 0.001


# ── p5 已缴增值税 ───────────────────────────────────────────────────────

def test_p5_vat_prepaid_subtracts_from_due():
    rep = _make_report(month=6, is_month={1: D("200000")})
    adj = TaxAdjust(vat_prepaid=D("900"))
    tmp = Path(tempfile.mkdtemp())
    out, _ = p5.generate(
        rep, 2026, 6,
        invoice_data={"invoiced": D("123456.78"), "uninvoiced": D("0"),
                      "tax_exempt": D("0"), "details": []},
        adjust=adj, out_path=tmp,
    )
    ws = openpyxl.load_workbook(out, data_only=True)["增值税申报主表"]
    row7 = _paid_row(ws)
    assert abs(float(row7[5]) - 900) < 0.001, "本期已缴税额应取录入值 900"


def test_p5_without_adjust_unchanged():
    """不传 adjust 时已缴税额仍为 0（保持既有行为）"""
    rep = _make_report(month=6, is_month={1: D("200000")})
    tmp = Path(tempfile.mkdtemp())
    out, _ = p5.generate(
        rep, 2026, 6,
        invoice_data={"invoiced": D("123456.78"), "uninvoiced": D("0"),
                      "tax_exempt": D("0"), "details": []},
        out_path=tmp,
    )
    ws = openpyxl.load_workbook(out, data_only=True)["增值税申报主表"]
    assert float(_paid_row(ws)[5] or 0) == 0
