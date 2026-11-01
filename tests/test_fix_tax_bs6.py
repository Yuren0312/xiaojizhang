# -*- coding: utf-8 -*-
"""税务输出修复回归测试（BS6 批）

覆盖：
  M4 税额 = 基数 × 税率 一律量化到分，不得有 12.3456 这类多位小数写入申报表
  M5 利润总额为负（亏损月）时税额/应补退置 0，不得出现负退税
  M6 p7 累计减除费用月数与累计收入覆盖月数一致（单月收入 5000×1）
  M9 p5 收入核对④差异 = ① − (② + ③无票收入)，不再恒为 0
  L4 p6 行4 标签与取值一致；资产/人数数据缺失时状态文案提示待补录
"""
from __future__ import annotations

import sys
import tempfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from accounting_report.models import ReportSet, ZERO, money  # noqa: E402
from accounting_report.writers import (  # noqa: E402
    p5_vat_declaration as p5,
    p6_income_tax_declaration as p6,
    p7_payroll_tax as p7,
)

CENT = Decimal("0.01")


def D(x) -> Decimal:
    return money(x)


def _is_two_decimal(v) -> bool:
    """单元格值量化到分：str 转 Decimal 后最多 2 位小数"""
    if v is None:
        return True
    d = Decimal(str(v))
    return d == d.quantize(CENT, rounding=ROUND_HALF_UP)


def _find_rows(ws, *keywords):
    """返回包含所有关键词的行列表（values_only 元组），关键词可出现在任意列"""
    out = []
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        cells = [str(c) for c in r if c is not None]
        if cells and all(any(k in c for c in cells) for k in keywords):
            out.append(r)
    return out


def _line(ws, line_no):
    """按第1列行次取数据行（p6 主表/明细表的行次列）"""
    rows = [r for r in ws.iter_rows(values_only=True) if r and r[0] == line_no]
    assert rows, f"行次 {line_no} 不存在"
    return rows[0]


def _label_row(ws, label):
    """第1列精确等于 label 的数据行（p5 主表）"""
    rows = [r for r in ws.iter_rows(values_only=True) if r and r[0] == label]
    assert rows, f"标签行 {label} 不存在"
    return rows[0]


def _make_report(month=6, is_month=None, is_ytd=None) -> ReportSet:
    return ReportSet(
        year=2026, month=month,
        bs_end={}, bs_year_start={},
        is_month=is_month or {}, is_ytd=is_ytd or {},
        cf_month={}, cf_ytd={}, entries=[],
    )


# ── M4：税额量化到分 ─────────────────────────────────────────────

def test_p5_vat_payable_quantized_to_cents():
    """p5：计税销售额 123456.78 → 应纳税额 1234.5678 应舍入为 1234.57

    收入与开票一致（123456.78），无票收入为 0；无票收入补差场景由
    test_regression_night 区块 36 覆盖（白班12 后计税依据 = 利润表收入）。"""
    rep = _make_report(is_month={1: D("123456.78")})
    tmp = Path(tempfile.mkdtemp())
    out, _tmpl = p5.generate(rep, 2026, 6, invoice_data={
        "invoiced": D("123456.78"), "uninvoiced": D("0"),
        "tax_exempt": D("0"),
        "details": [{"code": "11", "number": "1", "date": "2026-06-01",
                     "buyer": "客户", "service": "服务",
                     "amount": D("123456.78"), "tax": D("1234.57"),
                     "direction": "销项"}],
    }, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["增值税申报主表"]
    row4 = _label_row(ws, "本期应纳税额")
    cell_val = row4[4]  # 第5列「本期应纳税额」
    assert _is_two_decimal(cell_val), f"应纳税额应量化到分，实际 {cell_val}"
    assert abs(cell_val - 1234.57) < 0.001, f"应纳税额应为 1234.57，实际 {cell_val}"

    # 附加税费同样量化到分：1234.57 × 7% × 50% = 43.20995 → 43.21
    ws2 = wb["附加税费计算"]
    urban = _find_rows(ws2, "城市维护建设税")
    assert urban, "应存在城市维护建设税行"
    assert _is_two_decimal(urban[0][4]), f"附加税应量化到分，实际 {urban[0][4]}"
    assert abs(urban[0][4] - 43.21) < 0.001


def test_p5_tax_reduction_quantized_to_cents():
    """p5：减征额 = 基数 × (3% − 1%) 也应量化到分"""
    rep = _make_report(is_month={1: D("200000")})
    tmp = Path(tempfile.mkdtemp())
    out, _ = p5.generate(rep, 2026, 6, invoice_data={
        "invoiced": D("123456.78"), "uninvoiced": D("0"),
        "tax_exempt": D("0"), "details": [],
    }, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["增值税申报主表"]
    rows = _find_rows(ws, "减征额")
    assert rows, "应存在减征额行"
    assert _is_two_decimal(rows[0][4]), f"减征额应量化到分，实际 {rows[0][4]}"


def test_p6_tax_payable_quantized_to_cents():
    """p6：利润总额 12345.67 × 5% = 617.2835 → 应纳所得税 617.28"""
    rep = _make_report(month=8, is_ytd={1: D("100000"), 30: D("12345.67")})
    tmp = Path(tempfile.mkdtemp())
    out, _ = p6.generate(rep, 2026, 8, entity_type="corp",
                         asset_total=D("1000000"), employee_count=50,
                         out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["企业所得税A类预缴"]
    row10 = _line(ws, 10)
    row16 = _line(ws, 16)
    # 行10 = 12345.67 × 25% = 3086.4175 → 3086.42
    assert _is_two_decimal(row10[2]) and abs(row10[2] - 3086.42) < 0.001
    # 行16 = 12345.67 × 5% = 617.2835 → 617.28
    assert _is_two_decimal(row16[2]) and abs(row16[2] - 617.28) < 0.001


def test_calc_tax_and_individual_quantized_to_cents():
    """_calc_tax / _calc_individual_tax 返回值量化到分"""
    from decimal import Decimal as Dec
    # 36.6666 × 3% = 1.099998 → 1.10
    assert p7._calc_tax(Dec("36.6666")) == Dec("1.10")
    # 30000 × 5% = 1500 → 1500.00
    assert p6._calc_individual_tax(Dec("30000")) == Dec("1500.00")


# ── M5：亏损月税额 = 0，无负退税 ────────────────────────────────

def test_p6_loss_month_no_negative_tax():
    """亏损月（利润总额 < 0）：行10 应纳所得税 = 0、行16 应补退 = 0，不得出现负退税"""
    rep = _make_report(month=8, is_ytd={1: D("100000"), 30: D("-50000")})
    tmp = Path(tempfile.mkdtemp())
    out, _ = p6.generate(rep, 2026, 8, entity_type="corp",
                         asset_total=D("10000000"), employee_count=50,
                         out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["企业所得税A类预缴"]
    row10 = _line(ws, 10)
    row16 = _line(ws, 16)
    assert row10[2] == 0, f"亏损月应纳所得税额应为 0，实际 {row10[2]}"
    assert row16[2] == 0, f"亏损月应补退税额应为 0，实际 {row16[2]}"
    # 全表金额不得为负（行4/行8 允许负数表达亏损，税额相关行为须为 0）
    for r in ws.iter_rows(values_only=True):
        if isinstance(r[2], (int, float)) and r[0] in (10, 11, 12, 13, 14, 15, 16):
            assert r[2] >= 0, f"行 {r[0]} 金额不得为负，实际 {r[2]}"


# ── M6：p7 累计减除费用口径 ─────────────────────────────────────

def test_p7_single_month_deduction_5000():
    """8月新员工只有单月收入 20000：累计减除 = 5000×1，不得按自然月号减 40000"""
    tmp = Path(tempfile.mkdtemp())
    emps = [{"name": "新员工", "id_card": "X1", "monthly_income": D("20000")}]
    out, _ = p7.generate(2026, 8, employees=emps, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["个税扣缴明细"]
    row = _find_rows(ws, "新员工")
    assert row, "应存在新员工行"
    # 列11（累计减除费用）= 5000.0；列14（累计应纳税所得额）= 15000.0
    assert abs(row[0][10] - 5000.0) < 0.001, f"单月收入累计减除应 5000，实际 {row[0][10]}"
    assert abs(row[0][13] - 15000.0) < 0.001, f"累计应纳税所得额应 15000，实际 {row[0][13]}"


def test_p7_cumulative_months_used_when_provided():
    """提供了 cumulative_income + cumulative_months 时按 5000×月数 减除"""
    tmp = Path(tempfile.mkdtemp())
    emps = [{"name": "老员工", "id_card": "L1", "monthly_income": D("20000"),
             "cumulative_income": D("160000"), "cumulative_months": 8}]
    out, _ = p7.generate(2026, 8, employees=emps, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["个税扣缴明细"]
    row = _find_rows(ws, "老员工")
    assert row
    assert abs(row[0][10] - 40000.0) < 0.001, f"累计减除应 40000，实际 {row[0][10]}"
    assert abs(row[0][13] - 120000.0) < 0.001


def test_p7_cumulative_months_default_one():
    """有 cumulative_income 但未提供 cumulative_months：缺省按 1 个月，不虚减"""
    tmp = Path(tempfile.mkdtemp())
    emps = [{"name": "缺月数", "id_card": "M1", "monthly_income": D("20000"),
             "cumulative_income": D("160000")}]
    out, _ = p7.generate(2026, 8, employees=emps, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["个税扣缴明细"]
    row = _find_rows(ws, "缺月数")
    assert row
    assert abs(row[0][10] - 5000.0) < 0.001, f"缺省累计减除应 5000，实际 {row[0][10]}"


# ── M9：p5 收入核对差异真实值 ───────────────────────────────────

def test_p5_income_recheck_difference_real_value():
    """④ 差异 = ① − ② − ③无票收入，不再是恒为 0 的自减表达式。
    revenue=200000，销项发票=123456.78，无票收入=50000 → 差异=26543.22
    """
    rep = _make_report(is_month={1: D("200000")})
    tmp = Path(tempfile.mkdtemp())
    out, _ = p5.generate(rep, 2026, 6, invoice_data={
        "invoiced": D("123456.78"), "uninvoiced": D("50000"),
        "tax_exempt": D("0"),
        "details": [{"code": "11", "number": "1", "date": "2026-06-01",
                     "buyer": "客户", "service": "服务",
                     "amount": D("123456.78"), "tax": D("1234.57"),
                     "direction": "销项"}],
    }, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["发票与收入明细"]
    diff_rows = _find_rows(ws, "差异")
    assert diff_rows, "应存在差异行"
    label, val, _note = diff_rows[0][0], diff_rows[0][1], diff_rows[0][2]
    assert "① − ② − ③" in str(label), f"差异标签应说明公式，实际 {label}"
    assert abs(val - 26543.22) < 0.001, f"差异应为 26543.22，实际 {val}"


def test_p5_income_recheck_difference_zero_when_uninvoiced_missing():
    """发票台账未提供 uninvoiced 时，③按①−②推算，④差异保持 0（不虚报差异）"""
    rep = _make_report(is_month={1: D("200000")})
    tmp = Path(tempfile.mkdtemp())
    # invoice_data 不含 uninvoiced 字段
    out, _ = p5.generate(rep, 2026, 6, invoice_data={
        "invoiced": D("123456.78"), "tax_exempt": D("0"),
        "details": [{"code": "11", "number": "1", "date": "2026-06-01",
                     "buyer": "客户", "service": "服务",
                     "amount": D("123456.78"), "tax": D("1234.57"),
                     "direction": "销项"}],
    }, out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["发票与收入明细"]
    diff_rows = _find_rows(ws, "差异")
    assert diff_rows
    val = diff_rows[0][1]
    assert abs(val - 0.0) < 0.001, f"未提供 uninvoiced 时差异应为 0，实际 {val}"


# ── L4：p6 行4 标签 + 数据缺失状态 ──────────────────────────────

def test_p6_line4_label_is_profit_total():
    """行4 标签改为「利润总额」，不再标注 1−2−3（取值是含费用/投资收益的利润总额）"""
    rep = _make_report(month=8, is_ytd={1: D("100000"), 30: D("20000")})
    tmp = Path(tempfile.mkdtemp())
    out, _ = p6.generate(rep, 2026, 8, entity_type="corp",
                         asset_total=D("1000000"), employee_count=50,
                         out_path=tmp)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["企业所得税A类预缴"]
    row4 = _line(ws, 4)
    assert "利润总额" in str(row4[1]), f"行4 标签应含利润总额，实际 {row4[1]}"
    assert "1−2−3" not in str(row4[1]), f"行4 标签不应含 1−2−3，实际 {row4[1]}"


def test_p6_data_missing_status_warns_pending():
    """资产/人数数据缺失（模板）时，状态文案提示待补录，不得断言「符合小微」"""
    rep = _make_report(month=8, is_ytd={1: D("100000"), 30: D("20000")})
    tmp = Path(tempfile.mkdtemp())
    out, is_tmpl = p6.generate(rep, 2026, 8, entity_type="corp",
                               asset_total=None, employee_count=None,
                               out_path=tmp)
    assert is_tmpl is True, "资产/人数缺失应生成模板输出"
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["企业所得税A类预缴"]
    combined = "\n".join(str(r[0]) for r in ws.iter_rows(values_only=True) if r[0])
    assert "数据缺失" in combined and "待补录" in combined, \
        f"缺失时应提示待补录，实际: {combined[-200:]}"
    assert "✅ 符合小微企业条件" not in combined, "数据缺失时不得断言符合小微"

    # 有数据时仍正常判定
    out2, is_tmpl2 = p6.generate(rep, 2026, 8, entity_type="corp",
                                 asset_total=D("1000000"), employee_count=50,
                                 out_path=tmp)
    assert is_tmpl2 is False
    wb2 = openpyxl.load_workbook(out2, data_only=True)
    ws2 = wb2["企业所得税A类预缴"]
    combined2 = "\n".join(str(r[0]) for r in ws2.iter_rows(values_only=True) if r[0])
    assert "✅ 符合小微企业条件" in combined2
