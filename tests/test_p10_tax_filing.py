# -*- coding: utf-8 -*-
"""输出10（电子税务局报送格式）回归测试

这份文件的正确性只有一条标准：**和税局模板长得一样，且数填在对应行次上**。
电子税务局按固定模板校验上传文件，行次错位或样式被覆盖都会上传失败，
所以这里不只验数值，还要验表头、工作表名、模板样式有没有被 xlwt 默认样式冲掉。
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
import xlrd

from accounting_report.models import ReportSet
from accounting_report.writers import p10_tax_filing as p10

D = Decimal

COMPANY = "示例科技有限公司"
TAX_NO = "91110108MA00000000"


def _report() -> ReportSet:
    """2026-09（季末月）的一期报表，数值取自实际申报过的样例"""
    return ReportSet(
        year=2026, month=9,
        bs_end={1: D("48780.58"), 15: D("159282.92"), 53: D("159282.92")},
        bs_year_start={1: D("99368.42"), 15: D("99368.42"), 53: D("99368.42")},
        is_month={18: D("-1.88"), 30: D("1.88")},
        is_ytd={30: D("-41682.20")},
    )


@pytest.fixture(scope="module")
def generated(tmp_path_factory) -> list[Path]:
    out = tmp_path_factory.mktemp("tax_filing")
    return p10.write_tax_filing(_report(), out, company_name=COMPANY, tax_no=TAX_NO)


def _sheet(path: Path) -> xlrd.sheet.Sheet:
    return xlrd.open_workbook(str(path), formatting_info=True).sheet_by_index(0)


def test_template_exists() -> None:
    """模板必须随仓库分发，否则线上跑到输出10才炸"""
    assert p10.template_path().exists(), p10.template_path()


def test_two_files_named_by_period(generated: list[Path]) -> None:
    names = sorted(p.name for p in generated)
    assert names == ["2026-09 利润表（税局报送）.xls", "2026-09 资产负债表（税局报送）.xls"]


def test_single_sheet_only(generated: list[Path]) -> None:
    """电子税务局按单表上传，多表反而对不上"""
    for p in generated:
        wb = xlrd.open_workbook(str(p))
        assert len(wb.sheet_names()) == 1, wb.sheet_names()


def test_balance_sheet_header_and_values(generated: list[Path]) -> None:
    """表头三要素：纳税人识别号、纳税人名称、所属期。缺一个税局就退件。

    模板里所属期是日期格式单元格，xlrd 读回来是序列号，要按 datemode 还原。
    """
    wb = xlrd.open_workbook(str(generated[0]))
    sh = wb.sheet_by_index(0)
    assert sh.cell_value(2, 3) == TAX_NO
    assert sh.cell_value(2, 7) == COMPANY
    # 季末月按整季：2026-07-01 ~ 2026-09-30
    start = xlrd.xldate_as_tuple(sh.cell_value(3, 3), wb.datemode)[:3]
    end = xlrd.xldate_as_tuple(sh.cell_value(3, 7), wb.datemode)[:3]
    assert start == (2026, 7, 1), start
    assert end == (2026, 9, 30), end


def test_balance_sheet_fills_by_line_number(generated: list[Path]) -> None:
    """按模板行次取数：左右两栏各自的行次都要落到对应单元格"""
    sh = _sheet(generated[0])
    found = {}
    for r in range(sh.nrows):
        for line_col, end_col, start_col in (
            (p10.BS_LINE_L, p10.BS_END_L, p10.BS_START_L),
            (p10.BS_LINE_R, p10.BS_END_R, p10.BS_START_R),
        ):
            ln = sh.cell_value(r, line_col)
            if isinstance(ln, float) and ln:
                found[int(ln)] = (sh.cell_value(r, end_col), sh.cell_value(r, start_col))
    assert found[1] == (48780.58, 99368.42)      # 货币资金
    assert found[15] == (159282.92, 99368.42)    # 资产总计
    assert found[53] == (159282.92, 99368.42)    # 负债和所有者权益总计（右栏）


def test_income_statement_month_and_ytd(generated: list[Path]) -> None:
    sh = _sheet(generated[1])
    assert sh.cell_value(2, 3) == TAX_NO
    assert sh.cell_value(2, 5) == COMPANY
    found = {}
    for r in range(sh.nrows):
        ln = sh.cell_value(r, p10.IS_LINE)
        if isinstance(ln, float) and ln:
            found[int(ln)] = (sh.cell_value(r, p10.IS_MONTH), sh.cell_value(r, p10.IS_YTD))
    assert found[30] == (1.88, -41682.20)         # 净利润：本期 / 本年累计


def test_template_styles_preserved(generated: list[Path]) -> None:
    """xlwt 默认样式会把模板的表格线刷掉，这里抽查有数字的行是否仍有边框"""
    for p in generated:
        wb = xlrd.open_workbook(str(p), formatting_info=True)
        sh = wb.sheet_by_index(0)
        bordered = 0
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                xf = wb.xf_list[sh.cell_xf_index(r, c)]
                if xf.border.top_line_style or xf.border.bottom_line_style:
                    bordered += 1
        assert bordered > 100, f"{p.name} 边框单元格仅 {bordered} 个，模板样式疑似丢失"


def test_header_dates_follow_quarter_rule() -> None:
    """非季末月按当月填；1 月不能出现 month-2=负数 的越界"""
    assert p10._period_range(2026, 9) == (date(2026, 7, 1), date(2026, 9, 30))
    assert p10._period_range(2026, 10) == (date(2026, 10, 1), date(2026, 10, 31))
    assert p10._period_range(2026, 1) == (date(2026, 1, 1), date(2026, 1, 31))


def test_missing_report_numbers_written_as_zero(generated: list[Path]) -> None:
    """报表没给的行次一律写 0，不能留空——留空上传会被税局判为空值错误。

    取行次 2（短期投资）：本用例的报表里没有这一项。
    """
    sh = _sheet(generated[0])
    for r in range(sh.nrows):
        if sh.cell_value(r, p10.BS_LINE_L) == 2.0:
            assert sh.cell_value(r, p10.BS_END_L) == 0.0
            assert sh.cell_value(r, p10.BS_START_L) == 0.0
            return
    pytest.fail("模板里找不到行次 2")
