# -*- coding: utf-8 -*-
"""ReportSet → xlsx

版式来自 templates/报表模板.xlsx（从现有报表清空数值而来），
这里只填数值和日期，合并区、列宽、字体、边框、数字格式都保持原样。

零值的写法照搬样表的习惯：四张表留空，现金流量表季报显式写 0.00。
"""
from __future__ import annotations

import calendar
import shutil
from decimal import Decimal
from pathlib import Path

import openpyxl

from ..config import (COMPANY_NAME, SHEET_BS, SHEET_CF, SHEET_CF_Q, SHEET_IS,
                      SHEET_IS_Q, TEMPLATE_PATH)
from ..models import ReportSet, ZERO
from ..rules import layout

# 每张表：模板里的列名 → 从 ReportSet 的哪个字段取数
SOURCES = {
    SHEET_BS:   {"end": "bs_end", "year_start": "bs_year_start"},
    SHEET_IS:   {"ytd": "is_ytd", "month": "is_month"},
    SHEET_CF:   {"ytd": "cf_ytd", "month": "cf_month"},
    SHEET_IS_Q: {"ytd": "is_ytd", "q1": ("is_quarters", 1), "q2": ("is_quarters", 2),
                 "q3": ("is_quarters", 3), "q4": ("is_quarters", 4)},
    SHEET_CF_Q: {"ytd": "cf_ytd", "q1": ("cf_quarters", 1), "q2": ("cf_quarters", 2),
                 "q3": ("cf_quarters", 3), "q4": ("cf_quarters", 4)},
}


def _value(report: ReportSet, source, line: int) -> Decimal:
    if isinstance(source, tuple):
        field, quarter = source
        return getattr(report, field).get(line, {}).get(quarter, ZERO)
    return getattr(report, source).get(line, ZERO)


def _date_text(kind: str, year: int, month: int) -> str:
    if kind == "month_end":
        return f"{year:04d}-{month:02d}-{calendar.monthrange(year, month)[1]:02d}"
    return f"{year:04d}-{month:02d}"


def write(report: ReportSet, out_path: str | Path) -> Path:
    out_path = Path(out_path)
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"找不到报表模板 {TEMPLATE_PATH}，请先运行 python tools/build_template.py"
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(TEMPLATE_PATH, out_path)

    wb = openpyxl.load_workbook(out_path)
    spec = layout()

    for sheet, sources in SOURCES.items():
        ws = wb[sheet]
        conf = spec[sheet]
        blank = conf["blank"]

        ws[conf["date_cell"]] = _date_text(conf["date_kind"], report.year, report.month)
        # 模板里只留「企业名称: 」，公司名按配置填入，换个主体不用改模板
        ws[conf["company_cell"]] = f"企业名称: {COMPANY_NAME}"

        for line, cells in conf["lines"].items():
            line = int(line)
            for col_name, source in sources.items():
                col = cells.get(col_name)
                if not col:
                    continue
                value = _value(report, source, line)
                ws[f"{col}{cells['row']}"] = blank if value == ZERO else float(value)

    # 清理模板残留的操作提示文字（利润表/现金流量表右侧第 5 列起，如
    # 「温馨提示：冲销损益类科目记账方向…」）——最终报表不应带操作说明。
    for sn in (SHEET_IS, SHEET_CF, SHEET_IS_Q, SHEET_CF_Q):
        ws = wb[sn]
        for row in ws.iter_rows(min_col=5, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.value = None

    wb.save(out_path)
    return out_path
