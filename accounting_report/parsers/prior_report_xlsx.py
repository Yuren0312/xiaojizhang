# -*- coding: utf-8 -*-
"""上月由本程序生成的 xlsx → PriorPeriod

滚动出表时优先走这条路：比解析 PDF 准且快。首月没有上月输出，才回退到 PDF。
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from ..config import SHEET_BS, SHEET_CF, SHEET_CF_Q, SHEET_IS, SHEET_IS_Q
from ..models import PriorPeriod, money
from ..rules import layout

DATE_RE = re.compile(r"(20\d{2})-(\d{2})")


def _read(ws, conf, col_name) -> dict[int, "object"]:
    out = {}
    for line, cells in conf["lines"].items():
        col = cells.get(col_name)
        if col:
            out[int(line)] = money(ws[f"{col}{cells['row']}"].value)
    return out


def parse(xlsx_path: str | Path) -> PriorPeriod:
    xlsx_path = Path(xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    spec = layout()

    missing = [s for s in (SHEET_BS, SHEET_IS, SHEET_CF, SHEET_IS_Q, SHEET_CF_Q)
               if s not in wb.sheetnames]
    if missing:
        raise ValueError(f"{xlsx_path.name} 缺少工作表：{missing}")

    bs_conf = spec[SHEET_BS]
    raw_date = wb[SHEET_BS][bs_conf["date_cell"]].value
    m = DATE_RE.search(str(raw_date or ""))
    if not m:
        raise ValueError(f"{xlsx_path.name} 的 {bs_conf['date_cell']} 读不到期间日期")
    prior = PriorPeriod(year=int(m.group(1)), month=int(m.group(2)))

    prior.bs_end = _read(wb[SHEET_BS], bs_conf, "end")
    prior.bs_year_start = _read(wb[SHEET_BS], bs_conf, "year_start")
    prior.is_ytd = _read(wb[SHEET_IS], spec[SHEET_IS], "ytd")
    prior.cf_ytd = _read(wb[SHEET_CF], spec[SHEET_CF], "ytd")

    for field, sheet in (("is_quarters", SHEET_IS_Q), ("cf_quarters", SHEET_CF_Q)):
        conf = spec[sheet]
        columns = {q: _read(wb[sheet], conf, f"q{q}") for q in range(1, 5)}
        setattr(prior, field, {
            int(line): {q: columns[q].get(int(line)) for q in range(1, 5)}
            for line in conf["lines"]
        })

    return prior
