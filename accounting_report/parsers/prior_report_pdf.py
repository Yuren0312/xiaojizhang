# -*- coding: utf-8 -*-
"""上月 5 表 PDF → PriorPeriod

财务软件导出的 PDF 固定 5 页：资产负债表 / 利润表 / 现金流量表 / 利润表季报 / 现金流量表季报。
页序不写死，按标题识别。
"""
from __future__ import annotations

import re
from pathlib import Path

import fitz

from ..models import PriorPeriod, ZERO
from .pdf_table import parse_page

# 资产负债表左右分栏的 x 分界（页宽 595pt，左栏值列在 260 以内，右栏在 460 以上）
BS_SPLIT_X = 290.0

DATE_RE = re.compile(r"(20\d{2})-(\d{2})(?:-(\d{2}))?")


def _identify_pages(pdf_path: Path) -> dict[str, int]:
    """返回 {表名: 页号}"""
    doc = fitz.open(str(pdf_path))
    pages: dict[str, int] = {}
    for i, page in enumerate(doc):
        text = page.get_text()
        head = text[:200]
        quarterly = "第一季度" in text
        if head.startswith("资产负债表"):
            pages["资产负债表"] = i
        elif head.startswith("利润表"):
            pages["利润表季报" if quarterly else "利润表"] = i
        elif head.startswith("现金流量表"):
            pages["现金流量表季报" if quarterly else "现金流量表"] = i
    doc.close()

    missing = {"资产负债表", "利润表", "现金流量表", "利润表季报", "现金流量表季报"} - set(pages)
    if missing:
        raise ValueError(f"{pdf_path.name} 缺少这些表：{sorted(missing)}")
    return pages


def _period(pdf_path: Path, page_no: int) -> tuple[int, int]:
    doc = fitz.open(str(pdf_path))
    text = doc[page_no].get_text()
    doc.close()
    m = DATE_RE.search(text)
    if not m:
        raise ValueError(f"{pdf_path.name} 第 {page_no + 1} 页找不到期间日期")
    return int(m.group(1)), int(m.group(2))


def parse(pdf_path: str | Path) -> PriorPeriod:
    pdf_path = Path(pdf_path)
    pages = _identify_pages(pdf_path)
    year, month = _period(pdf_path, pages["资产负债表"])

    bs = parse_page(str(pdf_path), pages["资产负债表"],
                    ["期末余额", "年初余额"], lineno_split_x=BS_SPLIT_X)
    inc = parse_page(str(pdf_path), pages["利润表"], ["本年累计金额", "本月金额"])
    cf = parse_page(str(pdf_path), pages["现金流量表"], ["本年累计金额", "本月金额"])
    q_heads = ["本年累计", "第一季度", "第二季度", "第三季度", "第四季度"]
    inc_q = parse_page(str(pdf_path), pages["利润表季报"], q_heads)
    cf_q = parse_page(str(pdf_path), pages["现金流量表季报"], q_heads)

    _require_lines(bs, 53, "资产负债表")
    _require_lines(inc, 32, "利润表")
    _require_lines(cf, 22, "现金流量表")
    _require_lines(inc_q, 32, "利润表季报")
    _require_lines(cf_q, 22, "现金流量表季报")

    prior = PriorPeriod(year=year, month=month)
    # 资产负债表：列 0/2 = 期末余额（左/右栏），列 1/3 = 年初余额
    for line, cols in bs.items():
        prior.bs_end[line] = cols.get(0, cols.get(2, ZERO))
        prior.bs_year_start[line] = cols.get(1, cols.get(3, ZERO))
    for line, cols in inc.items():
        prior.is_ytd[line] = cols.get(0, ZERO)
    for line, cols in cf.items():
        prior.cf_ytd[line] = cols.get(0, ZERO)
    for line, cols in inc_q.items():
        prior.is_quarters[line] = {q: cols.get(q, ZERO) for q in range(1, 5)}
    for line, cols in cf_q.items():
        prior.cf_quarters[line] = {q: cols.get(q, ZERO) for q in range(1, 5)}

    return prior


def _require_lines(parsed: dict, expected: int, name: str) -> None:
    missing = set(range(1, expected + 1)) - set(parsed)
    if missing:
        raise ValueError(f"{name} 解析缺失行次 {sorted(missing)}，PDF 版式可能变了")
