# -*- coding: utf-8 -*-
"""输出10：电子税务局报送格式（会小企01表 资产负债表 / 会小企02表 利润表）

电子税务局「财务报表报送与信息采集（小企业会计准则）」只接受它自己的固定模板，
而引擎产出的会计报表是自排版式，两者列名、行次排布、表格线都不一致，实务上只能
人工誊抄一遍再提交。本模块直接以税局模板为底稿，把 ReportSet 按**行次**填进去，
产出与模板逐格一致的单表文件（值 / 字体 / 边框 / 对齐 / 数字格式 / 合并区 / 列宽），
下载即可上传，无需二次整理。

为什么不用 openpyxl：税局模板是老 .xls（BIFF8），样式存放在 XF 索引表里。
必须用 xlrd(formatting_info=True) 读出每个单元格的 XF，再经 xlwt 原样写回；
否则 xlwt 默认样式会把模板的表格线和底纹覆盖掉，导出的表反而不像原表。

数据口径（与用户已成功提交的样例一致）：
    资产负债表  期末余额 = 本期末，年初余额 = 本年年初
    利润表      本期金额 = 当月发生额，本年累计金额 = 本年累计

输出：
    {YYYY-MM} 资产负债表（税局报送）.xls
    {YYYY-MM} 利润表（税局报送）.xls
"""
from __future__ import annotations

import calendar
from datetime import date
from decimal import Decimal
from pathlib import Path

import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

from ..config import TEMPLATES_DIR
from ..models import ReportSet

TEMPLATE_NAME = "财务报表报送与信息采集（小企业会计准则）月季报.xls"
SHEET_BS = "资产负债表"
SHEET_IS = "利润表_月季报"

# 模板列位置（0-based）。资产负债表左右两栏各有一套「行次 / 期末 / 年初」
BS_LINE_L, BS_END_L, BS_START_L = 2, 3, 4
BS_LINE_R, BS_END_R, BS_START_R = 6, 7, 8
IS_LINE, IS_MONTH, IS_YTD = 2, 3, 4

# 表头锚点：(行, 列, 值)
BS_HEADER = [(2, 3, "tax_no"), (2, 7, "company"), (3, 3, "start"), (3, 7, "end")]
IS_HEADER = [(2, 3, "tax_no"), (2, 5, "company"), (3, 3, "start"), (3, 5, "end")]

HORZ = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7}
VERT = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4}
UNDERLINE = {0: 0, 1: 1, 2: 2, 33: 33, 34: 34}


def template_path() -> Path:
    return TEMPLATES_DIR / TEMPLATE_NAME


def _num(v) -> float:
    """Decimal/None → float（xlwt 不接受 Decimal）"""
    if v is None:
        return 0.0
    if isinstance(v, Decimal):
        return float(v)
    return float(v)


def _period_range(year: int, month: int) -> tuple[date, date]:
    """所属期：季末月（3/6/9/12）按整季填，其余月份按当月。

    电子税务局这张表叫「月季报」，季度申报时所属期起是季度首月 1 日。
    """
    last_day = calendar.monthrange(year, month)[1]
    if month in (3, 6, 9, 12):
        return date(year, month - 2, 1), date(year, month, last_day)
    return date(year, month, 1), date(year, month, last_day)


def _xf_to_style(src_wb, xf_index: int, cache: dict) -> xlwt.XFStyle:
    """把源工作簿的 XF 索引翻译成等价的 xlwt 样式（字体/对齐/边框/底纹/数字格式）"""
    if xf_index in cache:
        return cache[xf_index]

    xf = src_wb.xf_list[xf_index]
    st = xlwt.XFStyle()

    f = src_wb.font_list[xf.font_index]
    font = xlwt.Font()
    font.name = f.name
    font.height = f.height
    font.bold = bool(f.bold)
    font.italic = bool(f.italic)
    for attr in ("struck_out", "outline", "shadow"):
        if hasattr(f, attr):
            setattr(font, attr, bool(getattr(f, attr)))
    font.colour_index = getattr(f, "colour_index", 8)
    font.escapement = getattr(f, "escapement", 0)
    font.underline = UNDERLINE.get(getattr(f, "underline_type", 0), 0)
    if hasattr(f, "family"):
        font.family = f.family
    if hasattr(f, "charset"):
        font.charset = f.charset
    st.font = font

    a = xf.alignment
    al = xlwt.Alignment()
    al.horz = HORZ.get(a.hor_align, 0)
    al.vert = VERT.get(a.vert_align, 2)
    al.wrap = 1 if a.text_wrapped else 0
    al.rota = a.rotation
    al.inde = a.indent_level
    al.shri = 1 if a.shrink_to_fit else 0
    st.alignment = al

    b = xf.border
    bd = xlwt.Borders()
    bd.left, bd.right = b.left_line_style, b.right_line_style
    bd.top, bd.bottom = b.top_line_style, b.bottom_line_style
    bd.left_colour, bd.right_colour = b.left_colour_index, b.right_colour_index
    bd.top_colour, bd.bottom_colour = b.top_colour_index, b.bottom_colour_index
    st.borders = bd

    bg = xf.background
    pt = xlwt.Pattern()
    pt.pattern = bg.fill_pattern
    pt.pattern_fore_colour = bg.pattern_colour_index
    pt.pattern_back_colour = bg.background_colour_index
    st.pattern = pt

    st.num_format_str = src_wb.format_map[xf.format_key].format_str
    cache[xf_index] = st
    return st


def _build(tpl: Path, sheet_name: str, header: dict, body_writer, out_path: Path) -> None:
    """以模板单表为底稿填数并另存"""
    rb = xlrd.open_workbook(str(tpl), formatting_info=True)
    try:
        tpl_sh = rb.sheet_by_name(sheet_name)
    except xlrd.biffh.XLRDError as e:
        raise ValueError(f"模板缺少工作表「{sheet_name}」：{e}") from e

    wb = xl_copy(rb)
    keep = [ws for ws in wb._Workbook__worksheets if ws.name == sheet_name]
    if not keep:
        raise ValueError(f"复制模板后找不到工作表「{sheet_name}」")
    ws = keep[0]
    # 只保留目标表：电子税务局按单表上传，多表反而对不上
    wb._Workbook__worksheets = keep
    wb._Workbook__active_sheet = 0

    # xlutils 复制会带上 xlwt 默认页眉页脚，模板本身为空，清掉
    ws.header_str = b""
    ws.footer_str = b""

    # xlutils 不复制页面设置，按模板补齐（A4 纵向、缩放适应 1 页宽）
    for attr, val in [("left_margin", 0.9), ("right_margin", 0.9),
                      ("top_margin", 0.98), ("bottom_margin", 0.98),
                      ("header_margin", 0.51), ("footer_margin", 0.51),
                      ("paper_size_code", 9), ("portrait", 1), ("print_scaling", 100),
                      ("fit_num_pages", 0), ("fit_width_to_pages", 1),
                      ("fit_height_to_pages", 1), ("print_grid", 0), ("print_headers", 0),
                      ("print_centered_horz", 0), ("print_centered_vert", 0)]:
        if hasattr(ws, attr):
            setattr(ws, attr, val)

    # xlwt 默认调色板与模板自定义调色板不同（如索引 59 模板是浅灰，默认是深色），
    # 不搬过来会把底色刷黑
    try:
        for idx, rgb in rb.colour_map.items():
            if idx >= 8 and rgb:
                wb.set_colour_RGB(idx, *rgb)
    except Exception:
        pass

    cache: dict = {}

    def put(r: int, c: int, v) -> None:
        ws.write(r, c, v, _xf_to_style(rb, tpl_sh.cell_xf_index(r, c), cache))

    for (r, c), v in header.items():
        put(r, c, v)
    body_writer(put, tpl_sh)

    # xlwt 复制合并区时会重设外框，导致合并区域右/下边缘线变细；
    # 把合并区内非锚点单元格按模板原位样式回写一遍（内容留空）即可复原
    for (r1, r2, c1, c2) in tpl_sh.merged_cells:
        for r in range(r1, r2):
            for c in range(c1, c2):
                if (r, c) != (r1, c1):
                    put(r, c, "")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))


def _fill_bs(put, sh, report: ReportSet) -> None:
    """资产负债表：按模板上的行次取 ReportSet 的 bs_end / bs_year_start"""
    for r in range(sh.nrows):
        ln = sh.cell_value(r, BS_LINE_L)
        rn = sh.cell_value(r, BS_LINE_R)
        if isinstance(ln, float) and ln:
            line = int(ln)
            put(r, BS_END_L, _num(report.bs_end.get(line)))
            put(r, BS_START_L, _num(report.bs_year_start.get(line)))
        if isinstance(rn, float) and rn:
            line = int(rn)
            put(r, BS_END_R, _num(report.bs_end.get(line)))
            put(r, BS_START_R, _num(report.bs_year_start.get(line)))


def _fill_is(put, sh, report: ReportSet) -> None:
    """利润表：本期金额 = 当月，本年累计金额 = 累计"""
    for r in range(sh.nrows):
        ln = sh.cell_value(r, IS_LINE)
        if isinstance(ln, float) and ln:
            line = int(ln)
            put(r, IS_MONTH, _num(report.is_month.get(line)))
            put(r, IS_YTD, _num(report.is_ytd.get(line)))


def write_tax_filing(
    report: ReportSet,
    out_dir: Path | str,
    *,
    company_name: str = "",
    tax_no: str = "",
    template: Path | str | None = None,
) -> list[Path]:
    """生成电子税务局报送格式的资产负债表与利润表。

    Args:
        report: 本账期算好的报表（ReportSet，键为行次）。
        out_dir: 输出目录。
        company_name: 纳税人名称。
        tax_no: 纳税人识别号（统一社会信用代码）。
        template: 模板路径，默认用内置的月季报模板。

    Returns:
        生成的两个文件路径 [资产负债表, 利润表]。
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tpl = Path(template) if template else template_path()
    if not tpl.exists():
        raise FileNotFoundError(f"缺少电子税务局报送模板：{tpl}")

    start, end = _period_range(report.year, report.month)
    prefix = f"{report.year:04d}-{report.month:02d} "
    bs_path = out_dir / f"{prefix}资产负债表（税局报送）.xls"
    is_path = out_dir / f"{prefix}利润表（税局报送）.xls"

    common = {"tax_no": tax_no, "company": company_name, "start": start, "end": end}

    _build(tpl, SHEET_BS,
           {(r, c): common[key] for (r, c, key) in BS_HEADER},
           lambda put, sh: _fill_bs(put, sh, report), bs_path)

    _build(tpl, SHEET_IS,
           {(r, c): common[key] for (r, c, key) in IS_HEADER},
           lambda put, sh: _fill_is(put, sh, report), is_path)

    return [bs_path, is_path]
