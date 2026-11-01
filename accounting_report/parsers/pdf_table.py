# -*- coding: utf-8 -*-
"""PDF 表格的坐标解析基础设施

财务软件导出的 PDF 里，左右两栏在文本流里是交错的（`99,368.42   短期借款`），
按行读文本必然串栏。这里改用词坐标：

  1. 按 y 聚成视觉行
  2. 用表头词的中心确定各值列的 x 中心
  3. 每行以「行次」整数为锚点，数值按中心距就近归列

数值一律是「元至角分」两位小数（`49,977.25`），行次是 1~2 位纯整数，
两者用正则就能干净分开，不必依赖位置猜测。
"""
from __future__ import annotations

import re
from decimal import Decimal

from ..models import money
from ..utils.pdf_utils import find_table_words

# 49,977.25 / -835.63 / 0.00 ；必须带两位小数，才不会和行次混淆
NUMBER_RE = re.compile(r"^-?[\d,]*\d\.\d{2}$")
LINENO_RE = re.compile(r"^\d{1,2}$")

Word = tuple[float, float, float, float, str, int, int, int]


def is_number(text: str) -> bool:
    return bool(NUMBER_RE.match(text.strip()))


def is_lineno(text: str) -> bool:
    return bool(LINENO_RE.match(text.strip()))


def center_x(w: Word) -> float:
    return (w[0] + w[2]) / 2


def center_y(w: Word) -> float:
    return (w[1] + w[3]) / 2


def group_rows(words: list[Word], tol: float = 4.0) -> list[list[Word]]:
    """按 y 中心把词聚成视觉行，行内按 x 排序"""
    rows: list[tuple[float, list[Word]]] = []
    for w in sorted(words, key=center_y):
        cy = center_y(w)
        if rows and abs(cy - rows[-1][0]) <= tol:
            rows[-1][1].append(w)
        else:
            rows.append((cy, [w]))
    return [sorted(ws, key=lambda w: w[0]) for _, ws in rows]


def header_centers(words: list[Word], labels: list[str]) -> list[float]:
    """按给定表头文字（可重复，如资产负债表有两个「期末余额」）取列中心，从左到右"""
    found: list[float] = []
    for label in dict.fromkeys(labels):  # 去重但保序
        hits = [center_x(w) for w in words if w[4].strip() == label]
        if not hits:
            raise ValueError(f"PDF 中找不到表头 {label!r}，版式可能变了")
        found.extend(hits)
    return sorted(found)


def assign_columns(
    numbers: list[Word], centers: list[float], max_offset: float = 45.0
) -> dict[int, Decimal]:
    """数值按中心距就近归到第几列（0-based）。偏得太远说明版式变了，报错。"""
    out: dict[int, Decimal] = {}
    for w in numbers:
        cx = center_x(w)
        idx = min(range(len(centers)), key=lambda i: abs(cx - centers[i]))
        offset = abs(cx - centers[idx])
        if offset > max_offset:
            raise ValueError(
                f"数值 {w[4]!r}(x={cx:.1f}) 离最近的列中心 {centers[idx]:.1f} "
                f"有 {offset:.1f}pt，超出容差，请检查 PDF 版式"
            )
        if idx in out:
            raise ValueError(f"第 {idx} 列出现两个数值：{out[idx]} 与 {w[4]}")
        out[idx] = money(w[4])
    return out


def parse_page(
    pdf_path: str,
    page_no: int,
    value_headers: list[str],
    lineno_split_x: float | None = None,
) -> dict[int, dict[int, Decimal]]:
    """解析一页表格

    Args:
        value_headers: 值列表头文字，按从左到右的视觉顺序
        lineno_split_x: 资产负债表有左右两个「行次」列，传入分界 x；
                        其余表传 None

    Returns:
        {行次: {列序号: 金额}}，列序号与 value_headers 排序后的列一一对应
    """
    words = list(find_table_words(pdf_path, page_no))
    if not words:
        raise ValueError(f"{pdf_path} 第 {page_no + 1} 页没有可提取的文本")

    centers = header_centers(words, value_headers)
    result: dict[int, dict[int, Decimal]] = {}

    for row in group_rows(words):
        linenos = [w for w in row if is_lineno(w[4])]
        numbers = [w for w in row if is_number(w[4])]
        if not linenos:
            continue

        if lineno_split_x is None:
            if len(linenos) != 1:
                raise ValueError(f"一行里出现多个行次：{[w[4] for w in linenos]}")
            result[int(linenos[0][4])] = assign_columns(numbers, centers)
        else:
            # 左右分栏：行次和数值都按分界线切开，各自归各自的列
            half = len(centers) // 2
            for side, (lo, hi, cols) in enumerate(
                [(-1e9, lineno_split_x, centers[:half]),
                 (lineno_split_x, 1e9, centers[half:])]
            ):
                side_lines = [w for w in linenos if lo < center_x(w) <= hi]
                if not side_lines:
                    continue
                if len(side_lines) != 1:
                    raise ValueError(f"同侧出现多个行次：{[w[4] for w in side_lines]}")
                side_nums = [w for w in numbers if lo < center_x(w) <= hi]
                assigned = assign_columns(side_nums, cols)
                result[int(side_lines[0][4])] = {
                    k + side * half: v for k, v in assigned.items()
                }

    return result
