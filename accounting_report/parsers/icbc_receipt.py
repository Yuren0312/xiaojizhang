# -*- coding: utf-8 -*-
"""工商银行业务回单 PDF → 单笔凭证明细

回单是「标签: 值」的版面。它不作为记账的主数据源（对账单才是），
用途是交叉校验：金额、日期、摘要能否和对账单里的某笔流水对上。
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from ..models import money, ZERO
from ..utils.pdf_utils import find_table_words
from .pdf_table import center_x, group_rows

# 需要的标签 → 结果字段
WANTED = {
    "记账日期:": "date",
    "小写:": "amount",
    "摘要:": "summary",
    "付款人户名:": "payer",
    "收款人户名:": "payee",
    "凭证号码:": "voucher",
}


@dataclass
class Receipt:
    date: str = ""
    amount: Decimal = ZERO
    summary: str = ""
    payer: str = ""
    payee: str = ""
    voucher: str = ""
    source: str = ""


def parse(pdf_path: str | Path) -> Receipt:
    pdf_path = Path(pdf_path)
    words = list(find_table_words(str(pdf_path), 0))
    if not words:
        raise ValueError(f"{pdf_path.name} 没有可提取的文本")

    result = Receipt(source=pdf_path.name)

    for row in group_rows(words):
        # 行内按 x 排序后，每个标签后面、下一个标签之前的词就是它的值。
        # 边界是「任何以冒号结尾的词」，不只是我们关心的那几个——否则
        # 「付款人户名: 财付通… 付款人账号: 2433…」会把账号一起吞进户名。
        boundaries = [i for i, w in enumerate(row)
                      if w[4].strip().endswith((":", "："))]
        for pos, i in enumerate(boundaries):
            field = WANTED.get(row[i][4].strip())
            if field is None:
                continue
            end = boundaries[pos + 1] if pos + 1 < len(boundaries) else len(row)
            value = "".join(t[4].strip() for t in row[i + 1:end]
                            if center_x(t) > center_x(row[i]))
            if not value:
                continue
            if field == "amount":
                result.amount = money(value)
            else:
                setattr(result, field, value)

    return result
