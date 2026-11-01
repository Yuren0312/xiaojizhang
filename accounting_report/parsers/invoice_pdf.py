# -*- coding: utf-8 -*-
"""发票 PDF → 发票台账（体验官夜间轮新增）

解决「用户明明上传了发票 PDF，增值税底稿却提示缺发票台账数据源」的体验断裂：
扫描输入目录下的发票 PDF（电子发票，文本型），用 pdfplumber 提取关键字段，
生成符合 parsers/invoice_ledger.py 期望格式的台账 Excel。

识别字段：发票号码 / 开票日期 / 购方名称 / 销方名称 / 不含税金额 / 税额 / 价税合计
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import List

import pdfplumber

from ..models import money, ZERO


@dataclass
class PdfInvoice:
    """从 PDF 识别的一张发票"""
    number: str = ""
    date: str = ""
    buyer: str = ""          # 购方名称
    seller: str = ""         # 销方名称
    service: str = ""        # 货物/服务名称（*项目*名称）
    amount: Decimal = ZERO   # 不含税金额
    tax: Decimal = ZERO      # 税额
    total: Decimal = ZERO    # 价税合计
    source: str = ""


def _clean_num(s: str) -> Decimal:
    return money(s.replace(",", "").replace("，", "").replace("¥", "").strip())


def parse_invoice_text(text: str, source: str = "") -> PdfInvoice:
    """从电子发票 PDF 文本提取字段（容错：缺什么字段补什么）"""
    inv = PdfInvoice(source=source)
    m = re.search(r"发票号码[:：]\s*([0-9]+)", text)
    if m:
        inv.number = m.group(1)
    else:
        # 兜底：号码可能被分列布局拆到水印区，找 20 位纯数字
        m = re.search(r"(?<!\d)(\d{20})(?!\d)", text)
        if m:
            inv.number = m.group(1)
    m = re.search(r"开票日期[:：]\s*(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", text)
    if m:
        inv.date = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    else:
        # 兜底：日期可能被列布局拆散，全局找「年/月/日」格式
        m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
        if m:
            inv.date = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    # 购方/销方名称（文本级：同一行/相邻行/半角冒号）
    m = re.search(r"购\s*(?:方)?\s*名\s*称[:：]?\s*([^\s：:]{2,40})\s*销\s*(?:方)?\s*名\s*称[:：]?\s*([^\s：:]{2,40})", text)
    if m:
        inv.buyer, inv.seller = m.group(1).strip("，,。、"), m.group(2).strip("，,。、")
    else:
        m = re.search(r"购\s*(?:方)?\s*名\s*称[:：]?\s*([^\s：:]{2,40})", text)
        if m:
            inv.buyer = m.group(1).strip("，,。、")
        m = re.search(r"销\s*(?:方)?\s*名\s*称[:：]?\s*([^\s：:]{2,40})", text)
        if m:
            inv.seller = m.group(1).strip("，,。、")
    # 服务项目名称（*项目*明细，如 *餐饮服务*餐费）
    m = re.search(r"\*([^*]{2,20})\*([^*\s]{1,20})", text)
    if m:
        inv.service = f"{m.group(1)}-{m.group(2)}"
    # 合计行：不含税金额 + 税额
    # 兼容三种写法：
    #   「合 计 ¥26.53 ¥0.27」          （两个金额连续）
    #   「合计 ¥26.53 税额 ¥0.27」       （中间带「税额」标签，常见于新版数电票）
    #   「合计：不含税 ¥26.53 税额 ¥0.27」
    m = re.search(
        r"合\s*计\s*[:：]?\s*(?:不含税\s*[金额]?[:：]?)?[¥￥]?([\d,]+\.\d{2})"
        r"\s*(?:税额|税\s*额)?\s*[:：]?\s*[¥￥]?([\d,]+\.\d{2})",
        text,
    )
    if m:
        inv.amount, inv.tax = _clean_num(m.group(1)), _clean_num(m.group(2))
    # 价税合计（小写）
    m = re.search(r"（小写）\s*[¥￥]?\s*([\d,]+\.\d{2})", text)
    if m:
        inv.total = _clean_num(m.group(1))
    # 兜底：若未识别合计但有小写价税合计，金额按价税合计近似（税率 0 标注）
    if inv.amount == ZERO and inv.total > ZERO:
        inv.amount = inv.total
    return inv


def _extract_columns(page) -> tuple[str, str]:
    """电子发票 PDF 是左右列布局（左=购买方信息，右=销售方信息）。

    按 x 坐标中轴把单词分到左右两列，重组为两段文本，
    解决「购买方信息」竖排标题被拆成 购/买/方/信/息 单字、名称标签与值分离的问题。
    """
    try:
        words = page.extract_words(x_tolerance=1.5)
    except Exception:
        return "", ""
    if not words:
        return "", ""
    x0s = [w["x0"] for w in words]
    x1s = [w["x1"] for w in words]
    mid = (min(x0s) + max(x1s)) / 2
    left, right = [], []
    for w in words:
        if w["x1"] < mid:
            left.append(w["text"])
        elif w["x0"] > mid:
            right.append(w["text"])
        else:
            left.append(w["text"])
    return " ".join(left), " ".join(right)


_SKIP_VALUES = {"售", "销", "统一社会信用代码", "统一社会信用代码/纳税人识别号",
                "项目名称", "规格型号", "单", "位", "数", "量"}


def _party_from_column(col: str) -> str:
    """列内：名称标签后取值（值至少 2 个连续字符，排除标签词）"""
    m = re.search(r"名\s*称\s*[:：]\s*([^\s：:]{2,40})", col)
    if m:
        v = m.group(1).strip("，,。、")
        if v not in _SKIP_VALUES and len(v) >= 2:
            return v
    return ""


def _name_before_taxid(col: str) -> str:
    """列内：名称紧邻 15-20 位税号的模式（名称+纳税人识别号成对出现）"""
    m = re.search(r"([\u4e00-\u9fff（）()]{3,40}?)\s*([0-9A-Z]{15,20})", col)
    if m:
        v = m.group(1).strip("，,。、")
        if len(v) >= 3:
            return v
    return ""


def parse_invoice_pages(pages, source: str = "") -> PdfInvoice:
    """基于页面对象解析：先文本级，再按左右列补全购方/销方名称"""
    text = "\n".join((p.extract_text() or "") for p in pages)
    inv = parse_invoice_text(text, source=source)
    if inv.buyer and inv.seller:
        return inv
    lefts, rights = [], []
    for p in pages:
        l, r = _extract_columns(p)
        lefts.append(l)
        rights.append(r)
    left = " ".join(lefts)
    right = " ".join(rights)
    if not inv.buyer:
        inv.buyer = _party_from_column(left) or _name_before_taxid(left)
    if not inv.seller:
        inv.seller = _party_from_column(right) or _name_before_taxid(right)
    return inv


def parse_invoice_pdfs(pdf_files: List[Path]) -> List[PdfInvoice]:
    """批量解析发票 PDF，返回识别结果列表"""
    invoices: List[PdfInvoice] = []
    for pdf in pdf_files:
        try:
            with pdfplumber.open(str(pdf)) as doc:
                # 在 with 块内完成解析——pdfplumber 关闭后 doc.pages 不可用，
                # 若在块外调用 parse_invoice_pages 会抛异常被吞，发票台账静默变空
                inv = parse_invoice_pages(doc.pages, source=pdf.name)
            if inv.number or inv.amount > ZERO:
                invoices.append(inv)
        except Exception:
            continue  # 单张失败不阻断
    return invoices


def write_ledger(invoices: List[PdfInvoice], out_path: Path) -> Path:
    """生成符合 invoice_ledger 期望格式的发票台账 xlsx"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发票台账"
    headers = ["开票日期", "发票代码", "发票号码", "发票类型", "购方名称", "货物/服务名称",
               "不含税金额", "税额", "含税金额", "备注", "销方名称"]
    hf = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="2F5496")
    for j, h in enumerate(headers, 1):
        c = ws.cell(1, j, h)
        c.font = hf
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
    row = 2
    for inv in invoices:
        inv_type = "e_normal" if "电子" in inv.source else "normal"
        ws.cell(row, 1, inv.date)
        ws.cell(row, 3, inv.number)
        ws.cell(row, 4, inv_type)
        ws.cell(row, 5, inv.buyer)
        ws.cell(row, 6, inv.service)
        ws.cell(row, 7, float(inv.amount)).number_format = "#,##0.00"
        ws.cell(row, 8, float(inv.tax)).number_format = "#,##0.00"
        ws.cell(row, 9, float(inv.total)).number_format = "#,##0.00"
        ws.cell(row, 10, f"PDF解析:{inv.source}")
        ws.cell(row, 11, inv.seller)
        row += 1
    for col, w in zip("ABCDEFGHIJK", (12, 14, 22, 16, 26, 22, 14, 12, 14, 24, 28)):
        ws.column_dimensions[col].width = w
    wb.save(out_path)
    return out_path
