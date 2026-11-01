# -*- coding: utf-8 -*-
"""业务事项 → 分录

银行流水只能反映「钱动了」的业务。挂账费用、应收应付、计提折旧这类不动钱的业务，
银行对账单里一个字都没有，必须另有来源。这里支持两种：

  1. 输入/业务事项.xlsx —— 人工填，列：日期 / 摘要 / 借方科目 / 贷方科目 / 金额 / 来源
  2. 输入/票据记录.csv  —— 「财务票据识别录入」OCR 工具的产出，按供应商映射科目

两者都会被转成分录，与银行流水汇合。
"""
from __future__ import annotations

import csv
from pathlib import Path

import openpyxl

from ..models import JournalEntry, money, ZERO
from ..rules import bank_mapping

COLUMNS = ["日期", "摘要", "借方科目", "贷方科目", "金额", "来源"]

import re as _re


def _norm_date(value) -> str:
    """日期单元格 → YYYY-MM-DD

    兼容 Excel 日期、ISO 文本、手填的 2026-6-5 / 2026/6/5 / 20260605。
    无法识别的值原样返回，由调用方校验（防静默丢行）。
    """
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    # Excel 序列号数字（如 46173）→ 忽略，按不可识别处理
    if s.isdigit() and len(s) != 8:
        return s
    m = _re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    m = _re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{mo}-{d}"
    return s[:10]


def parse_xlsx(path: str | Path) -> list[JournalEntry]:
    path = Path(path)
    if not path.exists():
        return []

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    missing = [c for c in COLUMNS[:5] if c not in header]
    if missing:
        raise ValueError(f"{path.name} 缺少必需列：{missing}；应为 {COLUMNS}")
    idx = {name: header.index(name) for name in header if name}

    entries: list[JournalEntry] = []
    for rowno, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue

        def cell(name: str):
            i = idx.get(name)
            return row[i] if i is not None and i < len(row) else None

        amount = money(cell("金额"))
        if amount == ZERO:
            continue

        debit = str(cell("借方科目") or "").strip()
        credit = str(cell("贷方科目") or "").strip()
        if not debit or not credit:
            raise ValueError(f"{path.name} 第 {rowno} 行缺少借方或贷方科目")

        date = _norm_date(cell("日期"))
        # 日期缺失/非法必须显式报错：此前空日期/非 ISO 写法（2026/6/5、20260605）
        # 会在下游按月前缀过滤时被静默丢弃，用户以为已入账实则没入（体验坑）。
        if not _re.match(r"^\d{4}-\d{2}-\d{2}$", date):
            raise ValueError(
                f"{path.name} 第 {rowno} 行日期「{cell('日期') or ''}」无法识别，"
                "请填成 YYYY-MM-DD（如 2026-06-05）"
            )

        entries.append(JournalEntry(
            date=date,
            summary=str(cell("摘要") or "").strip(),
            debit=debit,
            credit=credit,
            amount=amount,
            source=str(cell("来源") or path.name).strip(),
        ))

    return entries


def parse_invoice_csv(path: str | Path, month_prefix: str) -> list[JournalEntry]:
    """票据记录.csv → 分录（借：按供应商映射的费用科目；贷：应付账款）

    只取发票日期落在本月的记录。匹配不到科目的供应商会被报错列出，不猜。
    """
    path = Path(path)
    if not path.exists():
        return []

    rules = bank_mapping()["bank_rules"]
    entries: list[JournalEntry] = []
    unmatched: list[str] = []

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            date = (row.get("发票日期") or "").strip()
            if not date.startswith(month_prefix):
                continue
            amount = money(row.get("金额"))
            if amount == ZERO:
                continue

            vendor = (row.get("供应商") or "").strip()
            account = next(
                (r["account"] for r in rules if r["match"] in vendor), None
            )
            if account is None:
                unmatched.append(f"{vendor}（{row.get('文件名', '')}）")
                continue

            entries.append(JournalEntry(
                date=date,
                summary=f"{vendor} 发票 {row.get('发票号码', '')}".strip(),
                debit=account,
                credit="应付账款",
                amount=amount,
                source=f"{path.name}:{row.get('文件名', '')}",
            ))

    if unmatched:
        raise ValueError(
            "以下票据的供应商无法映射到科目，请在 rules/mapping.yaml 补充规则：\n  "
            + "\n  ".join(unmatched)
        )
    return entries


def write_template(path: str | Path) -> Path:
    """生成一张空的业务事项表，方便手工填"""
    path = Path(path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "业务事项"
    ws.append(COLUMNS)
    for i, width in enumerate([12, 40, 16, 16, 14, 24], start=1):
        ws.column_dimensions[chr(64 + i)].width = width
    wb.save(path)
    return path
